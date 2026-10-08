"""Fail closed when a default runtime package has prohibited or unclear licensing metadata."""

from __future__ import annotations

import argparse
import re
import sys
from importlib import metadata
from typing import Iterable


PROHIBITED = re.compile(
    r"\b(?:GNU\s+(?:AFFERO\s+)?GENERAL\s+PUBLIC\s+LICENSE|"
    r"GNU\s+LESSER\s+GENERAL\s+PUBLIC\s+LICENSE|"
    r"AGPL|LGPL|GPL|EUPL)(?:-?V?\d+(?:\.\d+)*(?:[-+][A-Z0-9.]+)?)?\b",
    re.IGNORECASE,
)
OPERATORS = {"AND", "OR", "WITH"}
LICENSE_ALIASES = {
    "MIT": "MIT",
    "MITLICENSE": "MIT",
    "APACHE20": "Apache-2.0",
    "APACHESOFTWARELICENSE": "Apache-2.0",
    "BSD": "BSD",
    "BSDLICENSE": "BSD",
    "BSD2CLAUSE": "BSD-2-Clause",
    "BSD2CLAUSELICENSE": "BSD-2-Clause",
    "BSD3CLAUSE": "BSD-3-Clause",
    "BSD3CLAUSELICENSE": "BSD-3-Clause",
    "ISC": "ISC",
    "ISCLICENSE": "ISC",
    "PYTHONSOFTWAREFOUNDATIONLICENSE": "PSF-2.0",
    "PSF": "PSF-2.0",
    "PSF20": "PSF-2.0",
    "MOZILLAPUBLICLICENSE20": "MPL-2.0",
    "MPL20": "MPL-2.0",
    "ZLIB": "Zlib",
    "UNLICENSE": "Unlicense",
    "CC01": "CC0-1.0",
    "CCZERO10UNIVERSAL": "CC0-1.0",
    "ARTISTIC20": "Artistic-2.0",
    "ARTISTICLICENSE": "Artistic-2.0",
}
EXCEPTIONS = {"CLASSPATHEXCEPTION20", "LLVMEXCEPTION"}


class LicenseCheckError(ValueError):
    pass


def _normalize(value: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", value.upper())


def _validate_expression(value: str, source: str) -> None:
    tokens = []
    position = 0
    for match in re.finditer(r"[A-Za-z0-9][A-Za-z0-9.+-]*|[()]", value):
        if value[position : match.start()].strip():
            raise LicenseCheckError(f"malformed license expression in {source}: {value}")
        tokens.append(match.group())
        position = match.end()
    if value[position:].strip() or not tokens:
        raise LicenseCheckError(f"malformed license expression in {source}: {value}")

    index = 0

    def parse_primary() -> bool:
        nonlocal index
        if index == len(tokens):
            raise LicenseCheckError(f"malformed license expression in {source}: {value}")
        if tokens[index] == "(":
            index += 1
            parse_or()
            if index == len(tokens) or tokens[index] != ")":
                raise LicenseCheckError(f"malformed license expression in {source}: {value}")
            index += 1
            return False
        elif tokens[index] not in OPERATORS and tokens[index] != ")":
            index += 1
            return True
        else:
            raise LicenseCheckError(f"malformed license expression in {source}: {value}")

    def parse_with() -> None:
        nonlocal index
        bare_license = parse_primary()
        if index < len(tokens) and tokens[index].upper() == "WITH":
            if not bare_license:
                raise LicenseCheckError(f"malformed license expression in {source}: {value}")
            index += 1
            if index == len(tokens) or _normalize(tokens[index]) not in EXCEPTIONS:
                raise LicenseCheckError(f"malformed license expression in {source}: {value}")
            index += 1

    def parse_and() -> None:
        nonlocal index
        parse_with()
        while index < len(tokens) and tokens[index].upper() == "AND":
            index += 1
            parse_with()

    def parse_or() -> None:
        nonlocal index
        parse_and()
        while index < len(tokens) and tokens[index].upper() == "OR":
            index += 1
            parse_and()

    parse_or()
    if index != len(tokens):
        raise LicenseCheckError(f"malformed license expression in {source}: {value}")


def _check_value(value: str, source: str) -> None:
    if source == "License-Expression" or source.endswith("(License-Expression)"):
        _validate_expression(value, source)
    if PROHIBITED.search(value):
        raise LicenseCheckError(f"prohibited license family in {source}: {value}")
    if _normalize(value) in LICENSE_ALIASES:
        return
    classifier_name = re.sub(r"\s*\([^)]*\)\s*$", "", value)
    if _normalize(classifier_name) in LICENSE_ALIASES:
        return

    tokens = re.findall(r"[A-Za-z0-9][A-Za-z0-9.+-]*", value)
    if not tokens:
        raise LicenseCheckError(f"unclassifiable license metadata in {source}: {value!r}")
    operators = [token.upper() in OPERATORS for token in tokens]
    if operators[0] or operators[-1] or any(a and b for a, b in zip(operators, operators[1:])):
        raise LicenseCheckError(f"malformed license expression in {source}: {value}")
    for token in tokens:
        if token.upper() in OPERATORS:
            continue
        if _normalize(token) in LICENSE_ALIASES or _normalize(token) in EXCEPTIONS:
            continue
        raise LicenseCheckError(f"unclassifiable license identifier in {source}: {token}")


def check_metadata(name: str, values: Iterable[tuple[str, str]]) -> None:
    licenses = list(values)
    if not licenses:
        raise LicenseCheckError(f"missing license metadata: {name}")
    for source, value in licenses:
        _check_value(value, f"{name} ({source})")


def check_distribution(distribution: metadata.Distribution) -> None:
    name = distribution.metadata.get("Name", "<unnamed>")
    values: list[tuple[str, str]] = []
    for field in ("License-Expression", "License"):
        value = distribution.metadata.get(field)
        if value and value.strip():
            values.append((field, value.strip()))
    for classifier in distribution.metadata.get_all("Classifier", []):
        if classifier.startswith("License ::"):
            values.append(("Classifier", classifier.rsplit("::", 1)[-1].strip()))
    check_metadata(name, values)


def self_test() -> None:
    samples = (
        [("License-Expression", "MIT OR GPL-3.0-only")],
        [("License-Expression", "MIT OR")],
        [("License-Expression", "MIT BSD-3-Clause")],
        [("License-Expression", "MIT | BSD-3-Clause")],
        [("License-Expression", "(MIT OR BSD-3-Clause")],
        [("License-Expression", "MIT WITH Apache-2.0")],
        [("License-Expression", "(MIT) WITH LLVM-exception")],
        [("License", "GNU General Public License v3")],
        [("License-Expression", "MIT WITH EUPL-1.2")],
        [("License", "mystery license")],
        [],
    )
    for sample in samples:
        try:
            check_metadata("known-sample", sample)
        except LicenseCheckError:
            continue
        raise AssertionError(f"checker accepted violating/unknown metadata: {sample!r}")
    check_metadata("valid-mixed-sample", [("License-Expression", "MIT OR BSD-3-Clause")])
    check_metadata("valid-exception-sample", [("License-Expression", "MIT WITH LLVM-exception")])
    print("License checker rejects prohibited, malformed, unknown, and missing metadata samples.")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--sample-violation", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        self_test()
        return 0
    if args.sample_violation:
        try:
            check_metadata("known-sample", [("License-Expression", "MIT OR GPL-3.0-only")])
        except LicenseCheckError as error:
            print(error, file=sys.stderr)
            return 1
        print("License checker accepted a known violation.", file=sys.stderr)
        return 2

    failures = []
    for distribution in metadata.distributions():
        if distribution.metadata.get("Name", "").casefold() in {"pip", "setuptools", "excel-to-act"}:
            continue  # pip/setuptools bootstrap the venv; excel-to-act is the closure root.
        try:
            check_distribution(distribution)
        except LicenseCheckError as error:
            failures.append(str(error))
    if failures:
        print("Default runtime license check failed:", file=sys.stderr)
        print("\n".join(f"- {failure}" for failure in failures), file=sys.stderr)
        return 1
    print("Default runtime license check passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
