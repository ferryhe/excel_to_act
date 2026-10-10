# Default runtime dependency license policy

The default runtime dependency closure must not contain GPL, AGPL, or EUPL license families. This is a project dependency-selection policy, not legal advice or a legal conclusion.

CI checks the installed distributions in a fresh environment after installing the project without extras. It reads `License-Expression`, legacy `License`, and license `Classifier` fields from installed package metadata. The check includes every installed distribution except `pip` and `setuptools`, which may be present as isolated-environment bootstrap/build tools, and the project itself, which is the closure root rather than a dependency. It therefore covers the resolved runtime dependencies and their transitive dependencies without dev or opt-in packages.

License-family identifiers are normalized across punctuation and version spellings. A prohibited identifier anywhere in a mixed SPDX expression is rejected, including `OR` and `WITH` expressions. Missing metadata, unknown identifiers, and unclassifiable license text fail the check. Optional extras are outside this default-install gate and must be reviewed separately before adoption.
