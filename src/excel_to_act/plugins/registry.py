"""Simple in-process plugin registry."""

from __future__ import annotations

from collections import defaultdict
from typing import Any


class PluginRegistry:
    def __init__(self) -> None:
        self._plugins: dict[str, dict[str, Any]] = defaultdict(dict)

    def register(self, capability: str, name: str, plugin: Any, *, replace: bool = False) -> None:
        if not replace and name in self._plugins[capability]:
            raise ValueError(f"plugin already registered for {capability!r}: {name!r}")
        self._plugins[capability][name] = plugin

    def get(self, capability: str, name: str) -> Any:
        try:
            return self._plugins[capability][name]
        except KeyError as exc:
            raise KeyError(f"unknown plugin {capability!r}/{name!r}") from exc

    def names(self, capability: str) -> list[str]:
        return sorted(self._plugins.get(capability, {}))


registry = PluginRegistry()
