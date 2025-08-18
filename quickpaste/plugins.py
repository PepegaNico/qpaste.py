from typing import Callable, Dict, Any


class PluginManager:
    def __init__(self) -> None:
        self._plugins: Dict[str, Callable[..., Any]] = {}

    def register(self, name: str, func: Callable[..., Any]) -> None:
        self._plugins[name] = func

    def execute(self, name: str, *args, **kwargs) -> Any:
        if name not in self._plugins:
            raise KeyError(f"Plugin not found: {name}")
        return self._plugins[name](*args, **kwargs)

