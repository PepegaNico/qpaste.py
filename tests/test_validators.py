from quickpaste.validators import HotkeyValidator


def test_normalize_and_validate():
    assert HotkeyValidator.is_valid('ctrl+shift+1')
    assert HotkeyValidator.is_valid(' CTRL + SHIFT + b ')
    assert not HotkeyValidator.is_valid('alt+shift+1')
    assert not HotkeyValidator.is_valid('ctrl+shift+')


def test_next_available():
    used = {f'ctrl+shift+{ch}': True for ch in '123'}
    used = set(used)
    nxt = HotkeyValidator.next_available(used)
    assert nxt.startswith('ctrl+shift+')
