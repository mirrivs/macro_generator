"""XOR obfuscation helper for generated training artifacts.

Applies a simple repeating-key XOR to a generated file and writes the result
next to it with a ``.xored`` suffix (for example ``Handover_Protocol.docm`` ->
``Handover_Protocol.docm.xored``). XOR is symmetric, so applying the same
operation again with the same key restores the original bytes.
"""

from __future__ import annotations

from pathlib import Path

XOR_SUFFIX = ".xored"
# Caldera's default repeating XOR key (payload_encoder.py):
#     DEFAULT_KEY = [0x32, 0x45, 0x32, 0xca]
DEFAULT_KEY = "0x324532CA"


def _coerce_key(key) -> bytes:
    """Normalise a config value into the byte key used for XOR."""
    if isinstance(key, bytes):
        result = bytes(key)
    elif isinstance(key, bytearray):
        result = bytes(key)
    elif isinstance(key, int):
        result = bytes([key & 0xFF])
    else:
        text = str(key).strip()
        if text.lower().startswith("0x"):
            digits = text[2:]
            if len(digits) % 2:
                digits = "0" + digits
            try:
                result = bytes.fromhex(digits)
            except ValueError as exc:
                raise ValueError(f"Invalid hex XOR key {key!r}") from exc
        else:
            result = text.encode("utf-8")

    if not result:
        raise ValueError("XOR key must not be empty")
    return result


def load_xor_key(project_dir: Path, default: str = DEFAULT_KEY) -> bytes:
    """Read ``app.xor.key`` from config.yml, falling back to ``default``."""
    import yaml

    key = default
    config_path = project_dir / "config.yml"
    if config_path.is_file():
        try:
            data = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
        except (OSError, yaml.YAMLError):
            data = {}
        xor_cfg = (data.get("app") or {}).get("xor") or {}
        if xor_cfg.get("key") is not None:
            key = xor_cfg["key"]
    return _coerce_key(key)


def xor_bytes(data: bytes, key: bytes) -> bytes:
    """Return ``data`` XOR'd with a repeating ``key``."""
    if not key:
        raise ValueError("XOR key must not be empty")
    key_len = len(key)
    return bytes(byte ^ key[index % key_len] for index, byte in enumerate(data))


def encrypt_file(source: Path, key: bytes) -> Path:
    """Write ``source`` XOR'd with ``key`` to ``<source>.xored`` and return it."""
    if not source.is_file():
        raise FileNotFoundError(f"Cannot XOR missing file: {source}")

    xored_path = source.with_name(source.name + XOR_SUFFIX)
    xored_path.write_bytes(xor_bytes(source.read_bytes(), key))
    return xored_path


def main(argv: list[str] | None = None) -> int:
    """Encrypt/decrypt a file from the command line.

    ``python -m utils.xor <file>`` writes ``<file>.xored``; running it on a
    ``.xored`` file restores the original name. The key is read from
    ``config.yml`` (``app.xor.key``).
    """
    import sys

    project_dir = Path(__file__).resolve().parent.parent

    if argv is None:
        argv = sys.argv[1:]
    if len(argv) != 1:
        print("Usage: python -m utils.xor <file>", file=sys.stderr)
        print("  <file>         -> writes <file>.xored (XOR-encrypt)", file=sys.stderr)
        print("  <file>.xored   -> restores <file> (XOR-decrypt)", file=sys.stderr)
        return 2

    source = Path(argv[0]).resolve()
    if not source.is_file():
        print(f"File not found: {source}", file=sys.stderr)
        return 1

    key = load_xor_key(project_dir)
    data = xor_bytes(source.read_bytes(), key)

    if source.name.endswith(XOR_SUFFIX):
        target = source.with_name(source.name[: -len(XOR_SUFFIX)])
        action = "Restored"
    else:
        target = source.with_name(source.name + XOR_SUFFIX)
        action = "Encrypted"
    target.write_bytes(data)
    print(f"{action}: {source} -> {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
