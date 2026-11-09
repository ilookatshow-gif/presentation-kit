#!/usr/bin/env python3
"""Remove only an unmodified installation's exact owned files; keep other data."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import sys

NAME = "presentation-kit"
VERSION = "0.1.0"
MARKER = ".presentation-kit-install.json"
FILES = ("presentation_kit.py", "uninstall.py", "LICENSE", "NOTICE.md",
         "examples/synthetic-deck.json")


class UninstallError(Exception):
    pass


def safe_path(value):
    path = Path(os.path.abspath(str(value)))
    for part in (path,) + tuple(path.parents):
        if part.is_symlink():
            raise UninstallError("SYMLINK_REFUSED")
    return path


def uninstall(prefix, bin_dir):
    prefix, binary = safe_path(prefix), safe_path(Path(bin_dir) / NAME)
    marker = safe_path(prefix / MARKER)
    if not marker.is_file() or marker.stat().st_size > 20000:
        raise UninstallError("INSTALL_MARKER_MISSING")
    info = json.loads(marker.read_text(encoding="utf-8"))
    if (not isinstance(info, dict) or info.get("kind") != "presentation-kit-install"
            or info.get("version") != VERSION or info.get("prefix") != str(prefix)
            or info.get("binary") != str(binary) or not isinstance(info.get("files"), dict)
            or set(info["files"]) != set(FILES)):
        raise UninstallError("INSTALL_MARKER_INVALID")
    owned = [(safe_path(prefix / rel), info["files"][rel]) for rel in FILES]
    owned.append((binary, info.get("binary_sha256")))
    for path, digest in owned:
        if (not path.is_file() or path.stat().st_size > 5 * 1024 * 1024
                or hashlib.sha256(path.read_bytes()).hexdigest() != digest):
            raise UninstallError("INSTALL_MODIFIED_NO_REMOVAL")
    optional = prefix / ".onboarded"
    if optional.exists():
        safe_path(optional)
        if not optional.is_file() or optional.stat().st_size != 6 or optional.read_bytes() != (VERSION + "\n").encode("ascii"):
            raise UninstallError("ONBOARD_MARKER_MODIFIED")
        owned.append((optional, None))
    for path, unused in owned:
        path.unlink()
    marker.unlink()
    # Only remove exact known directories if empty; never recurse into user data.
    for directory in (prefix / "examples", prefix):
        try:
            directory.rmdir()
        except OSError:
            pass
    return {"ok": True, "command": "uninstall", "removed_files": len(owned) + 1,
            "other_files_preserved": True, "network_used": False}


class Parser(argparse.ArgumentParser):
    def error(self, message):
        raise UninstallError("CLI_USAGE")


def main(argv=None):
    try:
        parser = Parser(description="Remove exact unmodified owned files; preserve all other data.")
        parser.add_argument("--prefix", required=True)
        parser.add_argument("--bin-dir", required=True)
        args = parser.parse_args(argv)
        print(json.dumps(uninstall(args.prefix, args.bin_dir), sort_keys=True))
        return 0
    except (UninstallError, OSError, ValueError, TypeError):
        print(json.dumps({"ok": False, "error": {"code": "UNINSTALL_REFUSED", "message": "Installation is missing, modified or unsafe; inspect it without blind retry."}}, sort_keys=True))
        return 2


if __name__ == "__main__":
    sys.exit(main())
