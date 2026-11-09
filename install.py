#!/usr/bin/env python3
"""Original explicit-prefix installer; no network, dependencies or overwrite."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import sys

NAME = "presentation-kit"
VERSION = "0.1.0"
FILES = ("presentation_kit.py", "uninstall.py", "LICENSE", "NOTICE.md",
         "examples/synthetic-deck.json")
MARKER = ".presentation-kit-install.json"


class InstallError(Exception):
    pass


def safe_path(value):
    path = Path(os.path.abspath(str(value)))
    if path == Path(path.anchor) or path in (Path.home(), Path("/usr"), Path("/usr/local"), Path("/bin"),
                                     Path("/sbin"), Path("/etc"), Path("/var"), Path("/opt")):
        raise InstallError("PREFIX_TOO_BROAD")
    for part in (path,) + tuple(path.parents):
        if part.is_symlink():
            raise InstallError("SYMLINK_REFUSED")
    return path


def exclusive(path, data, mode=0o644):
    descriptor = os.open(str(path), os.O_CREAT | os.O_EXCL | os.O_WRONLY | getattr(os, "O_NOFOLLOW", 0), mode)
    with os.fdopen(descriptor, "wb") as handle:
        handle.write(data)


def install(prefix, bin_dir):
    if sys.version_info < (3, 9):
        raise InstallError("PYTHON_VERSION")
    source = Path(__file__).resolve().parent
    if os.name != "posix":
        raise InstallError("POSIX_INSTALLER_ONLY")
    prefix, bin_dir = safe_path(prefix), safe_path(bin_dir)
    if source == prefix or source in prefix.parents or source == bin_dir or source in bin_dir.parents:
        raise InstallError("INSTALL_INSIDE_SOURCE_REFUSED")
    if prefix == bin_dir or prefix in bin_dir.parents or bin_dir in prefix.parents:
        raise InstallError("OVERLAPPING_DESTINATIONS")
    if prefix.exists():
        raise InstallError("PREFIX_EXISTS_NO_OVERWRITE")
    if bin_dir.exists() and not bin_dir.is_dir():
        raise InstallError("BIN_NOT_DIRECTORY")
    binary = safe_path(bin_dir / NAME)
    if binary.exists():
        raise InstallError("COMMAND_EXISTS_NO_OVERWRITE")
    contents = {}
    for relative in FILES:
        path = safe_path(source / relative)
        if not path.is_file():
            raise InstallError("SOURCE_MISSING")
        contents[relative] = path.read_bytes()
    # A Python launcher uses the exact tested interpreter, no shell interpolation.
    wrapper = ("#!" + sys.executable + "\nimport runpy\n"
               + "runpy.run_path(" + repr(str(prefix / "presentation_kit.py")) + ", run_name='__main__')\n").encode("utf-8")
    if any(character.isspace() for character in sys.executable) or len(sys.executable) > 200:
        raise InstallError("INTERPRETER_PATH_UNSUPPORTED")
    prefix.parent.mkdir(parents=True, exist_ok=True)
    safe_path(prefix.parent)
    prefix.mkdir()
    try:
        for relative, data in contents.items():
            target = prefix / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            exclusive(target, data)
        bin_dir.mkdir(parents=True, exist_ok=True)
        safe_path(bin_dir)
        exclusive(binary, wrapper, 0o755)
        manifest = {"kind": "presentation-kit-install", "version": VERSION,
                    "prefix": str(prefix), "binary": str(binary),
                    "files": {relative: hashlib.sha256(data).hexdigest() for relative, data in contents.items()},
                    "binary_sha256": hashlib.sha256(wrapper).hexdigest()}
        exclusive(prefix / MARKER, (json.dumps(manifest, sort_keys=True, indent=2) + "\n").encode("utf-8"))
    except OSError:
        # Do not recursively remove an uncertain/racing installation. A precise error remains reviewable.
        raise InstallError("INSTALL_INCOMPLETE_NO_RETRY")
    return {"ok": True, "command": "install", "name": NAME, "version": VERSION,
            "installed_files": len(FILES), "binary": str(binary), "prefix": str(prefix),
            "network_used": False, "overwrite_supported": False}


class Parser(argparse.ArgumentParser):
    def error(self, message):
        raise InstallError("CLI_USAGE")


def main(argv=None):
    try:
        parser = Parser(description="Install a reviewed offline CLI into a new dedicated prefix.")
        parser.add_argument("--prefix", required=True)
        parser.add_argument("--bin-dir", required=True)
        args = parser.parse_args(argv)
        print(json.dumps(install(args.prefix, args.bin_dir), ensure_ascii=True, sort_keys=True))
        return 0
    except (InstallError, OSError, ValueError) as error:
        code = str(error) if isinstance(error, InstallError) else "INSTALL_IO_ERROR"
        print(json.dumps({"ok": False, "error": {"code": code, "message": "Installation refused or incomplete; no existing files are overwritten."}}, sort_keys=True))
        return 2


if __name__ == "__main__":
    sys.exit(main())
