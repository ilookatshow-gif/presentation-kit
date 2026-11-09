"""Behavior, artifact structure and file-boundary checks on invented inputs."""
import copy
import hashlib
import importlib.util
import io
import json
import os
from pathlib import Path
import shutil
import socket
import subprocess
import sys
import tempfile
import unittest
from unittest import mock
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("presentation_kit", ROOT / "presentation_kit.py")
KIT = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(KIT)


def change_part(raw, name, change):
    with zipfile.ZipFile(io.BytesIO(raw)) as archive:
        parts = {info.filename: archive.read(info) for info in archive.infolist()}
    parts[name] = change(parts[name])
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_STORED) as archive:
        for path, data in parts.items():
            archive.writestr(path, data)
    return output.getvalue()


class PresentationTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(prefix="presentation-tests-")
        self.temp = Path(self.directory.name).resolve()
        self.deck = json.loads((ROOT / "examples/synthetic-deck.json").read_text(encoding="utf-8"))
        self.source = self.temp / "deck.json"
        self.output = self.temp / "deck.pptx"
        self.save()

    def tearDown(self):
        self.directory.cleanup()

    def save(self, value=None):
        self.source.write_text(json.dumps(self.deck if value is None else value, ensure_ascii=False), encoding="utf-8")

    def error(self, code, function, *args, **kwargs):
        with self.assertRaises(KIT.KitError) as raised:
            function(*args, **kwargs)
        self.assertEqual(raised.exception.code, code)

    def build(self):
        result = KIT.build(self.source, self.output)
        self.assertTrue(result["ok"])
        return self.output.read_bytes()

    def test_demo_contains_five_slides_all_four_layouts(self):
        result = KIT.demo(self.output)
        self.assertEqual(result["slides"], 5)
        self.assertTrue(result["synthetic"])
        report = KIT.inspect(self.output)
        self.assertEqual({s["layout"] for s in report["outline"]}, set(KIT.LAYOUTS))

    def test_pptx_retains_input_text_as_native_objects(self):
        raw = self.build()
        with zipfile.ZipFile(io.BytesIO(raw)) as archive:
            for number, slide in enumerate(self.deck["slides"], 1):
                root = ET.fromstring(archive.read("ppt/slides/slide" + str(number) + ".xml"))
                actual = [x.text for x in root.findall(".//a:t", KIT.NS)]
                self.assertIn(slide["title"], actual)
                expected = []
                if slide["layout"] == "cover":
                    expected = [slide["subtitle"]]
                elif slide["layout"] == "bullets":
                    expected = slide["bullets"]
                elif slide["layout"] == "two_columns":
                    for side in ("left", "right"):
                        expected += [slide[side]["heading"]] + slide[side]["bullets"]
                else:
                    expected = slide["columns"] + [v for row in slide["rows"] for v in row]
                    self.assertEqual(len(root.findall(".//a:tc", KIT.NS)), 12)
                    self.assertEqual(root.find(".//a:graphicData", KIT.NS).get("uri"), "http://schemas.openxmlformats.org/drawingml/2006/table")
                for value in expected:
                    self.assertIn(value, actual)
                self.assertIn(self.deck["author"], actual)
            core = ET.fromstring(archive.read("docProps/core.xml"))
            self.assertEqual(core.findtext("dc:title", namespaces=KIT.NS), self.deck["title"])

    def test_no_raster_media_macros_external_assets(self):
        raw = self.build()
        with zipfile.ZipFile(io.BytesIO(raw)) as archive:
            self.assertFalse(any("media/" in n or "vba" in n.lower() or "embedding" in n.lower() for n in archive.namelist()))
            self.assertTrue(all(info.compress_type == zipfile.ZIP_STORED for info in archive.infolist()))
        self.assertEqual(KIT.inspect(self.output)["external_relationships"], 0)

    def test_opc_metadata_uses_importer_compatible_default_namespaces(self):
        raw = self.build()
        with zipfile.ZipFile(io.BytesIO(raw)) as archive:
            for name in archive.namelist():
                if name.endswith(".rels") or name in ("[Content_Types].xml", "docProps/app.xml"):
                    data = archive.read(name)
                    self.assertIn(b'xmlns="', data)
                    self.assertNotIn(b"ns0:", data)

    def test_pptx_is_deterministic_and_does_not_embed_local_path(self):
        first = self.build()
        second = self.temp / "other.pptx"
        KIT.build(self.source, second)
        self.assertEqual(first, second.read_bytes())
        self.assertNotIn(str(self.temp).encode(), first)
        with zipfile.ZipFile(io.BytesIO(first)) as archive:
            self.assertTrue(all(info.create_system == 3 for info in archive.infolist()))

    def test_xml_special_characters_are_data(self):
        phrase = 'A & B < C > D "quoted"'
        self.deck["slides"][1]["bullets"][0] = phrase
        self.save()
        self.build()
        self.assertIn(phrase, KIT.inspect(self.output)["outline"][1]["text"])

    def test_unicode_decoding_cp1252_stdout(self):
        env = dict(os.environ, PYTHONIOENCODING="cp1252", PYTHONDONTWRITEBYTECODE="1")
        cp = subprocess.run([sys.executable, "-B", str(ROOT / "presentation_kit.py"), "--json", "demo", "--output", str(self.output)], capture_output=True, env=env, cwd=self.temp)
        self.assertEqual(cp.returncode, 0, cp.stderr)
        self.assertTrue(json.loads(cp.stdout.decode("ascii"))["synthetic"])
        cp = subprocess.run([sys.executable, "-B", str(ROOT / "presentation_kit.py"), "--json", "inspect", str(self.output)], capture_output=True, env=env, cwd=self.temp)
        self.assertEqual(cp.returncode, 0, cp.stderr)
        self.assertIn(self.deck["slides"][0]["title"], json.loads(cp.stdout.decode("ascii"))["outline"][0]["text"])

    def test_network_is_not_used(self):
        with mock.patch.object(socket, "socket", side_effect=AssertionError("network not permitted")):
            self.build()
            self.assertEqual(KIT.inspect(self.output)["slides"], 5)

    def test_schema_and_doctor_discovery(self):
        self.assertEqual(KIT.schema()["layouts"]["table"], ["layout", "title", "columns", "rows"])
        cp = subprocess.run([sys.executable, "-B", str(ROOT / "presentation_kit.py"), "--json", "doctor"], capture_output=True, cwd=self.temp)
        self.assertEqual(cp.returncode, 0)
        self.assertEqual(json.loads(cp.stdout)["auth"], "not_required")

    def test_invalid_cli_json_without_private_path(self):
        cp = subprocess.run([sys.executable, "-B", str(ROOT / "presentation_kit.py"), "build", str(self.source)], capture_output=True)
        self.assertEqual(cp.returncode, 2)
        result = json.loads(cp.stdout)
        self.assertEqual(result["error"]["code"], "CLI_USAGE")
        self.assertNotIn(str(self.temp).encode(), cp.stdout)
        self.assertNotIn(b"Traceback", cp.stderr)

    def test_duplicate_json_keys_refused_before_output(self):
        self.source.write_text('{"title":"one","title":"two","slides":[]}', encoding="utf-8")
        self.error("INPUT_INVALID", KIT.build, self.source, self.output)
        self.assertFalse(self.output.exists())

    def test_invalid_utf8_refused(self):
        self.source.write_bytes(b"\xff\xfe")
        self.error("INPUT_INVALID", KIT.build, self.source, self.output)

    def test_input_size_limit(self):
        self.source.write_bytes(b" " * (KIT.MAX_INPUT + 1))
        self.error("INPUT_TOO_LARGE", KIT.build, self.source, self.output)

    def test_unknown_and_missing_top_fields(self):
        self.error("INPUT_INVALID", KIT.validate_deck, dict(self.deck, secret="x"))
        self.error("INPUT_INVALID", KIT.validate_deck, {"title": "only"})

    def test_empty_and_excessive_slides(self):
        self.error("INPUT_INVALID", KIT.validate_deck, dict(self.deck, slides=[]))
        self.error("INPUT_INVALID", KIT.validate_deck, dict(self.deck, slides=[self.deck["slides"][0]] * 31))

    def test_unknown_layout_and_fields(self):
        self.deck["slides"][0]["layout"] = "remote_image"
        self.error("INPUT_INVALID", KIT.validate_deck, self.deck)
        self.deck["slides"][0]["layout"] = "cover"
        self.deck["slides"][0]["remote_url"] = "https://example.invalid"
        self.error("INPUT_INVALID", KIT.validate_deck, self.deck)

    def test_text_control_characters_and_nontext_refused(self):
        self.deck["slides"][0]["title"] = "bad\x00text"
        self.error("INPUT_INVALID", KIT.validate_deck, self.deck)
        self.deck["slides"][0]["title"] = 12
        self.error("INPUT_INVALID", KIT.validate_deck, self.deck)

    def test_long_title_and_wide_author_refused(self):
        self.deck["slides"][0]["title"] = "x" * 71
        self.error("INPUT_INVALID", KIT.validate_deck, self.deck)
        self.deck["slides"][0]["title"] = "Valid title"
        self.deck["author"] = "W" * 100
        self.error("TEXT_OVERFLOW", KIT.validate_deck, self.deck)

    def test_bullets_must_fit_without_silent_truncation(self):
        self.deck["slides"][1]["bullets"] = ["word " * 30] * 6
        self.error("TEXT_OVERFLOW", KIT.validate_deck, self.deck)

    def test_two_columns_require_heading_and_bullets(self):
        self.deck["slides"][2]["left"] = {"heading": "Left"}
        self.error("INPUT_INVALID", KIT.validate_deck, self.deck)

    def test_table_requires_rectangular_bounded_text(self):
        self.deck["slides"][3]["rows"][0].pop()
        self.error("INPUT_INVALID", KIT.validate_deck, self.deck)
        self.deck["slides"][3]["rows"][0].append(15)
        self.error("INPUT_INVALID", KIT.validate_deck, self.deck)

    def test_no_partial_output_and_existing_output_preserved_on_invalid(self):
        self.output.write_bytes(b"existing")
        self.deck["slides"] = []
        self.save()
        self.error("INPUT_INVALID", KIT.build, self.source, self.output, True)
        self.assertEqual(self.output.read_bytes(), b"existing")

    def test_create_does_not_overwrite_existing_file(self):
        self.output.write_bytes(b"existing")
        self.error("OUTPUT_EXISTS", KIT.build, self.source, self.output)
        self.assertEqual(self.output.read_bytes(), b"existing")

    def test_explicit_replace_produces_the_same_bytes(self):
        first = self.build()
        KIT.build(self.source, self.output, True)
        self.assertEqual(first, self.output.read_bytes())

    def test_output_hardlink_does_not_replace_input(self):
        try:
            os.link(str(self.source), str(self.output))
        except OSError:
            self.skipTest("Host does not permit test hardlinks")
        original = self.source.read_bytes()
        self.error("OUTPUT_COLLISION", KIT.build, self.source, self.output, True)
        self.assertEqual(self.source.read_bytes(), original)

    @unittest.skipUnless(os.name == "posix", "POSIX symlink case")
    def test_symlink_source_and_output_refused(self):
        alias = self.temp / "alias.json"
        alias.symlink_to(self.source)
        self.error("UNSAFE_PATH", KIT.build, alias, self.output)
        self.output.symlink_to(self.source)
        self.error("UNSAFE_PATH", KIT.build, self.source, self.output, True)

    def test_missing_output_parent_refused(self):
        self.error("UNSAFE_PATH", KIT.build, self.source, self.temp / "missing" / "out.pptx")

    def test_inspect_rejects_arbitrary_zip_and_extra_part(self):
        raw = self.build()
        stream = io.BytesIO()
        with zipfile.ZipFile(stream, "w") as archive:
            archive.writestr("../escape.xml", "<safe/>")
        self.error("INVALID_PPTX", KIT.inspect_bytes, stream.getvalue())
        extra = io.BytesIO(raw)
        with zipfile.ZipFile(extra, "a") as archive:
            archive.writestr("extra.xml", "<extra/>")
        self.error("INVALID_PPTX", KIT.inspect_bytes, extra.getvalue())

    def test_inspect_rejects_duplicate_part(self):
        raw = self.build()
        stream = io.BytesIO(raw)
        with zipfile.ZipFile(stream, "a") as archive:
            archive.writestr("docProps/app.xml", "<extra/>")
        self.error("INVALID_PPTX", KIT.inspect_bytes, stream.getvalue())

    def test_inspect_rejects_macro_mime(self):
        raw = self.build()
        bad = change_part(raw, "[Content_Types].xml", lambda data: data.replace(KIT.MIME["presentation"].encode(), b"application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"))
        self.error("INVALID_PPTX", KIT.inspect_bytes, bad)

    def test_inspect_rejects_duplicate_content_type(self):
        raw = self.build()
        def mutation(data):
            node = ET.fromstring(data)
            node.append(copy.deepcopy(node[0]))
            return ET.tostring(node, encoding="utf-8")
        self.error("INVALID_PPTX", KIT.inspect_bytes, change_part(raw, "[Content_Types].xml", mutation))

    def test_inspect_rejects_external_and_missing_internal_relationship(self):
        raw = self.build()
        def external(data):
            root = ET.fromstring(data)
            root[0].set("TargetMode", "External")
            root[0].set("Target", "https://example.invalid")
            return ET.tostring(root, encoding="utf-8")
        self.error("INVALID_PPTX", KIT.inspect_bytes, change_part(raw, "_rels/.rels", external))
        def missing(data):
            root = ET.fromstring(data)
            root.remove(root[0])
            return ET.tostring(root, encoding="utf-8")
        self.error("INVALID_PPTX", KIT.inspect_bytes, change_part(raw, "ppt/_rels/presentation.xml.rels", missing))

    def test_inspect_rejects_slide_target_alias(self):
        raw = self.build()
        def mutation(data):
            root = ET.fromstring(data)
            root[3].set("Target", root[2].get("Target"))
            return ET.tostring(root, encoding="utf-8")
        self.error("INVALID_PPTX", KIT.inspect_bytes, change_part(raw, "ppt/_rels/presentation.xml.rels", mutation))

    def test_inspect_rejects_utf16_dtd_and_invalid_table_uri(self):
        raw = self.build()
        dtd = '<?xml version="1.0" encoding="utf-16"?><!DOCTYPE x [<!ENTITY y "bad">]><x>&y;</x>'
        self.error("INVALID_PPTX", KIT.inspect_bytes, change_part(raw, "docProps/core.xml", lambda data: dtd.encode("utf-16")))
        self.error("INVALID_PPTX", KIT.inspect_bytes, change_part(raw, "ppt/slides/slide4.xml", lambda data: data.replace(b"drawingml/2006/table", b"drawingml/2006/main/table")))


@unittest.skipUnless(os.name == "posix", "Native installer is explicitly POSIX only")
class InstallationTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(prefix="presentation-install-")
        self.temp = Path(self.directory.name).resolve()
        self.prefix, self.bin = self.temp / "installed", self.temp / "bin"
        self.env = dict(os.environ, PYTHONDONTWRITEBYTECODE="1", PYTHONIOENCODING="cp1252")

    def tearDown(self):
        self.directory.cleanup()

    def install(self):
        cp = subprocess.run([sys.executable, "-B", str(ROOT / "install.py"), "--prefix", str(self.prefix), "--bin-dir", str(self.bin)], capture_output=True, env=self.env, cwd=self.temp)
        self.assertEqual(cp.returncode, 0, cp.stdout + cp.stderr)
        self.assertTrue(json.loads(cp.stdout)["ok"])

    def command(self, *args):
        return subprocess.run([str(self.bin / "presentation-kit")] + list(args), capture_output=True, env=self.env, cwd=self.temp)

    def remove(self):
        return subprocess.run([sys.executable, "-B", str(self.prefix / "uninstall.py"), "--prefix", str(self.prefix), "--bin-dir", str(self.bin)], capture_output=True, env=self.env, cwd=self.temp)

    def test_install_run_from_other_cwd_and_preserve_user_files(self):
        self.install()
        self.assertEqual(self.command("--json", "doctor").returncode, 0)
        output = self.temp / "installed-demo.pptx"
        self.assertEqual(self.command("--json", "demo", "--output", str(output)).returncode, 0)
        self.assertEqual(json.loads(self.command("--json", "inspect", str(output)).stdout)["slides"], 5)
        user = self.prefix / "keep.txt"
        user.write_text("user data", encoding="utf-8")
        self.assertEqual(self.remove().returncode, 0)
        self.assertEqual(user.read_text(encoding="utf-8"), "user data")
        self.assertTrue(output.exists())

    def test_install_never_overwrites_existing_command(self):
        self.bin.mkdir()
        binary = self.bin / "presentation-kit"
        binary.write_bytes(b"existing")
        cp = subprocess.run([sys.executable, "-B", str(ROOT / "install.py"), "--prefix", str(self.prefix), "--bin-dir", str(self.bin)], capture_output=True, env=self.env)
        self.assertEqual(cp.returncode, 2)
        self.assertEqual(binary.read_bytes(), b"existing")
        self.assertFalse(self.prefix.exists())

    def test_onboard_once_and_precise_uninstall(self):
        self.install()
        self.assertFalse(json.loads(self.command("--json", "onboard").stdout)["already_shown"])
        self.assertTrue(json.loads(self.command("--json", "onboard").stdout)["already_shown"])
        self.assertNotIn(b"lookatsoul", self.command("--json", "doctor").stdout)
        self.assertEqual(self.remove().returncode, 0)
        self.assertFalse((self.bin / "presentation-kit").exists())

    def test_modified_install_is_not_removed(self):
        self.install()
        target = self.prefix / "presentation_kit.py"
        target.write_bytes(target.read_bytes() + b"\n# user edit\n")
        self.assertEqual(self.remove().returncode, 2)
        self.assertTrue(target.exists())
        self.assertTrue((self.bin / "presentation-kit").exists())


if __name__ == "__main__":
    unittest.main()
