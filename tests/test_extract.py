import struct
import zipfile
from pathlib import Path

import prism_ole_handler

import app

SAMPLE_DECK = Path(prism_ole_handler.__file__).parent / "test" / "test_01.pptx"


def test_sample_deck_gives_four_prism10_graphs(tmp_path):
    result = app.extract_document(SAMPLE_DECK, "ppt", tmp_path)
    assert [o["slide"] for o in result["objects"]] == [2, 3, 4, 5]
    assert result["skipped"] == []
    for o in result["objects"]:
        assert o["format"] == "prism" and o["progId"] == "Prism10.Document"
        assert zipfile.ZipFile(tmp_path / o["file"]).testzip() is None


def test_pzf_length_prefix_counts_itself():
    body = b"PCFFGRA4" + b"\x00" * 20 + b"\x01\x40"
    stream = struct.pack("<I", len(body) + 4) + body + b"leftover bytes"
    unwrapped = app.unwrap_length_prefix(stream)
    assert unwrapped == body
    assert app.classify_payload(unwrapped, None) == "pzf"


def test_non_office_file_reports_error(tmp_path):
    bad = tmp_path / "broken.pptx"
    bad.write_bytes(b"garbage")
    result = app.extract_document(bad, "ppt", tmp_path)
    assert "error" in result and result["objects"] == []
