"""Fixture-based tests for Word threaded comment replies."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from docx import Document
from lxml import etree

from tools.word_tools import DEFAULT_COMMENT_AUTHOR


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"


def _load_comments_root(docx_path: str | Path):
    with zipfile.ZipFile(docx_path, "r") as zf:
        xml = zf.read("word/comments.xml")
    return etree.fromstring(xml)


def _save_comments_root(docx_path: str | Path, root) -> None:
    with zipfile.ZipFile(docx_path, "r") as zf_in:
        items = zf_in.namelist()
        blobs = {name: zf_in.read(name) for name in items}

    blobs["word/comments.xml"] = etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone="yes",
    )

    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
        for name in items:
            zf_out.writestr(name, blobs[name])


@pytest.fixture
def word_doc_with_comment(temp_dir, word_advanced_tools, word_tools):
    """Create a DOCX with one base comment and return metadata."""
    path = temp_dir / "reply_base.docx"
    doc = Document()
    doc.add_paragraph("Please review this section.")
    doc.save(path)

    add = word_advanced_tools.tool_word_add_comment(
        file_path=str(path),
        target_text="review this section",
        comment_text="Initial reviewer comment",
        author="Reviewer",
    )
    assert add.get("success") is True

    comments = word_tools.tool_word_get_comments(str(path))
    assert comments.get("comment_count", 0) >= 1

    parent_id = comments["comments"][0]["id"]
    return {
        "path": str(path),
        "parent_id": str(parent_id),
    }


def test_reply_to_existing_comment(word_doc_with_comment, word_tools):
    """T1: Reply to existing comment should create linked reply comment."""
    file_path = word_doc_with_comment["path"]
    parent_id = word_doc_with_comment["parent_id"]

    res = word_tools.tool_word_reply_to_comment(
        file_path=file_path,
        comment_id=parent_id,
        text="Reply in thread",
        author="Rui Carmo",
    )
    assert res.get("success") is True
    assert res.get("parent_comment_id") == parent_id
    reply_id = res.get("reply_comment_id")
    assert reply_id is not None

    root = _load_comments_root(file_path)
    parent = root.find(f".//{{{W_NS}}}comment[@w:id='{parent_id}']", namespaces={"w": W_NS})
    reply = root.find(f".//{{{W_NS}}}comment[@w:id='{reply_id}']", namespaces={"w": W_NS})
    assert parent is not None
    assert reply is not None

    parent_para = parent.find(f"{{{W_NS}}}p")
    reply_para = reply.find(f"{{{W_NS}}}p")
    assert parent_para is not None
    assert reply_para is not None

    parent_para_id = parent_para.get(f"{{{W14_NS}}}paraId")
    assert parent_para_id
    assert reply_para.get(f"{{{W14_NS}}}paraIdParent") == parent_para_id


def test_reply_invalid_comment_id_returns_valid_ids(word_doc_with_comment, word_tools):
    """T2: Non-existent comment id should include valid IDs in error."""
    file_path = word_doc_with_comment["path"]

    res = word_tools.tool_word_reply_to_comment(
        file_path=file_path,
        comment_id="999",
        text="No-op",
    )
    assert "error" in res
    assert "Valid IDs" in res["error"]
    assert isinstance(res.get("valid_comment_ids"), list)


def test_reply_preserves_existing_replies(word_doc_with_comment, word_tools):
    """T3: Multiple replies should be preserved and appended."""
    file_path = word_doc_with_comment["path"]
    parent_id = word_doc_with_comment["parent_id"]

    r1 = word_tools.tool_word_reply_to_comment(file_path=file_path, comment_id=parent_id, text="First reply")
    r2 = word_tools.tool_word_reply_to_comment(file_path=file_path, comment_id=parent_id, text="Second reply")
    assert r1.get("success") is True
    assert r2.get("success") is True
    assert r1.get("reply_comment_id") != r2.get("reply_comment_id")

    comments = word_tools.tool_word_get_comments(file_path)
    texts = [c.get("text") for c in comments.get("comments", [])]
    assert "First reply" in texts
    assert "Second reply" in texts


def test_reply_adds_parent_paraid_when_missing(word_doc_with_comment, word_tools):
    """T4: Parent without w14:paraId gets one synthesized before linking reply."""
    file_path = word_doc_with_comment["path"]
    parent_id = word_doc_with_comment["parent_id"]

    root = _load_comments_root(file_path)
    parent = root.find(f".//{{{W_NS}}}comment[@w:id='{parent_id}']", namespaces={"w": W_NS})
    assert parent is not None
    parent_para = parent.find(f"{{{W_NS}}}p")
    assert parent_para is not None

    # Simulate old format lacking paraId.
    parent_para.attrib.pop(f"{{{W14_NS}}}paraId", None)
    _save_comments_root(file_path, root)

    res = word_tools.tool_word_reply_to_comment(
        file_path=file_path,
        comment_id=parent_id,
        text="Reply after synthetic paraId",
    )
    assert res.get("success") is True

    root_after = _load_comments_root(file_path)
    parent_after = root_after.find(f".//{{{W_NS}}}comment[@w:id='{parent_id}']", namespaces={"w": W_NS})
    reply_after = root_after.find(
        f".//{{{W_NS}}}comment[@w:id='{res.get('reply_comment_id')}']",
        namespaces={"w": W_NS},
    )
    parent_para_after = parent_after.find(f"{{{W_NS}}}p")
    reply_para_after = reply_after.find(f"{{{W_NS}}}p")

    new_parent_para_id = parent_para_after.get(f"{{{W14_NS}}}paraId")
    assert new_parent_para_id
    assert reply_para_after.get(f"{{{W14_NS}}}paraIdParent") == new_parent_para_id


def test_multiple_replies_have_unique_ids_and_paraids(word_doc_with_comment, word_tools):
    """T5: Sequential replies must keep unique comment IDs and para IDs."""
    file_path = word_doc_with_comment["path"]
    parent_id = word_doc_with_comment["parent_id"]

    ids = []
    for i in range(3):
        res = word_tools.tool_word_reply_to_comment(
            file_path=file_path,
            comment_id=parent_id,
            text=f"Reply #{i+1}",
        )
        assert res.get("success") is True
        ids.append(res.get("reply_comment_id"))

    assert len(ids) == len(set(ids))

    root = _load_comments_root(file_path)
    para_ids = []
    for cid in ids:
        node = root.find(f".//{{{W_NS}}}comment[@w:id='{cid}']", namespaces={"w": W_NS})
        para = node.find(f"{{{W_NS}}}p")
        para_ids.append(para.get(f"{{{W14_NS}}}paraId"))

    assert len(para_ids) == len(set(para_ids))


def test_reply_author_fallback_from_identity(word_doc_with_comment, word_tools):
    """T6: Missing author falls back to runtime identity/default."""
    file_path = word_doc_with_comment["path"]
    parent_id = word_doc_with_comment["parent_id"]

    word_tools._comment_author = "Fixture Reviewer"

    res = word_tools.tool_word_reply_to_comment(
        file_path=file_path,
        comment_id=parent_id,
        text="Author fallback reply",
    )
    assert res.get("success") is True
    assert res.get("author") in {"Fixture Reviewer", DEFAULT_COMMENT_AUTHOR}


def test_reply_to_output_path_leaves_source_unchanged(word_doc_with_comment, word_tools, temp_dir):
    """T7: output_path writes reply to new file and preserves original."""
    source = word_doc_with_comment["path"]
    parent_id = word_doc_with_comment["parent_id"]
    output = temp_dir / "reply_out.docx"

    before = word_tools.tool_word_get_comments(source)
    before_count = before.get("comment_count", 0)

    res = word_tools.tool_word_reply_to_comment(
        file_path=source,
        comment_id=parent_id,
        text="Reply in output copy",
        output_path=str(output),
    )
    assert res.get("success") is True
    assert output.exists()

    after_source = word_tools.tool_word_get_comments(source)
    assert after_source.get("comment_count", 0) == before_count

    after_output = word_tools.tool_word_get_comments(str(output))
    assert after_output.get("comment_count", 0) == before_count + 1
