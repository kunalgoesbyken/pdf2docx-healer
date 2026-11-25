import pymupdf
import pytest
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docx_healer import HealerConfig, heal
from docx_healer.link_safety import is_safe_uri, unlink_unsafe_targets

SAFE = [
    "http://example.com",
    "https://example.com/a?b=1#c",
    "HTTPS://EXAMPLE.COM",
    "mailto:ops@example.com",
    "ftp://files.example.com/x",
    "  https://example.com  ",
]
UNSAFE = [
    "file:///C:/Windows/System32/calc.exe",
    "FILE://server/share",
    "\\\\attacker\\share\\x.docx",
    "//attacker/share",
    "javascript:alert(1)",
    "smb://attacker/share",
    "ms-msdt:/id PCWDiagnostic",
    "fi\tle:///etc/passwd",
    "\x00file:///etc/passwd",
    "page2.html",
    "www.example.com",
    "",
]


@pytest.mark.parametrize("uri", SAFE)
def test_web_and_mail_schemes_are_safe(uri):
    assert is_safe_uri(uri)


@pytest.mark.parametrize("uri", UNSAFE)
def test_other_targets_are_unsafe(uri):
    assert not is_safe_uri(uri)


def _docx_with_links(path, targets):
    doc = Document()
    for target in targets:
        paragraph = doc.add_paragraph("before ")
        link = OxmlElement("w:hyperlink")
        if target.startswith("#"):
            link.set(qn("w:anchor"), target[1:])
        else:
            link.set(qn("r:id"), paragraph.part.relate_to(target, RT.HYPERLINK, is_external=True))
        run = paragraph.add_run("label " + target)
        style = OxmlElement("w:rStyle")
        style.set(qn("w:val"), "Hyperlink")
        run._r.get_or_add_rPr().append(style)
        link.append(run._r)
        paragraph._p.append(link)
        paragraph.add_run(" after")
    doc.save(str(path))
    return path


def _links(path):
    doc = Document(str(path))
    found = []
    for link in doc.element.body.iter(qn("w:hyperlink")):
        rid = link.get(qn("r:id"))
        found.append(doc.part.rels[rid].target_ref if rid else "#" + link.get(qn("w:anchor")))
    return found


def test_unsafe_links_become_plain_text_and_safe_ones_stay(tmp_path):
    path = _docx_with_links(
        tmp_path / "in.docx",
        ["https://example.com", "file:///C:/x.exe", "#sec1", "\\\\host\\share", "mailto:a@b.co"],
    )

    removed = unlink_unsafe_targets(path)

    assert removed == 2
    assert _links(path) == ["https://example.com", "#sec1", "mailto:a@b.co"]
    doc = Document(str(path))
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "before label file:///C:/x.exe after" in text
    assert "before label \\\\host\\share after" in text
    targets = [rel.target_ref for rel in doc.part.rels.values() if rel.is_external]
    assert "file:///C:/x.exe" not in targets
    assert "\\\\host\\share" not in targets
    for paragraph in doc.paragraphs:
        if "x.exe" in paragraph.text:
            assert not list(paragraph._p.iter(qn("w:rStyle")))


def _nested_link_docx(path, target):
    doc = Document()
    paragraph = doc.add_paragraph("start ")
    outer = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    rpr.append(OxmlElement("w:b"))
    outer.append(rpr)
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), paragraph.part.relate_to(target, RT.HYPERLINK, is_external=True))
    inner = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "linked"
    inner.append(text)
    link.append(inner)
    outer.append(link)
    tail = OxmlElement("w:t")
    tail.text = " tail"
    outer.append(tail)
    paragraph._p.append(outer)
    paragraph.add_run(" end")
    doc.save(str(path))
    return path


@pytest.mark.parametrize("target,kept", [("https://example.com", True), ("javascript:x", False)])
def test_links_nested_in_runs_are_lifted_in_order(tmp_path, target, kept):
    path = _nested_link_docx(tmp_path / "in.docx", target)

    unlink_unsafe_targets(path)

    doc = Document(str(path))
    paragraph = doc.paragraphs[0]
    assert paragraph.text == "start linked tail end"
    assert not [r for r in paragraph._p.iter(qn("w:r")) if r.getparent().tag == qn("w:r")]
    assert all(link.getparent() is paragraph._p for link in paragraph._p.iter(qn("w:hyperlink")))
    assert _links(path) == ([target] if kept else [])
    bold_tail = [r for r in paragraph._p.iter(qn("w:r")) if "tail" in "".join(r.itertext())]
    assert bold_tail[0].find(f"{qn('w:rPr')}/{qn('w:b')}") is not None


def test_document_without_unsafe_links_is_left_alone(tmp_path):
    path = _docx_with_links(tmp_path / "in.docx", ["https://example.com", "#a"])
    before = path.read_bytes()

    assert unlink_unsafe_targets(path) == 0
    assert path.read_bytes() == before


def _pdf_with_link(path, uri):
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_text((72, 100), "Click the report link here for the details.", fontsize=12)
    page.insert_text((72, 140), "Read more at the public site link as well.", fontsize=12)
    page.insert_link({"kind": pymupdf.LINK_URI, "from": pymupdf.Rect(70, 88, 300, 104), "uri": uri})
    page.insert_link({
        "kind": pymupdf.LINK_URI, "from": pymupdf.Rect(70, 128, 300, 144),
        "uri": "https://example.com/",
    })
    doc.save(str(path))
    return path


@pytest.mark.parametrize("mode", ["fidelity", "editable"])
@pytest.mark.parametrize(
    "uri",
    [
        "ms-msdt:/id PCWDiagnostic",
        "javascript:alert(1)",
        "file:///C:/Windows/calc.exe",
        "\\\\attacker\\share\\a",
    ],
)
def test_heal_never_writes_unsafe_link_targets(tmp_path, mode, uri):
    pdf = _pdf_with_link(tmp_path / "in.pdf", uri)

    heal(pdf, tmp_path / "out.docx", HealerConfig(verbose=False, mode=mode))

    doc = Document(str(tmp_path / "out.docx"))
    targets = [rel.target_ref for rel in doc.part.rels.values() if rel.is_external]
    assert uri not in targets
    assert "report link" in "\n".join(p.text for p in doc.paragraphs)


@pytest.mark.parametrize("mode", ["fidelity", "editable"])
def test_heal_keeps_unsafe_link_removal_with_hyperlink_pass_off(tmp_path, mode):
    pdf = _pdf_with_link(tmp_path / "in.pdf", "ms-msdt:/id PCWDiagnostic")

    heal(pdf, tmp_path / "out.docx", HealerConfig(verbose=False, mode=mode, fix_hyperlinks=False))

    doc = Document(str(tmp_path / "out.docx"))
    targets = [rel.target_ref for rel in doc.part.rels.values() if rel.is_external]
    assert "ms-msdt:/id PCWDiagnostic" not in targets
    assert "https://example.com/" in targets
