import importlib.util

from docx import Document
from docx.oxml.ns import qn

from docx_healer.list_detector import (
    _apply_list_via_xml,
    detect_number_format,
    fix_lists_in_docx,
)
from factories import make_docx

LEVEL_CHILD_ORDER = ["start", "numFmt", "lvlText", "lvlJc", "pPr"]


def local_names(element):
    return [child.tag.split("}")[1] for child in element]


def test_numbering_module_is_gone():
    assert importlib.util.find_spec("docx_healer.numbering") is None


def test_lettered_list_gets_one_shared_lower_letter_definition(tmp_path):
    def build(doc):
        for text in ("a. first item here", "b. second item here", "c. third item here"):
            doc.add_paragraph(text)

    path = make_docx(tmp_path / "a.docx", build)

    assert fix_lists_in_docx(path) == 3

    doc = Document(str(path))
    paragraphs = [p for p in doc.paragraphs if p.text.strip()]
    assert [p.text for p in paragraphs] == ["first item here", "second item here", "third item here"]
    num_ids = {p._p.pPr.numPr.numId.val for p in paragraphs}
    assert len(num_ids) == 1

    numbering = doc.part.numbering_part.element
    abstract = numbering.findall(qn("w:abstractNum"))[-1]
    assert {lvl.find(qn("w:numFmt")).get(qn("w:val")) for lvl in abstract.findall(qn("w:lvl"))} == {
        "lowerLetter"
    }


def test_abstract_nums_always_precede_nums(tmp_path):
    def build(doc):
        for text in ("a. first item here", "b. second item here"):
            doc.add_paragraph(text)
        doc.add_paragraph("A normal paragraph between the two lists.")
        for text in ("i. roman one here", "ii. roman two here"):
            doc.add_paragraph(text)

    path = make_docx(tmp_path / "a.docx", build)

    fix_lists_in_docx(path)

    numbering = Document(str(path)).part.numbering_part.element
    kinds = [child.tag.split("}")[1] for child in numbering]
    last_abstract = max(i for i, kind in enumerate(kinds) if kind == "abstractNum")
    first_num = min(i for i, kind in enumerate(kinds) if kind == "num")
    assert last_abstract < first_num


def test_injected_definition_covers_nine_levels_in_schema_order(tmp_path):
    path = make_docx(tmp_path / "a.docx", lambda doc: doc.add_paragraph("x"))
    doc = Document(str(path))
    paragraph = doc.paragraphs[0]

    _apply_list_via_xml(paragraph, "number", level=2, num_fmt="lowerRoman")

    abstract = doc.part.numbering_part.element.findall(qn("w:abstractNum"))[-1]
    levels = abstract.findall(qn("w:lvl"))
    assert len(levels) == 9
    assert all(local_names(lvl) == LEVEL_CHILD_ORDER for lvl in levels)
    assert [lvl.find(qn("w:lvlText")).get(qn("w:val")) for lvl in levels[:3]] == ["%1.", "%2.", "%3."]


def test_bullet_definition_puts_run_properties_after_paragraph_properties(tmp_path):
    path = make_docx(tmp_path / "a.docx", lambda doc: doc.add_paragraph("x"))
    doc = Document(str(path))

    _apply_list_via_xml(doc.paragraphs[0], "bullet")

    abstract = doc.part.numbering_part.element.findall(qn("w:abstractNum"))[-1]
    assert all(local_names(lvl) == LEVEL_CHILD_ORDER + ["rPr"] for lvl in abstract.findall(qn("w:lvl")))


def test_numbering_reference_follows_paragraph_style_in_schema_order(tmp_path):
    def build(doc):
        paragraph = doc.add_paragraph("x")
        paragraph.paragraph_format.keep_with_next = True
        paragraph.paragraph_format.space_after = None

    path = make_docx(tmp_path / "a.docx", build)
    doc = Document(str(path))

    _apply_list_via_xml(doc.paragraphs[0], "bullet", level=1)

    names = local_names(doc.paragraphs[0]._p.pPr)
    assert names.index("keepNext") < names.index("numPr")


def test_reusing_a_num_id_does_not_add_definitions(tmp_path):
    path = make_docx(tmp_path / "a.docx", lambda doc: [doc.add_paragraph("x"), doc.add_paragraph("y")])
    doc = Document(str(path))
    numbering = doc.part.numbering_part.element
    first, second = doc.paragraphs[:2]

    num_id = _apply_list_via_xml(first, "bullet")
    count = len(numbering.findall(qn("w:abstractNum")))
    assert _apply_list_via_xml(second, "bullet", level=1, _num_id=num_id) == num_id

    assert len(numbering.findall(qn("w:abstractNum"))) == count
    assert second._p.pPr.numPr.ilvl.val == 1


def test_detect_number_format():
    assert detect_number_format("a.") == "lowerLetter"
    assert detect_number_format("B)") == "upperLetter"
    assert detect_number_format("iv.") == "lowerRoman"
    assert detect_number_format("1.1.") == "decimal"
    assert detect_number_format("3.") == "decimal"
