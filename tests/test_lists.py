import importlib.util

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

from docx_healer.list_detector import _Numbering, fix_lists_in_docx
from docx_healer.list_markers import BULLET, LOWER_ROMAN
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

    _Numbering(doc).numbered((LOWER_ROMAN, "."), 4)

    abstract = doc.part.numbering_part.element.findall(qn("w:abstractNum"))[-1]
    levels = abstract.findall(qn("w:lvl"))
    assert len(levels) == 9
    assert all(local_names(lvl) == LEVEL_CHILD_ORDER for lvl in levels)
    assert [lvl.find(qn("w:lvlText")).get(qn("w:val")) for lvl in levels[:3]] == ["%1.", "%2.", "%3."]
    assert {lvl.find(qn("w:start")).get(qn("w:val")) for lvl in levels} == {"4"}


def test_bracket_styles_produce_matching_level_text(tmp_path):
    path = make_docx(tmp_path / "a.docx", lambda doc: doc.add_paragraph("x"))
    doc = Document(str(path))
    numbering = _Numbering(doc)

    for shape in (")", "()"):
        numbering.numbered(("decimal", shape), 1)

    texts = [
        abstract.findall(qn("w:lvl"))[0].find(qn("w:lvlText")).get(qn("w:val"))
        for abstract in doc.part.numbering_part.element.findall(qn("w:abstractNum"))[-2:]
    ]
    assert texts == ["%1)", "(%1)"]


def test_bullet_definition_puts_run_properties_after_paragraph_properties(tmp_path):
    path = make_docx(tmp_path / "a.docx", lambda doc: doc.add_paragraph("x"))
    doc = Document(str(path))

    _Numbering(doc).bullet()

    abstract = doc.part.numbering_part.element.findall(qn("w:abstractNum"))[-1]
    assert all(local_names(lvl) == LEVEL_CHILD_ORDER + ["rPr"] for lvl in abstract.findall(qn("w:lvl")))
    assert {lvl.find(qn("w:numFmt")).get(qn("w:val")) for lvl in abstract.findall(qn("w:lvl"))} == {BULLET}


def test_bullet_definition_is_created_once_per_document(tmp_path):
    path = make_docx(tmp_path / "a.docx", lambda doc: doc.add_paragraph("x"))
    doc = Document(str(path))
    numbering = _Numbering(doc)
    count = len(doc.part.numbering_part.element.findall(qn("w:abstractNum")))

    assert numbering.bullet() == numbering.bullet()

    assert len(doc.part.numbering_part.element.findall(qn("w:abstractNum"))) == count + 1


def test_numbering_reference_follows_paragraph_style_in_schema_order(tmp_path):
    def build(doc):
        for text in ("a. first item", "b. second item"):
            paragraph = doc.add_paragraph(text)
            paragraph.paragraph_format.keep_with_next = True

    path = make_docx(tmp_path / "a.docx", build)

    fix_lists_in_docx(path)

    names = local_names(Document(str(path)).paragraphs[0]._p.pPr)
    assert names.index("keepNext") < names.index("numPr")


def test_bullets_deeper_than_the_builtin_styles_use_the_shared_bullet_definition(tmp_path):
    def build(doc):
        for depth in range(5):
            paragraph = doc.add_paragraph("• item")
            paragraph.paragraph_format.left_indent = Pt(30 * depth)

    path = make_docx(tmp_path / "a.docx", build)

    assert fix_lists_in_docx(path) == 5

    paragraphs = Document(str(path)).paragraphs
    assert [p.style.name for p in paragraphs[:3]] == ["List Bullet", "List Bullet 2", "List Bullet 3"]
    deep = [p._p.pPr.numPr for p in paragraphs[3:]]
    assert [n.ilvl.val for n in deep] == [3, 4]
    assert len({n.numId.val for n in deep}) == 1
