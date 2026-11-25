import zipfile
from pathlib import Path
from typing import List, Optional, Tuple

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NAMESPACES = (
    f'xmlns:w="{W_NS}" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:v="urn:schemas-microsoft-com:vml"'
)
REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CT_BASE = "application/vnd.openxmlformats-officedocument.wordprocessingml."

Rel = Tuple[str, str, str, bool]


def numbering_xml(abstracts: str, nums: str) -> str:
    return f"<w:numbering {NAMESPACES}>{abstracts}{nums}</w:numbering>"


def lvl(
    ilvl: int,
    fmt: Optional[str] = "decimal",
    text: str = "%1.",
    start: Optional[int] = 1,
    extra: str = "",
) -> str:
    parts = [f'<w:lvl w:ilvl="{ilvl}">']
    if start is not None:
        parts.append(f'<w:start w:val="{start}"/>')
    if fmt is not None:
        parts.append(f'<w:numFmt w:val="{fmt}"/>')
    parts.append(extra)
    parts.append(f'<w:lvlText w:val="{text}"/></w:lvl>')
    return "".join(parts)


def abstract(abstract_id: int, levels: str) -> str:
    return f'<w:abstractNum w:abstractNumId="{abstract_id}">{levels}</w:abstractNum>'


def num(num_id: int, abstract_id: int, overrides: str = "") -> str:
    return f'<w:num w:numId="{num_id}"><w:abstractNumId w:val="{abstract_id}"/>{overrides}</w:num>'


def start_override(ilvl: int, value: int) -> str:
    return f'<w:lvlOverride w:ilvl="{ilvl}"><w:startOverride w:val="{value}"/></w:lvlOverride>'


def para(
    text: str = "",
    num_id: Optional[int] = None,
    ilvl: int = 0,
    style: Optional[str] = None,
    extra_ppr: str = "",
) -> str:
    ppr = ""
    if style:
        ppr += f'<w:pStyle w:val="{style}"/>'
    if num_id is not None:
        ppr += f'<w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="{num_id}"/></w:numPr>'
    ppr += extra_ppr
    ppr = f"<w:pPr>{ppr}</w:pPr>" if ppr else ""
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f"<w:p>{ppr}{run}</w:p>"


def styles_xml(styles: str) -> str:
    return f"<w:styles {NAMESPACES}>{styles}</w:styles>"


def style(style_id: str, name: str, based_on: Optional[str] = None, ppr: str = "") -> str:
    base = f'<w:basedOn w:val="{based_on}"/>' if based_on else ""
    ppr_xml = f"<w:pPr>{ppr}</w:pPr>" if ppr else ""
    return (
        f'<w:style w:type="paragraph" w:styleId="{style_id}"><w:name w:val="{name}"/>'
        f"{base}{ppr_xml}</w:style>"
    )


def build_docx(
    path: Path,
    body: str = "",
    numbering: Optional[str] = None,
    styles: Optional[str] = None,
    header: Optional[str] = None,
    rels: Optional[List[Rel]] = None,
) -> Path:
    document = f"<w:document {NAMESPACES}><w:body>{body}</w:body></w:document>"
    content_types = [
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        f'<Override PartName="/word/document.xml" ContentType="{CT_BASE}document.main+xml"/>',
    ]
    doc_rels: List[str] = []
    parts = {}
    if numbering is not None:
        parts["word/numbering.xml"] = numbering
        content_types.append(
            f'<Override PartName="/word/numbering.xml" ContentType="{CT_BASE}numbering+xml"/>'
        )
        doc_rels.append(
            f'<Relationship Id="rIdN" Type="{REL_BASE}numbering" Target="numbering.xml"/>'
        )
    if styles is not None:
        parts["word/styles.xml"] = styles
        content_types.append(
            f'<Override PartName="/word/styles.xml" ContentType="{CT_BASE}styles+xml"/>'
        )
        doc_rels.append(f'<Relationship Id="rIdS" Type="{REL_BASE}styles" Target="styles.xml"/>')
    if header is not None:
        parts["word/header1.xml"] = f"<w:hdr {NAMESPACES}>{header}</w:hdr>"
        content_types.append(
            f'<Override PartName="/word/header1.xml" ContentType="{CT_BASE}header+xml"/>'
        )
        doc_rels.append(f'<Relationship Id="rIdH" Type="{REL_BASE}header" Target="header1.xml"/>')
    for rel_id, kind, target, external in rels or []:
        mode = ' TargetMode="External"' if external else ""
        doc_rels.append(
            f'<Relationship Id="{rel_id}" Type="{REL_BASE}{kind}" Target="{target}"{mode}/>'
        )
    types_xml = (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        + "".join(content_types)
        + "</Types>"
    )
    root_rels = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{REL_BASE}officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    )
    document_rels = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(doc_rels)
        + "</Relationships>"
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", types_xml)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("word/document.xml", document)
        archive.writestr("word/_rels/document.xml.rels", document_rels)
        for name, content in parts.items():
            archive.writestr(name, content)
    return path
