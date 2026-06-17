"""OOXML numbering XML builder for custom Word list definitions."""

import random
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def create_custom_bullet_definition(numbering_elm, bullet_char='\u2022',
                                     bullet_font='Symbol', levels=1,
                                     left_indents_cm=None, hanging_indents_cm=None):
    """Create a custom bullet abstract numbering definition and num instance.
    
    Args:
        numbering_elm: The <w:numbering> root element from document.part.numbering_part.
        bullet_char: The bullet character (e.g., '\u2022' for •).
        bullet_font: Font name for the bullet character.
        levels: Number of nesting levels (1-9).
        left_indents_cm: List of left indent values per level in cm.
        hanging_indents_cm: List of hanging indent values per level in cm.
    
    Returns:
        (abstract_num_id, num_id) tuple.
    """
    WML_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    
    if left_indents_cm is None:
        left_indents_cm = [1.27 * (i + 1) for i in range(levels)]
    if hanging_indents_cm is None:
        hanging_indents_cm = [0.63] * levels
    
    abstract_num_id = _next_abstract_num_id(numbering_elm)
    num_id = _next_num_id(numbering_elm)
    
    nsid_val = ''.join(random.choice('0123456789ABCDEF') for _ in range(8))
    
    abs_num = OxmlElement('w:abstractNum')
    abs_num.set(qn('w:abstractNumId'), str(abstract_num_id))
    
    # nsid
    nsid = OxmlElement('w:nsid')
    nsid.set(qn('w:val'), nsid_val)
    abs_num.append(nsid)
    
    # multiLevelType
    mlt = OxmlElement('w:multiLevelType')
    mlt.set(qn('w:val'), 'hybridMultilevel')
    abs_num.append(mlt)
    
    # template
    tmpl = OxmlElement('w:tmpl')
    tmpl.set(qn('w:val'), nsid_val)
    abs_num.append(tmpl)
    
    for ilvl in range(min(levels, 9)):
        lvl = OxmlElement('w:lvl')
        lvl.set(qn('w:ilvl'), str(ilvl))
        lvl.set(qn('w:tplc'), str(ilvl))
        
        # start
        start = OxmlElement('w:start')
        start.set(qn('w:val'), '1')
        lvl.append(start)
        
        # numFmt
        numFmt = OxmlElement('w:numFmt')
        numFmt.set(qn('w:val'), 'bullet')
        lvl.append(numFmt)
        
        # lvlText
        lvlText = OxmlElement('w:lvlText')
        lvlText.set(qn('w:val'), bullet_char)
        lvl.append(lvlText)
        
        # lvlJc
        lvlJc = OxmlElement('w:lvlJc')
        lvlJc.set(qn('w:val'), 'left')
        lvl.append(lvlJc)
        
        # pPr with indentation
        pPr = OxmlElement('w:pPr')
        ind = OxmlElement('w:ind')
        left_twips = int(left_indents_cm[ilvl] * 567)
        hang_twips = int(hanging_indents_cm[ilvl] * 567)
        ind.set(qn('w:left'), str(left_twips))
        ind.set(qn('w:hanging'), str(hang_twips))
        pPr.append(ind)
        lvl.append(pPr)
        
        # rPr with font
        rPr = OxmlElement('w:rPr')
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), bullet_font)
        rFonts.set(qn('w:hAnsi'), bullet_font)
        rFonts.set(qn('w:hint'), 'default')
        rPr.append(rFonts)
        lvl.append(rPr)
        
        abs_num.append(lvl)
    
    numbering_elm.append(abs_num)
    
    # Create num instance
    num = OxmlElement('w:num')
    num.set(qn('w:numId'), str(num_id))
    absNumId_elm = OxmlElement('w:abstractNumId')
    absNumId_elm.set(qn('w:val'), str(abstract_num_id))
    num.append(absNumId_elm)
    numbering_elm.append(num)
    
    return abstract_num_id, num_id


def create_custom_numbered_definition(numbering_elm, format_str='%1.',
                                       num_fmt='decimal', levels=1,
                                       left_indents_cm=None, hanging_indents_cm=None):
    """Create a custom numbered abstract numbering definition and num instance.
    
    Args:
        numbering_elm: The <w:numbering> root element.
        format_str: Format string like '%1.', '%1)', '(%1)'.
        num_fmt: 'decimal', 'upperLetter', 'lowerLetter', 'upperRoman', 'lowerRoman'.
        levels: Number of nesting levels (1-9).
        left_indents_cm: List of left indent values per level in cm.
        hanging_indents_cm: List of hanging indent values per level in cm.
    
    Returns:
        (abstract_num_id, num_id) tuple.
    """
    if left_indents_cm is None:
        left_indents_cm = [1.27 * (i + 1) for i in range(levels)]
    if hanging_indents_cm is None:
        hanging_indents_cm = [0.63] * levels
    
    abstract_num_id = _next_abstract_num_id(numbering_elm)
    num_id = _next_num_id(numbering_elm)
    
    nsid_val = ''.join(random.choice('0123456789ABCDEF') for _ in range(8))
    
    abs_num = OxmlElement('w:abstractNum')
    abs_num.set(qn('w:abstractNumId'), str(abstract_num_id))
    
    nsid = OxmlElement('w:nsid')
    nsid.set(qn('w:val'), nsid_val)
    abs_num.append(nsid)
    
    mlt = OxmlElement('w:multiLevelType')
    mlt.set(qn('w:val'), 'hybridMultilevel')
    abs_num.append(mlt)
    
    tmpl = OxmlElement('w:tmpl')
    tmpl.set(qn('w:val'), nsid_val)
    abs_num.append(tmpl)
    
    for ilvl in range(min(levels, 9)):
        lvl = OxmlElement('w:lvl')
        lvl.set(qn('w:ilvl'), str(ilvl))
        lvl.set(qn('w:tplc'), str(ilvl))
        
        start = OxmlElement('w:start')
        start.set(qn('w:val'), '1')
        lvl.append(start)
        
        numFmt = OxmlElement('w:numFmt')
        numFmt.set(qn('w:val'), num_fmt)
        lvl.append(numFmt)
        
        lvlText = OxmlElement('w:lvlText')
        lvlText.set(qn('w:val'), format_str)
        lvl.append(lvlText)
        
        lvlJc = OxmlElement('w:lvlJc')
        lvlJc.set(qn('w:val'), 'left')
        lvl.append(lvlJc)
        
        pPr = OxmlElement('w:pPr')
        ind = OxmlElement('w:ind')
        ind.set(qn('w:left'), str(int(left_indents_cm[ilvl] * 567)))
        ind.set(qn('w:hanging'), str(int(hanging_indents_cm[ilvl] * 567)))
        pPr.append(ind)
        lvl.append(pPr)
        
        abs_num.append(lvl)
    
    numbering_elm.append(abs_num)
    
    num = OxmlElement('w:num')
    num.set(qn('w:numId'), str(num_id))
    absNumId_elm = OxmlElement('w:abstractNumId')
    absNumId_elm.set(qn('w:val'), str(abstract_num_id))
    num.append(absNumId_elm)
    numbering_elm.append(num)
    
    return abstract_num_id, num_id


def _next_abstract_num_id(numbering_elm):
    """Find the next available abstractNumId."""
    ids = numbering_elm.xpath(
        './w:abstractNum/@w:abstractNumId',
        namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    )
    return max((int(i) for i in ids), default=-1) + 1


def _next_num_id(numbering_elm):
    """Find the next available numId."""
    ids = numbering_elm.xpath(
        './w:num/@w:numId',
        namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    )
    return max((int(i) for i in ids), default=0) + 1
