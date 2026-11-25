"""Shared constants for docx_healer."""

# Glyphs that only ever mark a bullet when they open a line
BULLET_CODEPOINTS = frozenset({
    0x00B7, 0x2022, 0x2023, 0x2043, 0x204C, 0x204D, 0x2219,
    0x25A0, 0x25A1, 0x25AA, 0x25AB, 0x25B6, 0x25B8, 0x25BA,
    0x25C6, 0x25C7, 0x25CB, 0x25CF, 0x25D8, 0x25E6,
    0x2610, 0x2611, 0x2612, 0x2619, 0x2713, 0x2714, 0x2756,
    0x2765, 0x2767, 0x2794, 0x27A2, 0x27A4, 0x29BE, 0x29BF,
})

# Characters that also occur in prose or maths, so one alone is not a bullet
WEAK_BULLETS = frozenset("-*+~>»–—")

# Font fallback chain: PyMuPDF name -> [Windows preferred, macOS, Linux, universal]
FONT_FALLBACK_MAP = {
    'SimSun': ['SimSun', '\u5b8b\u4f53', 'Microsoft YaHei', '\u5fae\u8f6f\u96c5\u9ed1', 'Arial Unicode MS', 'Noto Sans CJK SC'],
    'SimHei': ['SimHei', '\u9ed1\u4f53', 'Microsoft YaHei', '\u5fae\u8f6f\u96c5\u9ed1', 'Arial Unicode MS', 'Noto Sans CJK SC'],
    'KaiTi': ['KaiTi', '\u6977\u4f53', 'SimSun', 'Microsoft YaHei', 'Noto Sans CJK SC'],
    'FangSong': ['FangSong', '\u4eff\u5b8b', 'SimSun', 'Noto Serif CJK SC'],
    'FangSong_GB2312': ['FangSong', '\u4eff\u5b8b', 'SimSun', 'Noto Serif CJK SC'],
    'MS-Mincho': ['MS Mincho', 'Yu Mincho', 'Meiryo', 'Noto Serif CJK JP'],
    'MS-Gothic': ['MS Gothic', 'Yu Gothic', 'Meiryo', 'Noto Sans CJK JP'],
    'MS-PGothic': ['MS PGothic', 'Yu Gothic', 'Meiryo', 'Noto Sans CJK JP'],
    'HYGoThic-Medium': ['Apple SD Gothic Neo', 'Malgun Gothic', 'Gulim', 'Noto Sans CJK KR'],
    'HYSMyeongJo-Medium': ['AppleMyungjo', 'Batang', 'Noto Serif CJK KR'],
    'AppleMyungjo': ['Apple SD Gothic Neo', 'Batang', 'Noto Serif CJK KR'],
    'TimesNewRomanPSMT': ['Times New Roman', 'Times', 'Liberation Serif'],
    'TimesNewRomanPS-BoldMT': ['Times New Roman', 'Times', 'Liberation Serif'],
    'TimesNewRomanPS-ItalicMT': ['Times New Roman', 'Times', 'Liberation Serif'],
    'ArialMT': ['Arial', 'Helvetica', 'Liberation Sans'],
    'Arial-BoldMT': ['Arial', 'Helvetica', 'Liberation Sans'],
    'CourierNewPSMT': ['Courier New', 'Consolas', 'Liberation Mono'],
    'Helvetica': ['Arial', 'Helvetica', 'Liberation Sans'],
    'Helvetica-Bold': ['Arial', 'Helvetica', 'Liberation Sans'],
    'Symbol': ['Symbol', 'Segoe UI Symbol'],
    'ZapfDingbats': ['Segoe UI Symbol', 'Symbol'],
    'UNNAMED': ['Arial', 'Helvetica', 'Liberation Sans'],
    'helv': ['Arial', 'Helvetica', 'Liberation Sans'],
}

CHAR_RANGE_FONT_MAP = [
    ((0x4E00, 0x9FFF), 'SimSun'),
    ((0x3400, 0x4DBF), 'SimSun'),
    ((0x20000, 0x2A6DF), 'SimSun'),
    ((0x3040, 0x309F), 'Meiryo'),
    ((0x30A0, 0x30FF), 'Meiryo'),
    ((0xAC00, 0xD7AF), 'Malgun Gothic'),
    ((0x0600, 0x06FF), 'Arabic Typesetting'),
    ((0x0590, 0x05FF), 'David'),
    ((0x0E00, 0x0E7F), 'Leelawadee UI'),
    ((0x0900, 0x097F), 'Mangal'),
]
