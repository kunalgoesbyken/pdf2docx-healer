# Benchmark

A reproducible comparison of plain `pdf2docx` and `pdf2docx-healer` on a fixed set of public PDFs. Every number comes from the PDF itself (PyMuPDF) and from the produced DOCX (XML read directly), never from log text.

## Run it

```bash
pip install -e ".[dev]"
python -m bench.fetch                 # download and verify the corpus into bench/.cache/pdfs
python -m bench.run                   # every tool on every non-slow document, plus the synthetic PDF
python -m bench.run --include-slow    # also the 492-page document
python -m bench.run --only irs-fw9-form synthetic --tools pdf2docx healer-fidelity
```

Results land in `bench/results/<date>.json` (every row, plus commit and package versions) and `bench/results/<date>.md` (the scoreboard below, ready to paste). Both folders are git-ignored; downloads go to `bench/.cache`. Downloaded files are only ever opened with PyMuPDF and never executed. `python -m bench.fetch --pin` records new SHA-256 values and page counts into `corpus.json`.

The workflow `.github/workflows/bench.yml` runs the same thing on Ubuntu when started by hand (Actions, Benchmark, Run workflow) and uploads the results as an artifact.

## Tools

| Tool | What runs |
| --- | --- |
| `pdf2docx` | `Converter(pdf).convert(docx)` with defaults |
| `healer-fidelity` | `heal(pdf, docx, HealerConfig(verbose=False))` |
| `healer-editable` | `heal(pdf, docx, HealerConfig(mode="editable", verbose=False))`; a checkout where it still raises `NotImplementedError` is reported as "not available" and left out of the averages |
| `pdf2docx-plus` | `pdf2docx-plus convert IN --output OUT`, default profile `fidelity` (the others are `fast` and `semantic`). It needs its own environment: set `BENCH_PDF2DOCX_PLUS` to its executable (or put it on `PATH`). It exits 1 when any page fails even though it writes the DOCX, so the output file decides and the logged `failed=N` is kept as a note on the row. |
| `libreoffice` | `soffice --headless --infilter=writer_pdf_import --convert-to docx`, with its own user profile (`BENCH_SOFFICE_PROFILE`, else a throwaway folder). Set `BENCH_SOFFICE` to the executable if it is not on `PATH`. |

A tool that is not installed is reported as "not available". To measure another checkout of the healer, or to add tools to an earlier result, use `--healer-src PATH/src` and `--merge bench/results/<date>.json`; rows of the same document and tool are replaced and everything else is kept.

Each conversion runs in its own child process with a timeout (900 s, four times that for slow documents); on a timeout the whole process tree is killed. A hang, a crash or a Python exception is recorded for that one row and the run goes on. Seconds are the conversion only, not interpreter start-up.

## Corpus

| Id | Category | Pages | Licence | Source | What it tests |
| --- | --- | --- | --- | --- | --- |
| `llama-two-column` | two-column | 27 | CC BY 4.0 | [page](https://arxiv.org/abs/2302.13971v1) | Two-column research paper: reading order across columns, 36 external links, tables and figures, no outline. |
| `irs-pub15-lists` | report-lists | 57 | Public domain (US federal government work, 17 U.S.C. 105) | [page](https://www.irs.gov/publications/p15) | Single-column government report with bullets, numbered and lettered lists, nested levels and many tables. |
| `irs-fw9-form` | form-tables | 6 | Public domain (US federal government work, 17 U.S.C. 105) | [page](https://www.irs.gov/forms-pubs/about-form-w-9) | Government form: ruled boxes, checkbox rows, field tables and small print in several columns. |
| `slides-hack-for-heritage` | slides | 8 | CC BY-SA 4.0 | [page](https://commons.wikimedia.org/wiki/File:Hack_for_Heritage_2020_%E2%80%93_presentation_dag_1.pdf) | Slide deck: one idea per page, large short text boxes, images on most slides, links in text. |
| `mixtral-outline` | outline | 13 | CC BY 4.0 | [page](https://arxiv.org/abs/2401.04088v1) | Paper with a PDF outline (bookmarks) of 10 entries over two levels, figures and tables. |
| `nist-csf2-links` | internal-links | 32 | Public domain (US federal government work, 17 U.S.C. 105) | [page](https://csrc.nist.gov/pubs/cswp/29/the-nist-cybersecurity-framework-csf-20/final) | Government report with 47 internal cross-reference links, 56 external links, an outline and bullet lists. |
| `zh-us-constitution` | cjk | 29 | Public domain (as declared on the Wikimedia Commons file page) | [page](https://commons.wikimedia.org/wiki/File:Constitution_of_the_United_States_(Traditional_Chinese).pdf) | CJK text interleaved with English, per-ideograph tokens, font fallback for Traditional Chinese. |
| `ar-wikipedia-education` | rtl-arabic | 2 | CC BY-SA 3.0 | [page](https://commons.wikimedia.org/wiki/File:Wikipedia-in-Arabic-Education-Program.pdf) | Arabic right-to-left text with diacritics, many small images, mixed Latin words. |
| `he-wikimedia-article` | rtl-hebrew | 8 | CC BY-SA 4.0 | [page](https://commons.wikimedia.org/wiki/File:2.3_6-215_%D7%A4%D7%96%D7%99_%D7%91%D7%91%D7%A9%D7%A2%D7%A8_15.4.2005_-_%D7%9E%D7%90%D7%9E%D7%A8%D7%95_%D7%A9%D7%9C_%D7%A4%D7%96%D7%99_%D7%91%D7%91%D7%A9%D7%A2%D7%A8_%D7%94%D7%9E%D7%A1%D7%9B%D7%9D_%D7%90%D7%AA_%D7%9E%D7%94%D7%9C%D7%9B%D7%99%D7%95_%D7%9C%D7%A4%D7%99%D7%AA%D7%95%D7%97_%D7%94%D7%9E%D7%9B%D7%9C%D7%9C%D7%95%D7%AA_%D7%94%D7%90%D7%A7%D7%93%D7%9E%D7%99%D7%95%D7%AA_%D7%94%D7%9E%D7%AA%D7%95%D7%A7%D7%A6%D7%91%D7%95%D7%AA.pdf) | Hebrew right-to-left running text over eight pages, Hebrew-only paragraphs with a few Latin words. |
| `hi-mizoram-syllabus` | devanagari | 36 | Public domain (as declared on the Wikimedia Commons file page) | [page](https://commons.wikimedia.org/wiki/File:M.A._Hindi_Syllabus_Mizoram_University.pdf) | Devanagari text with combining vowel signs and conjuncts, mixed with English, and a large course table. |
| `us-patent-174465-scan` | scanned | 6 | Public domain (US patent published in 1876) | [page](https://patents.google.com/patent/US174465A/en) | Image-only scan with no text layer: six page images, nothing to extract without OCR. |
| `wmil-financial-statements` | table-heavy | 10 | CC BY-SA 3.0 | [page](https://commons.wikimedia.org/wiki/File:Wikimedia_Israel_Financial_Statements_2012.pdf) | Table-heavy financial statement (balance sheet style): six ruled tables of figures with totals, English text. |
| `nist-sp800-53r5-long` (slow, optional) | long | 492 | Public domain (US federal government work, 17 U.S.C. 105) | [page](https://csrc.nist.gov/pubs/sp/800/53/r5/upd1/final) | Long document (492 pages): 358 outline entries, about 7,500 internal links, 226 external links, tables and lists. Slow and optional. |

Plus one generated PDF, `synthetic` (`bench/synthetic.py`), built with exact ground truth: 5 headings over 3 levels (also written as the PDF outline), 25 list items whose visible number and level are known (bullets at 3 levels, `1. 2. a. b. 3.`, a list that starts at `3.`, `i. ii. iii.`, `a) b) c)`, and `1. 1.1 1.2 2. 2.1`), 3 external and 2 internal links, one ruled table and one image.

The corpus is only linked, never committed. `fetch` checks the `%PDF-` header, the `%%EOF` marker, that PyMuPDF opens it and the SHA-256, and refuses to use a file that fails any of these.

## Ground truth

Read from the PDF with PyMuPDF (`bench/truth.py`), cached by file hash:

- words in extraction order, per page (`page.get_text()`, which for most documents is the content-stream order, so it is the reading order for LaTeX papers)
- pages, outline entries (`get_toc`), external URIs (`get_links`, kind `LINK_URI`), internal links (kind `LINK_GOTO`), tables (`find_tables`), images (`get_image_info`)

## Metrics

All ratios are 0 to 1, higher is better. A metric the document does not define (no outline, no links, no words) is left out of that document and shown as `-`.

| Metric | Definition |
| --- | --- |
| Word recall, precision | Words are NFKC-folded and case-folded; letters, digits and combining marks stay together (Devanagari and Arabic words are not split), and every CJK ideograph or kana is one word. The two sequences are aligned in order: equal prefix and suffix are trimmed, words that sit in n-grams (3, then 2, then 1 words) occurring exactly once on both sides are anchors, the longest increasing chain of anchors is kept, and the gaps between anchors use a two-pointer diff that can skip at most 32 words. Recall is matched reference words over reference words, precision is matched output words over output words. A block that moved far away is lost, as it would be for a reader. Cost is O(n log n), so a 100,000-word document takes well under a second. |
| Recall any order | Same words as a multiset, ignoring order. The gap to word recall is the reading-order loss. |
| Pages lost | A PDF page with at least 5 words is lost when fewer than half of its 8-character windows (page text with spaces removed) occur anywhere in the output text. Computed from content, so fused or reordered words do not count as a lost page, a dropped page does. Reported as lost/text pages. |
| Heading recall, heading level | For each PDF outline entry: is there an output heading paragraph (style `Heading N`, or an `outlineLvl`, direct or from the style chain) with the same words? Heading level counts only matches at the same level. Needs a PDF outline, so most documents show `-`. |
| List real | Synthetic only. Fraction of the expected list items found as paragraphs that carry real numbering (`w:numPr`) of the right kind. |
| List level | Synthetic only. Fraction whose `w:ilvl` equals the expected level. |
| Number shown | Synthetic only. The visible label is rendered from `numbering.xml` following ECMA-376 (`w:start`, `w:startOverride`, `w:numFmt`, `w:lvlText`, `w:lvlRestart`, `w:isLgl`; paragraphs of different `w:num` on one abstract list continue each other unless `startOverride` restarts them) and must equal the PDF label exactly (`3.`, `iii.`, `c)`, `1.1`). A label typed into the text counts as visible but not as a real list; a numbered paragraph that also has the label typed in is a doubled label and does not count. |
| Ext links | Distinct external URIs of the PDF that appear as hyperlinks (relationship, `HYPERLINK` field) in the DOCX. |
| Int links | Output hyperlinks whose `w:anchor` names an existing bookmark (a `bookmarkStart` with a matching `bookmarkEnd`), over the PDF's internal links, capped at 1. Dangling anchors are listed per row. |
| Tables, Images | min(count)/max(count) of DOCX versus PDF, 1 when both are zero. Phantom layout tables lower the score. `find_tables` finds ruled tables, so the PDF count is a reference signal, not an oracle. |
| Seconds | Wall time of the conversion in the child process. |

Visible numbers count rendered labels as words, so turning `3. text` into a real list is not scored as lost text, and a wrong or doubled label costs precision.

An output that python-docx cannot open, or that repeats a `w:bookmarkStart` id, is `invalid`. Failed, crashed, timed-out and invalid runs score 0 on every metric the PDF defines, so a tool cannot look better by failing on hard files. Unavailable tools are listed but not averaged.

## Known limits

- Reference reading order is the PDF's own extraction order. For infographics and slides that order is arbitrary, so word recall is low for every tool there; Recall any order shows how much text survived.
- Lists, visible numbers and levels are measured exactly only on the synthetic PDF. For real PDFs the benchmark reports how many list paragraphs the output has and which levels it uses, but there is no independent list truth.
- Heading metrics need a PDF outline. Documents without one say nothing about heading detection.
- Tables and Images compare counts with what PyMuPDF finds in the PDF. A PDF whose figures are vector drawings has no images, so a converter that rasterises them scores 0 on Images and one that drops them scores 1; read the count columns before reading that as quality.
- OCR is not scored (it needs Tesseract). The scanned document only shows what each tool does with image-only pages by default.
- Numbering is rendered by this scorer, not by Word. Unknown `numFmt` values render as decimal and are listed in the JSON (`approx_formats`); `w:start` omitted means 0 as the standard says.

## Results

Run 2026-11-25: pdf2docx-healer at commit `8e4bbf8`, benchmark at commit `3033935+dirty` (python 3.12.4, pymupdf 2.1, pdf2docx 0.5.13, python-docx 1.2.0, lxml 6.1.3, pdf2docx-healer 0.2.0, platform Windows-11-10.0.26300-SP0, pdf2docx-plus 0.6.10, libreoffice build 861effc5289d).

| Tool | Ok | Failed | Word recall | Word precision | Recall any order | Heading recall | Heading level | List real | List level | Number shown | Ext links | Int links | Tables | Images | Pages lost | Median s |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| pdf2docx | 13 | 0 | 0.849 | 0.869 | 0.959 | 0.00 | 0.00 | 0.00 | 0.00 | 1.00 | 0.99 | 0.00 | 0.33 | 0.66 | 0/230 | 12.5 |
| healer-fidelity | 13 | 0 | 0.849 | 0.870 | 0.959 | 0.52 | 0.52 | 0.76 | 0.68 | 1.00 | 0.99 | 0.95 | 0.33 | 0.66 | 0/230 | 7.3 |
| healer-editable | 13 | 0 | 0.910 | 0.922 | 0.984 | 0.59 | 0.59 | 0.68 | 0.68 | 0.68 | 0.99 | 0.86 | 0.80 | 0.63 | 0/230 | 5.2 |
| pdf2docx-plus | 13 | 0 | 0.855 | 0.861 | 0.987 | 0.00 | 0.00 | 0.68 | 0.52 | 0.88 | 0.99 | 0.00 | 0.64 | 0.66 | 0/230 | 8.7 |
| libreoffice | 13 | 0 | 0.700 | 0.691 | 0.934 | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 | 0.23 | 0.93 | 36/230 | 12.6 |

| Document | Converter | Status | Recall | Precision | Pages lost | List paras | Ext links | Int links | Tables d/p | Images d/p | Seconds |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| llama-two-column | pdf2docx | ok | 0.619 | 0.748 | 0/27 | 0 | 1.00 | - | 82/17 | 7/0 | 80.0 |
| llama-two-column | healer-fidelity | ok | 0.619 | 0.748 | 0/27 | 0 | 1.00 | - | 82/17 | 7/0 | 82.1 |
| llama-two-column | healer-editable | ok | 0.956 | 0.975 | 0/27 | 7 | 1.00 | - | 14/17 | 5/0 | 19.0 |
| llama-two-column | pdf2docx-plus | ok | 0.780 | 0.780 | 0/27 | 5 | 1.00 | - | 4/17 | 7/0 | 59.2 |
| llama-two-column | libreoffice | ok | 0.710 | 0.709 | 0/27 | 0 | 0.00 | - | 0/17 | 0/0 | 33.9 |
| irs-pub15-lists | pdf2docx | ok | 0.953 | 0.955 | 0/57 | 0 | 1.00 | - | 87/16 | 43/2 | 61.1 |
| irs-pub15-lists | healer-fidelity | ok | 0.953 | 0.955 | 0/57 | 8 | 0.99 | - | 87/16 | 43/2 | 57.4 |
| irs-pub15-lists | healer-editable | ok | 0.926 | 0.951 | 0/57 | 279 | 1.00 | - | 8/16 | 0/2 | 28.5 |
| irs-pub15-lists | pdf2docx-plus | ok | 0.950 | 0.949 | 0/57 | 309 | 0.99 | - | 33/16 | 31/2 | 98.1 |
| irs-pub15-lists | libreoffice | ok | 0.555 | 0.532 | 0/57 | 0 | 0.00 | - | 0/16 | 2/2 | 84.7 |
| irs-fw9-form | pdf2docx | ok | 0.964 | 0.977 | 0/6 | 0 | - | - | 29/5 | 0/0 | 12.5 |
| irs-fw9-form | healer-fidelity | ok | 0.964 | 0.977 | 0/6 | 41 | - | - | 29/5 | 0/0 | 5.9 |
| irs-fw9-form | healer-editable | ok | 0.966 | 0.972 | 0/6 | 42 | - | - | 3/5 | 0/0 | 5.0 |
| irs-fw9-form | pdf2docx-plus | ok | 0.979 | 0.985 | 0/6 | 56 | - | - | 19/5 | 0/0 | 8.7 |
| irs-fw9-form | libreoffice | ok | 0.618 | 0.618 | 0/6 | 0 | - | - | 0/5 | 0/0 | 10.0 |
| slides-hack-for-heritage | pdf2docx | ok | 0.819 | 0.819 | 0/7 | 0 | 1.00 | - | 17/0 | 13/17 | 4.6 |
| slides-hack-for-heritage | healer-fidelity | ok | 0.819 | 0.819 | 0/7 | 0 | 1.00 | - | 17/0 | 13/17 | 4.4 |
| slides-hack-for-heritage | healer-editable | ok | 0.867 | 0.867 | 0/7 | 4 | 1.00 | - | 0/0 | 11/17 | 2.9 |
| slides-hack-for-heritage | pdf2docx-plus | ok | 0.724 | 0.724 | 0/7 | 2 | 1.00 | - | 0/0 | 13/17 | 4.2 |
| slides-hack-for-heritage | libreoffice | ok | 0.814 | 0.814 | 0/7 | 0 | 0.00 | - | 0/0 | 17/17 | 7.8 |
| mixtral-outline | pdf2docx | ok | 0.862 | 0.934 | 0/13 | 0 | 1.00 | - | 27/6 | 11/8 | 567.2 |
| mixtral-outline | healer-fidelity | ok | 0.862 | 0.934 | 0/13 | 0 | 1.00 | - | 27/6 | 11/8 | 226.9 |
| mixtral-outline | healer-editable | ok | 0.933 | 0.938 | 0/13 | 6 | 1.00 | - | 5/6 | 22/8 | 7.7 |
| mixtral-outline | pdf2docx-plus | ok | 0.952 | 0.955 | 0/13 | 41 | 1.00 | - | 12/6 | 11/8 | 243.3 |
| mixtral-outline | libreoffice | ok | 0.947 | 0.944 | 0/13 | 0 | 0.00 | - | 0/6 | 8/8 | 17.2 |
| nist-csf2-links | pdf2docx | ok | 0.948 | 0.959 | 0/32 | 0 | 0.97 | 0.00 | 6/3 | 12/13 | 10.1 |
| nist-csf2-links | healer-fidelity | ok | 0.948 | 0.959 | 0/32 | 0 | 0.97 | 0.89 | 6/3 | 12/13 | 9.6 |
| nist-csf2-links | healer-editable | ok | 0.939 | 0.983 | 0/32 | 69 | 0.97 | 0.72 | 1/3 | 10/13 | 6.0 |
| nist-csf2-links | pdf2docx-plus | ok | 0.938 | 0.942 | 0/32 | 43 | 0.97 | 0.00 | 4/3 | 12/13 | 11.4 |
| nist-csf2-links | libreoffice | ok | 0.952 | 0.947 | 0/32 | 0 | 0.00 | 0.00 | 0/3 | 13/13 | 12.6 |
| zh-us-constitution | pdf2docx | ok | 1.000 | 1.000 | 0/29 | 0 | - | - | 29/8 | 0/0 | 22.0 |
| zh-us-constitution | healer-fidelity | ok | 1.000 | 1.000 | 0/29 | 0 | - | - | 29/8 | 0/0 | 9.3 |
| zh-us-constitution | healer-editable | ok | 0.958 | 0.958 | 0/29 | 0 | - | - | 6/8 | 0/0 | 9.0 |
| zh-us-constitution | pdf2docx-plus | ok | 0.993 | 0.990 | 0/29 | 2 | - | - | 1/8 | 0/0 | 16.3 |
| zh-us-constitution | libreoffice | ok | 0.627 | 0.627 | 0/29 | 0 | - | - | 0/8 | 0/0 | 14.1 |
| ar-wikipedia-education | pdf2docx | ok | 0.478 | 0.483 | 0/2 | 0 | - | - | 7/0 | 5/45 | 19.7 |
| ar-wikipedia-education | healer-fidelity | ok | 0.478 | 0.483 | 0/2 | 0 | - | - | 7/0 | 5/45 | 7.3 |
| ar-wikipedia-education | healer-editable | ok | 0.463 | 0.470 | 0/2 | 0 | - | - | 0/0 | 16/45 | 5.2 |
| ar-wikipedia-education | pdf2docx-plus | ok | 0.477 | 0.481 | 0/2 | 0 | - | - | 0/0 | 5/45 | 7.0 |
| ar-wikipedia-education | libreoffice | ok | 0.382 | 0.359 | 0/2 | 0 | - | - | 0/0 | 6/45 | 41.9 |
| he-wikimedia-article | pdf2docx | ok | 0.597 | 0.614 | 0/8 | 0 | - | - | 9/3 | 1/0 | 7.9 |
| he-wikimedia-article | healer-fidelity | ok | 0.597 | 0.614 | 0/8 | 0 | - | - | 9/3 | 1/0 | 5.2 |
| he-wikimedia-article | healer-editable | ok | 0.961 | 0.965 | 0/8 | 0 | - | - | 3/3 | 0/0 | 3.0 |
| he-wikimedia-article | pdf2docx-plus | ok | 0.597 | 0.613 | 0/8 | 0 | - | - | 3/3 | 1/0 | 4.8 |
| he-wikimedia-article | libreoffice | ok | 0.584 | 0.578 | 0/8 | 0 | - | - | 0/3 | 0/0 | 8.6 |
| hi-mizoram-syllabus | pdf2docx | ok | 0.983 | 0.985 | 0/36 | 0 | - | - | 137/3 | 0/0 | 114.3 |
| hi-mizoram-syllabus | healer-fidelity | ok | 0.984 | 0.986 | 0/36 | 44 | - | - | 137/3 | 0/0 | 37.6 |
| hi-mizoram-syllabus | healer-editable | ok | 0.989 | 1.000 | 0/36 | 287 | - | - | 2/3 | 0/0 | 7.0 |
| hi-mizoram-syllabus | pdf2docx-plus | ok | 0.947 | 0.961 | 0/36 | 499 | - | - | 3/3 | 0/0 | 25.0 |
| hi-mizoram-syllabus | libreoffice | ok | 0.250 | 0.253 | 36/36 | 0 | - | - | 0/3 | 0/0 | 29.3 |
| us-patent-174465-scan | pdf2docx | ok | - | - | 0/0 | 0 | - | - | 0/0 | 6/6 | 4.9 |
| us-patent-174465-scan | healer-fidelity | ok | - | - | 0/0 | 0 | - | - | 0/0 | 6/6 | 2.1 |
| us-patent-174465-scan | healer-editable | ok | - | - | 0/0 | 0 | - | - | 0/0 | 0/6 | 1.1 |
| us-patent-174465-scan | pdf2docx-plus | ok | - | - | 0/0 | 0 | - | - | 0/0 | 6/6 | 3.0 |
| us-patent-174465-scan | libreoffice | ok | - | - | 0/0 | 0 | - | - | 0/0 | 6/6 | 7.5 |
| wmil-financial-statements | pdf2docx | ok | 0.958 | 0.959 | 0/10 | 0 | - | - | 14/6 | 0/0 | 5.5 |
| wmil-financial-statements | healer-fidelity | ok | 0.958 | 0.959 | 0/10 | 15 | - | - | 14/6 | 0/0 | 2.2 |
| wmil-financial-statements | healer-editable | ok | 0.959 | 0.981 | 0/10 | 11 | - | - | 5/6 | 0/0 | 3.1 |
| wmil-financial-statements | pdf2docx-plus | ok | 0.946 | 0.971 | 0/10 | 19 | - | - | 0/6 | 0/0 | 1.9 |
| wmil-financial-statements | libreoffice | ok | 0.963 | 0.910 | 0/10 | 0 | - | - | 0/6 | 0/0 | 7.6 |
| synthetic | pdf2docx | ok | 1.000 | 1.000 | 0/3 | 0 | 1.00 | 0.00 | 1/1 | 1/1 | 1.6 |
| synthetic | healer-fidelity | ok | 1.000 | 1.000 | 0/3 | 19 | 1.00 | 1.00 | 1/1 | 1/1 | 0.8 |
| synthetic | healer-editable | ok | 1.000 | 1.000 | 0/3 | 17 | 1.00 | 1.00 | 1/1 | 1/1 | 1.3 |
| synthetic | pdf2docx-plus | ok | 0.983 | 0.983 | 0/3 | 17 | 1.00 | 0.00 | 1/1 | 1/1 | 0.9 |
| synthetic | libreoffice | ok | 1.000 | 1.000 | 0/3 | 0 | 0.00 | 0.00 | 0/1 | 1/1 | 5.4 |

