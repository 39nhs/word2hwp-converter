# -*- coding: utf-8 -*-
"""Word(.docx) -> 한글(.hwp / .hwpx) 변환 엔진.

한컴오피스나 MS Word 가 설치되어 있지 않아도 동작한다.
- 한컴오피스(한글)가 설치되어 있으면 한글 프로그램으로 직접 변환(가장 정확)
- 없으면 python-hwpx 로 문서를 다시 만들어 저장
  (글꼴·크기·굵게/기울임/밑줄·색·정렬·들여쓰기·줄간격·문단간격·
   표(병합·배경색·테두리)·그림·쪽 크기/여백·쪽 번호·머리말/꼬리말 유지)
"""
from __future__ import annotations

import copy
import io
import os
import re
import zipfile
from dataclasses import dataclass, field, replace
from typing import Callable

from lxml import etree

# ---------------------------------------------------------------------------
# 네임스페이스 / 단위
# ---------------------------------------------------------------------------
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
V = "urn:schemas-microsoft-com:vml"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
HP_NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def w(tag: str) -> str:
    return f"{{{W}}}{tag}"


def wattr(el, name: str, default=None):
    if el is None:
        return default
    return el.get(f"{{{W}}}{name}", default)


TWIP_TO_HWP = 5          # 1 twip = 1/20 pt, 1 HWPUNIT = 1/100 pt
EMU_TO_HWP = 1 / 127     # 1 pt = 12700 EMU
TWIP_TO_MM = 25.4 / 1440

FONT_ALIASES = {
    "malgun gothic": "맑은 고딕", "batang": "바탕", "batangche": "바탕체",
    "gulim": "굴림", "gulimche": "굴림체", "dotum": "돋움", "dotumche": "돋움체",
    "gungsuh": "궁서", "gungsuhche": "궁서체", "nanumgothic": "나눔고딕",
    "nanummyeongjo": "나눔명조", "nanum gothic": "나눔고딕", "nanum myeongjo": "나눔명조",
    "nanumbarungothic": "나눔바른고딕", "hcr batang": "함초롬바탕", "hcr dotum": "함초롬돋움",
}

HIGHLIGHT = {
    "yellow": "#FFFF00", "green": "#00FF00", "cyan": "#00FFFF", "magenta": "#FF00FF",
    "blue": "#0000FF", "red": "#FF0000", "darkBlue": "#000080", "darkCyan": "#008080",
    "darkGreen": "#008000", "darkMagenta": "#800080", "darkRed": "#800000",
    "darkYellow": "#808000", "darkGray": "#808080", "lightGray": "#C0C0C0",
    "black": "#000000", "white": "#FFFFFF",
}

# Symbol 글꼴의 글자 자리 → 유니코드 (Word 에서 그리스 문자·기호를 Symbol 글꼴로 넣은 경우)
_SYMBOL_MAP = dict(zip("abgdezhqiklmnxoprstufcywjvABGDEZHQIKLMNXOPRSTUFCYWJV",
                       "αβγδεζηθικλμνξοπρστυφχψωϕϖΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩϑς"))
_SYMBOL_MAP.update({"\xb4": "×", "\xb1": "±", "\xa3": "≤", "\xb3": "≥", "\xb9": "≠", "\xbb": "≈",
                    "\xa5": "∞", "\xae": "→", "\xac": "←", "\xb7": "•", "\xb0": "°", "\xd6": "√",
                    "\xe5": "∑", "\xf2": "∫", "\xb6": "∂", "\xd1": "∇", "\xce": "∈", "\xa2": "′",
                    "\xb2": "″", "\xbe": "─", "\xde": "⇒", "\xdb": "⇔", "\xb8": "÷", "\xd7": "·",
                    "\x60": "‾", "\x22": "∀", "\x24": "∃", "\x27": "∋", "\x40": "≅", "\x5e": "⊥"})
WINGDINGS = {0xFC: "✓", 0xFB: "✗", 0xA7: "▪", 0xD8: "➢", 0x6C: "●", 0x6E: "■", 0x71: "❑",
             0x76: "❖", 0xA8: "◻", 0x9F: "•", 0xE0: "→", 0xE8: "➔", 0x4A: "☺", 0x77: "◆", 0xAB: "★"}


def symbol_char(ch: str) -> str:
    code = ord(ch)
    if 0xF000 <= code <= 0xF0FF:
        ch = chr(code - 0xF000)
    return _SYMBOL_MAP.get(ch, ch)


PAGE_FIELDS = {"PAGE": "pagenum", "NUMPAGES": "numpages", "SECTIONPAGES": "numpages"}

HWP_LINE_WIDTHS = [0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0]


def on_off(el, default=True) -> bool | None:
    """w:b, w:i 같은 토글 속성 해석. el 이 없으면 None."""
    if el is None:
        return None
    val = wattr(el, "val")
    if val is None:
        return default
    return val.lower() not in ("0", "false", "off", "none")


def norm_font(name: str | None) -> str | None:
    if not name:
        return None
    return FONT_ALIASES.get(name.strip().lower(), name.strip())


def hex_color(val: str | None) -> str | None:
    if not val or val.lower() == "auto":
        return None
    if re.fullmatch(r"[0-9A-Fa-f]{6}", val):
        return "#" + val.upper()
    return None


# ---------------------------------------------------------------------------
# Word 수식(OMML) → LaTeX → 한글 수식 스크립트
# ---------------------------------------------------------------------------
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_INVALID_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")

_UNI2TEX = {
    "α": r"\alpha", "β": r"\beta", "γ": r"\gamma", "δ": r"\delta", "ε": r"\epsilon", "ϵ": r"\epsilon",
    "ζ": r"\zeta", "η": r"\eta", "θ": r"\theta", "ι": r"\iota", "κ": r"\kappa", "λ": r"\lambda",
    "μ": r"\mu", "µ": r"\mu", "ν": r"\nu", "ξ": r"\xi", "π": r"\pi", "ρ": r"\rho", "σ": r"\sigma",
    "τ": r"\tau", "υ": r"\upsilon", "φ": r"\phi", "ϕ": r"\phi", "χ": r"\chi", "ψ": r"\psi", "ω": r"\omega",
    "Γ": r"\Gamma", "Δ": r"\Delta", "∆": r"\Delta", "Θ": r"\Theta", "Λ": r"\Lambda", "Ξ": r"\Xi",
    "Π": r"\Pi", "Σ": r"\Sigma", "Φ": r"\Phi", "Ψ": r"\Psi", "Ω": r"\Omega",
    "×": r"\times", "·": r"\cdot", "⋅": r"\cdot", "∙": r"\cdot", "÷": r"\div", "±": r"\pm", "∓": r"\mp",
    "≤": r"\le", "≥": r"\ge", "≠": r"\neq", "≈": r"\approx", "∞": r"\infty", "→": r"\to",
    "←": r"\leftarrow", "⇒": r"\Rightarrow", "⇔": r"\Leftrightarrow", "∂": r"\partial", "∇": r"\nabla",
    "∈": r"\in", "∑": r"\sum", "∏": r"\prod", "∫": r"\int", "…": r"\cdots", "⋯": r"\cdots",
    "−": "-", "–": "-", "′": "'", "∝": r"\propto", "≡": r"\equiv", "∼": r"\sim", "°": r"^{\circ}",
}
_NARY = {"∑": r"\sum", "∏": r"\prod", "∫": r"\int", "∬": r"\iint", "∭": r"\iiint", "∮": r"\oint",
         "⋃": r"\bigcup", "⋂": r"\bigcap"}
_ACC = {"̂": r"\hat", "^": r"\hat", "̄": r"\bar", "¯": r"\bar", "̇": r"\dot", "⃗": r"\vec",
        "̃": r"\tilde", "~": r"\tilde", "̈": r"\ddot"}


def _tex_text(s: str, upright: bool) -> str:
    if upright and any(ord(ch) >= 128 and ch not in _UNI2TEX for ch in s):
        # 한글 등이 섞인 일반 글자는 통째로 글자 상자로
        safe = s.replace("\\", "").replace("{", "(").replace("}", ")")
        parts = re.split(r"([%#&$_^~])", safe)
        return "".join((r"\text{" + p + "}") if i % 2 == 0 and p else ("\\" + p if p in "%#&$_" else "")
                       for i, p in enumerate(parts))
    out = []
    plain = []

    def flush():
        if plain:
            word = "".join(plain)
            out.append(r"\mathrm{" + word + "}" if upright and word.strip() else word)
            plain.clear()
    for ch in s:
        if ch in _UNI2TEX:
            flush()
            out.append(" " + _UNI2TEX[ch] + " ")
        elif ch in "{}%#&$_":
            flush()
            out.append("\\" + ch)
        elif ch in "\\^~":
            flush()
        elif ord(ch) < 128:
            plain.append(ch)
        else:  # 한글 등 기타 글자
            flush()
            out.append(r"\text{" + ch + "}")
    flush()
    return "".join(out)


def omml_to_latex(el) -> str:
    def m(tag):
        return f"{{{M}}}{tag}"

    def kids(e):
        return "".join(conv(c) for c in e) if e is not None else ""

    def val(e, tag, attr="val"):
        x = e.find(m(tag)) if e is not None else None
        return x.get(m(attr)) if x is not None else None

    def conv(e):
        if not isinstance(e.tag, str):
            return ""
        ns, tag = etree.QName(e).namespace, etree.QName(e).localname
        if ns == W:
            if tag == "r":
                return _tex_text("".join(t.text or "" for t in e.iter(w("t"))), True)
            return ""
        if ns != M:
            return ""
        if tag in ("oMathPara", "oMath", "e", "num", "den", "sub", "sup", "deg", "lim", "fName"):
            return kids(e)
        if tag == "r":
            rpr = e.find(m("rPr"))
            upright = rpr is not None and (rpr.find(m("nor")) is not None or val(rpr, "sty") == "p")
            return _tex_text("".join(t.text or "" for t in e.findall(m("t"))), upright)
        if tag == "f":
            num, den = kids(e.find(m("num"))), kids(e.find(m("den")))
            if val(e.find(m("fPr")), "type") == "lin":
                return f"{{{num}}}/{{{den}}}"
            return rf"\frac{{{num}}}{{{den}}}"
        if tag == "sSup":
            return f"{{{kids(e.find(m('e')))}}}^{{{kids(e.find(m('sup')))}}}"
        if tag == "sSub":
            return f"{{{kids(e.find(m('e')))}}}_{{{kids(e.find(m('sub')))}}}"
        if tag == "sSubSup":
            return f"{{{kids(e.find(m('e')))}}}_{{{kids(e.find(m('sub')))}}}^{{{kids(e.find(m('sup')))}}}"
        if tag == "sPre":
            return f"{{}}_{{{kids(e.find(m('sub')))}}}^{{{kids(e.find(m('sup')))}}}{{{kids(e.find(m('e')))}}}"
        if tag == "rad":
            deg = kids(e.find(m("deg")))
            hide = val(e.find(m("radPr")), "degHide") in ("1", "on", "true")
            body = kids(e.find(m("e")))
            return rf"\sqrt[{deg}]{{{body}}}" if deg.strip() and not hide else rf"\sqrt{{{body}}}"
        if tag == "nary":
            pr = e.find(m("naryPr"))
            op = _NARY.get(val(pr, "chr") or "∫", r"\int")
            sub, sup = kids(e.find(m("sub"))), kids(e.find(m("sup")))
            s = op
            if sub.strip():
                s += f"_{{{sub}}}"
            if sup.strip():
                s += f"^{{{sup}}}"
            return s + " " + kids(e.find(m("e")))
        if tag == "d":
            pr = e.find(m("dPr"))
            beg = val(pr, "begChr")
            end = val(pr, "endChr")
            sep = val(pr, "sepChr") or ","
            beg = "(" if beg is None else beg
            end = ")" if end is None else end
            inner = (" " + sep + " ").join(kids(x) for x in e.findall(m("e")))
            tex = {"(": "(", ")": ")", "[": "[", "]": "]", "{": r"\{", "}": r"\}", "|": "|", "": "."}
            return rf"\left{tex.get(beg, '(')} {inner} \right{tex.get(end, ')')}"
        if tag == "func":
            return kids(e.find(m("fName"))) + " " + kids(e.find(m("e")))
        if tag == "limLow":
            return f"{{{kids(e.find(m('e')))}}}_{{{kids(e.find(m('lim')))}}}"
        if tag == "limUpp":
            return f"{{{kids(e.find(m('e')))}}}^{{{kids(e.find(m('lim')))}}}"
        if tag == "acc":
            cmd = _ACC.get(val(e.find(m("accPr")), "chr") or "̂", r"\hat")
            return rf"{cmd}{{{kids(e.find(m('e')))}}}"
        if tag == "bar":
            pos = val(e.find(m("barPr")), "pos")
            return (r"\underline{" if pos == "bot" else r"\overline{") + kids(e.find(m("e"))) + "}"
        if tag in ("eqArr", "m"):
            rows = e.findall(m("e")) if tag == "eqArr" else e.findall(m("mr"))
            return r" , ".join(kids(r) for r in rows)
        if tag in ("groupChr", "box", "borderBox", "phant"):
            return kids(e.find(m("e")))
        if tag.endswith("Pr"):
            return ""
        return kids(e)

    return re.sub(r"\s+", " ", conv(el)).strip()


def omml_plain_text(el) -> str:
    return "".join(t.text or "" for t in el.iter(f"{{{M}}}t", w("t")))


def omml_to_eqedit(el) -> str | None:
    try:
        from hwpx.equation.authoring import latex_to_eqedit
        tex = omml_to_latex(el)
        if not tex:
            return None
        return latex_to_eqedit(tex)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# 글꼴 줄 높이 (Word 의 "1줄" 높이 = 글꼴의 winAscent+winDescent)
# ---------------------------------------------------------------------------
_FONT_INDEX: dict | None = None


def _font_files():
    dirs = [os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")]
    for d in dirs:
        if os.path.isdir(d):
            for n in os.listdir(d):
                if n.lower().endswith((".ttf", ".otf", ".ttc")):
                    yield os.path.join(d, n)


def _parse_font(data: bytes, offset: int):
    """(이름들, 줄높이 비율) — name/OS2/head 테이블만 읽는다"""
    import struct
    num = struct.unpack(">H", data[offset + 4:offset + 6])[0]
    tables = {}
    for i in range(num):
        rec = offset + 12 + i * 16
        tag = data[rec:rec + 4].decode("latin1")
        _, toff, tlen = struct.unpack(">III", data[rec + 4:rec + 16])
        tables[tag] = (toff, tlen)
    names = set()
    if "name" in tables:
        toff, _ = tables["name"]
        _, count, stor = struct.unpack(">HHH", data[toff:toff + 6])
        for i in range(count):
            pid, eid, lid, nid, ln, off = struct.unpack(">HHHHHH", data[toff + 6 + i * 12:toff + 18 + i * 12])
            if nid not in (1, 4, 16):
                continue
            raw = data[toff + stor + off:toff + stor + off + ln]
            try:
                s = raw.decode("utf-16-be") if pid in (0, 3) else raw.decode("latin1")
            except UnicodeDecodeError:
                continue
            names.add(s.strip().lower())
    ratio = None
    if "OS/2" in tables and "head" in tables:
        hoff, _ = tables["head"]
        upem = struct.unpack(">H", data[hoff + 18:hoff + 20])[0]
        ooff, _ = tables["OS/2"]
        win_asc, win_desc = struct.unpack(">HH", data[ooff + 74:ooff + 78])
        if upem:
            ratio = (win_asc + win_desc) / upem
    return names, ratio


_FONT_PATHS: dict = {}
_WIDTH_CACHE: dict = {}
_PIL_FONTS: dict = {}


def _pil_font(name: str | None):
    """글꼴 이름 → PIL 글꼴 (없으면 비슷한 기본 글꼴)"""
    font_line_ratio(None)  # 색인 만들기
    key = (name or "").strip().lower()
    if key in _PIL_FONTS:
        return _PIL_FONTS[key]
    cand = _FONT_PATHS.get(key)
    if cand is None:
        serif = any(k in key for k in ("바탕", "batang", "명조", "myeongjo", "궁서", "serif", "times", "신명"))
        for alt in (("바탕", "batang") if serif else ("맑은 고딕", "malgun gothic", "굴림", "gulim")):
            cand = _FONT_PATHS.get(alt)
            if cand:
                break
    font = None
    if cand:
        try:
            from PIL import ImageFont
            font = ImageFont.truetype(cand[0], 100, index=cand[1])
        except Exception:
            font = None
    _PIL_FONTS[key] = font
    return font


def char_width_em(ch: str, font_name: str | None) -> float:
    k = (font_name, ch)
    if k in _WIDTH_CACHE:
        return _WIDTH_CACHE[k]
    f = _pil_font(font_name)
    if f is not None:
        try:
            v = f.getlength(ch) / 100
        except Exception:
            v = None
    else:
        v = None
    if not v:
        v = 1.0 if ord(ch) >= 0x2E80 else 0.33 if ch == " " else 0.55
    _WIDTH_CACHE[k] = v
    return v


def count_lines(text: str, size: float, avail_pt: float, ea: str | None, latin: str | None) -> int:
    """한글처럼 줄 나눔을 흉내 내서 줄 수 세기 (영어 낱말은 안 쪼갬, 한글은 글자 단위)"""
    import math
    total_lines = 0
    for part in (text or "").split("\n"):
        tokens = re.findall(r" +|[!-~]+|[^!-~ ]", part.replace("\t", "    "))
        lines, cur = 1, 0.0
        for tok in tokens:
            w_ = sum(char_width_em(ch, latin if ord(ch) < 128 else ea) for ch in tok) * size
            if tok.startswith(" "):
                cur += w_
                continue
            if cur + w_ > avail_pt and cur > 0:
                lines += 1
                cur = 0.0
                if w_ > avail_pt:  # 한 낱말이 칸보다 넓음
                    lines += math.ceil(w_ / avail_pt) - 1
                    cur = w_ % avail_pt
                    continue
            cur += w_
        total_lines += lines
    return total_lines


def font_line_ratio(name: str | None) -> float:
    global _FONT_INDEX
    if _FONT_INDEX is None:
        _FONT_INDEX = {}
        import struct
        for path in _font_files():
            try:
                with open(path, "rb") as fh:
                    data = fh.read()
                offsets = [0]
                if data[:4] == b"ttcf":
                    n = struct.unpack(">I", data[8:12])[0]
                    offsets = list(struct.unpack(">%dI" % n, data[12:12 + 4 * n]))
                for fi, off in enumerate(offsets):
                    names, ratio = _parse_font(data, off)
                    if ratio:
                        for nm in names:
                            _FONT_INDEX.setdefault(nm, ratio)
                            _FONT_PATHS.setdefault(nm, (path, fi))
            except Exception:
                continue
    if name:
        r = _FONT_INDEX.get(name.strip().lower())
        if r:
            return min(max(r, 1.0), 2.0)
        # 이 PC 에 없는 글꼴: 자주 쓰는 한글 글꼴의 대략값
        for key, val in (("함초롬", 1.4), ("hcr", 1.4), ("맑은", 1.33), ("malgun", 1.33),
                         ("나눔", 1.15), ("nanum", 1.15)):
            if key in name.lower():
                return val
    return 1.0


# ---------------------------------------------------------------------------
# 서식 모델
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class RunFmt:
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    color: str | None = None
    size: float = 10.0
    font_ea: str | None = None
    font_latin: str | None = None
    highlight: str | None = None
    script: str | None = None      # "sup" / "sub"
    ratio: int | None = None
    hidden: bool = False
    caps: bool = False


@dataclass
class ParaFmt:
    align: str | None = None        # LEFT/CENTER/RIGHT/JUSTIFY/DISTRIBUTE
    left: int = 0                   # twips
    right: int = 0
    first_line: int = 0             # 양수: 첫줄 들여쓰기, 음수: 내어쓰기
    before: int = 0                 # twips
    after: int = 0
    line: int | None = None         # spacing w:line
    line_rule: str = "auto"
    keep_next: bool = False
    keep_lines: bool = False
    page_break_before: bool = False
    num_id: str | None = None
    num_level: int = 0
    contextual: bool = False
    tabs: dict = field(default_factory=dict)   # pos(twips) -> (종류, 채움)


# ---------------------------------------------------------------------------
# DOCX 패키지 읽기
# ---------------------------------------------------------------------------
class DocxPackage:
    def __init__(self, path: str):
        self.path = path
        with open(path, "rb") as fh:
            head = fh.read(8)
        if head.startswith(b"\xd0\xcf\x11\xe0"):
            raise ValueError("암호가 걸린 문서이거나 예전 Word 형식(.doc)입니다. "
                             "Word 에서 암호를 풀고 .docx 로 저장한 뒤 다시 넣어 주세요.")
        if not head.startswith(b"PK"):
            raise ValueError("Word(.docx) 파일이 아니거나 파일이 손상되었습니다.")
        try:
            self.zip = zipfile.ZipFile(path)
            self.names = set(self.zip.namelist())
        except zipfile.BadZipFile:
            raise ValueError("파일이 손상되어 열 수 없습니다. (다운로드가 덜 되었는지 확인해 주세요)")
        main = "word/document.xml"
        ct = self._xml("[Content_Types].xml")
        if ct is not None:  # 본문 파일 이름이 다른 경우 (일부 프로그램)
            for o in ct:
                if (o.get("ContentType") or "").endswith("document.main+xml") or \
                        (o.get("ContentType") or "").endswith("document.macroEnabled.main+xml") or \
                        (o.get("ContentType") or "").endswith("template.main+xml"):
                    main = (o.get("PartName") or "/" + main).lstrip("/")
        self.main_part = main
        try:
            self.document = self._xml(main)
        except etree.XMLSyntaxError:
            raise ValueError("문서 내용이 손상되어 읽을 수 없습니다.")
        if self.document is None:
            raise ValueError("올바른 Word(.docx) 파일이 아닙니다.")
        self.rels = self.part_rels(main)

        def related(kind, default):
            t = next((t for t, typ in self.rels.values() if typ.endswith("/" + kind)), None)
            try:
                return self._xml(t or default)
            except etree.XMLSyntaxError:
                return None
        self.styles_root = related("styles", "word/styles.xml")
        self.numbering_root = related("numbering", "word/numbering.xml")
        self.theme_fonts = self._theme_fonts()
        self.styles: dict[str, etree._Element] = {}
        self.default_para_style = None
        self.default_char_style = None
        self.default_table_style = None
        self.doc_rpr = None
        self.doc_ppr = None
        if self.styles_root is not None:
            for st in self.styles_root.iter(w("style")):
                sid = wattr(st, "styleId")
                if sid:
                    self.styles[sid] = st
                if wattr(st, "default") in ("1", "true"):
                    t = wattr(st, "type")
                    if t == "paragraph":
                        self.default_para_style = sid
                    elif t == "character":
                        self.default_char_style = sid
                    elif t == "table":
                        self.default_table_style = sid
            dd = self.styles_root.find(w("docDefaults"))
            if dd is not None:
                self.doc_rpr = dd.find(f"{w('rPrDefault')}/{w('rPr')}")
                self.doc_ppr = dd.find(f"{w('pPrDefault')}/{w('pPr')}")
        self._num_cache: dict[str, dict] = {}

    def _xml(self, name):
        if name not in self.names:
            return None
        parser = etree.XMLParser(resolve_entities=False, huge_tree=True, remove_blank_text=False)
        return etree.fromstring(self.zip.read(name), parser)

    def _rels(self, name) -> dict[str, tuple[str, str]]:
        root = self._xml(name)
        out = {}
        if root is None:
            return out
        base = os.path.dirname(name.replace("_rels/", "").replace(".rels", ""))
        for rel in root:
            target = rel.get("Target", "")
            mode = rel.get("TargetMode", "")
            if mode != "External":
                target = os.path.normpath(os.path.join(base, target)).replace("\\", "/")
                if target.startswith("/"):
                    target = target[1:]
            out[rel.get("Id")] = (target, rel.get("Type", ""))
        return out

    def part_rels(self, part_name: str):
        d, b = os.path.split(part_name)
        return self._rels(f"{d}/_rels/{b}.rels")

    def read(self, name):
        return self.zip.read(name) if name in self.names else None

    def note_text(self, kind: str, note_id: str | None) -> str:
        """각주/미주 내용 (문단은 줄바꿈으로 이음)"""
        cache = self.__dict__.setdefault("_notes", {})
        if kind not in cache:
            notes = {}
            target = next((t for t, typ in self.rels.values() if typ.endswith("/" + kind + "s")), None)
            root = self._xml(target) if target else self._xml(f"word/{kind}s.xml")
            if root is not None:
                for n in root.findall(w(kind)):
                    if wattr(n, "type") in ("separator", "continuationSeparator", "continuationNotice"):
                        continue
                    paras = []
                    for p in n.iter(w("p")):
                        txt = "".join((t.text or "") if t.tag == w("t") else "\t"
                                      for t in p.iter(w("t"), w("tab"))
                                      if not any(a.tag == w("del") for a in t.iterancestors()))
                        paras.append(_INVALID_XML.sub("", txt))
                    notes[wattr(n, "id")] = "\n".join(paras).strip()
            cache[kind] = notes
        return cache[kind].get(note_id, "")

    def _theme_fonts(self):
        fonts = {}
        theme_name = next((n for n in self.names if re.match(r"word/theme/theme\d*\.xml$", n)), None)
        root = self._xml(theme_name) if theme_name else None
        if root is None:
            return fonts
        for kind in ("major", "minor"):
            el = root.find(f".//{{{A}}}{kind}Font")
            if el is None:
                continue
            latin = el.find(f"{{{A}}}latin")
            ea = el.find(f"{{{A}}}ea")
            hang = None
            for f in el.findall(f"{{{A}}}font"):
                if f.get("script") == "Hang":
                    hang = f.get("typeface")
            fonts[f"{kind}HAnsi"] = latin.get("typeface") if latin is not None else None
            fonts[f"{kind}Ascii"] = fonts[f"{kind}HAnsi"]
            fonts[f"{kind}EastAsia"] = (ea.get("typeface") if ea is not None and ea.get("typeface") else None) or hang
            fonts[f"{kind}Bidi"] = None
        return fonts

    # ---- 스타일 체인 -----------------------------------------------------
    def style_chain(self, sid: str | None) -> list[etree._Element]:
        chain, seen = [], set()
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            st = self.styles[sid]
            chain.append(st)
            based = st.find(w("basedOn"))
            sid = wattr(based, "val")
        return list(reversed(chain))  # 기반 스타일이 먼저

    # ---- 번호 매기기 -----------------------------------------------------
    def numbering_level(self, num_id: str, ilvl: int):
        root = self.numbering_root
        if root is None:
            return None
        num = None
        for n in root.findall(w("num")):
            if wattr(n, "numId") == num_id:
                num = n
                break
        if num is None:
            return None
        abs_id = wattr(num.find(w("abstractNumId")), "val")
        lvl = None
        for ov in num.findall(w("lvlOverride")):
            if wattr(ov, "ilvl") == str(ilvl) and ov.find(w("lvl")) is not None:
                lvl = ov.find(w("lvl"))
        if lvl is None:
            for an in root.findall(w("abstractNum")):
                if wattr(an, "abstractNumId") == abs_id:
                    for l in an.findall(w("lvl")):
                        if wattr(l, "ilvl") == str(ilvl):
                            lvl = l
                    # numStyleLink 따라가기는 생략
                    break
        return abs_id, lvl


# ---------------------------------------------------------------------------
# 서식 해석
# ---------------------------------------------------------------------------
def apply_rpr(fmt: dict, rpr) -> None:
    if rpr is None:
        return
    for tag, key in (("b", "bold"), ("i", "italic"), ("strike", "strike"), ("dstrike", "strike"),
                     ("vanish", "hidden"), ("caps", "caps")):
        v = on_off(rpr.find(w(tag)))
        if v is not None:
            fmt[key] = v
    u = rpr.find(w("u"))
    if u is not None:
        fmt["underline"] = (wattr(u, "val", "single") or "single") != "none"
    c = rpr.find(w("color"))
    if c is not None:
        fmt["color"] = hex_color(wattr(c, "val"))
    sz = rpr.find(w("sz"))
    if sz is not None and wattr(sz, "val"):
        try:
            fmt["size"] = int(float(wattr(sz, "val"))) / 2
        except ValueError:
            pass
    hl = rpr.find(w("highlight"))
    if hl is not None:
        fmt["highlight"] = HIGHLIGHT.get(wattr(hl, "val") or "", None)
    shd = rpr.find(w("shd"))
    if shd is not None and fmt.get("highlight") is None:
        fill = hex_color(wattr(shd, "fill"))
        if fill and fill not in ("#FFFFFF",):
            fmt["highlight"] = fill
    va = rpr.find(w("vertAlign"))
    if va is not None:
        v = wattr(va, "val")
        fmt["script"] = "sup" if v == "superscript" else "sub" if v == "subscript" else None
    sc = rpr.find(w("w"))
    if sc is not None and wattr(sc, "val"):
        try:
            fmt["ratio"] = int(wattr(sc, "val"))
        except ValueError:
            pass
    rf = rpr.find(w("rFonts"))
    if rf is not None:
        for key, attrs in (("font_latin", ("ascii", "hAnsi")), ("font_ea", ("eastAsia",))):
            for a in attrs:
                if wattr(rf, a):
                    fmt[key] = ("", wattr(rf, a))
                    break
                if wattr(rf, a + "Theme"):
                    fmt[key] = ("theme", wattr(rf, a + "Theme"))
                    break


def apply_ppr(fmt: ParaFmt, ppr) -> None:
    if ppr is None:
        return
    jc = ppr.find(w("jc"))
    if jc is not None:
        v = wattr(jc, "val", "left")
        fmt.align = {"left": "LEFT", "start": "LEFT", "center": "CENTER", "right": "RIGHT",
                     "end": "RIGHT", "both": "JUSTIFY", "distribute": "DISTRIBUTE",
                     "lowKashida": "JUSTIFY", "mediumKashida": "JUSTIFY",
                     "highKashida": "JUSTIFY", "thaiDistribute": "DISTRIBUTE"}.get(v, "LEFT")
    ind = ppr.find(w("ind"))
    if ind is not None:
        def iv(*names):
            for n in names:
                val = wattr(ind, n)
                if val is not None:
                    try:
                        return int(float(val))
                    except ValueError:
                        return None
            return None
        v = iv("left", "start")
        if v is not None:
            fmt.left = v
        v = iv("right", "end")
        if v is not None:
            fmt.right = v
        fl, hg = iv("firstLine"), iv("hanging")
        if hg is not None:
            fmt.first_line = -hg
        elif fl is not None:
            fmt.first_line = fl
    sp = ppr.find(w("spacing"))
    if sp is not None:
        for attr, key in (("before", "before"), ("after", "after")):
            val = wattr(sp, attr)
            auto = wattr(sp, attr + "Autospacing")
            if auto in ("1", "true", "on"):
                setattr(fmt, key, 280)
            elif val is not None:
                try:
                    setattr(fmt, key, int(float(val)))
                except ValueError:
                    pass
        if wattr(sp, "line"):
            try:
                fmt.line = int(float(wattr(sp, "line")))
                fmt.line_rule = wattr(sp, "lineRule", "auto")
            except ValueError:
                pass
    for tag, key in (("keepNext", "keep_next"), ("keepLines", "keep_lines"),
                     ("pageBreakBefore", "page_break_before"), ("contextualSpacing", "contextual")):
        v = on_off(ppr.find(w(tag)))
        if v is not None:
            setattr(fmt, key, v)
    tabs = ppr.find(w("tabs"))
    if tabs is not None:
        fmt.tabs = dict(fmt.tabs)
        for t in tabs.findall(w("tab")):
            try:
                pos = int(float(wattr(t, "pos", "0")))
            except ValueError:
                continue
            val = wattr(t, "val", "left")
            if val == "clear":
                fmt.tabs.pop(pos, None)
                continue
            kind = {"center": "CENTER", "right": "RIGHT", "end": "RIGHT", "decimal": "DECIMAL"}.get(val, "LEFT")
            leader = {"dot": "DOT", "middleDot": "DOT", "hyphen": "DASH", "underscore": "SOLID",
                      "heavy": "SOLID"}.get(wattr(t, "leader", "none"), "NONE")
            fmt.tabs[pos] = (kind, leader)
    numpr = ppr.find(w("numPr"))
    if numpr is not None:
        nid = wattr(numpr.find(w("numId")), "val")
        lvl = wattr(numpr.find(w("ilvl")), "val")
        if nid is not None:
            fmt.num_id = nid
        if lvl is not None:
            fmt.num_level = int(lvl)


# ---------------------------------------------------------------------------
# 번호 매기기 문자열
# ---------------------------------------------------------------------------
def _roman(n: int) -> str:
    vals = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
            (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    out = ""
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out


def _letters(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


GANADA = "가나다라마바사아자차카타파하"
CHOSUNG = "ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ"
CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"


def format_number(n: int, fmt: str) -> str:
    if fmt in ("decimal", "decimalZero"):
        return f"{n:02d}" if fmt == "decimalZero" else str(n)
    if fmt == "upperRoman":
        return _roman(n)
    if fmt == "lowerRoman":
        return _roman(n).lower()
    if fmt == "upperLetter":
        return _letters(n)
    if fmt == "lowerLetter":
        return _letters(n).lower()
    if fmt in ("ganada", "koreanDigital"):
        return GANADA[(n - 1) % len(GANADA)]
    if fmt == "chosung":
        return CHOSUNG[(n - 1) % len(CHOSUNG)]
    if fmt in ("decimalEnclosedCircle", "decimalEnclosedCircleChinese"):
        return CIRCLED[n - 1] if 0 < n <= len(CIRCLED) else str(n)
    if fmt == "none":
        return ""
    return str(n)


class Numberer:
    def __init__(self, pkg: DocxPackage):
        self.pkg = pkg
        self.counters: dict[str, list[int]] = {}

    def label(self, num_id: str, ilvl: int):
        """(label 텍스트, 단계 pPr, 단계 rPr) 반환"""
        if num_id in (None, "0"):
            return None
        res = self.pkg.numbering_level(num_id, ilvl)
        if not res or res[1] is None:
            return None
        abs_id, lvl = res
        key = abs_id or num_id
        counts = self.counters.setdefault(key, [0] * 9)
        start = int(wattr(lvl.find(w("start")), "val", "1") or 1)
        if counts[ilvl] == 0:
            counts[ilvl] = start
        else:
            counts[ilvl] += 1
        for i in range(ilvl + 1, 9):
            counts[i] = 0
        num_fmt = wattr(lvl.find(w("numFmt")), "val", "decimal")
        text = wattr(lvl.find(w("lvlText")), "val", "") or ""
        if num_fmt == "bullet":
            label = text if text and text not in ("", "", "", "", "o") else "•"
            label = {"": "•", "": "▪", "": "➢", "": "✓", "o": "◦"}.get(text, label)
        else:
            def repl(m):
                k = int(m.group(1)) - 1
                lv_res = self.pkg.numbering_level(num_id, k)
                lfmt = num_fmt
                if lv_res and lv_res[1] is not None:
                    lfmt = wattr(lv_res[1].find(w("numFmt")), "val", "decimal")
                return format_number(counts[k] or 1, lfmt)
            label = re.sub(r"%(\d)", repl, text)
        suff = wattr(lvl.find(w("suff")), "val", "tab")
        sep = {"tab": "\t", "space": " ", "nothing": ""}.get(suff, "\t")
        return label + sep, lvl.find(w("pPr")), lvl.find(w("rPr"))


# ---------------------------------------------------------------------------
# 변환기
# ---------------------------------------------------------------------------
class Converter:
    def __init__(self, src: str, log: Callable[[str], None] | None = None):
        from hwpx import HwpxDocument  # 지연 import
        self.pkg = DocxPackage(src)
        self.doc = HwpxDocument.new()
        self.log = log or (lambda m: None)
        self.numberer = Numberer(self.pkg)
        self._char_cache: dict[RunFmt, str] = {}
        self._para_cache: dict[tuple, str] = {}
        self._fill_cache: dict[tuple, str] = {}
        self._first_para_used = False
        self._pending_break = False
        self.warnings: list[str] = []
        self.page_num_fmt: dict[int, tuple] = {}
        self._est = None   # 표 셀 높이 추정용 수집 목록
        self._pending_boxes: list = []   # 문단 뒤에 넣을 글상자
        self._hf = None   # 지금 채우는 머리말/꼬리말
        self.base_para_pr = self.doc.paragraphs[0].para_pr_id_ref or "0"

    # ---------------- 서식 → HWPX id -----------------------------------
    def resolve_font(self, spec, kind: str) -> str | None:
        if spec is None:
            return None
        src, name = spec
        if src == "theme":
            name = self.pkg.theme_fonts.get(name)
        return norm_font(name)

    def run_fmt(self, para_style: str | None, rpr_list) -> RunFmt:
        d: dict = {}
        apply_rpr(d, self.pkg.doc_rpr)
        for st in self.pkg.style_chain(para_style or self.pkg.default_para_style):
            apply_rpr(d, st.find(w("rPr")))
        for rpr in rpr_list:
            if rpr is None:
                continue
            rs = rpr.find(w("rStyle"))
            if rs is not None:
                for st in self.pkg.style_chain(wattr(rs, "val")):
                    apply_rpr(d, st.find(w("rPr")))
            apply_rpr(d, rpr)
        ea = self.resolve_font(d.pop("font_ea", None), "ea")
        latin = self.resolve_font(d.pop("font_latin", None), "latin")
        if ea is None:
            ea = norm_font(self.pkg.theme_fonts.get("minorEastAsia")) or latin or "맑은 고딕"
        if latin is None:
            latin = ea
        return RunFmt(font_ea=ea, font_latin=latin, **d)

    def char_pr(self, f: RunFmt) -> str:
        if f in self._char_cache:
            return self._char_cache[f]
        fonts = {}
        if f.font_ea:
            for lang in ("HANGUL", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER"):
                fonts[lang] = f.font_ea
        if f.font_latin:
            fonts["LATIN"] = f.font_latin
        try:
            cid = self.doc.styles.ensure_run(
                bold=f.bold, italic=f.italic, underline=f.underline, strike=f.strike,
                color=f.color or "#000000", size=max(1.0, f.size), highlight=f.highlight,
                script=f.script, ratio=f.ratio if f.ratio and 50 <= f.ratio <= 200 else None,
                font_by_lang=fonts or None,
            )
        except Exception as e:  # 서식 하나가 실패해도 계속
            self.warnings.append(f"글자 서식 일부 생략: {e}")
            cid = self.doc.styles.ensure_run(bold=f.bold, italic=f.italic, size=max(1.0, f.size))
        self._char_cache[f] = cid
        return cid

    def para_fmt(self, ppr, style_id: str | None, extra_ppr=None, table_style: str | None = None) -> ParaFmt:
        pf = ParaFmt()
        apply_ppr(pf, self.pkg.doc_ppr)
        if table_style:
            for st in self.pkg.style_chain(table_style):
                apply_ppr(pf, st.find(w("pPr")))
        chain = self.pkg.style_chain(style_id or self.pkg.default_para_style)
        for st in chain:
            apply_ppr(pf, st.find(w("pPr")))
        if extra_ppr is not None:      # 번호 단계 들여쓰기
            apply_ppr(pf, extra_ppr)
        apply_ppr(pf, ppr)
        return pf

    def line_percent(self, pf: ParaFmt, font_size: float, font: str | None) -> int:
        # Word 의 1줄 높이(글꼴마다 다름)를 한글 비율 줄간격으로 환산
        factor = 1.2 * font_line_ratio(font)
        if pf.line is None:
            pct = 100 * factor
        elif pf.line_rule == "auto":
            pct = pf.line / 240 * 100 * factor
        else:  # exact / atLeast  (twips)
            pt = pf.line / 20
            pct = pt / max(font_size, 1) * 100
            if pf.line_rule == "atLeast":
                pct = max(pct, 100 * factor)
        return int(min(max(round(pct), 60), 500))

    def para_pr(self, pf: ParaFmt, font_size: float, font: str | None, align_override=None) -> str:
        align = align_override or pf.align or "LEFT"
        first = pf.first_line
        left = pf.left
        if first < 0:           # 한글의 내어쓰기: 첫 줄이 왼쪽 여백에서 시작
            left = max(0, left + first)
        before = pf.before
        after = pf.after
        # 탭 위치: Word 는 쪽 여백 기준, 한글은 문단 왼쪽 여백 기준
        tabs = tuple(sorted((p - left, k, l) for p, (k, l) in pf.tabs.items() if p - left > 0))
        key = (align, left, pf.right, first, before, after,
               self.line_percent(pf, font_size, font), pf.keep_next, pf.keep_lines, pf.page_break_before, tabs)
        if key in self._para_cache:
            return self._para_cache[key]
        hdr = self.doc.oxml.headers[0]
        tab_id = None
        if tabs:
            try:
                tab_id = hdr.ensure_tab_definition(
                    tab_stops=[{"pos": p * TWIP_TO_HWP, "type": k, "leader": l} for p, k, l in tabs])
            except Exception as e:
                self.warnings.append(f"탭 설정 생략: {e}")
        pid = hdr.ensure_paragraph_format(
            tab_pr_id_ref=tab_id,
            base_para_pr_id=self.base_para_pr,
            alignment=align,
            line_spacing_percent=key[6],
            margins={"left": left * TWIP_TO_HWP, "right": pf.right * TWIP_TO_HWP,
                     "intent": first * TWIP_TO_HWP, "prev": before * TWIP_TO_HWP,
                     "next": after * TWIP_TO_HWP},
            break_setting={"keep_with_next": pf.keep_next, "keep_lines": pf.keep_lines,
                           "page_break_before": pf.page_break_before},
        )
        self._para_cache[key] = pid
        return pid

    # ---------------- 문단 내용 수집 -----------------------------------
    def collect_runs(self, p_el, para_style, part_rels):
        """문단 안의 (종류, 값, RunFmt) 목록. 종류: text / image / pagebreak / pagenum"""
        items = []
        state = {"field": None}   # None / "instr" / "result"
        instr = []

        def walk(el, rpr_stack):
            tag = etree.QName(el).localname if isinstance(el.tag, str) else ""
            ns = etree.QName(el).namespace if isinstance(el.tag, str) else ""
            if ns == MC and tag == "AlternateContent":
                fb = el.find(f"{{{MC}}}Fallback")
                ch = el.find(f"{{{MC}}}Choice")
                target = ch if ch is not None else fb
                if target is not None:
                    for c in target:
                        walk(c, rpr_stack)
                return
            if ns == M and tag in ("oMathPara", "oMath"):
                items.append(("equation", el, self.run_fmt(para_style, rpr_stack)))
                return
            if ns != W:
                return
            if tag in ("pPr", "del", "moveFrom", "proofErr", "bookmarkStart", "bookmarkEnd",
                       "commentRangeStart", "commentRangeEnd", "permStart", "permEnd"):
                return
            if tag == "fldSimple":
                ins = (wattr(el, "instr") or "").strip().split()
                fk = PAGE_FIELDS.get(ins[0].upper()) if ins else None
                if fk:
                    items.append((fk, None, self.run_fmt(para_style, rpr_stack)))
                    return
                for c in el:
                    walk(c, rpr_stack)
                return
            if tag in ("hyperlink", "ins", "smartTag", "sdt", "sdtContent", "customXml", "moveTo", "dir", "bdo"):
                for c in el:
                    if isinstance(c.tag, str) and etree.QName(c).localname in ("sdtPr", "sdtEndPr", "smartTagPr", "customXmlPr"):
                        continue
                    walk(c, rpr_stack)
                return
            if tag == "r":
                rpr = el.find(w("rPr"))
                fmt = self.run_fmt(para_style, rpr_stack + [rpr])
                for c in el:
                    if not isinstance(c.tag, str):
                        continue
                    ctag = etree.QName(c).localname
                    if ctag == "fldChar":
                        t = wattr(c, "fldCharType")
                        if t == "begin":
                            state["field"] = "instr"
                            instr.clear()
                        elif t == "separate":
                            code = " ".join(instr).strip().split()
                            fk = PAGE_FIELDS.get(code[0].upper()) if code else None
                            if fk:
                                items.append((fk, None, fmt))
                                state["field"] = "skip"
                            else:
                                state["field"] = "result"
                        elif t == "end":
                            if state["field"] == "instr":
                                code = " ".join(instr).strip().split()
                                fk = PAGE_FIELDS.get(code[0].upper()) if code else None
                                if fk:
                                    items.append((fk, None, fmt))
                            state["field"] = None
                        continue
                    if state["field"] in ("instr",):
                        if ctag == "instrText":
                            instr.append(c.text or "")
                        continue
                    if state["field"] == "skip":
                        continue
                    if fmt.hidden:
                        continue
                    if ctag == "t":
                        txt = _INVALID_XML.sub("", c.text or "")
                        if fmt.caps:
                            txt = txt.upper()
                        if "symbol" in (fmt.font_latin or "").lower():
                            txt = "".join(symbol_char(ch) for ch in txt)
                        items.append(("text", txt, fmt))
                    elif ctag == "tab":
                        items.append(("text", "\t", fmt))
                    elif ctag in ("br", "cr"):
                        if wattr(c, "type") == "page":
                            items.append(("pagebreak", None, fmt))
                        elif wattr(c, "type") == "column":
                            items.append(("pagebreak", None, fmt))
                        else:
                            items.append(("text", "\n", fmt))
                    elif ctag == "noBreakHyphen":
                        items.append(("text", "-", fmt))
                    elif ctag == "softHyphen":
                        pass
                    elif ctag == "sym":
                        ch = wattr(c, "char")
                        font = (wattr(c, "font") or "").lower()
                        try:
                            code = int(ch, 16)
                            if code >= 0xF000:
                                code -= 0xF000
                            s = chr(code)
                            if "symbol" in font:
                                s = symbol_char(s)
                            elif "wingdings" in font:
                                s = WINGDINGS.get(code, "•")
                            items.append(("text", s, fmt))
                        except (TypeError, ValueError):
                            pass
                    elif ctag in ("drawing", "pict", "object", "AlternateContent"):
                        for kind, val in self.extract_objects(c, part_rels):
                            items.append((kind, val, fmt))
                    elif ctag in ("footnoteReference", "endnoteReference"):
                        note = self.pkg.note_text("footnote" if ctag == "footnoteReference" else "endnote",
                                                  wattr(c, "id"))
                        if note:
                            items.append(("footnote" if ctag == "footnoteReference" else "endnote", note, fmt))
                return
            for c in el:
                walk(c, rpr_stack)

        for c in p_el:
            walk(c, [])
        return items

    def extract_objects(self, el, part_rels):
        """그림·글상자·차트 등 개체 → [(종류, 값)]"""
        tag = etree.QName(el).localname
        if tag == "AlternateContent":
            for branch in (el.find(f"{{{MC}}}Choice"), el.find(f"{{{MC}}}Fallback")):
                if branch is None:
                    continue
                res = []
                for c in branch:
                    if isinstance(c.tag, str):
                        res += self.extract_objects(c, part_rels)
                if res:
                    return res
            return []
        out = []
        # 글상자: 가장 바깥 글상자 내용만
        boxes = [t for t in el.iter(w("txbxContent"))
                 if not any(a.tag == w("txbxContent") for a in t.iterancestors())]
        for box in boxes:
            ext = None
            for anc in box.iterancestors():
                if isinstance(anc.tag, str) and etree.QName(anc).localname in ("inline", "anchor"):
                    ext = anc.find(f"{{{WP}}}extent")
                    break
            width = int(int(ext.get("cx")) / 635) if ext is not None and ext.get("cx") else None
            # 테두리: 선 없음(noFill)이면 테두리 없이
            no_line = any(etree.QName(n).localname == "noFill" and etree.QName(n.getparent()).localname == "ln"
                          for n in el.iter(f"{{{A}}}noFill"))
            if el.find(f".//{{{V}}}textbox") is not None:
                shape = el.find(f".//{{{V}}}textbox").getparent()
                no_line = no_line or (shape is not None and shape.get("stroked") in ("f", "false"))
            out.append(("textbox", (box, width, not no_line)))
        imgs = [i for i in self.extract_images(el, part_rels)]
        out += [("image", i) for i in imgs]
        if not out:
            uris = [g.get("uri", "") for g in el.iter(f"{{{A}}}graphicData")]
            if any("chart" in u for u in uris):
                out.append(("text", "[차트]"))
                self.warnings.append("차트는 그림으로 옮기지 못해 [차트] 표시로 남겼습니다.")
            elif any("diagram" in u for u in uris):
                texts = [t.text for t in el.iter(f"{{{A}}}t") if t.text]
                out.append(("text", " / ".join(texts) if texts else "[SmartArt]"))
                self.warnings.append("SmartArt 도형은 글자만 옮겼습니다.")
        return out

    def extract_images(self, el, part_rels):
        out = []
        for blip in el.iter(f"{{{A}}}blip"):
            if any(a.tag == w("txbxContent") for a in blip.iterancestors()):
                continue  # 글상자 안 그림은 글상자 내용으로 처리
            rid = blip.get(f"{{{R}}}embed")
            ext = None
            # 가장 가까운 extent 찾기
            anc = blip
            while anc is not None and ext is None:
                anc = anc.getparent()
                if anc is None:
                    break
                if etree.QName(anc).localname in ("inline", "anchor"):
                    ext = anc.find(f"{{{WP}}}extent")
            if rid and rid in part_rels:
                target = part_rels[rid][0]
                data = self.pkg.read(target)
                if data:
                    cx = int(ext.get("cx")) if ext is not None else None
                    cy = int(ext.get("cy")) if ext is not None else None
                    out.append((data, target, cx, cy))
        if not out:  # VML 그림 (예전 Word)
            for im in el.iter(f"{{{V}}}imagedata"):
                if any(a.tag == w("txbxContent") for a in im.iterancestors()):
                    continue
                rid = im.get(f"{{{R}}}id")
                if rid and rid in part_rels:
                    data = self.pkg.read(part_rels[rid][0])
                    if data:
                        out.append((data, part_rels[rid][0], None, None))
        return out

    # ---------------- 그림 -------------------------------------------
    def prepare_image(self, data: bytes, name: str):
        ext = os.path.splitext(name)[1].lower().lstrip(".")
        if ext in ("jpg", "jpeg", "png", "gif", "bmp"):
            return data, "jpg" if ext == "jpeg" else ext
        try:  # emf/wmf/tiff 등 → PNG
            from PIL import Image
            im = Image.open(io.BytesIO(data))
            buf = io.BytesIO()
            im.save(buf, "PNG")
            return buf.getvalue(), "png"
        except Exception:
            return None, None

    def image_size(self, data, cx, cy, max_w):
        if cx and cy:
            wv, hv = int(cx * EMU_TO_HWP), int(cy * EMU_TO_HWP)
        else:
            try:
                from PIL import Image
                im = Image.open(io.BytesIO(data))
                dpi = im.info.get("dpi", (96, 96))[0] or 96
                wv = int(im.width / dpi * 7200)
                hv = int(im.height / dpi * 7200)
            except Exception:
                wv, hv = 28800, 21600
        if max_w and wv > max_w:
            hv = int(hv * max_w / wv)
            wv = max_w
        return max(wv, 100), max(hv, 100)

    def add_image_to_paragraph(self, para, img, max_w):
        data, name, cx, cy = img
        data2, fmt = self.prepare_image(data, name)
        if data2 is None:
            self.warnings.append(f"지원하지 않는 그림 형식이라 생략: {os.path.basename(name)}")
            return False
        wv, hv = self.image_size(data2, cx, cy, max_w)
        item = self.doc.media.add_image(data2, fmt)
        item_id = getattr(item, "id", None) or getattr(item, "item_id", None) or str(item)
        para.add_picture(str(item_id), width=wv, height=hv)
        return True

    # ---------------- 문단 출력 ---------------------------------------
    def new_body_paragraph(self, para_pr, char_pr):
        if not self._first_para_used:
            self._first_para_used = True
            p = self.doc.paragraphs[0]
            p.para_pr_id_ref = para_pr
            try:
                p.char_pr_id_ref = char_pr
            except Exception:
                pass
            return p
        return self.doc.add_paragraph("", section=self.doc.sections[0], para_pr_id_ref=para_pr,
                                      char_pr_id_ref=char_pr, include_run=False, inherit_style=False)

    def emit_paragraph(self, p_el, factory, part_rels, max_w, table_style=None, is_cell=False):
        """p_el 을 1개 이상의 한글 문단으로 출력. factory(para_pr, char_pr) -> 새 문단"""
        ppr = p_el.find(w("pPr"))
        style_id = wattr(ppr.find(w("pStyle")) if ppr is not None else None, "val")
        pf_tmp = self.para_fmt(ppr, style_id, table_style=table_style)
        num_ppr = num_rpr = None
        label = None
        if pf_tmp.num_id and pf_tmp.num_id != "0":
            res = self.numberer.label(pf_tmp.num_id, pf_tmp.num_level)
            if res:
                label, num_ppr, num_rpr = res
        pf = self.para_fmt(ppr, style_id, extra_ppr=num_ppr, table_style=table_style) if num_ppr is not None else pf_tmp

        mark_rpr = ppr.find(w("rPr")) if ppr is not None else None
        mark_fmt = self.run_fmt(style_id, [mark_rpr])
        items = self.collect_runs(p_el, style_id, part_rels)
        if label:
            lab_fmt = self.run_fmt(style_id, [num_rpr]) if num_rpr is not None else mark_fmt
            first_text = next((f for k, v, f in items if k == "text" and v.strip()), None)
            if num_rpr is None and first_text is not None:
                lab_fmt = first_text
            # 기호 글꼴(Symbol/Wingdings)로 된 글머리표는 한글에서 □ 로 깨지므로 본문 글꼴 사용
            symbolic = ("symbol", "wingdings", "webdings")
            if any(s in (lab_fmt.font_latin or "").lower() or s in (lab_fmt.font_ea or "").lower() for s in symbolic):
                base = first_text or mark_fmt
                lab_fmt = replace(lab_fmt, font_latin=base.font_latin, font_ea=base.font_ea)
            items.insert(0, ("text", label, lab_fmt))

        if self._pending_break and not is_cell:
            pf.page_break_before = True
            self._pending_break = False

        # 페이지 나누기 기준으로 문단 쪼개기
        chunks, cur = [], []
        for it in items:
            if it[0] == "pagebreak":
                chunks.append(cur)
                cur = []
            else:
                cur.append(it)
        chunks.append(cur)

        paragraphs = []
        for ci, chunk in enumerate(chunks):
            if ci > 0 and not chunk and ci == len(chunks) - 1:
                # 문단 끝의 페이지 나누기 → 다음 문단에 적용
                if not is_cell:
                    self._pending_break = True
                break
            sizes = [f.size for k, v, f in chunk if k == "text" and v.strip()] or [mark_fmt.size]
            main_size = max(set(sizes), key=sizes.count)
            fonts = [f.font_ea for k, v, f in chunk if k == "text" and v.strip()] or [mark_fmt.font_ea]
            pf_c = replace(pf) if ci == 0 else replace(pf, page_break_before=True, before=0)
            # 수식만 있는 문단(Word 의 "디스플레이 수식")은 수식 정렬을 따름 (기본 가운데)
            eqs = [v for k, v, f in chunk if k == "equation"]
            others = [v for k, v, f in chunk if k == "text" and v.strip()]
            if eqs and not others and any(etree.QName(e).localname == "oMathPara" for e in eqs):
                jc = None
                for e in eqs:
                    j = e.find(f"{{{M}}}oMathParaPr/{{{M}}}jc")
                    if j is not None:
                        jc = j.get(f"{{{M}}}val")
                pf_c.align = {"left": "LEFT", "right": "RIGHT"}.get(jc or "", "CENTER")
            para_pr = self.para_pr(pf_c, main_size, fonts[0])
            first_fmt = next((f for k, v, f in chunk if k == "text"), mark_fmt)
            para = factory(para_pr, self.char_pr(first_fmt if chunk else mark_fmt))
            paragraphs.append(para)
            has_content = False
            buf, buf_fmt = "", None
            for kind, val, f in chunk:
                if kind == "text":
                    if buf_fmt is not None and f != buf_fmt:
                        para.add_run(buf, char_pr_id_ref=self.char_pr(buf_fmt))
                        has_content = True
                        buf = ""
                    buf += val
                    buf_fmt = f
                else:
                    if buf:
                        para.add_run(buf, char_pr_id_ref=self.char_pr(buf_fmt))
                        has_content = True
                        buf, buf_fmt = "", None
                    if kind == "image":
                        if self.add_image_to_paragraph(para, val, max_w):
                            has_content = True
                    elif kind in ("pagenum", "numpages"):
                        if self.add_page_field(para, kind, f):
                            has_content = True
                    elif kind == "equation":
                        if self.add_equation(para, val, f):
                            has_content = True
                    elif kind in ("footnote", "endnote"):
                        # 각주 본문은 위첨자가 아닌 보통 글자 (조금 작게)
                        note_fmt = replace(f, script=None, size=max(f.size - 1, 8))
                        try:
                            if kind == "footnote":
                                self.doc.notes.add_footnote(val, paragraph=para, char_pr_id_ref=self.char_pr(note_fmt))
                            else:
                                self.doc.notes.add_endnote(val, paragraph=para, char_pr_id_ref=self.char_pr(note_fmt))
                            has_content = True
                        except Exception as e:
                            para.add_run(f"[{'각주' if kind == 'footnote' else '미주'}: {val}]",
                                         char_pr_id_ref=self.char_pr(f))
                            self.warnings.append(f"각주를 본문 괄호로 옮김: {e}")
                            has_content = True
                    elif kind == "textbox":
                        self._pending_boxes.append(val)
            if buf or buf_fmt is not None:
                para.add_run(buf, char_pr_id_ref=self.char_pr(buf_fmt))
                has_content = True
            if not has_content:
                para.add_run("", char_pr_id_ref=self.char_pr(mark_fmt))
            if self._est is not None:
                text = "".join(v for k, v, f in chunk if k == "text")
                img_h = sum(self.image_size(v[0], v[2], v[3], max_w)[1] for k, v, f in chunk if k == "image")
                lat = next((f.font_latin for k, v, f in chunk if k == "text" and v.strip()), mark_fmt.font_latin)
                self._est.append((text, main_size, self.line_percent(pf_c, main_size, fonts[0]),
                                  pf_c.before, pf_c.after, img_h, fonts[0], lat,
                                  (pf_c.left + max(pf_c.first_line, 0) + pf_c.right) / 20))
        return paragraphs

    def add_equation(self, para, el, fmt) -> bool:
        """Word 수식 → 한글 수식(고칠 수 있는 수식 개체). 안 되면 글자로."""
        script = omml_to_eqedit(el)
        if script:
            try:
                para.add_equation(script, base_unit=int(max(fmt.size, 5) * 100),
                                  char_pr_id_ref=self.char_pr(fmt))
                return True
            except Exception as e:
                self.warnings.append(f"수식 일부를 글자로 옮김: {e}")
        text = omml_plain_text(el)
        if text:
            para.add_run(text, char_pr_id_ref=self.char_pr(replace(fmt, italic=True)))
            return True
        return False

    def textbox_table(self, info):
        """글상자 → 테두리 있는 1칸 표 (w:tbl 를 만들어 표 변환에 넘김)"""
        box, width, bordered = info
        width = width or 6000
        tbl = etree.Element(w("tbl"))
        pr = etree.SubElement(tbl, w("tblPr"))
        b = etree.SubElement(pr, w("tblBorders"))
        for side in ("top", "left", "bottom", "right"):
            e = etree.SubElement(b, w(side))
            e.set(w("val"), "single" if bordered else "nil")
            e.set(w("sz"), "6")
            e.set(w("color"), "000000")
        mar = etree.SubElement(pr, w("tblCellMar"))
        for side, v in (("top", 72), ("bottom", 72), ("left", 144), ("right", 144)):
            e = etree.SubElement(mar, w(side))
            e.set(w("w"), str(v))
            e.set(w("type"), "dxa")
        grid = etree.SubElement(tbl, w("tblGrid"))
        etree.SubElement(grid, w("gridCol")).set(w("w"), str(width))
        tc = etree.SubElement(etree.SubElement(tbl, w("tr")), w("tc"))
        for child in box:
            if isinstance(child.tag, str):
                tc.append(copy.deepcopy(child))
        if tc.find(w("p")) is None and tc.find(w("tbl")) is None:
            etree.SubElement(tc, w("p"))
        return tbl

    def flush_textboxes(self, host, section, part_rels, max_w):
        boxes, self._pending_boxes = self._pending_boxes, []
        heights = 0
        for info in boxes:
            try:
                tbl = self.textbox_table(info)
                if host is None:
                    self.emit_table_in_section(tbl, section, part_rels, max_w)
                else:
                    heights += self.emit_table(tbl, host, part_rels, max_w) or 0
            except Exception as e:
                self.warnings.append(f"글상자 변환 실패: {e}")
        return heights

    def estimate_height(self, est, width_hwp: int) -> int:
        """셀 내용 높이 추정 (HWPUNIT)"""
        import math
        total = 0.0
        avail_pt = max(width_hwp / 100, 20)
        for item in est:
            text, size, pct, before, after, img_h = item[:6]
            ea, lat, indent = (item[6], item[7], item[8]) if len(item) > 8 else (None, None, 0)
            # 실제 글꼴 너비로 한글의 줄 나눔을 흉내 냄
            # (한글은 파일에 저장된 행 높이를 그대로 쓰므로 줄 수가 정확해야 글자가 칸을 넘지 않음)
            lines = count_lines(text, size, max(avail_pt - indent, 10) * 0.98, ea, lat) if size else 0
            h = lines * size * pct / 100 if size else 0
            if img_h:
                h = max(h, img_h / 100 + size * 0.3)
            total += h + (before + after) / 20
        return int(total * 100)

    # ---------------- 표 ---------------------------------------------
    def border_of(self, el):
        """w:top 등 → (켜짐, 색, 두께 mm)"""
        if el is None:
            return None
        val = wattr(el, "val", "single")
        if val in ("nil", "none"):
            return (False, None, 0)
        try:
            sz = int(wattr(el, "sz", "4"))
        except ValueError:
            sz = 4
        mm = sz / 8 * 0.3528
        kind = {"double": "DOUBLE_SLIM", "triple": "SLIM_THICK_SLIM", "dotted": "DOT", "dashed": "DASH",
                "dashSmallGap": "DASH", "dotDash": "DASH_DOT", "dotDotDash": "DASH_DOT_DOT",
                "thinThickSmallGap": "THICK_SLIM", "thickThinSmallGap": "SLIM_THICK",
                "wave": "WAVE", "doubleWave": "WAVE"}.get(val, "SOLID")
        return (True, hex_color(wattr(el, "color")) or "#000000", mm, kind)

    def table_borders(self, tbl_pr, table_style):
        b = {}
        def take(container):
            if container is None:
                return
            for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
                el = container.find(w(side))
                if el is None and side == "left":
                    el = container.find(w("start"))
                if el is None and side == "right":
                    el = container.find(w("end"))
                r = self.border_of(el)
                if r is not None:
                    b[side] = r
        for st in self.pkg.style_chain(table_style):
            tp = st.find(w("tblPr"))
            if tp is not None:
                take(tp.find(w("tblBorders")))
        if tbl_pr is not None:
            take(tbl_pr.find(w("tblBorders")))
        return b

    @staticmethod
    def read_margins(el) -> dict:
        out = {}
        if el is None:
            return out
        for side, alts in (("left", ("left", "start")), ("right", ("right", "end")),
                           ("top", ("top",)), ("bottom", ("bottom",))):
            for a in alts:
                m = el.find(w(a))
                if m is not None and wattr(m, "type", "dxa") in ("dxa", None):
                    try:
                        out[side] = max(0, int(float(wattr(m, "w", "0"))))
                    except ValueError:
                        pass
                    break
        return out

    def border_fill_id(self, sides: dict, fill: str | None):
        active = [s for s in ("left", "right", "top", "bottom") if sides.get(s) and sides[s][0]]
        color = next((sides[s][1] for s in active), "#000000")
        width_mm = max([sides[s][2] for s in active] or [0.12])
        width = min(HWP_LINE_WIDTHS, key=lambda x: abs(x - width_mm))
        btype = next((sides[s][3] for s in active if len(sides[s]) > 3), "SOLID")
        key = (tuple(active), color, width, fill, btype)
        if key in self._fill_cache:
            return self._fill_cache[key]
        kwargs = dict(border_color=color, border_width=f"{width} mm", active_borders=active, border_type=btype)
        if fill:
            kwargs["fill_color"] = fill
        if not active:
            kwargs["border_type"] = "SOLID"
            kwargs["active_borders"] = []
        try:
            bid = self.doc.styles.ensure_border_fill(**kwargs)
        except Exception as e:
            self.warnings.append(f"표 테두리 일부 생략: {e}")
            bid = self.doc.styles.ensure_border_fill(border_color=color)
        self._fill_cache[key] = bid
        return bid

    def emit_table(self, tbl_el, host, part_rels, max_w):
        """host: None(본문) 또는 셀 객체"""
        tbl_pr = tbl_el.find(w("tblPr"))
        tstyle = wattr(tbl_pr.find(w("tblStyle")) if tbl_pr is not None else None, "val") or self.pkg.default_table_style
        grid = [int(float(wattr(g, "w", "0") or 0)) for g in tbl_el.findall(f"{w('tblGrid')}/{w('gridCol')}")]
        rows = tbl_el.findall(w("tr"))
        if not rows:
            return
        # 셀 배치 계산
        layout = []   # (r, c, tc, colspan)
        vmerge_open: dict[int, list] = {}
        merges = []
        ncols = len(grid)
        for r, tr in enumerate(rows):
            c = 0
            trpr = tr.find(w("trPr"))
            gb = wattr(trpr.find(w("gridBefore")) if trpr is not None else None, "val")
            if gb:
                c += int(gb)
            for tc in tr.findall(w("tc")):
                tcpr = tc.find(w("tcPr"))
                span = int(wattr(tcpr.find(w("gridSpan")) if tcpr is not None else None, "val", "1") or 1)
                vm = tcpr.find(w("vMerge")) if tcpr is not None else None
                vval = wattr(vm, "val", "continue") if vm is not None else None
                if vval == "continue" and c in vmerge_open:
                    vmerge_open[c][2] = r          # 아래로 늘리기
                else:
                    entry = [r, c, r, c + span - 1, tc]
                    if vval == "restart":
                        vmerge_open[c] = entry
                    else:
                        vmerge_open.pop(c, None)
                    layout.append(entry)
                c += span
            ncols = max(ncols, c)
        avail_tw = int((max_w or 42520) / TWIP_TO_HWP)
        if not grid or len(grid) < ncols or sum(grid) < 100 * ncols:
            # 열 너비 정보가 없거나 이상함 → 첫 줄 칸 너비(tcW) 또는 균등 분배
            widths = []
            for tc in rows[0].findall(w("tc")):
                tcw = tc.find(f"{w('tcPr')}/{w('tcW')}")
                span = int(wattr(tc.find(f"{w('tcPr')}/{w('gridSpan')}"), "val", "1") or 1)
                v = 0
                if tcw is not None and wattr(tcw, "type", "dxa") == "dxa":
                    try:
                        v = int(float(wattr(tcw, "w", "0")))
                    except ValueError:
                        v = 0
                widths += [max(v // span, 0)] * span
            if len(widths) == ncols and all(x >= 200 for x in widths):
                grid = widths
            else:
                grid = [max(avail_tw // max(ncols, 1), 300)] * ncols
        grid = [max(g, 150) for g in grid]
        nrows = len(rows)
        total_w = sum(grid) * TWIP_TO_HWP
        if max_w:
            total_w = min(total_w, max_w)
        total_w = max(total_w, 1000 * ncols)

        borders = self.table_borders(tbl_pr, tstyle)
        # 셀 안쪽 여백 (Word 기본: 좌우 0.19cm, 위아래 0)
        tbl_mar = {"left": 108, "right": 108, "top": 0, "bottom": 0}
        for st_pr in [st.find(w("tblPr")) for st in self.pkg.style_chain(tstyle)] + [tbl_pr]:
            tbl_mar.update(self.read_margins(st_pr.find(w("tblCellMar")) if st_pr is not None else None))
        if host is None:
            anchor_pf = ParaFmt(align="LEFT")
            anchor_pr = self.para_pr(anchor_pf, 10, None)
            if self._pending_break:
                anchor_pr = self.para_pr(replace(anchor_pf, page_break_before=True), 10, None)
                self._pending_break = False
            table = self.doc.add_table(nrows, ncols, width=total_w, para_pr_id_ref=anchor_pr)
        else:
            table = host.add_table(nrows, ncols, width=total_w)
        try:
            table.set_column_widths(grid)
        except Exception:
            pass
        # 표 위치: 글자처럼 취급하지 않음 → 쪽 경계에서 나뉠 수 있음 (Word 와 같음)
        jc = wattr(tbl_pr.find(w("jc")) if tbl_pr is not None else None, "val")
        pos = table.element.find(f"{{{HP_NS}}}pos")
        if pos is not None:
            pos.set("treatAsChar", "0")
            pos.set("affectLSpacing", "0")
            pos.set("flowWithText", "1")
            pos.set("allowOverlap", "0")
            pos.set("vertRelTo", "PARA")
            pos.set("vertAlign", "TOP")
            pos.set("vertOffset", "0")
            pos.set("horzRelTo", "COLUMN")
            pos.set("horzAlign", {"center": "CENTER", "right": "RIGHT", "end": "RIGHT"}.get(jc, "LEFT"))
            pos.set("horzOffset", "0")
        table.element.set("textWrap", "TOP_AND_BOTTOM")
        # 쪽 경계에서 "나눔"(행 단위): 행 하나가 두 쪽으로 쪼개지지 않고 통째로 다음 쪽으로 감 (Word 와 같은 모양)
        table.element.set("pageBreak", "TABLE")
        # 머리글 행 반복
        header_rows = set()
        for r, tr in enumerate(rows):
            th = tr.find(f"{w('trPr')}/{w('tblHeader')}")
            if th is not None and on_off(th):
                header_rows.add(r)
            else:
                break
        if header_rows:
            table.element.set("repeatHeader", "1")

        # 셀 채우기
        cell_need: dict = {}
        for (r0, c0, r1, c1, tc) in layout:
            if r0 >= nrows or c0 >= ncols:
                continue
            cell = table.cell(r0, c0)
            if r0 in header_rows:
                cell.element.set("header", "1")
            tcpr = tc.find(w("tcPr"))
            # 테두리 & 배경
            sides = {
                "top": borders.get("top") if r0 == 0 else borders.get("insideH"),
                "bottom": borders.get("bottom") if r1 >= nrows - 1 else borders.get("insideH"),
                "left": borders.get("left") if c0 == 0 else borders.get("insideV"),
                "right": borders.get("right") if c1 >= ncols - 1 else borders.get("insideV"),
            }
            fill = None
            if tcpr is not None:
                tcb = tcpr.find(w("tcBorders"))
                if tcb is not None:
                    for side, alt in (("top", None), ("bottom", None), ("left", "start"), ("right", "end")):
                        el = tcb.find(w(side))
                        if el is None and alt:
                            el = tcb.find(w(alt))
                        bres = self.border_of(el)
                        if bres is not None:
                            sides[side] = bres
                shd = tcpr.find(w("shd"))
                if shd is not None:
                    fill = hex_color(wattr(shd, "fill"))
            try:
                cell.element.set("borderFillIDRef", self.border_fill_id(sides, fill))
            except Exception as e:
                self.warnings.append(f"셀 테두리 생략: {e}")
            # 안쪽 여백
            mar = dict(tbl_mar)
            if tcpr is not None:
                mar.update(self.read_margins(tcpr.find(w("tcMar"))))
            try:
                from hwpx.oxml.table_sizes import set_cell_margins
                set_cell_margins(cell, **{k: int(v * TWIP_TO_HWP) for k, v in mar.items()})
            except Exception as e:
                self.warnings.append(f"셀 여백 생략: {e}")
            # 세로 정렬
            va = wattr(tcpr.find(w("vAlign")) if tcpr is not None else None, "val")
            sub = cell.element.find("{http://www.hancom.co.kr/hwpml/2011/paragraph}subList")
            if sub is not None:
                sub.set("vertAlign", {"center": "CENTER", "bottom": "BOTTOM"}.get(va or "top", "TOP"))
            # 내용
            cell_w = (sum(grid[c0:c1 + 1]) - mar["left"] - mar["right"]) * TWIP_TO_HWP
            content_h = self.fill_cell(cell, tc, part_rels, cell_w, tstyle)
            need = content_h + int((mar["top"] + mar["bottom"]) * TWIP_TO_HWP) + 100
            cell_need[(r0, c0)] = (r1, need)

        # 원본에 칸이 없는 자리(gridBefore/After 등)는 테두리 없이
        covered = set()
        for (r0, c0, r1, c1, tc) in layout:
            for rr in range(r0, r1 + 1):
                for cc in range(c0, c1 + 1):
                    covered.add((rr, cc))
        empty_fill = None
        for rr in range(nrows):
            for cc in range(ncols):
                if (rr, cc) not in covered:
                    try:
                        empty_fill = empty_fill or self.border_fill_id({}, None)
                        table.cell(rr, cc).element.set("borderFillIDRef", empty_fill)
                    except Exception:
                        pass

        # 행 높이: 내용 추정값과 Word 의 행 높이 중 큰 값
        row_h = [0] * nrows
        for r, tr in enumerate(rows[:nrows]):
            trh = tr.find(f"{w('trPr')}/{w('trHeight')}")
            if trh is not None:
                try:
                    v = int(float(wattr(trh, "val", "0"))) * TWIP_TO_HWP
                    row_h[r] = v if wattr(trh, "hRule") == "exact" else max(row_h[r], v)
                except ValueError:
                    pass
        for (r0, c0), (r1, need) in cell_need.items():
            if r1 == r0:
                row_h[r0] = max(row_h[r0], need)
        for (r0, c0), (r1, need) in cell_need.items():
            if r1 > r0 and r1 < nrows:
                have = sum(row_h[r0:r1 + 1])
                if need > have:
                    row_h[r1] += need - have
        row_h = [max(h, 600) for h in row_h]
        for r in range(nrows):
            for c in range(ncols):
                try:
                    table.cell(r, c).set_size(height=row_h[r])
                except Exception:
                    pass

        # 병합 (아래/오른쪽부터)
        for (r0, c0, r1, c1, tc) in sorted(layout, key=lambda e: (-e[0], -e[1])):
            if (r1 > r0 or c1 > c0) and r1 < nrows and c1 < ncols:
                try:
                    table.merge_cells(r0, c0, r1, c1)
                except Exception as e:
                    self.warnings.append(f"셀 병합 실패({r0},{c0}): {e}")
        try:
            table.set_column_widths(grid)
        except Exception:
            pass
        total_h = sum(row_h)
        sz = table.element.find(f"{{{HP_NS}}}sz")
        if sz is not None:
            sz.set("height", str(total_h))
        table.mark_dirty()
        return total_h

    def fill_cell(self, cell, tc, part_rels, cell_w, tstyle):
        saved, self._est = self._est, []
        try:
            self._fill_cell(cell, tc, part_rels, cell_w, tstyle)
            return self.estimate_height(self._est, cell_w)
        finally:
            self._est = saved

    def _fill_cell(self, cell, tc, part_rels, cell_w, tstyle):
        existing = list(cell.paragraphs)
        first = [True]

        def factory(para_pr, char_pr):
            if first[0] and existing:
                first[0] = False
                p = existing[0]
                for run in list(p.runs):
                    try:
                        run.element.getparent().remove(run.element)
                    except Exception:
                        pass
                p.para_pr_id_ref = para_pr
                return p
            first[0] = False
            p = cell.add_paragraph("", para_pr_id_ref=para_pr, char_pr_id_ref=char_pr)
            for run in list(p.runs):
                run.element.getparent().remove(run.element)
            return p

        for child in tc:
            if not isinstance(child.tag, str):
                continue
            tag = etree.QName(child).localname
            if tag == "p":
                self.emit_paragraph(child, factory, part_rels, cell_w, table_style=tstyle, is_cell=True)
                if self._pending_boxes:
                    first[0] = False
                    h = self.flush_textboxes(cell, None, part_rels, cell_w)
                    self._est.append(("", 0, 0, 0, 0, h))
            elif tag == "tbl":
                first[0] = False
                outer = self._est
                inner_h = self.emit_table(child, cell, part_rels, cell_w) or 0
                self._est = outer
                self._est.append(("", 0, 0, 0, 0, inner_h))
            elif tag == "sdt":
                for p in child.iter(w("p")):
                    self.emit_paragraph(p, factory, part_rels, cell_w, table_style=tstyle, is_cell=True)

    # ---------------- 구역(쪽 설정) ------------------------------------
    def apply_section(self, sect_pr, section_index):
        if sect_pr is None:
            return
        pnt = sect_pr.find(w("pgNumType"))
        if pnt is not None:
            fmt = {"lowerRoman": "ROMAN_SMALL", "upperRoman": "ROMAN_CAPITAL",
                   "lowerLetter": "LATIN_SMALL", "upperLetter": "LATIN_CAPITAL",
                   "decimalEnclosedCircle": "CIRCLED_DIGIT"}.get(wattr(pnt, "fmt", ""), None)
            start = wattr(pnt, "start")
            self.page_num_fmt[section_index] = (fmt, int(start) if start and start.isdigit() else None)
        pg = sect_pr.find(w("pgSz"))
        mar = sect_pr.find(w("pgMar"))
        all_cols = sect_pr.findall(w("cols"))
        cols = next((c for c in reversed(all_cols) if wattr(c, "num")), all_cols[-1] if all_cols else None)
        try:
            n = int(wattr(cols, "num", "1") or 1) if cols is not None else 1
            if n > 1:
                gap = int(float(wattr(cols, "space", "720") or 720)) * TWIP_TO_HWP
                sep = wattr(cols, "sep") in ("1", "true", "on")
                self.doc.page.set_columns(min(n, 10), same_gap=gap,
                                          separator_type="SOLID" if sep else None,
                                          section_index=section_index)
        except Exception as e:
            self.warnings.append(f"다단 설정 생략: {e}")
        try:
            if pg is not None:
                wv = int(float(wattr(pg, "w", "11906"))) * TWIP_TO_HWP
                hv = int(float(wattr(pg, "h", "16838"))) * TWIP_TO_HWP
                land = wattr(pg, "orient") == "landscape"
                if land and wv < hv:
                    wv, hv = hv, wv
                self.doc.page.set_size(width=min(wv, hv) if land else wv,
                                       height=max(wv, hv) if land else hv,
                                       orientation="NARROWLY" if land else "WIDELY",
                                       section_index=section_index)
            if mar is not None:
                def m(name, default):
                    return abs(int(float(wattr(mar, name, default)))) * TWIP_TO_HWP
                top = m("top", "1440")
                bottom = m("bottom", "1440")
                header = m("header", "720")
                footer = m("footer", "720")
                # Word 위/아래 여백은 머리말 포함 거리, 한글은 머리말 영역을 따로 더함
                self.doc.page.set_margins(
                    left=m("left", "1440"), right=m("right", "1440"),
                    top=max(top - header, 0), bottom=max(bottom - footer, 0),
                    header=header, footer=footer, gutter=m("gutter", "0"),
                    section_index=section_index)
        except Exception as e:
            self.warnings.append(f"쪽 설정 일부 생략: {e}")
        # 머리말/꼬리말
        for kind in ("header", "footer"):
            ref = None
            for r in sect_pr.findall(w(f"{kind}Reference")):
                if wattr(r, "type", "default") == "default":
                    ref = r
            if ref is None:
                continue
            rid = ref.get(f"{{{R}}}id")
            target = self.pkg.rels.get(rid, (None,))[0]
            root = self.pkg._xml(target) if target else None
            if root is None:
                continue
            self.emit_header_footer(kind, root, target, section_index, sect_pr)

    def emit_header_footer(self, kind, root, part_name, section_index, sect_pr=None):
        """머리말/꼬리말: 본문과 같은 방식으로 변환 (글자 서식·그림·쪽 번호·전체 쪽 수)"""
        rels = self.pkg.part_rels(part_name)
        paras = []
        for el in root.iter(w("p")):
            # 글상자 안 문단은 제외 (글상자 처리에서 다룸)
            if any(a.tag == w("txbxContent") for a in el.iterancestors()):
                continue
            paras.append(el)
        has_content = False
        for p in paras:
            for k, v, f in self.collect_runs(p, wattr(p.find(f"{w('pPr')}/{w('pStyle')}"), "val"), rels):
                if (k == "text" and v.strip()) or k in ("image", "pagenum", "numpages", "equation"):
                    has_content = True
        if not has_content:
            return
        label = "머리말" if kind == "header" else "꼬리말"
        try:
            hf = (self.doc.page.set_header if kind == "header" else self.doc.page.set_footer)(
                text="", section_index=section_index)
        except Exception as e:
            self.warnings.append(f"{label} 생략: {e}")
            return
        state = {"first": True}

        def factory(para_pr, char_pr):
            if state["first"] and hf.paragraphs:
                state["first"] = False
                para = hf.paragraphs[0]
                for r in list(para.element):
                    if etree.QName(r).localname == "run":
                        para.element.remove(r)
            else:
                state["first"] = False
                hf.add_paragraph()
                para = hf.paragraphs[-1]
            para.para_pr_id_ref = para_pr
            return para

        max_w = self.text_width(sect_pr) if sect_pr is not None else 42520
        self._hf = (hf, self.page_num_fmt.get(section_index, (None, None))[0] or "DIGIT")
        try:
            for p in paras:
                self.emit_paragraph(p, factory, rels, max_w, is_cell=True)
                self._pending_boxes = []
        except Exception as e:
            self.warnings.append(f"{label} 일부 생략: {e}")
        finally:
            self._hf = None

    def add_page_field(self, para, kind, fmt):
        """머리말/꼬리말 안에 쪽 번호(PAGE) 또는 전체 쪽 수(TOTAL_PAGE) 자동 번호 넣기"""
        if not self._hf:
            return False
        hf, num_fmt = self._hf
        before = set(id(r) for r in para.element)
        try:
            if kind == "pagenum":
                hf.add_page_number_field(paragraph=para.element, format_type=num_fmt)
            else:
                hf._add_total_page_counter(para.element, num_fmt)
        except Exception as e:
            self.warnings.append(f"쪽 번호 넣기 실패: {e}")
            return False
        cp = self.char_pr(fmt)
        for r in list(para.element):
            if id(r) in before or etree.QName(r).localname != "run":
                continue
            if r.find(f".//{{{HP_NS}}}pageNum") is not None:
                para.element.remove(r)   # "- 1 -" 위치 컨트롤은 빼고 번호 글자만 남김
            else:
                r.set("charPrIDRef", str(cp))
        return True

    # ---------------- 본문 ---------------------------------------------
    def text_width(self, sect_pr):
        try:
            pg = sect_pr.find(w("pgSz"))
            mar = sect_pr.find(w("pgMar"))
            wv = int(float(wattr(pg, "w", "11906")))
            if wattr(pg, "orient") == "landscape":
                wv = max(wv, int(float(wattr(pg, "h", "16838"))))
            return (wv - int(float(wattr(mar, "left", "1440"))) - int(float(wattr(mar, "right", "1440")))) * TWIP_TO_HWP
        except Exception:
            return 42520

    def convert(self, progress: Callable[[float], None] | None = None):
        body = self.pkg.document.find(w("body"))
        rels = self.pkg.rels
        blocks = []          # [(element, section_no)]
        sect_prs = []        # 구역별 sectPr
        sec = 0

        def gather(container):
            nonlocal sec
            for el in container:
                if not isinstance(el.tag, str):
                    continue
                tag = etree.QName(el).localname
                if tag == "p":
                    blocks.append((el, sec))
                    sp = el.find(f"{w('pPr')}/{w('sectPr')}")
                    if sp is not None:
                        sect_prs.append(sp)
                        sec += 1
                elif tag == "tbl":
                    blocks.append((el, sec))
                elif tag in ("sdt",):
                    content = el.find(w("sdtContent"))
                    if content is not None:
                        gather(content)
                elif tag == "customXml":
                    gather(el)
                elif tag == "sectPr":
                    sect_prs.append(el)
        gather(body)
        if not sect_prs:
            sect_prs.append(etree.Element(w("sectPr")))
        # 마지막 구역 뒤에 빈 구역이 생기지 않도록
        nsec = len(sect_prs)
        for _ in range(1, nsec):
            self.doc.add_section()
        for i, sp in enumerate(sect_prs):
            self.apply_section(sp, i)

        total = max(len(blocks), 1)
        current_sec = 0
        for i, (el, s) in enumerate(blocks):
            s = min(s, nsec - 1)
            current_sec = s
            max_w = self.text_width(sect_prs[s])
            section = self.doc.sections[s]
            tag = etree.QName(el).localname

            def factory(para_pr, char_pr, _section=section, _s=s):
                if _s == 0:
                    return self.new_body_paragraph(para_pr, char_pr)
                paras = _section.paragraphs
                if len(paras) == 1 and not getattr(self, f"_sec_used_{_s}", False):
                    setattr(self, f"_sec_used_{_s}", True)
                    p = paras[0]
                    p.para_pr_id_ref = para_pr
                    return p
                return self.doc.add_paragraph("", section=_section, para_pr_id_ref=para_pr,
                                              char_pr_id_ref=char_pr, include_run=False, inherit_style=False)
            try:
                if tag == "p":
                    self.emit_paragraph(el, factory, rels, max_w)
                    if self._pending_boxes:
                        if s == 0 and not self._first_para_used:
                            self._first_para_used = True
                        self.flush_textboxes(None, section, rels, max_w)
                else:
                    if s == 0 and not self._first_para_used:
                        # 첫 문단(구역 정보 보유)을 빈 문단으로 두고 표를 그 뒤에
                        self.new_body_paragraph(self.para_pr(ParaFmt(), 10, None), self.char_pr(RunFmt()))
                    if s > 0 and not getattr(self, f"_sec_used_{s}", False):
                        setattr(self, f"_sec_used_{s}", True)
                    self.emit_table_in_section(el, section, rels, max_w)
            except Exception as e:
                self.warnings.append(f"{i + 1}번째 블록 변환 중 오류(건너뜀): {e}")
            if progress and i % 10 == 0:
                progress(i / total)
        # 구역별 쪽 번호 다시 시작
        for si, (_fmt, start) in self.page_num_fmt.items():
            if start is None or si >= len(self.doc.sections):
                continue
            paras = self.doc.sections[si].paragraphs
            if paras and (si > 0 or start != 1):
                try:
                    self.doc.page.restart_page_number(paras[0], number=start)
                except Exception as e:
                    self.warnings.append(f"쪽 번호 다시 시작 생략: {e}")
        if progress:
            progress(1.0)

    def emit_table_in_section(self, tbl_el, section, rels, max_w):
        orig = self.doc.add_table

        def add_table(rows, cols, **kw):
            return orig(rows, cols, section=section, **kw)
        self.doc.add_table = add_table
        try:
            self.emit_table(tbl_el, None, rels, max_w)
        finally:
            self.doc.add_table = orig

    def finalize(self):
        # Word 처럼 글꼴 고유의 빈칸 너비 사용 ("글꼴에 어울리는 빈칸")
        HH = "{http://www.hancom.co.kr/hwpml/2011/head}"
        for hdr in self.doc.oxml.headers:
            for cp in hdr.element.iter(f"{HH}charPr"):
                cp.set("useFontSpace", "1")
            try:
                hdr.mark_dirty()
            except Exception:
                hdr._dirty = True

    def save(self, out_path: str):
        self.finalize()
        if out_path.lower().endswith(".hwp"):
            try:
                self.doc.save_to_path(out_path)
                return out_path
            except Exception as e:
                alt = os.path.splitext(out_path)[0] + ".hwpx"
                self.warnings.append(f".hwp 저장이 안 되어 .hwpx 로 저장했습니다 ({e})")
                if os.path.exists(out_path) and os.path.getsize(out_path) == 0:
                    os.remove(out_path)
                self.doc.save_to_path(alt)
                return alt
        self.doc.save_to_path(out_path)
        return out_path


# ---------------------------------------------------------------------------
# 한컴오피스가 설치된 경우: 한글 프로그램으로 직접 변환
# ---------------------------------------------------------------------------
def hancom_available() -> bool:
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "HWPFrame.HwpObject"))
        return True
    except Exception:
        return False


def convert_with_hancom(src: str, dst: str) -> str:
    import pythoncom
    import win32com.client as win32
    pythoncom.CoInitialize()
    hwp = None
    try:
        hwp = win32.Dispatch("HWPFrame.HwpObject")
        try:
            hwp.XHwpWindows.Item(0).Visible = False
        except Exception:
            pass
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        if not hwp.Open(os.path.abspath(src), "", "forceopen:true;suspendpassword:true"):
            raise RuntimeError("한글 프로그램이 파일을 열지 못했습니다.")
        fmt = "HWPX" if dst.lower().endswith(".hwpx") else "HWP"
        hwp.SaveAs(os.path.abspath(dst), fmt, "")
        return dst
    finally:
        if hwp is not None:
            try:
                hwp.Clear(1)
                hwp.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# 공개 함수
# ---------------------------------------------------------------------------
def unique_path(path: str) -> str:
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 2
    while os.path.exists(f"{base} ({i}){ext}"):
        i += 1
    return f"{base} ({i}){ext}"


def convert_file(src: str, out_dir: str | None = None, fmt: str = "hwp",
                 use_hancom: bool = True, overwrite: bool = False,
                 progress: Callable[[float], None] | None = None) -> tuple[str, list[str]]:
    """src(.docx) → 한글 파일. (저장 경로, 경고 목록) 반환"""
    src = os.path.abspath(src)
    ext = os.path.splitext(src)[1].lower()
    if ext not in (".docx", ".docm", ".dotx", ".dotm", ".doc"):
        raise ValueError("Word 파일(.docx)만 변환할 수 있습니다.")
    if not os.path.isfile(src):
        raise ValueError("파일을 찾을 수 없습니다.")
    if os.path.getsize(src) == 0:
        raise ValueError("빈 파일(0바이트)입니다. 파일이 제대로 저장/다운로드되었는지 확인해 주세요.")
    out_dir = out_dir or os.path.dirname(src)
    dst = os.path.join(out_dir, os.path.splitext(os.path.basename(src))[0] + "." + fmt)
    if not overwrite:
        dst = unique_path(dst)
    if use_hancom and hancom_available():
        try:
            return convert_with_hancom(src, dst), ["한글 프로그램으로 변환했습니다."]
        except Exception as e:
            note = f"한글 프로그램 변환 실패, 내장 변환기로 진행: {e}"
        else:
            note = None
    else:
        note = None
    if not os.path.isdir(out_dir):
        raise ValueError(f"저장할 폴더가 없습니다: {out_dir}")
    tmpdir = None
    work_src = src
    with open(src, "rb") as fh:
        is_zip = fh.read(2) == b"PK"
    if ext == ".doc" and not is_zip:
        # 예전 .doc → LibreOffice 가 있으면 .docx 로 바꿔서 진행
        work_src, tmpdir = doc_to_docx(src)
        note = (note + " / " if note else "") + "예전 .doc 파일이라 LibreOffice로 .docx 변환 후 진행했습니다."
    try:
        conv = Converter(work_src)
        conv.convert(progress)
        try:
            out = conv.save(dst)
        except PermissionError:
            raise ValueError("저장할 수 없습니다. 같은 이름의 파일이 한글에서 열려 있거나 "
                             "폴더에 쓰기 권한이 없습니다.")
        except OSError as e:
            raise ValueError(f"저장 중 오류가 났습니다: {e}")
    finally:
        if tmpdir:
            import shutil
            shutil.rmtree(tmpdir, ignore_errors=True)
    warnings = ([note] if note else []) + list(dict.fromkeys(conv.warnings))
    return out, warnings


def find_soffice() -> str | None:
    import shutil
    for cand in (shutil.which("soffice"),
                 r"C:\Program Files\LibreOffice\program\soffice.exe",
                 r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if cand and os.path.exists(cand):
            return cand
    return None


def doc_to_docx(src: str):
    import subprocess
    import tempfile
    soffice = find_soffice()
    if not soffice:
        raise ValueError("예전 형식(.doc) 파일입니다. Word 에서 '다른 이름으로 저장 → Word 문서(.docx)' 후 "
                         "다시 넣어 주세요. (LibreOffice 가 설치되어 있으면 자동으로 바꿔 줍니다)")
    tmpdir = tempfile.mkdtemp(prefix="w2h_")
    tmp_src = os.path.join(tmpdir, "input.doc")
    import shutil
    shutil.copyfile(src, tmp_src)
    try:
        subprocess.run([soffice, "--headless", "--norestore", "--convert-to", "docx",
                        "--outdir", tmpdir, tmp_src], capture_output=True, timeout=180,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except Exception as e:
        shutil.rmtree(tmpdir, ignore_errors=True)
        raise ValueError(f".doc 변환에 실패했습니다: {e}")
    out = os.path.join(tmpdir, "input.docx")
    if not os.path.exists(out):
        shutil.rmtree(tmpdir, ignore_errors=True)
        raise ValueError(".doc 파일을 변환하지 못했습니다. Word 에서 .docx 로 저장한 뒤 넣어 주세요.")
    return out, tmpdir


if __name__ == "__main__":
    import sys
    for f in sys.argv[1:]:
        try:
            out, warns = convert_file(f)
            print("완료:", out)
            for wmsg in warns[:20]:
                print("  -", wmsg)
        except Exception as e:
            print("실패:", f, e)
