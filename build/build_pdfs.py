"""Gera os PDFs do treinamento: apostila, prova, gabarito e roteiro do instrutor.

A apostila tem caixas de vídeo interativas: um botão "Mostrar / ocultar vídeo"
alterna a visibilidade do quadro do vídeo (JavaScript de formulário PDF, que
funciona no Adobe Acrobat Reader e no Foxit). O quadro abre o link definido em
build/videos.json.
"""
import json
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor, white
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.fonts import addMapping
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table,
                                TableStyle, PageBreak, KeepTogether, Flowable, NextPageTemplate,
                                CondPageBreak)
from reportlab.platypus.tableofcontents import TableOfContents

import brand as B
from content import MODULES, SCHEDULE, MATERIALS, CHALLENGE, EXAM, VIDEOS

# ---------------------------------------------------------------- fontes e estilos
# Ubuntu (Ubuntu Font Licence) fica em build/assets/fonts, então a build não depende das fontes do sistema.
# Cobre todo o teclado ABNT2 (acentos, ç, ª, º, §, ¹²³, £, ¢, ¬, °). O "negrito" usa a Ubuntu Medium.
FONTS = B.ASSETS / "fonts"
for name, file in [("UB", "Regular"), ("UB-B", "Medium"), ("UB-I", "Italic"), ("UB-BI", "MediumItalic"), ("UB-L", "Light")]:
    pdfmetrics.registerFont(TTFont(name, str(FONTS / f"Ubuntu-{file}.ttf")))
for bold, ital, name in [(0, 0, "UB"), (1, 0, "UB-B"), (0, 1, "UB-I"), (1, 1, "UB-BI")]:
    addMapping("UB", bold, ital, name)
# a Ubuntu não tem ✓, ◯ nem ●: esses marcadores usam a Segoe UI Symbol
pdfmetrics.registerFont(TTFont("SYM", "C:/Windows/Fonts/seguisym.ttf"))


def SYM(ch):
    return f'<font name="SYM">{ch}</font>'

H = HexColor
DARK = H("#" + B.BG)
TXT = H("#2B2B2B")
SOFT = H("#5E645F")
RULE = H("#D9DDD5")
TINT = H("#F2F6EE")
TINT2 = H("#EAF4F6")
LIME, GREEN, TEAL, BLUE, DEEP = (H("#" + c) for c in (B.LIME, B.GREEN, B.TEAL, B.BLUE, B.DEEP))
# versões escuras das cores da marca para TEXTO sobre fundo branco (contraste AA)
TXTC = {1: H("#4A7F1C"), 2: H("#4A7F1C"), 3: H("#2E7D3A"), 4: H("#1D7563"), 5: H("#1B6A87"), 6: H("#175A7C")}
SHAPEC = {1: LIME, 2: LIME, 3: GREEN, 4: TEAL, 5: BLUE, 6: DEEP}
STAGES = ["Empatizar", "Definir", "Idear", "Prototipar", "Testar"]
STAGE_TXT = ["Entender as pessoas e o contexto", "Encontrar o problema certo", "Gerar muitas soluções", "Tornar a ideia tangível", "Validar com pessoas reais"]
STAGEC = [LIME, GREEN, TEAL, BLUE, DEEP]

PW, PH = A4
M = 52  # margem
CW = PW - 2 * M

ST = {
    "h1": ParagraphStyle("H1", fontName="UB-B", fontSize=25, leading=30, textColor=TXT, spaceAfter=4),
    "sub": ParagraphStyle("sub", fontName="UB", fontSize=13, leading=17, textColor=SOFT, spaceAfter=14),
    "h2": ParagraphStyle("H2", fontName="UB-B", fontSize=14.5, leading=19, textColor=TXT, spaceBefore=14, spaceAfter=6),
    "p": ParagraphStyle("p", fontName="UB", fontSize=10.5, leading=15.5, textColor=TXT, spaceAfter=7),
    "small": ParagraphStyle("small", fontName="UB", fontSize=9, leading=12.5, textColor=SOFT),
    "cell": ParagraphStyle("cell", fontName="UB", fontSize=9.5, leading=13, textColor=TXT),
    "cellb": ParagraphStyle("cellb", fontName="UB-B", fontSize=9.5, leading=13, textColor=TXT),
    "cellh": ParagraphStyle("cellh", fontName="UB-B", fontSize=9.5, leading=13, textColor=white),
    "label": ParagraphStyle("label", fontName="UB-B", fontSize=8.5, leading=11, textColor=SOFT),
    "q": ParagraphStyle("q", fontName="UB-B", fontSize=11, leading=15.5, textColor=TXT, spaceAfter=6),
}


def P(t, s="p", **kw):
    st = ST[s] if not kw else ParagraphStyle(s + str(id(kw)), parent=ST[s], **kw)
    return Paragraph(t, st)


def hexs(c):
    return "#" + c.hexval()[2:]


# ---------------------------------------------------------------- flowables
class Badge(Flowable):
    """Número grande do módulo + rótulo do encontro."""

    def __init__(self, n, label, color):
        super().__init__()
        self.n, self.label, self.color = n, label, color

    def wrap(self, aw, ah):
        return aw, 74

    def draw(self):
        c = self.canv
        c.setFillColor(self.color)
        c.setFont("UB-B", 54)
        c.drawString(-2, 16, f"{self.n:02d}")
        c.setFont("UB-B", 8.5)
        c.setFillColor(SOFT)
        c.drawString(0, 2, self.label.upper())


class Stages(Flowable):
    def wrap(self, aw, ah):
        self.w = aw
        return aw, 104

    def draw(self):
        c = self.canv
        cw = self.w / 5
        for i, (st, tx) in enumerate(zip(STAGES, STAGE_TXT)):
            cx = cw * i + cw / 2
            if i < 4:
                c.setStrokeColor(RULE)
                c.setLineWidth(2)
                c.line(cx + 22, 78, cx + cw - 22, 78)
            c.setFillColor(STAGEC[i])
            c.circle(cx, 78, 18, stroke=0, fill=1)
            c.setFillColor(DARK)
            c.setFont("UB-B", 13)
            c.drawCentredString(cx, 73.5, str(i + 1))
            c.setFillColor(TXT)
            c.setFont("UB-B", 10.5)
            c.drawCentredString(cx, 46, st)
            p = Paragraph(tx, ParagraphStyle("st", fontName="UB", fontSize=8.5, leading=11, textColor=SOFT, alignment=TA_CENTER))
            w, h = p.wrap(cw - 10, 40)
            p.drawOn(c, cw * i + 5, 38 - h)


class EmpathyMap(Flowable):
    def wrap(self, aw, ah):
        self.w = aw
        return aw, 236

    def draw(self):
        c = self.canv
        w = self.w
        qw, qh, g = (w - 8) / 2, 70, 8
        quads = [("Pensa e sente", "Preocupações, sonhos, medos"), ("Vê", "Ambiente, amigos, ofertas"),
                 ("Fala e faz", "Atitudes, comportamento"), ("Ouve", "Amigos, chefe, influenciadores")]
        cols = [LIME, GREEN, TEAL, BLUE]
        for i, (h, b) in enumerate(quads):
            x = (i % 2) * (qw + g)
            y = 236 - (i // 2 + 1) * (qh + g)
            c.setFillColor(TINT)
            c.roundRect(x, y, qw, qh, 8, stroke=0, fill=1)
            c.setFillColor(cols[i])
            c.circle(x + 16 if i % 2 == 0 else x + qw - 16, y + qh - 18, 5, stroke=0, fill=1)
            c.setFillColor(TXT)
            c.setFont("UB-B", 11)
            if i % 2 == 0:
                c.drawString(x + 28, y + qh - 22, h)
                c.setFont("UB", 9)
                c.setFillColor(SOFT)
                c.drawString(x + 28, y + qh - 38, b)
            else:
                c.drawRightString(x + qw - 28, y + qh - 22, h)
                c.setFont("UB", 9)
                c.setFillColor(SOFT)
                c.drawRightString(x + qw - 28, y + qh - 38, b)
        cy = 236 - qh - g / 2
        c.setFillColor(white)
        c.circle(w / 2, cy, 34, stroke=0, fill=1)
        c.setFillColor(GREEN)
        c.circle(w / 2, cy, 29, stroke=0, fill=1)
        c.setFillColor(DARK)
        c.setFont("UB-B", 9.5)
        c.drawCentredString(w / 2, cy - 3.5, "Persona")
        for i, (h, b) in enumerate([("Dores", "Frustrações, obstáculos, riscos"), ("Ganhos", "Desejos, necessidades, medidas de sucesso")]):
            x = i * (qw + g)
            c.setFillColor(TINT2)
            c.roundRect(x, 0, qw, 64, 8, stroke=0, fill=1)
            c.setFillColor(TXT)
            c.setFont("UB-B", 11)
            c.drawString(x + 14, 40, h)
            c.setFont("UB", 9)
            c.setFillColor(SOFT)
            c.drawString(x + 14, 22, b)


class Matrix(Flowable):
    def wrap(self, aw, ah):
        self.w = aw
        return aw, 200

    def draw(self):
        c = self.canv
        x0 = 28
        qw, qh, g = (self.w - x0 - 8) / 2, 84, 8
        quads = [("Vitórias rápidas", "Alto impacto, baixo esforço: faça já", LIME),
                 ("Grandes apostas", "Alto impacto, alto esforço: planeje", BLUE),
                 ("Tarefas extras", "Baixo impacto, baixo esforço: se sobrar tempo", H("#9AA09A")),
                 ("Evite", "Baixo impacto, alto esforço: descarte", H("#C0584F"))]
        for i, (h, b, col) in enumerate(quads):
            x = x0 + (i % 2) * (qw + g)
            y = 200 - 14 - (i // 2 + 1) * qh - (i // 2) * g
            c.setFillColor(TINT)
            c.roundRect(x, y, qw, qh, 8, stroke=0, fill=1)
            c.setFillColor(col)
            c.circle(x + 18, y + qh - 20, 7, stroke=0, fill=1)
            c.setFillColor(TXT)
            c.setFont("UB-B", 11.5)
            c.drawString(x + 32, y + qh - 24, h)
            p = Paragraph(b, ParagraphStyle("mx", fontName="UB", fontSize=9, leading=12, textColor=SOFT))
            w, hh = p.wrap(qw - 44, 50)
            p.drawOn(c, x + 32, y + qh - 34 - hh)
        c.setStrokeColor(SOFT)
        c.setLineWidth(1.2)
        c.line(14, 12, 14, 186)
        c.line(14, 186, 10, 178)
        c.line(14, 186, 18, 178)
        c.line(x0, 4, self.w, 4)
        c.line(self.w, 4, self.w - 8, 0)
        c.line(self.w, 4, self.w - 8, 8)
        c.saveState()
        c.translate(6, 100)
        c.rotate(90)
        c.setFont("UB-B", 7.5)
        c.setFillColor(SOFT)
        c.drawCentredString(0, 0, "IMPACTO")
        c.restoreState()
        c.setFont("UB-B", 7.5)
        c.setFillColor(SOFT)
        c.drawCentredString(x0 + (self.w - x0) / 2, -8, "ESFORÇO")


VIDEO_SLOTS = {}  # key -> dict(page, video_rect, toggle_rect, color)


class VideoBox(Flowable):
    """Espaço de vídeo. O quadro interativo e o botão são adicionados depois (pypdf)."""
    BW, BH = 400, 225

    def __init__(self, key, n, color):
        super().__init__()
        self.key, self.n, self.color = key, n, color

    def wrap(self, aw, ah):
        self.w = aw
        return aw, self.BH + 40

    def draw(self):
        c = self.canv
        c.setFillColor(self.color)
        c.circle(7, self.BH + 25, 7, stroke=0, fill=1)
        c.setFillColor(white)
        p = c.beginPath()
        p.moveTo(4.8, self.BH + 21.5)
        p.lineTo(4.8, self.BH + 28.5)
        p.lineTo(10.6, self.BH + 25)
        p.close()
        c.drawPath(p, stroke=0, fill=1)
        c.setFillColor(TXT)
        c.setFont("UB-B", 11)
        c.drawString(20, self.BH + 21, "Vídeo do módulo")
        c.setFont("UB", 9)
        c.setFillColor(SOFT)
        c.drawString(20, self.BH + 9, VIDEOS[self.key])
        # espaço reservado (visível quando o vídeo está oculto)
        bx = (self.w - self.BW) / 2
        c.setStrokeColor(RULE)
        c.setDash(4, 3)
        c.setFillColor(H("#FAFBF9"))
        c.roundRect(bx, 0, self.BW, self.BH, 8, stroke=1, fill=1)
        c.setDash()
        c.setFillColor(SOFT)
        c.setFont("UB", 9.5)
        c.drawCentredString(self.w / 2, self.BH / 2 + 4, "Vídeo oculto.")
        c.drawCentredString(self.w / 2, self.BH / 2 - 10, "Use o botão \u201cMostrar / ocultar vídeo\u201d para exibi-lo.")
        ax, ay = c.absolutePosition(0, 0)
        VIDEO_SLOTS[self.key] = {
            "page": c.getPageNumber() - 1,
            "video": (ax + bx, ay, ax + bx + self.BW, ay + self.BH),
            "toggle": (ax + self.w - 150, ay + self.BH + 12, ax + self.w, ay + self.BH + 34),
            "color": self.color.hexval()[2:],
        }


class Lines(Flowable):
    def __init__(self, n, gap=20):
        super().__init__()
        self.n, self.gap = n, gap

    def wrap(self, aw, ah):
        self.w = aw
        return aw, self.n * self.gap + 4

    def draw(self):
        c = self.canv
        c.setStrokeColor(RULE)
        c.setLineWidth(0.6)
        for i in range(self.n):
            y = i * self.gap + 2
            c.line(0, y, self.w, y)


def box(title, text, n=1):
    t = Table([[P(f"<b>{title}</b>", "p", textColor=TXTC[n], spaceAfter=2)], [P(text, "p", spaceAfter=0)]], colWidths=[CW])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), TINT), ("LEFTPADDING", (0, 0), (-1, -1), 14),
                           ("RIGHTPADDING", (0, 0), (-1, -1), 14), ("TOPPADDING", (0, 0), (0, 0), 10),
                           ("BOTTOMPADDING", (0, -1), (-1, -1), 11), ("TOPPADDING", (0, 1), (-1, -1), 0),
                           ("ROUNDEDCORNERS", [8, 8, 8, 8])]))
    return KeepTogether([Spacer(1, 4), t, Spacer(1, 10)])


def bullets(items, color):
    rows = [[P(f'<font name="SYM" color="{hexs(color)}">\u25CF</font>', "p", spaceAfter=0), P(it, "p", spaceAfter=0)] for it in items]
    t = Table(rows, colWidths=[16, CW - 16])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                           ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                           ("TOPPADDING", (0, 0), (-1, -1), 0)]))
    return [t, Spacer(1, 6)]


def table(header, rows, n=1, widths=None):
    data = [[P(h, "cellh") for h in header]] + [[P(c, "cellb" if j == 0 else "cell") for j, c in enumerate(r)] for r in rows]
    if widths is None:
        widths = [CW / len(header)] * len(header)
    t = Table(data, colWidths=widths, repeatRows=1)
    st = [("BACKGROUND", (0, 0), (-1, 0), TXTC[n]), ("VALIGN", (0, 0), (-1, -1), "TOP"),
          ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8),
          ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
          ("LINEBELOW", (0, 1), (-1, -1), 0.5, RULE)]
    for i in range(1, len(data)):
        if i % 2 == 0:
            st.append(("BACKGROUND", (0, i), (-1, i), H("#F7F8F6")))
    t.setStyle(TableStyle(st))
    return [Spacer(1, 4), t, Spacer(1, 10)]


def exercise(title, items, n):
    parts = [P(f"<b>{title}</b>", "p", textColor=TXTC[n])]
    for i, it in enumerate(items):
        parts += [P(f"{i + 1}. {it}", "p", spaceAfter=4), Lines(3), Spacer(1, 6)]
    t = Table([[parts]], colWidths=[CW])
    t.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, RULE), ("LEFTPADDING", (0, 0), (-1, -1), 14),
                           ("RIGHTPADDING", (0, 0), (-1, -1), 14), ("TOPPADDING", (0, 0), (-1, -1), 12),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 8), ("ROUNDEDCORNERS", [8, 8, 8, 8])]))
    return [Spacer(1, 8), t]


# ---------------------------------------------------------------- páginas
def draw_cover(c, doc, kind, title2):
    c.saveState()
    c.setFillColor(DARK)
    c.rect(0, 0, PW, PH, stroke=0, fill=1)
    c.drawImage(str(B.ASSETS / "logo_full_t.png"), PW - 300, PH - 330, width=250, preserveAspectRatio=True, mask="auto", anchor="n", height=160)
    c.setFillColor(LIME)
    c.setFont("UB-B", 10)
    c.drawString(M, PH - 400, kind.upper())
    c.setFillColor(H("#" + B.TEXT))
    c.setFont("UB-B", 38)
    c.drawString(M, PH - 450, "Design Thinking")
    c.setFillColor(LIME)
    c.drawString(M, PH - 494, "na Prática")
    c.setFillColor(H("#" + B.MUTED))
    c.setFont("UB", 14)
    c.drawString(M, PH - 528, title2)
    c.setFont("UB", 11)
    c.drawString(M, 120, f"Instrutor: {B.INSTRUCTOR}")
    c.drawString(M, 102, B.HOURS)
    c.setFillColor(H("#" + B.DIM))
    c.setFont("UB", 9)
    c.drawString(M, 50, f"{B.BRAND}  ·  {B.TAGLINE}")
    # circuito decorativo
    c.setStrokeColor(BLUE)
    c.setLineWidth(2)
    for pts in [[(PW - 230, 150), (PW - 110, 150)], [(PW - 230, 118), (PW - 150, 118), (PW - 126, 94), (PW - 70, 94)], [(PW - 230, 86), (PW - 190, 86), (PW - 168, 62), (PW - 100, 62)]]:
        for (x1, y1), (x2, y2) in zip(pts, pts[1:]):
            c.line(x1, y1, x2, y2)
        c.setFillColor(DARK)
        c.circle(pts[-1][0] + 7, pts[-1][1], 7, stroke=1, fill=1)
    c.restoreState()


def draw_body(c, doc):
    c.saveState()
    c.drawImage(str(B.ASSETS / "logo_icon_t.png"), PW - M - 38, PH - 40, width=38, height=25, mask="auto", preserveAspectRatio=True)
    c.setFont("UB", 8)
    c.setFillColor(SOFT)
    c.drawString(M, PH - 32, f"{B.COURSE.upper()}  ·  {doc.doc_label}")
    c.setStrokeColor(RULE)
    c.setLineWidth(0.5)
    c.line(M, 40, PW - M, 40)
    c.drawString(M, 28, f"{B.BRAND}  ·  {B.TAGLINE}")
    c.drawRightString(PW - M, 28, str(c.getPageNumber()))
    c.restoreState()


def draw_back(c, doc):
    c.saveState()
    c.setFillColor(DARK)
    c.rect(0, 0, PW, PH, stroke=0, fill=1)
    c.drawImage(str(B.ASSETS / "logo_full_t.png"), (PW - 280) / 2, PH / 2 - 40, width=280, height=175, mask="auto", preserveAspectRatio=True)
    c.setFillColor(H("#" + B.MUTED))
    c.setFont("UB", 11)
    c.drawCentredString(PW / 2, PH / 2 - 80, f"{B.INSTRUCTOR}  ·  hayaidatasystems@gmail.com")
    c.setFillColor(H("#" + B.DIM))
    c.setFont("UB", 8.5)
    c.drawCentredString(PW / 2, 60, "Material de uso exclusivo dos participantes do treinamento. Reprodução não autorizada.")
    c.restoreState()


class Doc(BaseDocTemplate):
    def __init__(self, path, label, cover_kind, cover_title, **kw):
        super().__init__(str(path), pagesize=A4, leftMargin=M, rightMargin=M, topMargin=58, bottomMargin=56,
                         title=f"{B.COURSE} — {label}", author=B.INSTRUCTOR, subject=B.SUBTITLE, creator=B.BRAND, **kw)
        self.doc_label = label.upper()
        frame = Frame(M, 56, CW, PH - 58 - 56, id="f", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
        self.addPageTemplates([
            PageTemplate("cover", [frame], onPage=lambda c, d: draw_cover(c, d, cover_kind, cover_title)),
            PageTemplate("body", [frame], onPage=draw_body),
            PageTemplate("back", [frame], onPage=draw_back),
        ])

    def afterFlowable(self, f):
        if isinstance(f, Paragraph) and f.style.name == "H1":
            self.notify("TOCEntry", (0, f.getPlainText(), self.page))


def start(story):
    story += [Spacer(1, 1), NextPageTemplate("body"), PageBreak()]


# ---------------------------------------------------------------- apostila
def build_handout(out):
    doc = Doc(out, "Apostila do participante", "Apostila do participante", B.SUBTITLE)
    s = []
    start(s)
    s.append(P("Como usar esta apostila", "h1"))
    s.append(P("Material de apoio aos três encontros do treinamento", "sub"))
    s.append(P("Esta apostila acompanha os slides e as dinâmicas. Use-a para revisar os conceitos, fazer os exercícios e registrar as ideias da sua equipe para o Desafio Final."))
    s += bullets([
        "<b>Vídeos:</b> cada módulo tem um espaço de vídeo. No <b>Adobe Acrobat Reader</b> (gratuito), use o botão <b>Mostrar / ocultar vídeo</b> para exibir ou esconder o quadro, e clique no quadro para abrir o vídeo.",
        "<b>Exercícios:</b> ao fim de cada módulo há perguntas práticas com espaço para respostas. Aplique-as a um problema real da sua empresa.",
        "<b>Desafio Final:</b> no Encontro 3, sua equipe aplica todo o processo. O canvas está no final desta apostila.",
    ], LIME)
    toc = TableOfContents()
    toc.levelStyles = [ParagraphStyle("toc", fontName="UB", fontSize=11, leading=20, textColor=TXT)]
    toc.dotsMinLevel = 0
    s += [P("Sumário", "h2"), toc]
    s.append(P("Agenda", "h2"))
    rows = []
    for i, enc in enumerate(SCHEDULE):
        mods = ", ".join(f"Módulo {m['n']}: {m['title']}" for m in MODULES if m["encontro"] == i + 1)
        extra = {0: "Dinâmicas 1 e 2", 1: "Dinâmicas 3 e 4", 2: "Desafio Final e Prova"}[i]
        rows.append([f"Encontro {i + 1} (4h)", f"{mods}. {extra}."])
    s += table(["Encontro", "Conteúdo"], rows, 1, [110, CW - 110])

    for m in MODULES:
        n, col = m["n"], SHAPEC[m["n"]]
        s += [PageBreak(), Badge(n, f"Encontro {m['encontro']}  ·  Módulo {n}", col), P(m["title"], "h1"), P(m["short"], "sub")]
        obj = Table([[P("<b>Objetivos do módulo</b>", "small", textColor=TXTC[n])]] + [[P(SYM("\u2713") + "  " + o, "p", spaceAfter=0)] for o in m["objectives"]], colWidths=[CW])
        obj.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), TINT), ("LEFTPADDING", (0, 0), (-1, -1), 14),
                                 ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                                 ("TOPPADDING", (0, 0), (0, 0), 10), ("BOTTOMPADDING", (0, -1), (-1, -1), 10),
                                 ("ROUNDEDCORNERS", [8, 8, 8, 8])]))
        s += [obj, Spacer(1, 6)]
        for blk in m["handout"]:
            kind = blk[0]
            if kind == "h":
                s += [CondPageBreak(90), P(blk[1], "h2")]
            elif kind == "p":
                s.append(P(blk[1]))
            elif kind == "bullets":
                s += bullets(blk[1], col)
            elif kind == "box":
                s.append(box(blk[1], blk[2], n))
            elif kind == "table":
                w = None
                if len(blk[1]) == 3:
                    w = [110, (CW - 110) / 2, (CW - 110) / 2]
                s += table(blk[1], blk[2], n, w)
            elif kind == "stages":
                s += [Spacer(1, 4), Stages(), Spacer(1, 8)]
            elif kind == "empathy":
                s += [Spacer(1, 6), EmpathyMap(), Spacer(1, 12)]
            elif kind == "matrix":
                s += [Spacer(1, 6), Matrix(), Spacer(1, 18)]
            elif kind == "video":
                s += [CondPageBreak(VideoBox.BH + 60), Spacer(1, 8), VideoBox(blk[1], n, col), Spacer(1, 12)]
            elif kind == "exercise":
                s += [CondPageBreak(220)] + exercise(blk[1], blk[2], n)

    # Desafio final
    s += [PageBreak(), Badge(7, "Encontro 3  ·  Aplicação", LIME), P("Desafio Final", "h1"), P(CHALLENGE["time"], "sub")]
    s.append(P(CHALLENGE["brief"]))
    s.append(P("Entregas por etapa", "h2"))
    s += table(["Etapa", "O que entregar"], CHALLENGE["deliverables"], 1, [110, CW - 110])
    s.append(P("Cronograma sugerido", "h2"))
    s += table(["Tempo", "Atividade"], CHALLENGE["timeline"], 4, [110, CW - 110])
    s.append(P("Critérios de avaliação do pitch", "h2"))
    s += table(["Critério", "Pergunta-chave", "Pontos"], [(a, b, str(c)) for a, b, c in CHALLENGE["rubric"]], 5, [130, CW - 190, 60])

    # Canvas
    s += [PageBreak(), P("Canvas do Desafio Final", "h1"), P("Equipe: ______________________________   Problema: ______________________________", "sub")]
    for i, (st, d) in enumerate(CHALLENGE["deliverables"]):
        cell = [P(f"<b>{i + 1}. {st}</b>", "p", textColor=TXTC[[1, 3, 4, 5, 6][i]], spaceAfter=1), P(d, "small"), Spacer(1, 78)]
        t = Table([[cell]], colWidths=[CW])
        t.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, RULE), ("LEFTPADDING", (0, 0), (-1, -1), 12),
                               ("TOPPADDING", (0, 0), (-1, -1), 8), ("ROUNDEDCORNERS", [8, 8, 8, 8])]))
        s += [t, Spacer(1, 6)]

    s += [PageBreak(), P("Anotações", "h1"), Spacer(1, 10), Lines(30, 21)]
    s += [NextPageTemplate("back"), PageBreak(), Spacer(1, 1)]
    doc.multiBuild(s)


# ---------------------------------------------------------------- prova e gabarito
def exam_header(s, title, sub):
    s.append(P(title, "h1"))
    s.append(P(sub, "sub"))


def build_exam(out):
    doc = Doc(out, "Prova", "Avaliação final", "Prova  ·  6 questões")
    s = []
    start(s)
    exam_header(s, "Prova final", "Design Thinking na Prática  ·  6 questões  ·  25 minutos")
    ident = Table([[P("Nome:", "label"), "", P("Data:", "label"), ""], [P("Empresa:", "label"), "", P("Nota:", "label"), ""]],
                  colWidths=[60, CW - 60 - 50 - 110, 50, 110], rowHeights=[26, 26])
    ident.setStyle(TableStyle([("LINEBELOW", (1, 0), (1, -1), 0.6, RULE), ("LINEBELOW", (3, 0), (3, -1), 0.6, RULE), ("VALIGN", (0, 0), (-1, -1), "BOTTOM")]))
    s += [ident, Spacer(1, 12)]
    s.append(box("Instruções", "Prova individual e sem consulta. Leia cada situação com atenção e marque <b>apenas uma</b> alternativa. Cada questão vale 1 ponto.", 5))
    for i, q in enumerate(EXAM):
        rows = [[P(SYM("\u25EF") + f"  <b>{'abcd'[j]})</b>", "p", spaceAfter=0), P(o, "p", spaceAfter=0)] for j, o in enumerate(q["opts"])]
        t = Table(rows, colWidths=[42, CW - 42])
        t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 4), ("TOPPADDING", (0, 0), (-1, -1), 0)]))
        s.append(KeepTogether([P(f"Questão {i + 1}", "label", textColor=TXTC[i + 1]), Spacer(1, 2), P(q["q"], "q"), t, Spacer(1, 14)]))
    s += [CondPageBreak(140), P("Folha de respostas", "h2"), P("Transcreva suas respostas marcando um X.", "small"), Spacer(1, 6)]
    grid = [[""] + [f"Q{i + 1}" for i in range(6)]] + [[l] + [""] * 6 for l in "abcd"]
    t = Table(grid, colWidths=[40] + [50] * 6, rowHeights=20)
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.6, RULE), ("FONTNAME", (0, 0), (-1, -1), "UB-B"),
                           ("FONTSIZE", (0, 0), (-1, -1), 9.5), ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("BACKGROUND", (0, 0), (-1, 0), TINT), ("BACKGROUND", (0, 0), (0, -1), TINT)]))
    s.append(t)
    doc.multiBuild(s)


def build_key(out):
    doc = Doc(out, "Gabarito — uso do instrutor", "Gabarito comentado", "Uso exclusivo do instrutor")
    s = []
    start(s)
    exam_header(s, "Gabarito comentado", "Uso exclusivo do instrutor. Não distribuir aos participantes.")
    grid = [[f"Q{i + 1}" for i in range(6)], [("abcd"[q["ans"]]).upper() for q in EXAM]]
    t = Table(grid, colWidths=[CW / 6] * 6, rowHeights=[22, 34])
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.6, RULE), ("FONTNAME", (0, 0), (-1, -1), "UB-B"),
                           ("FONTSIZE", (0, 0), (-1, 0), 9.5), ("FONTSIZE", (0, 1), (-1, 1), 18),
                           ("TEXTCOLOR", (0, 1), (-1, 1), TXTC[3]), ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("BACKGROUND", (0, 0), (-1, 0), TINT)]))
    s += [t, Spacer(1, 16)]
    mods = [m["title"] for m in MODULES]
    for i, q in enumerate(EXAM):
        items = [P(f"Questão {i + 1}  ·  Módulo {i + 1}: {mods[i]}", "label", textColor=TXTC[i + 1]), Spacer(1, 2), P(q["q"], "q")]
        for j, o in enumerate(q["opts"]):
            ok = j == q["ans"]
            mark = SYM("\u25CF" if ok else "\u25EF")
            items.append(P(f"{mark}  <b>{'abcd'[j]})</b> {o}", "p", spaceAfter=2, textColor=TXTC[3] if ok else SOFT, fontName="UB-B" if ok else "UB"))
        items.append(box("Por que esta é a resposta", q["why"], 3))
        s.append(KeepTogether(items))
    s.append(P("Critérios de nota sugeridos", "h2"))
    s += table(["Acertos", "Conceito", "Sugestão"], [("6", "Excelente", "Domínio completo do processo"), ("4–5", "Bom", "Revisar os módulos das questões erradas"), ("0–3", "Em desenvolvimento", "Retomar a apostila e os exercícios")], 3, [80, 130, CW - 210])
    doc.multiBuild(s)


# ---------------------------------------------------------------- roteiro do instrutor
def build_guide(out):
    doc = Doc(out, "Roteiro do instrutor", "Roteiro do instrutor", "Cronograma, dinâmicas e materiais")
    s = []
    start(s)
    s.append(P("Visão geral", "h1"))
    s.append(P("Treinamento prático e in-company, de 12 horas", "sub"))
    s += table(["Item", "Descrição"], [
        ("Objetivo", "Capacitar equipes a resolver problemas do negócio com foco no cliente, usando as 5 etapas do Design Thinking."),
        ("Público", "Gestores, equipes de atendimento, vendas, marketing, produto, TI e operações. Não exige conhecimento prévio."),
        ("Formato", "3 encontros de 4 horas: exposição curta, vídeos, dinâmicas em equipe e um desafio real da empresa."),
        ("Turma ideal", "12 a 30 participantes, em equipes de 4 a 6 pessoas."),
        ("Avaliação", "Desafio Final (pitch com rubrica de 100 pontos) + prova individual de 6 questões."),
        ("Entregáveis", "Slides, apostila do participante, prova, gabarito e este roteiro."),
    ], 1, [100, CW - 100])
    s.append(P("Antes do treinamento", "h2"))
    s += bullets([
        "Converse com o contratante e colete 2 ou 3 problemas reais da empresa para usar nas dinâmicas e no Desafio Final.",
        "Grave e insira os vídeos nos slides (botões \u201cAssistir vídeo\u201d abrem slides ocultos) e publique os links em build/videos.json para a apostila.",
        "Imprima a apostila (opcional), a prova (1 por pessoa) e o canvas do desafio em A3 (1 por equipe).",
        "Teste o projetor, o som e o modo de apresentação: os slides de vídeo ficam ocultos e só aparecem pelo botão.",
        "As animações dos slides rodam sozinhas. Só dois momentos pedem um clique: na correção da prova (revela a alternativa correta) e na abertura do estudo de caso (mostra a pergunta à turma).",
        "Personalize o slide \u201cSobre o instrutor\u201d e troque os exemplos genéricos por exemplos do setor do cliente.",
    ], LIME)
    s.append(P("Materiais", "h2"))
    s += bullets(MATERIALS, TEAL)

    acts = {}
    for m in MODULES:
        for sp in m["slides"]:
            if sp["type"] == "activity":
                acts.setdefault(m["encontro"], []).append(sp)
    for i, enc in enumerate(SCHEDULE):
        s += [PageBreak(), Badge(i + 1, "Encontro · 4 horas", [LIME, TEAL, BLUE][i]), P(enc["title"], "h1"), P(enc["goal"], "sub")]
        s += table(["Início", "Fim", "Atividade", "Recursos"], enc["rows"], [1, 4, 5][i], [44, 44, CW - 88 - 130, 130])
        for a in acts.get(i + 1, []):
            parts = [P(f"{a['title']}  ·  {a['time']}", "h2"), P(f"<b>Objetivo:</b> {a['goal']}")]
            parts += bullets([f"<b>Passo {k + 1}.</b> {st}" for k, st in enumerate(a["steps"])], [LIME, TEAL, BLUE][i])
            parts.append(P(f"<b>Materiais:</b> {a['materials']}"))
            parts.append(box("Dica do instrutor", a["notes"], [1, 4, 5][i]))
            s.append(KeepTogether(parts))
        if i == 2:
            s.append(P("Desafio Final: condução", "h2"))
            s += bullets([
                "Mostre a rubrica ANTES de começar: as equipes precisam saber como serão avaliadas.",
                "Anuncie o tempo de cada etapa (cronograma do slide) e circule entre as equipes fazendo perguntas, sem dar respostas.",
                "No teste cruzado, cada equipe testa o protótipo de outra: isso gera aprendizado real e movimenta a sala.",
                "Pitch: 5 minutos por equipe + 2 minutos de perguntas. Use a ficha de pontuação a seguir.",
            ], BLUE)

    s += [PageBreak(), P("Ficha de pontuação do Desafio Final", "h1"), P("Pontue cada equipe de 0 a 25 em cada critério.", "sub")]
    crit = [c[0] for c in CHALLENGE["rubric"]]
    data = [["Equipe"] + crit + ["Total"]] + [["" for _ in range(len(crit) + 2)] for _ in range(6)]
    t = Table([[P(c, "cellh") for c in data[0]]] + data[1:], colWidths=[100] + [(CW - 160) / 4] * 4 + [60], rowHeights=[None] + [34] * 6)
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), TXTC[5]), ("GRID", (0, 0), (-1, -1), 0.6, RULE),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    s += [t, Spacer(1, 12)]
    s += table(["Critério", "O que observar"], [(a, b) for a, b, _ in CHALLENGE["rubric"]], 5, [130, CW - 130])
    s.append(P("Após o treinamento", "h2"))
    s += bullets([
        "Aplique um formulário rápido de avaliação (satisfação, aplicabilidade, o que melhorar).",
        "Envie ao contratante um resumo com os problemas trabalhados, as ideias priorizadas e as notas.",
        "Ofereça um acompanhamento: implementar a melhor ideia e medir o resultado com dados.",
    ], TEAL)
    doc.multiBuild(s)


# ---------------------------------------------------------------- vídeos interativos (pypdf)
def add_video_widgets(pdf_path, urls):
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import (DictionaryObject, NameObject, NumberObject, ArrayObject, FloatObject,
                               TextStringObject, DecodedStreamObject, BooleanObject)

    r = PdfReader(str(pdf_path))
    w = PdfWriter(clone_from=r)

    def font(base):
        f = DictionaryObject({NameObject("/Type"): NameObject("/Font"), NameObject("/Subtype"): NameObject("/Type1"),
                              NameObject("/BaseFont"): NameObject(base), NameObject("/Encoding"): NameObject("/WinAnsiEncoding")})
        return w._add_object(f)

    helv, helvb = font("/Helvetica"), font("/Helvetica-Bold")

    def esc(t):
        return t.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)").encode("cp1252", "replace")

    def rgbf(hx):
        return " ".join(f"{int(hx[i:i + 2], 16) / 255:.3f}" for i in (0, 2, 4))

    def xobj(wd, ht, ops):
        st = DecodedStreamObject()
        st.set_data(ops)
        st.update({NameObject("/Type"): NameObject("/XObject"), NameObject("/Subtype"): NameObject("/Form"),
                   NameObject("/BBox"): ArrayObject([FloatObject(0), FloatObject(0), FloatObject(wd), FloatObject(ht)]),
                   NameObject("/Resources"): DictionaryObject({NameObject("/Font"): DictionaryObject({NameObject("/F1"): helv, NameObject("/F2"): helvb})})})
        return w._add_object(st)

    def text_w(t, size, bold=False):
        return pdfmetrics.stringWidth(t, "Helvetica-Bold" if bold else "Helvetica", size)

    fields = ArrayObject()
    for key, slot in VIDEO_SLOTS.items():
        page = w.pages[slot["page"]]
        col = slot["color"]
        url = urls.get(key, "").strip()
        x1, y1, x2, y2 = slot["video"]
        vw, vh = x2 - x1, y2 - y1
        title = VIDEOS[key]
        line2 = "Clique para assistir" if url else "Vídeo disponível em breve"
        u = (url[:70] + "...") if len(url) > 73 else url
        ops = [b"q", f"{rgbf(B.BG)} rg 0 0 {vw:.1f} {vh:.1f} re f".encode(),
               f"{rgbf(col)} rg".encode()]
        cx, cy, rr = vw / 2, vh / 2 + 22, 26
        k = 0.5523 * rr
        ops.append(f"{cx + rr:.1f} {cy:.1f} m {cx + rr:.1f} {cy + k:.1f} {cx + k:.1f} {cy + rr:.1f} {cx:.1f} {cy + rr:.1f} c "
                   f"{cx - k:.1f} {cy + rr:.1f} {cx - rr:.1f} {cy + k:.1f} {cx - rr:.1f} {cy:.1f} c "
                   f"{cx - rr:.1f} {cy - k:.1f} {cx - k:.1f} {cy - rr:.1f} {cx:.1f} {cy - rr:.1f} c "
                   f"{cx + k:.1f} {cy - rr:.1f} {cx + rr:.1f} {cy - k:.1f} {cx + rr:.1f} {cy:.1f} c f".encode())
        ops.append(f"{rgbf(B.BG)} rg {cx - 8:.1f} {cy - 11:.1f} m {cx - 8:.1f} {cy + 11:.1f} l {cx + 12:.1f} {cy:.1f} l f".encode())
        for txt, size, bold, color, yy in [(title, 11, True, B.TEXT, vh / 2 - 22), (line2, 9.5, False, col, vh / 2 - 40), (u, 7.5, False, B.MUTED, 14)]:
            if not txt:
                continue
            tw = text_w(txt, size, bold)
            ops.append(b"BT /" + (b"F2 " if bold else b"F1 ") + f"{size} Tf {rgbf(color)} rg {(vw - tw) / 2:.1f} {yy:.1f} Td (".encode() + esc(txt) + b") Tj ET")
        ops.append(b"Q")
        ap = xobj(vw, vh, b"\n".join(ops))
        video = DictionaryObject({
            NameObject("/Type"): NameObject("/Annot"), NameObject("/Subtype"): NameObject("/Widget"),
            NameObject("/FT"): NameObject("/Btn"), NameObject("/Ff"): NumberObject(65536),
            NameObject("/T"): TextStringObject(f"video_{key}"), NameObject("/TU"): TextStringObject(f"Vídeo: {title}"),
            NameObject("/Rect"): ArrayObject([FloatObject(v) for v in (x1, y1, x2, y2)]),
            NameObject("/F"): NumberObject(4), NameObject("/P"): page.indirect_reference,
            NameObject("/AP"): DictionaryObject({NameObject("/N"): ap}),
            NameObject("/MK"): DictionaryObject(),
        })
        if url:
            video[NameObject("/A")] = DictionaryObject({NameObject("/S"): NameObject("/URI"), NameObject("/URI"): TextStringObject(url)})
        vref = w._add_object(video)

        tx1, ty1, tx2, ty2 = slot["toggle"]
        tw_, th_ = tx2 - tx1, ty2 - ty1
        lbl = "Mostrar / ocultar vídeo"
        lw = text_w(lbl, 8.5, True)
        r_ = th_ / 2
        tops = [b"q", f"{rgbf(col)} rg".encode(),
                f"{r_:.1f} 0 m {tw_ - r_:.1f} 0 l {tw_:.1f} 0 {tw_:.1f} {th_:.1f} {tw_ - r_:.1f} {th_:.1f} c {r_:.1f} {th_:.1f} l 0 {th_:.1f} 0 0 {r_:.1f} 0 c f".encode(),
                b"BT /F2 8.5 Tf " + f"{rgbf(B.BG)} rg {(tw_ - lw) / 2:.1f} {th_ / 2 - 3:.1f} Td (".encode() + esc(lbl) + b") Tj ET", b"Q"]
        tap = xobj(tw_, th_, b"\n".join(tops))
        js = (f'var f = this.getField("video_{key}"); '
              f'f.display = (f.display == display.hidden) ? display.visible : display.hidden;')
        toggle = DictionaryObject({
            NameObject("/Type"): NameObject("/Annot"), NameObject("/Subtype"): NameObject("/Widget"),
            NameObject("/FT"): NameObject("/Btn"), NameObject("/Ff"): NumberObject(65536),
            NameObject("/T"): TextStringObject(f"toggle_{key}"), NameObject("/TU"): TextStringObject("Mostrar ou ocultar o vídeo"),
            NameObject("/Rect"): ArrayObject([FloatObject(v) for v in (tx1, ty1, tx2, ty2)]),
            NameObject("/F"): NumberObject(0),  # visível na tela, não sai na impressão
            NameObject("/P"): page.indirect_reference,
            NameObject("/AP"): DictionaryObject({NameObject("/N"): tap}),
            NameObject("/MK"): DictionaryObject(),
            NameObject("/A"): DictionaryObject({NameObject("/S"): NameObject("/JavaScript"), NameObject("/JS"): TextStringObject(js)}),
        })
        tref = w._add_object(toggle)
        if "/Annots" not in page:
            page[NameObject("/Annots")] = ArrayObject()
        page["/Annots"].extend([vref, tref])
        fields.extend([vref, tref])

    w._root_object[NameObject("/AcroForm")] = DictionaryObject({
        NameObject("/Fields"): fields, NameObject("/NeedAppearances"): BooleanObject(False),
        NameObject("/DR"): DictionaryObject({NameObject("/Font"): DictionaryObject({NameObject("/Helv"): helv})}),
    })
    with open(pdf_path, "wb") as fh:
        w.write(fh)
    return len(VIDEO_SLOTS)


def load_urls():
    p = B.BUILD / "videos.json"
    data = json.loads(p.read_text(encoding="utf-8"))
    return {k: v for k, v in data.items() if not k.startswith("_")}


def build_all_pdfs():
    out = B.ROOT
    handout = out / "DT_na_Pratica_Apostila.pdf"
    build_handout(handout)
    n = add_video_widgets(handout, load_urls())
    build_exam(out / "DT_na_Pratica_Prova.pdf")
    build_key(out / "DT_na_Pratica_Gabarito.pdf")
    build_guide(out / "DT_na_Pratica_Roteiro_Instrutor.pdf")
    print(f"PDFs gerados em {out} (apostila com {n} espaços de vídeo)")


if __name__ == "__main__":
    build_all_pdfs()
