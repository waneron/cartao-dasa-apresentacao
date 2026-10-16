"""One page da etapa Gold da carga Detalhe, no visual do material Dasa."""

from reportlab.lib.colors import Color, HexColor, white
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

pdfmetrics.registerFont(TTFont("Inter", "/usr/share/fonts/truetype/macos/Inter-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Inter-Med", "/usr/share/fonts/truetype/macos/Inter-Medium.ttf"))
pdfmetrics.registerFont(TTFont("Inter-Sem", "/usr/share/fonts/truetype/macos/Inter-SemiBold.ttf"))
pdfmetrics.registerFont(TTFont("Inter-Bold", "/usr/share/fonts/truetype/macos/Inter-Bold.ttf"))

NAVY = HexColor("#002060")
INK = HexColor("#000E40")
SKY = HexColor("#ABE6FF")
SAND = HexColor("#FFF6D9")
MIST = HexColor("#F2F2F2")
LINE = HexColor("#D5DBE3")
MUTED = HexColor("#3E4A5C")
PAGE = (960, 540)
OUT = "/workspace/Carga-Detalhe-Gold-onepage.pdf"


def draw_wordmark(c, x, y, h):
    """Símbolo DASA no azul do logo: D de um bojo, A aberto, S, A aberto."""
    c.saveState()
    c.setFillColor(INK)
    c.setStrokeColor(INK)
    c.setLineCap(0)
    c.setLineJoin(0)
    c.setMiterLimit(2)
    t = h * 0.155

    def letter_d(ox):
        w = h * 0.68
        p = c.beginPath()
        p.moveTo(ox, y)
        p.lineTo(ox, y + h)
        p.lineTo(ox + w * 0.36, y + h)
        p.curveTo(ox + w * 0.92, y + h * 1.02, ox + w * 1.02, y + h * 0.72, ox + w * 0.98, y + h * 0.50)
        p.curveTo(ox + w * 1.02, y + h * 0.28, ox + w * 0.92, y - h * 0.02, ox + w * 0.36, y)
        p.close()
        p.moveTo(ox + t, y + t)
        p.curveTo(ox + w * 0.78, y + t * 0.2, ox + w * 0.80, y + h * 0.28, ox + w * 0.78, y + h * 0.50)
        p.curveTo(ox + w * 0.80, y + h * 0.72, ox + w * 0.78, y + h - t * 0.2, ox + t, y + h - t)
        p.lineTo(ox + t, y + t)
        p.close()
        c.drawPath(p, fill=1, stroke=0)
        return w

    def letter_a(ox):
        w = h * 0.56
        c.setLineWidth(t)
        c.setLineCap(0)
        c.setLineJoin(0)
        peak_x = ox + w * 0.50
        c.line(ox + t * 0.15, y, peak_x, y + h)
        c.line(peak_x, y + h, ox + w - t * 0.15, y)
        return w

    def letter_s(ox):
        w = h * 0.42
        c.setLineWidth(t * 0.95)
        c.setLineCap(1)
        p = c.beginPath()
        p.moveTo(ox + w * 0.92, y + h * 0.74)
        p.curveTo(ox + w * 0.55, y + h * 1.08, ox - w * 0.08, y + h * 0.92, ox + w * 0.16, y + h * 0.58)
        p.curveTo(ox + w * 0.42, y + h * 0.28, ox + w * 0.95, y + h * 0.46, ox + w * 0.88, y + h * 0.30)
        p.curveTo(ox + w * 0.72, y + h * 0.02, ox + w * 0.05, y - h * 0.06, ox + w * 0.08, y + h * 0.22)
        c.drawPath(p, stroke=1, fill=0)
        return w

    gap = h * 0.16
    cursor = x
    cursor += letter_d(cursor) + gap
    cursor += letter_a(cursor) + gap * 0.85
    cursor += letter_s(cursor) + gap
    cursor += letter_a(cursor)
    c.restoreState()
    return cursor - x


def wrap(text, font, size, width):
    words = text.split()
    lines, cur = [], ""
    for word in words:
        trial = word if not cur else f"{cur} {word}"
        if pdfmetrics.stringWidth(trial, font, size) <= width:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = word
    if cur:
        lines.append(cur)
    return lines


def draw_paragraph(c, text, x, y, font, size, width, color, leading):
    c.setFillColor(color)
    c.setFont(font, size)
    for line in wrap(text, font, size, width):
        c.drawString(x, y, line)
        y -= leading
    return y


def card(c, x, y, w, h, fill, border=None):
    c.setFillColor(fill)
    if border:
        c.setStrokeColor(border)
        c.setLineWidth(1)
        c.roundRect(x, y, w, h, 4, fill=1, stroke=1)
    else:
        c.setStrokeColor(fill)
        c.roundRect(x, y, w, h, 4, fill=1, stroke=0)


def kpi(c, x, y, w, h, fill, value, label, note, value_color, label_color, border=None):
    card(c, x, y, w, h, fill, border)
    c.setFillColor(value_color)
    c.setFont("Inter-Bold", 22)
    c.drawString(x + 14, y + h - 34, value)
    c.setFillColor(label_color)
    c.setFont("Inter-Sem", 8)
    c.drawString(x + 14, y + 24, label)
    c.setFont("Inter", 7.5)
    c.drawString(x + 14, y + 12, note)


def section_head(c, x, y, w, text):
    c.setFillColor(NAVY)
    c.rect(x, y - 4, w, 18, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("Inter-Sem", 8.5)
    c.drawString(x + 8, y + 1, text)


def build():
    c = canvas.Canvas(OUT, pagesize=PAGE)
    c.setTitle("Carga Detalhe · etapa Gold")
    c.setAuthor("Data Provider")
    width, height = PAGE

    # faixa superior no azul do cabeçalho Dasa
    c.setFillColor(NAVY)
    c.rect(0, height - 8, width, 8, fill=1, stroke=0)

    logo_w = draw_wordmark(c, 32, height - 46, 18)
    c.setFillColor(INK)
    c.setFont("Inter-Bold", 18)
    c.drawString(32 + logo_w + 16, height - 40, "Carga Detalhe")
    c.setFillColor(NAVY)
    c.setFont("Inter", 18)
    title_w = pdfmetrics.stringWidth("Carga Detalhe", "Inter-Bold", 18)
    c.drawString(32 + logo_w + 16 + title_w + 8, height - 40, "·  etapa Gold")

    c.setFillColor(MUTED)
    c.setFont("Inter", 8.5)
    c.drawString(
        32 + logo_w + 16,
        height - 56,
        "Notebook da carga Detalhe  ·  banco CORPORATE_PRD  ·  15/10 e 16/10  ·  horário de Brasília",
    )

    c.setStrokeColor(NAVY)
    c.setLineWidth(1.25)
    c.line(32, height - 68, width - 32, height - 68)

    # indicadores
    cards_y = height - 162
    card_h = 78
    gap = 10
    card_w = (width - 64 - 3 * gap) / 4
    x0 = 32
    kpi(
        c, x0, cards_y, card_w, card_h, NAVY,
        "75,41%", "MENOS TEMPO DE EXECUÇÃO", "52 min 55 s  →  13 min 01 s",
        white, SKY,
    )
    kpi(
        c, x0 + (card_w + gap), cards_y, card_w, card_h, SKY,
        "13 min 01 s", "DURAÇÃO EM 16/10", "Concluída com sucesso às 02h13",
        INK, NAVY,
    )
    kpi(
        c, x0 + 2 * (card_w + gap), cards_y, card_w, card_h, white,
        "39 min 54 s", "A MENOS NESTA RODADA", "Diferença entre as duas durações",
        INK, NAVY, NAVY,
    )
    kpi(
        c, x0 + 3 * (card_w + gap), cards_y, card_w, card_h, SAND,
        "75,7%", "MENOS CPU ACUMULADA", "163,4  →  39,7 vCore-min",
        INK, NAVY,
    )

    # tabela
    table_x = 32
    table_w = 560
    head_y = cards_y - 28
    section_head(c, table_x, head_y, table_w, "COMPARATIVO  ·  AZURE MONITOR  ·  CAPACIDADE DE 10 vCORES NAS DUAS RODADAS")

    rows = [
        ("Janela da Gold", "01h58 – 02h51", "02h00 – 02h13", False),
        ("Duração medida", "52 min 55 s", "13 min 01 s", True),
        ("CPU média do banco", "30,88%", "30,49%", False),
        ("Pico de CPU nos minutos da carga", "98%", "80%", True),
        ("CPU acumulada estimada", "163,4 vCore-min", "39,7 vCore-min", True),
        ("Leitura física média", "2,14%", "4,30%", False),
        ("Gravação do log média", "9,78%", "9,86%", False),
    ]
    col_w = (170, 195, 195)
    row_h = 26
    y = head_y - 4 - row_h
    c.setFillColor(HexColor("#1A3A7A"))
    c.rect(table_x, y, table_w, row_h, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("Inter-Sem", 8)
    headers = ("Indicador", "15/10  ·  rodada anterior", "16/10  ·  publicação incremental")
    xx = table_x
    for i, header in enumerate(headers):
        c.drawString(xx + 8, y + 9, header)
        xx += col_w[i]

    y -= row_h
    for i, (label, before, after, emphasis) in enumerate(rows):
        bg = SKY if emphasis else (white if i % 2 == 0 else MIST)
        c.setFillColor(bg)
        c.rect(table_x, y, table_w, row_h, fill=1, stroke=0)
        c.setStrokeColor(LINE)
        c.setLineWidth(0.3)
        c.line(table_x, y, table_x + table_w, y)
        c.setFillColor(INK)
        c.setFont("Inter-Sem" if emphasis else "Inter", 8)
        c.drawString(table_x + 8, y + 9, label)
        c.setFont("Inter", 8)
        c.drawString(table_x + col_w[0] + 8, y + 9, before)
        c.setFont("Inter-Sem", 8)
        c.drawString(table_x + col_w[0] + col_w[1] + 8, y + 9, after)
        y -= row_h

    c.setStrokeColor(NAVY)
    c.setLineWidth(0.6)
    c.rect(table_x, y, table_w, (head_y - 4) - y, fill=0, stroke=1)

    # barras de duração, alinhadas ao tom da tabela
    bar_top = y - 16
    c.setFillColor(INK)
    c.setFont("Inter-Sem", 8)
    c.drawString(table_x, bar_top, "Duração das janelas")
    max_bar = 360
    dur_before = 52 + 55 / 60
    dur_after = 13 + 1 / 60
    scale = max_bar / dur_before
    by = bar_top - 16
    c.setFillColor(NAVY)
    c.roundRect(table_x, by, dur_before * scale, 8, 2, fill=1, stroke=0)
    c.setFillColor(MUTED)
    c.setFont("Inter", 7.5)
    c.drawString(table_x + dur_before * scale + 6, by, "15/10  ·  52 min 55 s")
    by -= 14
    c.setFillColor(HexColor("#3D8FBF"))
    c.roundRect(table_x, by, max(dur_after * scale, 8), 8, 2, fill=1, stroke=0)
    c.setFillColor(MUTED)
    c.drawString(table_x + dur_after * scale + 6, by, "16/10  ·  13 min 01 s")

    # painel direito
    panel_x = table_x + table_w + 16
    panel_w = width - 32 - panel_x
    section_head(c, panel_x, head_y, panel_w, "O QUE MUDOU")

    bullets = [
        "Cerca de sete truncates retirados do notebook da etapa Gold.",
        "Código reestruturado para publicação incremental nas sete tabelas.",
        "A rodada de 16/10 terminou com sucesso às 02h13.",
    ]
    ty = head_y - 28
    for bullet in bullets:
        c.setFillColor(NAVY)
        c.circle(panel_x + 7, ty + 3, 2.2, fill=1, stroke=0)
        ty = draw_paragraph(c, bullet, panel_x + 16, ty, "Inter", 8, panel_w - 22, INK, 11)
        ty -= 8

    box_top = ty - 4
    box_h = 118
    box_y = box_top - box_h
    card(c, panel_x, box_y, panel_w, box_h, SAND)
    c.setFillColor(NAVY)
    c.setFont("Inter-Sem", 8)
    c.drawString(panel_x + 10, box_y + box_h - 16, "COMO LER O NÚMERO")
    reading = (
        "A CPU média ficou estável, de 30,88% para 30,49%, com os mesmos 10 núcleos. "
        "O pico nos minutos da carga caiu de 98% para 80%. "
        "A CPU acumulada estimada caiu de 163,4 para 39,7 vCore-minutos porque a segunda execução durou 13 min 01 s, contra 52 min 55 s. "
        "A leitura física média subiu de 2,14% para 4,30%. A gravação do log permaneceu em 9,8%."
    )
    draw_paragraph(c, reading, panel_x + 10, box_y + box_h - 32, "Inter", 7.6, panel_w - 20, INK, 10)

    # rodapé
    foot_y = 28
    c.setStrokeColor(LINE)
    c.setLineWidth(0.6)
    c.line(32, foot_y + 22, width - 32, foot_y + 22)
    note = (
        "Fonte: Azure Monitor, médias de 1 minuto. Métricas do banco inteiro, incluindo processos concorrentes. "
        "A CPU acumulada é a integração do uso ao longo de cada janela. Volumes e durações diferem entre as rodadas. "
        "O comparativo não isola o consumo exclusivo da Gold e não mede redução de custo."
    )
    draw_paragraph(c, note, 32, foot_y + 8, "Inter", 6.6, width - 64, MUTED, 8.2)

    c.showPage()
    c.save()
    print(OUT)


if __name__ == "__main__":
    build()
