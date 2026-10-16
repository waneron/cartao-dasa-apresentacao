"""One page executiva da publicação, no visual das apresentações Dasa."""

from reportlab.lib.colors import HexColor, white
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

pdfmetrics.registerFont(TTFont("Inter", "/usr/share/fonts/truetype/macos/Inter-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Inter-Med", "/usr/share/fonts/truetype/macos/Inter-Medium.ttf"))
pdfmetrics.registerFont(TTFont("Inter-Sem", "/usr/share/fonts/truetype/macos/Inter-SemiBold.ttf"))
pdfmetrics.registerFont(TTFont("Inter-Bold", "/usr/share/fonts/truetype/macos/Inter-Bold.ttf"))

NAVY = HexColor("#1B2A4A")
TEAL = HexColor("#2EC4B6")
CARD = HexColor("#F2F4F7")
MIST = HexColor("#F7F8FA")
LINE = HexColor("#E4E8EE")
MUTED = HexColor("#5C6B7A")
PALE = HexColor("#E7F6F4")
PAGE = (960, 540)
OUT = "/workspace/Carga-Detalhe-Gold-onepage.pdf"
LOGO = "/workspace/assets/dasa-logo.png"
FOOTER = "/workspace/assets/dasa-footer.png"


def card(c, x, y, w, h):
    c.setFillColor(CARD)
    c.setStrokeColor(CARD)
    c.roundRect(x, y, w, h, 8, fill=1, stroke=0)
    c.setFillColor(TEAL)
    c.rect(x + 10, y + 12, 3.5, h - 24, fill=1, stroke=0)


def build():
    c = canvas.Canvas(OUT, pagesize=PAGE)
    c.setTitle("Tempo de publicação reduzido em 75%")
    c.setAuthor("Data Provider")
    width, height = PAGE

    logo = ImageReader(LOGO)
    logo_h = 16
    logo_w = logo_h * (90 / 25)
    c.drawImage(logo, 36, height - 40, logo_w, logo_h, mask="auto", preserveAspectRatio=True)

    c.setStrokeColor(NAVY)
    c.setLineWidth(1.1)
    c.line(36, height - 52, width - 36, height - 52)

    c.setFillColor(TEAL)
    c.rect(36, height - 96, 4, 28, fill=1, stroke=0)
    c.setFillColor(NAVY)
    c.setFont("Inter-Bold", 24)
    c.drawString(52, height - 90, "Tempo de publicação reduzido em 75%")

    cards = [
        ("75,41%", "menos tempo"),
        ("13 min 01 s", "duração em 16/10"),
        ("39 min 54 s", "a menos nesta rodada"),
        ("75,7%", "menos CPU acumulada"),
    ]
    gap = 12
    card_w = (width - 72 - 3 * gap) / 4
    card_h = 78
    card_y = height - 196
    for i, (value, label) in enumerate(cards):
        x = 36 + i * (card_w + gap)
        card(c, x, card_y, card_w, card_h)
        c.setFillColor(NAVY)
        c.setFont("Inter-Bold", 18)
        c.drawString(x + 22, card_y + 38, value)
        c.setFillColor(MUTED)
        c.setFont("Inter", 9)
        c.drawString(x + 22, card_y + 20, label)

    # comparativo
    table_x = 36
    table_w = 560
    section_y = card_y - 36
    c.setFillColor(NAVY)
    c.setFont("Inter-Sem", 11)
    c.drawString(table_x, section_y, "Comparativo")

    rows = [
        ("Janela", "01h58 – 02h51", "02h00 – 02h13", False),
        ("Duração", "52 min 55 s", "13 min 01 s", True),
        ("Capacidade", "10 vCores", "10 vCores", False),
        ("CPU média", "30,88%", "30,49%", False),
        ("Pico de CPU", "98%", "80%", True),
        ("CPU acumulada", "163,4 vCore-min", "39,7 vCore-min", True),
        ("Leitura física", "2,14%", "4,30%", False),
        ("Gravação do log", "9,78%", "9,86%", False),
    ]
    col_w = (168, 196, 196)
    row_h = 22
    y = section_y - 14 - row_h
    c.setFillColor(NAVY)
    c.rect(table_x, y, table_w, row_h, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("Inter-Sem", 8)
    headers = ("Indicador", "15/10", "16/10")
    xx = table_x
    for i, header in enumerate(headers):
        c.drawString(xx + 12, y + 7, header)
        xx += col_w[i]

    y -= row_h
    for i, (label, before, after, emphasis) in enumerate(rows):
        c.setFillColor(PALE if emphasis else (white if i % 2 == 0 else MIST))
        c.rect(table_x, y, table_w, row_h, fill=1, stroke=0)
        c.setStrokeColor(LINE)
        c.setLineWidth(0.4)
        c.line(table_x, y, table_x + table_w, y)
        c.setFillColor(NAVY)
        c.setFont("Inter-Sem" if emphasis else "Inter", 8.5)
        c.drawString(table_x + 12, y + 6.5, label)
        c.setFont("Inter", 8.5)
        c.drawString(table_x + col_w[0] + 12, y + 6.5, before)
        c.setFont("Inter-Sem" if emphasis else "Inter", 8.5)
        c.drawString(table_x + col_w[0] + col_w[1] + 12, y + 6.5, after)
        y -= row_h

    # o que mudou
    panel_x = table_x + table_w + 16
    panel_w = width - 36 - panel_x
    panel_top = section_y + 14
    panel_h = 168
    panel_bottom = panel_top - panel_h
    card(c, panel_x, panel_bottom, panel_w, panel_h)
    c.setFillColor(NAVY)
    c.setFont("Inter-Sem", 11)
    c.drawString(panel_x + 22, panel_top - 28, "O que mudou")
    bullets = [
        "Cerca de sete truncates retirados.",
        "Código reestruturado para publicação incremental nas sete tabelas.",
        "A execução de 16/10 terminou às 02h13.",
    ]
    ty = panel_top - 52
    for bullet in bullets:
        c.setFillColor(TEAL)
        c.circle(panel_x + 24, ty + 3, 2.2, fill=1, stroke=0)
        c.setFillColor(NAVY)
        c.setFont("Inter", 9)
        # wrap inside the card
        words = bullet.split()
        line = ""
        for word in words:
            trial = word if not line else f"{line} {word}"
            if c.stringWidth(trial, "Inter", 9) <= panel_w - 48:
                line = trial
            else:
                c.drawString(panel_x + 34, ty, line)
                ty -= 12
                line = word
        c.drawString(panel_x + 34, ty, line)
        ty -= 20

    c.setFillColor(MUTED)
    c.setFont("Inter", 7)
    c.drawString(
        36,
        46,
        "Azure Monitor, médias de 1 minuto. Métricas do banco inteiro. A CPU acumulada é estimada na janela e não mede custo.",
    )

    footer = ImageReader(FOOTER)
    footer_h = width * (54 / 1385)
    c.drawImage(footer, 0, 0, width, footer_h, preserveAspectRatio=False, anchor="sw", mask="auto")

    c.showPage()
    c.save()
    print(OUT)


if __name__ == "__main__":
    build()
