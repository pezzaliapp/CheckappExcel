"""Genera l'immagine Open Graph e il favicon PNG per CheckappExcel."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

BASE = Path(__file__).resolve().parent

# ---------------------------------------------------------------------------
# OG image 1200x630
# ---------------------------------------------------------------------------
W, H = 1200, 630
img = Image.new("RGB", (W, H), "#1F77B4")
draw = ImageDraw.Draw(img)

# Sfondo: gradiente lineare blu -> verde (come l'header)
C1 = (0x1F, 0x77, 0xB4)   # blu header
C2 = (0x2C, 0xA0, 0x2C)   # verde header
for y in range(H):
    t = y / (H - 1)
    r = int(C1[0] + (C2[0] - C1[0]) * t)
    g = int(C1[1] + (C2[1] - C1[1]) * t)
    b = int(C1[2] + (C2[2] - C1[2]) * t)
    draw.line([(0, y), (W, y)], fill=(r, g, b))

# Diagonale: banda decorativa con "celle colorate" in basso
cell_colors = ["#FFF2CC", "#C6EFCE", "#FCE4D6", "#D9E1F2", "#F8CBAD"]
cw, ch = 170, 70
gy = H - ch - 40
for i, c in enumerate(cell_colors):
    x = 60 + i * (cw + 14)
    draw.rounded_rectangle([x, gy, x + cw, gy + ch], radius=10, fill=c, outline="#FFFFFF", width=2)

# Tenta di caricare font di sistema; fallback a default
def load_font(names, size):
    for name in names:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()

font_title = load_font(
    ["DejaVuSans-Bold.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
     "Arial Bold.ttf", "Arial.ttf"], 96)
font_sub   = load_font(
    ["DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
     "Arial.ttf"], 40)
font_small = load_font(
    ["DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
     "Arial.ttf"], 28)

# Icona a sinistra: due rettangoli sovrapposti (foglio + foglio) con check
icon_x, icon_y = 80, 110
# foglio 1 (bianco, leggera rotazione simulata con ombra)
draw.rounded_rectangle([icon_x, icon_y, icon_x + 150, icon_y + 190],
                        radius=14, fill="#FFFFFF")
# "righe" del primo foglio
for i in range(4):
    y = icon_y + 30 + i * 30
    draw.rectangle([icon_x + 18, y, icon_x + 132, y + 12], fill="#DDE3EB")
# foglio 2 sovrapposto
draw.rounded_rectangle([icon_x + 60, icon_y + 50, icon_x + 210, icon_y + 240],
                        radius=14, fill="#FFFFFF")
for i in range(4):
    y = icon_y + 80 + i * 30
    draw.rectangle([icon_x + 78, y, icon_x + 192, y + 12], fill="#C6EFCE")
# check verde sopra il secondo foglio
check_x, check_y = icon_x + 135, icon_y + 55
draw.ellipse([check_x, check_y, check_x + 60, check_y + 60], fill="#2CA02C", outline="#FFFFFF", width=3)
draw.line([(check_x + 15, check_y + 32), (check_x + 27, check_y + 42), (check_x + 47, check_y + 20)],
           fill="#FFFFFF", width=6)

# Titolo
title_x = 340
draw.text((title_x, 150), "CheckappExcel", fill="#FFFFFF", font=font_title)

# Sottotitolo
draw.text((title_x, 260), "Confronta listini Excel/CSV",
           fill="#FFFFFF", font=font_sub)
draw.text((title_x, 312), "per codice prodotto — gratis, nel browser",
           fill="#E8F0FA", font=font_small)

# URL in basso a destra
url = "alessandropezzali.it/CheckappExcel"
bbox = draw.textbbox((0, 0), url, font=font_small)
url_w = bbox[2] - bbox[0]
draw.text((W - url_w - 40, H - 44), url, fill="#FFFFFF", font=font_small)

img.save(BASE / "og-image.png", "PNG", optimize=True)

# ---------------------------------------------------------------------------
# Favicon PNG 180x180 per iOS / apple-touch-icon
# ---------------------------------------------------------------------------
fav = Image.new("RGBA", (180, 180), (0, 0, 0, 0))
fd = ImageDraw.Draw(fav)
# sfondo blu arrotondato
fd.rounded_rectangle([0, 0, 180, 180], radius=38, fill="#1F77B4")
# due rettangoli bianchi sovrapposti
fd.rounded_rectangle([26, 36, 110, 150], radius=8, fill="#FFFFFF")
fd.rounded_rectangle([66, 50, 154, 160], radius=8, fill="#FFFFFF")
# righe colorate nel secondo
for i, c in enumerate(["#C6EFCE", "#FCE4D6", "#FFF2CC"]):
    y = 66 + i * 22
    fd.rectangle([78, y, 142, y + 12], fill=c)
# badge check verde
fd.ellipse([108, 100, 170, 162], fill="#2CA02C", outline="#FFFFFF", width=4)
fd.line([(122, 132), (134, 144), (156, 118)], fill="#FFFFFF", width=7)

fav.save(BASE / "apple-touch-icon.png", "PNG", optimize=True)

print("Creati:")
print(" -", BASE / "og-image.png")
print(" -", BASE / "apple-touch-icon.png")
