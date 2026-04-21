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

# ---------------------------------------------------------------------------
# Icone PWA: 192x192 e 512x512 standard + 512x512 maskable (safe zone)
# ---------------------------------------------------------------------------
def make_icon(size, maskable=False):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    # sfondo arrotondato blu. Per maskable -> safe zone: il contenuto vero
    # sta nel 80% centrale, quindi il badge diventa un quadrato pieno
    if maskable:
        # Android taglia in cerchio/squircle, quindi sfondo pieno (no bordi arrotondati)
        d.rectangle([0, 0, size, size], fill="#1F77B4")
        safe_margin = int(size * 0.12)  # area sicura 76% al centro
    else:
        r = int(size * 0.22)
        d.rounded_rectangle([0, 0, size, size], radius=r, fill="#1F77B4")
        safe_margin = int(size * 0.14)

    s = size - 2 * safe_margin
    # proporzioni replicate dall'apple-touch (180x180 origine)
    def sx(px): return safe_margin + int(px / 180 * s)
    def sy(px): return safe_margin + int(px / 180 * s)

    # foglio 1
    d.rounded_rectangle([sx(26), sy(36), sx(110), sy(150)],
                        radius=max(4, int(size*0.045)), fill="#FFFFFF")
    # foglio 2
    d.rounded_rectangle([sx(66), sy(50), sx(154), sy(160)],
                        radius=max(4, int(size*0.045)), fill="#FFFFFF")
    for i, c in enumerate(["#C6EFCE", "#FCE4D6", "#FFF2CC"]):
        y = 66 + i * 22
        d.rectangle([sx(78), sy(y), sx(142), sy(y + 12)], fill=c)
    # badge check verde
    d.ellipse([sx(108), sy(100), sx(170), sy(162)],
              fill="#2CA02C", outline="#FFFFFF", width=max(2, size//45))
    # check
    d.line([(sx(122), sy(132)), (sx(134), sy(144)), (sx(156), sy(118))],
           fill="#FFFFFF", width=max(3, size//26))

    # per i non-maskable convertiamo in RGB mantenendo il background trasparente
    # ma in realtà il PNG RGBA va bene uguale
    return img

make_icon(192).save(BASE / "icon-192.png", "PNG", optimize=True)
make_icon(512).save(BASE / "icon-512.png", "PNG", optimize=True)
make_icon(512, maskable=True).save(BASE / "icon-512-maskable.png", "PNG", optimize=True)

print("Creati:")
print(" -", BASE / "og-image.png")
print(" -", BASE / "apple-touch-icon.png")
print(" -", BASE / "icon-192.png")
print(" -", BASE / "icon-512.png")
print(" -", BASE / "icon-512-maskable.png")
