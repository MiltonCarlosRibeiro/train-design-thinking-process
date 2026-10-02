"""Gera recursos gráficos a partir do logo: logo recortado e ícone."""
from PIL import Image
from brand import LOGO_SRC, ASSETS, BG, rgb

ASSETS.mkdir(parents=True, exist_ok=True)
im = Image.open(LOGO_SRC).convert("RGB")
bg = rgb(BG)

def bbox_nonbg(img, tol=28):
    w, h = img.size
    px = img.load()
    xs, ys = [], []
    for y in range(h):
        for x in range(w):
            r, g, b = px[x, y]
            if max(abs(r - bg[0]), abs(g - bg[1]), abs(b - bg[2])) > tol:
                xs.append(x); ys.append(y)
    return min(xs), min(ys), max(xs) + 1, max(ys) + 1

x0, y0, x1, y1 = bbox_nonbg(im)
pad = 10
full = im.crop((max(0, x0 - pad), max(0, y0 - pad), min(im.width, x1 + pad), min(im.height, y1 + pad)))
full.save(ASSETS / "logo_full.png")

# ícone = parte superior (acima do texto "HayaiDataSystems")
icon = im.crop((150, 0, 470, 205))
ix0, iy0, ix1, iy1 = bbox_nonbg(icon)
icon = icon.crop((max(0, ix0 - 6), max(0, iy0 - 6), min(icon.width, ix1 + 6), min(icon.height, iy1 + 6)))
icon.save(ASSETS / "logo_icon.png")

# versão com fundo transparente (para PDFs com fundo claro/escuro)
def transparent(img, name):
    rgba = img.convert("RGBA")
    data = []
    for r, g, b, a in rgba.getdata():
        d = max(abs(r - bg[0]), abs(g - bg[1]), abs(b - bg[2]))
        alpha = 0 if d < 14 else min(255, int((d - 14) * 255 / 40)) if d < 54 else 255
        data.append((r, g, b, alpha))
    rgba.putdata(data)
    rgba.save(ASSETS / name)

transparent(full, "logo_full_t.png")
transparent(icon, "logo_icon_t.png")
print("full", full.size, "icon", icon.size)
