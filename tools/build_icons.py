"""Draw the app logo («کارمند و رشد»: employee bust with rising golden bars) and write every icon file.

Output:
  personnel_evaluation_report/static/logo.svg      – UI badge and browser-tab icon
  personnel_evaluation_report/static/favicon.png   – PNG fallback for the browser tab
  packaging/app.ico                                 – Windows exe + installer icon
  packaging/app.icns                                – macOS app icon
  packaging/logo-1024.png                           – master raster

Run:  .venv/bin/python tools/build_icons.py
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent.parent
TOP, BOTTOM = (0x4B, 0x4B, 0xB0), (0x2E, 0x2E, 0x78)   # indigo gradient (app accent)
GOLD = (0xF2, 0xB3, 0x3D)
WHITE = (255, 255, 255)

# Geometry on a 512 × 512 canvas — shared by the SVG and the raster drawing.
TILE = (16, 16, 496, 496, 112)                 # x0, y0, x1, y1, corner radius
HEAD = (256, 190, 70)                          # cx, cy, r
BODY = (130, 272, 382, 528)                    # ellipse box; its upper half is the shoulders
BARS = [(300, 330, 334, 400), (344, 296, 378, 400), (388, 256, 422, 400)]
BAR_R = 8

SVG = f"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512">
  <defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
    <stop offset="0" stop-color="#4b4bb0"/><stop offset="1" stop-color="#2e2e78"/></linearGradient></defs>
  <rect x="16" y="16" width="480" height="480" rx="112" fill="url(#g)"/>
  <circle cx="256" cy="190" r="70" fill="#fff"/>
  <path d="M130 400 A126 128 0 0 1 382 400 Z" fill="#fff"/>
  <rect x="300" y="330" width="34" height="70" rx="8" fill="#f2b33d"/>
  <rect x="344" y="296" width="34" height="104" rx="8" fill="#f2b33d"/>
  <rect x="388" y="256" width="34" height="144" rx="8" fill="#f2b33d"/>
</svg>
"""


def draw(size: int) -> Image.Image:
    """Render at 4× and downsample for smooth edges."""
    ss = size * 4
    k = ss / 512
    s = lambda *v: [round(x * k) for x in v]   # noqa: E731
    # diagonal gradient, clipped to the rounded tile
    grad = Image.new("RGB", (ss, ss))
    gp = grad.load()
    for y in range(ss):
        for x in range(0, ss):
            t = (x + y) / (2 * ss)
            gp[x, y] = tuple(round(a + (b - a) * t) for a, b in zip(TOP, BOTTOM))
    mask = Image.new("L", (ss, ss), 0)
    x0, y0, x1, y1, r = TILE
    ImageDraw.Draw(mask).rounded_rectangle(s(x0, y0, x1, y1), radius=round(r * k), fill=255)
    img = Image.new("RGBA", (ss, ss), (0, 0, 0, 0))
    img.paste(grad, (0, 0), mask)
    d = ImageDraw.Draw(img)
    cx, cy, hr = HEAD
    d.ellipse(s(cx - hr, cy - hr, cx + hr, cy + hr), fill=WHITE)
    d.pieslice(s(*BODY), 180, 360, fill=WHITE)
    for bx0, by0, bx1, by1 in BARS:
        d.rounded_rectangle(s(bx0, by0, bx1, by1), radius=round(BAR_R * k), fill=GOLD)
    return img.resize((size, size), Image.Resampling.LANCZOS)


def main() -> None:
    static = ROOT / "personnel_evaluation_report" / "static"
    pack = ROOT / "packaging"
    (static / "logo.svg").write_text(SVG, encoding="utf-8")
    master = draw(1024)
    master.save(pack / "logo-1024.png")
    draw(64).save(static / "favicon.png")
    master.save(pack / "app.ico", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    master.save(pack / "app.icns")
    print("icons written")


if __name__ == "__main__":
    main()
