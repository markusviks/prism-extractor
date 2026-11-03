"""
Draws the Prism Extractor app icon and builds AppIcon.icns.

A slide with an empty dashed slot, and the graph that was in it lifting out above.
Run with:  .venv/bin/python icon/make_icon.py   (needs Pillow; macOS for iconutil)
"""

import random
import shutil
import subprocess
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

HERE = Path(__file__).parent
S = 2  # draw at 2x and downsample for smooth edges
W = 1024 * S


def px(*v):
    return [round(x * S) for x in v]


def shadow(base, box, radius, offset, blur, alpha):
    layer = Image.new("RGBA", base.size, (0, 0, 0, 0))
    x0, y0, x1, y1 = box
    ImageDraw.Draw(layer).rounded_rectangle(
        px(x0, y0 + offset, x1, y1 + offset), radius=radius * S, fill=(10, 15, 40, alpha))
    base.alpha_composite(layer.filter(ImageFilter.GaussianBlur(blur * S)))


def gradient(size, top, bottom):
    g = Image.new("RGBA", size)
    d = ImageDraw.Draw(g)
    w, h = size
    for y in range(h):
        t = y / (h - 1)
        d.line([(0, y), (w, y)], fill=tuple(round(a + (b - a) * t) for a, b in zip(top, bottom)) + (255,))
    return g


def dashed_rect(d, box, dash, gap, width, fill):
    x0, y0, x1, y1 = box
    for (ax, ay, bx, by) in [(x0, y0, x1, y0), (x1, y0, x1, y1), (x1, y1, x0, y1), (x0, y1, x0, y0)]:
        length = abs(bx - ax) + abs(by - ay)
        pos = 0
        while pos < length:
            end = min(pos + dash, length)
            t0, t1 = pos / length, end / length
            d.line(px(ax + (bx - ax) * t0, ay + (by - ay) * t0, ax + (bx - ax) * t1, ay + (by - ay) * t1),
                   fill=fill, width=width * S)
            pos += dash + gap


def draw():
    img = Image.new("RGBA", (W, W), (0, 0, 0, 0))

    # Background tile (macOS-style rounded square with a soft drop shadow)
    tile = (100, 100, 924, 924)
    shadow(img, tile, 185, 12, 18, 90)
    mask = Image.new("L", img.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle(px(*tile), radius=185 * S, fill=255)
    img.paste(gradient(img.size, (52, 92, 230), (24, 34, 110)), (0, 0), mask)

    d = ImageDraw.Draw(img)

    ink = (40, 48, 80, 255)
    slot_ink = (122, 140, 205, 255)

    # The slide, with an empty dashed slot (and an up arrow) where the graph used to be
    slide = (196, 548, 828, 862)
    shadow(img, slide, 26, 10, 14, 110)
    d.rounded_rectangle(px(*slide), radius=26 * S, fill=(236, 240, 252, 255))
    dashed_rect(d, (322, 640, 702, 822), 22, 14, 6, slot_ink)
    d.line(px(512, 790, 512, 690), fill=slot_ink, width=12 * S)
    d.polygon(px(472, 704, 512, 664, 552, 704), fill=slot_ink)

    # The graph card, lifted out of the slide
    card = (302, 168, 722, 548)
    shadow(img, card, 34, 26, 30, 150)
    d.rounded_rectangle(px(*card), radius=34 * S, fill=(255, 255, 255, 255))

    # Axes
    ax_x, ax_top, ax_y, ax_right = 362, 214, 492, 682
    d.line(px(ax_x, ax_top, ax_x, ax_y), fill=ink, width=8 * S)
    d.line(px(ax_x - 4, ax_y, ax_right, ax_y), fill=ink, width=8 * S)

    # Column scatter plot: dots, mean line and error bars per group
    rng = random.Random(3)
    groups = [(432, 404, (255, 122, 89)), (522, 330, (43, 179, 163)), (612, 270, (255, 184, 48))]
    for cx, mean, color in groups:
        for _ in range(8):
            x = cx + rng.uniform(-30, 30)
            y = mean + max(-52, min(52, rng.gauss(0, 30)))
            r = 11
            d.ellipse(px(x - r, y - r, x + r, y + r), fill=color + (255,))
        sd = 44
        d.line(px(cx, mean - sd, cx, mean + sd), fill=ink, width=6 * S)
        for yy in (mean - sd, mean + sd):
            d.line(px(cx - 14, yy, cx + 14, yy), fill=ink, width=6 * S)
        d.line(px(cx - 32, mean, cx + 32, mean), fill=ink, width=8 * S)

    return img.resize((1024, 1024), Image.LANCZOS)


def build_icns(master):
    iconset = HERE / "AppIcon.iconset"
    shutil.rmtree(iconset, ignore_errors=True)
    iconset.mkdir()
    for size in (16, 32, 128, 256, 512):
        master.resize((size, size), Image.LANCZOS).save(iconset / f"icon_{size}x{size}.png")
        master.resize((size * 2, size * 2), Image.LANCZOS).save(iconset / f"icon_{size}x{size}@2x.png")
    subprocess.run(["iconutil", "-c", "icns", str(iconset), "-o", str(HERE / "AppIcon.icns")], check=True)
    shutil.rmtree(iconset)


if __name__ == "__main__":
    master = draw()
    master.save(HERE / "AppIcon.png")
    build_icns(master)
    print("Wrote", HERE / "AppIcon.png", "and", HERE / "AppIcon.icns")
