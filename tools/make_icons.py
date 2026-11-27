#!/usr/bin/env python3
"""Build PWA icons (glowing potsticker) from a real photo.

Usage: python3 tools/make_icons.py path/to/potsticker.jpg [left,top,size]   (optional square crop in pixels)
The photo should show a single potsticker, roughly centred. Without an argument a
placeholder drawing is generated (NOT a real photo - replace it before launch).
"""
import sys, math, random
from PIL import Image, ImageDraw, ImageFilter, ImageChops

BG_IN, BG_OUT = (74, 40, 20), (24, 13, 8)
GLOW = (255, 190, 90)

def placeholder(size=900):
    random.seed(3)
    im = Image.new("RGBA", (size, size), (0, 0, 0, 0)); d = ImageDraw.Draw(im)
    w, h = size * .86, size * .5; cx, cy = size / 2, size * .54
    d.ellipse([cx - w / 2, cy - h / 2 + size * .04, cx + w / 2, cy + h / 2 + size * .04], fill=(150, 80, 30, 255))
    d.ellipse([cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2], fill=(236, 190, 110, 255))
    for i in range(9):
        x = cx - w * .36 + i * w * .09
        d.arc([x - 20, cy - h * .42, x + 48, cy + h * .1], 200, 340, fill=(190, 130, 60, 255), width=7)
    for _ in range(40):
        x = cx + random.uniform(-w * .4, w * .4); y = cy + random.uniform(-h * .3, h * .3)
        d.ellipse([x, y, x + 9, y + 6], fill=(205, 150, 70, 255))
    return im

def cutout(im, size):
    im = im.convert("RGBA"); s = min(im.size)
    l, t = (im.width - s) // 2, (im.height - s) // 2
    im = im.crop((l, t, l + s, t + s)).resize((size, size), Image.LANCZOS)
    if im.getextrema()[3][0] == 255:            # photo without alpha -> soft round mask
        m = Image.new("L", (size, size), 0)
        ImageDraw.Draw(m).ellipse([size * .04, size * .04, size * .96, size * .96], fill=255)
        im.putalpha(m.filter(ImageFilter.GaussianBlur(size * .02)))
    return im

def build(photo, size, safe):
    bg = Image.new("RGB", (size, size)); px = bg.load()
    for y in range(size):
        for x in range(size):
            t = min(1, math.hypot(x - size / 2, y - size * .52) / (size * .72)); px[x, y] = tuple(int(a + (b - a) * t) for a, b in zip(BG_IN, BG_OUT))
    canvas = bg.convert("RGBA")
    inner = int(size * safe); item = cutout(photo, inner); off = ((size - inner) // 2,) * 2
    a = item.getchannel("A")
    for blur, strength in ((size * .09, .9), (size * .035, .7)):          # glow halo
        halo = Image.new("RGBA", (size, size), GLOW + (0,)); halo.putalpha(ImageChops.multiply(Image.new("L", (size, size), 0), Image.new("L", (size, size), 0)))
        m = Image.new("L", (size, size), 0); m.paste(a, off)
        m = m.filter(ImageFilter.GaussianBlur(blur)).point(lambda v: int(min(255, v * strength * 1.6)))
        canvas = Image.composite(Image.new("RGBA", (size, size), GLOW + (255,)), canvas, m)
    canvas.alpha_composite(item, off)
    return canvas.convert("RGB")

if __name__ == "__main__":
    photo = Image.open(sys.argv[1]) if len(sys.argv) > 1 else placeholder()
    if len(sys.argv) > 2:
        l, t, n = map(int, sys.argv[2].split(",")); photo = photo.convert("RGB").crop((l, t, l + n, t + n))
    for name, size, safe in (("icon-192.png", 192, .72), ("icon-512.png", 512, .72), ("icon-maskable-512.png", 512, .56), ("apple-touch-icon.png", 180, .72)):
        build(photo, size, safe).save(f"icons/{name}", optimize=True); print("wrote icons/" + name)
