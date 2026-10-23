#!/usr/bin/env python3
"""يولّد أيقونات وشعار التطبيق من لوجو المكتب (logo-original.jpg) بشكل أوضح وأكتر عمقاً.
الاستخدام: python3 make_icons.py [logo.jpg]"""
import sys, os
from PIL import Image, ImageDraw, ImageFilter, ImageChops, ImageEnhance

S = 1024
src = sys.argv[1] if len(sys.argv) > 1 else "logo-original.jpg"

def enhanced(path):
    im = Image.open(path).convert("RGB")
    im = im.resize((S, S), Image.LANCZOS)
    im = im.filter(ImageFilter.UnsharpMask(radius=3, percent=140, threshold=2))
    im = ImageEnhance.Contrast(im).enhance(1.12)
    im = ImageEnhance.Color(im).enhance(1.25)
    im = ImageEnhance.Brightness(im).enhance(1.05)
    return im

def disc(im, d):
    """قص دايرة قطرها d مع حواف ناعمة"""
    c = im.resize((d, d), Image.LANCZOS).convert("RGBA")
    m = Image.new("L", (d * 4, d * 4), 0)
    ImageDraw.Draw(m).ellipse((0, 0, d * 4 - 1, d * 4 - 1), fill=255)
    c.putalpha(m.resize((d, d), Image.LANCZOS))
    return c

def bg_gradient(rounded):
    g = Image.new("RGB", (1, S))
    for y in range(S):
        t = y / (S - 1)
        g.putpixel((0, y), (int(30 - 22 * t), int(70 - 55 * t), int(120 - 95 * t)))
    b = g.resize((S, S)).convert("RGBA")
    if rounded:
        m = Image.new("L", (S, S), 0)
        ImageDraw.Draw(m).rounded_rectangle((0, 0, S - 1, S - 1), radius=int(S * .22), fill=255)
        o = Image.new("RGBA", (S, S), (0, 0, 0, 0)); o.paste(b, (0, 0), m); return o
    return b

def compose(im, frac, rounded):
    d = int(S * frac)
    base = bg_gradient(rounded)
    x = (S - d) // 2
    d_img = disc(im, d)
    # ظل تحت الدايرة
    sh = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(sh).ellipse((x, x + int(S * .035), x + d, x + d + int(S * .035)), fill=(0, 0, 0, 190))
    base = Image.alpha_composite(base, sh.filter(ImageFilter.GaussianBlur(S * .025)))
    layer = Image.new("RGBA", (S, S), (0, 0, 0, 0)); layer.paste(d_img, (x, x))
    base = Image.alpha_composite(base, layer)
    # حلقة ذهبية رفيعة + لمعة زجاجية من فوق للعمق
    ring = Image.new("RGBA", (S * 2, S * 2), (0, 0, 0, 0))
    rd = ImageDraw.Draw(ring)
    rd.ellipse((x * 2, x * 2, (x + d) * 2, (x + d) * 2), outline=(255, 205, 80, 255), width=int(S * .012))
    base = Image.alpha_composite(base, ring.resize((S, S), Image.LANCZOS))
    gl = Image.new("L", (S * 2, S * 2), 0)
    ImageDraw.Draw(gl).ellipse(((x + d * .08) * 2, (x + d * .03) * 2, (x + d * .92) * 2, (x + d * .5) * 2), fill=46)
    cm = Image.new("L", (S * 2, S * 2), 0)
    ImageDraw.Draw(cm).ellipse((x * 2, x * 2, (x + d) * 2, (x + d) * 2), fill=255)
    gl = ImageChops.multiply(gl, cm).resize((S, S), Image.LANCZOS).filter(ImageFilter.GaussianBlur(10))
    white = Image.new("RGBA", (S, S), (255, 255, 255, 0)); white.putalpha(gl)
    return Image.alpha_composite(base, white)

os.makedirs("icons", exist_ok=True)
im = enhanced(src)
a = compose(im, .90, True)
a.resize((512, 512), Image.LANCZOS).save("icons/icon-512.png")
a.resize((192, 192), Image.LANCZOS).save("icons/icon-192.png")
compose(im, .90, False).convert("RGB").resize((180, 180), Image.LANCZOS).save("icons/apple-touch-icon.png")
compose(im, .70, False).convert("RGB").resize((512, 512), Image.LANCZOS).save("icons/icon-maskable-512.png")
disc(im, S).save("icons/logo.png")          # الشعار بخلفية شفافة (دايرة)
a.save("icons/icon-1024.png")
print("تم")
