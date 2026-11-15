"""Tạo icon ứng dụng assets/icon.ico (không cần file ảnh nguồn)."""
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

S = 1024          # vẽ ở kích thước lớn rồi thu nhỏ để mượt
ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "assets" / "icon.ico"
FONT_BOLD = ROOT / "fonts" / "DejaVuSans-Bold.ttf"

BLUE = (30, 111, 217, 255)
BLUE_DARK = (21, 84, 168, 255)
RED = (219, 68, 55, 255)
LINE = (108, 124, 145, 255)
PAPER = (255, 255, 255, 255)


def make() -> Image.Image:
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))

    # ── nền: bo góc, có gradient dọc nhẹ ────────────────────────────────
    bg = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    bd = ImageDraw.Draw(bg)
    for y in range(S):
        t = y / S
        c = tuple(int(BLUE[i] * (1 - t) + BLUE_DARK[i] * t) for i in range(3))
        bd.line([(0, y), (S, y)], fill=c + (255,))
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, S - 1, S - 1), radius=int(S * 0.215), fill=255)
    img.paste(bg, (0, 0), mask)

    # vệt sáng trên (được mask trong nền nên không tràn ra ngoài)
    glow = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(glow).ellipse((-S * 0.45, -S * 0.85, S * 0.85, S * 0.35),
                                 fill=(255, 255, 255, 30))
    img.paste(glow, (0, 0), Image.composite(glow, Image.new("RGBA", (S, S), (0, 0, 0, 0)), mask))

    d = ImageDraw.Draw(img)

    # ── tờ giấy trắng, góc phải trên gập ────────────────────────────────
    left, top, right, bottom = S * 0.245, S * 0.135, S * 0.755, S * 0.865
    fold = S * 0.15
    d.polygon([(left, top), (right - fold, top), (right, top + fold),
               (right, bottom), (left, bottom)], fill=PAPER)
    d.polygon([(right - fold, top), (right - fold, top + fold), (right, top + fold)],
              fill=(206, 220, 239, 255))

    # ── 3 dòng chữ ──────────────────────────────────────────────────────
    x0, x1 = left + S * 0.07, right - S * 0.07
    y = top + S * 0.16
    for frac in (1.0, 0.82, 0.58):
        d.rounded_rectangle((x0, y, x0 + (x1 - x0) * frac, y + S * 0.05),
                            radius=S * 0.025, fill=LINE)
        y += S * 0.115

    # ── dải đỏ "PDF" ở chân tờ giấy ─────────────────────────────────────
    band_h = S * 0.145
    band_top = bottom - band_h
    d.rounded_rectangle((left + S * 0.07, band_top, right - S * 0.07, bottom - S * 0.055),
                        radius=S * 0.035, fill=RED)
    try:
        f = ImageFont.truetype(str(FONT_BOLD), int(S * 0.105))
        text = "PDF"
        bb = d.textbbox((0, 0), text, font=f)
        d.text(((left + right) / 2 - (bb[2] - bb[0]) / 2 - bb[0],
                (band_top + bottom - S * 0.055) / 2 - (bb[3] - bb[1]) / 2 - bb[1]),
               text, font=f, fill=PAPER)
    except Exception:
        pass

    return img


def main() -> Path:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    big = make()
    big.resize((256, 256), Image.LANCZOS).save(
        OUT, format="ICO",
        sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)],
    )
    for px in (256, 64, 32, 16):
        big.resize((px, px), Image.LANCZOS).save(OUT.parent / f"icon_{px}.png")
    return OUT


if __name__ == "__main__":
    print("Đã tạo:", main())
