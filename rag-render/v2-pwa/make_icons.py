"""產生 PWA 圖示（需要 Pillow）：python make_icons.py"""
from PIL import Image, ImageDraw

G, W = (11, 138, 100), (255, 255, 255)


def draw(size, safe):  # safe：內容佔畫布比例（maskable 需留安全區）
    S = 1024
    img = Image.new("RGB", (S, S), G)
    d = ImageDraw.Draw(img)
    w, h = S * safe * 0.56, S * safe * 0.72
    x0, y0 = (S - w) / 2, (S - h) / 2
    fold = w * 0.28
    d.polygon([(x0, y0), (x0 + w - fold, y0), (x0 + w, y0 + fold), (x0 + w, y0 + h), (x0, y0 + h)], fill=W)
    d.polygon([(x0 + w - fold, y0), (x0 + w - fold, y0 + fold), (x0 + w, y0 + fold)], fill=(190, 226, 212))
    lw = w * 0.56
    for i, f in enumerate((1, 0.85, 1, 0.6)):
        y = y0 + h * (0.42 + i * 0.13)
        d.rounded_rectangle([x0 + w * 0.2, y, x0 + w * 0.2 + lw * f, y + h * 0.045], radius=h * 0.022, fill=G)
    return img.resize((size, size), Image.LANCZOS)


if __name__ == "__main__":
    draw(192, 0.9).save("static/icons/icon-192.png")
    draw(512, 0.9).save("static/icons/icon-512.png")
    draw(512, 0.7).save("static/icons/icon-maskable-512.png")
    draw(180, 0.9).save("static/icons/apple-touch-icon.png")
