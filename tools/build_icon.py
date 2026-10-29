"""메신저 작업 도우미 아이콘 — 쪽지 서식 도우미와 같은 계열(남색 둥근 사각형 + 흰 기호 + 청록 막대).
흰 말풍선 안에 남색 종이비행기 = 메신저 쪽지 자동 발송."""
from pathlib import Path
from PIL import Image, ImageDraw
ROOT = Path(__file__).resolve().parents[1]
NAVY, TEAL, WHITE = (2, 48, 100, 255), (50, 189, 184, 255), (255, 255, 255, 255)

def draw(size):
    s = 1024
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((40, 40, s - 40, s - 40), radius=200, fill=NAVY)
    # 말풍선
    d.rounded_rectangle((170, 170, 854, 690), radius=150, fill=WHITE)
    d.polygon([(300, 640), (420, 640), (250, 800)], fill=WHITE)
    # 종이비행기 (남색)
    d.polygon([(285, 440), (735, 265), (600, 600)], fill=NAVY)          # 몸통
    d.polygon([(285, 440), (735, 265), (455, 495)], fill=(40, 90, 160, 255))  # 날개 음영
    d.polygon([(455, 495), (735, 265), (505, 585)], fill=(14, 66, 128, 255))
    # 청록 막대 (계열 표시)
    d.rounded_rectangle((470, 800, s - 170, 880), radius=40, fill=TEAL)
    return img.resize((size, size), Image.Resampling.LANCZOS)

sizes = [16, 20, 24, 32, 40, 48, 64, 128, 256]
imgs = [draw(n) for n in sizes]
imgs[-1].save(ROOT / "assets" / "app-icon.ico", format="ICO", sizes=[(n, n) for n in sizes], append_images=imgs[:-1])
imgs[-1].save(ROOT / "assets" / "app-icon-preview-256.png")
draw(64).save(ROOT / "assets" / "app-icon-64.png")
prev = Image.new("RGBA", (420, 140), (240, 240, 240, 255))
x = 10
for n in (16, 24, 32, 48, 128):
    im = draw(n); prev.paste(im, (x, 70 - n // 2), im); x += n + 20
prev.save(ROOT / "assets" / "sizes-preview.png")
print("ok")
