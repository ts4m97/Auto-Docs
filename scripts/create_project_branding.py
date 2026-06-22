from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "assets"
DOCS_ASSETS = ROOT / "docs" / "assets"


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    candidates = [
        Path("C:/Windows/Fonts/segoeuib.ttf" if bold else "C:/Windows/Fonts/segoeui.ttf"),
        Path("C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf"),
    ]
    for path in candidates:
        if path.exists():
            return ImageFont.truetype(str(path), size)
    return ImageFont.load_default()


def rounded_rectangle(draw: ImageDraw.ImageDraw, xy, radius, fill, outline=None, width=1):
    draw.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline, width=width)


def draw_text(draw: ImageDraw.ImageDraw, xy, text: str, size: int, fill, bold: bool = False):
    draw.text(xy, text, font=font(size, bold), fill=fill)


def create_banner() -> None:
    width, height = 1280, 640
    image = Image.new("RGB", (width, height), "#f6f7fb")
    draw = ImageDraw.Draw(image)

    for y in range(height):
        t = y / height
        r = int(246 * (1 - t) + 229 * t)
        g = int(247 * (1 - t) + 236 * t)
        b = int(251 * (1 - t) + 246 * t)
        draw.line([(0, y), (width, y)], fill=(r, g, b))

    icon_path = ASSETS / "autodocs-icon.png"
    icon = Image.open(icon_path).convert("RGBA").resize((168, 168), Image.Resampling.LANCZOS)
    image.paste(icon, (88, 86), icon)

    draw_text(draw, (288, 98), "Auto Docs", 66, "#102a43", True)
    draw_text(draw, (292, 178), "Offline Word template automation for Windows", 30, "#334155")
    draw_text(draw, (292, 224), "Fill DOCX placeholders from forms or Excel, then export DOCX, PDF, PNG and print in batches.", 22, "#475569")

    chips = ["PySide6", "DOCX", "Excel", "PDF", "Batch Print", "Offline"]
    x = 292
    for chip in chips:
        text_w = int(draw.textlength(chip, font=font(18, True)))
        rounded_rectangle(draw, [x, 282, x + text_w + 30, 324], 21, "#ffffff", "#cbd5e1", 2)
        draw_text(draw, (x + 15, 291), chip, 18, "#1d4ed8", True)
        x += text_w + 42

    panel = [84, 374, 1196, 586]
    rounded_rectangle(draw, panel, 18, "#ffffff", "#d8dee9", 2)
    draw_text(draw, (118, 402), "[[customer_name]]", 24, "#2563eb", True)
    draw_text(draw, (118, 444), "[[contract_no]]", 24, "#2563eb", True)
    draw_text(draw, (118, 486), "[[effective_date]]", 24, "#2563eb", True)

    arrow_y = 462
    draw.line([(420, arrow_y), (516, arrow_y)], fill="#94a3b8", width=5)
    draw.polygon([(516, arrow_y), (496, arrow_y - 14), (496, arrow_y + 14)], fill="#94a3b8")

    rounded_rectangle(draw, [560, 398, 1160, 552], 14, "#eff6ff", "#bfdbfe", 2)
    draw_text(draw, (592, 418), "Nguyen Van A", 28, "#172033", True)
    draw_text(draw, (592, 466), "HD-2026-0001", 28, "#172033", True)
    draw_text(draw, (592, 514), "22/06/2026", 28, "#172033", True)

    DOCS_ASSETS.mkdir(parents=True, exist_ok=True)
    image.save(DOCS_ASSETS / "hero.png", quality=95)


def create_ui_preview() -> None:
    width, height = 1280, 760
    image = Image.new("RGB", (width, height), "#eef2f7")
    draw = ImageDraw.Draw(image)

    rounded_rectangle(draw, [42, 40, 1238, 720], 20, "#f6f7fb", "#cbd5e1", 2)
    rounded_rectangle(draw, [42, 40, 340, 720], 20, "#ffffff", "#e2e8f0", 2)
    draw.rectangle([320, 42, 340, 718], fill="#ffffff")

    icon = Image.open(ASSETS / "autodocs-icon.png").convert("RGBA").resize((58, 58), Image.Resampling.LANCZOS)
    image.paste(icon, (72, 72), icon)
    draw_text(draw, (144, 78), "Auto Docs", 30, "#102a43", True)
    draw_text(draw, (144, 116), "Offline document forms", 17, "#64748b")

    rounded_rectangle(draw, [72, 166, 308, 210], 9, "#ffffff", "#cbd5e1", 2)
    draw_text(draw, (90, 177), "Tim mau...", 17, "#64748b")

    templates = [
        ("Hop dong mua ban", "Hop dong - 23 truong"),
        ("Bao gia san pham", "Kinh doanh - 30 truong"),
        ("Quyet dinh nhan su", "Nhan su - 14 truong"),
    ]
    y = 238
    for index, (name, meta) in enumerate(templates):
        fill = "#eaf1ff" if index == 2 else "#ffffff"
        outline = "#93b4f2" if index == 2 else "#e2e8f0"
        rounded_rectangle(draw, [72, y, 308, y + 74], 10, fill, outline, 2)
        draw_text(draw, (90, y + 13), name, 18, "#172033", True)
        draw_text(draw, (90, y + 41), meta, 15, "#64748b")
        y += 86

    rounded_rectangle(draw, [74, 620, 306, 666], 9, "#2563eb", "#2563eb", 2)
    draw_text(draw, (118, 632), "Them mau Word", 17, "#ffffff", True)

    draw_text(draw, (382, 78), "Quyet dinh nhan su", 34, "#102a43", True)
    draw_text(draw, (382, 122), "Nhan su - 14 placeholder - 04_quyet_dinh_nhan_su.docx", 18, "#64748b")

    rounded_rectangle(draw, [382, 166, 1198, 282], 10, "#ffffff", "#e2e8f0", 2)
    draw_text(draw, (408, 190), "Placeholder trong mau", 20, "#172033", True)
    columns = ["[[ho_ten_nhan_su]]", "[[ma_nhan_vien]]", "[[ngay_hieu_luc]]", "[[phong_ban]]"]
    x = 408
    for col in columns:
        rounded_rectangle(draw, [x, 226, x + 180, 258], 6, "#f8fafc", "#e2e8f0", 1)
        draw_text(draw, (x + 10, 233), col, 14, "#2563eb", True)
        x += 192

    tabs = [("Nhap thu cong", True), ("Nhap tu Excel", False), ("In hang loat", False), ("Lich su", False)]
    x = 382
    for label, active in tabs:
        fill = "#ffffff" if active else "#e9eef6"
        color = "#1d4ed8" if active else "#475467"
        rounded_rectangle(draw, [x, 312, x + 150, 354], 8, fill, "#e2e8f0", 1)
        draw_text(draw, (x + 18, 323), label, 16, color, True)
        x += 160

    rounded_rectangle(draw, [382, 374, 812, 650], 10, "#ffffff", "#e2e8f0", 2)
    draw_text(draw, (408, 398), "Thong tin can dien", 20, "#172033", True)
    fields = ["Ho ten nhan su", "Ma nhan vien", "Chuc danh moi", "Ngay hieu luc"]
    y = 444
    for field in fields:
        draw_text(draw, (410, y + 10), field, 16, "#334155", True)
        rounded_rectangle(draw, [570, y, 782, y + 38], 8, "#ffffff", "#cbd5e1", 1)
        y += 52

    rounded_rectangle(draw, [840, 374, 1198, 650], 10, "#ffffff", "#e2e8f0", 2)
    draw_text(draw, (866, 398), "Quy cach xuat file", 20, "#172033", True)
    rounded_rectangle(draw, [866, 444, 1172, 486], 8, "#ffffff", "#cbd5e1", 1)
    draw_text(draw, (884, 455), "[[so_quyet_dinh]] - [[ho_ten_nhan_su]]", 15, "#64748b")
    rounded_rectangle(draw, [866, 526, 1008, 572], 8, "#2563eb", "#2563eb", 2)
    draw_text(draw, (902, 539), "Xuat DOCX", 17, "#ffffff", True)
    rounded_rectangle(draw, [1024, 526, 1138, 572], 8, "#ffffff", "#cbd5e1", 2)
    draw_text(draw, (1052, 539), "PDF", 17, "#172033", True)

    DOCS_ASSETS.mkdir(parents=True, exist_ok=True)
    image.save(DOCS_ASSETS / "app-preview.png", quality=95)


def main() -> None:
    create_banner()
    create_ui_preview()
    print(f"Wrote project branding assets to {DOCS_ASSETS}")


if __name__ == "__main__":
    main()
