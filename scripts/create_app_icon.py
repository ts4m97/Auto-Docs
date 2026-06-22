from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter


ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "assets"
SIZES = [16, 24, 32, 48, 64, 128, 256]


def rounded_rectangle(draw: ImageDraw.ImageDraw, xy, radius, fill, outline=None, width=1):
    draw.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline, width=width)


def create_icon(size: int) -> Image.Image:
    scale = size / 256
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))

    shadow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    shadow_draw = ImageDraw.Draw(shadow)
    rounded_rectangle(
        shadow_draw,
        [int(42 * scale), int(24 * scale), int(210 * scale), int(232 * scale)],
        int(24 * scale),
        (15, 23, 42, 72),
    )
    shadow = shadow.filter(ImageFilter.GaussianBlur(max(1, int(8 * scale))))
    image.alpha_composite(shadow)

    draw = ImageDraw.Draw(image)
    rounded_rectangle(
        draw,
        [int(34 * scale), int(18 * scale), int(206 * scale), int(226 * scale)],
        int(24 * scale),
        (255, 255, 255, 255),
        (203, 213, 225, 255),
        max(1, int(3 * scale)),
    )

    fold = [
        (int(158 * scale), int(18 * scale)),
        (int(206 * scale), int(66 * scale)),
        (int(158 * scale), int(66 * scale)),
    ]
    draw.polygon(fold, fill=(219, 234, 254, 255), outline=(147, 180, 242, 255))
    draw.line(
        [(int(158 * scale), int(18 * scale)), (int(158 * scale), int(66 * scale))],
        fill=(147, 180, 242, 255),
        width=max(1, int(2 * scale)),
    )

    accent = (37, 99, 235, 255)
    dark = (23, 32, 51, 255)
    muted = (100, 116, 139, 255)
    green = (16, 185, 129, 255)

    if size >= 32:
        line_y = [90, 118, 146]
        for y in line_y:
            draw.rounded_rectangle(
                [int(64 * scale), int(y * scale), int(176 * scale), int((y + 10) * scale)],
                radius=max(1, int(5 * scale)),
                fill=(226, 232, 240, 255),
            )

    bracket_width = max(2, int(9 * scale))
    bracket_top = int(92 * scale)
    bracket_bottom = int(166 * scale)
    left_x = int(70 * scale)
    right_x = int(168 * scale)
    hook = int(22 * scale)
    draw.line([(left_x + hook, bracket_top), (left_x, bracket_top), (left_x, bracket_bottom), (left_x + hook, bracket_bottom)], fill=accent, width=bracket_width)
    draw.line([(right_x - hook, bracket_top), (right_x, bracket_top), (right_x, bracket_bottom), (right_x - hook, bracket_bottom)], fill=accent, width=bracket_width)

    dot_radius = max(2, int(7 * scale))
    for x in [102, 119, 136]:
        draw.ellipse(
            [
                int(x * scale) - dot_radius,
                int(129 * scale) - dot_radius,
                int(x * scale) + dot_radius,
                int(129 * scale) + dot_radius,
            ],
            fill=dark,
        )

    badge_r = int(34 * scale)
    badge_c = (int(178 * scale), int(194 * scale))
    draw.ellipse(
        [badge_c[0] - badge_r, badge_c[1] - badge_r, badge_c[0] + badge_r, badge_c[1] + badge_r],
        fill=green,
        outline=(255, 255, 255, 255),
        width=max(1, int(5 * scale)),
    )
    tick = [
        (int(160 * scale), int(194 * scale)),
        (int(174 * scale), int(207 * scale)),
        (int(199 * scale), int(178 * scale)),
    ]
    draw.line(tick, fill=(255, 255, 255, 255), width=max(2, int(9 * scale)), joint="curve")

    if size < 32:
        compact = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(compact)
        rounded_rectangle(draw, [2, 2, size - 3, size - 3], max(3, size // 6), (37, 99, 235, 255))
        draw.line([(size // 3, size // 3), (size // 4, size // 3), (size // 4, size * 2 // 3), (size // 3, size * 2 // 3)], fill=(255, 255, 255, 255), width=max(1, size // 10))
        draw.line([(size * 2 // 3, size // 3), (size * 3 // 4, size // 3), (size * 3 // 4, size * 2 // 3), (size * 2 // 3, size * 2 // 3)], fill=(255, 255, 255, 255), width=max(1, size // 10))
        return compact

    return image


def main() -> None:
    ASSETS.mkdir(parents=True, exist_ok=True)
    png = create_icon(1024)
    png.save(ASSETS / "autodocs-icon.png")
    images = [create_icon(size) for size in SIZES]
    images[-1].save(
        ASSETS / "autodocs.ico",
        sizes=[(size, size) for size in SIZES],
        append_images=images[:-1],
    )
    print(f"Wrote {ASSETS / 'autodocs.ico'}")


if __name__ == "__main__":
    main()
