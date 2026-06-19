from PIL import Image
from pathlib import Path

src = Path("logo.png")
dst = Path("logo.ico")

img = Image.open(src).convert("RGBA")

img.save(
    dst,
    format="ICO",
    sizes=[
        (16, 16),
        (24, 24),
        (32, 32),
        (48, 48),
        (64, 64),
        (128, 128),
        (256, 256),
    ],
)

print(f"Created: {dst.resolve()}")