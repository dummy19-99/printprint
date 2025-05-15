from pdf2image import convert_from_path
from PIL import Image

# ▼ 設定値
pdf_path = r"C:\Users\woo.jihun\Desktop\Skywalker\sample.pdf"  # ← PDFファイルのパス
dpi = 300
poppler_path = r"C:\tools\poppler\Library\bin"  # ← Popplerを解凍した場所に合わせて修正

# ▼ 用紙サイズ（mm → px）
paper_width_mm = 24.9
paper_height_mm = 22.4
paper_width_px = int(paper_width_mm * dpi / 25.4)
paper_height_px = int(paper_height_mm * dpi / 25.4)

# ▼ PDF → 画像（1ページ目）
images = convert_from_path(pdf_path, dpi=dpi, poppler_path=poppler_path)
image = images[0]
original_width, original_height = image.size

# ▼ 上部を固定し、必要な高さだけ切り出し
crop_height = min(paper_height_px, original_height)
cropped_image = image.crop((0, 0, original_width, crop_height))

# ▼ 横幅が違う場合は調整（縦はそのまま）
cropped_image = cropped_image.resize((paper_width_px, crop_height), Image.LANCZOS)

# ▼ 用紙サイズにパディング（下に余白）
final_image = Image.new("RGB", (paper_width_px, paper_height_px), "white")
final_image.paste(cropped_image, (0, 0))  # 上から配置

# ▼ 保存・表示（テスト用）
final_image.save("cropped_sample.png")
final_image.show()
