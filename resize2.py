from pdf2image import convert_from_path
from PIL import Image

# 変換設定
pdf_path = r"C:\Users\woo.jihun\Desktop\Skywalker\sample.pdf"
dpi = 300

# 用紙サイズ（mm → px）
paper_width_mm = 24.9
paper_height_mm = 22.4
paper_width_px = int(paper_width_mm * dpi / 25.4)
paper_height_px = int(paper_height_mm * dpi / 25.4)

# PDF → 画像変換（1ページ目だけ処理）
images = convert_from_path(pdf_path, dpi=dpi)
image = images[0]

# 元サイズ取得
original_width, original_height = image.size

# 上部を固定（例：上部50%を固定）
fixed_ratio = 0.5
fixed_height = int(original_height * fixed_ratio)
shrink_target_height = paper_height_px

# 残り部分（下部）を縮小
lower_part = image.crop((0, fixed_height, original_width, original_height))
lower_part_height = original_height - fixed_height
scaling_ratio = (shrink_target_height - fixed_height) / lower_part_height

# 下部リサイズ
resized_lower = lower_part.resize(
    (original_width, int(lower_part_height * scaling_ratio)),
    resample=Image.LANCZOS
)

# 上部切り出し（そのまま）
upper_part = image.crop((0, 0, original_width, fixed_height))

# 上下結合
new_height = fixed_height + resized_lower.height
final_image = Image.new("RGB", (original_width, new_height))
final_image.paste(upper_part, (0, 0))
final_image.paste(resized_lower, (0, fixed_height))

# サイズが用紙サイズを超えていれば最終リサイズ
final_image = final_image.resize((paper_width_px, paper_height_px), Image.LANCZOS)

# 保存
final_image.save("resized_sample.png")
final_image.show()
