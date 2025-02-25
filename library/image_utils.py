import os

from PIL import Image


TEMP_DIR = "__temp"


# 压缩图像，返回压缩后的图像路径
def compress_png(src_path: str, quality: int = 90) -> str:
    img = Image.open(src_path)
    # 创建一个白色底色的新图像
    white_bg = Image.new("RGB", img.size, (255, 255, 255))
    # 将 PNG 图像粘贴到白色背景上
    white_bg.paste(img, mask=img.split()[3])  # 使用 alpha 通道作为掩码
    # 保存为 JPG 格式
    src_file_name = os.path.basename(src_path)
    dst_file_name = src_file_name.replace(".png", ".jpg")
    dst_path = os.path.join(TEMP_DIR, dst_file_name)
    if not os.path.exists(os.path.dirname(dst_path)):
        os.makedirs(os.path.dirname(dst_path))

    white_bg.save(dst_path, "JPEG", quality=quality)
    return dst_path
