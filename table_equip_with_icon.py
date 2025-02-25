import pandas as pd
import os
import shutil
from PIL import Image

import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from library.excel_auto_fit import ExcelAutoFit
from table_equip import dump_weapon_data, get_weapon_types
from library.utils import rare_enum_to_value, reindex_column, seperate_enum_value

TEMP_DIR = "__temp"
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)


def compress_image(src_path: str, dst_path: str):
    try:
        img = Image.open(src_path)
        # 创建一个白色底色的新图像
        white_bg = Image.new("RGB", img.size, (255, 255, 255))
        # 将 PNG 图像粘贴到白色背景上
        white_bg.paste(img, mask=img.split()[3])  # 使用 alpha 通道作为掩码
        # 保存为 JPG 格式
        white_bg.save(dst_path, "JPEG", quality=90)
    except Exception as e:
        print(f"压缩图像时出错: {e}")


def apply_icons(
    workbook: openpyxl.Workbook,
    icon_dir: str,
    weapon_types: dict[str, int],
    skip_row: int = 0,
):
    for sheet_name in workbook.sheetnames:
        sheet = writer.book[sheet_name]

        # 遍历表头，查找Icon列的位置
        icon_col = -1
        for i, col in enumerate(sheet.iter_cols()):
            if col[skip_row].value == "Icon":
                icon_col = i
                break
        if icon_col == -1:
            print("Can't find Icon column")
            return
        # 调整列宽
        sheet.column_dimensions[get_column_letter(icon_col + 1)].width = 13
        # 遍历Icon列，读取文件名并插入图片
        for row in sheet.iter_rows(
            min_row=2 + skip_row, min_col=icon_col + 1, max_col=icon_col + 1
        ):
            for cell in row:
                icon_id = cell.value
                if not icon_id:
                    continue
                weapon_int_id = weapon_types[sheet_name]
                # example: tex_it0500_0039_IMLM4.tex.241106027.png
                icon_file_name = f"tex_it{weapon_int_id:02d}00_0{icon_id:03d}_IMLM4.tex.241106027.png"
                icon_path = os.path.join(
                    icon_dir,
                    icon_file_name,
                )
                compressed_path = os.path.join(TEMP_DIR, f"{icon_file_name[:-4]}.jpg")
                # 压缩图片
                compress_image(icon_path, compressed_path)
                print(f"Compressed at {compressed_path}")
                icon_path = compressed_path

                if os.path.exists(icon_path):
                    img = OpenpyxlImage(icon_path)
                    img.width = 100
                    img.height = 100
                    cell.value = ""
                    # 将图片插入到当前格
                    sheet.add_image(
                        img,
                        get_column_letter(icon_col + 1) + str(cell.row),
                    )
                    # 调整行高
                    sheet.row_dimensions[cell.row].height = 80
                else:
                    print(f"File not found: {icon_path}")


if __name__ == "__main__":
    print("Dumping weapon data...")
    weapon_sheets_serial = dump_weapon_data(keep_serial_id=True)
    weapon_sheets = dump_weapon_data(keep_serial_id=False)

    icon_dir = "C:/Users/Eigeen/Downloads/tex"

    with pd.ExcelWriter("WeaponDataWithIcon.xlsx") as writer:
        for weapon_type, weapon_data in weapon_sheets.items():
            # 提取Id enum value
            weapon_data_serial = weapon_sheets_serial[weapon_type]
            ids = []
            for _, row in weapon_data_serial.iterrows():
                int_id, _ = seperate_enum_value(row["Id"])
                ids.append(int_id)
            # 添加Icon列，内容为武器Id enum value
            weapon_data["Icon"] = ids
            weapon_data = reindex_column(weapon_data, "Icon", next_to="Name")

            weapon_data.to_excel(writer, sheet_name=weapon_type, index=False)

        print("Styling...")
        weapon_types = get_weapon_types()
        autofit = ExcelAutoFit()
        autofit.style_workbook(writer.book, max_width=60)

        print("Applying wrap text alignment...")
        align_wrap = Alignment(wrapText=True)
        for sheet_name in writer.sheets:
            sheet = writer.book[sheet_name]
            # RARE修正
            for row in range(1, sheet.max_row + 1):
                for col in range(1, sheet.max_column + 1):
                    cell = sheet.cell(row=row, column=col)
                    if isinstance(cell.value, str):
                        cell.value = rare_enum_to_value(cell.value)
            # 为所有非表头格应用自动换行
            for row in sheet.iter_rows(min_row=3):
                for cell in row:
                    cell.alignment = align_wrap
            # 首行增加一行，加签名
            sheet.insert_rows(1)
            sheet.cell(row=1, column=1).value = (
                "解包&制表：Eigeen本征  版本：1.0  不包含首日补丁"
            )

        print("Applying icons...")
        apply_icons(writer.book, icon_dir, weapon_types, skip_row=1)
