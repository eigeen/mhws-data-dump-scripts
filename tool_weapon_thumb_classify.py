import os
import shutil
import pandas as pd

from library.utils import seperate_enum_value
from table_equip import dump_weapon_data, dump_weapon_series_data, get_weapon_types

if __name__ == "__main__":
    weapon_types = get_weapon_types()
    weapon_data_serial = dump_weapon_data(keep_serial_id=True)

    # example: tex_it1100_0063_IMLM4.tex.241106027
    input_all_dir = "C:/Users/Eigeen/Downloads/tex"
    output_all_dir = "tex_classified"

    for weapon_type, data in weapon_data_serial.items():
        weapon_type_id_int = weapon_types[weapon_type]
        for _, row in data.iterrows():
            serial_id = row["Id"]
            int_id, enum_id = seperate_enum_value(serial_id)
            tex_name = (
                f"tex_it{weapon_type_id_int:02d}00_0{int_id:03d}_IMLM4.tex.241106027.png"
            )
            output_dir = os.path.join(output_all_dir, f"{weapon_type_id_int:02d}")
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            output_file_name = f"{enum_id}_{row["Name"]}.png"
            output_path = os.path.join(output_dir, output_file_name)
            input_path = f"{input_all_dir}/{tex_name}"
            try:
                shutil.copyfile(input_path, output_path)
            except FileNotFoundError:
                print(f"File not found: {input_path}")
