import json
import pandas as pd

from library.excel_auto_fit import ExcelAutoFit
from library.item_db import ItemDB
from library.text_db import load_text_db
from library.utils import minify_nested_serial, reindex_column, remove_enum_value
from parse_whistle_tone import ToneParser
from table_equip import (
    dump_weapon_data,
    get_gun_shell_type_name,
    get_slash_axe_bin_name,
    process_loading_bin,
)
from table_general import dump_user3_data_general, load_enum_internal

item_db = ItemDB("item_db.json")
text_db = load_text_db("texts_db.json")

# 火、水、电、冰、龙、毒、麻、眠、NONE、榴弹
SHELL_INDEX_TO_NAME = {
    0: "FIRE",
    1: "WATER",
    2: "ELEC",
    3: "ICE",
    4: "DRAGON",
    5: "POISON",
    6: "PARALYSE",
    7: "SLEEP",
    8: "Unknown",
    9: "GRENADE",
}


def get_item_name_mapping(x):
    item = item_db.get_entry_by_id(str(x))
    if item is None:
        return ""
    return item.raw_name


def process_artian_gun_shell(shell_lv_list: list[str]) -> list[str]:
    if len(shell_lv_list) != 10:
        raise ValueError("Artian Shell list length should be 20")
    shell_names = []
    for idx, shell_lv in enumerate(shell_lv_list):
        if shell_lv == "NONE":
            shell_names.append(None)
            continue
        shell_type = SHELL_INDEX_TO_NAME.get(idx)
        if shell_type is None:
            print(f"Unknown shell type: {shell_lv}")
            shell_names.append(None)
            continue
        if shell_type == "Unknown":
            shell_names.append("Unknown")
        shell_name = get_gun_shell_type_name(shell_type)
        shell_names.append(shell_name)

    return shell_names


def dump_artian_judge(path: str) -> pd.DataFrame:
    data = None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    table = []
    for cData in data[0]["app.user_data.ArtianJudgeItemData"]["_Values"]:
        cData = cData["app.user_data.ArtianJudgeItemData.cData"]

        row = {}
        for key, value in cData.items():
            if key.startswith("_"):
                key = key[1:]
            value = minify_nested_serial(value)
            row[key] = value
        table.append(row)

    df = pd.DataFrame(table)
    df["ItemName"] = df["ItemId"].apply(get_item_name_mapping)
    return df


def dump_artian_performance(
    path: str,
    bonus_data: pd.DataFrame,
    weapon_sheets: dict[str, pd.DataFrame],
    enum_internal: dict[str, dict],
    whistle_high_freq_data: pd.DataFrame,
    whistle_hibiki_data: pd.DataFrame,
) -> pd.DataFrame:
    data = None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    table = []
    for cData in data[0]["app.user_data.ArtianPerformanceData"]["_Values"]:
        cData = cData["app.user_data.ArtianPerformanceData.cData"]

        row = {}
        for key, value in cData.items():
            if key.startswith("_"):
                key = key[1:]
            if key == "Wp05UniqueType":
                pass
            value = minify_nested_serial(value)
            value = remove_enum_value(value)

            text = text_db.get_text_by_guid(str(value))
            if text:
                value = text.replace("\n", "").replace("\r", "")

            if key == "IsLoaded":
                # 重新排序
                # sb capcom
                value = [
                    value[5],  # CLOSE
                    value[6],  # STRONG
                    value[7],  # PENETRATE
                    value[1],  # PARALYSE
                    value[0],  # POISON
                    value[2],  # SLEEP
                    value[3],  # BLAST
                    value[4],  # STAMINA
                ]
                value = process_loading_bin(value, enum_internal)
            elif key == "Wp08BinType":
                # 斩斧瓶子
                value = get_slash_axe_bin_name(value)
            elif key == "Wp05MusicSkillHighFreqType":
                # 笛子
                v = whistle_high_freq_data.loc[
                    whistle_high_freq_data["HighFreqType"] == value, "SkillName"
                ]
                if len(v) > 0:
                    value = v.values[0]
            elif key == "Wp05HibikiSkillType":
                # 笛子
                v = whistle_hibiki_data.loc[
                    whistle_hibiki_data["HiblkiSkillType"] == value, "SkillName"
                ]
                if len(v) > 0:
                    value = v.values[0]

            row[key] = value
        table.append(row)

    df = pd.DataFrame(table)

    def _apply_bonus_name(x: str):
        bonus_values = bonus_data.loc[bonus_data["BonusId"] == x, "Name"].values
        if len(bonus_values) == 0:
            return ""
        return bonus_values[0]

    df["BonusName"] = df["BonusId"].apply(_apply_bonus_name)

    # 弩炮弹药处理
    df["ShellName"] = None
    shell_lv_list_all = df["ShellLv"]
    for row_idx, shell_lv_list in shell_lv_list_all.items():
        shell_name_list = process_artian_gun_shell(shell_lv_list)
        df.at[row_idx, "ShellName"] = shell_name_list

    # 合并Shell名字，等级和数量列
    def _extract_shell_lv(lv_name: str) -> int:
        if lv_name.startswith("SL_"):
            return int(lv_name[3:]) + 1
        else:
            raise ValueError(f"Unknown shell level: {lv_name}")

    df["ExtraShellInfo"] = None
    for row_idx in df.index:
        shell_name_list = df.at[row_idx, "ShellName"]
        shell_lv_list = df.at[row_idx, "ShellLv"]
        shell_num_list = df.at[row_idx, "BowgunShellNum"]
        merged = []
        for shell_name, shell_lv, shell_num in zip(
            shell_name_list, shell_lv_list, shell_num_list
        ):
            if shell_name is None or shell_lv == "NONE" or shell_num <= 0:
                continue
            formatted = f"{shell_name} Lv{_extract_shell_lv(shell_lv)} x{shell_num}"
            merged.append(formatted)
        df.at[row_idx, "ExtraShellInfo"] = merged

    df = reindex_column(df, "ExtraShellInfo", next_to="ShellLv")
    df.drop(columns=["ShellName", "ShellLv", "BowgunShellNum"], inplace=True)

    # 处理笛子旋律
    parser = ToneParser()
    parser.set_text_db(text_db)
    parser.load_tone_table(
        "natives/STM/GameDesign/Player/ActionData/Wp05/UserData/Wp05MusicSkillToneTable.user.3.json"
    )
    parser.load_tone_color_table(
        "natives/STM/GameDesign/Player/ActionData/Wp05/UserData/Wp05MusicSkillToneColorTable.user.3.json"
    )
    parser.load_music_skill_data(
        "natives/STM/GameDesign/Common/Player/ActionGuide/MusicSkillData_Wp05.user.3.json"
    )
    parser.set_whistle_data(df)
    df = parser.parse()
    df.drop(
        columns=[
            "MusicSkills",
            "Wp05UniqueType",
        ],
        inplace=True,
    )
    df = reindex_column(df, "MusicSkillNames", next_to="Wp09BinType")

    return df


def dump_artian_bonus(path: str) -> pd.DataFrame:
    return dump_user3_data_general(path, "app.user_data.ArtianBonusData")


def dump_artian_parts(path: str) -> pd.DataFrame:
    return dump_user3_data_general(path, "app.user_data.ArtianPartsData")


def dump_artian_weapon_type(path: str, parts_data: pd.DataFrame) -> pd.DataFrame:
    data = None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    table = []
    for cData in data[0]["app.user_data.ArtianWeaponTypeData"]["_Values"]:
        cData = cData["app.user_data.ArtianWeaponTypeData.cData"]

        row = {}
        for key, value in cData.items():
            if key.startswith("_"):
                key = key[1:]
            value = minify_nested_serial(value)
            value = remove_enum_value(value)

            if key == "PartsType":
                for idx, val in enumerate(value):
                    part_name = parts_data.loc[
                        parts_data["PartsType"] == val, "Name"
                    ].values[0]
                    value[idx] = part_name.strip()
            row[key] = value
        table.append(row)

    df = pd.DataFrame(table)
    return df


if __name__ == "__main__":
    enum_internal = load_enum_internal()
    whistle_high_freq_data = dump_user3_data_general(
        "natives/STM/GameDesign/Common/Player/ActionGuide/HighFreqData_Wp05.user.3.json",
        "app.user_data.HighFreqData_Wp05",
    )
    whistle_hibiki_data = dump_user3_data_general(
        "natives/STM/GameDesign/Common/Player/ActionGuide/HibikiData_Wp05.user.3.json",
        "app.user_data.HibikiData_Wp05",
    )

    bonus_data = dump_artian_bonus(
        "natives/STM/GameDesign/Facility/ArtianBonusData.user.3.json"
    )
    parts_data = dump_artian_parts(
        "natives/STM/GameDesign/Facility/ArtianPartsData.user.3.json"
    )

    weapon_sheets = dump_weapon_data()

    sheets = {
        "ArtianPerformanceData": dump_artian_performance(
            "natives/STM/GameDesign/Facility/ArtianPerformanceData.user.3.json",
            bonus_data,
            weapon_sheets,
            enum_internal,
            whistle_high_freq_data,
            whistle_hibiki_data,
        ),
        "ArtianBonusData": bonus_data,
        "ArtianPartsData": parts_data,
        "ArtianWeaponTypeData": dump_artian_weapon_type(
            "natives/STM/GameDesign/Facility/ArtianWeaponTypeData.user.3.json",
            parts_data,
        ),
        "ArtianJudgeItemData": dump_artian_judge(
            "natives/STM/GameDesign/Facility/ArtianJudgeItemData.user.3.json"
        ),
    }

    with pd.ExcelWriter("ArtianCollection.xlsx", engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    autofit = ExcelAutoFit()
    autofit.style_excel("ArtianCollection.xlsx")
