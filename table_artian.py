import json
import pandas as pd

from library.excel_auto_fit import ExcelAutoFit
from library.item_db import ItemDB
from library.text_db import load_text_db
from library.utils import minify_nested_serial, remove_enum_value
from table_equip import dump_weapon_data

item_db = ItemDB("item_db.json")
text_db = load_text_db("texts_db.json")


def get_item_name_mapping(x):
    item = item_db.get_entry_by_id(str(x))
    if item is None:
        return ""
    return item.raw_name


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
    path: str, bonus_data: pd.DataFrame, weapon_sheets: dict[str, pd.DataFrame]
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
            value = minify_nested_serial(value)
            value = remove_enum_value(value)

            text = text_db.get_text_by_guid(str(value))
            if text:
                value = text.replace("\n", "").replace("\r", "")
                
            if key == "Wp05UniqueType":
                pass
            
            row[key] = value
        table.append(row)

    df = pd.DataFrame(table)

    def apply_bonus_name(x: str):
        bonus_values = bonus_data.loc[bonus_data["BonusId"] == x, "Name"].values
        if len(bonus_values) == 0:
            return ""
        return bonus_values[0]

    df["BonusName"] = df["BonusId"].apply(apply_bonus_name)
    return df


def dump_artian_bonus(path: str) -> pd.DataFrame:
    data = None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    table = []
    for cData in data[0]["app.user_data.ArtianBonusData"]["_Values"]:
        cData = cData["app.user_data.ArtianBonusData.cData"]

        row = {}
        for key, value in cData.items():
            if key.startswith("_"):
                key = key[1:]
            value = minify_nested_serial(value)

            text = text_db.get_text_by_guid(str(value))
            if text:
                value = text.replace("\n", "").replace("\r", "")
            row[key] = value
        table.append(row)

    df = pd.DataFrame(table)
    return df


def dump_artian_parts(path: str) -> pd.DataFrame:
    data = None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    table = []
    for cData in data[0]["app.user_data.ArtianPartsData"]["_Values"]:
        cData = cData["app.user_data.ArtianPartsData.cData"]

        row = {}
        for key, value in cData.items():
            if key.startswith("_"):
                key = key[1:]
            value = minify_nested_serial(value)

            text = text_db.get_text_by_guid(str(value))
            if text:
                value = text.replace("\n", "").replace("\r", "")
            row[key] = value
        table.append(row)

    df = pd.DataFrame(table)
    return df


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
