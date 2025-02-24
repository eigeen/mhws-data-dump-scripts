import json
import pandas as pd

from library.excel_auto_fit import ExcelAutoFit
from library.item_db import ItemDB
from library.utils import remove_enum_value, reindex_column
from table_equip import (
    dump_armor_series_data,
    get_weapon_types,
    dump_weapon_data,
    dump_armor_data,
)
from table_enemy import dump_enemy_data
from table_quest import get_mission_data

item_db = ItemDB("item_db.json")
weapon_types = get_weapon_types()


def dump_weapon_recipe_data(
    path: str,
    weapon_data: pd.DataFrame,
    enemy_data: pd.DataFrame,
    mission_data: pd.DataFrame,
    weapon_type: str,
    weapon_types: dict[str, int],
) -> pd.DataFrame:
    data = None
    with open(path, "r") as f:
        data = json.load(f)

    table = []
    for cData in data[0]["app.user_data.WeaponRecipeData"]["_Values"]:
        cData = cData["app.user_data.WeaponRecipeData.cData"]
        row = {}
        for key, value in cData.items():
            if key.startswith("_"):
                key = key[1:]
            if key == "GunLance":
                key = "Gunlance"  # sb capcom
            if key in weapon_types and weapon_type != key:
                continue

            value = remove_enum_value(value)
            if key == weapon_type:
                key = "Name"
                v = weapon_data.loc[weapon_data["Id"] == value, "Name"]
                if len(v) > 0:
                    value = v.iloc[0]
            if key == "KeyEnemyId" and value != "INVALID":
                v = enemy_data.loc[enemy_data["enemyId"] == value, "EnemyName"]
                if len(v) > 0:
                    value = v.iloc[0]

            if isinstance(value, list):
                for i, v in enumerate(value):
                    item = item_db.get_entry_by_id(str(v))
                    if item:
                        value[i] = item.raw_name
            elif isinstance(value, str):
                item = item_db.get_entry_by_id(str(value))
                if item:
                    value = item.raw_name

            if value == "INVALID":
                value = ""

            row[key] = value
        table.append(row)

    df = pd.DataFrame(table)
    # 增加KeyStory列
    stories = []
    for _, row in df.iterrows():
        id = row["KeyStoryNo"]
        if not id:
            stories.append("")
            continue
        v = mission_data.loc[mission_data["MissionIDSerial"] == id, "SetLGuideMsgData"]
        if len(v) > 0:
            stories.append(v.iloc[0])
        else:
            stories.append("")
    df["KeyStory"] = stories
    df = reindex_column(df, column="KeyStory", next_to="KeyStoryNo")

    # 列重命名
    df.rename(columns={"KeyItemId": "KeyItem", "KeyEnemyId": "KeyEnemy"}, inplace=True)

    return df


def dump_armor_recipe_data(
    path: str,
    armor_series_data: pd.DataFrame,
    armor_data: pd.DataFrame,
    mission_data: pd.DataFrame,
) -> pd.DataFrame:
    data = None
    with open(path, "r") as f:
        data = json.load(f)

    table = []
    for cData in data[0]["app.user_data.ArmorRecipeData"]["_Values"]:
        cData = cData["app.user_data.ArmorRecipeData.cData"]
        row = {}
        for key, value in cData.items():
            if key.startswith("_"):
                key = key[1:]

            value = remove_enum_value(value)

            if key == "KeyEnemyId" and value != "INVALID":
                v = enemy_data.loc[enemy_data["enemyId"] == value, "EnemyName"]
                if len(v) > 0:
                    value = v.iloc[0]
            # Series 名
            if key == "SeriesId":
                v = armor_series_data.loc[armor_series_data["Series"] == value, "Name"]
                if len(v) > 0:
                    value = v.iloc[0]

            # ItemName
            if isinstance(value, list):
                for i, v in enumerate(value):
                    if v == "INVALID":
                        value[i] = None
                        continue
                    item = item_db.get_entry_by_id(str(v))
                    if item:
                        value[i] = item.raw_name
                value = list(filter(lambda x: x is not None, value))
            elif isinstance(value, str):
                if value == "INVALID":
                    value = ""
                item = item_db.get_entry_by_id(str(value))
                if item:
                    value = item.raw_name

            row[key] = value
        table.append(row)

    df = pd.DataFrame(table)
    # 新增装备名称列
    name_col = []
    for _, row in df.iterrows():
        series = row["SeriesId"]
        parts_type = row["PartsType"]
        v = armor_data.loc[
            (armor_data["Series"] == series) & (armor_data["PartsType"] == parts_type),
            "Name",
        ]
        if len(v) > 0:
            name_col.append(v.iloc[0])
        else:
            name_col.append("")
    df["Name"] = name_col
    df = reindex_column(df, column="Name", next_to="PartsType")

    # 增加KeyStory列
    stories = []
    for _, row in df.iterrows():
        id = row["KeyStoryNo"]
        if not id:
            stories.append("")
            continue
        v = mission_data.loc[mission_data["MissionIDSerial"] == id, "SetLGuideMsgData"]
        if len(v) > 0:
            stories.append(v.iloc[0])
        else:
            stories.append("")
    df["KeyStory"] = stories
    df = reindex_column(df, column="KeyStory", next_to="KeyStoryNo")

    df.rename(columns={"KeyItemId": "KeyItem", "KeyEnemyId": "KeyEnemy"}, inplace=True)

    return df


# 合并 Item 和 ItemNum
def _merge_item_and_num(df: pd.DataFrame) -> pd.DataFrame:
    item_and_nums = []
    for _, row in df.iterrows():
        item_and_num = []
        for item, num in zip(row["Item"], row["ItemNum"]):
            if num == 0:
                continue
            item_and_num.append(f"{item} x{num}")
        item_and_nums.append(item_and_num)
    df["ItemAndNum"] = item_and_nums
    df = reindex_column(df, column="ItemAndNum", next_to="ItemNum")
    df.drop(columns=["Item", "ItemNum"], inplace=True)
    return df


if __name__ == "__main__":
    weapon_sheet = dump_weapon_data()
    enemy_data = dump_enemy_data(
        "natives/STM/GameDesign/Common/Enemy/EnemyData.user.3.json",
        "natives/STM/GameDesign/Common/Enemy/EnemySpecies.user.3.json",
    )
    mission_data = get_mission_data()

    armor_series_data = dump_armor_series_data(
        "natives/STM/GameDesign/Common/Equip/ArmorSeriesData.user.3.json"
    )
    armor_data = dump_armor_data()
    armor_recipe_data = dump_armor_recipe_data(
        "natives/STM/GameDesign/Common/Equip/ArmorRecipeData.user.3.json",
        armor_series_data,
        armor_data,
        mission_data,
    )
    # 合并 Item 和 ItemNum
    armor_recipe_data = _merge_item_and_num(armor_recipe_data)

    weapon_sheets = {}
    for weapon_type in weapon_types.keys():
        print(f"Dumping {weapon_type} recipe data...")
        data = dump_weapon_recipe_data(
            f"natives/STM/GameDesign/Common/Weapon/{weapon_type}Recipe.user.3.json",
            weapon_sheet[weapon_type],
            enemy_data,
            mission_data,
            weapon_type,
            weapon_types,
        )
        # 合并 Item 和 ItemNum
        data = _merge_item_and_num(data)

        weapon_sheets[weapon_type] = data

    auto_fit = ExcelAutoFit()
    with pd.ExcelWriter("EquipRecipeCollection.xlsx") as writer:
        armor_recipe_data.to_excel(writer, sheet_name="Armor", index=False)

        for weapon_type, data in weapon_sheets.items():
            data.to_excel(writer, sheet_name=f"Wp_{weapon_type}", index=False)

        auto_fit.style_workbook(writer.book)
