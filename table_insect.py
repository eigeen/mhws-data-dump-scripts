import json
import pandas as pd

from library.excel_auto_fit import ExcelAutoFit
from library.utils import (
    remove_enum_value,
    minify_nested_serial,
    reindex_column,
)
from library.rare import apply_fix_rare_colors
from library.text_db import get_global_text_db
from library.item_db import get_global_item_db
from table_quest import get_mission_data

text_db = get_global_text_db()
item_db = get_global_item_db()


def dump_insect_data(path: str) -> pd.DataFrame:
    data = None
    with open(path, "r") as f:
        data = json.load(f)

    table = []
    for cData in data[0]["app.user_data.RodInsectData"]["_Values"]:
        cData = cData["app.user_data.RodInsectData.cData"]
        row = {}
        for key, value in cData.items():
            if key.startswith("_"):
                key = key[1:]
            if isinstance(value, str):
                text = text_db.get_text_by_guid(value)
                if text:
                    value = text

            value = minify_nested_serial(value)
            value = remove_enum_value(value)
            row[key] = value
        table.append(row)

    df = pd.DataFrame(table)
    # 排序
    df = reindex_column(df, column=["ModelID"], to_end=True)
    return df


def dump_insect_recipe_data(
    path: str, insect_data: pd.DataFrame, mission_data: pd.DataFrame
) -> pd.DataFrame:
    data = None
    with open(path, "r") as f:
        data = json.load(f)

    table = []
    for cData in data[0]["app.user_data.RodInsectRecipeData"]["_Values"]:
        cData = cData["app.user_data.RodInsectRecipeData.cData"]
        row = {}
        for key, value in cData.items():
            if key.startswith("_"):
                key = key[1:]

            value = minify_nested_serial(value)
            value = remove_enum_value(value)
            if isinstance(value, str):
                if value in {"NONE", "INVALID"}:
                    value = ""
            if key == "ItemId":
                for i, item_enum in enumerate(value):
                    if item_enum == "INVALID":
                        value[i] = None
                    item = item_db.get_entry_by_id(item_enum)
                    if item:
                        value[i] = item.raw_name
                value = list(filter(lambda x: x is not None, value))

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
    df = reindex_column(df, column=["KeyStory"], next_to="KeyStoryNo")
    # 替换ID列
    names = []
    for _, row in df.iterrows():
        id_enum = row["ID"]
        v = insect_data.loc[insect_data["Id"] == id_enum, "Name"]
        if len(v) > 0:
            names.append(v.iloc[0])
        else:
            names.append("")
    df["ID"] = names
    # 替换PrevID列
    prev_names = []
    for _, row in df.iterrows():
        id_enum = row["PrevID"]
        v = insect_data.loc[insect_data["Id"] == id_enum, "Name"]
        if len(v) > 0:
            prev_names.append(v.iloc[0])
        else:
            prev_names.append("")
    df["PrevID"] = prev_names

    df.rename(columns={"ID": "Name", "PrevID": "PrevName"}, inplace=True)
    df = reindex_column(df, column=["Name", "PrevName"], next_to="Index")

    return df


if __name__ == "__main__":
    mission_data = get_mission_data()

    insect_data = dump_insect_data(
        "natives/STM/GameDesign/Common/Weapon/RodInsectData.user.3.json"
    )
    insect_recipe_data = dump_insect_recipe_data(
        "natives/STM/GameDesign/Common/Equip/RodInsectRecipeData.user.3.json",
        insect_data,
        mission_data,
    )

    autofit = ExcelAutoFit()
    with pd.ExcelWriter("RodInsectCollection.xlsx") as writer:
        insect_data.to_excel(writer, sheet_name="RodInsectData")
        insect_recipe_data.to_excel(writer, sheet_name="RodInsectRecipeData")

        apply_fix_rare_colors(writer.book)
        autofit.style_workbook(writer.book)
