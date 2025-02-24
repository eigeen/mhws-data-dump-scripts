import json
import os
import re
import pandas as pd

from library.item_db import ItemDB
from library.excel_auto_fit import ExcelAutoFit
from table_quest import get_mission_data

item_db = ItemDB("item_db.json")


def minify_nested_obj(obj: dict) -> str | dict:
    if isinstance(obj, dict) and len(obj) == 1:
        first_key = list(obj.keys())[0]
        value = obj[first_key]
        if "_Value" in value:
            return value["_Value"]
        # elif "userdataPath" in value:
        #     return value["userdataPath"]
        for k, v in value.items():
            value[k] = minify_nested_obj(v)

    if isinstance(obj, list):
        for i in range(len(obj)):
            obj[i] = minify_nested_obj(obj[i])

    return obj


def dump_mission_reward(path):
    data = None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    reward_db = []
    for cData in data[0]["app.user_data.MissionRewardData"]["_Values"]:
        cData = cData["app.user_data.MissionRewardData.cData"]
        reward = {}
        for key, value in cData.items():
            if key.startswith("_"):
                key = key[1:]
            val = minify_nested_obj(value)
            if isinstance(val, str):
                item = item_db.get_entry_by_id(val)
                if item:
                    val = item.raw_name
            if isinstance(val, list):
                for k in range(len(val)):
                    item = item_db.get_entry_by_id(val[k])
                    if item:
                        val[k] = item.raw_name
            # if isinstance(val, list):
            #     val = ", ".join(str(val))
            reward[key] = val
        reward_db.append(reward)

    df = pd.DataFrame(reward_db)
    return df


def dump_common_reward(path):
    data = None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    reward_db = []
    for cData in data[0]["app.user_data.QuestGeneralRewardData"]["_Values"]:
        cData = cData["app.user_data.QuestGeneralRewardData.cData"]
        reward = {}
        for key, value in cData.items():
            if key.startswith("_"):
                key = key[1:]
            val = minify_nested_obj(value)
            if isinstance(val, str):
                item = item_db.get_entry_by_id(val)
                if item:
                    val = item.raw_name
            reward[key] = val
        reward_db.append(reward)

    df = pd.DataFrame(reward_db)
    return df


sheets = {
    "MissionData": get_mission_data(),
    "MissionRewardData": dump_mission_reward(
        "natives/STM/GameDesign/Mission/_UserData/_Reward/MissionRewardData.user.3.json"
    ),
    "CommonRewardData": dump_common_reward(
        "natives/STM/GameDesign/Mission/_UserData/_Reward/CommonRewardData.user.3.json"
    ),
}

with pd.ExcelWriter("MissionDataCollection.xlsx", engine="openpyxl") as writer:
    for sheet_name, df in sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)


autofit = ExcelAutoFit()
autofit.style_excel("MissionDataCollection.xlsx")
