import json
import os
import re
import csv

from library.text_db import load_text_db

text_db = load_text_db("texts_db.json")
re_mission_id = re.compile(r"\[[\d-]+\]MISSION_(\d+)")

PATH_ROOT = "natives/STM/"
USELESS_COL_NAMES = set([
    "MissionPrefab",
    "SetSGuideMsgDataList",
    "EmSetDataList",
    "EnemySetDataList",
    "MissionGmSetPrefab"
])


def get_mission_ud_paths() -> list[str]:
    mission_list_data = None
    with open(
        "natives/STM/GameDesign/Mission/_UserData/MissionListData_00.user.3.json",
        "r",
        encoding="utf-8",
    ) as f:
        mission_list_data = json.load(f)

    paths = []

    for data in mission_list_data[0]["app.user_data.MissionListData"]["_DataList"]:
        path = data["app.user_data.MissionData"]["userdataPath"]
        paths.append(path)

    return paths


def minify_nested_obj(obj: dict) -> str | dict:
    if isinstance(obj, dict) and len(obj) == 1:
        first_key = list(obj.keys())[0]
        value = obj[first_key]
        if "_Value" in value:
            return value["_Value"]
        elif "userdataPath" in value:
            return value["userdataPath"]
        for k, v in value.items():
            value[k] = minify_nested_obj(v)

    if isinstance(obj, list):
        for i in range(len(obj)):
            obj[i] = minify_nested_obj(obj[i])

    return obj


def sort_by_mission_id(obj: dict) -> int:
    id1_num = int(re_mission_id.search(obj["MissionIDSerial"]).group(1))
    return id1_num


def get_mission_data() -> list[dict]:
    mission_ud_paths = get_mission_ud_paths()
    mission_ud_paths = map(lambda path: PATH_ROOT + path + ".3.json", mission_ud_paths)

    mission_datas = []
    for mission_ud_path in mission_ud_paths:
        mission_data = None
        try:
            with open(mission_ud_path, "r", encoding="utf-8") as f:
                mission_data = json.load(f)
        except Exception as e:
            print(f"Error loading mission data: {e}")
            continue

        mission_data = mission_data[0]["app.user_data.MissionData"]
        data = {}
        for key, value in mission_data.items():
            if key.startswith("_"):
                key = key[1:]

            # try minify serial id
            value = minify_nested_obj(value)

            if key == "SetLGuideMsgData":
                id = value["app.user_data.MissionData.GuideMsgParts"]["SetMsgID"]
                msg = text_db.get_text_by_guid(id)
                value = msg or ""
            elif key == "SetSGuideMsgDataList":
                curr_i = 0
                for i, id in enumerate(value):
                    id = id["app.user_data.MissionData.GuideMsgParts"]["SetMsgID"]
                    curr_i = i
                    if i < 5:
                        data[f"SetSGuideMsgData{i}"] = (
                            text_db.get_text_by_guid(id) or ""
                        )
                if curr_i != 4:
                    for i in range(curr_i, 5):
                        data[f"SetSGuideMsgData{i}"] = ""
            elif key in USELESS_COL_NAMES:
                continue
            # # ignore nested objects
            # if isinstance(value, dict):
            #     continue
            # if isinstance(value, list):
            #     continue

            data[key] = value
        mission_datas.append(data)

    mission_datas.sort(key=sort_by_mission_id)
    return mission_datas


# with open("Missions.csv", "w", encoding="utf-8", newline="") as f:
#     writer = csv.DictWriter(f, fieldnames=mission_datas[0].keys())
#     writer.writeheader()
#     writer.writerows(mission_datas)
