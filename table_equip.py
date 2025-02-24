import json
import pandas as pd

from library.excel_auto_fit import ExcelAutoFit
from library.text_db import load_text_db
from library.utils import (
    is_guid_like,
    minify_nested_serial,
    remove_enum_value,
    reindex_column,
)
from library.rare import apply_rare_colors
from table_skill import dump_skill_common_data

text_db = load_text_db("texts_db.json")


def get_weapon_types() -> dict[str, int]:
    return {
        "LongSword": 0,
        "ShortSword": 1,
        "TwinSword": 2,
        "Tachi": 3,
        "Hammer": 4,
        "Whistle": 5,
        "Lance": 6,
        "Gunlance": 7,
        "SlashAxe": 8,
        "ChargeAxe": 9,
        "Rod": 10,
        "Bow": 11,
        "HeavyBowgun": 12,
        "LightBowgun": 13,
    }


def _dump_weapon_data(
    path: str,
    weapon_type: str,
    weapon_types: dict[str, int],
    skill_common_data: pd.DataFrame,
    keep_serial_id: bool = False,
) -> pd.DataFrame:
    data = None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    table = []
    for cData in data[0]["app.user_data.WeaponData"]["_Values"]:
        row = {}
        for key, value in cData["app.user_data.WeaponData.cData"].items():
            if key.startswith("_"):
                key = key[1:]
            key_lower = key.lower()
            if key_lower == weapon_type:
                key = "Id"
            if key_lower in weapon_types and key_lower != weapon_type:
                continue
            # 批量处理列名带 wpxx 的列，排除无关列
            exclude = False
            for wp_type, wp_type_id in weapon_types.items():
                wp_short_id = f"wp{wp_type_id:02d}"
                if key_lower.find(wp_short_id) != -1 and wp_type != weapon_type:
                    exclude = True
                    break
            if exclude:
                continue
            if weapon_type != "rod" and key == "RodInsectLv":
                continue
            if weapon_type not in {"heavybowgun", "lightbowgun"} and key in {
                "MainShell",
                "ShellLv",
                "ShellNum",
                "RapidShellNum",
                "IsRappid",
                "CustomizePattern",
                "DispSilencer",
                "DispBarrel",
            }:
                continue
            if weapon_type != "heavybowgun" and key in {
                "EnergyEfficiency",
                "AmmoStrength",
                "EnergyShellTypeNormal",
                "EnergyShellTypeNormal",
                "EnergyShellTypePower",
                "EnergyShellTypeWeak",
            }:
                continue
            if weapon_type != "bow" and key == "isLoadingBin":
                continue
            if weapon_type in {"lightbowgun", "heavybowgun", "bow"} and key in {
                "SharpnessValList",
                "TakumiValList",
            }:
                continue

            value = minify_nested_serial(value)
            # 处理Skill
            if key == "Skill":
                for i, skill in enumerate(value):
                    if skill.find("NONE") != -1:
                        continue
                    skill_name = skill_common_data.loc[
                        skill_common_data["skillId"] == skill, "skillName"
                    ]
                    if len(skill_name) > 0:
                        value[i] = skill_name.values[0]

            if is_guid_like(str(value)):
                # process guids
                text = text_db.get_text_by_guid(value)
                if text:
                    value = text.replace("\n", "").replace("\r", "")
                else:
                    value = ""

            # 处理SlotLevel
            if key == "SlotLevel":
                for i, level in enumerate(value):
                    if level.find("NONE") != -1:
                        value[i] = 0
                    elif level.find("Lv1") != -1:
                        value[i] = 1
                    elif level.find("Lv2") != -1:
                        value[i] = 2
                    elif level.find("Lv3") != -1:
                        value[i] = 3
                    else:
                        print(f"Unknown level: {level}")

            if not keep_serial_id:
                value = remove_enum_value(value)

            row[key] = value
        table.append(row)

    df = pd.DataFrame(table)
    # 拆分skill列
    skill_names = df["Skill"]
    skill_levels = df["SkillLevel"]
    all_skills = []
    for i in range(len(skill_names)):
        skills = []
        for j in range(len(skill_names[i])):
            level = skill_levels[i][j]
            if level == 0:
                skills.append(None)
            else:
                skills.append(f"{skill_names[i][j]}: {level}")
        skills = list(filter(lambda x: x is not None, skills))
        all_skills.append(skills)
    df["SkillAndLevel"] = all_skills

    df.drop(columns=["Skill", "SkillLevel"], inplace=True)
    df = reindex_column(df, "SkillAndLevel", next_to="SlotLevel")

    df = reindex_column(df, ["ModelId", "CustomModelId"], to_end=True)

    return df


def dump_weapon_data(keep_serial_id: bool = False) -> dict[str, pd.DataFrame]:
    weapon_types = get_weapon_types()
    weapon_types_lower = {}
    for key, value in weapon_types.items():
        weapon_types_lower[key.lower()] = value

    weapon_paths = {}
    for weapon_type in weapon_types.keys():
        weapon_paths[weapon_type] = (
            f"natives/STM/GameDesign/Common/Weapon/{weapon_type}.user.3.json"
        )

    skill_common_data = dump_skill_common_data(
        "natives/STM/GameDesign/Common/Equip/SkillCommonData.user.3.json"
    )

    sheets = {}
    for weapon_type, path in weapon_paths.items():
        df = _dump_weapon_data(
            weapon_paths[weapon_type],
            weapon_type.lower(),
            weapon_types_lower,
            skill_common_data,
            keep_serial_id=keep_serial_id,
        )
        sheets[weapon_type] = df
    return sheets


def _dump_armor_data(
    path: str,
    skill_common_data: pd.DataFrame,
    armor_series_data: pd.DataFrame,
) -> pd.DataFrame:
    data = None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    table = []
    for cData in data[0]["app.user_data.ArmorData"]["_Values"]:
        row = {}
        for key, value in cData["app.user_data.ArmorData.cData"].items():
            if key.startswith("_"):
                key = key[1:]

            value = minify_nested_serial(value)
            # 处理Skill
            if key == "Skill":
                for i, skill in enumerate(value):
                    if skill.find("NONE") != -1:
                        continue
                    skill_name = skill_common_data.loc[
                        skill_common_data["skillId"] == skill, "skillName"
                    ]
                    if len(skill_name) > 0:
                        value[i] = skill_name.values[0]

            if is_guid_like(str(value)):
                # process guids
                text = text_db.get_text_by_guid(value)
                if text:
                    value = text.replace("\n", "").replace("\r", "")
                else:
                    value = ""
            value = remove_enum_value(value)
            # 处理SlotLevel
            if key == "SlotLevel":
                for i, level in enumerate(value):
                    if level == "NONE":
                        value[i] = 0
                    elif level == "Lv1":
                        value[i] = 1
                    elif level == "Lv2":
                        value[i] = 2
                    elif level == "Lv3":
                        value[i] = 3
                    else:
                        print(f"Unknown level: {level}")
            # 处理Series
            if key == "Series":
                v = armor_series_data.loc[armor_series_data["Series"] == value, "Name"]
                if len(v) > 0:
                    value = v.iloc[0]

            row[key] = value
        table.append(row)

    df = pd.DataFrame(table)
    # 拆分skill列
    skill_names = df["Skill"]
    skill_levels = df["SkillLevel"]
    all_skills = []
    for i in range(len(skill_names)):
        skills = []
        for j in range(len(skill_names[i])):
            level = skill_levels[i][j]
            if level == 0:
                skills.append(None)
            else:
                skills.append(f"{skill_names[i][j]}: {level}")
        skills = list(filter(lambda x: x is not None, skills))
        all_skills.append(skills)
    df["SkillAndLevel"] = all_skills

    df.drop(columns=["Skill", "SkillLevel"], inplace=True)
    df = reindex_column(df, "SkillAndLevel", next_to="SlotLevel")

    return df


def dump_weapon_series_data(path: str) -> pd.DataFrame:
    data = None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    table = []
    for cData in data[0]["app.user_data.WeaponSeriesData"]["_Values"]:
        cData = cData["app.user_data.WeaponSeriesData.cData"]
        row = {}
        for key, value in cData.items():
            if key.startswith("_"):
                key = key[1:]

            value = minify_nested_serial(value)
            value = remove_enum_value(value)
            text = text_db.get_text_by_guid(str(value))
            if text:
                value = text
            row[key] = value
        table.append(row)
    df = pd.DataFrame(table)
    return df


def dump_armor_series_data(path: str) -> pd.DataFrame:
    data = None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    table = []
    for cData in data[0]["app.user_data.ArmorSeriesData"]["_Values"]:
        cData = cData["app.user_data.ArmorSeriesData.cData"]
        row = {}
        for key, value in cData.items():
            if key.startswith("_"):
                key = key[1:]

            value = minify_nested_serial(value)
            value = remove_enum_value(value)
            text = text_db.get_text_by_guid(str(value))
            if text:
                value = text
            row[key] = value
        table.append(row)
    return pd.DataFrame(table)


if __name__ == "__main__":
    skill_common_data = dump_skill_common_data(
        "natives/STM/GameDesign/Common/Equip/SkillCommonData.user.3.json"
    )
    armorseries_data = dump_armor_series_data(
        "natives/STM/GameDesign/Common/Equip/ArmorSeriesData.user.3.json"
    )
    armor_data = _dump_armor_data(
        "natives/STM/GameDesign/Common/Equip/ArmorData.user.3.json",
        skill_common_data,
        armorseries_data,
    )
    sheets = dump_weapon_data()
    with pd.ExcelWriter("EquipCollection.xlsx") as writer:
        armor_data.to_excel(writer, sheet_name="Armor", index=False)

        for sheet_name, df in sheets.items():
            # 武器表名前加 Wp_ 前缀
            sheet_name = f"Wp_{sheet_name}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        print("Applying autofit...")
        autofit = ExcelAutoFit()
        autofit.style_workbook(writer.book)

        print("Applying rare colors...")
        apply_rare_colors(writer.book)
