import json

import pandas as pd

from library.utils import minify_nested_serial, remove_enum_value

# 笛子
# app.Wp05Def.UNIQUE_TYPE_Serializable
#   -> app.Wp05MusicSkillToneTable
#   -> app.Wp05MusicSkillToneColorTable
#   -> app.user_data.MusicSkillData_Wp05


class ToneParser(object):
    def __init__(self):
        self._tone_data = None
        self._tone_color_data = None
        self._music_skill_data = None
        self._text_db = None

    def set_text_db(self, text_db):
        self._text_db = text_db

    # app.Wp05MusicSkillToneTable
    def load_tone_table(self, path: str):
        data = None
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)

        table = []
        for cData in data[0]["app.Wp05MusicSkillToneTable"]["_Datas"]:
            row = {}
            for key, value in cData["app.Wp05MusicSkillToneTable.cData"].items():
                if key.startswith("_"):
                    key = key[1:]
                value = remove_enum_value(value)

                row[key] = value
            table.append(row)
        df = pd.DataFrame(table)
        self._tone_data = df

    # app.Wp05MusicSkillToneColorTable
    def load_tone_color_table(self, path: str):
        data = None
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)

        table = []
        for cData in data[0]["app.Wp05MusicSkillToneColorTable"]["_Datas"]:
            row = {}
            for key, value in cData["app.Wp05MusicSkillToneColorTable.cData"].items():
                if key.startswith("_"):
                    key = key[1:]
                value = remove_enum_value(value)

                row[key] = value
            table.append(row)
        df = pd.DataFrame(table)
        self._tone_color_data = df

    # app.user_data.MusicSkillData_Wp05
    def load_music_skill_data(self, path: str):
        data = None
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)

        table = []
        for cData in data[0]["app.user_data.MusicSkillData_Wp05"]["_Values"]:
            row = {}
            for key, value in cData["app.user_data.MusicSkillData_Wp05.cData"].items():
                if key.startswith("_"):
                    key = key[1:]
                value = minify_nested_serial(value)
                value = remove_enum_value(value)
                if isinstance(value, str) and self._text_db is not None:
                    text = self._text_db.get_text_by_guid(value)
                    if text:
                        value = text.replace("\n", "").replace("\r", "")

                row[key] = value
            table.append(row)
        df = pd.DataFrame(table)
        self._music_skill_data = df

    # app.Wp05Def.UNIQUE_TYPE_Serializable
    def set_whistle_data(self, data: pd.DataFrame):
        self._whistle_data = data

    def parse(self) -> pd.DataFrame:
        assert self._tone_data is not None
        assert self._tone_color_data is not None
        assert self._music_skill_data is not None
        assert self._whistle_data is not None

        # app.Wp05Def.UNIQUE_TYPE_Serializable
        unique_types = self._whistle_data["Wp05UniqueType"]
        colors_all = [self._get_colors_by_type(t) for t in unique_types]
        skills_all = [self._get_skills_by_colors(c) for c in colors_all]

        output_data = self._whistle_data.copy(deep=True)
        output_data["MusicSkills"] = skills_all

        skill_names = []
        for skills in skills_all:
            skill_names.append([self._get_skill_name(s) for s in skills])
        skill_names = list(filter(lambda x: x is not None, skill_names))
        output_data["MusicSkillNames"] = skill_names
        return output_data

    def _get_colors_by_type(self, unique_type: str) -> list[str] | None:
        colors = self._tone_data.loc[
            self._tone_data["UniqueType"] == unique_type,
            ["ToneColor1", "ToneColor2", "ToneColor3"],
        ]
        if len(colors) == 0:
            return None
        return colors.values[0].tolist()

    def _get_skills_by_colors(self, colors: list[str]) -> list[str]:
        skills = []
        for _, row in self._tone_color_data.iterrows():
            skill = row["MusicSkill"]
            skill_colors = {
                row["ToneColor1"],
                row["ToneColor2"],
                row["ToneColor3"],
                row["ToneColor4"],
            }
            skill_colors.discard("INVALID")
            all_contains = True
            for color in colors:
                if color not in skill_colors:
                    all_contains = False
                    break
            if all_contains:
                skills.append(skill)
        return skills

    def _get_skill_name(self, skill_id: str) -> str | None:
        if skill_id == "INVALID":
            return None
        v = self._music_skill_data.loc[
            self._music_skill_data["MusicSkillType"] == skill_id, "MusicSkillName"
        ]
        if len(v) == 0:
            return None
        return v.values[0]


if __name__ == "__main__":
    parser = ToneParser()

    from library.text_db import load_text_db

    parser.set_text_db(load_text_db("texts_db.json"))

    parser.load_tone_table(
        "natives/STM/GameDesign/Player/ActionData/Wp05/UserData/Wp05MusicSkillToneTable.user.3.json"
    )
    parser.load_tone_color_table(
        "natives/STM/GameDesign/Player/ActionData/Wp05/UserData/Wp05MusicSkillToneColorTable.user.3.json"
    )
    parser.load_music_skill_data(
        "natives/STM/GameDesign/Common/Player/ActionGuide/MusicSkillData_Wp05.user.3.json"
    )

    from table_equip import dump_weapon_data

    weapon_sheets = dump_weapon_data()
    parser.set_whistle_data(weapon_sheets["Whistle"])

    result = parser.parse()
    print(result)
