import json
import re
import pandas as pd

from library.excel_auto_fit import ExcelAutoFit
from library.text_db import load_text_db
from library.utils import is_guid_like

text_db = load_text_db("texts_db.json")


def dump_enemy_data(path: str, species_data: pd.DataFrame) -> pd.DataFrame:
    data = None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    table = []
    for cData in data[0]["app.user_data.EnemyData"]["_Values"]:
        row = {}
        for key, value in cData["app.user_data.EnemyData.cData"].items():
            if key.startswith("_"):
                key = key[1:]

            if is_guid_like(str(value)):
                # process guids
                text = text_db.get_text_by_guid(value)
                if text:
                    value = text.replace("\n", "").replace("\r", "")
                else:
                    value = ""
            if key == "Species":
                specie_name = species_data.loc[
                    species_data["EmSpecies"] == value, "EmSpeciesName"
                ]
                if len(specie_name) > 0:
                    value = specie_name.values[0]

            row[key] = value
        table.append(row)
    df = pd.DataFrame(table)
    return df


def dump_species_data(path: str) -> pd.DataFrame:
    data = None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    table = []
    for cData in data[0]["app.user_data.EnemySpeciesData"]["_Values"]:
        row = {}
        for key, value in cData["app.user_data.EnemySpeciesData.cData"].items():
            if key.startswith("_"):
                key = key[1:]
            if is_guid_like(str(value)):
                # process guids
                text = text_db.get_text_by_guid(value)
                if text:
                    value = text.replace("\n", "").replace("\r", "")
                else:
                    value = ""

            row[key] = value
        table.append(row)
    df = pd.DataFrame(table)
    return df


species_data = dump_species_data(
    "natives/STM/GameDesign/Common/Enemy/EnemySpecies.user.3.json"
)
enemy_data = dump_enemy_data(
    "natives/STM/GameDesign/Common/Enemy/EnemyData.user.3.json", species_data
)
# 统计种族数量
species_counts = enemy_data["Species"].value_counts()
species_counts = species_counts.reset_index()
species_counts.columns = ["EmSpeciesName", "Count"]
species_data = species_data.merge(species_counts, on="EmSpeciesName")

sheets = {
    "EnemyData": enemy_data,
    "SpeciesData": species_data,
}

with pd.ExcelWriter("EnemyCollection.xlsx", engine="openpyxl") as writer:
    for sheet_name, df in sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)


autofit = ExcelAutoFit()
autofit.style_excel("EnemyCollection.xlsx")
