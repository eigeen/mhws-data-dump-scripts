import csv
import json
import re

from library.text_db import load_text_db

re_guid_like = re.compile(
    r"\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b"
)
text_db = load_text_db("texts_db.json")

enemy_data = None
with open("natives/STM/GameDesign/Common/Enemy/EnemyData.user.3.json", "r") as f:
    enemy_data = json.load(f)

table = []
for cData in enemy_data[0]["app.user_data.EnemyData"]["_Values"]:
    row = {}
    for field, value in cData["app.user_data.EnemyData.cData"].items():
        col_name = field
        if col_name.startswith("_"):
            col_name = col_name[1:]
        if re_guid_like.match(str(value)):
            # process guids
            text = text_db.get_text_by_guid(value)
            if text:
                text = text.replace("\n", "").replace("\r", "")
                value = text
            else:
                value = ""

        row[col_name] = value
    table.append(row)

with open("enemy_data.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=table[0].keys())
    writer.writeheader()
    writer.writerows(table)
