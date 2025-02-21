import csv
import json

from library.item_db import ItemDB

item_db = ItemDB("item_db.json")

recipe_data = None
with open("natives/STM/GameDesign/Common/Item/ItemRecipe.user.3.json", "r") as f:
    recipe_data = json.load(f)

table = []
for cData in recipe_data[0]["app.user_data.cItemRecipe"]["_Values"]:
    cData = cData["app.user_data.cItemRecipe.cData"]

    row = {}
    for col_name, col_value in cData.items():
        if col_name.startswith("_"):
            col_name = col_name[1:]
        # replace item_id with item_name
        if type(col_value) == list:
            for i, item_id in enumerate(col_value):
                item_entry = item_db.get_entry_by_id(str(item_id))
                if item_entry:
                    col_value[i] = item_entry.raw_name
            col_value = ", ".join(col_value)

        item_entry = item_db.get_entry_by_id(str(col_value))
        if item_entry:
            col_value = item_entry.raw_name

        row[col_name] = col_value

    table.append(row)

with open("ItemRecipe.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=table[0].keys())
    writer.writeheader()
    writer.writerows(table)
