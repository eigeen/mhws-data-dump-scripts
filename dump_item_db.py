import json
import os

from library.text_db import load_text_db
from library.utils import is_guid_like, remove_enum_value

text_db = load_text_db("texts_db.json")


def dump_item_db(path: str) -> list[dict[str, str]]:
    content = None
    with open(path, "r", encoding="utf-8") as f:
        content = json.load(f)

    item_db = []
    for cData in content[0]["app.user_data.ItemData"]["_Values"]:
        cData = cData["app.user_data.ItemData.cData"]

        entry = {}
        for col_name, col_data in cData.items():
            if is_guid_like(str(col_data)):
                text = text_db.get_text_by_guid(col_data)
                if text:
                    text = text.replace("\n", "").replace("\r", "")
                    col_data = text
                else:
                    col_data = ""
            else:
                col_data = remove_enum_value(col_data)
            entry[col_name] = col_data
        item_db.append(entry)

    return item_db


item_db = dump_item_db("natives/STM/GameDesign/Common/Item/itemData.user.3.json")
with open("item_db.json", "w", encoding="utf-8") as f:
    json.dump(item_db, f, ensure_ascii=False, indent=4)
