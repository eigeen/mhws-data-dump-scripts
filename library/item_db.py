from dataclasses import dataclass
import json
import re

re_item_id = re.compile(r"\[[-\d]+\]ITEM_(\d+)")


@dataclass
class ItemEntry:
    index: int
    id: str
    raw_name: str
    raw_explain: str


class ItemDB:
    def __init__(self, path):
        self.items = {}

        data = None
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not data:
            raise ValueError("Invalid ItemDB data")

        for item in data:
            self.items[item["_ItemId"]] = ItemEntry(
                item["_Index"],
                item["_ItemId"],
                item["_RawName"],
                item["_RawExplain"],
            )

    def get_entry_by_id(self, item_id: str) -> ItemEntry | None:
        match = re_item_id.match(item_id)
        if not match:
            return self.items.get(item_id)
        else:
            return self.items.get(f"ITEM_{match.group(1)}")


if __name__ == "__main__":
    item_db = ItemDB("item_db.json")
    assert item_db.get_entry_by_id("[2]ITEM_0000") is not None
