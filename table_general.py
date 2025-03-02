import json
import pandas as pd


def dump_enum_maker(path: str) -> pd.DataFrame:
    data = None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    table = []
    for cData in data[0]["ace.user_data.EnumMaker"]["_DataList"]:
        row = {}
        for key, value in cData["ace.user_data.EnumMaker.cData"].items():
            if key.startswith("_"):
                key = key[1:]
            row[key] = value
        table.append(row)
    df = pd.DataFrame(table)
    return df


# 创建单元格内嵌的图片元数据，用于后续设置图片
def create_icon_flag(path: str, metadata: dict = {}) -> str:
    metadata["path"] = path
    metadata_str = json.dumps(metadata)
    return f"![{metadata_str}]"


# 解析单元格内嵌的图片元数据
def parse_icon_flag(text: str) -> dict | None:
    if text.startswith("![") and text.endswith("]"):
        metadata_str = text[2:-1]
        metadata = json.loads(metadata_str)
        return metadata
    return None


def load_enum_internal() -> dict[str, dict]:
    data = None
    with open("Enums_Internal.json", "r", encoding="utf-8") as f:
        data = json.load(f)
    return data
