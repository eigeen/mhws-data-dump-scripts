import re
import pandas as pd

re_guid_like = re.compile(
    r"\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b"
)
re_enum_value = re.compile(r"^\[(.*?)\](.*)$")
re_rare = re.compile(r"^RARE(\d+)$")


def is_guid_like(text: str) -> bool:
    return re_guid_like.match(text) is not None


def minify_nested_serial(obj: dict) -> str | dict:
    if isinstance(obj, dict) and len(obj) == 1:
        first_key = list(obj.keys())[0]
        value = obj[first_key]
        if "_Value" in value:
            return value["_Value"]
        for k, v in value.items():
            value[k] = minify_nested_serial(v)

    if isinstance(obj, list):
        for i in range(len(obj)):
            obj[i] = minify_nested_serial(obj[i])

    return obj


# 移除枚举值前缀，只保留枚举名
def remove_enum_value(text: object) -> str | object:
    if isinstance(text, list):
        return [remove_enum_value(t) for t in text]
    if not isinstance(text, str):
        return text
    match = re_enum_value.match(text)
    if match:
        return match.group(2)
    return text


# 将枚举字符串拆分成枚举值和枚举名
def seperate_enum_value(text: str) -> tuple[int, str] | None:
    if not isinstance(text, str):
        return None
    match = re_enum_value.match(text)
    if match:
        return int(match.group(1)), match.group(2)
    return None


def rare_enum_to_value(text: str) -> str | int:
    match = re_rare.match(text)
    if match:
        return int(match.group(1)) + 1
    return text


# 将指定的一个或多个列排序到指定的列之后
def reindex_column(
    df: pd.DataFrame,
    column: str | list[str],
    next_to: str = "",
    to_end: bool = False,
):
    if next_to == "" and not to_end:
        raise ValueError("Either next_to or to_end must be specified")
    if to_end:
        next_to = df.columns[-1]

    if isinstance(column, str):
        column = [column]
    for c in column:
        if c not in df.columns:
            raise ValueError(f"Column {c} not found in DataFrame")
    if next_to not in df.columns:
        raise ValueError(f"Column {next_to} not found in DataFrame")

    new_columns = df.columns.tolist()
    column.reverse()
    for c in column:
        new_columns.remove(c)
        new_columns.insert(new_columns.index(next_to) + 1, c)
    return df.reindex(columns=new_columns)


# 将指定的一个或多个行排序到指定的行之后
def reindex_row(
    df: pd.DataFrame,
    row: str | list[str],
    next_to: str = "",
    to_end: bool = False,
):
    if next_to == "" and not to_end:
        raise ValueError("Either next_to or to_end must be specified")
    if to_end:
        next_to = df.index[-1]

    if isinstance(row, str):
        row = [row]
    for r in row:
        if r not in df.index:
            raise ValueError(f"Row {r} not found in DataFrame")
    if next_to not in df.index:
        raise ValueError(f"Row {next_to} not found in DataFrame")

    new_index = df.index.tolist()
    row.reverse()
    for r in row:
        new_index.remove(r)
        new_index.insert(new_index.index(next_to) + 1, r)
    return df.reindex(index=new_index)
