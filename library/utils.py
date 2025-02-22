import re

re_guid_like = re.compile(
    r"\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b"
)


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
