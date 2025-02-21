import re

re_guid_like = re.compile(
    r"\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b"
)


def is_guid_like(text: str) -> bool:
    return re_guid_like.match(text) is not None
