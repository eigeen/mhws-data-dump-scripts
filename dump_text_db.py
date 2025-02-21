import json
import os
import re

re_xref_id = re.compile(r"(<[Rr][Ee][Ff] (.*?)>)")
re_xref_em_id = re.compile(r"(<EMID (.*?)>)")


def dump_text_db_json(input_dir: str) -> dict[str, dict]:
    db_entries = {}

    for root, dirs, files in os.walk(input_dir):
        for file in files:
            if not file.endswith(".msg.23.json"):
                continue

            data = None
            with open(os.path.join(root, file), "r", encoding="utf-8") as f:
                data = f.read()

            data = json.loads(data)
            for entry in data["entries"]:
                if entry["content"] == None:
                    continue
                contents = content_list_to_dict(entry["content"])
                db_entry = {
                    "name": entry["name"],
                    "guid": entry["guid"],
                    "belongs_to": file,
                    "contents": contents,
                }
                db_entries[entry["name"]] = db_entry

    db_entries = process_xref(db_entries)
    return list(db_entries.values())


def content_list_to_dict(contents: list) -> dict:
    content_dict = {}
    for i, content in enumerate(contents):
        if not content:
            continue
        content_dict[i] = content
    return content_dict


def process_xref(db_entries: dict[str, dict]) -> dict[str, dict]:
    for entry in db_entries.values():
        new_entry = replace_xref_tag_in_entry(entry, db_entries)
        db_entries[entry["name"]] = new_entry
    return db_entries


def replace_xref_tag_in_entry(entry: dict, db_entries: dict[str, dict]) -> dict:
    for i, content in entry["contents"].items():
        content = replace_xref_tag_in_content(content, i, db_entries)
        entry["contents"][i] = content
    return entry


def replace_xref_tag_in_content(
    content: str, lang_id: int, db_entries: dict[str, dict]
) -> str:
    line = content
    xrefs = re_xref_id.findall(line)
    for xref in xrefs:
        xref_tag = xref[0]
        xref_name = xref[1]
        xref_entry = db_entries.get(xref_name)
        if xref_entry:
            line = line.replace(xref_tag, xref_entry["contents"][lang_id], 1)
            line = replace_xref_tag_in_content(line, lang_id, db_entries)
    xrefs = re_xref_em_id.findall(line)
    for xref in xrefs:
        xref_tag = xref[0]
        xref_name = xref[1]
        tag = f"EnemyText_NAME_{xref_name}"
        xref_entry = db_entries.get(tag)
        if xref_entry:
            line = line.replace(xref_tag, xref_entry["contents"][lang_id], 1)
            line = replace_xref_tag_in_content(line, lang_id, db_entries)
    return line


output = dump_text_db_json("natives")
with open("texts_db.json", "w", encoding="utf-8") as f:
    json.dump(output, f, ensure_ascii=False, indent=4)
