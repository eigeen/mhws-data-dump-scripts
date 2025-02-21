import json
import os
import re

from library.text_db import load_text_db

text_db = load_text_db("texts_db.json")
re_xref_id = re.compile(r"<REF (.*?)>")
re_xref_em_id = re.compile(r"<EMID (.*?)>")


def dump_text_txt(data: str) -> str:
    data_json = json.loads(data)
    text = ""
    for entry in data_json["entries"]:
        line = f"{entry["content"][13]}"

        # line = line.replace("\n", "\\n").replace("\r", "\\r")
        line = line.replace("\n", "").replace("\r", "")
        text += line + "\n"
    return text


out_dir = "./texts"
if not os.path.exists(out_dir):
    os.makedirs(out_dir)

# walk dir to search all *.msg.23.json files
for root, dirs, files in os.walk("natives"):
    for file in files:
        if not file.endswith(".msg.23.json"):
            continue

        data = None
        text = None

        with open(os.path.join(root, file), "r", encoding="utf-8") as f:
            data = f.read()

        print(f"Processing {file}...")
        try:
            text = dump_text_txt(data)
        except Exception as e:
            print(f"Error while processing {file}: {e}\n")
            continue

        out_file_path = os.path.join(
            out_dir, root, file.replace(".msg.23.json", ".txt")
        )
        if not os.path.exists(os.path.dirname(out_file_path)):
            os.makedirs(os.path.dirname(out_file_path))
        with open(
            out_file_path,
            "w",
            encoding="utf-8",
        ) as f2:
            f2.write(text)
