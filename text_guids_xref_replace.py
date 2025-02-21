import os
import re

from library.text_db import load_text_db

text_db = load_text_db("texts_db.json")

input_dir = "natives"
output_dir = "natives_xref_replaced"

re_guid_like = re.compile(
    r"\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b"
)

for root, dirs, files in os.walk(input_dir):
    for file in files:
        if file.endswith(".msg.23.json"):
            continue

        output_file_path = os.path.join(output_dir, root, file)
        # if os.path.exists(output_file_path):
        #     continue

        file_path = os.path.join(root, file)
        print(f"Processing {file_path}")
        content = ""
        with open(file_path, "r", encoding="utf-8") as f:
            content = f.read()

        changed = False
        for guid in re_guid_like.findall(content):
            text = text_db.get_text_by_guid(guid)
            if text:
                text = text.replace("\r", "").replace("\n", "")
                content = content.replace(guid, text, 1)
                changed = True

        if not changed:
            continue

        if not os.path.exists(os.path.dirname(output_file_path)):
            os.makedirs(os.path.dirname(output_file_path))
        with open(output_file_path, "w", encoding="utf-8") as f:
            f.write(content)
