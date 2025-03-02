import json

inputs = ["re_chunk_000_1.0.2.0.json", "re_chunk_000_1.0.1.0.json"]
output = "diff.json"

print("Loading data...")
entry_datas = []
for input in inputs:
    with open(input, "r") as f:
        data = json.load(f)
        data_processed = {}
        for entry in data["entries"]:
            data_path_hash = (
                entry["entry"]["hash_name_upper"] + entry["entry"]["hash_name_lower"]
            )
            data_processed[data_path_hash] = entry
        entry_datas.append(data_processed)

print("Comparing data...")
# 相同文件差异
diff = {"missing": [], "extra": [], "modified": []}
# 从第一个开始对比
for path_hash, entry in entry_datas[0].items():
    for other_data in entry_datas[1:]:
        other_entry = other_data.get(path_hash)
        if not other_entry:
            diff["extra"].append(entry)
        else:
            if entry["entry"]["checksum"] != other_entry["entry"]["checksum"]:
                diff["modified"].append(entry)

print(
    f"missing: {len(diff['missing'])}, extra: {len(diff['extra'])}, modified: {len(diff['modified'])}"
)

with open(output, "w") as f:
    json.dump(diff, f, indent=4)
# 从第二个开始对比新增文件
