import json

with open("data/evaluation_data.json", "r", encoding="utf-8") as f:
    data = json.load(f)

# Check all members
for m in ["何維安", "陳泳璇", "張芳媐", "姚品瑄", "胡喻翔", "林文琇", "薛筑瑄", "戴佑珍", "張希慈"]:
    peer_records = [e for e in data if e["relation"] == "同事" and e["target"] == m]
    self_record = [e for e in data if e["relation"] == "自評" and e["target"] == m]
    print(f"Member: {m:<6} | Peer Reviews: {len(peer_records)} | Has Self-Eval: {len(self_record) > 0}")

