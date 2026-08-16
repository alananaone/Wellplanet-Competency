import urllib.request
import json
import csv
import os

GAS_URL = "https://script.google.com/macros/s/AKfycbxM-5YB3AX_CRK6APM3-dxGPUK7A2anQLrWSRwDK0_cZubdUu3pcUSl9lTPy5ahxXytgg/exec"
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CSV_PATH = os.path.join(BASE_DIR, "data", "好好星球文化基金會內部 360 年中成長評估 (回覆) - 表單回覆 1.csv")
JSON_DATA_PATH = os.path.join(BASE_DIR, "data", "evaluation_data.json")
ROOT_JSON_PATH = os.path.join(BASE_DIR, "evaluation_data.json")

JOB_ROLES_MAP = {
    "張希慈": "執行長",
    "陳泳璇": "行政經理",
    "林文琇": "美感設計師",
    "胡喻翔": "專案經理",
    "張芳媐": "營運經理兼執行長特助",
    "何維安": "品牌經理",
    "姚品瑄": "部門儲備主管",
    "薛筑瑄": "專案經理",
    "戴佑珍": "專案經理"
}

def sync_data():
    print(f"Connecting to Google Apps Script Web App:\n{GAS_URL}")
    req = urllib.request.Request(GAS_URL, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            raw_rows = json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        print(f"Failed to fetch from GAS: {e}")
        return False

    if not raw_rows or len(raw_rows) < 2:
        print("Data from GAS is empty or invalid.")
        return False

    header = raw_rows[0]
    data_rows = raw_rows[1:]

    # Save to CSV
    os.makedirs(os.path.dirname(CSV_PATH), exist_ok=True)
    with open(CSV_PATH, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(raw_rows)
    print(f"Saved {len(raw_rows)} rows to CSV: {CSV_PATH}")

    entries = []
    for row in data_rows:
        if len(row) < 4: continue
        timestamp = str(row[0] or "").strip()
        email = str(row[1] or "").strip()
        target = str(row[2] or "").strip()
        relation = str(row[3] or "").strip()
        job_role = JOB_ROLES_MAP.get(target, str(row[58] if len(row) > 58 else "").strip())

        entry = {
            "timestamp": timestamp,
            "email": email,
            "target": target,
            "relation": relation,
            "job_role": job_role
        }

        if relation == "主管":
            scores = {}
            for c in range(4, 20):
                val = str(row[c] if len(row) > c else "").strip()
                try:
                    scores[header[c]] = float(val) if val else None
                except:
                    scores[header[c]] = None
            
            entry["supervisor_eval"] = {
                "q_scores": scores,
                "q4_help_easy": scores.get(header[4]),
                "q5_guidance_freq": scores.get(header[5]),
                "q6_improve_degree": scores.get(header[6]),
                "q7_cross_dept": scores.get(header[7]),
                "q8_resource_eval": scores.get(header[8]),
                "q9_constructive_mistake": scores.get(header[9]),
                "q10_recognition": scores.get(header[10]),
                "q11_overall_performance": scores.get(header[11]),
                "q12_trust_express": scores.get(header[12]),
                "q13_diversity_listen": scores.get(header[13]),
                "q14_experiment_try": scores.get(header[14]),
                "q15_experiment_psych_safety": scores.get(header[15]),
                "q16_sustain_praise": scores.get(header[16]),
                "q17_sustain_boundary": scores.get(header[17]),
                "q18_nps_recommend": scores.get(header[18]),
                "q19_satisfaction": scores.get(header[19]),
                "q20_vision_mission": str(row[20] if len(row) > 20 else "").strip(),
                "q21_improvement_advice": str(row[21] if len(row) > 21 else "").strip(),
                "q22_other_comments": str(row[22] if len(row) > 22 else "").strip(),
                "q23_starlight_thanks": str(row[23] if len(row) > 23 else "").strip()
            }
        elif relation == "同事":
            scores = {}
            for c in range(24, 37):
                val = str(row[c] if len(row) > c else "").strip()
                try:
                    scores[header[c]] = float(val) if val else None
                except:
                    scores[header[c]] = None
            
            entry["peer_eval"] = {
                "q_scores": scores,
                "q24_cooperation": scores.get(header[24]),
                "q25_detail_oriented": scores.get(header[25]),
                "q26_on_time": scores.get(header[26]),
                "q27_flexibility": scores.get(header[27]),
                "q28_follow_up": scores.get(header[28]),
                "q29_transparency": scores.get(header[29]),
                "q30_open_to_opposing": scores.get(header[30]),
                "q31_constructive_opinions": scores.get(header[31]),
                "q32_growth_mindset": scores.get(header[32]),
                "q33_share_knowledge": scores.get(header[33]),
                "q34_praise_peers": scores.get(header[34]),
                "q35_boundary_respect": scores.get(header[35]),
                "q36_nps_recommend": scores.get(header[36]),
                "q37_improvement_advice": str(row[37] if len(row) > 37 else "").strip(),
                "q38_other_comments": str(row[38] if len(row) > 38 else "").strip(),
                "q39_starlight_thanks": str(row[39] if len(row) > 39 else "").strip()
            }
        elif relation == "自評":
            top3_stable = [s.strip() for s in str(row[52] if len(row) > 52 else "").split(",") if s.strip()]
            top3_practice = [s.strip() for s in str(row[53] if len(row) > 53 else "").split(",") if s.strip()]
            values = {
                "信任": str(row[54] if len(row) > 54 else "").strip(),
                "多元": str(row[55] if len(row) > 55 else "").strip(),
                "實驗": str(row[56] if len(row) > 56 else "").strip(),
                "可持續": str(row[57] if len(row) > 57 else "").strip()
            }

            competencies = []
            reflection = {}

            for c in range(59, len(row)):
                val = str(row[c]).strip()
                if not val: continue
                col_name = header[c].strip()
                if "在未來的一年中" in col_name or "卡關" in col_name or "願景和使命" in col_name:
                    reflection[col_name] = val
                elif "記得也要花" in col_name or "下一階段進行" in col_name:
                    continue
                else:
                    competencies.append({"title": col_name, "answer": val})

            entry["self_eval"] = {
                "job_role": job_role,
                "top3_stable": top3_stable,
                "top3_practice": top3_practice,
                "values": values,
                "competencies": competencies,
                "reflection": reflection
            }

        entries.append(entry)

    with open(JSON_DATA_PATH, "w", encoding="utf-8") as f:
        json.dump(entries, f, ensure_ascii=False, indent=2)

    with open(ROOT_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(entries, f, ensure_ascii=False, indent=2)

    print(f"Successfully processed and saved {len(entries)} parsed entries to JSON files!")
    return True

if __name__ == "__main__":
    sync_data()
