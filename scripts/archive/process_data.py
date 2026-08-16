import csv
import json

# Supervisor hierarchy
SUPERVISOR_TEAMS = {
    "張希慈": ["何維安", "陳泳璇", "張芳媐", "姚品瑄", "胡喻翔"],
    "何維安": ["林文琇"],
    "姚品瑄": ["薛筑瑄", "戴佑珍"]
}

# Role definitions for self-evaluation columns
JOB_BLOCKS = {
    "執行長": {
        "competencies": [
            (59, "策略決策與組織方向設定"),
            (60, "組織治理與財務風險管理"),
            (61, "關鍵利害關係人關係建立與維繫"),
            (62, "公關、演講與媒體關係"),
            (63, "核心團隊培養與組織文化建立"),
            (64, "其他創造的好事與預防的壞事")
        ],
        "reflection": [
            (65, "未來一年想創造的價值"),
            (66, "最常感到卡關/掙扎的階段"),
            (67, "卡關具體原因描述"),
            (68, "希望組織當時提供的幫助"),
            (69, "對願景使命理解與實踐的最大轉變")
        ]
    },
    "行政經理": {
        "competencies": [
            (71, "人事薪資與人資系統管理"),
            (72, "法規與政府公文管理"),
            (73, "董事會與治理作業執行"),
            (74, "財務核銷與內控執行"),
            (75, "總務採購與行政庶務管理"),
            (76, "其他創造的好事與預防的壞事")
        ],
        "reflection": [
            (77, "未來一年想創造的價值"),
            (78, "最常感到卡關/掙扎的階段"),
            (79, "卡關具體原因描述"),
            (80, "希望組織當時提供的幫助"),
            (81, "對願景使命理解與實踐的最大轉變")
        ]
    },
    "營運經理": {
        "competencies": [
            (83, "組織營運流程設計與優化"),
            (84, "制度文件與 SOP 建置"),
            (85, "人力資源策略與選用育留"),
            (86, "組織文化落實與制度轉化"),
            (87, "總務採購與行政庶務管理"),
            (88, "其他創造的好事與預防的壞事")
        ],
        "reflection": [
            (89, "未來一年想創造的價值"),
            (90, "最常感到卡關/掙扎的階段"),
            (91, "卡關具體原因描述"),
            (92, "希望組織當時提供的幫助"),
            (93, "對願景使命理解與實踐的最大轉變")
        ]
    },
    "專案部門儲備主管": {
        "competencies": [
            (95, "部門策略規劃與專案組合管理"),
            (96, "專案經理管理與培育"),
            (97, "部門預算與資源配置管理"),
            (98, "跨 LAB 專業掌握（CGL、SL、SPL）"),
            (99, "利害關係人管理與衝突協調"),
            (100, "其他創造的好事與預防的壞事")
        ],
        "reflection": [
            (101, "未來一年想創造的價值"),
            (102, "最常感到卡關/掙扎的階段"),
            (103, "卡關具體原因描述"),
            (104, "希望組織當時提供的幫助"),
            (105, "對願景使命理解與實踐的最大轉變")
        ]
    },
    "CGL專案經理": {
        "competencies": [
            (107, "專案企劃與現場執行"),
            (108, "專案時程與預算規劃管理"),
            (109, "需求研究與方案迭代"),
            (110, "外部夥伴關係經營"),
            (111, "多元教學設計與現場引導"),
            (112, "其他創造的好事與預防的壞事")
        ],
        "reflection": [
            (113, "未來一年想創造的價值"),
            (114, "最常感到卡關/掙扎的階段"),
            (115, "卡關具體原因描述"),
            (116, "希望組織當時提供的幫助"),
            (117, "對願景使命理解與實踐的最大轉變")
        ]
    },
    "SL專案經理": {
        "competencies": [
            (119, "專案企劃與現場執行"),
            (120, "專案時程與預算規劃管理"),
            (121, "需求研究與方案迭代"),
            (122, "外部夥伴關係經營"),
            (123, "社群經營與培力需求回應"),
            (124, "其他創造的好事與預防的壞事")
        ],
        "reflection": [
            (125, "未來一年想創造的價值"),
            (126, "最常感到卡關/掙扎的階段"),
            (127, "卡關具體原因描述"),
            (128, "希望組織當時提供的幫助"),
            (129, "對願景使命理解與實踐的最大轉變")
        ]
    },
    "品牌經理": {
        "competencies": [
            (131, "品牌定位與外部溝通一致性"),
            (132, "行銷策略與議題倡議"),
            (133, "品牌活動策劃與策展敘事"),
            (134, "內部品牌管理與雇主品牌"),
            (135, "品牌危機與聲譽風險處理"),
            (136, "其他創造的好事與預防的壞事")
        ],
        "reflection": [
            (137, "未來一年想創造的價值"),
            (138, "最常感到卡關/掙扎的階段"),
            (139, "卡關具體原因描述"),
            (140, "希望組織當時提供的幫助"),
            (141, "對願景使命理解與實踐的最大轉變")
        ]
    },
    "視覺設計師": {
        "competencies": [
            (143, "品牌識別系統設計與維護"),
            (144, "視覺設計實務"),
            (145, "需求釐清與創意提案"),
            (146, "其他創造的好事與預防的壞事")
        ],
        "reflection": [
            (147, "未來一年想創造的價值"),
            (148, "最常感到卡關/掙扎的階段"),
            (149, "卡關具體原因描述"),
            (150, "希望組織當時提供的幫助"),
            (151, "對願景使命理解與實踐的最大轉變")
        ]
    }
}

def parse_csv_data(filepath):
    with open(filepath, "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        header = next(reader)
        raw_rows = list(reader)

    entries = []
    for row in raw_rows:
        timestamp = row[0].strip()
        email = row[1].strip()
        target = row[2].strip()
        relation = row[3].strip()
        job_role = row[58].strip() if len(row) > 58 else ""
        
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
                val = row[c].strip()
                scores[header[c]] = float(val) if val else None
            
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
                "q20_vision_mission": row[20].strip() if len(row) > 20 else "",
                "q21_improvement_advice": row[21].strip() if len(row) > 21 else "",
                "q22_other_comments": row[22].strip() if len(row) > 22 else "",
                "q23_starlight_thanks": row[23].strip() if len(row) > 23 else ""
            }
        elif relation == "同事":
            scores = {}
            for c in range(24, 37):
                val = row[c].strip()
                scores[header[c]] = float(val) if val else None
                
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
                "q37_improvement_advice": row[37].strip() if len(row) > 37 else "",
                "q38_other_comments": row[38].strip() if len(row) > 38 else "",
                "q39_starlight_thanks": row[39].strip() if len(row) > 39 else ""
            }
        elif relation == "自評":
            top3_stable = [s.strip() for s in row[52].split(",") if s.strip()] if len(row) > 52 else []
            top3_practice = [s.strip() for s in row[53].split(",") if s.strip()] if len(row) > 53 else []
            
            trust_text = row[54].strip() if len(row) > 54 else ""
            diversity_text = row[55].strip() if len(row) > 55 else ""
            experiment_text = row[56].strip() if len(row) > 56 else ""
            sustainability_text = row[57].strip() if len(row) > 57 else ""
            
            competencies = []
            reflection = {}
            
            job_def = JOB_BLOCKS.get(job_role)
            if job_def:
                for idx, title in job_def["competencies"]:
                    ans = row[idx].strip() if len(row) > idx else ""
                    competencies.append({"title": title, "answer": ans})
                for idx, title in job_def["reflection"]:
                    ans = row[idx].strip() if len(row) > idx else ""
                    reflection[title] = ans
                    
            entry["self_eval"] = {
                "job_role": job_role,
                "top3_stable": top3_stable,
                "top3_practice": top3_practice,
                "values": {
                    "信任": trust_text,
                    "多元": diversity_text,
                    "實驗": experiment_text,
                    "可持續": sustainability_text
                },
                "competencies": competencies,
                "reflection": reflection
            }
            
        entries.append(entry)

    with open("evaluation_data.json", "w", encoding="utf-8") as f:
        json.dump(entries, f, ensure_ascii=False, indent=2)

    print("Saved evaluation_data.json with", len(entries), "entries.")
    return entries

if __name__ == "__main__":
    parse_csv_data("好好星球文化基金會內部 360 年中成長評估 (回覆) - 表單回覆 1.csv")
