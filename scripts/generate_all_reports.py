import json
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

DATA_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "evaluation_data.json")
if not os.path.exists(DATA_PATH):
    DATA_PATH = os.path.join(os.path.dirname(__file__), "..", "evaluation_data.json")

with open(DATA_PATH, "r", encoding="utf-8") as f:
    all_entries = json.load(f)

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "reports_excel")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Unified, restrained Sage & Stone Colors
COLOR_SAGE_HEADER = "E4ECD3"
COLOR_STONE_HEADER = "F2EEE6"
COLOR_WHITE = "FFFFFF"

thin_side = Side(border_style="thin", color="D5CEC5")
thin_border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)

font_main_title = Font(name="微軟正黑體", size=13, bold=True, color="2E2827")
font_chunk_title = Font(name="微軟正黑體", size=11, bold=True, color="2E2827")
font_sub_header = Font(name="微軟正黑體", size=10, bold=True, color="2E2827")
font_body = Font(name="微軟正黑體", size=9.5, color="2E2827")
font_body_bold = Font(name="微軟正黑體", size=9.5, bold=True, color="2E2827")

align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="top", wrap_text=True)
align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)

# OFFICIAL ORG & TITLES
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

SUPERVISOR_TEAMS = {
    "張希慈": ["何維安", "陳泳璇", "張芳媐", "姚品瑄", "胡喻翔"],
    "何維安": ["林文琇"],
    "姚品瑄": ["薛筑瑄", "戴佑珍"],
    "張希慈_執行長": ["張希慈"]
}

MEMBER_SUPERVISOR_MAP = {
    "何維安": "張希慈", "陳泳璇": "張希慈", "張芳媐": "張希慈", "姚品瑄": "張希慈", "胡喻翔": "張希慈",
    "林文琇": "何維安", "薛筑瑄": "姚品瑄", "戴佑珍": "姚品瑄", "張希慈": "董事會"
}

ALL_MEMBERS = ["何維安", "陳泳璇", "張芳媐", "姚品瑄", "胡喻翔", "林文琇", "薛筑瑄", "戴佑珍", "張希慈"]

ROLE_COMPETENCIES = {
    "執行長": ["策略決策與組織方向設定", "組織治理與財務風險管理", "關鍵利害關係人關係建立與維繫", "公關、演講與媒體關係", "核心團隊培養與組織文化建立", "其他創造的好事與預防的壞事"],
    "行政經理": ["人事薪資與人資系統管理", "法規與政府公文管理", "董事會與治理作業執行", "財務核銷與內控執行", "總務採購與行政庶務管理", "其他創造的好事與預防的壞事"],
    "美感設計師": ["品牌識別系統設計與維護", "視覺與美感設計實務", "需求釐清與創意提案", "其他創造的好事與預防的壞事"],
    "專案經理": ["專案企劃與現場執行", "專案時程與預算規劃管理", "需求研究與方案迭代", "外部夥伴關係經營", "多元教學設計與現場引導", "其他創造的好事與預防的壞事"],
    "營運經理兼執行長特助": ["組織營運流程設計與優化", "制度文件與 SOP 建置", "人力資源策略與選用育留", "執行長幕僚協調與跨部門推動", "總務採購與行政庶務管理", "其他創造的好事與預防的壞事"],
    "品牌經理": ["品牌定位與外部溝通一致性", "行銷策略與議題倡議", "品牌活動策劃與策展敘事", "內部品牌管理與雇主品牌", "品牌危機與聲譽風險處理", "其他創造的好事與預防的壞事"],
    "部門儲備主管": ["部門策略規劃與專案組合管理", "專案經理管理與培育", "部門預算與資源配置管理", "跨領域專業掌握與推進", "利害關係人管理與衝突協調", "其他創造的好事與預防的壞事"]
}

def style_cell_range(ws, start_row, start_col, end_row, end_col, font=None, fill_hex=None, alignment=None):
    fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type="solid") if fill_hex else None
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=r, column=c)
            if font: cell.font = font
            if fill: cell.fill = fill
            cell.border = thin_border
            if alignment: cell.alignment = alignment

def add_subordinate_alignment_sheet(wb, member_name):
    ws = wb.create_sheet(title=f"【{member_name}】自評vs主管評")
    ws.views.sheetView[0].showGridLines = True
    job_role = JOB_ROLES_MAP.get(member_name, "專案經理")
    sup_name = MEMBER_SUPERVISOR_MAP.get(member_name, "主管")

    peer_records = [e for e in all_entries if e["relation"] == "同事" and e["target"] == member_name]
    num_peers = len(peer_records)

    self_entry = next((e for e in all_entries if e["relation"] == "自評" and e["target"] == member_name), None)
    has_self = bool(self_entry and self_entry.get("self_eval"))
    se = self_entry.get("self_eval") if has_self else None

    # Title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    ws.cell(row=1, column=1, value=f"好好星球文化基金會 360 年中成長評估 — 【{member_name}】部屬自評與主管評核對照表（主管專用）")
    style_cell_range(ws, 1, 1, 1, 7, font_main_title, COLOR_STONE_HEADER, Alignment(horizontal="left", vertical="center", indent=1))
    ws.row_dimensions[1].height = 28

    # Metadata
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=7)
    ws.cell(row=2, column=1, value=f"評估對象：{member_name}（${job_role}） ｜ 直屬主管：{sup_name} ｜ 同儕樣本：{num_peers} 位 ｜ 自評狀態：{'已完成自評' if has_self else '尚未自評'}")
    style_cell_range(ws, 2, 1, 2, 7, font_sub_header, COLOR_WHITE, Alignment(horizontal="left", vertical="center", indent=1))
    ws.row_dimensions[2].height = 22

    r = 3
    # 1. 組織文化 (一列一個面向)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    ws.cell(row=r, column=1, value="一、組織文化實踐：部屬自評實例 vs 主管評核回饋（一列一個面向）")
    style_cell_range(ws, r, 1, r, 7, font_sub_header, COLOR_STONE_HEADER, Alignment(horizontal="left", vertical="center", indent=1))
    ws.row_dimensions[r].height = 22
    r += 1

    cult_headers = ["文化面向", "文化定義與行為指引", "部屬自評實例（STAR）", "自評等級", "主管評定", "落差分析", "主管評核回饋與觀察"]
    for i, h in enumerate(cult_headers):
        ws.cell(row=r, column=i+1, value=h)
    style_cell_range(ws, r, 1, r, 7, font_sub_header, COLOR_STONE_HEADER, align_header)
    ws.row_dimensions[r].height = 22
    r += 1

    cult_rows = [
        ("【信任】", "獨立行動與決策、主動協作、雙向溝通，在高彈性下給予彼此信任", se.get("values", {}).get("信任", "（部屬未填寫）") if has_self else "（部屬未填寫）"),
        ("【多元】", "尊重差異、多元工作方法、主動表達不同觀點與想法", se.get("values", {}).get("多元", "（部屬未填寫）") if has_self else "（部屬未填寫）"),
        ("【實驗】", "透過開放心態嘗試修正與反思，勇於檢討及給予回饋", se.get("values", {}).get("實驗", "（部屬未填寫）") if has_self else "（部屬未填寫）"),
        ("【可持續】", "內在韌性、自我照顧、彈性的人際與工作邊界", se.get("values", {}).get("可持續", "（部屬未填寫）") if has_self else "（部屬未填寫）"),
    ]

    for c_title, c_desc, c_self in cult_rows:
        ws.cell(row=r, column=1, value=c_title)
        ws.cell(row=r, column=2, value=c_desc)
        ws.cell(row=r, column=3, value=c_self)
        ws.cell(row=r, column=4, value="不適用")
        ws.cell(row=r, column=5, value="不適用")
        ws.cell(row=r, column=6, value="質化對齊")
        ws.cell(row=r, column=7, value="【待主管填寫回饋】")

        style_cell_range(ws, r, 1, r, 7, font_body, COLOR_WHITE, align_left)
        ws.cell(row=r, column=1).alignment = align_center
        ws.cell(row=r, column=1).font = font_body_bold
        ws.cell(row=r, column=4).alignment = align_center
        ws.cell(row=r, column=5).alignment = align_center
        ws.cell(row=r, column=6).alignment = align_center
        ws.row_dimensions[r].height = max(24, min(100, len(c_self or "") // 35 * 16 + 18))
        r += 1

    r += 1
    # 2. 專業職能 (一列一個面向，自評 vs 主管評並列)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    ws.cell(row=r, column=1, value=f"二、專業職能：自評 vs 主管評分並列對照【{job_role}】（逐項比對認知差異）")
    style_cell_range(ws, r, 1, r, 7, font_sub_header, COLOR_STONE_HEADER, Alignment(horizontal="left", vertical="center", indent=1))
    ws.row_dimensions[r].height = 22
    r += 1

    comp_headers = ["職能項目", "職能定義與說明", "部屬自評實例（STAR）", "部屬自評（Lv. 1～5）", "主管評定（Lv. 1～5）", "落差分析", "主管評語與回饋（Feedback）"]
    for i, h in enumerate(comp_headers):
        ws.cell(row=r, column=i+1, value=h)
    style_cell_range(ws, r, 1, r, 7, font_sub_header, COLOR_STONE_HEADER, align_header)
    ws.row_dimensions[r].height = 22
    r += 1

    dv = DataValidation(type="list", formula1='"L1,L2,L3,L4,L5"', allow_blank=True)
    dv.error ='請從選單選擇等級 L1～L5'
    dv.errorTitle = '評分無效'
    dv.prompt = '請選取：L1（Start）、L2（Grow）、L3（Keep）、L4（Good）、L5（Amazing！）'
    dv.promptTitle = '等級選單'
    ws.add_data_validation(dv)

    comp_titles = ROLE_COMPETENCIES.get(job_role, [])
    for c_title in comp_titles:
        self_ans = "（部屬未填寫自評實例）"
        if has_self and se.get("competencies"):
            item = next((x for x in se["competencies"] if x.get("title") == c_title), None)
            if item and item.get("answer"):
                self_ans = item["answer"]

        ws.cell(row=r, column=1, value=c_title)
        ws.cell(row=r, column=2, value="核心專業職能")
        ws.cell(row=r, column=3, value=self_ans)
        ws.cell(row=r, column=4, value="待評")
        ws.cell(row=r, column=5, value="")
        ws.cell(row=r, column=6, value="待主管評")
        ws.cell(row=r, column=7, value="【待主管填寫回饋】")

        style_cell_range(ws, r, 1, r, 7, font_body, COLOR_WHITE, align_left)
        ws.cell(row=r, column=1).font = font_body_bold
        ws.cell(row=r, column=4).alignment = align_center
        ws.cell(row=r, column=5).alignment = align_center
        ws.cell(row=r, column=5).font = font_body_bold
        ws.cell(row=r, column=6).alignment = align_center

        dv.add(ws.cell(row=r, column=4))
        dv.add(ws.cell(row=r, column=5))

        ws.row_dimensions[r].height = max(24, min(100, len(self_ans or "") // 35 * 16 + 18))
        r += 1

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 55
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 35

# Build Supervisor Workbook
for sup_name, members in SUPERVISOR_TEAMS.items():
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default sheet
    for m in members:
        add_subordinate_alignment_sheet(wb, m)
    fname = f"【{sup_name}主管專用】部屬自評vs主管評對照彙整包.xlsx"
    fpath = os.path.join(OUTPUT_DIR, fname)
    wb.save(fpath)
    print(f"Generated {fpath}")

# Build Full Organization Workbook
wb_full = openpyxl.Workbook()
wb_full.remove(wb_full.active)
for m in ALL_MEMBERS:
    add_subordinate_alignment_sheet(wb_full, m)
full_path = os.path.join(OUTPUT_DIR, "好好星球_360年中成長評估_主管分流與完整彙整表.xlsx")
wb_full.save(full_path)
print(f"Generated {full_path}")
