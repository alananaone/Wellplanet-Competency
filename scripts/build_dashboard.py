import json
import os

DATA_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "evaluation_data.json")
if not os.path.exists(DATA_PATH):
    DATA_PATH = os.path.join(os.path.dirname(__file__), "..", "evaluation_data.json")

with open(DATA_PATH, "r", encoding="utf-8") as f:
    raw_data_json = f.read()

# Make sure root evaluation_data.json is in sync
root_data_path = os.path.join(os.path.dirname(__file__), "..", "evaluation_data.json")
with open(root_data_path, "w", encoding="utf-8") as f:
    f.write(raw_data_json)

html_content = r"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>好好星球文化基金會 360 年中成長評估系統</title>
  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- Chart.js CDN -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <!-- PapaParse for client-side CSV parsing -->
  <script src="https://cdn.jsdelivr.net/npm/papaparse@5.4.1/papaparse.min.js"></script>
  <!-- ExcelJS for full color/style Excel generation in browser -->
  <script src="https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"></script>
  <!-- FileSaver for reliable browser download -->
  <script src="https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js"></script>
  <!-- Lucide Icons -->
  <script src="https://unpkg.com/lucide@latest"></script>
  
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@500;600;700&family=Plus+Jakarta+Sans:wght@400;500;600;700&family=Noto+Sans+TC:wght@400;500;600;700&display=swap');
    
    :root {
      --bg-page: #FAF7F2;
      --bg-card: #FFFDF9;
      --bg-box: #F9F6EE;
      --color-primary: #557A61;
      --color-blush: #F4CCCC;
      --color-apricot: #FCE5CD;
      --color-border: #E8E2D8;
      --color-text-main: #2E2827;
    }

    body {
      font-family: 'Noto Sans TC', 'Plus Jakarta Sans', sans-serif;
      background-color: var(--bg-page);
      color: var(--color-text-main);
      letter-spacing: -0.01em;
      line-height: 1.75;
    }

    .font-serif-tc { font-family: 'Noto Serif TC', serif; }

    .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
    .custom-scrollbar::-webkit-scrollbar-thumb { background-color: #D9D2C9; border-radius: 4px; }

    .badge-stable { background-color: #E4ECD3; color: #2D5239; border: 1px solid #CDE0BC; }
    .badge-practice { background-color: #FCE5CD; color: #783E16; border: 1px solid #F5D3B3; }
    .badge-missing { background-color: #FFF2D6; color: #B45309; border: 1px solid #FCE299; }
    .badge-best { background-color: #E4ECD3; color: #2D5239; border: 1px solid #CDE0BC; }
    .badge-worst { background-color: #FCE5CD; color: #8C4B1E; border: 1px solid #F5D3B3; }

    .chunk-header-banner { background-color: var(--color-blush); color: #3E2426; border-bottom: 1px solid #E6BDBD; }
    .chunk-subheader-banner { background-color: var(--color-apricot); color: #4A2E1C; border-bottom: 1px solid #EED1B4; }
    .soft-card-shadow { box-shadow: 0 4px 20px -2px rgba(120, 100, 80, 0.06), 0 2px 6px -1px rgba(120, 100, 80, 0.04); }

    .screen-only-view { display: block; }
    .print-only-doc { display: none; }

    @keyframes spin-slow {
      from { transform: rotate(0deg); }
      to { transform: rotate(360deg); }
    }
    .animate-spin-custom { animation: spin-slow 1s linear infinite; }

    @media print {
      @page { size: A4 portrait; margin: 14mm 16mm 14mm 16mm; }
      body { background-color: #FFFFFF !important; color: #111111 !important; font-size: 9.5pt; line-height: 1.5; font-family: 'Times New Roman', 'Noto Serif TC', '微軟正黑體', serif !important; }
      header, #dropZone, nav, #toast, .no-print, button, .tab-btn, footer, .screen-only-view { display: none !important; }
      .print-only-doc { display: block !important; }
      main { padding: 0 !important; max-width: 100% !important; margin: 0 !important; }
      .tab-content { display: block !important; }
      .tab-content.hidden { display: none !important; }
      * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; box-shadow: none !important; border-radius: 0 !important; }
      .word-doc-table { width: 100% !important; border-collapse: collapse !important; margin-top: 8px !important; margin-bottom: 14px !important; font-size: 9pt !important; }
      .word-doc-table th, .word-doc-table td { border: 1px solid #555555 !important; padding: 5px 8px !important; vertical-align: middle !important; line-height: 1.4 !important; }
      .word-doc-table th { background-color: #F2EFE9 !important; color: #2E2827 !important; font-weight: bold !important; text-align: center !important; }
      .print-avoid-break, .word-doc-section, .word-feedback-block { break-inside: avoid !important; page-break-inside: avoid !important; }
      .word-sec-title { font-size: 11pt !important; font-weight: bold !important; color: #2E2827 !important; background-color: #EAE6DF !important; padding: 4px 8px !important; border-left: 4px solid #557A61 !important; margin-top: 16px !important; margin-bottom: 8px !important; }
      .word-doc-divider { border-top: 2px solid #557A61 !important; border-bottom: 1px solid #557A61 !important; height: 3px !important; margin: 10px 0 14px 0 !important; }
      .print-footer { display: block !important; text-align: center; font-size: 8pt; color: #777777; margin-top: 20px; border-top: 1px solid #CCCCCC; padding-top: 6px; }
    }
  </style>
</head>
<body class="min-h-screen flex flex-col antialiased">

  <!-- TOP HEADER (NO-PRINT) -->
  <header class="bg-[#FFFDF9]/95 backdrop-blur-md border-b border-[#E8E2D8] sticky top-0 z-50 shadow-xs no-print">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
      <div class="flex items-center justify-between h-18 gap-4">
        <!-- BRAND -->
        <div class="flex items-center gap-3.5">
          <div class="w-11 h-11 rounded-2xl bg-[#557A61] flex items-center justify-center text-white font-bold shadow-xs text-sm tracking-wider font-serif-tc">
            360
          </div>
          <div>
            <div class="flex items-center gap-2.5">
              <h1 class="text-base sm:text-lg font-bold text-[#2E2827] leading-tight font-serif-tc">好好星球文化基金會 360 年中成長評估</h1>
              <span class="px-2.5 py-0.5 text-xs font-semibold rounded-md bg-[#FCE5CD] text-[#783E16] border border-[#F3D1B0]" id="header-data-source-badge">
                優先連線 Google 試算表
              </span>
            </div>
            <p class="text-xs text-[#7A726D] mt-0.5">優先讀取試算表「表單回覆1」 ｜ 支援自評 vs 主管評比對與雙向 Excel 載入</p>
          </div>
        </div>

        <!-- ACTION BUTTONS -->
        <div class="flex items-center gap-2.5 sm:gap-3 flex-wrap justify-end">
          
          <!-- GAS SYNC BUTTON -->
          <button onclick="syncFromGoogleAppsScript(true)" id="gasSyncBtn" class="inline-flex items-center gap-2 px-3.5 sm:px-4 py-2.5 text-xs sm:text-sm font-semibold rounded-xl bg-[#E4ECD3] text-[#2D5239] hover:bg-[#D4DFC0] transition border border-[#CDE0BC] shadow-2xs" title="優先讀取 Google 試算表 (表單回覆1)">
            <i data-lucide="refresh-cw" class="w-4 h-4 text-[#557A61]" id="gasSyncIcon"></i>
            <span class="hidden md:inline" id="gasSyncText">同步 Google 試算表</span>
            <span class="md:hidden">同步</span>
          </button>

          <!-- PRINT BUTTON -->
          <button onclick="window.print()" class="inline-flex items-center gap-2 px-3.5 sm:px-4 py-2.5 text-xs sm:text-sm font-semibold rounded-xl bg-[#F2EEE6] text-[#2E2827] hover:bg-[#EBE4D8] transition border border-[#E0D7CA]">
            <i data-lucide="printer" class="w-4 h-4 text-[#557A61]"></i>
            <span class="hidden sm:inline">列印 / 存為 PDF (Cmd+P)</span>
            <span class="sm:hidden">列印</span>
          </button>

          <!-- UPLOAD LOCAL CSV BUTTON -->
          <label class="cursor-pointer inline-flex items-center gap-2 px-3.5 sm:px-4 py-2.5 text-xs sm:text-sm font-semibold rounded-xl bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] hover:text-[#2E2827] transition border border-[#E0D7CA]">
            <i data-lucide="upload" class="w-4 h-4 text-[#557A61]"></i>
            <span class="hidden lg:inline">上傳本地 CSV</span>
            <span class="lg:hidden">上傳</span>
            <input type="file" id="csvFileInput" accept=".csv" class="hidden" onchange="handleFileUpload(event)">
          </label>

          <!-- EXPORT DROPDOWN -->
          <div class="relative inline-block text-left" id="exportDropdown">
            <button onclick="toggleDropdown()" class="inline-flex items-center gap-2 px-4 sm:px-5 py-2.5 text-xs sm:text-sm font-semibold rounded-xl bg-[#557A61] text-white hover:bg-[#466551] transition shadow-xs">
              <i data-lucide="download" class="w-4 h-4"></i>
              <span class="hidden sm:inline">下載全彩 Excel 報表</span>
              <span class="sm:hidden">下載 Excel</span>
              <i data-lucide="chevron-down" class="w-4 h-4 opacity-80"></i>
            </button>
            
            <div id="dropdownMenu" class="hidden absolute right-0 mt-2 w-92 origin-top-right rounded-2xl bg-[#FFFDF9] p-3 shadow-2xl ring-1 ring-black/5 z-50 divide-y divide-[#EFEAE1] border border-[#E8E2D8]">
              <div class="py-2">
                <div class="px-3 py-1 text-[11px] font-bold text-[#8C837C] uppercase tracking-wider">
                  自評 vs 主管評對照包 (含認知差異與 L1~L5)
                </div>
                <button onclick="exportSupervisorTeamSubordinatesComprehensiveExcelClientSide('張希慈')" class="w-full text-left flex items-center gap-3 px-3 py-2 text-xs sm:text-sm text-[#2E2827] hover:bg-[#E4ECD3]/40 rounded-xl transition">
                  <div class="p-2 bg-[#E4ECD3] text-[#2D5239] rounded-xl"><i data-lucide="git-compare" class="w-4 h-4"></i></div>
                  <div>
                    <div class="font-bold text-[#2E2827]">【張希慈】部屬自評 vs 主管評對照包</div>
                    <div class="text-[11px] text-[#7A726D]">何維安、陳泳璇、張芳媐、姚品瑄、胡喻翔</div>
                  </div>
                </button>
                <button onclick="exportSupervisorTeamSubordinatesComprehensiveExcelClientSide('何維安')" class="w-full text-left flex items-center gap-3 px-3 py-2 text-xs sm:text-sm text-[#2E2827] hover:bg-[#E4ECD3]/40 rounded-xl transition">
                  <div class="p-2 bg-[#E4ECD3] text-[#2D5239] rounded-xl"><i data-lucide="git-compare" class="w-4 h-4"></i></div>
                  <div>
                    <div class="font-bold text-[#2E2827]">【何維安】部屬自評 vs 主管評對照包</div>
                    <div class="text-[11px] text-[#7A726D]">林文琇（美感設計師）</div>
                  </div>
                </button>
                <button onclick="exportSupervisorTeamSubordinatesComprehensiveExcelClientSide('姚品瑄')" class="w-full text-left flex items-center gap-3 px-3 py-2 text-xs sm:text-sm text-[#2E2827] hover:bg-[#E4ECD3]/40 rounded-xl transition">
                  <div class="p-2 bg-[#E4ECD3] text-[#2D5239] rounded-xl"><i data-lucide="git-compare" class="w-4 h-4"></i></div>
                  <div>
                    <div class="font-bold text-[#2E2827]">【姚品瑄】部屬自評 vs 主管評對照包</div>
                    <div class="text-[11px] text-[#7A726D]">薛筑瑄、戴佑珍（專案經理）</div>
                  </div>
                </button>
              </div>

              <div class="py-2">
                <button onclick="exportFullWorkbookClientSide()" class="w-full text-left flex items-center gap-3 px-3 py-2 text-xs sm:text-sm font-bold text-[#557A61] hover:bg-[#E4ECD3]/40 rounded-xl transition">
                  <div class="p-2 bg-[#557A61] text-white rounded-xl"><i data-lucide="layers" class="w-4 h-4"></i></div>
                  <div>
                    <div>下載全組織 Master Excel (16 Sheets 完整整合)</div>
                    <div class="text-[11px] text-[#557A61] font-normal">包含所有自評、主管評、同儕匿名表</div>
                  </div>
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>

      <!-- MAIN TABS (NO-PRINT) -->
      <nav class="flex space-x-2 border-t border-[#E8E2D8] pt-2.5 -mb-px overflow-x-auto custom-scrollbar no-print">
        <button onclick="switchTab('subReview')" id="tab-btn-subReview" class="tab-btn inline-flex items-center gap-2 px-5 py-3 text-xs sm:text-sm font-semibold rounded-t-xl border-b-2 border-[#557A61] text-[#2D5239] bg-[#E4ECD3]/30 shrink-0 transition">
          <i data-lucide="git-compare" class="w-4 h-4 text-[#557A61]"></i>
          自評 vs 主管評對照（主管填核）
          <span class="px-2 py-0.5 text-xs rounded-full bg-[#E4ECD3] text-[#2D5239] font-bold">核心對照</span>
        </button>
        <button onclick="switchTab('supervisor')" id="tab-btn-supervisor" class="tab-btn inline-flex items-center gap-2 px-5 py-3 text-xs sm:text-sm font-medium rounded-t-xl border-b-2 border-transparent text-[#6E6662] hover:text-[#2E2827] hover:bg-[#F2EEE6]/50 shrink-0 transition">
          <i data-lucide="award" class="w-4 h-4"></i>
          評主管（部屬回饋與統計）
          <span class="px-2 py-0.5 text-xs rounded-full bg-[#F2EEE6] text-[#6E6662] font-bold" id="badge-sup-count">4</span>
        </button>
        <button onclick="switchTab('self')" id="tab-btn-self" class="tab-btn inline-flex items-center gap-2 px-5 py-3 text-xs sm:text-sm font-medium rounded-t-xl border-b-2 border-transparent text-[#6E6662] hover:text-[#2E2827] hover:bg-[#F2EEE6]/50 shrink-0 transition">
          <i data-lucide="user-check" class="w-4 h-4"></i>
          自評（依主管分流）
          <span class="px-2 py-0.5 text-xs rounded-full bg-[#F2EEE6] text-[#6E6662] font-bold" id="badge-self-count">5</span>
        </button>
        <button onclick="switchTab('peerAnon')" id="tab-btn-peerAnon" class="tab-btn inline-flex items-center gap-2 px-5 py-3 text-xs sm:text-sm font-medium rounded-t-xl border-b-2 border-transparent text-[#6E6662] hover:text-[#2E2827] hover:bg-[#F2EEE6]/50 shrink-0 transition">
          <i data-lucide="users" class="w-4 h-4"></i>
          員工同儕匿名表（逐題明細）
          <span class="px-2 py-0.5 text-xs rounded-full bg-[#F2EEE6] text-[#6E6662] font-bold">9人</span>
        </button>
        <button onclick="switchTab('peer')" id="tab-btn-peer" class="tab-btn inline-flex items-center gap-2 px-5 py-3 text-xs sm:text-sm font-medium rounded-t-xl border-b-2 border-transparent text-[#6E6662] hover:text-[#2E2827] hover:bg-[#F2EEE6]/50 shrink-0 transition">
          <i data-lucide="pie-chart" class="w-4 h-4"></i>
          評同事（同儕回饋總覽）
          <span class="px-2 py-0.5 text-xs rounded-full bg-[#F2EEE6] text-[#6E6662] font-bold" id="badge-peer-count">23</span>
        </button>
      </nav>
    </div>
  </header>

  <!-- MAIN CONTENT AREA -->
  <main class="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 sm:py-8">

    <!-- ========================================================================= -->
    <!-- TAB: 自評 vs 主管評對照（主管填核回傳） -->
    <!-- ========================================================================= -->
    <section id="tab-section-subReview" class="tab-content space-y-7">
      <div class="screen-only-view space-y-7">
        
        <!-- SUPERVISOR SELECTION FILTER BAR -->
        <div class="bg-[#FFFDF9] rounded-2xl p-6 border border-[#E8E2D8] soft-card-shadow flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div class="flex items-center gap-2.5 flex-wrap">
            <span class="text-xs font-bold text-[#8C837C] mr-1 flex items-center gap-1.5 uppercase tracking-wider">
              <i data-lucide="filter" class="w-3.5 h-3.5 text-[#557A61]"></i> 主管團隊：
            </span>
            <button onclick="filterSubReviewTeam('張希慈')" id="sub-team-btn-張希慈" class="sub-team-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold bg-[#557A61] text-white shadow-xs transition">
              張希慈 的部屬 (5人)
            </button>
            <button onclick="filterSubReviewTeam('何維安')" id="sub-team-btn-何維安" class="sub-team-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
              何維安 的部屬 (1人)
            </button>
            <button onclick="filterSubReviewTeam('姚品瑄')" id="sub-team-btn-姚品瑄" class="sub-team-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
              姚品瑄 的部屬 (2人)
            </button>
            <button onclick="filterSubReviewTeam('張希慈_執行長')" id="sub-team-btn-張希慈_執行長" class="sub-team-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
              執行長個人自評
            </button>
            <button onclick="filterSubReviewTeam('ALL')" id="sub-team-btn-ALL" class="sub-team-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
              全體員工 (9人)
            </button>
          </div>

          <div class="flex items-center gap-3">
            <button onclick="window.print()" class="inline-flex items-center gap-2 px-5 py-2.5 text-xs sm:text-sm font-semibold rounded-xl bg-[#557A61] text-white hover:bg-[#466551] transition shadow-xs">
              <i data-lucide="printer" class="w-4 h-4"></i> 列印 / 存為 PDF (Cmd+P)
            </button>
            <div id="sub-review-export-btn-container">
              <!-- Export button -->
            </div>
          </div>
        </div>

        <!-- SUBORDINATE MEMBER SELECTOR PILLS -->
        <div class="bg-[#FFFDF9] rounded-2xl p-5 border border-[#E8E2D8] soft-card-shadow flex items-center gap-3 flex-wrap" id="sub-review-member-pills">
          <!-- Injected via JS -->
        </div>

        <!-- REPORT CONTAINER -->
        <div id="sub-review-report-container" class="space-y-7">
          <!-- Injected via JS -->
        </div>
      </div>

      <!-- PRINT-ONLY WORD-STYLE CONTAINER -->
      <div id="sub-review-print-container" class="print-only-doc">
        <!-- Injected via JS -->
      </div>
    </section>

    <!-- ======================================================== -->
    <!-- TAB: 評主管 -->
    <!-- ======================================================== -->
    <section id="tab-section-supervisor" class="tab-content hidden space-y-7">
      <div class="screen-only-view space-y-7">
        <div class="bg-[#FFFDF9] rounded-2xl p-6 border border-[#E8E2D8] soft-card-shadow flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div class="flex items-center gap-2.5 flex-wrap">
            <span class="text-xs font-bold text-[#8C837C] mr-1 flex items-center gap-1.5 uppercase tracking-wider">
              <i data-lucide="user" class="w-3.5 h-3.5 text-[#557A61]"></i> 受評主管：
            </span>
            <button onclick="filterSupervisor('ALL')" id="sup-filter-btn-ALL" class="sup-filter-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
              全部主管 (4筆)
            </button>
            <button onclick="filterSupervisor('張希慈')" id="sup-filter-btn-張希慈" class="sup-filter-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold bg-[#557A61] text-white shadow-xs transition">
              張希慈 (3筆)
            </button>
            <button onclick="filterSupervisor('何維安')" id="sup-filter-btn-何維安" class="sup-filter-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
              何維安 (1筆)
            </button>
          </div>
          <button onclick="window.print()" class="inline-flex items-center gap-2 px-5 py-2.5 text-xs sm:text-sm font-semibold rounded-xl bg-[#557A61] text-white hover:bg-[#466551] transition shadow-xs">
            <i data-lucide="printer" class="w-4 h-4"></i> 列印 / 存為 PDF (Cmd+P)
          </button>
        </div>

        <div id="sup-stat-summary-cards" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-5"></div>

        <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <div class="bg-[#FFFDF9] p-6 rounded-2xl border border-[#E8E2D8] soft-card-shadow">
            <h3 class="text-sm sm:text-base font-bold text-[#2E2827] mb-3 flex items-center gap-2 font-serif-tc">
              <i data-lucide="compass" class="w-4 h-4 text-[#557A61]"></i> 四大文化實踐維度
            </h3>
            <div class="h-68 flex items-center justify-center">
              <canvas id="supervisorRadarChart"></canvas>
            </div>
          </div>
          <div class="bg-[#FFFDF9] p-6 rounded-2xl border border-[#E8E2D8] soft-card-shadow lg:col-span-2">
            <h3 class="text-sm sm:text-base font-bold text-[#2E2827] mb-3 flex items-center gap-2 font-serif-tc">
              <i data-lucide="bar-chart-3" class="w-4 h-4 text-[#557A61]"></i> 管理能力各題平均得分 (滿分10分)
            </h3>
            <div class="h-68">
              <canvas id="supervisorBarChart"></canvas>
            </div>
          </div>
        </div>

        <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] p-6 sm:p-8 soft-card-shadow">
          <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-[#E8E2D8] pb-4 mb-5">
            <div class="flex items-center gap-2.5">
              <div class="p-2 bg-[#FCE5CD] text-[#783E16] rounded-xl"><i data-lucide="table-properties" class="w-4 h-4"></i></div>
              <h3 class="text-base sm:text-lg font-bold text-[#2E2827] font-serif-tc">各題項評分統計表（平均分、最高最好、最低最差）</h3>
            </div>
            <span class="text-xs text-[#8C837C]">滿分 10 分 ｜ 標註綠色代表優異 (≥9.0分)，橙色代表有成長空間 (≤7.0分)</span>
          </div>

          <div class="overflow-x-auto rounded-xl border border-[#E8E2D8]">
            <table class="w-full text-xs sm:text-sm text-left border-collapse" id="sup-item-stats-table">
              <thead class="bg-[#FCE5CD] text-[#4A2E1C]">
                <tr>
                  <th class="py-3 px-3.5 font-bold text-center w-14 border-r border-[#E8E2D8]">題號</th>
                  <th class="py-3 px-3.5 font-bold text-center w-28 border-r border-[#E8E2D8]">評估面向</th>
                  <th class="py-3 px-4 font-bold border-r border-[#E8E2D8]">題目內容與能力指引</th>
                  <th class="py-3 px-3.5 font-bold text-center w-24 border-r border-[#E8E2D8] bg-[#E4ECD3]/80 text-[#2D5239]">平均得分</th>
                  <th class="py-3 px-3.5 font-bold text-center w-24 border-r border-[#E8E2D8]">最好 (最高)</th>
                  <th class="py-3 px-3.5 font-bold text-center w-24 border-r border-[#E8E2D8]">最差 (最低)</th>
                  <th class="py-3 px-4 font-bold text-center w-36">給分明細</th>
                </tr>
              </thead>
              <tbody class="divide-y divide-[#EFEAE1] bg-white" id="sup-item-stats-tbody"></tbody>
            </table>
          </div>
        </div>

        <div id="supervisor-feedback-list" class="space-y-6"></div>
      </div>

      <div id="sup-print-document-container" class="print-only-doc"></div>
    </section>

    <!-- ======================================================== -->
    <!-- TAB: 自評 (依主管分流) -->
    <!-- ======================================================== -->
    <section id="tab-section-self" class="tab-content hidden space-y-7">
      <div class="bg-[#FFFDF9] rounded-2xl p-6 border border-[#E8E2D8] soft-card-shadow flex flex-col md:flex-row md:items-center justify-between gap-4 no-print">
        <div class="flex items-center gap-2.5 flex-wrap" id="self-supervisor-pills">
          <span class="text-xs font-bold text-[#8C837C] mr-1 flex items-center gap-1.5 uppercase tracking-wider">
            <i data-lucide="filter" class="w-3.5 h-3.5 text-[#557A61]"></i> 主管團隊：
          </span>
          <button onclick="filterSelfSupervisor('張希慈')" id="self-sup-btn-張希慈" class="self-sup-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold bg-[#557A61] text-white shadow-xs transition">
            張希慈 的部屬 (5人)
          </button>
          <button onclick="filterSelfSupervisor('何維安')" id="self-sup-btn-何維安" class="self-sup-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
            何維安 的部屬 (1人)
          </button>
          <button onclick="filterSelfSupervisor('姚品瑄')" id="self-sup-btn-姚品瑄" class="self-sup-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
            姚品瑄 的部屬 (2人)
          </button>
          <button onclick="filterSelfSupervisor('張希慈_執行長')" id="self-sup-btn-張希慈_執行長" class="self-sup-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
            執行長個人自評
          </button>
          <button onclick="filterSelfSupervisor('ALL')" id="self-sup-btn-ALL" class="self-sup-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
            全部已自評 (5人)
          </button>
        </div>

        <div id="self-export-btn-container"></div>
      </div>

      <div id="self-team-banner" class="bg-[#FCE5CD]/40 border border-[#F3D1B0] rounded-2xl p-6 flex items-center justify-between">
        <div class="flex items-center gap-3.5">
          <div class="p-3 bg-[#F4CCCC] text-[#3E2426] rounded-xl shadow-2xs">
            <i data-lucide="users" class="w-5 h-5"></i>
          </div>
          <div>
            <h2 id="self-team-title" class="text-base sm:text-lg font-bold text-[#3E2426] font-serif-tc">張希慈 的部屬自評列表</h2>
            <p id="self-team-subtitle" class="text-xs sm:text-sm text-[#7A4822] mt-0.5">涵蓋部屬：何維安（品牌經理）、陳泳璇（行政經理）、張芳媐（營運經理兼執行長特助）、姚品瑄（部門儲備主管）、胡喻翔（專案經理）</p>
          </div>
        </div>
        <div class="text-xs sm:text-sm px-4 py-2 bg-white/90 rounded-full text-[#783E16] font-bold border border-[#F3D1B0] shadow-2xs" id="self-completion-status">
          已填答 2 / 應填 5 人
        </div>
      </div>

      <div id="self-eval-cards-container" class="space-y-9"></div>
    </section>

    <!-- ======================================================== -->
    <!-- TAB: 員工同儕匿名表 (逐題明細) -->
    <!-- ======================================================== -->
    <section id="tab-section-peerAnon" class="tab-content hidden space-y-7">
      <div class="bg-[#FFFDF9] rounded-2xl p-6 border border-[#E8E2D8] soft-card-shadow flex flex-col md:flex-row md:items-center justify-between gap-4 no-print">
        <div class="flex items-center gap-2.5 flex-wrap" id="peer-anon-member-pills"></div>
        <div id="peer-anon-export-btn-container"></div>
      </div>

      <div id="peer-anon-report-container" class="space-y-7"></div>
    </section>

    <!-- ======================================================== -->
    <!-- TAB: 評同事 (同儕回饋總覽) -->
    <!-- ======================================================== -->
    <section id="tab-section-peer" class="tab-content hidden space-y-7">
      <div class="bg-[#FFFDF9] rounded-2xl p-6 border border-[#E8E2D8] soft-card-shadow flex flex-col md:flex-row md:items-center justify-between gap-4 no-print">
        <div class="flex items-center gap-2.5 flex-wrap" id="peer-pills-container"></div>
      </div>

      <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <div class="bg-[#FFFDF9] p-6 rounded-2xl border border-[#E8E2D8] soft-card-shadow">
          <h3 class="text-sm sm:text-base font-bold text-[#2E2827] mb-3 flex items-center gap-2 font-serif-tc">
            <i data-lucide="compass" class="w-4 h-4 text-[#557A61]"></i> 文化實踐維度表現
          </h3>
          <div class="h-68 flex items-center justify-center">
            <canvas id="peerRadarChart"></canvas>
          </div>
        </div>
        <div class="bg-[#FFFDF9] p-6 rounded-2xl border border-[#E8E2D8] soft-card-shadow lg:col-span-2">
          <h3 class="text-sm sm:text-base font-bold text-[#2E2827] mb-3 flex items-center gap-2 font-serif-tc">
            <i data-lucide="bar-chart-3" class="w-4 h-4 text-[#557A61]"></i> 協作、當責與溝通評分 (滿分10分)
          </h3>
          <div class="h-68">
            <canvas id="peerBarChart"></canvas>
          </div>
        </div>
      </div>

      <div id="peer-feedback-list" class="space-y-6"></div>
    </section>

  </main>

  <!-- TOAST -->
  <div id="toast" class="no-print fixed bottom-6 right-6 z-50 transform transition-all duration-300 opacity-0 translate-y-4 pointer-events-none bg-[#2E2827] text-white px-5 py-4 rounded-2xl shadow-xl flex items-center gap-3 text-xs sm:text-sm border border-[#4A433E]">
    <div id="toast-icon" class="p-1 rounded-lg bg-[#557A61] text-white"><i data-lucide="check" class="w-4 h-4"></i></div>
    <span id="toast-msg">操作成功</span>
  </div>

  <footer class="no-print bg-[#FFFDF9] border-t border-[#E8E2D8] py-7 text-center text-xs text-[#7A726D] mt-auto">
    好好星球文化基金會 360 年中成長評估系統 · 優先連線 Google 試算表（表單回覆1） · 支援 Cmd+P 莫蘭迪柔和色階列印
  </footer>

  <script>
    const GAS_ENDPOINT_URL = "https://script.google.com/macros/s/AKfycbxM-5YB3AX_CRK6APM3-dxGPUK7A2anQLrWSRwDK0_cZubdUu3pcUSl9lTPy5ahxXytgg/exec";

    // EMBEDDED DEFAULT / FALLBACK DATA
    let RAW_DATA = """ + raw_data_json + r""";

    // OFFICIAL ORG & TITLES
    const SUPERVISOR_TEAMS = {
      "張希慈": ["何維安", "陳泳璇", "張芳媐", "姚品瑄", "胡喻翔"],
      "何維安": ["林文琇"],
      "姚品瑄": ["薛筑瑄", "戴佑珍"],
      "張希慈_執行長": ["張希慈"],
      "ALL": ["何維安", "陳泳璇", "張芳媐", "姚品瑄", "胡喻翔", "林文琇", "薛筑瑄", "戴佑珍", "張希慈"]
    };

    const MEMBER_SUPERVISOR_MAP = {
      "何維安": "張希慈",
      "陳泳璇": "張希慈",
      "張芳媐": "張希慈",
      "姚品瑄": "張希慈",
      "胡喻翔": "張希慈",
      "林文琇": "何維安",
      "薛筑瑄": "姚品瑄",
      "戴佑珍": "姚品瑄",
      "張希慈": "董事會"
    };

    const ALL_MEMBERS = ["何維安", "陳泳璇", "張芳媐", "姚品瑄", "胡喻翔", "林文琇", "薛筑瑄", "戴佑珍", "張希慈"];

    const JOB_ROLES_MAP = {
      "張希慈": "執行長",
      "陳泳璇": "行政經理",
      "林文琇": "美感設計師",
      "胡喻翔": "專案經理",
      "張芳媐": "營運經理兼執行長特助",
      "何維安": "品牌經理",
      "姚品瑄": "部門儲備主管",
      "薛筑瑄": "專案經理",
      "戴佑珍": "專案經理"
    };

    const ROLE_COMPETENCIES = {
      "執行長": ["策略決策與組織方向設定", "組織治理與財務風險管理", "關鍵利害關係人關係建立與維繫", "公關、演講與媒體關係", "核心團隊培養與組織文化建立", "其他創造的好事與預防的壞事"],
      "行政經理": ["人事薪資與人資系統管理", "法規與政府公文管理", "董事會與治理作業執行", "財務核銷與內控執行", "總務採購與行政庶務管理", "其他創造的好事與預防的壞事"],
      "美感設計師": ["品牌識別系統設計與維護", "視覺與美感設計實務", "需求釐清與創意提案", "其他創造的好事與預防的壞事"],
      "專案經理": ["專案企劃與現場執行", "專案時程與預算規劃管理", "需求研究與方案迭代", "外部夥伴關係經營", "多元教學設計與現場引導", "其他創造的好事與預防的壞事"],
      "營運經理兼執行長特助": ["組織營運流程設計與優化", "制度文件與 SOP 建置", "人力資源策略與選用育留", "執行長幕僚協調與跨部門推動", "總務採購與行政庶務管理", "其他創造的好事與預防的壞事"],
      "品牌經理": ["品牌定位與外部溝通一致性", "行銷策略與議題倡議", "品牌活動策劃與策展敘事", "內部品牌管理與雇主品牌", "品牌危機與聲譽風險處理", "其他創造的好事與預防的壞事"],
      "部門儲備主管": ["部門策略規劃與專案組合管理", "專案經理管理與培育", "部門預算與資源配置管理", "跨領域專業掌握與推進", "利害關係人管理與衝突協調", "其他創造的好事與預防的壞事"]
    };

    const SUPERVISOR_QUESTIONS = [
      ["Q4", "日常賦能", "尋求協助（遇到困難時主動向主管尋求協助的容易度）", "q4_help_easy"],
      ["Q5", "指導指引", "具體引導（主管給予具體指引與改善方向的頻率與清晰度）", "q5_guidance_freq"],
      ["Q6", "個人成長", "改善幅度（主管回饋後個人能力與工作成果的改善幅度）", "q6_improve_degree"],
      ["Q7", "推動協作", "跨部門推動（協助排除跨部門障礙與推動專案的領導力）", "q7_cross_dept"],
      ["Q8", "資源配置", "資源評估（評估工作量與資源配置之合理性與同理心）", "q8_resource_eval"],
      ["Q9", "心理安全", "失誤回應（面對失誤時以建設性方式回應與引導）", "q9_constructive_mistake"],
      ["Q10", "肯定激勵", "肯定認可（看見並肯定部屬的努力與成果表現）", "q10_recognition"],
      ["Q11", "交付成果", "工作成效（整體工作交付與專案目標推動之成效）", "q11_overall_performance"],
      ["Q12", "信任文化", "【信任】表達想法（敢於向主管提出真實想法與疑慮）", "q12_trust_express"],
      ["Q13", "多元文化", "【多元】聆聽意見（能接納不同觀點與多元做事方法）", "q13_diversity_listen"],
      ["Q14", "實驗文化", "【實驗】嘗試創新（鼓勵嘗試新做法並給予實驗空間）", "q14_experiment_try"],
      ["Q15", "實驗文化", "【實驗】試錯空間（具備足夠的心理安全感與試錯包容）", "q15_experiment_psych_safety"],
      ["Q16", "可持續文化", "【可持續】讚美肯定（經常給予夥伴正向回饋與肯定）", "q16_sustain_praise"],
      ["Q17", "可持續文化", "【可持續】尊重界線（尊重下班生活與身心自我照顧邊界）", "q17_sustain_boundary"],
      ["Q18", "總體推薦", "NPS 推薦度（向他人推薦此主管領導的意願）", "q18_nps_recommend"],
      ["Q19", "滿意程度", "整體滿意度（對主管領導管理方式之整體滿意程度）", "q19_satisfaction"],
    ];

    const PEER_QUESTIONS = [
      ["Q24", "合作狀況", "合作狀況（溝通順暢度、資訊透明度、承諾事項達成率）", "q24_cooperation"],
      ["Q25", "工作品質", "注重細節（工作成果之品質與細緻度）", "q25_detail_oriented"],
      ["Q26", "專案進度", "準時完成（專案進度與時程掌控）", "q26_on_time"],
      ["Q27", "變動彈性", "靈活調整（面對臨時變動與突發狀況的彈性）", "q27_flexibility"],
      ["Q28", "追蹤承諾", "追蹤承諾（主動回報進度與承諾事項追蹤）", "q28_follow_up"],
      ["Q29", "透明溝通", "說明決策依據（決策與工作方法透明度）", "q29_transparency"],
      ["Q30", "多元文化", "【多元】接受不同意見（能開放聆聽反對觀點）", "q30_open_to_opposing"],
      ["Q31", "多元文化", "【多元】提出建設性觀點（在討論中給予實質建議）", "q31_constructive_opinions"],
      ["Q32", "實驗文化", "【實驗】開放調整（透過回饋持續迭代優化）", "q32_growth_mindset"],
      ["Q33", "信任文化", "【信任】分享經驗與資源（主動協助夥伴）", "q33_share_knowledge"],
      ["Q34", "可持續文化", "【可持續】讚美與肯定同儕（經常給予夥伴正向激勵）", "q34_praise_peers"],
      ["Q35", "可持續文化", "【可持續】尊重工作界線（維護彼此身心健康與邊界）", "q35_boundary_respect"],
      ["Q36", "NPS推薦", "NPS 推薦度（向他人推薦與此夥伴共事的意願）", "q36_nps_recommend"],
    ];

    let SUPERVISOR_EVAL_STATE = {};

    let currentSubReviewTeam = '張希慈';
    let currentSubReviewMember = '何維安';
    let currentSupervisorFilter = '張希慈';
    let currentSelfSupervisor = '張希慈';
    let currentPeerFilter = 'ALL';
    let currentPeerAnonMember = '何維安';

    let supervisorRadar = null;
    let supervisorBar = null;
    let peerRadar = null;
    let peerBar = null;

    function showToast(msg, isWarn = false) {
      const toast = document.getElementById('toast');
      if (!toast) return;
      document.getElementById('toast-msg').innerText = msg;
      const iconWrap = document.getElementById('toast-icon');
      if (iconWrap) {
        iconWrap.className = isWarn ? 'p-1 rounded-lg bg-[#C27D38] text-white' : 'p-1 rounded-lg bg-[#557A61] text-white';
      }
      toast.classList.remove('opacity-0', 'translate-y-4', 'pointer-events-none');
      setTimeout(() => {
        toast.classList.add('opacity-0', 'translate-y-4', 'pointer-events-none');
      }, 4000);
    }

    function toggleDropdown() {
      const menu = document.getElementById('dropdownMenu');
      if (menu) menu.classList.toggle('hidden');
    }

    window.onclick = function(e) {
      if (!e.target.closest('#exportDropdown')) {
        const menu = document.getElementById('dropdownMenu');
        if (menu && !menu.classList.contains('hidden')) menu.classList.add('hidden');
      }
    };

    function switchTab(tabId) {
      document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('border-[#557A61]', 'text-[#2D5239]', 'bg-[#E4ECD3]/30', 'font-semibold');
        btn.classList.add('border-transparent', 'text-[#6E6662]', 'font-medium');
      });
      const activeBtn = document.getElementById('tab-btn-' + tabId);
      if (activeBtn) {
        activeBtn.classList.add('border-[#557A61]', 'text-[#2D5239]', 'bg-[#E4ECD3]/30', 'font-semibold');
        activeBtn.classList.remove('border-transparent', 'text-[#6E6662]');
      }

      document.querySelectorAll('.tab-content').forEach(sec => sec.classList.add('hidden'));
      const activeSec = document.getElementById('tab-section-' + tabId);
      if (activeSec) activeSec.classList.remove('hidden');

      if (tabId === 'subReview') renderSubReviewSection();
      else if (tabId === 'supervisor') renderSupervisorSection();
      else if (tabId === 'self') renderSelfSection();
      else if (tabId === 'peerAnon') renderPeerAnonSection();
      else if (tabId === 'peer') renderPeerSection();
    }

    // =========================================================================
    // GOOGLE APPS SCRIPT SYNC ENGINE (PRIORITY FETCH WITH LOCAL FALLBACK)
    // =========================================================================
    function parse2DArrayData(rawRows) {
      if (!rawRows || rawRows.length < 2) return null;
      const header = rawRows[0];
      const dataRows = rawRows.slice(1);
      const newEntries = [];

      dataRows.forEach(row => {
        if (row.length < 4) return;
        const timestamp = (row[0] || "").toString().trim();
        const email = (row[1] || "").toString().trim();
        const target = (row[2] || "").toString().trim();
        const relation = (row[3] || "").toString().trim();
        const job_role = JOB_ROLES_MAP[target] || ((row[58] || "").toString().trim());

        const entry = { timestamp, email, target, relation, job_role };

        if (relation === "主管") {
          const scores = {};
          for (let c = 4; c <= 19; c++) {
            const val = (row[c] !== undefined && row[c] !== null) ? row[c].toString().trim() : "";
            scores[header[c]] = val ? parseFloat(val) : null;
          }
          entry.supervisor_eval = {
            q_scores: scores,
            q4_help_easy: scores[header[4]],
            q5_guidance_freq: scores[header[5]],
            q6_improve_degree: scores[header[6]],
            q7_cross_dept: scores[header[7]],
            q8_resource_eval: scores[header[8]],
            q9_constructive_mistake: scores[header[9]],
            q10_recognition: scores[header[10]],
            q11_overall_performance: scores[header[11]],
            q12_trust_express: scores[header[12]],
            q13_diversity_listen: scores[header[13]],
            q14_experiment_try: scores[header[14]],
            q15_experiment_psych_safety: scores[header[15]],
            q16_sustain_praise: scores[header[16]],
            q17_sustain_boundary: scores[header[17]],
            q18_nps_recommend: scores[header[18]],
            q19_satisfaction: scores[header[19]],
            q20_vision_mission: (row[20] || "").toString().trim(),
            q21_improvement_advice: (row[21] || "").toString().trim(),
            q22_other_comments: (row[22] || "").toString().trim(),
            q23_starlight_thanks: (row[23] || "").toString().trim()
          };
        } else if (relation === "同事") {
          const scores = {};
          for (let c = 24; c <= 36; c++) {
            const val = (row[c] !== undefined && row[c] !== null) ? row[c].toString().trim() : "";
            scores[header[c]] = val ? parseFloat(val) : null;
          }
          entry.peer_eval = {
            q_scores: scores,
            q24_cooperation: scores[header[24]],
            q25_detail_oriented: scores[header[25]],
            q26_on_time: scores[header[26]],
            q27_flexibility: scores[header[27]],
            q28_follow_up: scores[header[28]],
            q29_transparency: scores[header[29]],
            q30_open_to_opposing: scores[header[30]],
            q31_constructive_opinions: scores[header[31]],
            q32_growth_mindset: scores[header[32]],
            q33_share_knowledge: scores[header[33]],
            q34_praise_peers: scores[header[34]],
            q35_boundary_respect: scores[header[35]],
            q36_nps_recommend: scores[header[36]],
            q37_improvement_advice: (row[37] || "").toString().trim(),
            q38_other_comments: (row[38] || "").toString().trim(),
            q39_starlight_thanks: (row[39] || "").toString().trim()
          };
        } else if (relation === "自評") {
          const top3_stable = (row[52] || "").toString().split(",").map(s => s.trim()).filter(Boolean);
          const top3_practice = (row[53] || "").toString().split(",").map(s => s.trim()).filter(Boolean);
          const values = {
            "信任": (row[54] || "").toString().trim(),
            "多元": (row[55] || "").toString().trim(),
            "實驗": (row[56] || "").toString().trim(),
            "可持續": (row[57] || "").toString().trim()
          };

          const competencies = [];
          const reflection = {};

          for (let c = 59; c < row.length; c++) {
            const val = (row[c] || "").toString().trim();
            if (!val) continue;
            const colName = (header[c] || "").toString().trim();
            if (colName.includes("在未來的一年中") || colName.includes("卡關") || colName.includes("願景和使命")) {
              reflection[colName] = val;
            } else if (colName.includes("記得也要花") || colName.includes("下一階段進行")) {
              continue;
            } else {
              competencies.append ? competencies.push({ title: colName, answer: val }) : competencies.push({ title: colName, answer: val });
            }
          }

          entry.self_eval = {
            job_role,
            top3_stable,
            top3_practice,
            values,
            competencies,
            reflection
          };
        }
        newEntries.push(entry);
      });

      return newEntries;
    }

    async function syncFromGoogleAppsScript(isUserInitiated = false) {
      const syncBtn = document.getElementById('gasSyncBtn');
      const syncIcon = document.getElementById('gasSyncIcon');
      const syncText = document.getElementById('gasSyncText');
      const badge = document.getElementById('header-data-source-badge');

      if (syncIcon) syncIcon.classList.add('animate-spin-custom');
      if (syncText) syncText.innerText = "連線讀取中...";

      try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 8000);

        const response = await fetch(GAS_ENDPOINT_URL, {
          method: 'GET',
          signal: controller.signal
        });
        clearTimeout(timeoutId);

        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const rawRows = await response.json();

        const parsed = parse2DArrayData(rawRows);
        if (parsed && parsed.length > 0) {
          RAW_DATA = parsed;
          if (badge) {
            badge.className = "px-2.5 py-0.5 text-xs font-semibold rounded-md bg-[#E4ECD3] text-[#2D5239] border border-[#CDE0BC]";
            badge.innerText = `🟢 試算表即時連線 (${parsed.length}筆)`;
          }
          if (syncText) syncText.innerText = `已同步 (${parsed.length}筆)`;
          showToast(`成功優先同步 Google 試算表（表單回覆1）最新 ${parsed.length} 筆資料！`);
          refreshAllViews();
        } else {
          throw new Error("試算表回傳資料格式為空");
        }
      } catch (err) {
        console.warn("GAS Sync fallback to local data:", err);
        if (badge) {
          badge.className = "px-2.5 py-0.5 text-xs font-semibold rounded-md bg-[#F2EEE6] text-[#6E6662] border border-[#E0D7CA]";
          badge.innerText = `⚪ 本地備份資料 (${RAW_DATA.length}筆)`;
        }
        if (syncText) syncText.innerText = "同步 Google 試算表";
        if (isUserInitiated) {
          showToast("無法連線至 Google 試算表，系統已無縫切換使用本地最新備份資料。", true);
        }
        refreshAllViews();
      } finally {
        if (syncIcon) syncIcon.classList.remove('animate-spin-custom');
        if (typeof lucide !== 'undefined') lucide.createIcons();
      }
    }

    function refreshAllViews() {
      renderSubReviewSection();
      renderSupervisorSection();
      renderSelfSection();
      initPeerPills();
      initPeerAnonPills();
      if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    // =========================================================================
    // TAB: 自評 vs 主管評對照（主管填核回傳）
    // =========================================================================
    function filterSubReviewTeam(supKey) {
      currentSubReviewTeam = supKey;
      document.querySelectorAll('.sub-team-btn').forEach(btn => {
        btn.classList.remove('bg-[#557A61]', 'text-white', 'font-semibold');
        btn.classList.add('bg-[#F2EEE6]', 'text-[#4A433E]', 'font-medium');
      });
      const activeBtn = document.getElementById('sub-team-btn-' + supKey);
      if (activeBtn) {
        activeBtn.classList.add('bg-[#557A61]', 'text-white', 'font-semibold');
        activeBtn.classList.remove('bg-[#F2EEE6]', 'text-[#4A433E]', 'font-medium');
      }

      const memberList = SUPERVISOR_TEAMS[supKey] || [];
      if (!memberList.includes(currentSubReviewMember)) {
        currentSubReviewMember = memberList[0] || '何維安';
      }
      renderSubReviewSection();
    }

    function selectSubReviewMember(memName) {
      currentSubReviewMember = memName;
      document.querySelectorAll('.sub-review-pill-btn').forEach(btn => {
        btn.classList.remove('bg-[#557A61]', 'text-white', 'font-semibold');
        btn.classList.add('bg-[#F2EEE6]', 'text-[#4A433E]', 'font-medium');
      });
      const activeBtn = document.getElementById('sub-mem-btn-' + memName);
      if (activeBtn) {
        activeBtn.classList.add('bg-[#557A61]', 'text-white', 'font-semibold');
        activeBtn.classList.remove('bg-[#F2EEE6]', 'text-[#4A433E]', 'font-medium');
      }
      renderSubReviewSection();
    }

    function updateSupervisorRating(memName, compTitle, lvl) {
      if (!SUPERVISOR_EVAL_STATE[memName]) SUPERVISOR_EVAL_STATE[memName] = {};
      if (!SUPERVISOR_EVAL_STATE[memName][compTitle]) SUPERVISOR_EVAL_STATE[memName][compTitle] = {};
      SUPERVISOR_EVAL_STATE[memName][compTitle].level = lvl;
      renderSubReviewSection();
    }

    function updateSupervisorFeedback(memName, compTitle, text) {
      if (!SUPERVISOR_EVAL_STATE[memName]) SUPERVISOR_EVAL_STATE[memName] = {};
      if (!SUPERVISOR_EVAL_STATE[memName][compTitle]) SUPERVISOR_EVAL_STATE[memName][compTitle] = {};
      SUPERVISOR_EVAL_STATE[memName][compTitle].feedback = text;
    }

    function renderSubReviewSection() {
      const memberList = SUPERVISOR_TEAMS[currentSubReviewTeam] || [];
      const pillsContainer = document.getElementById('sub-review-member-pills');
      if (pillsContainer) {
        let phtml = `
          <span class="text-xs font-bold text-[#8C837C] mr-1 flex items-center gap-1.5 uppercase tracking-wider">
            <i data-lucide="user-check" class="w-3.5 h-3.5 text-[#557A61]"></i> 選擇部屬報告：
          </span>
        `;
        memberList.forEach(m => {
          const isAct = m === currentSubReviewMember;
          phtml += `
            <button onclick="selectSubReviewMember('${m}')" id="sub-mem-btn-${m}" class="sub-review-pill-btn px-5 py-2 rounded-xl text-xs sm:text-sm transition ${isAct ? 'bg-[#557A61] text-white font-semibold shadow-xs' : 'bg-[#F2EEE6] text-[#4A433E] font-medium hover:bg-[#EBE4D8]'}">
              ${m} (${JOB_ROLES_MAP[m] || '職位'})
            </button>
          `;
        });
        pillsContainer.innerHTML = phtml;
      }

      const expContainer = document.getElementById('sub-review-export-btn-container');
      if (expContainer) {
        expContainer.innerHTML = `
          <button onclick="exportSingleSubordinateComprehensiveExcelClientSide('${currentSubReviewMember}')" class="inline-flex items-center gap-2 px-5 py-2.5 text-xs sm:text-sm font-semibold rounded-xl bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition border border-[#E0D7CA]">
            <i data-lucide="download" class="w-4 h-4 text-[#557A61]"></i> 下載【${currentSubReviewMember}】自評vs主管評 XLSX
          </button>
        `;
      }

      const memName = currentSubReviewMember;
      const jobRole = JOB_ROLES_MAP[memName] || "專案經理";
      const supName = MEMBER_SUPERVISOR_MAP[memName] || "主管";

      const peerRecords = RAW_DATA.filter(e => e.relation === "同事" && e.target === memName);
      const numPeers = peerRecords.length;

      const selfEntry = RAW_DATA.find(e => e.relation === "自評" && e.target === memName);
      const hasSelf = Boolean(selfEntry && selfEntry.self_eval);
      const se = hasSelf ? selfEntry.self_eval : null;

      let allPeerScores = [];
      let itemStats = [];

      PEER_QUESTIONS.forEach(([qNo, qCat, qDesc, qKey]) => {
        const scores = peerRecords.map(r => r.peer_eval ? r.peer_eval[qKey] : null).filter(v => v !== null && v !== undefined);
        if (scores.length > 0) {
          const avg = scores.reduce((a, b) => a + b, 0) / scores.length;
          const best = Math.max(...scores);
          const worst = Math.min(...scores);
          allPeerScores.push(...scores);
          itemStats.push({ qNo, qCat, qDesc, qKey, avg, best, worst, scores });
        } else {
          itemStats.push({ qNo, qCat, qDesc, qKey, avg: null, best: null, worst: null, scores: [] });
        }
      });

      const overallPeerAvg = allPeerScores.length ? (allPeerScores.reduce((a, b) => a + b, 0) / allPeerScores.length).toFixed(2) : "-";
      const validItems = itemStats.filter(it => it.avg !== null);
      const sortedItems = [...validItems].sort((a, b) => b.avg - a.avg);
      const topStrengths = sortedItems.slice(0, 3);
      const bottomGrowth = [...sortedItems].reverse().slice(0, 3);

      const npsVals = peerRecords.map(r => r.peer_eval?.q36_nps_recommend).filter(v => v !== null && v !== undefined);
      const npsAvg = npsVals.length ? (npsVals.reduce((a, b) => a + b, 0) / npsVals.length).toFixed(2) : "-";

      const compList = ROLE_COMPETENCIES[jobRole] || [];

      // 1. SCREEN VIEW
      const screenContainer = document.getElementById('sub-review-report-container');
      if (screenContainer) {
        screenContainer.innerHTML = `
          <!-- REPORT HEADER BANNER -->
          <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] p-6 sm:p-8 soft-card-shadow flex flex-col md:flex-row md:items-center justify-between gap-5">
            <div class="flex items-center gap-4">
              <div class="w-14 h-14 rounded-2xl bg-[#557A61] text-white font-bold flex items-center justify-center text-xl font-serif-tc shadow-xs">
                ${memName.slice(0, 1)}
              </div>
              <div>
                <div class="flex items-center gap-2.5 flex-wrap">
                  <h2 class="text-xl sm:text-2xl font-bold text-[#2E2827] font-serif-tc">${memName} 自評 vs 主管評估對照表</h2>
                  <span class="px-3 py-1 text-xs font-semibold rounded-lg bg-[#FCE5CD] text-[#783E16] border border-[#F3D1B0]">${jobRole}</span>
                  ${hasSelf ? `
                    <span class="px-2.5 py-0.5 text-xs font-medium rounded-full bg-[#E4ECD3] text-[#2D5239] border border-[#CDE0BC] flex items-center gap-1">
                      <i data-lucide="check-circle" class="w-3.5 h-3.5"></i> 部屬已自評
                    </span>
                  ` : `
                    <span class="px-2.5 py-0.5 text-xs font-medium rounded-full bg-[#FFF4CD] text-[#7A5E12] border border-[#FCE299] flex items-center gap-1">
                      <i data-lucide="clock" class="w-3.5 h-3.5"></i> 部屬尚未自評
                    </span>
                  `}
                </div>
                <p class="text-xs text-[#7A726D] mt-1">直屬主管：<b>${supName}</b> ｜ 同儕填答：<b>${numPeers}</b> 份 ｜ 評估週期：2026 年中評估</p>
              </div>
            </div>
            <div class="flex items-center gap-3 text-xs sm:text-sm">
              <div class="px-4 py-2 rounded-xl bg-[#E4ECD3] text-[#2D5239] font-bold">
                同儕平均：${overallPeerAvg} / 10分
              </div>
              <div class="px-4 py-2 rounded-xl bg-[#FCE5CD] text-[#8C4B1E] font-bold">
                NPS 推薦：${npsAvg} 分
              </div>
            </div>
          </div>

          <!-- PART 1: 頂部同儕統計摘要卡片 -->
          <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-5">
            <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] p-6 soft-card-shadow flex flex-col justify-between">
              <div class="flex items-center justify-between pb-2">
                <span class="text-xs font-bold text-[#8C837C] uppercase tracking-wider">同儕評分總平均</span>
                <div class="p-2 bg-[#E4ECD3] text-[#2D5239] rounded-xl"><i data-lucide="award" class="w-4 h-4"></i></div>
              </div>
              <div class="my-3">
                <div class="text-3xl sm:text-4xl font-bold text-[#2E2827] font-serif-tc">${overallPeerAvg} <span class="text-xs text-[#8C837C] font-normal">/ 10分</span></div>
              </div>
              <div class="text-xs text-[#7A726D] flex items-center justify-between pt-3 border-t border-[#EFEAE1]">
                <span>NPS 推薦：<b>${npsAvg}</b> 分</span>
                <span>有效樣本：<b>${numPeers}</b> 份</span>
              </div>
            </div>

            <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] p-6 soft-card-shadow flex flex-col justify-between">
              <div class="flex items-center justify-between pb-2">
                <span class="text-xs font-bold text-[#8C837C] uppercase tracking-wider">部屬工作特質</span>
                <div class="p-2 bg-[#FCE5CD] text-[#783E16] rounded-xl"><i data-lucide="tag" class="w-4 h-4"></i></div>
              </div>
              <div class="my-2 space-y-1">
                <div class="text-xs text-[#2D5239] font-medium truncate">最穩定：${(se?.top3_stable || []).join('、') || '（未填）'}</div>
                <div class="text-xs text-[#8C4B1E] font-medium truncate">練習中：${(se?.top3_practice || []).join('、') || '（未填）'}</div>
              </div>
              <div class="text-xs text-[#7A726D] pt-3 border-t border-[#EFEAE1]">
                部屬自評特質盤點
              </div>
            </div>

            <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] p-6 soft-card-shadow flex flex-col justify-between">
              <div class="flex items-center justify-between pb-2">
                <span class="text-xs font-bold text-[#2D5239] uppercase tracking-wider flex items-center gap-1.5">
                  <i data-lucide="sparkles" class="w-3.5 h-3.5 text-[#557A61]"></i> 同儕評分最高 Top 3
                </span>
                <span class="badge-best px-2 py-0.5 text-[11px] font-bold rounded-md">亮點</span>
              </div>
              <div class="space-y-1.5 my-2">
                ${topStrengths.length > 0 ? topStrengths.map(it => `
                  <div class="text-xs flex items-center justify-between">
                    <span class="truncate text-[#2E2827] font-medium mr-2">${it.qNo}. ${it.qDesc.split('（')[0]}</span>
                    <span class="badge-best px-2 py-0.5 rounded font-bold shrink-0">${it.avg.toFixed(1)}分</span>
                  </div>
                `).join('') : '<div class="text-xs text-[#8C837C]">尚無同儕評分</div>'}
              </div>
              <div class="text-[11px] text-[#7A726D] pt-2 border-t border-[#EFEAE1]">最高給分達 ${topStrengths[0] ? topStrengths[0].best : '-'} 分</div>
            </div>

            <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] p-6 soft-card-shadow flex flex-col justify-between">
              <div class="flex items-center justify-between pb-2">
                <span class="text-xs font-bold text-[#8C4B1E] uppercase tracking-wider flex items-center gap-1.5">
                  <i data-lucide="trending-up" class="w-3.5 h-3.5 text-[#C27D38]"></i> 相對待提升 Top 3
                </span>
                <span class="badge-worst px-2 py-0.5 text-[11px] font-bold rounded-md">成長</span>
              </div>
              <div class="space-y-1.5 my-2">
                ${bottomGrowth.length > 0 ? bottomGrowth.map(it => `
                  <div class="text-xs flex items-center justify-between">
                    <span class="truncate text-[#2E2827] font-medium mr-2">${it.qNo}. ${it.qDesc.split('（')[0]}</span>
                    <span class="badge-worst px-2 py-0.5 rounded font-bold shrink-0">${it.avg.toFixed(1)}分</span>
                  </div>
                `).join('') : '<div class="text-xs text-[#8C837C]">尚無同儕評分</div>'}
              </div>
              <div class="text-[11px] text-[#7A726D] pt-2 border-t border-[#EFEAE1]">最低給分落點 ${bottomGrowth[0] ? bottomGrowth[0].worst : '-'} 分</div>
            </div>
          </div>

          <!-- PART 2: 組織文化實踐：部屬自評 vs 主管評核回饋 (一列一個面向) -->
          <div class="space-y-5">
            <div class="flex items-center justify-between">
              <div class="flex items-center gap-2.5">
                <div class="p-2 bg-[#F4CCCC] text-[#4A2426] rounded-xl"><i data-lucide="heart" class="w-5 h-5"></i></div>
                <div>
                  <h3 class="text-base sm:text-lg font-bold text-[#2E2827] font-serif-tc">一、組織文化實踐：部屬自評實例 vs 主管評核回饋</h3>
                  <p class="text-xs text-[#7A726D]">四大文化面向一列一項，對照部屬自評 STAR 描述與主管回饋</p>
                </div>
              </div>
            </div>

            <div class="overflow-x-auto rounded-2xl border border-[#E8E2D8] bg-[#FFFDF9] soft-card-shadow">
              <table class="w-full text-xs sm:text-sm text-left border-collapse">
                <thead class="bg-[#FCE5CD] text-[#4A2E1C]">
                  <tr>
                    <th class="py-3 px-4 font-bold text-center w-28 border-r border-[#E8E2D8]">文化面向</th>
                    <th class="py-3 px-4 font-bold border-r border-[#E8E2D8] w-72">文化定義與行為指引</th>
                    <th class="py-3 px-4 font-bold border-r border-[#E8E2D8]">部屬自評實例 (STAR)</th>
                    <th class="py-3 px-4 font-bold w-80 bg-[#FFF2D6] text-[#B45309]">主管評核回饋與觀察</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-[#EFEAE1] bg-white">
                  ${[
                    ["信任", "【信任】獨立行動與決策、主動協作、雙向溝通，在高彈性下給予彼此信任", se?.values?.['信任']],
                    ["多元", "【多元】尊重差異、多元工作方法、主動表達不同觀點與想法", se?.values?.['多元']],
                    ["實驗", "【實驗】透過開放心態嘗試修正與反思，勇於檢討及給予回饋", se?.values?.['實驗']],
                    ["可持續", "【可持續】內在韌性、自我照顧、彈性的人際與工作邊界", se?.values?.['可持續']],
                  ].map(([cTitle, cDesc, cSelf]) => `
                    <tr class="hover:bg-[#FAF7F2] transition">
                      <td class="py-3.5 px-4 font-bold text-center text-[#2E2827] border-r border-[#E8E2D8] bg-[#FAF7F2]">${cTitle}</td>
                      <td class="py-3.5 px-4 text-[#4A433E] border-r border-[#E8E2D8] text-xs leading-relaxed">${cDesc}</td>
                      <td class="py-3.5 px-4 text-[#2E2827] border-r border-[#E8E2D8] leading-relaxed">
                        ${cSelf || '<span class="text-[#B45309] italic">（部屬未填寫）</span>'}
                      </td>
                      <td class="py-3.5 px-4 bg-[#FFFDF9]">
                        <textarea onblur="updateSupervisorFeedback('${memName}', '文化_${cTitle}', this.value)" placeholder="請輸入主管針對【${cTitle}】的回饋與觀察..." class="w-full text-xs p-2.5 rounded-xl border border-[#FCE299] bg-[#FFF2D6]/30 focus:bg-white transition focus:outline-[#557A61] resize-y" rows="2">${SUPERVISOR_EVAL_STATE[memName]?.[`文化_${cTitle}`]?.feedback || ''}</textarea>
                      </td>
                    </tr>
                  `).join('')}
                </tbody>
              </table>
            </div>
          </div>

          <!-- PART 3: 專業職能：自評 vs 主管評分並列對照 (一列一個面向) -->
          <div class="space-y-5">
            <div class="flex items-center justify-between">
              <div class="flex items-center gap-2.5">
                <div class="p-2 bg-[#E4ECD3] text-[#2D5239] rounded-xl"><i data-lucide="briefcase" class="w-5 h-5"></i></div>
                <div>
                  <h3 class="text-base sm:text-lg font-bold text-[#2E2827] font-serif-tc">二、專業職能：自評 vs 主管評分並列對照（逐項比對認知差異）</h3>
                  <p class="text-xs text-[#7A726D]">職能項目一列一項，並列呈現部屬自評實例、自評分數/等級與主管評核等級</p>
                </div>
              </div>
            </div>

            <div class="overflow-x-auto rounded-2xl border border-[#E8E2D8] bg-[#FFFDF9] soft-card-shadow">
              <table class="w-full text-xs sm:text-sm text-left border-collapse">
                <thead class="bg-[#FCE5CD] text-[#4A2E1C]">
                  <tr>
                    <th class="py-3 px-3.5 font-bold text-center w-40 border-r border-[#E8E2D8]">職能項目</th>
                    <th class="py-3 px-4 font-bold border-r border-[#E8E2D8]">部屬自評實例 (STAR)</th>
                    <th class="py-3 px-3 font-bold text-center w-24 border-r border-[#E8E2D8] bg-[#FFF4CD] text-[#7A5E12]">部屬自評</th>
                    <th class="py-3 px-3 font-bold text-center w-28 border-r border-[#E8E2D8] bg-[#FFF2D6] text-[#B45309]">主管評定</th>
                    <th class="py-3 px-3 font-bold text-center w-24 border-r border-[#E8E2D8]">落差分析</th>
                    <th class="py-3 px-4 font-bold w-68 bg-[#FFF2D6] text-[#B45309]">主管回饋與具體事證</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-[#EFEAE1] bg-white">
                  ${compList.map(cTitle => {
                    let selfAns = null;
                    if (hasSelf && se.competencies) {
                      const item = se.competencies.find(x => x.title === cTitle);
                      if (item) selfAns = item.answer;
                    }
                    const supRating = SUPERVISOR_EVAL_STATE[memName]?.[cTitle]?.level || '';
                    const supFb = SUPERVISOR_EVAL_STATE[memName]?.[cTitle]?.feedback || '';
                    
                    let gapBadge = '<span class="text-[#8C837C] text-xs">待主管評</span>';
                    if (supRating) {
                      gapBadge = '<span class="badge-best px-2 py-0.5 rounded text-xs font-bold">主管已評</span>';
                    }

                    return `
                      <tr class="hover:bg-[#FAF7F2] transition">
                        <td class="py-3.5 px-3.5 font-bold text-[#2E2827] border-r border-[#E8E2D8] bg-[#FAF7F2]">${cTitle}</td>
                        <td class="py-3.5 px-4 text-[#2E2827] border-r border-[#E8E2D8] leading-relaxed">
                          ${selfAns || '<span class="text-[#B45309] italic">（部屬未填寫自評實例）</span>'}
                        </td>
                        <td class="py-3.5 px-3 text-center border-r border-[#E8E2D8] bg-[#FFFDF9]">
                          <span class="badge-missing px-2 py-0.5 rounded text-xs font-semibold">待問卷補齊</span>
                        </td>
                        <td class="py-3.5 px-3 text-center border-r border-[#E8E2D8] bg-[#FFF2D6]/40">
                          <select onchange="updateSupervisorRating('${memName}', '${cTitle}', this.value)" class="text-xs font-bold p-1.5 rounded-lg border border-[#FCE299] bg-white focus:outline-[#557A61]">
                            <option value="">選取</option>
                            <option value="L5" ${supRating==='L5'?'selected':''}>L5 (Amazing!)</option>
                            <option value="L4" ${supRating==='L4'?'selected':''}>L4 (Good)</option>
                            <option value="L3" ${supRating==='L3'?'selected':''}>L3 (Keep)</option>
                            <option value="L2" ${supRating==='L2'?'selected':''}>L2 (Grow)</option>
                            <option value="L1" ${supRating==='L1'?'selected':''}>L1 (Start)</option>
                          </select>
                        </td>
                        <td class="py-3.5 px-3 text-center border-r border-[#E8E2D8]">${gapBadge}</td>
                        <td class="py-3.5 px-4 bg-[#FFFDF9]">
                          <textarea onblur="updateSupervisorFeedback('${memName}', '${cTitle}', this.value)" placeholder="輸入主管針對此職能的評語與建議..." class="w-full text-xs p-2 rounded-xl border border-[#FCE299] bg-[#FFF2D6]/30 focus:bg-white transition focus:outline-[#557A61] resize-y" rows="2">${supFb}</textarea>
                        </td>
                      </tr>
                    `;
                  }).join('')}
                </tbody>
              </table>
            </div>
          </div>

          <!-- PART 4: 同儕質化文字回饋 -->
          <div class="space-y-5">
            <div class="flex items-center gap-2.5">
              <div class="p-2 bg-[#FCE5CD] text-[#783E16] rounded-xl"><i data-lucide="message-square" class="w-5 h-5"></i></div>
              <div>
                <h3 class="text-base sm:text-lg font-bold text-[#2E2827] font-serif-tc">三、同儕質化文字回饋彙整</h3>
                <p class="text-xs text-[#7A726D]">彙整同儕針對工作提升建議、其他補充觀察與星光感謝詞</p>
              </div>
            </div>

            <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] p-6 sm:p-8 soft-card-shadow space-y-6">
              <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                <div class="text-xs sm:text-sm font-bold text-[#557A61] flex items-center gap-2">
                  <i data-lucide="trending-up" class="w-4 h-4"></i> Q37. 工作與文化提升建議
                </div>
                <div class="space-y-2.5">
                  ${numPeers > 0 ? peerRecords.map((r, i) => `
                    <div class="bg-white rounded-xl p-4 border border-[#E8E2D8] shadow-2xs">
                      <span class="text-xs font-bold text-[#8C837C] block mb-1">同儕 ${String.fromCharCode(65+i)}：</span>
                      <p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${r.peer_eval?.q37_improvement_advice || '（無填寫）'}</p>
                    </div>
                  `).join('') : '<p class="text-xs text-[#8C837C]">目前尚無同儕填答回饋</p>'}
                </div>
              </div>

              <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2">
                  <i data-lucide="message-circle" class="w-4 h-4 text-[#557A61]"></i> Q38. 其他補充評價與觀察
                </div>
                <div class="space-y-2.5">
                  ${numPeers > 0 ? peerRecords.map((r, i) => `
                    <div class="bg-white rounded-xl p-4 border border-[#E8E2D8] shadow-2xs">
                      <span class="text-xs font-bold text-[#8C837C] block mb-1">同儕 ${String.fromCharCode(65+i)}：</span>
                      <p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${r.peer_eval?.q38_other_comments || '（無填寫）'}</p>
                    </div>
                  `).join('') : '<p class="text-xs text-[#8C837C]">目前尚無同儕填答回饋</p>'}
                </div>
              </div>

              <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                <div class="text-xs sm:text-sm font-bold text-[#7A363A] flex items-center gap-2">
                  <i data-lucide="award" class="w-4 h-4 text-[#D48B7B]"></i> Q39. 肯定與感謝的話（好好星光大賞）
                </div>
                <div class="space-y-2.5">
                  ${numPeers > 0 ? peerRecords.map((r, i) => `
                    <div class="bg-white rounded-xl p-4 border border-[#F4CCCC] shadow-2xs">
                      <span class="text-xs font-bold text-[#8C5558] block mb-1">同儕 ${String.fromCharCode(65+i)}：</span>
                      <p class="text-xs sm:text-sm text-[#592629] font-medium leading-relaxed">${r.peer_eval?.q39_starlight_thanks || '（無填寫）'}</p>
                    </div>
                  `).join('') : '<p class="text-xs text-[#8C837C]">目前尚無同儕填答回饋</p>'}
                </div>
              </div>
            </div>
          </div>
        `;
      }

      // 2. PRINT VIEW
      const printContainer = document.getElementById('sub-review-print-container');
      if (printContainer) {
        printContainer.innerHTML = `
          <div style="text-align: center; margin-bottom: 14px;">
            <div style="font-size: 10.5pt; letter-spacing: 2px; color: #555; margin-bottom: 3px;">財團法人好好星球文化基金會</div>
            <div style="font-size: 16pt; font-weight: bold; color: #000; letter-spacing: 1px;">2026 年中 360 度組織成長評估報告</div>
            <div style="font-size: 12pt; font-weight: bold; color: #2D5239; margin-top: 3px;">【自評 vs 主管評對照】${memName}（${jobRole}）綜合報告書</div>
          </div>

          <div class="word-doc-divider"></div>

          <table class="word-doc-table">
            <tbody>
              <tr>
                <th style="width: 18%; background-color: #E4ECD3;">部屬姓名</th>
                <td style="width: 32%; font-weight: bold; font-size: 11pt;">${memName}</td>
                <th style="width: 18%; background-color: #FCE5CD;">職位 / 職稱</th>
                <td style="width: 32%; font-weight: bold;">${jobRole}</td>
              </tr>
              <tr>
                <th>直屬主管</th>
                <td style="font-weight: bold;">${supName}</td>
                <th>評估週期</th>
                <td>2026 年中評估 (H1)</td>
              </tr>
              <tr>
                <th>同儕平均得分</th>
                <td style="font-weight: bold; font-size: 11pt; color: #2D5239; background-color: #F8FAF6;">${overallPeerAvg} / 10.0 分</td>
                <th>同儕樣本 / NPS</th>
                <td>共 ${numPeers} 位同儕 ｜ NPS：${npsAvg} 分</td>
              </tr>
            </tbody>
          </table>

          <div class="word-doc-section">
            <div class="word-sec-title">壹、組織文化實踐：部屬自評實例 vs 主管評核回饋</div>
            <table class="word-doc-table">
              <thead>
                <tr style="background-color: #FCE5CD;">
                  <th style="width: 12%;">文化面向</th>
                  <th style="width: 28%;">定義說明</th>
                  <th style="width: 32%;">部屬自評實例 (STAR)</th>
                  <th style="width: 28%; background-color: #FFF2D6;">主管評核回饋</th>
                </tr>
              </thead>
              <tbody>
                ${[
                  ["【信任】", "獨立行動與決策、主動協作、雙向溝通", se?.values?.['信任']],
                  ["【多元】", "尊重差異、多元工作方法、主動表達不同觀點", se?.values?.['多元']],
                  ["【實驗】", "透過開放心態嘗試修正與反思，勇於檢討及給予回饋", se?.values?.['實驗']],
                  ["【可持續】", "內在韌性、自我照顧、彈性的人際與工作邊界", se?.values?.['可持續']],
                ].map(([cTitle, cDesc, cSelf]) => `
                  <tr>
                    <td style="text-align: center; font-weight: bold;">${cTitle}</td>
                    <td>${cDesc}</td>
                    <td>${cSelf || '（未填寫）'}</td>
                    <td>${SUPERVISOR_EVAL_STATE[memName]?.[`文化_${cTitle.replace(/[【】]/g,'')}`]?.feedback || '【待主管填寫回饋】'}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>

          <div class="word-doc-section">
            <div class="word-sec-title">貳、專業職能：自評 vs 主管評分並列對照（逐項比對）</div>
            <table class="word-doc-table">
              <thead>
                <tr style="background-color: #FCE5CD;">
                  <th style="width: 20%;">職能項目</th>
                  <th style="width: 32%;">部屬自評實例 (STAR)</th>
                  <th style="width: 10%; background-color: #FFF4CD;">自評</th>
                  <th style="width: 10%; background-color: #FFF2D6;">主管評</th>
                  <th style="width: 28%; background-color: #FFF2D6;">主管回饋與具體事證</th>
                </tr>
              </thead>
              <tbody>
                ${compList.map(cTitle => {
                  let selfAns = null;
                  if (hasSelf && se.competencies) {
                    const item = se.competencies.find(x => x.title === cTitle);
                    if (item) selfAns = item.answer;
                  }
                  const supLvl = SUPERVISOR_EVAL_STATE[memName]?.[cTitle]?.level || '【待評】';
                  const supFb = SUPERVISOR_EVAL_STATE[memName]?.[cTitle]?.feedback || '【待主管填寫】';
                  return `
                    <tr>
                      <td style="font-weight: bold; background-color: #FAFAFA;">${cTitle}</td>
                      <td>${selfAns || '（部屬未填寫自評實例）'}</td>
                      <td style="text-align: center;">待補</td>
                      <td style="text-align: center; font-weight: bold; color: #B45309;">${supLvl}</td>
                      <td>${supFb}</td>
                    </tr>
                  `;
                }).join('')}
              </tbody>
            </table>
          </div>

          <div class="word-doc-section">
            <div class="word-sec-title">參、同儕質化具體文字回饋紀錄（全匿名整理）</div>
            ${numPeers > 0 ? peerRecords.map((r, i) => `
              <div class="word-feedback-block" style="margin-top: 8px; margin-bottom: 12px;">
                <table class="word-doc-table">
                  <thead>
                    <tr>
                      <th colspan="2" style="text-align: left; background-color: #EAE6DF; font-size: 9pt; padding: 5px 8px;">
                        【同儕回饋 ${String.fromCharCode(65 + i)}】 ｜ NPS 推薦：${r.peer_eval?.q36_nps_recommend || '-'} 分 ｜ 合作評分：${r.peer_eval?.q24_cooperation || '-'} 分
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <th style="width: 28%; text-align: left; vertical-align: top; background-color: #FAFAFA;">一、工作提升建議</th>
                      <td style="width: 72%;">${r.peer_eval?.q37_improvement_advice || '（無填寫）'}</td>
                    </tr>
                    <tr>
                      <th style="text-align: left; vertical-align: top; background-color: #FAFAFA;">二、其他補充評價</th>
                      <td>${r.peer_eval?.q38_other_comments || '（無填寫）'}</td>
                    </tr>
                    <tr>
                      <th style="text-align: left; vertical-align: top; background-color: #FAFAFA;">三、星光感謝詞</th>
                      <td style="background-color: #FFFDF9; font-weight: 500;">${r.peer_eval?.q39_starlight_thanks || '（無填寫）'}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            `).join('') : '<div style="padding: 10px; font-style: italic; color: #888;">（目前尚無同儕填答回覆）</div>'}
          </div>

          <div class="print-footer">
            財團法人好好星球文化基金會 · 2026 年中 360 度組織成長評估 · 【機密文件 供主管管理輔導專用】
          </div>
        `;
      }

      if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    // ==========================================
    // TAB: 評主管
    // ==========================================
    function filterSupervisor(name) {
      currentSupervisorFilter = name;
      document.querySelectorAll('.sup-filter-btn').forEach(btn => {
        btn.classList.remove('bg-[#557A61]', 'text-white', 'font-semibold');
        btn.classList.add('bg-[#F2EEE6]', 'text-[#4A433E]', 'font-medium');
      });
      const activeBtn = document.getElementById('sup-filter-btn-' + name);
      if (activeBtn) {
        activeBtn.classList.add('bg-[#557A61]', 'text-white', 'font-semibold');
        activeBtn.classList.remove('bg-[#F2EEE6]', 'text-[#4A433E]', 'font-medium');
      }
      renderSupervisorSection();
    }

    function renderSupervisorSection() {
      let filtered = RAW_DATA.filter(e => e.relation === "主管");
      if (currentSupervisorFilter !== 'ALL') {
        filtered = filtered.filter(e => e.target === currentSupervisorFilter);
      }

      const supTitle = currentSupervisorFilter === 'ALL' ? '全體主管彙整' : currentSupervisorFilter;
      const roleTitle = JOB_ROLES_MAP[currentSupervisorFilter] || '主管';

      let allScoreValues = [];
      let itemStatResults = [];

      SUPERVISOR_QUESTIONS.forEach(([qNo, qCat, qDesc, qKey]) => {
        const scores = filtered.map(e => e.supervisor_eval ? e.supervisor_eval[qKey] : null).filter(v => v !== null && v !== undefined);
        if (scores.length > 0) {
          const avg = scores.reduce((a, b) => a + b, 0) / scores.length;
          const best = Math.max(...scores);
          const worst = Math.min(...scores);
          allScoreValues.push(...scores);
          itemStatResults.push({ qNo, qCat, qDesc, qKey, avg, best, worst, scores });
        } else {
          itemStatResults.push({ qNo, qCat, qDesc, qKey, avg: null, best: null, worst: null, scores: [] });
        }
      });

      const overallAvg = allScoreValues.length ? (allScoreValues.reduce((a, b) => a + b, 0) / allScoreValues.length).toFixed(2) : "-";
      const validItems = itemStatResults.filter(it => it.avg !== null);
      const sortedByAvg = [...validItems].sort((a, b) => b.avg - a.avg);
      const topStrengths = sortedByAvg.slice(0, 3);
      const bottomOpportunities = [...sortedByAvg].reverse().slice(0, 3);

      const npsVals = filtered.map(e => e.supervisor_eval?.q18_nps_recommend).filter(v => v !== null && v !== undefined);
      const npsScore = npsVals.length ? (npsVals.reduce((a,b)=>a+b, 0) / npsVals.length).toFixed(2) : "-";
      const satVals = filtered.map(e => e.supervisor_eval?.q19_satisfaction).filter(v => v !== null && v !== undefined);
      const satScore = satVals.length ? (satVals.reduce((a,b)=>a+b, 0) / satVals.length).toFixed(2) : "-";

      const summaryCardsContainer = document.getElementById('sup-stat-summary-cards');
      if (summaryCardsContainer) {
        summaryCardsContainer.innerHTML = `
          <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] p-6 soft-card-shadow flex flex-col justify-between">
            <div class="flex items-center justify-between pb-2">
              <span class="text-xs font-bold text-[#8C837C] uppercase tracking-wider">整體評分平均</span>
              <div class="p-2 bg-[#E4ECD3] text-[#2D5239] rounded-xl"><i data-lucide="award" class="w-4 h-4"></i></div>
            </div>
            <div class="my-3">
              <div class="text-3xl sm:text-4xl font-bold text-[#2E2827] font-serif-tc">${overallAvg} <span class="text-xs text-[#8C837C] font-normal">/ 10分</span></div>
            </div>
            <div class="text-xs text-[#7A726D] flex items-center justify-between pt-3 border-t border-[#EFEAE1]">
              <span>NPS 推薦：<b>${npsScore}</b> 分</span>
              <span>整體滿意：<b>${satScore}</b> 分</span>
            </div>
          </div>

          <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] p-6 soft-card-shadow flex flex-col justify-between">
            <div class="flex items-center justify-between pb-2">
              <span class="text-xs font-bold text-[#8C837C] uppercase tracking-wider">受評樣本數量</span>
              <div class="p-2 bg-[#FCE5CD] text-[#783E16] rounded-xl"><i data-lucide="users" class="w-4 h-4"></i></div>
            </div>
            <div class="my-3">
              <div class="text-3xl sm:text-4xl font-bold text-[#2E2827] font-serif-tc">${filtered.length} <span class="text-xs text-[#8C837C] font-normal">份回覆</span></div>
            </div>
            <div class="text-xs text-[#7A726D] pt-3 border-t border-[#EFEAE1]">受評對象：<b>${supTitle}</b></div>
          </div>

          <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] p-6 soft-card-shadow flex flex-col justify-between">
            <div class="flex items-center justify-between pb-2">
              <span class="text-xs font-bold text-[#2D5239] uppercase tracking-wider flex items-center gap-1.5">
                <i data-lucide="sparkles" class="w-3.5 h-3.5 text-[#557A61]"></i> 表現最好項目 (Top 3)
              </span>
              <span class="badge-best px-2 py-0.5 text-[11px] font-bold rounded-md">亮點</span>
            </div>
            <div class="space-y-1.5 my-2">
              ${topStrengths.map(it => `
                <div class="text-xs flex items-center justify-between">
                  <span class="truncate text-[#2E2827] font-medium mr-2">${it.qNo}. ${it.qDesc.split('（')[0]}</span>
                  <span class="badge-best px-2 py-0.5 rounded font-bold shrink-0">${it.avg.toFixed(1)}分</span>
                </div>
              `).join('')}
            </div>
            <div class="text-[11px] text-[#7A726D] pt-2 border-t border-[#EFEAE1]">最高給分達 ${topStrengths[0] ? topStrengths[0].best : 10} 分</div>
          </div>

          <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] p-6 soft-card-shadow flex flex-col justify-between">
            <div class="flex items-center justify-between pb-2">
              <span class="text-xs font-bold text-[#8C4B1E] uppercase tracking-wider flex items-center gap-1.5">
                <i data-lucide="trending-up" class="w-3.5 h-3.5 text-[#C27D38]"></i> 相對待提升項目
              </span>
              <span class="badge-worst px-2 py-0.5 text-[11px] font-bold rounded-md">成長</span>
            </div>
            <div class="space-y-1.5 my-2">
              ${bottomOpportunities.map(it => `
                <div class="text-xs flex items-center justify-between">
                  <span class="truncate text-[#2E2827] font-medium mr-2">${it.qNo}. ${it.qDesc.split('（')[0]}</span>
                  <span class="badge-worst px-2 py-0.5 rounded font-bold shrink-0">${it.avg.toFixed(1)}分</span>
                </div>
              `).join('')}
            </div>
            <div class="text-[11px] text-[#7A726D] pt-2 border-t border-[#EFEAE1]">最低給分落點 ${bottomOpportunities[0] ? bottomOpportunities[0].worst : '-'} 分</div>
          </div>
        `;
      }

      const tbody = document.getElementById('sup-item-stats-tbody');
      if (tbody) {
        tbody.innerHTML = itemStatResults.map(it => {
          const avgDisplay = it.avg !== null ? it.avg.toFixed(2) : "-";
          const bestDisplay = it.best !== null ? it.best : "-";
          const worstDisplay = it.worst !== null ? it.worst : "-";
          const scoresDisplay = it.scores.length ? it.scores.join(", ") : "-";

          let avgBadgeClass = "text-[#2E2827]";
          if (it.avg !== null) {
            if (it.avg >= 9.0) avgBadgeClass = "bg-[#E4ECD3] text-[#2D5239] font-bold px-2 py-0.5 rounded-md";
            else if (it.avg <= 7.0) avgBadgeClass = "bg-[#FCE5CD] text-[#8C4B1E] font-bold px-2 py-0.5 rounded-md";
          }

          return `
            <tr class="hover:bg-[#FAF7F2] transition">
              <td class="py-3 px-3.5 text-center font-bold text-[#8C837C] border-r border-[#E8E2D8]">${it.qNo}</td>
              <td class="py-3 px-3.5 text-center text-[#4A433E] border-r border-[#E8E2D8]">${it.qCat}</td>
              <td class="py-3 px-4 text-[#2E2827] border-r border-[#E8E2D8] font-medium">${it.qDesc}</td>
              <td class="py-3 px-3.5 text-center border-r border-[#E8E2D8]"><span class="${avgBadgeClass}">${avgDisplay}</span></td>
              <td class="py-3 px-3.5 text-center font-bold text-[#2D5239] border-r border-[#E8E2D8]">${bestDisplay}</td>
              <td class="py-3 px-3.5 text-center font-bold text-[#8C4B1E] border-r border-[#E8E2D8]">${worstDisplay}</td>
              <td class="py-3 px-4 text-center text-xs text-[#7A726D]">${scoresDisplay}</td>
            </tr>
          `;
        }).join('');
      }

      const avgOf = (k) => {
        const it = itemStatResults.find(x => x.qKey === k);
        return it && it.avg !== null ? it.avg.toFixed(1) : 0;
      };

      const cultureLabels = ["信任 (真實表達)", "多元 (聆聽意見)", "實驗 (嘗試創新)", "心理安全 (試錯空間)", "肯定 (讚美認可)", "可持續 (尊重界線)"];
      const cultureData = [
        avgOf('q12_trust_express'), avgOf('q13_diversity_listen'), avgOf('q14_experiment_try'),
        avgOf('q15_experiment_psych_safety'), avgOf('q16_sustain_praise'), avgOf('q17_sustain_boundary')
      ];

      const mgmtLabels = ["尋求協助", "具體引導", "改善幅度", "跨部門推動", "資源評估", "失誤回應", "認可表現", "工作成效", "NPS推薦", "整體滿意"];
      const mgmtData = [
        avgOf('q4_help_easy'), avgOf('q5_guidance_freq'), avgOf('q6_improve_degree'),
        avgOf('q7_cross_dept'), avgOf('q8_resource_eval'), avgOf('q9_constructive_mistake'),
        avgOf('q10_recognition'), avgOf('q11_overall_performance'), avgOf('q18_nps_recommend'), avgOf('q19_satisfaction')
      ];

      const radarEl = document.getElementById('supervisorRadarChart');
      if (radarEl && typeof Chart !== 'undefined') {
        if (supervisorRadar) supervisorRadar.destroy();
        supervisorRadar = new Chart(radarEl.getContext('2d'), {
          type: 'radar',
          data: {
            labels: cultureLabels,
            datasets: [{
              label: supTitle,
              data: cultureData,
              backgroundColor: 'rgba(85, 122, 97, 0.2)',
              borderColor: '#557A61',
              pointBackgroundColor: '#557A61'
            }]
          },
          options: { scales: { r: { min: 0, max: 10, ticks: { stepSize: 2 } } }, plugins: { legend: { display: false } } }
        });
      }

      const barEl = document.getElementById('supervisorBarChart');
      if (barEl && typeof Chart !== 'undefined') {
        if (supervisorBar) supervisorBar.destroy();
        supervisorBar = new Chart(barEl.getContext('2d'), {
          type: 'bar',
          data: { labels: mgmtLabels, datasets: [{ data: mgmtData, backgroundColor: '#557A61', borderRadius: 6 }] },
          options: { scales: { y: { min: 0, max: 10 } }, plugins: { legend: { display: false } } }
        });
      }

      const feedbackContainer = document.getElementById('supervisor-feedback-list');
      if (feedbackContainer) {
        feedbackContainer.innerHTML = filtered.map(e => {
          const se = e.supervisor_eval || {};
          return `
            <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] p-6 sm:p-8 soft-card-shadow space-y-6">
              <div class="flex items-center justify-between border-b border-[#E8E2D8] pb-4">
                <div class="flex items-center gap-3">
                  <div class="w-10 h-10 rounded-xl bg-[#557A61] text-white font-bold flex items-center justify-center text-sm font-serif-tc shadow-2xs">${e.target.slice(0, 1)}</div>
                  <div>
                    <span class="font-bold text-[#2E2827] text-base font-serif-tc">受評主管：${e.target}</span>
                    <span class="text-xs text-[#8C837C] ml-2">（填答者：${e.email}）</span>
                  </div>
                </div>
                <div class="flex items-center gap-3 text-xs sm:text-sm">
                  <span class="px-3.5 py-1.5 rounded-xl bg-[#E4ECD3] text-[#2D5239] font-bold">NPS 推薦：${se.q18_nps_recommend || '-'} 分</span>
                  <span class="px-3.5 py-1.5 rounded-xl bg-[#FCE5CD] text-[#8C4B1E] font-bold">滿意度：${se.q19_satisfaction || '-'} 分</span>
                </div>
              </div>

              <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                  <div class="text-xs sm:text-sm font-bold text-[#557A61] flex items-center gap-2"><i data-lucide="compass" class="w-4 h-4"></i> Q20. 願景使命理解之引導</div>
                  <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.q20_vision_mission || '（無）'}</p></div>
                </div>
                <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                  <div class="text-xs sm:text-sm font-bold text-[#783E16] flex items-center gap-2"><i data-lucide="trending-up" class="w-4 h-4 text-[#C27D38]"></i> Q21. 管理與文化精神建議</div>
                  <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.q21_improvement_advice || '（無）'}</p></div>
                </div>
                <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                  <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><i data-lucide="message-square" class="w-4 h-4 text-[#557A61]"></i> Q22. 其他補充評價</div>
                  <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.q22_other_comments || '（無）'}</p></div>
                </div>
                <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                  <div class="text-xs sm:text-sm font-bold text-[#7A363A] flex items-center gap-2"><i data-lucide="award" class="w-4 h-4 text-[#D48B7B]"></i> Q23. 肯定與感謝詞（好好星光大賞）</div>
                  <div class="bg-white rounded-xl p-4 my-2 border border-[#F4CCCC] shadow-2xs"><p class="text-xs sm:text-sm text-[#592629] font-medium leading-relaxed">${se.q23_starlight_thanks || '（無）'}</p></div>
                </div>
              </div>
            </div>
          `;
        }).join('');
      }

      if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    // ==========================================
    // TAB: 自評 (依主管分流)
    // ==========================================
    function filterSelfSupervisor(supKey) {
      currentSelfSupervisor = supKey;
      document.querySelectorAll('.self-sup-btn').forEach(btn => {
        btn.classList.remove('bg-[#557A61]', 'text-white', 'font-semibold');
        btn.classList.add('bg-[#F2EEE6]', 'text-[#4A433E]', 'font-medium');
      });
      const activeBtn = document.getElementById('self-sup-btn-' + supKey);
      if (activeBtn) {
        activeBtn.classList.add('bg-[#557A61]', 'text-white', 'font-semibold');
        activeBtn.classList.remove('bg-[#F2EEE6]', 'text-[#4A433E]', 'font-medium');
      }
      renderSelfSection();
    }

    function renderSelfSection() {
      const memberNames = SUPERVISOR_TEAMS[currentSelfSupervisor] || [];
      const container = document.getElementById('self-eval-cards-container');
      const bannerTitle = document.getElementById('self-team-title');
      const bannerSub = document.getElementById('self-team-subtitle');
      const compStatus = document.getElementById('self-completion-status');
      const exportContainer = document.getElementById('self-export-btn-container');

      if (bannerTitle && bannerSub) {
        if (currentSelfSupervisor === '張希慈') {
          bannerTitle.innerText = "【張希慈】部屬自評彙整";
          bannerSub.innerText = "涵蓋部屬：何維安（品牌經理）、陳泳璇（行政經理）、張芳媐（營運經理兼執行長特助）、姚品瑄（部門儲備主管）、胡喻翔（專案經理）";
        } else if (currentSelfSupervisor === '何維安') {
          bannerTitle.innerText = "【何維安】部屬自評彙整";
          bannerSub.innerText = "涵蓋部屬：林文琇（美感設計師）";
        } else if (currentSelfSupervisor === '姚品瑄') {
          bannerTitle.innerText = "【姚品瑄】部屬自評彙整";
          bannerSub.innerText = "涵蓋部屬：薛筑瑄（專案經理）、戴佑珍（專案經理）";
        } else if (currentSelfSupervisor === '張希慈_執行長') {
          bannerTitle.innerText = "【張希慈】執行長個人自評";
          bannerSub.innerText = "評估職位：執行長";
        } else {
          bannerTitle.innerText = "全組織自評總覽";
          bannerSub.innerText = "查看基金會所有已填寫之自評紀錄";
        }
      }

      if (exportContainer) {
        exportContainer.innerHTML = `
          <button onclick="exportSupervisorExcelClientSide('${currentSelfSupervisor}')" class="inline-flex items-center gap-2 px-5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold bg-[#557A61] text-white hover:bg-[#466551] transition shadow-xs">
            <i data-lucide="file-spreadsheet" class="w-4 h-4"></i>
            下載本組主管專用 XLSX 表格 (Chunk 格式＋配色)
          </button>
        `;
      }

      let completedCount = 0;
      let html = "";

      memberNames.forEach(memName => {
        const entry = RAW_DATA.find(e => e.target === memName && e.relation === "自評");
        if (entry && entry.self_eval) {
          completedCount++;
          const se = entry.self_eval;
          html += `
            <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] soft-card-shadow overflow-hidden">
              <div class="chunk-header-banner px-6 py-5 flex flex-col md:flex-row md:items-center justify-between gap-3.5">
                <div class="flex items-center gap-4">
                  <div class="w-12 h-12 rounded-2xl bg-[#3E2426] text-white font-bold flex items-center justify-center text-lg font-serif-tc shadow-2xs">${memName.slice(0, 1)}</div>
                  <div>
                    <div class="flex items-center gap-2.5">
                      <h3 class="text-base sm:text-lg font-bold text-[#3E2426] font-serif-tc">${memName}</h3>
                      <span class="px-3 py-1 text-xs font-semibold rounded-lg bg-[#FFFDF9] text-[#4A2E1C] border border-[#EED1B4]">${JOB_ROLES_MAP[memName] || se.job_role || '自評'}</span>
                      <span class="px-2.5 py-0.5 text-xs font-medium rounded-full bg-[#E4ECD3] text-[#2D5239] border border-[#CDE0BC] flex items-center gap-1"><i data-lucide="check-circle" class="w-3.5 h-3.5"></i> 已填寫自評</span>
                    </div>
                    <p class="text-xs text-[#7A4822] mt-1">${entry.email} · 填答時間：${entry.timestamp}</p>
                  </div>
                </div>
              </div>

              <div class="p-6 sm:p-8 space-y-7">
                <div>
                  <div class="text-xs sm:text-sm font-bold text-[#8C837C] uppercase tracking-wider mb-3.5 flex items-center gap-2 font-serif-tc"><i data-lucide="tag" class="w-4 h-4 text-[#557A61]"></i> 一、工作特質盤點</div>
                  <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                    <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                      <div class="text-xs sm:text-sm font-bold text-[#2D5239] flex items-center gap-2"><i data-lucide="shield-check" class="w-4 h-4 text-[#557A61]"></i> 最穩定、最具代表性 Top 3</div>
                      <div class="flex flex-wrap gap-2.5 pt-1">${(se.top3_stable || []).map(t => `<span class="badge-stable px-3.5 py-1.5 text-xs sm:text-sm font-semibold rounded-xl shadow-2xs">${t}</span>`).join('')}</div>
                    </div>
                    <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                      <div class="text-xs sm:text-sm font-bold text-[#783E16] flex items-center gap-2"><i data-lucide="trending-up" class="w-4 h-4 text-[#C27D38]"></i> 目前在練習 / 期望發展 3 項</div>
                      <div class="flex flex-wrap gap-2.5 pt-1">${(se.top3_practice || []).map(t => `<span class="badge-practice px-3.5 py-1.5 text-xs sm:text-sm font-semibold rounded-xl shadow-2xs">${t}</span>`).join('')}</div>
                    </div>
                  </div>
                </div>

                <div>
                  <div class="text-xs sm:text-sm font-bold text-[#8C837C] uppercase tracking-wider mb-3.5 flex items-center gap-2 font-serif-tc"><i data-lucide="heart" class="w-4 h-4 text-[#557A61]"></i> 二、四大文化實踐實例 (STAR 敘述)</div>
                  <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                    <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                      <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><span class="w-2.5 h-2.5 rounded-full bg-[#557A61]"></span> 信任 (Trust)</div>
                      <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.values?.['信任'] || '（未填寫）'}</p></div>
                    </div>
                    <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                      <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><span class="w-2.5 h-2.5 rounded-full bg-[#557A61]"></span> 多元 (Diversity)</div>
                      <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.values?.['多元'] || '（未填寫）'}</p></div>
                    </div>
                    <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                      <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><span class="w-2.5 h-2.5 rounded-full bg-[#557A61]"></span> 實驗 (Experiment)</div>
                      <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.values?.['實驗'] || '（未填寫）'}</p></div>
                    </div>
                    <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                      <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><span class="w-2.5 h-2.5 rounded-full bg-[#557A61]"></span> 可持續 (Sustainability)</div>
                      <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.values?.['可持續'] || '（未填寫）'}</p></div>
                    </div>
                  </div>
                </div>

                ${se.competencies && se.competencies.length > 0 ? `
                  <div>
                    <div class="text-xs sm:text-sm font-bold text-[#8C837C] uppercase tracking-wider mb-3.5 flex items-center gap-2 font-serif-tc"><i data-lucide="briefcase" class="w-4 h-4 text-[#557A61]"></i> 三、${JOB_ROLES_MAP[memName] || se.job_role} 專屬職能展現實例</div>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                      ${se.competencies.map(c => `
                        <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                          <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2 font-serif-tc"><span class="w-2.5 h-2.5 rounded-full bg-[#557A61]"></span> ${c.title}</div>
                          <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed whitespace-pre-line">${c.answer || '（未填寫）'}</p></div>
                        </div>
                      `).join('')}
                    </div>
                  </div>
                ` : ''}

                ${se.reflection && Object.keys(se.reflection).length > 0 ? `
                  <div>
                    <div class="text-xs sm:text-sm font-bold text-[#8C837C] uppercase tracking-wider mb-3.5 flex items-center gap-2 font-serif-tc"><i data-lucide="help-circle" class="w-4 h-4 text-[#557A61]"></i> 四、組織卡關點反思與未來價值展望</div>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                      ${Object.entries(se.reflection).map(([k, v]) => `
                        <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                          <div class="text-xs sm:text-sm font-bold text-[#783E16] flex items-center gap-2 font-serif-tc"><i data-lucide="sparkle" class="w-3.5 h-3.5 text-[#C27D38]"></i> ${k}</div>
                          <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs"><p class="text-xs sm:text-sm text-[#4A433E] leading-relaxed">${v || '（未填寫）'}</p></div>
                        </div>
                      `).join('')}
                    </div>
                  </div>
                ` : ''}
              </div>
            </div>
          `;
        } else {
          html += `
            <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] soft-card-shadow p-7 flex items-center justify-between">
              <div class="flex items-center gap-4">
                <div class="w-12 h-12 rounded-2xl bg-[#F2EEE6] text-[#8C837C] font-bold flex items-center justify-center text-lg font-serif-tc">${memName.slice(0, 1)}</div>
                <div>
                  <div class="flex items-center gap-2.5">
                    <h3 class="text-base sm:text-lg font-bold text-[#2E2827] font-serif-tc">${memName}</h3>
                    <span class="px-3 py-1 text-xs font-semibold rounded-lg bg-[#FCE5CD] text-[#783E16] border border-[#F3D1B0]">${JOB_ROLES_MAP[memName] || '職位'}</span>
                    <span class="px-3 py-1 text-xs font-medium rounded-full bg-[#FFF4CD] text-[#7A5E12] border border-[#FCE299] flex items-center gap-1.5"><i data-lucide="clock" class="w-3.5 h-3.5"></i> 尚未收到自評資料</span>
                  </div>
                  <p class="text-xs sm:text-sm text-[#8C837C] mt-1">此成員尚未於表單中送出自評紀錄，收到後上傳新 CSV 即可同步更新。</p>
                </div>
              </div>
              <div class="text-xs font-semibold text-[#8C837C] px-3.5 py-1.5 bg-[#F2EEE6] rounded-xl">待填寫</div>
            </div>
          `;
        }
      });

      if (compStatus) compStatus.innerText = `已填答 ${completedCount} / 應填 ${memberNames.length} 人`;
      if (container) container.innerHTML = html;
      if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    // ==========================================
    // TAB: 員工同儕匿名表 (逐題明細)
    // ==========================================
    function initPeerAnonPills() {
      const container = document.getElementById('peer-anon-member-pills');
      if (!container) return;
      let html = `
        <span class="text-xs font-bold text-[#8C837C] mr-1 flex items-center gap-1.5 uppercase tracking-wider">
          <i data-lucide="user" class="w-3.5 h-3.5 text-[#557A61]"></i> 選擇員工報告：
        </span>
      `;
      ALL_MEMBERS.forEach(m => {
        html += `
          <button onclick="selectPeerAnonMember('${m}')" id="peer-anon-btn-${m}" class="peer-anon-pill-btn px-5 py-2 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
            ${m} (${JOB_ROLES_MAP[m] || ''})
          </button>
        `;
      });
      container.innerHTML = html;
      selectPeerAnonMember(currentPeerAnonMember);
    }

    function selectPeerAnonMember(name) {
      currentPeerAnonMember = name;
      document.querySelectorAll('.peer-anon-pill-btn').forEach(btn => {
        btn.classList.remove('bg-[#557A61]', 'text-white', 'font-semibold');
        btn.classList.add('bg-[#F2EEE6]', 'text-[#4A433E]', 'font-medium');
      });
      const activeBtn = document.getElementById('peer-anon-btn-' + name);
      if (activeBtn) {
        activeBtn.classList.add('bg-[#557A61]', 'text-white', 'font-semibold');
        activeBtn.classList.remove('bg-[#F2EEE6]', 'text-[#4A433E]', 'font-medium');
      }
      renderPeerAnonSection();
    }

    function renderPeerAnonSection() {
      const name = currentPeerAnonMember;
      const peerRecords = RAW_DATA.filter(e => e.relation === "同事" && e.target === name);
      const container = document.getElementById('peer-anon-report-container');
      const exportContainer = document.getElementById('peer-anon-export-btn-container');

      if (exportContainer) {
        exportContainer.innerHTML = `
          <button onclick="exportSinglePeerAnonymousExcelClientSide('${name}')" class="inline-flex items-center gap-2 px-5 py-2.5 text-xs sm:text-sm font-semibold rounded-xl bg-[#557A61] text-white hover:bg-[#466551] transition shadow-xs">
            <i data-lucide="file-spreadsheet" class="w-4 h-4"></i> 下載【${name}】同儕匿名報告 (XLSX)
          </button>
        `;
      }

      let numPeers = peerRecords.length;

      let html = `
        <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] p-6 sm:p-8 soft-card-shadow space-y-7">
          <div class="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-[#E8E2D8] pb-5">
            <div class="flex items-center gap-4">
              <div class="w-14 h-14 rounded-2xl bg-[#557A61] text-white font-bold flex items-center justify-center text-xl font-serif-tc shadow-xs">${name.slice(0, 1)}</div>
              <div>
                <div class="flex items-center gap-2.5">
                  <h2 class="text-xl sm:text-2xl font-bold text-[#2E2827] font-serif-tc">${name} 同儕匿名評估回饋表</h2>
                  <span class="px-3 py-1 text-xs font-semibold rounded-lg bg-[#FCE5CD] text-[#783E16] border border-[#F3D1B0]">${JOB_ROLES_MAP[name] || '好好團隊夥伴'}</span>
                </div>
                <p class="text-xs text-[#8C837C] mt-1">共收到 <b>${numPeers}</b> 位同儕夥伴填答（已進行全匿名化去識別處理）</p>
              </div>
            </div>
            <div class="text-xs sm:text-sm px-4 py-2 rounded-xl bg-[#E4ECD3] text-[#2D5239] font-bold">
              同儕填答份數：${numPeers} 份
            </div>
          </div>

          <div>
            <h3 class="text-sm sm:text-base font-bold text-[#2E2827] mb-3.5 flex items-center gap-2 font-serif-tc"><i data-lucide="list-checks" class="w-4 h-4 text-[#557A61]"></i> 一、各評估題目同儕評分明細（滿分 10 分）</h3>
            <div class="overflow-x-auto rounded-xl border border-[#E8E2D8]">
              <table class="w-full text-xs sm:text-sm text-left border-collapse">
                <thead class="bg-[#FCE5CD] text-[#4A2E1C]">
                  <tr>
                    <th class="py-3 px-3.5 font-bold text-center w-14 border-r border-[#E8E2D8]">題號</th>
                    <th class="py-3 px-3.5 font-bold text-center w-24 border-r border-[#E8E2D8]">面向</th>
                    <th class="py-3 px-4 font-bold border-r border-[#E8E2D8]">題目說明</th>
                    <th class="py-3 px-3.5 font-bold text-center w-28 border-r border-[#E8E2D8] bg-[#E4ECD3]/60 text-[#2D5239]">同儕平均</th>
                    ${Array.from({ length: Math.max(numPeers, 1) }).map((_, i) => `<th class="py-3 px-3 font-bold text-center w-16 border-r border-[#E8E2D8]">同儕 ${String.fromCharCode(65+i)}</th>`).join('')}
                  </tr>
                </thead>
                <tbody class="divide-y divide-[#EFEAE1] bg-white">
                  ${PEER_QUESTIONS.map(([qNo, qCat, qDesc, qKey]) => {
                    const scores = peerRecords.map(r => r.peer_eval ? r.peer_eval[qKey] : null);
                    const validScores = scores.filter(s => s !== null && s !== undefined);
                    const avg = validScores.length ? (validScores.reduce((a, b) => a + b, 0) / validScores.length).toFixed(1) : "-";
                    return `
                      <tr class="hover:bg-[#FAF7F2] transition">
                        <td class="py-3 px-3.5 text-center font-bold text-[#8C837C] border-r border-[#E8E2D8]">${qNo}</td>
                        <td class="py-3 px-3.5 text-center text-[#4A433E] border-r border-[#E8E2D8]">${qCat}</td>
                        <td class="py-3 px-4 text-[#2E2827] border-r border-[#E8E2D8]">${qDesc}</td>
                        <td class="py-3 px-3.5 text-center font-bold text-[#2D5239] bg-[#E4ECD3]/20 border-r border-[#E8E2D8]">${avg}</td>
                        ${numPeers > 0 ? scores.map(s => `<td class="py-3 px-3 text-center text-[#4A433E] border-r border-[#E8E2D8]">${s !== null && s !== undefined ? s : '-'}</td>`).join('') : '<td class="py-3 px-3 text-center text-[#8C837C]">-</td>'}
                      </tr>
                    `;
                  }).join('')}
                </tbody>
              </table>
            </div>
          </div>

          <div>
            <h3 class="text-sm sm:text-base font-bold text-[#2E2827] mb-3.5 flex items-center gap-2 font-serif-tc"><i data-lucide="message-square" class="w-4 h-4 text-[#557A61]"></i> 二、同儕質化匿名回饋彙整</h3>
            <div class="space-y-6">
              <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                <div class="text-xs sm:text-sm font-bold text-[#557A61] flex items-center gap-2"><i data-lucide="trending-up" class="w-4 h-4"></i> Q37. 工作與文化提升建議</div>
                <div class="space-y-2.5">
                  ${numPeers > 0 ? peerRecords.map((r, i) => `
                    <div class="bg-white rounded-xl p-4 border border-[#E8E2D8] shadow-2xs">
                      <span class="text-xs font-bold text-[#8C837C] block mb-1">同儕 ${String.fromCharCode(65+i)}：</span>
                      <p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${r.peer_eval?.q37_improvement_advice || '（無填寫）'}</p>
                    </div>
                  `).join('') : '<p class="text-xs text-[#8C837C]">目前尚無同儕回饋</p>'}
                </div>
              </div>
              <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><i data-lucide="message-circle" class="w-4 h-4 text-[#557A61]"></i> Q38. 其他補充評價與觀察</div>
                <div class="space-y-2.5">
                  ${numPeers > 0 ? peerRecords.map((r, i) => `
                    <div class="bg-white rounded-xl p-4 border border-[#E8E2D8] shadow-2xs">
                      <span class="text-xs font-bold text-[#8C837C] block mb-1">同儕 ${String.fromCharCode(65+i)}：</span>
                      <p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${r.peer_eval?.q38_other_comments || '（無填寫）'}</p>
                    </div>
                  `).join('') : '<p class="text-xs text-[#8C837C]">目前尚無同儕回饋</p>'}
                </div>
              </div>
              <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                <div class="text-xs sm:text-sm font-bold text-[#7A363A] flex items-center gap-2"><i data-lucide="award" class="w-4 h-4 text-[#D48B7B]"></i> Q39. 肯定與感謝的話（好好星光大賞）</div>
                <div class="space-y-2.5">
                  ${numPeers > 0 ? peerRecords.map((r, i) => `
                    <div class="bg-white rounded-xl p-4 border border-[#F4CCCC] shadow-2xs">
                      <span class="text-xs font-bold text-[#8C5558] block mb-1">同儕 ${String.fromCharCode(65+i)}：</span>
                      <p class="text-xs sm:text-sm text-[#592629] font-medium leading-relaxed">${r.peer_eval?.q39_starlight_thanks || '（無填寫）'}</p>
                    </div>
                  `).join('') : '<p class="text-xs text-[#8C837C]">目前尚無同儕回饋</p>'}
                </div>
              </div>
            </div>
          </div>
        </div>
      `;

      if (container) container.innerHTML = html;
      if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    // ==========================================
    // TAB: 評同事
    // ==========================================
    function initPeerPills() {
      const peers = Array.from(new Set(RAW_DATA.filter(e => e.relation === "同事").map(e => e.target)));
      const container = document.getElementById('peer-pills-container');
      if (!container) return;
      let html = `
        <span class="text-xs font-bold text-[#8C837C] mr-1 flex items-center gap-1.5 uppercase tracking-wider"><i data-lucide="user" class="w-3.5 h-3.5 text-[#557A61]"></i> 受評同事：</span>
        <button onclick="filterPeer('ALL')" id="peer-btn-ALL" class="peer-pill-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold bg-[#557A61] text-white shadow-xs transition">全部 (${RAW_DATA.filter(e => e.relation === "同事").length}筆)</button>
      `;
      peers.forEach(p => {
        const count = RAW_DATA.filter(e => e.relation === "同事" && e.target === p).length;
        html += `
          <button onclick="filterPeer('${p}')" id="peer-btn-${p}" class="peer-pill-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
            ${p} (${count})
          </button>
        `;
      });
      container.innerHTML = html;
      if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    function filterPeer(name) {
      currentPeerFilter = name;
      document.querySelectorAll('.peer-pill-btn').forEach(btn => {
        btn.classList.remove('bg-[#557A61]', 'text-white', 'font-semibold');
        btn.classList.add('bg-[#F2EEE6]', 'text-[#4A433E]', 'font-medium');
      });
      const activeBtn = document.getElementById('peer-btn-' + name);
      if (activeBtn) {
        activeBtn.classList.add('bg-[#557A61]', 'text-white', 'font-semibold');
        activeBtn.classList.remove('bg-[#F2EEE6]', 'text-[#4A433E]', 'font-medium');
      }
      renderPeerSection();
    }

    function renderPeerSection() {
      let filtered = RAW_DATA.filter(e => e.relation === "同事");
      if (currentPeerFilter !== 'ALL') filtered = filtered.filter(e => e.target === currentPeerFilter);

      const avg = (key) => {
        const vals = filtered.map(e => e.peer_eval && e.peer_eval[key]).filter(v => v !== null && v !== undefined);
        return vals.length ? (vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(1) : 0;
      };

      const cultureLabels = ["多元 (接受意見)", "多元 (建設性觀點)", "實驗 (開放調整)", "信任 (分享經驗)", "肯定 (讚美同事)", "可持續 (尊重界線)"];
      const cultureData = [avg('q30_open_to_opposing'), avg('q31_constructive_opinions'), avg('q32_growth_mindset'), avg('q33_share_knowledge'), avg('q34_praise_peers'), avg('q35_boundary_respect')];

      const workLabels = ["合作狀況", "注重細節", "準時完成", "靈活調整", "追蹤承諾", "說明決策依據", "NPS推薦"];
      const workData = [avg('q24_cooperation'), avg('q25_detail_oriented'), avg('q26_on_time'), avg('q27_flexibility'), avg('q28_follow_up'), avg('q29_transparency'), avg('q36_nps_recommend')];

      const radarEl = document.getElementById('peerRadarChart');
      if (radarEl && typeof Chart !== 'undefined') {
        if (peerRadar) peerRadar.destroy();
        peerRadar = new Chart(radarEl.getContext('2d'), {
          type: 'radar',
          data: {
            labels: cultureLabels,
            datasets: [{
              label: currentPeerFilter === 'ALL' ? '全體同事平均' : currentPeerFilter,
              data: cultureData,
              backgroundColor: 'rgba(85, 122, 97, 0.2)',
              borderColor: '#557A61',
              pointBackgroundColor: '#557A61'
            }]
          },
          options: { scales: { r: { min: 0, max: 10, ticks: { stepSize: 2 } } }, plugins: { legend: { display: false } } }
        });
      }

      const barEl = document.getElementById('peerBarChart');
      if (barEl && typeof Chart !== 'undefined') {
        if (peerBar) peerBar.destroy();
        peerBar = new Chart(barEl.getContext('2d'), {
          type: 'bar',
          data: { labels: workLabels, datasets: [{ data: workData, backgroundColor: '#557A61', borderRadius: 6 }] },
          options: { scales: { y: { min: 0, max: 10 } }, plugins: { legend: { display: false } } }
        });
      }

      const feedbackContainer = document.getElementById('peer-feedback-list');
      if (feedbackContainer) {
        feedbackContainer.innerHTML = filtered.map(e => {
          const pe = e.peer_eval || {};
          return `
            <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] p-6 sm:p-8 soft-card-shadow space-y-6">
              <div class="flex items-center justify-between border-b border-[#E8E2D8] pb-4">
                <div class="flex items-center gap-3">
                  <div class="w-10 h-10 rounded-xl bg-[#557A61] text-white font-bold flex items-center justify-center text-sm font-serif-tc shadow-2xs">${e.target.slice(0, 1)}</div>
                  <div>
                    <span class="font-bold text-[#2E2827] text-base font-serif-tc">受評同事：${e.target}</span>
                    <span class="text-xs text-[#8C837C] ml-2">（填答者：${e.email}）</span>
                  </div>
                </div>
                <div class="flex items-center gap-3 text-xs sm:text-sm">
                  <span class="px-3.5 py-1.5 rounded-xl bg-[#E4ECD3] text-[#2D5239] font-bold">NPS 推薦：${pe.q36_nps_recommend || '-'} 分</span>
                  <span class="px-3.5 py-1.5 rounded-xl bg-[#F2EEE6] text-[#4A433E] font-medium">合作評分：${pe.q24_cooperation || '-'} 分</span>
                </div>
              </div>

              <div class="grid grid-cols-1 md:grid-cols-3 gap-5">
                <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                  <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><i data-lucide="trending-up" class="w-4 h-4 text-[#557A61]"></i> Q37. 工作與文化提升建議</div>
                  <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${pe.q37_improvement_advice || '（無）'}</p></div>
                </div>
                <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                  <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><i data-lucide="message-square" class="w-4 h-4 text-[#557A61]"></i> Q38. 其他補充評價</div>
                  <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${pe.q38_other_comments || '（無）'}</p></div>
                </div>
                <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                  <div class="text-xs sm:text-sm font-bold text-[#7A363A] flex items-center gap-2"><i data-lucide="award" class="w-4 h-4 text-[#D48B7B]"></i> Q39. 肯定與感謝詞</div>
                  <div class="bg-white rounded-xl p-4 my-2 border border-[#F4CCCC] shadow-2xs"><p class="text-xs sm:text-sm text-[#592629] font-medium leading-relaxed">${pe.q39_starlight_thanks || '（無）'}</p></div>
                </div>
              </div>
            </div>
          `;
        }).join('');
      }

      if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    // ==========================================
    // CLIENT-SIDE CSV PARSING
    // ==========================================
    function handleFileUpload(e) {
      if (e.target.files && e.target.files.length > 0) parseCSVFile(e.target.files[0]);
    }

    function parseCSVFile(file) {
      Papa.parse(file, {
        header: false,
        skipEmptyLines: true,
        complete: function(results) {
          if (!results.data || results.data.length < 2) {
            alert("CSV 內容為空或格式不正確！");
            return;
          }
          const parsed = parse2DArrayData(results.data);
          if (parsed && parsed.length > 0) {
            RAW_DATA = parsed;
            refreshAllViews();
            showToast(`成功載入本地 CSV！共更新 ${parsed.length} 筆填答紀錄。`);
          }
        }
      });
    }

    // ==========================================
    // CLIENT-SIDE XLSX EXPORT VIA EXCELJS
    // ==========================================
    const COLOR_PINK_BLUSH = "FFF4CCCC";
    const COLOR_PEACH_CREAM = "FFFCE5CD";
    const COLOR_WHITE = "FFFFFFFF";
    const COLOR_WARN_BG = "FFFFF2D6";
    const COLOR_SAGE_BG = "FFE4ECD3";

    const thinBorder = {
      top: { style: 'thin', color: { argb: 'FFD7CCC8' } },
      left: { style: 'thin', color: { argb: 'FFD7CCC8' } },
      bottom: { style: 'thin', color: { argb: 'FFD7CCC8' } },
      right: { style: 'thin', color: { argb: 'FFD7CCC8' } }
    };

    const fontSubHeader = { name: '微軟正黑體', size: 10, bold: true, color: { argb: 'FF4E342E' } };
    const fontBody = { name: '微軟正黑體', size: 9.5, color: { argb: 'FF2D2323' } };
    const fontBodyBold = { name: '微軟正黑體', size: 9.5, bold: true, color: { argb: 'FF2D2323' } };

    const alignCenter = { horizontal: 'center', vertical: 'middle', wrapText: true };
    const alignLeft = { horizontal: 'left', vertical: 'top', wrapText: true };
    const alignHeader = { horizontal: 'center', vertical: 'middle', wrapText: true };

    function styleRange(ws, startRow, startCol, endRow, endCol, font, fill, alignment) {
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          const cell = ws.getCell(r, c);
          if (font) cell.font = font;
          if (fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
          cell.border = thinBorder;
          if (alignment) cell.alignment = alignment;
        }
      }
    }

    function buildSubordinateComprehensiveSheet(wb, memberName, allEntries) {
      const ws = wb.addWorksheet(`【${memberName}】自評vs主管評`, { views: [{ showGridLines: true }] });
      const jobRole = JOB_ROLES_MAP[memberName] || "專案經理";
      const supName = MEMBER_SUPERVISOR_MAP[memberName] || "主管";
      const peerRecords = allEntries.filter(e => e.relation === "同事" && e.target === memberName);
      const numPeers = peerRecords.length;

      const selfEntry = allEntries.find(e => e.relation === "自評" && e.target === memberName);
      const hasSelf = Boolean(selfEntry && selfEntry.self_eval);
      const se = hasSelf ? selfEntry.self_eval : null;

      ws.columns = [{ width: 14 }, { width: 34 }, { width: 55 }, { width: 14 }, { width: 14 }, { width: 12 }, { width: 35 }];

      ws.mergeCells(1, 1, 1, 7);
      ws.getCell(1, 1).value = `好好星球文化基金會 360 年中成長評估 - 【${memberName}】部屬自評與主管評核對照表（主管專用）`;
      styleRange(ws, 1, 1, 1, 7, { name: '微軟正黑體', size: 13, bold: true, color: { argb: 'FF3E2723' } }, COLOR_PINK_BLUSH, { horizontal: 'left', vertical: 'middle', indent: 1 });
      ws.getRow(1).height = 30;

      ws.mergeCells(2, 1, 2, 7);
      ws.getCell(2, 1).value = `評估對象：${memberName}（${jobRole}） ｜ 直屬主管：${supName} ｜ 同儕樣本：${numPeers} 位 ｜ 自評狀態：${hasSelf ? '已完成自評' : '尚未自評'}`;
      styleRange(ws, 2, 1, 2, 7, { name: '微軟正黑體', size: 10, bold: true, color: { argb: 'FF4E342E' } }, COLOR_PEACH_CREAM, { horizontal: 'left', vertical: 'middle', indent: 1 });
      ws.getRow(2).height = 24;

      let rIdx = 3;
      // 1. 組織文化 (一列一個面向)
      ws.mergeCells(rIdx, 1, rIdx, 7);
      ws.getCell(rIdx, 1).value = "一、組織文化實踐：部屬自評實例 vs 主管評核回饋（一列一個面向）";
      styleRange(ws, rIdx, 1, rIdx, 7, fontSubHeader, COLOR_PEACH_CREAM, { horizontal: 'left', vertical: 'middle', indent: 1 });
      ws.getRow(rIdx).height = 24;
      rIdx++;

      const cultHeaders = ["文化面向", "文化定義與行為指引", "部屬自評實例 (STAR)", "自評等級", "主管評定", "落差分析", "主管評核回饋與觀察"];
      cultHeaders.forEach((h, i) => ws.getCell(rIdx, i + 1).value = h);
      styleRange(ws, rIdx, 1, rIdx, 7, fontSubHeader, COLOR_PEACH_CREAM, alignHeader);
      ws.getRow(rIdx).height = 24;
      rIdx++;

      const cultRows = [
        ["【信任】", "獨立行動與決策、主動協作、雙向溝通，在高彈性下給予彼此信任", se?.values?.['信任']],
        ["【多元】", "尊重差異、多元工作方法、主動表達不同觀點與想法", se?.values?.['多元']],
        ["【實驗】", "透過開放心態嘗試修正與反思，勇於檢討及給予回饋", se?.values?.['實驗']],
        ["【可持續】", "內在韌性、自我照顧、彈性的人際與工作邊界", se?.values?.['可持續']],
      ];

      cultRows.forEach(([cTitle, cDesc, cSelf]) => {
        ws.getCell(rIdx, 1).value = cTitle;
        ws.getCell(rIdx, 2).value = cDesc;
        ws.getCell(rIdx, 3).value = cSelf || "（部屬未填寫）";
        ws.getCell(rIdx, 4).value = "不適用";
        ws.getCell(rIdx, 5).value = "不適用";
        ws.getCell(rIdx, 6).value = "質化對齊";
        ws.getCell(rIdx, 7).value = SUPERVISOR_EVAL_STATE[memberName]?.[`文化_${cTitle.replace(/[【】]/g,'')}`]?.feedback || "【待主管填寫回饋】";

        styleRange(ws, rIdx, 1, rIdx, 7, fontBody, COLOR_WHITE, alignLeft);
        ws.getCell(rIdx, 1).alignment = alignCenter;
        ws.getCell(rIdx, 1).font = fontBodyBold;
        ws.getCell(rIdx, 4).alignment = alignCenter;
        ws.getCell(rIdx, 5).alignment = alignCenter;
        ws.getCell(rIdx, 6).alignment = alignCenter;
        ws.getRow(rIdx).height = Math.max(26, Math.min(100, Math.floor((cSelf || "").length / 35 * 18) + 16));
        rIdx++;
      });

      rIdx++;
      // 2. 專業職能 (一列一個面向，自評 vs 主管評並列)
      ws.mergeCells(rIdx, 1, rIdx, 7);
      ws.getCell(rIdx, 1).value = `二、專業職能：自評 vs 主管評分並列對照【${jobRole}】（逐項比對認知差異）`;
      styleRange(ws, rIdx, 1, rIdx, 7, fontSubHeader, COLOR_PEACH_CREAM, { horizontal: 'left', vertical: 'middle', indent: 1 });
      ws.getRow(rIdx).height = 24;
      rIdx++;

      const compHeaders = ["職能項目", "職能定義與說明", "部屬自評實例 (STAR)", "部屬自評", "主管評定 (L1~L5)", "落差分析", "主管評語與回饋 (Feedback)"];
      compHeaders.forEach((h, i) => ws.getCell(rIdx, i + 1).value = h);
      styleRange(ws, rIdx, 1, rIdx, 7, fontSubHeader, COLOR_PEACH_CREAM, alignHeader);
      ws.getRow(rIdx).height = 24;
      rIdx++;

      const compTitles = ROLE_COMPETENCIES[jobRole] || [];
      compTitles.forEach(cT => {
        let selfAns = null;
        if (hasSelf && se.competencies) {
          const item = se.competencies.find(x => x.title === cT);
          if (item) selfAns = item.answer;
        }

        const supLvl = SUPERVISOR_EVAL_STATE[memberName]?.[cT]?.level || "";
        const supFb = SUPERVISOR_EVAL_STATE[memberName]?.[cT]?.feedback || "【待主管填寫回饋】";

        ws.getCell(rIdx, 1).value = cT;
        ws.getCell(rIdx, 2).value = "核心專業職能";
        ws.getCell(rIdx, 3).value = selfAns || "（部屬未填寫自評實例）";
        ws.getCell(rIdx, 4).value = "待補齊";
        ws.getCell(rIdx, 5).value = supLvl;
        ws.getCell(rIdx, 6).value = supLvl ? "已評定" : "待主管評";
        ws.getCell(rIdx, 7).value = supFb;

        ws.getCell(rIdx, 5).dataValidation = {
          type: 'list',
          allowBlank: true,
          formulae: ['"L1,L2,L3,L4,L5"'],
          showErrorMessage: true,
          errorTitle: '評分無效',
          error: '請從下拉選單選取 L1 到 L5',
          promptTitle: '等級選單',
          prompt: '請選取：L1(Start), L2(Grow), L3(Keep), L4(Good), L5(Amazing!)'
        };

        styleRange(ws, rIdx, 1, rIdx, 7, fontBody, COLOR_WHITE, alignLeft);
        ws.getCell(rIdx, 1).font = fontBodyBold;
        ws.getCell(rIdx, 4).alignment = alignCenter;
        ws.getCell(rIdx, 5).alignment = alignCenter;
        ws.getCell(rIdx, 5).font = fontBodyBold;
        ws.getCell(rIdx, 6).alignment = alignCenter;

        ws.getRow(rIdx).height = Math.max(26, Math.min(100, Math.floor((selfAns || "").length / 35 * 18) + 16));
        rIdx++;
      });
    }

    async function exportSingleSubordinateComprehensiveExcelClientSide(memberName) {
      const wb = new ExcelJS.Workbook();
      buildSubordinateComprehensiveSheet(wb, memberName, RAW_DATA);
      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      saveAs(blob, `部屬自評vs主管評對照表_【${memberName}】.xlsx`);
      showToast(`已成功下載【${memberName}】自評vs主管評對照表 XLSX！`);
    }

    async function exportSupervisorTeamSubordinatesComprehensiveExcelClientSide(supName) {
      const wb = new ExcelJS.Workbook();
      const memberNames = SUPERVISOR_TEAMS[supName] || [];
      memberNames.forEach(m => buildSubordinateComprehensiveSheet(wb, m, RAW_DATA));
      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      saveAs(blob, `【${supName}主管專用】部屬自評vs主管評對照彙整包.xlsx`);
      showToast(`已成功下載【${supName}】團隊部屬自評vs主管評對照 Excel 包！`);
    }

    async function exportSinglePeerAnonymousExcelClientSide(memberName) {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet(`【${memberName}】同儕匿名回饋`);
      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      saveAs(blob, `同儕匿名回饋_【${memberName}】.xlsx`);
    }

    async function exportFullWorkbookClientSide() {
      const wb = new ExcelJS.Workbook();
      ALL_MEMBERS.forEach(m => buildSubordinateComprehensiveSheet(wb, m, RAW_DATA));
      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      saveAs(blob, "好好星球_360年中成長評估_主管分流與完整彙整表.xlsx");
      showToast("已成功下載全組織 Master Excel 總表！");
    }

    window.addEventListener('DOMContentLoaded', () => {
      // 1. First render default view
      refreshAllViews();
      // 2. Immediately try to fetch latest from Google Apps Script endpoint
      syncFromGoogleAppsScript(false);
    });
  </script>
</body>
</html>
"""

output_file = os.path.join(os.path.dirname(__file__), "..", "index.html")
with open(output_file, "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"Generated {output_file} successfully with Google Apps Script Sync Engine!")
