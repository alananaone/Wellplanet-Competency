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
  <!-- ExcelJS for full color/style Excel generation in browser matching VVN reference -->
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
      --bg-surface: #F2EEE6;
      --color-primary: #557A61;
      --color-primary-dark: #3F5E4A;
      --color-primary-light: #E4ECD3;
      --color-primary-text: #2D5239;
      --color-border: #E2DDD5;
      --color-border-subtle: #EDE8DF;
      --color-text-main: #2E2827;
      --color-text-muted: #6E6662;

      /* Reference Colors from VVN Excel */
      --vvn-pink-title: #F4CCCC;
      --vvn-peach-sec: #FCE5CD;
      --vvn-warm-feedback: #FFF2D6;
      --vvn-collab-blue: #CFE2F3;
      --vvn-culture-purple: #D9D2E9;
      --vvn-avg-green: #E4ECD3;
      --vvn-rating-l5: #E4EDF7;
      --vvn-rating-l4: #E3F1E6;
      --vvn-rating-l3: #FFF7E0;
      --vvn-rating-l2: #FCE8EC;
      --vvn-rating-l1: #FCE8EC;
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
    .custom-scrollbar::-webkit-scrollbar-thumb { background-color: #D6CEC3; border-radius: 4px; }

    .badge-sage { background-color: #E4ECD3; color: #2D5239; border: 1px solid #CDE0BC; }
    .badge-stone { background-color: #F2EEE6; color: #4A433E; border: 1px solid #E0D7CA; }
    .badge-peach { background-color: #FCE5CD; color: #4E342E; border: 1px solid #EAD1B8; }
    .badge-pink { background-color: #F4CCCC; color: #3E2723; border: 1px solid #E2B6B6; }

    .soft-card-shadow { box-shadow: 0 4px 16px -2px rgba(90, 80, 70, 0.05), 0 2px 6px -1px rgba(90, 80, 70, 0.03); }

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
      header, #dropZone, nav, #toast, .no-print, button, .tab-btn, footer, .screen-only-view, #selfCompetencyModal { display: none !important; }
      .print-only-doc { display: block !important; }
      main { padding: 0 !important; max-width: 100% !important; margin: 0 !important; }
      .tab-content { display: block !important; }
      .tab-content.hidden { display: none !important; }
      * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; box-shadow: none !important; border-radius: 0 !important; }
      .word-doc-table { width: 100% !important; border-collapse: collapse !important; margin-top: 8px !important; margin-bottom: 14px !important; font-size: 9pt !important; }
      .word-doc-table th, .word-doc-table td { border: 1px solid #555555 !important; padding: 5px 8px !important; vertical-align: middle !important; line-height: 1.4 !important; }
      .word-doc-table th { background-color: #FCE5CD !important; color: #4E342E !important; font-weight: bold !important; text-align: center !important; }
      .print-avoid-break, .word-doc-section, .word-feedback-block { break-inside: avoid !important; page-break-inside: avoid !important; }
      .word-sec-title { font-size: 11pt !important; font-weight: bold !important; color: #4E342E !important; background-color: #FCE5CD !important; padding: 4px 8px !important; border-left: 4px solid #557A61 !important; margin-top: 16px !important; margin-bottom: 8px !important; }
      .word-doc-divider { border-top: 2px solid #557A61 !important; border-bottom: 1px solid #557A61 !important; height: 3px !important; margin: 10px 0 14px 0 !important; }
      .print-footer { display: block !important; text-align: center; font-size: 8pt; color: #777777; margin-top: 20px; border-top: 1px solid #CCCCCC; padding-top: 6px; }
    }
  </style>
</head>
<body class="min-h-screen flex flex-col antialiased">

  <!-- TOP HEADER (NO-PRINT) -->
  <header class="bg-[#FFFDF9]/95 backdrop-blur-md border-b border-[#E2DDD5] sticky top-0 z-40 shadow-xs no-print">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
      <div class="flex items-center justify-between h-18 gap-4">
        <!-- BRAND -->
        <div class="flex items-center gap-3.5">
          <div class="w-11 h-11 rounded-2xl bg-[#557A61] flex items-center justify-center text-white font-bold shadow-xs text-sm tracking-wider font-serif-tc">
            360
          </div>
          <div>
            <div class="flex items-center gap-2.5 flex-wrap">
              <h1 class="text-base sm:text-lg font-bold text-[#2E2827] leading-tight font-serif-tc">好好星球文化基金會 360 年中成長評估</h1>
              <span class="px-2.5 py-0.5 text-xs font-semibold rounded-md bg-[#F2EEE6] text-[#4A433E] border border-[#E0D7CA]" id="header-data-source-badge">
                優先連線 Google 試算表
              </span>
            </div>
            <p class="text-xs text-[#7A726D] mt-0.5">主管評核 Excel 交付系統 ｜ 包含完整自評文字事證與題庫總表定義</p>
          </div>
        </div>

        <!-- ACTION BUTTONS -->
        <div class="flex items-center gap-2 sm:gap-3 flex-wrap justify-end">
          
          <!-- GAS SYNC BUTTON -->
          <button onclick="syncFromGoogleAppsScript(true)" id="gasSyncBtn" class="inline-flex items-center gap-2 px-3.5 sm:px-4 py-2.5 text-xs sm:text-sm font-semibold rounded-xl bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition border border-[#E0D7CA] shadow-2xs" title="優先讀取 Google 試算表（表單回覆1）">
            <i data-lucide="refresh-cw" class="w-4 h-4 text-[#557A61]" id="gasSyncIcon"></i>
            <span class="hidden lg:inline" id="gasSyncText">同步 Google 試算表</span>
            <span class="lg:hidden">同步</span>
          </button>

          <!-- DIRECT DOWNLOAD CURRENT SUPERVISOR EXCEL BUTTON -->
          <div class="relative inline-block text-left" id="exportDropdown">
            <button onclick="toggleDropdown()" class="inline-flex items-center gap-2 px-4 sm:px-5 py-2.5 text-xs sm:text-sm font-semibold rounded-xl bg-[#557A61] text-white hover:bg-[#466551] transition shadow-xs">
              <i data-lucide="file-spreadsheet" class="w-4 h-4"></i>
              <span class="hidden sm:inline">下載主管評估 Excel</span>
              <span class="sm:hidden">下載 Excel</span>
              <i data-lucide="chevron-down" class="w-4 h-4 opacity-80"></i>
            </button>
            
            <div id="dropdownMenu" class="hidden absolute right-0 mt-2 w-96 origin-top-right rounded-2xl bg-[#FFFDF9] p-3 shadow-2xl ring-1 ring-black/5 z-50 divide-y divide-[#EFEAE1] border border-[#E2DDD5]">
              <div class="py-2">
                <div class="px-3 py-1 text-[11px] font-bold text-[#8C837C] uppercase tracking-wider">
                  分主管 Excel 回覆包（含自評對照 ＋ 同儕評分分頁）
                </div>
                <button onclick="exportSupervisorTeamExcel('何維安')" class="w-full text-left flex items-center gap-3 px-3 py-2 text-xs sm:text-sm text-[#2E2827] hover:bg-[#F2EEE6] rounded-xl transition">
                  <div class="p-2 bg-[#FCE5CD] text-[#4E342E] rounded-xl"><i data-lucide="file-spreadsheet" class="w-4 h-4"></i></div>
                  <div>
                    <div class="font-bold text-[#2E2827]">VVN品牌主管專用_【主管評下屬用】</div>
                    <div class="text-[11px] text-[#7A726D]">部屬：林文琇（美感設計師）</div>
                  </div>
                </button>
                <button onclick="exportSupervisorTeamExcel('姚品瑄')" class="w-full text-left flex items-center gap-3 px-3 py-2 text-xs sm:text-sm text-[#2E2827] hover:bg-[#F2EEE6] rounded-xl transition">
                  <div class="p-2 bg-[#FCE5CD] text-[#4E342E] rounded-xl"><i data-lucide="file-spreadsheet" class="w-4 h-4"></i></div>
                  <div>
                    <div class="font-bold text-[#2E2827]">姚品瑄專案主管專用_【主管評下屬用】</div>
                    <div class="text-[11px] text-[#7A726D]">部屬：薛筑瑄、戴佑珍（專案經理）</div>
                  </div>
                </button>
                <button onclick="exportSupervisorTeamExcel('張希慈')" class="w-full text-left flex items-center gap-3 px-3 py-2 text-xs sm:text-sm text-[#2E2827] hover:bg-[#F2EEE6] rounded-xl transition">
                  <div class="p-2 bg-[#FCE5CD] text-[#4E342E] rounded-xl"><i data-lucide="file-spreadsheet" class="w-4 h-4"></i></div>
                  <div>
                    <div class="font-bold text-[#2E2827]">張希慈執行長主管專用_【主管評下屬用】</div>
                    <div class="text-[11px] text-[#7A726D]">何維安、陳泳璇、張芳媐、姚品瑄、胡喻翔</div>
                  </div>
                </button>
                <button onclick="exportSupervisorTeamExcel('張希慈_執行長')" class="w-full text-left flex items-center gap-3 px-3 py-2 text-xs sm:text-sm text-[#2E2827] hover:bg-[#F2EEE6] rounded-xl transition">
                  <div class="p-2 bg-[#FCE5CD] text-[#4E342E] rounded-xl"><i data-lucide="file-spreadsheet" class="w-4 h-4"></i></div>
                  <div>
                    <div class="font-bold text-[#2E2827]">張希慈個人自評_【主管評下屬用】</div>
                    <div class="text-[11px] text-[#7A726D]">執行長個人自評與同儕評估</div>
                  </div>
                </button>
              </div>

              <div class="py-2">
                <button onclick="exportFullMasterWorkbook()" class="w-full text-left flex items-center gap-3 px-3 py-2 text-xs sm:text-sm font-bold text-[#557A61] hover:bg-[#F2EEE6] rounded-xl transition">
                  <div class="p-2 bg-[#557A61] text-white rounded-xl"><i data-lucide="layers" class="w-4 h-4"></i></div>
                  <div>
                    <div>好好星球_360年中成長評估_【主管評下屬用】Master總表</div>
                    <div class="text-[11px] text-[#557A61] font-normal">包含所有 9 位成員之自評對照表與同儕評分分頁</div>
                  </div>
                </button>
              </div>
            </div>
          </div>

        </div>
      </div>

      <!-- MAIN TABS (NO-PRINT) -->
      <nav class="flex space-x-2 border-t border-[#E2DDD5] pt-2.5 -mb-px overflow-x-auto custom-scrollbar no-print">
        <button onclick="switchTab('subReview')" id="tab-btn-subReview" class="tab-btn inline-flex items-center gap-2 px-5 py-3 text-xs sm:text-sm font-semibold rounded-t-xl border-b-2 border-[#557A61] text-[#2D5239] bg-[#E4ECD3]/40 shrink-0 transition">
          <i data-lucide="file-text" class="w-4 h-4 text-[#557A61]"></i>
          自評 vs 主管評對照（主管填核回覆表）
          <span class="px-2 py-0.5 text-xs rounded-full bg-[#FCE5CD] text-[#4E342E] font-bold">主管核心</span>
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
    <!-- TAB: 自評 vs 主管評對照（主管填核回覆表） -->
    <!-- ========================================================================= -->
    <section id="tab-section-subReview" class="tab-content space-y-7">
      <div class="screen-only-view space-y-7">
        
        <!-- SUPERVISOR SELECTION FILTER BAR -->
        <div class="bg-[#FFFDF9] rounded-2xl p-6 border border-[#E2DDD5] soft-card-shadow flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div class="flex items-center gap-2.5 flex-wrap">
            <span class="text-xs font-bold text-[#8C837C] mr-1 flex items-center gap-1.5 uppercase tracking-wider">
              <i data-lucide="filter" class="w-3.5 h-3.5 text-[#557A61]"></i> 主管團隊：
            </span>
            <button onclick="filterSubReviewTeam('何維安')" id="sub-team-btn-何維安" class="sub-team-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold bg-[#557A61] text-white shadow-xs transition">
              何維安 的部屬（1人）
            </button>
            <button onclick="filterSubReviewTeam('姚品瑄')" id="sub-team-btn-姚品瑄" class="sub-team-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
              姚品瑄 的部屬（2人）
            </button>
            <button onclick="filterSubReviewTeam('張希慈')" id="sub-team-btn-張希慈" class="sub-team-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
              張希慈 的部屬（5人）
            </button>
            <button onclick="filterSubReviewTeam('張希慈_執行長')" id="sub-team-btn-張希慈_執行長" class="sub-team-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
              執行長個人自評
            </button>
            <button onclick="filterSubReviewTeam('ALL')" id="sub-team-btn-ALL" class="sub-team-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
              全體員工（9人）
            </button>
          </div>

          <!-- SIMPLIFIED DIRECT DOWNLOAD ACTIONS -->
          <div class="flex items-center gap-3">
            <button onclick="exportCurrentSubordinateExcel()" class="inline-flex items-center gap-2 px-4 py-2.5 text-xs sm:text-sm font-semibold rounded-xl bg-[#557A61] text-white hover:bg-[#466551] transition shadow-xs">
              <i data-lucide="download" class="w-4 h-4"></i>
              <span id="btn-sub-review-export-text">下載【林文琇】專用回覆 Excel</span>
            </button>
            <button onclick="exportCurrentSupervisorTeamExcel()" class="inline-flex items-center gap-2 px-4 py-2.5 text-xs sm:text-sm font-semibold rounded-xl bg-[#FCE5CD] text-[#4E342E] hover:bg-[#EAD1B8] transition border border-[#EAD1B8] shadow-2xs">
              <i data-lucide="file-spreadsheet" class="w-4 h-4"></i>
              <span id="btn-sup-team-export-text">下載【何維安】主管完整包</span>
            </button>
          </div>
        </div>

        <!-- SUBORDINATE MEMBER SELECTOR PILLS -->
        <div class="bg-[#FFFDF9] rounded-2xl p-5 border border-[#E2DDD5] soft-card-shadow flex items-center gap-3 flex-wrap" id="sub-review-member-pills"></div>

        <!-- REPORT CONTAINER -->
        <div id="sub-review-report-container" class="space-y-7"></div>
      </div>

      <!-- PRINT-ONLY WORD-STYLE CONTAINER -->
      <div id="sub-review-print-container" class="print-only-doc"></div>
    </section>

    <!-- ======================================================== -->
    <!-- TAB: 評主管 -->
    <!-- ======================================================== -->
    <section id="tab-section-supervisor" class="tab-content hidden space-y-7">
      <div class="screen-only-view space-y-7">
        <div class="bg-[#FFFDF9] rounded-2xl p-6 border border-[#E2DDD5] soft-card-shadow flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div class="flex items-center gap-2.5 flex-wrap">
            <span class="text-xs font-bold text-[#8C837C] mr-1 flex items-center gap-1.5 uppercase tracking-wider">
              <i data-lucide="user" class="w-3.5 h-3.5 text-[#557A61]"></i> 受評主管：
            </span>
            <button onclick="filterSupervisor('ALL')" id="sup-filter-btn-ALL" class="sup-filter-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
              全部主管（4筆）
            </button>
            <button onclick="filterSupervisor('張希慈')" id="sup-filter-btn-張希慈" class="sup-filter-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold bg-[#557A61] text-white shadow-xs transition">
              張希慈（3筆）
            </button>
            <button onclick="filterSupervisor('何維安')" id="sup-filter-btn-何維安" class="sup-filter-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
              何維安（1筆）
            </button>
          </div>
          <button onclick="window.print()" class="inline-flex items-center gap-2 px-5 py-2.5 text-xs sm:text-sm font-semibold rounded-xl bg-[#557A61] text-white hover:bg-[#466551] transition shadow-xs">
            <i data-lucide="printer" class="w-4 h-4"></i> 列印／存為 PDF（Cmd＋P）
          </button>
        </div>

        <div id="sup-stat-summary-cards" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-5"></div>

        <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <div class="bg-[#FFFDF9] p-6 rounded-2xl border border-[#E2DDD5] soft-card-shadow">
            <h3 class="text-sm sm:text-base font-bold text-[#2E2827] mb-3 flex items-center gap-2 font-serif-tc">
              <i data-lucide="compass" class="w-4 h-4 text-[#557A61]"></i> 四大文化實踐維度
            </h3>
            <div class="h-68 flex items-center justify-center">
              <canvas id="supervisorRadarChart"></canvas>
            </div>
          </div>
          <div class="bg-[#FFFDF9] p-6 rounded-2xl border border-[#E2DDD5] soft-card-shadow lg:col-span-2">
            <h3 class="text-sm sm:text-base font-bold text-[#2E2827] mb-3 flex items-center gap-2 font-serif-tc">
              <i data-lucide="bar-chart-3" class="w-4 h-4 text-[#557A61]"></i> 管理能力各題平均得分（滿分 10 分）
            </h3>
            <div class="h-68">
              <canvas id="supervisorBarChart"></canvas>
            </div>
          </div>
        </div>

        <div class="bg-[#FFFDF9] rounded-2xl border border-[#E2DDD5] p-6 sm:p-8 soft-card-shadow">
          <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-[#E2DDD5] pb-4 mb-5">
            <div class="flex items-center gap-2.5">
              <div class="p-2 bg-[#FCE5CD] text-[#4E342E] rounded-xl"><i data-lucide="table-properties" class="w-4 h-4"></i></div>
              <h3 class="text-base sm:text-lg font-bold text-[#2E2827] font-serif-tc">各題項評分統計表（平均分、最高最好、最低最差）</h3>
            </div>
            <span class="text-xs text-[#8C837C]">滿分 10 分 ｜ 標註綠色代表優異（≥9.0分），深灰代表有成長空間（≤7.0分）</span>
          </div>

          <div class="overflow-x-auto rounded-xl border border-[#E2DDD5]">
            <table class="w-full text-xs sm:text-sm text-left border-collapse" id="sup-item-stats-table">
              <thead class="bg-[#FCE5CD] text-[#4E342E]">
                <tr>
                  <th class="py-3 px-3.5 font-bold text-center w-14 border-r border-[#E2DDD5]">題號</th>
                  <th class="py-3 px-3.5 font-bold text-center w-28 border-r border-[#E2DDD5]">評估面向</th>
                  <th class="py-3 px-4 font-bold border-r border-[#E2DDD5]">題目內容與能力指引</th>
                  <th class="py-3 px-3.5 font-bold text-center w-24 border-r border-[#E2DDD5] bg-[#E4ECD3]/80 text-[#2D5239]">平均得分</th>
                  <th class="py-3 px-3.5 font-bold text-center w-24 border-r border-[#E2DDD5]">最好（最高）</th>
                  <th class="py-3 px-3.5 font-bold text-center w-24 border-r border-[#E2DDD5]">最差（最低）</th>
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
      <div class="bg-[#FFFDF9] rounded-2xl p-6 border border-[#E2DDD5] soft-card-shadow flex flex-col md:flex-row md:items-center justify-between gap-4 no-print">
        <div class="flex items-center gap-2.5 flex-wrap" id="self-supervisor-pills">
          <span class="text-xs font-bold text-[#8C837C] mr-1 flex items-center gap-1.5 uppercase tracking-wider">
            <i data-lucide="filter" class="w-3.5 h-3.5 text-[#557A61]"></i> 主管團隊：
          </span>
          <button onclick="filterSelfSupervisor('何維安')" id="self-sup-btn-何維安" class="self-sup-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold bg-[#557A61] text-white shadow-xs transition">
            何維安 的部屬（1人）
          </button>
          <button onclick="filterSelfSupervisor('姚品瑄')" id="self-sup-btn-姚品瑄" class="self-sup-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
            姚品瑄 的部屬（2人）
          </button>
          <button onclick="filterSelfSupervisor('張希慈')" id="self-sup-btn-張希慈" class="self-sup-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
            張希慈 的部屬（5人）
          </button>
          <button onclick="filterSelfSupervisor('張希慈_執行長')" id="self-sup-btn-張希慈_執行長" class="self-sup-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
            執行長個人自評
          </button>
          <button onclick="filterSelfSupervisor('ALL')" id="self-sup-btn-ALL" class="self-sup-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
            全部已自評（5人）
          </button>
        </div>

        <div id="self-export-btn-container"></div>
      </div>

      <div id="self-team-banner" class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-6 flex items-center justify-between">
        <div class="flex items-center gap-3.5">
          <div class="p-3 bg-[#FCE5CD] text-[#4E342E] rounded-xl shadow-2xs">
            <i data-lucide="users" class="w-5 h-5"></i>
          </div>
          <div>
            <h2 id="self-team-title" class="text-base sm:text-lg font-bold text-[#2E2827] font-serif-tc">何維安 的部屬自評列表</h2>
            <p id="self-team-subtitle" class="text-xs sm:text-sm text-[#6E6662] mt-0.5">涵蓋部屬：林文琇（美感設計師）</p>
          </div>
        </div>
        <div class="text-xs sm:text-sm px-4 py-2 bg-white rounded-full text-[#4E342E] font-bold border border-[#EAD1B8] shadow-2xs" id="self-completion-status">
          已填答 1 ／ 應填 1 人
        </div>
      </div>

      <div id="self-eval-cards-container" class="space-y-9"></div>
    </section>

    <!-- ======================================================== -->
    <!-- TAB: 員工同儕匿名表 (逐題明細) -->
    <!-- ======================================================== -->
    <section id="tab-section-peerAnon" class="tab-content hidden space-y-7">
      <div class="bg-[#FFFDF9] rounded-2xl p-6 border border-[#E2DDD5] soft-card-shadow flex flex-col md:flex-row md:items-center justify-between gap-4 no-print">
        <div class="flex items-center gap-2.5 flex-wrap" id="peer-anon-member-pills"></div>
        <div id="peer-anon-export-btn-container"></div>
      </div>

      <div id="peer-anon-report-container" class="space-y-7"></div>
    </section>

    <!-- ======================================================== -->
    <!-- TAB: 評同事 (同儕回饋總覽) -->
    <!-- ======================================================== -->
    <section id="tab-section-peer" class="tab-content hidden space-y-7">
      <div class="bg-[#FFFDF9] rounded-2xl p-6 border border-[#E2DDD5] soft-card-shadow flex flex-col md:flex-row md:items-center justify-between gap-4 no-print">
        <div class="flex items-center gap-2.5 flex-wrap" id="peer-pills-container"></div>
      </div>

      <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <div class="bg-[#FFFDF9] p-6 rounded-2xl border border-[#E2DDD5] soft-card-shadow">
          <h3 class="text-sm sm:text-base font-bold text-[#2E2827] mb-3 flex items-center gap-2 font-serif-tc">
            <i data-lucide="compass" class="w-4 h-4 text-[#557A61]"></i> 文化實踐維度表現
          </h3>
          <div class="h-68 flex items-center justify-center">
            <canvas id="peerRadarChart"></canvas>
          </div>
        </div>
        <div class="bg-[#FFFDF9] p-6 rounded-2xl border border-[#E2DDD5] soft-card-shadow lg:col-span-2">
          <h3 class="text-sm sm:text-base font-bold text-[#2E2827] mb-3 flex items-center gap-2 font-serif-tc">
            <i data-lucide="bar-chart-3" class="w-4 h-4 text-[#557A61]"></i> 協作、當責與溝通評分（滿分 10 分）
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

  <footer class="no-print bg-[#FFFDF9] border-t border-[#E2DDD5] py-7 text-center text-xs text-[#7A726D] mt-auto">
    好好星球文化基金會 360 年中成長評估系統 · 主管評核 Excel 交付系統 · 優先連線 Google 試算表（表單回覆1）
  </footer>

  <script>
    const GAS_ENDPOINT_URL = "https://script.google.com/macros/s/AKfycbxM-5YB3AX_CRK6APM3-dxGPUK7A2anQLrWSRwDK0_cZubdUu3pcUSl9lTPy5ahxXytgg/exec";

    // EMBEDDED DEFAULT / FALLBACK DATA
    let RAW_DATA = """ + raw_data_json + r""";

    // OFFICIAL ORG & TITLES
    const SUPERVISOR_TEAMS = {
      "何維安": ["林文琇"],
      "姚品瑄": ["薛筑瑄", "戴佑珍"],
      "張希慈": ["何維安", "陳泳璇", "張芳媐", "姚品瑄", "胡喻翔"],
      "張希慈_執行長": ["張希慈"],
      "ALL": ["林文琇", "薛筑瑄", "戴佑珍", "何維安", "陳泳璇", "張芳媐", "姚品瑄", "胡喻翔", "張希慈"]
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

    const ALL_MEMBERS = ["林文琇", "薛筑瑄", "戴佑珍", "何維安", "陳泳璇", "張芳媐", "姚品瑄", "胡喻翔", "張希慈"];

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

    const MEMBER_COMPETENCY_KEY = {
      "張希慈": "執行長",
      "陳泳璇": "行政經理",
      "林文琇": "美感設計師",
      "胡喻翔": "專案經理_CGL",
      "張芳媐": "營運經理兼執行長特助",
      "何維安": "品牌經理",
      "姚品瑄": "部門儲備主管",
      "薛筑瑄": "專案經理_CGL",
      "戴佑珍": "專案經理_SL"
    };

    // COMPLETE COMPETENCY DEFINITIONS EXTRACTED FROM 題庫總表
    const COMPETENCY_DEFINITIONS = {
      // 執行長
      "策略決策與組織方向設定": "1. 能將願景使命轉譯為可執行的年度策略目標，讓各部門知道自己的工作與組織方向的關係。\n2. 面對多個可行方向時，能說明選擇與捨棄的判準，而非僅宣布結論。\n3. 能整合各部門資料與不同利害關係人需求，提出資源分配邏輯。\n4. 能在環境變動時重新排序優先順序，並向團隊說明調整的理由。",
      "組織治理與財務風險管理": "1. 能維持董事會、主管機關與內部決策機制的正常運作與資訊透明。\n2. 能在重大決策前辨識法規、財務或聲譽風險，並提出因應方式。\n3. 能掌握組織整體財務狀況，在收支結構出現警訊時提前處理。\n4. 能建立可被檢驗的決策紀錄與授權範圍。",
      "關鍵利害關係人關係建立與維繫": "1. 能主動開發新的資源來源（人力、廠商、捐贈、補助、媒體）。\n2. 能維繫既有重要關係人的長期互動，而非僅在需要時才聯繫。\n3. 能在合作條件協商中兼顧組織利益與夥伴關係。\n4. 能為組織節省成本或爭取到原本不會有的資源。",
      "公關、演講與媒體關係": "1. 對外發言能清楚傳達基金會的理念，且與內部說法一致。\n2. 能因應不同受眾（媒體、捐款人、學校、政府）調整表達方式而不失核心訊息。\n3. 能在爭議或負面訊息出現時妥適回應，降低對組織的傷害。\n4. 公開露出能實際轉化為認識、信任或資源。",
      "核心團隊培養與組織文化建立": "1. 能辨識主管層的發展需求，給予具體的授權與練習機會。\n2. 能在關鍵決策上讓主管參與，而非全部由自己決定。\n3. 能在團隊出現文化落差時，直接處理而非迴避。\n4. 能建立讓夥伴敢於表達不同意見的討論場域。",

      // 行政經理
      "人事薪資與人資系統管理": "1. 準確執行薪資計算、勞健保與退休金投保作業。\n2. 維護人事資料與差勤紀錄的完整性與即時性。\n3. 確保人資系統的資料一致、可查、可交接。\n4. 能主動掌握法令變更並即時更新作業流程。",
      "法規與政府公文管理": "1. 掌握勞動法規、主管機關函文與法規異動。\n2. 於期限內完成公文收發、申報、陳報與回覆。\n3. 將法規與主管機關要求轉化為組織內部可執行的作業方式。\n4. 維護公文檔案與法定紀錄之完整保存。",
      "董事會與治理作業執行": "1. 依主管機關與組織章程要求，完成董事會會前準備、議程編排與資料寄送。\n2. 執行現場會議支援與精確完備之會議紀錄撰寫。\n3. 確實完成後續主管機關備查、變更登記等法定程序。\n4. 確保基金會法人治理文件與印鑑管理之合規與安全。",
      "財務核銷與內控執行": "1. 依核銷規範與補助專案要求審核單據憑證。\n2. 嚴格控管核銷流程、付款時程與合規性，及早辨識異常與風險。\n3. 確保帳務資料正確、透明且經得起內外部查核。\n4. 主動向專案同仁說明核銷規範與提供改善建議。",
      "總務採購與行政庶務管理": "1. 負責採購評估、廠商聯繫、詢比議價與發包作業。\n2. 建立並維持財產清冊、辦公物資與設備耗材管理。\n3. 維持辦公空間環境運作安全與日常行政庶務運作。\n4. 在成本控管、品質與時效之間做出合理且具說服力的判斷。",

      // 美感設計師
      "品牌識別系統設計與維護": "1. 設計 LOGO、標準字、品牌顏色與 CIS，確保品牌在名片、官網、招牌、包裝等載體一致。\n2. 能理解並實踐品牌定位與目標客群，將抽象策略轉化為具體視覺元素。\n3. 能建立並維護可供他人使用的視覺規範或模板。\n4. 能在既有識別系統下延伸出新的活動主題視覺，而不破壞一致性。",
      "視覺與美感設計實務": "1. 精通 Photoshop、Illustrator、Canva 等軟體，運用留白與字體設計建立層次感。\n2. 理解紙材、顏色模式（CMYK／RGB）及特殊加工，確保設計落地不失真。\n3. 能完成現場活動印刷品的設計、輸出、佈置與採購。\n4. 能在預算與時程限制下提出可行的材質與工法方案。",
      "需求釐清與創意提案": "1. 能在接案時主動問清楚使用情境、受眾與成功標準。\n2. 能蒐集參考資料並提出一個以上的方向供選擇。\n3. 能說明設計選擇背後的理由，而非只呈現成品。\n4. 能在收到修改意見時辨識真正的問題，而非逐條照改。",

      // 專案經理 (共通 & CGL / Soul LAB)
      "專案企劃與現場執行": "1. 事前與專案負責人及關係人確認任務細節，熟悉角色任務並預防突發狀況。\n2. 事中能獨立完成分內任務，並依時間與重要程度排序完成順序；能主動觀察服務對象、關係人及夥伴的需求，適時給予協助。\n3. 能將企劃內容落地成可執行的方案（流程表、物資、人力配置），具備從 0 到 1 企劃課程、活動與體驗的能力。",
      "專案時程與預算規劃管理": "1. 能盤點專案從起始到結尾的所有工作項目，評估各項所需時程。\n2. 能在專案中依實際狀況更新時程，即時與關係人同步。\n3. 能在專案啟動前合理分配各項目預算，過程中進行財務紀錄與評估。\n4. 能在時程或預算出現異常時排解問題，仍能在範圍內完成目標。",
      "需求研究與方案迭代": "1. 能設計並實施多元研究方法（深度訪談、問卷調查、現場觀察），主動蒐集服務對象的好奇與疑惑。\n2. 能識別、篩選並整理出有效資訊，不被個別意見帶著走；能從資料中歸納出使用者需求與行為模式。\n3. 能針對研究結果實際調整課程或方案，於下一次專案執行前完成迭代。",
      "外部夥伴關係經營": "1. 能向外部關係人（學校、講師、合作單位）提出清楚的合作需求與預期成果。\n2. 能考量對方的排程與習慣，在對方舒適的狀態下共同完成任務。\n3. 能讓實習生、志工與小組長有效分擔執行性工作，騰出正職做規劃型工作的時間。\n4. 能在外部夥伴結束合作時，讓他們對基金會留下正向感受。",
      "多元教學設計與現場引導": "1. 能依課程目標選擇合適的教學手段，並說明選擇理由。\n2. 能設計出有層次的課程流程（暖身、主活動、收斂、回顧）。\n3. 能在課程中安排讓孩子動手、動身體或表達的環節。\n4. 能說明自己的教育理念，並在課程設計中看得到這個理念。\n5. 能在現場主持引導課程進行，掌握節奏與時間。\n6. 能在課後回顧中辨識哪些設計有效、哪些需要調整。",
      "社群經營與培力需求回應": "1. 能創造與 Soul LAB 文化一致的社群氛圍（溫暖、療癒、支持）。\n2. 能讓服務對象在社群中彼此創造良好的社群氛圍（互助、回饋）。\n3. 主動觀察服務對象在社群內部的狀態、和 Soul LAB 互動的狀態。\n4. 能順暢和服務對象溝通、傳遞資訊，長期維持良好的互動關係。",

      // 營運經理兼執行長特助
      "組織營運流程設計與優化": "1. 能依組織目標規劃年度營運重點與執行節奏。\n2. 能設計並優化跨部門協作流程，減少重工與資訊落差。\n3. 能建立會議與資訊流通機制，讓決策所需資訊在會前就到位。\n4. 能建立並追蹤營運成效指標。\n5. 能在多專案並行時維持整體運作秩序。",
      "制度文件與 SOP 建置": "1. 能辨識哪些反覆發生的工作需要制度化，並產出可被他人使用的文件。\n2. 制度文件的用語與結構清楚，不需作者在旁說明也能執行。\n3. 能定期檢視既有制度是否還符合現況並進行更新。\n4. 能確認制度內容符合勞動法規與主管機關要求。",
      "人力資源策略與選用育留": "1. 能依組織發展階段規劃人力配置策略與職務說明書。\n2. 能設計招募與甄選流程，辨識與組織價值契合的人才。\n3. 能建立新人訓練與在職培訓機制。\n4. 能設計績效回饋與薪酬調整制度。\n5. 能評估人才流動對組織穩定與策略目標的影響。",
      "員工關係處理與勞動合規": "1. 能在衝突或申訴發生時依既定程序處理，並保護當事人的權益。\n2. 能依法辦理任用、留停、離職等人事程序，文件完備。\n3. 能主動掌握法規變動並調整內部作法。\n4. 能規劃並執行員工福利制度（聚會、旅行、年度活動、津貼）。",
      "組織文化落實與制度轉化": "1. 能把抽象的文化語言轉成具體可執行的制度或會議設計。\n2. 能設計讓夥伴實際練習文化行為的場合（工作坊、儀式、回饋機制）。\n3. 能觀察到文化落差的訊號，並提出結構性的處理方式。\n4. 能在制度設計時兼顧文化一致性，而非只求效率。",

      // 品牌經理
      "品牌定位與外部溝通一致性": "1. 能清楚界定組織使命、核心價值與品牌定位。\n2. 能建立並維護品牌識別系統（含視覺與敘事架構）。\n3. 能確保各部門對外傳遞一致的品牌語言與敘事。\n4. 能透過資訊透明與成效揭露建立公信力。",
      "行銷策略與議題倡議": "1. 能規劃整合行銷與議題倡議策略以提升品牌能見度。\n2. 能透過數位媒體與實體活動強化品牌認知與參與。\n3. 能透過數據分析與市場洞察調整品牌策略方向。\n4. 能運用品牌影響力吸引企業、政府與多元資源支持。",
      "品牌活動策劃與策展敘事": "1. 能策劃展覽、論壇或公共活動，將理念轉化為具體體驗。\n2. 能設計策展敘事脈絡，使受眾理解品牌價值。\n3. 能監督活動執行品質並管理協力廠商。\n4. 能在活動後評估成效並累積可複製的經驗。",
      "內部品牌管理與雇主品牌": "1. 能將品牌核心價值轉化為制度與行為準則。\n2. 能設計內部溝通機制以強化品牌認同。\n3. 能建立雇主品牌形象以吸引理念契合的人才。\n4. 能協助團隊成員成為品牌的倡議者與代言人。",
      "品牌危機與聲譽風險處理": "1. 能預判可能引發爭議的訊息或行動，事前調整。\n2. 能在負面訊息出現時快速判斷回應層級與方式。\n3. 能預防內外品牌形象失調並進行調整。\n4. 能建立危機處理的內部流程與發言原則。",

      // 部門儲備主管
      "部門策略規劃與專案組合管理": "1. 能制定部門年度發展策略，並確保各專案符合組織願景與策略方向。\n2. 能評估新專案的可行性與資源需求，做出啟動、調整或暫停的判斷。\n3. 能建立專案成效指標與定期檢核機制。\n4. 能在各專案之間協調資源與優先順序。",
      "專案經理管理與培育": "1. 能明確界定每位專案經理的權責範圍與目標，不留模糊地帶。\n2. 能定期進行專案檢核與回顧會議，而非只在出問題時介入。\n3. 能辨識專案經理在職能上的優勢與待加強處，給予具體回饋與指導。\n4. 能在專案經理遇到困難時給予支持，並協助排除跨部門或外部障礙。",
      "部門預算與資源配置管理": "1. 能統籌部門年度預算規劃，並在執行中控管落差。\n2. 能依專案優先順序分配人力與資源，並說明分配理由。\n3. 能監控部門整體投入與產出效益。\n4. 能在資源不足時做出取捨，而非平均分配。",
      "跨 LAB 專業掌握（CGL 教育專業與 SL 社群培力）": "1. 能理解 CGL 的教育設計邏輯（多元教學手段、現場引導、與兒童互動），並據以檢核課程品質。\n2. 能理解 Soul LAB 的社群培力邏輯（服務對象需求辨識、最小行動引導、社群氛圍營造），並據以檢核方案設計。\n3. 能辨識兩個 LAB 在方法上的差異，不用單一標準要求兩邊。\n4. 能促成兩個 LAB 之間的經驗交流與資源共用。",
      "利害關係人管理與衝突協調": "1. 能與外部合作單位及專案利害關係人進行有效溝通以促進專案運作。\n2. 能擔任團隊的溝通中心，促進成員互助合作、建立默契與信任感。\n3. 能有效化解專案團隊運作的各種危機。\n4. 能應用協商策略促進利害關係人支持專案的運作。"
    };

    const ROLE_COMPETENCIES_MAP = {
      "執行長": ["策略決策與組織方向設定", "組織治理與財務風險管理", "關鍵利害關係人關係建立與維繫", "公關、演講與媒體關係", "核心團隊培養與組織文化建立"],
      "行政經理": ["人事薪資與人資系統管理", "法規與政府公文管理", "董事會與治理作業執行", "財務核銷與內控執行", "總務採購與行政庶務管理"],
      "美感設計師": ["品牌識別系統設計與維護", "視覺與美感設計實務", "需求釐清與創意提案"],
      "專案經理_CGL": ["專案企劃與現場執行", "專案時程與預算規劃管理", "需求研究與方案迭代", "外部夥伴關係經營", "多元教學設計與現場引導"],
      "專案經理_SL": ["專案企劃與現場執行", "專案時程與預算規劃管理", "需求研究與方案迭代", "外部夥伴關係經營", "社群經營與培力需求回應"],
      "營運經理兼執行長特助": ["組織營運流程設計與優化", "制度文件與 SOP 建置", "人力資源策略與選用育留", "員工關係處理與勞動合規", "組織文化落實與制度轉化"],
      "品牌經理": ["品牌定位與外部溝通一致性", "行銷策略與議題倡議", "品牌活動策劃與策展敘事", "內部品牌管理與雇主品牌", "品牌危機與聲譽風險處理"],
      "部門儲備主管": ["部門策略規劃與專案組合管理", "專案經理管理與培育", "部門預算與資源配置管理", "跨 LAB 專業掌握（CGL 教育專業與 SL 社群培力）", "利害關係人管理與衝突協調"]
    };

    const RATING_STANDARDS = [
      { title: "Amazing!", lvl: "L5 (9.0-10.0)", desc: "遠超職位期待，表現為團隊之標竿與典範。", fb: "讚賞其突出貢獻，探討經驗複製機制。", color: "#E4EDF7" },
      { title: "Good", lvl: "L4 (7.0-8.9)", desc: "優於職位期待，持續展現高標準成果。", fb: "肯定並具體指出是哪些行為讓它超出標準，並設定具挑戰性的下一步目標。", color: "#E3F1E6" },
      { title: "Keep", lvl: "L3 (5.0-6.9)", desc: "符合職位門檻，展現穩定的工作交付。", fb: "確認穩定度，維持既有節奏，指出下一階可以再往前的地方，選一到兩項深化", color: "#FFF7E0" },
      { title: "Grow", lvl: "L2 (3.0-4.9)", desc: "部分符合，部分能力/行為仍在建立階段。", fb: "說明落差在哪，聚焦一項具體可練習的行為，納入下一期 IDP，設定可觀察的行為指標", color: "#FCE8EC" },
      { title: "Start", lvl: "L1（1.0-2.9)", desc: "目前的具體事證還看不到這項職能，或事證與職能要求落差明顯。", fb: "說明落差在哪，聚焦一項具體可練習的行為，納入下一期 IDP，設定可觀察的行為指標", color: "#FCE8EC" }
    ];

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

    // IN-MEMORY & LOCAL STORAGE
    const LOCAL_STORAGE_SUP_KEY = "WELLPLANET_SUPERVISOR_EVAL_STATE_V2";
    let SUPERVISOR_EVAL_STATE = {};

    let currentSubReviewTeam = '何維安';
    let currentSubReviewMember = '林文琇';
    let currentSupervisorFilter = '何維安';
    let currentSelfSupervisor = '何維安';
    let currentPeerFilter = 'ALL';
    let currentPeerAnonMember = '林文琇';

    let supervisorRadar = null;
    let supervisorBar = null;
    let peerRadar = null;
    let peerBar = null;

    function loadLocalStorageStates() {
      try {
        const savedSup = localStorage.getItem(LOCAL_STORAGE_SUP_KEY);
        if (savedSup) {
          SUPERVISOR_EVAL_STATE = JSON.parse(savedSup);
        }
      } catch (e) {
        console.warn("Failed to load SUPERVISOR_EVAL_STATE from localStorage", e);
      }
    }

    function persistSupervisorEvalState() {
      try {
        localStorage.setItem(LOCAL_STORAGE_SUP_KEY, JSON.stringify(SUPERVISOR_EVAL_STATE));
      } catch (e) {
        console.warn("Failed to save SUPERVISOR_EVAL_STATE to localStorage", e);
      }
    }

    function showToast(msg, isWarn = false) {
      const toast = document.getElementById('toast');
      if (!toast) return;
      document.getElementById('toast-msg').innerText = msg;
      const iconWrap = document.getElementById('toast-icon');
      if (iconWrap) {
        iconWrap.className = isWarn ? 'p-1 rounded-lg bg-[#4A433E] text-white' : 'p-1 rounded-lg bg-[#557A61] text-white';
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
        btn.classList.remove('border-[#557A61]', 'text-[#2D5239]', 'bg-[#E4ECD3]/40', 'font-semibold');
        btn.classList.add('border-transparent', 'text-[#6E6662]', 'font-medium');
      });
      const activeBtn = document.getElementById('tab-btn-' + tabId);
      if (activeBtn) {
        activeBtn.classList.add('border-[#557A61]', 'text-[#2D5239]', 'bg-[#E4ECD3]/40', 'font-semibold');
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
    // TAB: 自評 vs 主管評對照（主管填核回覆表）
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
        currentSubReviewMember = memberList[0] || '林文琇';
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
      persistSupervisorEvalState();
      renderSubReviewSection();
    }

    function updateSupervisorFeedback(memName, compTitle, text) {
      if (!SUPERVISOR_EVAL_STATE[memName]) SUPERVISOR_EVAL_STATE[memName] = {};
      if (!SUPERVISOR_EVAL_STATE[memName][compTitle]) SUPERVISOR_EVAL_STATE[memName][compTitle] = {};
      SUPERVISOR_EVAL_STATE[memName][compTitle].feedback = text;
      persistSupervisorEvalState();
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
              ${m}（${JOB_ROLES_MAP[m] || '職位'}）
            </button>
          `;
        });
        pillsContainer.innerHTML = phtml;
      }

      // Update Export Action Button Texts
      const btnSubExport = document.getElementById('btn-sub-review-export-text');
      if (btnSubExport) btnSubExport.innerText = `下載【${currentSubReviewMember}】專用回覆 Excel`;

      const btnSupExport = document.getElementById('btn-sup-team-export-text');
      if (btnSupExport) btnSupExport.innerText = `下載【${currentSubReviewTeam.replace('_執行長','')}】主管完整包`;

      const memName = currentSubReviewMember;
      const jobRole = JOB_ROLES_MAP[memName] || "專案經理";
      const supName = MEMBER_SUPERVISOR_MAP[memName] || "主管";

      const peerRecords = RAW_DATA.filter(e => e.relation === "同事" && e.target === memName);
      const numPeers = peerRecords.length;

      const selfEntry = RAW_DATA.find(e => e.relation === "自評" && e.target === memName);
      const hasSelf = Boolean(selfEntry && selfEntry.self_eval);
      const se = hasSelf ? selfEntry.self_eval : null;

      const roleKey = MEMBER_COMPETENCY_KEY[memName] || "專案經理_CGL";
      const compList = ROLE_COMPETENCIES_MAP[roleKey] || [];

      // SCREEN VIEW (ALL TEXTS VISIBLE, 4-COL CULTURE TABLE, ONLY COMPETENCY SCORE COLUMN SHOWS 尚不公布)
      const screenContainer = document.getElementById('sub-review-report-container');
      if (screenContainer) {
        screenContainer.innerHTML = `
          <!-- REPORT HEADER BANNER -->
          <div class="rounded-2xl border border-[#E2DDD5] p-6 sm:p-8 soft-card-shadow flex flex-col md:flex-row md:items-center justify-between gap-5 bg-[#FFFDF9]">
            <div class="flex items-center gap-4">
              <div class="w-14 h-14 rounded-2xl bg-[#F4CCCC] text-[#3E2723] font-bold flex items-center justify-center text-xl font-serif-tc shadow-xs border border-[#E2B6B6]">
                ${memName.slice(0, 1)}
              </div>
              <div>
                <div class="flex items-center gap-2.5 flex-wrap">
                  <h2 class="text-xl sm:text-2xl font-bold text-[#2E2827] font-serif-tc">【${memName}】部屬自評與主管評核對照表（主管專用）</h2>
                  <span class="px-3 py-1 text-xs font-semibold rounded-lg bg-[#FCE5CD] text-[#4E342E] border border-[#EAD1B8]">${jobRole}</span>
                  ${hasSelf ? `
                    <span class="px-2.5 py-0.5 text-xs font-medium rounded-full bg-[#E4ECD3] text-[#2D5239] border border-[#CDE0BC] flex items-center gap-1">
                      <i data-lucide="check-circle" class="w-3.5 h-3.5"></i> 已自評
                    </span>
                  ` : `
                    <span class="px-2.5 py-0.5 text-xs font-medium rounded-full bg-[#F2EEE6] text-[#6E6662] border border-[#E0D7CA] flex items-center gap-1">
                      <i data-lucide="clock" class="w-3.5 h-3.5"></i> 尚未自評
                    </span>
                  `}
                </div>
                <p class="text-xs text-[#7A726D] mt-1">評估對象：<b>${memName}（${jobRole}）</b> ｜ 直屬主管：<b>${supName}</b> ｜ 同儕填答樣本：<b>${numPeers}</b> 位 ｜ 自評狀態：<b>${hasSelf ? '已自評' : '尚未自評'}</b></p>
              </div>
            </div>
            
            <div class="flex items-center gap-3 text-xs sm:text-sm">
              <button onclick="exportCurrentSubordinateExcel()" class="inline-flex items-center gap-2 px-4 py-2.5 rounded-xl bg-[#557A61] text-white hover:bg-[#466551] transition font-bold shadow-xs">
                <i data-lucide="file-spreadsheet" class="w-4 h-4"></i> 下載【${memName}】回覆 Excel
              </button>
            </div>
          </div>

          <!-- PART 1: 組織文化實踐：部屬自評實例 vs 主管評核回饋 (4 欄：移除部屬自評與主管評定) -->
          <div class="space-y-4">
            <div class="bg-[#FCE5CD] px-5 py-3.5 rounded-2xl border border-[#EAD1B8] flex items-center justify-between">
              <div class="flex items-center gap-2.5">
                <i data-lucide="heart" class="w-5 h-5 text-[#4E342E]"></i>
                <h3 class="text-sm sm:text-base font-bold text-[#4E342E] font-serif-tc">一、組織文化實踐：部屬自評實例 vs 主管評核回饋（一列一個面向）</h3>
              </div>
            </div>

            <div class="overflow-x-auto rounded-2xl border border-[#E2DDD5] bg-[#FFFDF9] soft-card-shadow">
              <table class="w-full text-xs sm:text-sm text-left border-collapse">
                <thead class="bg-[#FCE5CD] text-[#4E342E]">
                  <tr>
                    <th class="py-3 px-4 font-bold text-center w-28 border-r border-[#E2DDD5]">文化面向</th>
                    <th class="py-3 px-4 font-bold border-r border-[#E2DDD5] w-72">文化定義與行為指引</th>
                    <th class="py-3 px-4 font-bold border-r border-[#E2DDD5]">部屬自評實例 (STAR)</th>
                    <th class="py-3 px-4 font-bold w-96 bg-[#FFF2D6] text-[#4E342E]">主管評核回饋與觀察</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-[#EFEAE1] bg-white">
                  ${[
                    ["【信任】", "獨立行動與決策、主動協作、雙向溝通，在高彈性下給予彼此信任", se?.values?.['信任']],
                    ["【多元】", "尊重差異、多元工作方法、主動表達不同觀點與想法", se?.values?.['多元']],
                    ["【實驗】", "透過開放心態嘗試修正與反思，勇於檢討及給予回饋", se?.values?.['實驗']],
                    ["【可持續】", "內在韌性、自我照顧、彈性的人際與工作邊界", se?.values?.['可持續']],
                  ].map(([cTitle, cDesc, cSelf]) => `
                    <tr class="hover:bg-[#FAF7F2] transition">
                      <td class="py-3.5 px-4 font-bold text-center text-[#2E2827] border-r border-[#E2DDD5]">${cTitle}</td>
                      <td class="py-3.5 px-4 text-[#4A433E] border-r border-[#E2DDD5] text-xs leading-relaxed">${cDesc}</td>
                      <td class="py-3.5 px-4 text-[#2E2827] border-r border-[#E2DDD5] leading-relaxed">
                        ${cSelf || '<span class="text-[#7A726D] italic">（部屬未填寫）</span>'}
                      </td>
                      <td class="py-3.5 px-4 bg-[#FFF2D6]/30">
                        <textarea onblur="updateSupervisorFeedback('${memName}', '文化_${cTitle.replace(/[【】]/g,'')}', this.value)" placeholder="請輸入主管針對【${cTitle}】的回饋與觀察..." class="w-full text-xs p-2.5 rounded-xl border border-[#ECD394] bg-[#FFF2D6]/60 focus:bg-white transition focus:outline-[#557A61] resize-y" rows="2">${SUPERVISOR_EVAL_STATE[memName]?.[`文化_${cTitle.replace(/[【】]/g,'')}`]?.feedback || ''}</textarea>
                      </td>
                    </tr>
                  `).join('')}
                </tbody>
              </table>
            </div>
          </div>

          <!-- PART 2: 專業職能：自評 vs 主管評分並列對照 (文字全部公布，等級統一寫「尚不公布」以避免影響主管評分) -->
          <div class="space-y-4">
            <div class="bg-[#FCE5CD] px-5 py-3.5 rounded-2xl border border-[#EAD1B8] flex items-center justify-between">
              <div class="flex items-center gap-2.5">
                <i data-lucide="briefcase" class="w-5 h-5 text-[#4E342E]"></i>
                <h3 class="text-sm sm:text-base font-bold text-[#4E342E] font-serif-tc">二、專業職能：自評 vs 主管評分並列對照【${jobRole}】（逐項比對認知差異）</h3>
              </div>
            </div>

            <div class="overflow-x-auto rounded-2xl border border-[#E2DDD5] bg-[#FFFDF9] soft-card-shadow">
              <table class="w-full text-xs sm:text-sm text-left border-collapse">
                <thead class="bg-[#FCE5CD] text-[#4E342E]">
                  <tr>
                    <th class="py-3 px-3.5 font-bold text-center w-40 border-r border-[#E2DDD5]">職能項目</th>
                    <th class="py-3 px-4 font-bold border-r border-[#E2DDD5] w-80">職能定義與說明（題庫總表）</th>
                    <th class="py-3 px-4 font-bold border-r border-[#E2DDD5]">部屬自評實例 (STAR)</th>
                    <th class="py-3 px-3 font-bold text-center w-28 border-r border-[#E2DDD5]">部屬自評</th>
                    <th class="py-3 px-3 font-bold text-center w-44 border-r border-[#E2DDD5] bg-[#FFF2D6] text-[#4E342E]">***主管評定 (L1~L5)</th>
                    <th class="py-3 px-4 font-bold w-80 bg-[#FFF2D6] text-[#4E342E]">***主管評語與回饋 (Feedback)</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-[#EFEAE1] bg-white">
                  ${compList.map(cTitle => {
                    const cDef = COMPETENCY_DEFINITIONS[cTitle] || "核心專業職能";
                    let selfAns = null;
                    if (hasSelf && se.competencies) {
                      const item = se.competencies.find(x => x.title === cTitle || cTitle.includes(x.title) || x.title.includes(cTitle));
                      if (item) selfAns = item.answer;
                    }
                    const supRating = SUPERVISOR_EVAL_STATE[memName]?.[cTitle]?.level || '';
                    const supFb = SUPERVISOR_EVAL_STATE[memName]?.[cTitle]?.feedback || '';

                    return `
                      <tr class="hover:bg-[#FAF7F2] transition">
                        <td class="py-3.5 px-3.5 font-bold text-[#2E2827] border-r border-[#E2DDD5]">${cTitle}</td>
                        <td class="py-3.5 px-4 text-[#4A433E] border-r border-[#E2DDD5] text-xs leading-relaxed whitespace-pre-line">${cDef}</td>
                        <td class="py-3.5 px-4 text-[#2E2827] border-r border-[#E2DDD5] leading-relaxed">
                          ${selfAns || '<span class="text-[#7A726D] italic">（部屬未填寫自評實例）</span>'}
                        </td>
                        <td class="py-3.5 px-3 text-center border-r border-[#E2DDD5] font-bold text-[#8C6D1F]">
                          <span class="bg-[#FFF2D6] px-2 py-0.5 rounded border border-[#ECD394]">尚不公布</span>
                        </td>
                        <td class="py-3.5 px-3 text-center border-r border-[#E2DDD5] bg-[#FFF2D6]/30">
                          <select onchange="updateSupervisorRating('${memName}', '${cTitle}', this.value)" class="text-xs font-bold p-2 rounded-xl border border-[#ECD394] bg-[#FFF2D6] focus:bg-white focus:outline-[#557A61] w-full">
                            <option value="">選取評級</option>
                            <option value="Amazing! (Lv5)" ${supRating.includes('Amazing')||supRating==='L5'?'selected':''}>Amazing! (Lv5)</option>
                            <option value="Good (L4)" ${supRating.includes('Good')||supRating==='L4'?'selected':''}>Good (L4)</option>
                            <option value="Keep (L3)" ${supRating.includes('Keep')||supRating==='L3'?'selected':''}>Keep (L3)</option>
                            <option value="Grow (L2)" ${supRating.includes('Grow')||supRating==='L2'?'selected':''}>Grow (L2)</option>
                            <option value="Start (Lv1)" ${supRating.includes('Start')||supRating==='L1'?'selected':''}>Start (Lv1)</option>
                            <option value="尚未評分" ${supRating==='尚未評分'?'selected':''}>尚未評分</option>
                          </select>
                        </td>
                        <td class="py-3.5 px-4 bg-[#FFF2D6]/30">
                          <textarea onblur="updateSupervisorFeedback('${memName}', '${cTitle}', this.value)" placeholder="請主管輸入針對此職能的具體評語與回饋..." class="w-full text-xs p-2 rounded-xl border border-[#ECD394] bg-[#FFF2D6]/60 focus:bg-white transition focus:outline-[#557A61] resize-y" rows="3">${supFb}</textarea>
                        </td>
                      </tr>
                    `;
                  }).join('')}
                </tbody>
              </table>
            </div>
          </div>

          <!-- PART 3: 評級標準 (三、評級標準) -->
          <div class="space-y-4">
            <div class="bg-[#FBE4EA] px-5 py-3.5 rounded-2xl border border-[#F5C2CD] flex items-center justify-between">
              <div class="flex items-center gap-2.5">
                <i data-lucide="award" class="w-5 h-5 text-[#8C1D40]"></i>
                <h3 class="text-sm sm:text-base font-bold text-[#8C1D40] font-serif-tc">三、評級標準（不對員工公布具體分數與總分，只回饋評級）</h3>
              </div>
            </div>

            <div class="overflow-x-auto rounded-2xl border border-[#E2DDD5] bg-[#FFFDF9] soft-card-shadow">
              <table class="w-full text-xs sm:text-sm text-left border-collapse">
                <thead class="bg-[#E8638A] text-white">
                  <tr>
                    <th class="py-3 px-4 font-bold text-center w-36 border-r border-white/20">評級（最終呈現）</th>
                    <th class="py-3 px-4 font-bold text-center w-40 border-r border-white/20">Level (分數落點）</th>
                    <th class="py-3 px-4 font-bold border-r border-white/20 w-80">定義</th>
                    <th class="py-3 px-4 font-bold">回饋方式參考</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-[#EFEAE1]">
                  ${RATING_STANDARDS.map(st => `
                    <tr style="background-color: ${st.color};" class="transition">
                      <td class="py-3 px-4 font-bold text-center text-[#2E2827] border-r border-[#D7CCC8]">${st.title}</td>
                      <td class="py-3 px-4 font-bold text-center text-[#4A433E] border-r border-[#D7CCC8]">${st.lvl}</td>
                      <td class="py-3 px-4 text-[#2E2827] border-r border-[#D7CCC8]">${st.desc}</td>
                      <td class="py-3 px-4 text-[#2E2827]">${st.fb}</td>
                    </tr>
                  `).join('')}
                </tbody>
              </table>
            </div>
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

      const overallAvg = allScoreValues.length ? (allScoreValues.reduce((a, b) => a + b, 0) / allScoreValues.length).toFixed(2) : "—";
      const validItems = itemStatResults.filter(it => it.avg !== null);
      const sortedByAvg = [...validItems].sort((a, b) => b.avg - a.avg);
      const topStrengths = sortedByAvg.slice(0, 3);
      const bottomOpportunities = [...sortedByAvg].reverse().slice(0, 3);

      const npsVals = filtered.map(e => e.supervisor_eval?.q18_nps_recommend).filter(v => v !== null && v !== undefined);
      const npsScore = npsVals.length ? (npsVals.reduce((a,b)=>a+b, 0) / npsVals.length).toFixed(2) : "—";
      const satVals = filtered.map(e => e.supervisor_eval?.q19_satisfaction).filter(v => v !== null && v !== undefined);
      const satScore = satVals.length ? (satVals.reduce((a,b)=>a+b, 0) / satVals.length).toFixed(2) : "—";

      const summaryCardsContainer = document.getElementById('sup-stat-summary-cards');
      if (summaryCardsContainer) {
        summaryCardsContainer.innerHTML = `
          <div class="bg-[#FFFDF9] rounded-2xl border border-[#E2DDD5] p-6 soft-card-shadow flex flex-col justify-between">
            <div class="flex items-center justify-between pb-2">
              <span class="text-xs font-bold text-[#8C837C] uppercase tracking-wider">整體評分平均</span>
              <div class="p-2 bg-[#FCE5CD] text-[#4E342E] rounded-xl"><i data-lucide="award" class="w-4 h-4"></i></div>
            </div>
            <div class="my-3">
              <div class="text-3xl sm:text-4xl font-bold text-[#2E2827] font-serif-tc">${overallAvg} <span class="text-xs text-[#8C837C] font-normal">／ 10 分</span></div>
            </div>
            <div class="text-xs text-[#7A726D] flex items-center justify-between pt-3 border-t border-[#EFEAE1]">
              <span>NPS 推薦：<b>${npsScore}</b> 分</span>
              <span>整體滿意：<b>${satScore}</b> 分</span>
            </div>
          </div>

          <div class="bg-[#FFFDF9] rounded-2xl border border-[#E2DDD5] p-6 soft-card-shadow flex flex-col justify-between">
            <div class="flex items-center justify-between pb-2">
              <span class="text-xs font-bold text-[#8C837C] uppercase tracking-wider">受評樣本數量</span>
              <div class="p-2 bg-[#F2EEE6] text-[#4A433E] rounded-xl"><i data-lucide="users" class="w-4 h-4"></i></div>
            </div>
            <div class="my-3">
              <div class="text-3xl sm:text-4xl font-bold text-[#2E2827] font-serif-tc">${filtered.length} <span class="text-xs text-[#8C837C] font-normal">份回覆</span></div>
            </div>
            <div class="text-xs text-[#7A726D] pt-3 border-t border-[#EFEAE1]">受評對象：<b>${supTitle}</b></div>
          </div>

          <div class="bg-[#FFFDF9] rounded-2xl border border-[#E2DDD5] p-6 soft-card-shadow flex flex-col justify-between">
            <div class="flex items-center justify-between pb-2">
              <span class="text-xs font-bold text-[#2D5239] uppercase tracking-wider flex items-center gap-1.5">
                <i data-lucide="sparkles" class="w-3.5 h-3.5 text-[#557A61]"></i> 表現最好項目（Top 3）
              </span>
              <span class="badge-sage px-2 py-0.5 text-[11px] font-bold rounded-md">亮點</span>
            </div>
            <div class="space-y-1.5 my-2">
              ${topStrengths.map(it => `
                <div class="text-xs flex items-center justify-between">
                  <span class="truncate text-[#2E2827] font-medium mr-2">${it.qNo}．${it.qDesc.split('（')[0]}</span>
                  <span class="badge-sage px-2 py-0.5 rounded font-bold shrink-0">${it.avg.toFixed(1)}分</span>
                </div>
              `).join('')}
            </div>
            <div class="text-[11px] text-[#7A726D] pt-2 border-t border-[#EFEAE1]">最高給分達 ${topStrengths[0] ? topStrengths[0].best : 10} 分</div>
          </div>

          <div class="bg-[#FFFDF9] rounded-2xl border border-[#E2DDD5] p-6 soft-card-shadow flex flex-col justify-between">
            <div class="flex items-center justify-between pb-2">
              <span class="text-xs font-bold text-[#4A433E] uppercase tracking-wider flex items-center gap-1.5">
                <i data-lucide="trending-up" class="w-3.5 h-3.5 text-[#557A61]"></i> 相對待提升項目
              </span>
              <span class="badge-stone px-2 py-0.5 text-[11px] font-bold rounded-md">成長</span>
            </div>
            <div class="space-y-1.5 my-2">
              ${bottomOpportunities.map(it => `
                <div class="text-xs flex items-center justify-between">
                  <span class="truncate text-[#2E2827] font-medium mr-2">${it.qNo}．${it.qDesc.split('（')[0]}</span>
                  <span class="badge-stone px-2 py-0.5 rounded font-bold shrink-0">${it.avg.toFixed(1)}分</span>
                </div>
              `).join('')}
            </div>
            <div class="text-[11px] text-[#7A726D] pt-2 border-t border-[#EFEAE1]">最低給分落點 ${bottomOpportunities[0] ? bottomOpportunities[0].worst : '—'} 分</div>
          </div>
        `;
      }

      const tbody = document.getElementById('sup-item-stats-tbody');
      if (tbody) {
        tbody.innerHTML = itemStatResults.map(it => {
          const avgDisplay = it.avg !== null ? it.avg.toFixed(2) : "—";
          const bestDisplay = it.best !== null ? it.best : "—";
          const worstDisplay = it.worst !== null ? it.worst : "—";
          const scoresDisplay = it.scores.length ? it.scores.join("、") : "—";

          let avgBadgeClass = "text-[#2E2827]";
          if (it.avg !== null) {
            if (it.avg >= 9.0) avgBadgeClass = "bg-[#E4ECD3] text-[#2D5239] font-bold px-2 py-0.5 rounded-md";
            else if (it.avg <= 7.0) avgBadgeClass = "bg-[#F2EEE6] text-[#4A433E] font-bold px-2 py-0.5 rounded-md";
          }

          return `
            <tr class="hover:bg-[#FAF7F2] transition">
              <td class="py-3 px-3.5 text-center font-bold text-[#8C837C] border-r border-[#E2DDD5]">${it.qNo}</td>
              <td class="py-3 px-3.5 text-center text-[#4A433E] border-r border-[#E2DDD5]">${it.qCat}</td>
              <td class="py-3 px-4 text-[#2E2827] border-r border-[#E2DDD5] font-medium">${it.qDesc}</td>
              <td class="py-3 px-3.5 text-center border-r border-[#E2DDD5]"><span class="${avgBadgeClass}">${avgDisplay}</span></td>
              <td class="py-3 px-3.5 text-center font-bold text-[#2D5239] border-r border-[#E2DDD5]">${bestDisplay}</td>
              <td class="py-3 px-3.5 text-center font-bold text-[#4A433E] border-r border-[#E2DDD5]">${worstDisplay}</td>
              <td class="py-3 px-4 text-center text-xs text-[#7A726D]">${scoresDisplay}</td>
            </tr>
          `;
        }).join('');
      }

      const avgOf = (k) => {
        const it = itemStatResults.find(x => x.qKey === k);
        return it && it.avg !== null ? it.avg.toFixed(1) : 0;
      };

      const cultureLabels = ["信任（真實表達）", "多元（聆聽意見）", "實驗（嘗試創新）", "心理安全（試錯空間）", "肯定（讚美認可）", "可持續（尊重界線）"];
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
            <div class="bg-[#FFFDF9] rounded-2xl border border-[#E2DDD5] p-6 sm:p-8 soft-card-shadow space-y-6">
              <div class="flex items-center justify-between border-b border-[#E2DDD5] pb-4">
                <div class="flex items-center gap-3">
                  <div class="w-10 h-10 rounded-xl bg-[#F4CCCC] text-[#3E2723] font-bold flex items-center justify-center text-sm font-serif-tc shadow-2xs border border-[#E2B6B6]">${e.target.slice(0, 1)}</div>
                  <div>
                    <span class="font-bold text-[#2E2827] text-base font-serif-tc">受評主管：${e.target}</span>
                    <span class="text-xs text-[#8C837C] ml-2">（填答者：${e.email}）</span>
                  </div>
                </div>
                <div class="flex items-center gap-3 text-xs sm:text-sm">
                  <span class="px-3.5 py-1.5 rounded-xl bg-[#E4ECD3] text-[#2D5239] font-bold">NPS 推薦：${se.q18_nps_recommend || '—'} 分</span>
                  <span class="px-3.5 py-1.5 rounded-xl bg-[#F2EEE6] text-[#4A433E] font-medium border border-[#E0D7CA]">滿意度：${se.q19_satisfaction || '—'} 分</span>
                </div>
              </div>

              <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                <div class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-5 sm:p-6 space-y-3">
                  <div class="text-xs sm:text-sm font-bold text-[#557A61] flex items-center gap-2"><i data-lucide="compass" class="w-4 h-4"></i> Q20．願景使命理解之引導</div>
                  <div class="bg-white rounded-xl p-4 my-2 border border-[#E2DDD5] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.q20_vision_mission || '（無）'}</p></div>
                </div>
                <div class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-5 sm:p-6 space-y-3">
                  <div class="text-xs sm:text-sm font-bold text-[#4A433E] flex items-center gap-2"><i data-lucide="trending-up" class="w-4 h-4 text-[#557A61]"></i> Q21．管理與文化精神建議</div>
                  <div class="bg-white rounded-xl p-4 my-2 border border-[#E2DDD5] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.q21_improvement_advice || '（無）'}</p></div>
                </div>
                <div class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-5 sm:p-6 space-y-3">
                  <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><i data-lucide="message-square" class="w-4 h-4 text-[#557A61]"></i> Q22．其他補充評價</div>
                  <div class="bg-white rounded-xl p-4 my-2 border border-[#E2DDD5] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.q22_other_comments || '（無）'}</p></div>
                </div>
                <div class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-5 sm:p-6 space-y-3">
                  <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><i data-lucide="award" class="w-4 h-4 text-[#557A61]"></i> Q23．肯定與感謝詞（好好星光大賞）</div>
                  <div class="bg-white rounded-xl p-4 my-2 border border-[#E2DDD5] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] font-medium leading-relaxed">${se.q23_starlight_thanks || '（無）'}</p></div>
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
        if (currentSelfSupervisor === '何維安') {
          bannerTitle.innerText = "【何維安】部屬自評彙整";
          bannerSub.innerText = "涵蓋部屬：林文琇（美感設計師）";
        } else if (currentSelfSupervisor === '姚品瑄') {
          bannerTitle.innerText = "【姚品瑄】部屬自評彙整";
          bannerSub.innerText = "涵蓋部屬：薛筑瑄（專案經理）、戴佑珍（專案經理）";
        } else if (currentSelfSupervisor === '張希慈') {
          bannerTitle.innerText = "【張希慈】部屬自評彙整";
          bannerSub.innerText = "涵蓋部屬：何維安（品牌經理）、陳泳璇（行政經理）、張芳媐（營運經理兼執行長特助）、姚品瑄（部門儲備主管）、胡喻翔（專案經理）";
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
          <button onclick="exportSupervisorTeamExcel('${currentSelfSupervisor}')" class="inline-flex items-center gap-2 px-5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold bg-[#557A61] text-white hover:bg-[#466551] transition shadow-xs">
            <i data-lucide="file-spreadsheet" class="w-4 h-4"></i>
            下載本組主管專用 XLSX 表格
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
            <div class="bg-[#FFFDF9] rounded-2xl border border-[#E2DDD5] soft-card-shadow overflow-hidden">
              <div class="bg-[#FCE5CD] px-6 py-5 flex flex-col md:flex-row md:items-center justify-between gap-3.5 border-b border-[#EAD1B8]">
                <div class="flex items-center gap-4">
                  <div class="w-12 h-12 rounded-2xl bg-[#F4CCCC] text-[#3E2723] font-bold flex items-center justify-center text-lg font-serif-tc shadow-2xs border border-[#E2B6B6]">${memName.slice(0, 1)}</div>
                  <div>
                    <div class="flex items-center gap-2.5">
                      <h3 class="text-base sm:text-lg font-bold text-[#2E2827] font-serif-tc">${memName}</h3>
                      <span class="px-3 py-1 text-xs font-semibold rounded-lg bg-white text-[#4A433E] border border-[#EAD1B8]">${JOB_ROLES_MAP[memName] || se.job_role || '自評'}</span>
                      <span class="px-2.5 py-0.5 text-xs font-medium rounded-full bg-[#E4ECD3] text-[#2D5239] border border-[#CDE0BC] flex items-center gap-1"><i data-lucide="check-circle" class="w-3.5 h-3.5"></i> 已填寫自評</span>
                    </div>
                    <p class="text-xs text-[#6E6662] mt-1">${entry.email} ｜ 填答時間：${entry.timestamp}</p>
                  </div>
                </div>
              </div>

              <div class="p-6 sm:p-8 space-y-7">
                <div>
                  <div class="text-xs sm:text-sm font-bold text-[#8C837C] uppercase tracking-wider mb-3.5 flex items-center gap-2 font-serif-tc"><i data-lucide="tag" class="w-4 h-4 text-[#557A61]"></i> 一、工作特質盤點</div>
                  <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                    <div class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-5 sm:p-6 space-y-3">
                      <div class="text-xs sm:text-sm font-bold text-[#2D5239] flex items-center gap-2"><i data-lucide="shield-check" class="w-4 h-4 text-[#557A61]"></i> 最穩定、最具代表性 Top 3</div>
                      <div class="flex flex-wrap gap-2.5 pt-1">${(se.top3_stable || []).map(t => `<span class="badge-sage px-3.5 py-1.5 text-xs sm:text-sm font-semibold rounded-xl shadow-2xs">${t}</span>`).join('')}</div>
                    </div>
                    <div class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-5 sm:p-6 space-y-3">
                      <div class="text-xs sm:text-sm font-bold text-[#4A433E] flex items-center gap-2"><i data-lucide="trending-up" class="w-4 h-4 text-[#557A61]"></i> 目前在練習／期望發展 3 項</div>
                      <div class="flex flex-wrap gap-2.5 pt-1">${(se.top3_practice || []).map(t => `<span class="badge-stone px-3.5 py-1.5 text-xs sm:text-sm font-semibold rounded-xl shadow-2xs">${t}</span>`).join('')}</div>
                    </div>
                  </div>
                </div>

                <div>
                  <div class="text-xs sm:text-sm font-bold text-[#8C837C] uppercase tracking-wider mb-3.5 flex items-center gap-2 font-serif-tc"><i data-lucide="heart" class="w-4 h-4 text-[#557A61]"></i> 二、四大文化實踐實例（STAR 敘述）</div>
                  <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                    <div class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-5 sm:p-6 space-y-3">
                      <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><span class="w-2.5 h-2.5 rounded-full bg-[#557A61]"></span> 信任（Trust）</div>
                      <div class="bg-white rounded-xl p-4 my-2 border border-[#E2DDD5] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.values?.['信任'] || '（未填寫）'}</p></div>
                    </div>
                    <div class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-5 sm:p-6 space-y-3">
                      <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><span class="w-2.5 h-2.5 rounded-full bg-[#557A61]"></span> 多元（Diversity）</div>
                      <div class="bg-white rounded-xl p-4 my-2 border border-[#E2DDD5] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.values?.['多元'] || '（未填寫）'}</p></div>
                    </div>
                    <div class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-5 sm:p-6 space-y-3">
                      <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><span class="w-2.5 h-2.5 rounded-full bg-[#557A61]"></span> 實驗（Experiment）</div>
                      <div class="bg-white rounded-xl p-4 my-2 border border-[#E2DDD5] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.values?.['實驗'] || '（未填寫）'}</p></div>
                    </div>
                    <div class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-5 sm:p-6 space-y-3">
                      <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><span class="w-2.5 h-2.5 rounded-full bg-[#557A61]"></span> 可持續（Sustainability）</div>
                      <div class="bg-white rounded-xl p-4 my-2 border border-[#E2DDD5] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.values?.['可持續'] || '（未填寫）'}</p></div>
                    </div>
                  </div>
                </div>

                ${se.competencies && se.competencies.length > 0 ? `
                  <div>
                    <div class="text-xs sm:text-sm font-bold text-[#8C837C] uppercase tracking-wider mb-3.5 flex items-center gap-2 font-serif-tc"><i data-lucide="briefcase" class="w-4 h-4 text-[#557A61]"></i> 三、${JOB_ROLES_MAP[memName] || se.job_role} 專屬職能展現實例</div>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                      ${se.competencies.map(c => `
                        <div class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-5 sm:p-6 space-y-3">
                          <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2 font-serif-tc"><span class="w-2.5 h-2.5 rounded-full bg-[#557A61]"></span> ${c.title}</div>
                          <div class="bg-white rounded-xl p-4 my-2 border border-[#E2DDD5] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed whitespace-pre-line">${c.answer || '（未填寫）'}</p></div>
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
                        <div class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-5 sm:p-6 space-y-3">
                          <div class="text-xs sm:text-sm font-bold text-[#4A433E] flex items-center gap-2 font-serif-tc"><i data-lucide="sparkle" class="w-3.5 h-3.5 text-[#557A61]"></i> ${k}</div>
                          <div class="bg-white rounded-xl p-4 my-2 border border-[#E2DDD5] shadow-2xs"><p class="text-xs sm:text-sm text-[#4A433E] leading-relaxed">${v || '（未填寫）'}</p></div>
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
            <div class="bg-[#FFFDF9] rounded-2xl border border-[#E2DDD5] soft-card-shadow p-7 flex items-center justify-between">
              <div class="flex items-center gap-4">
                <div class="w-12 h-12 rounded-2xl bg-[#F2EEE6] text-[#8C837C] font-bold flex items-center justify-center text-lg font-serif-tc">${memName.slice(0, 1)}</div>
                <div>
                  <div class="flex items-center gap-2.5">
                    <h3 class="text-base sm:text-lg font-bold text-[#2E2827] font-serif-tc">${memName}</h3>
                    <span class="px-3 py-1 text-xs font-semibold rounded-lg bg-[#F2EEE6] text-[#4A433E] border border-[#E0D7CA]">${JOB_ROLES_MAP[memName] || '職位'}</span>
                    <span class="px-3 py-1 text-xs font-medium rounded-full bg-[#F2EEE6] text-[#6E6662] border border-[#E0D7CA] flex items-center gap-1.5"><i data-lucide="clock" class="w-3.5 h-3.5"></i> 尚未收到自評資料</span>
                  </div>
                  <p class="text-xs sm:text-sm text-[#8C837C] mt-1">此成員尚未於表單中送出自評紀錄，收到後上傳新 CSV 即可同步更新。</p>
                </div>
              </div>
              <div class="text-xs font-semibold text-[#8C837C] px-3.5 py-1.5 bg-[#F2EEE6] rounded-xl">待填寫</div>
            </div>
          `;
        }
      });

      if (compStatus) compStatus.innerText = `已填答 ${completedCount} ／ 應填 ${memberNames.length} 人`;
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
            ${m}（${JOB_ROLES_MAP[m] || ''}）
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
          <button onclick="exportCurrentSubordinateExcel()" class="inline-flex items-center gap-2 px-5 py-2.5 text-xs sm:text-sm font-semibold rounded-xl bg-[#557A61] text-white hover:bg-[#466551] transition shadow-xs">
            <i data-lucide="file-spreadsheet" class="w-4 h-4"></i> 下載【${name}】專用回覆 XLSX
          </button>
        `;
      }

      let numPeers = peerRecords.length;

      let html = `
        <div class="bg-[#FFFDF9] rounded-2xl border border-[#E2DDD5] p-6 sm:p-8 soft-card-shadow space-y-7">
          <div class="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-[#E2DDD5] pb-5">
            <div class="flex items-center gap-4">
              <div class="w-14 h-14 rounded-2xl bg-[#F4CCCC] text-[#3E2723] font-bold flex items-center justify-center text-xl font-serif-tc shadow-xs border border-[#E2B6B6]">${name.slice(0, 1)}</div>
              <div>
                <div class="flex items-center gap-2.5">
                  <h2 class="text-xl sm:text-2xl font-bold text-[#2E2827] font-serif-tc">${name} 同儕評估與自評綜合報告（主管專用）</h2>
                  <span class="px-3 py-1 text-xs font-semibold rounded-lg bg-[#FCE5CD] text-[#4E342E] border border-[#EAD1B8]">${JOB_ROLES_MAP[name] || '好好團隊夥伴'}</span>
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
            <div class="overflow-x-auto rounded-xl border border-[#E2DDD5]">
              <table class="w-full text-xs sm:text-sm text-left border-collapse">
                <thead class="bg-[#FCE5CD] text-[#4E342E]">
                  <tr>
                    <th class="py-3 px-3.5 font-bold text-center w-14 border-r border-[#E2DDD5]">題號</th>
                    <th class="py-3 px-3.5 font-bold text-center w-24 border-r border-[#E2DDD5]">面向</th>
                    <th class="py-3 px-4 font-bold border-r border-[#E2DDD5]">題目說明</th>
                    <th class="py-3 px-3.5 font-bold text-center w-28 border-r border-[#E2DDD5] bg-[#E4ECD3]/60 text-[#2D5239]">同儕平均</th>
                    ${Array.from({ length: Math.max(numPeers, 1) }).map((_, i) => `<th class="py-3 px-3 font-bold text-center w-16 border-r border-[#E2DDD5]">同儕 ${String.fromCharCode(65+i)}</th>`).join('')}
                  </tr>
                </thead>
                <tbody class="divide-y divide-[#EFEAE1] bg-white">
                  ${PEER_QUESTIONS.map(([qNo, qCat, qDesc, qKey]) => {
                    const scores = peerRecords.map(r => r.peer_eval ? r.peer_eval[qKey] : null);
                    const validScores = scores.filter(s => s !== null && s !== undefined);
                    const avg = validScores.length ? (validScores.reduce((a, b) => a + b, 0) / validScores.length).toFixed(1) : "—";
                    return `
                      <tr class="hover:bg-[#FAF7F2] transition">
                        <td class="py-3 px-3.5 text-center font-bold text-[#8C837C] border-r border-[#E2DDD5]">${qNo}</td>
                        <td class="py-3 px-3.5 text-center text-[#4A433E] border-r border-[#E2DDD5]">${qCat}</td>
                        <td class="py-3 px-4 text-[#2E2827] border-r border-[#E2DDD5]">${qDesc}</td>
                        <td class="py-3 px-3.5 text-center font-bold text-[#2D5239] bg-[#E4ECD3]/20 border-r border-[#E2DDD5]">${avg}</td>
                        ${numPeers > 0 ? scores.map(s => `<td class="py-3 px-3 text-center text-[#4A433E] border-r border-[#E2DDD5]">${s !== null && s !== undefined ? s : '—'}</td>`).join('') : '<td class="py-3 px-3 text-center text-[#8C837C]">—</td>'}
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
              <div class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-5 sm:p-6 space-y-3">
                <div class="text-xs sm:text-sm font-bold text-[#557A61] flex items-center gap-2"><i data-lucide="trending-up" class="w-4 h-4"></i> Q37．工作與文化提升建議</div>
                <div class="space-y-2.5">
                  ${numPeers > 0 ? peerRecords.map((r, i) => `
                    <div class="bg-white rounded-xl p-4 border border-[#E2DDD5] shadow-2xs">
                      <span class="text-xs font-bold text-[#8C837C] block mb-1">同儕 ${String.fromCharCode(65+i)}：</span>
                      <p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${r.peer_eval?.q37_improvement_advice || '（無填寫）'}</p>
                    </div>
                  `).join('') : '<p class="text-xs text-[#8C837C]">目前尚無同儕回饋</p>'}
                </div>
              </div>
              <div class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-5 sm:p-6 space-y-3">
                <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><i data-lucide="message-circle" class="w-4 h-4 text-[#557A61]"></i> Q38．其他補充評價與觀察</div>
                <div class="space-y-2.5">
                  ${numPeers > 0 ? peerRecords.map((r, i) => `
                    <div class="bg-white rounded-xl p-4 border border-[#E2DDD5] shadow-2xs">
                      <span class="text-xs font-bold text-[#8C837C] block mb-1">同儕 ${String.fromCharCode(65+i)}：</span>
                      <p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${r.peer_eval?.q38_other_comments || '（無填寫）'}</p>
                    </div>
                  `).join('') : '<p class="text-xs text-[#8C837C]">目前尚無同儕回饋</p>'}
                </div>
              </div>
              <div class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-5 sm:p-6 space-y-3">
                <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><i data-lucide="award" class="w-4 h-4 text-[#557A61]"></i> Q39．肯定與感謝的話（好好星光大賞）</div>
                <div class="space-y-2.5">
                  ${numPeers > 0 ? peerRecords.map((r, i) => `
                    <div class="bg-white rounded-xl p-4 border border-[#E0D7CA] shadow-2xs">
                      <span class="text-xs font-bold text-[#6E6662] block mb-1">同儕 ${String.fromCharCode(65+i)}：</span>
                      <p class="text-xs sm:text-sm text-[#2E2827] font-medium leading-relaxed">${r.peer_eval?.q39_starlight_thanks || '（無填寫）'}</p>
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
        <button onclick="filterPeer('ALL')" id="peer-btn-ALL" class="peer-pill-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold bg-[#557A61] text-white shadow-xs transition">全部（${RAW_DATA.filter(e => e.relation === "同事").length}筆）</button>
      `;
      peers.forEach(p => {
        const count = RAW_DATA.filter(e => e.relation === "同事" && e.target === p).length;
        html += `
          <button onclick="filterPeer('${p}')" id="peer-btn-${p}" class="peer-pill-btn px-5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
            ${p}（${count}）
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

      const cultureLabels = ["多元（接受意見）", "多元（建設性觀點）", "實驗（開放調整）", "信任（分享經驗）", "肯定（讚美同事）", "可持續（尊重界線）"];
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
            <div class="bg-[#FFFDF9] rounded-2xl border border-[#E2DDD5] p-6 sm:p-8 soft-card-shadow space-y-6">
              <div class="flex items-center justify-between border-b border-[#E2DDD5] pb-4">
                <div class="flex items-center gap-3">
                  <div class="w-10 h-10 rounded-xl bg-[#F4CCCC] text-[#3E2723] font-bold flex items-center justify-center text-sm font-serif-tc shadow-2xs border border-[#E2B6B6]">${e.target.slice(0, 1)}</div>
                  <div>
                    <span class="font-bold text-[#2E2827] text-base font-serif-tc">受評同事：${e.target}</span>
                    <span class="text-xs text-[#8C837C] ml-2">（填答者：${e.email}）</span>
                  </div>
                </div>
                <div class="flex items-center gap-3 text-xs sm:text-sm">
                  <span class="px-3.5 py-1.5 rounded-xl bg-[#E4ECD3] text-[#2D5239] font-bold">NPS 推薦：${pe.q36_nps_recommend || '—'} 分</span>
                  <span class="px-3.5 py-1.5 rounded-xl bg-[#F2EEE6] text-[#4A433E] font-medium border border-[#E0D7CA]">合作評分：${pe.q24_cooperation || '—'} 分</span>
                </div>
              </div>

              <div class="grid grid-cols-1 md:grid-cols-3 gap-5">
                <div class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-5 sm:p-6 space-y-3">
                  <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><i data-lucide="trending-up" class="w-4 h-4 text-[#557A61]"></i> Q37．工作與文化提升建議</div>
                  <div class="bg-white rounded-xl p-4 my-2 border border-[#E2DDD5] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${pe.q37_improvement_advice || '（無）'}</p></div>
                </div>
                <div class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-5 sm:p-6 space-y-3">
                  <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><i data-lucide="message-square" class="w-4 h-4 text-[#557A61]"></i> Q38．其他補充評價</div>
                  <div class="bg-white rounded-xl p-4 my-2 border border-[#E2DDD5] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${pe.q38_other_comments || '（無）'}</p></div>
                </div>
                <div class="bg-[#FCE5CD]/40 border border-[#EAD1B8] rounded-2xl p-5 sm:p-6 space-y-3">
                  <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2"><i data-lucide="award" class="w-4 h-4 text-[#557A61]"></i> Q39．肯定與感謝詞</div>
                  <div class="bg-white rounded-xl p-4 my-2 border border-[#E2DDD5] shadow-2xs"><p class="text-xs sm:text-sm text-[#2E2827] font-medium leading-relaxed">${pe.q39_starlight_thanks || '（無）'}</p></div>
                </div>
              </div>
            </div>
          `;
        }).join('');
      }

      if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    // =========================================================================
    // EXCELJS ENGINE MATCHING VVN REFERENCE PALETTE & STRUCTURE EXACTLY
    // =========================================================================
    const COLOR_TITLE_BG = "FFF4CCCC";
    const COLOR_SECTION_BG = "FFFCE5CD";
    const COLOR_FEEDBACK_BG = "FFFFF2D6";
    const COLOR_COLLAB_BG = "FFCFE2F3";
    const COLOR_CULTURE_BG = "FFD9D2E9";
    const COLOR_AVG_BG = "FFE4ECD3";
    const COLOR_RATING_TITLE_BG = "FFFBE4EA";
    const COLOR_RATING_HDR_BG = "FFE8638A";
    const COLOR_L5_BG = "FFE4EDF7";
    const COLOR_L4_BG = "FFE3F1E6";
    const COLOR_L3_BG = "FFFFF7E0";
    const COLOR_L2_BG = "FFFCE8EC";
    const COLOR_L1_BG = "FFFCE8EC";
    const COLOR_WHITE = "FFFFFFFF";

    const thinBorder = {
      top: { style: 'thin', color: { argb: 'FFD7CCC8' } },
      left: { style: 'thin', color: { argb: 'FFD7CCC8' } },
      bottom: { style: 'thin', color: { argb: 'FFD7CCC8' } },
      right: { style: 'thin', color: { argb: 'FFD7CCC8' } }
    };

    const fontMainTitle = { name: '微軟正黑體', size: 13, bold: true, color: { argb: 'FF3E2723' } };
    const fontSubHeader = { name: '微軟正黑體', size: 10.5, bold: true, color: { argb: 'FF4E342E' } };
    const fontTableHdr = { name: '微軟正黑體', size: 10, bold: true, color: { argb: 'FF4E342E' } };
    const fontBody = { name: '微軟正黑體', size: 9.5, color: { argb: 'FF2D2323' } };
    const fontBodyBold = { name: '微軟正黑體', size: 9.5, bold: true, color: { argb: 'FF2D2323' } };
    const fontRatingHdr = { name: '微軟正黑體', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };

    function styleRange(ws, startRow, startCol, endRow, endCol, font, fillHex, alignment) {
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          const cell = ws.getCell(r, c);
          if (font) cell.font = font;
          if (fillHex) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillHex } };
          cell.border = thinBorder;
          if (alignment) cell.alignment = alignment;
        }
      }
    }

    function addRatingStandardsBlockToSheet(ws, startR) {
      let r = startR;
      ws.mergeCells(r, 1, r, 4);
      ws.getCell(r, 1).value = "三、評級標準（不對員工公布具體分數與總分，只回饋評級）";
      styleRange(ws, r, 1, r, 4, { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FF8C1D40' } }, COLOR_RATING_TITLE_BG, { horizontal: 'left', vertical: 'middle', indent: 1 });
      ws.getRow(r).height = 26;
      r++;

      const headers = ["評級（最終呈現）", "Level (分數落點）", "定義", "回饋方式參考"];
      headers.forEach((h, i) => ws.getCell(r, i + 1).value = h);
      styleRange(ws, r, 1, r, 4, fontRatingHdr, COLOR_RATING_HDR_BG, { horizontal: 'center', vertical: 'middle', wrapText: true });
      ws.getRow(r).height = 24;
      r++;

      const standards = [
        ["Amazing!", "L5 (9.0-10.0)", "遠超職位期待，表現為團隊之標竿與典範。", "讚賞其突出貢獻，探討經驗複製機制。", COLOR_L5_BG],
        ["Good", "L4 (7.0-8.9)", "優於職位期待，持續展現高標準成果。", "肯定並具體指出是哪些行為讓它超出標準，並設定具挑戰性的下一步目標。", COLOR_L4_BG],
        ["Keep", "L3 (5.0-6.9)", "符合職位門檻，展現穩定的工作交付。", "確認穩定度，維持既有節奏，指出下一階可以再往前的地方，選一到兩項深化", COLOR_L3_BG],
        ["Grow", "L2 (3.0-4.9)", "部分符合，部分能力/行為仍在建立階段。", "說明落差在哪，聚焦一項具體可練習的行為，納入下一期 IDP，設定可觀察的行為指標", COLOR_L2_BG],
        ["Start", "L1（1.0-2.9)", "目前的具體事證還看不到這項職能，或事證與職能要求落差明顯。", "說明落差在哪，聚焦一項具體可練習的行為，納入下一期 IDP，設定可觀察的行為指標", COLOR_L1_BG],
      ];

      standards.forEach(([t, lvl, d, fb, bg]) => {
        ws.getCell(r, 1).value = t;
        ws.getCell(r, 2).value = lvl;
        ws.getCell(r, 3).value = d;
        ws.getCell(r, 4).value = fb;

        styleRange(ws, r, 1, r, 4, fontBody, bg, { horizontal: 'left', vertical: 'top', wrapText: true });
        ws.getCell(r, 1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        ws.getCell(r, 1).font = fontBodyBold;
        ws.getCell(r, 2).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        ws.getRow(r).height = 30;
        r++;
      });

      return r;
    }

    function buildSubordinateSupervisorWorksheet(wb, memberName) {
      const ws = wb.addWorksheet(`【${memberName}】自評vs主管評`, { views: [{ showGridLines: true }] });
      const jobRole = JOB_ROLES_MAP[memberName] || "專案經理";
      const supName = MEMBER_SUPERVISOR_MAP[memberName] || "主管";

      const peerRecords = RAW_DATA.filter(e => e.relation === "同事" && e.target === memberName);
      const numPeers = peerRecords.length;

      const selfEntry = RAW_DATA.find(e => e.relation === "自評" && e.target === memberName);
      const hasSelf = Boolean(selfEntry && selfEntry.self_eval);
      const se = hasSelf ? selfEntry.self_eval : null;

      ws.columns = [{ width: 18 }, { width: 45 }, { width: 45 }, { width: 14 }, { width: 20 }, { width: 42 }];

      // Title
      ws.mergeCells(1, 1, 1, 6);
      ws.getCell(1, 1).value = `好好星球文化基金會 360 年中成長評估 - 【${memberName}】部屬自評與主管評核對照表（主管專用）`;
      styleRange(ws, 1, 1, 1, 6, fontMainTitle, COLOR_TITLE_BG, { horizontal: 'left', vertical: 'middle', indent: 1 });
      ws.getRow(1).height = 30;

      // Metadata
      ws.mergeCells(2, 1, 2, 6);
      const statusStr = hasSelf ? "已自評" : "尚未自評";
      ws.getCell(2, 1).value = `評估對象：${memberName}（${jobRole}） ｜ 直屬主管：${supName} ｜ 同儕樣本：${numPeers} 位 ｜ 自評狀態：${statusStr}`;
      styleRange(ws, 2, 1, 2, 6, fontSubHeader, COLOR_SECTION_BG, { horizontal: 'left', vertical: 'middle', indent: 1 });
      ws.getRow(2).height = 24;

      let r = 3;
      // 1. 組織文化 (4 欄結構：移除「部屬自評」與「主管評定」，Col 4~6 合併為「主管評核回饋與觀察」)
      ws.mergeCells(r, 1, r, 6);
      ws.getCell(r, 1).value = "一、組織文化實踐：部屬自評實例 vs 主管評核回饋（一列一個面向）";
      styleRange(ws, r, 1, r, 6, fontSubHeader, COLOR_SECTION_BG, { horizontal: 'left', vertical: 'middle', indent: 1 });
      ws.getRow(r).height = 24;
      r++;

      ws.getCell(r, 1).value = "文化面向";
      ws.getCell(r, 2).value = "文化定義與行為指引";
      ws.getCell(r, 3).value = "部屬自評實例 (STAR)";
      ws.mergeCells(r, 4, r, 6);
      ws.getCell(r, 4).value = "主管評核回饋與觀察";
      styleRange(ws, r, 1, r, 6, fontTableHdr, COLOR_SECTION_BG, { horizontal: 'center', vertical: 'middle', wrapText: true });
      ws.getRow(r).height = 24;
      r++;

      const cultRows = [
        ["【信任】", "獨立行動與決策、主動協作、雙向溝通，在高彈性下給予彼此信任", se?.values?.['信任']],
        ["【多元】", "尊重差異、多元工作方法、主動表達不同觀點與想法", se?.values?.['多元']],
        ["【實驗】", "透過開放心態嘗試修正與反思，勇於檢討及給予回饋", se?.values?.['實驗']],
        ["【可持續】", "內在韌性、自我照顧、彈性的人際與工作邊界", se?.values?.['可持續']],
      ];

      cultRows.forEach(([cTitle, cDesc, cSelf]) => {
        const selfVal = cSelf || "（部屬未填寫）";
        ws.getCell(r, 1).value = cTitle;
        ws.getCell(r, 2).value = cDesc;
        ws.getCell(r, 3).value = selfVal; // 全部公布文字
        ws.mergeCells(r, 4, r, 6);
        ws.getCell(r, 4).value = SUPERVISOR_EVAL_STATE[memberName]?.[`文化_${cTitle.replace(/[【】]/g,'')}`]?.feedback || "【待主管填寫回饋】";

        styleRange(ws, r, 1, r, 6, fontBody, COLOR_WHITE, { horizontal: 'left', vertical: 'top', wrapText: true });
        ws.getCell(r, 1).alignment = { horizontal: 'center', vertical: 'middle' };
        ws.getCell(r, 1).font = fontBodyBold;
        for (let c = 4; c <= 6; c++) {
          ws.getCell(r, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR_FEEDBACK_BG } };
        }
        ws.getRow(r).height = Math.max(30, Math.min(120, Math.floor((cSelf || "").length / 35 * 18) + 18));
        r++;
      });

      r++;
      // 2. 專業職能 (6 欄結構：文字全部公布，等級統一寫「尚不公布」以避免影響主管評分)
      ws.mergeCells(r, 1, r, 6);
      ws.getCell(r, 1).value = `二、專業職能：自評 vs 主管評分並列對照【${jobRole}】（逐項比對認知差異）`;
      styleRange(ws, r, 1, r, 6, fontSubHeader, COLOR_SECTION_BG, { horizontal: 'left', vertical: 'middle', indent: 1 });
      ws.getRow(r).height = 24;
      r++;

      const compHeaders = ["職能項目", "職能定義與說明", "部屬自評實例 (STAR)", "部屬自評", "***主管評定 (L1~L5)", "***主管評語與回饋 (Feedback)"];
      compHeaders.forEach((h, i) => ws.getCell(r, i + 1).value = h);
      styleRange(ws, r, 1, r, 6, fontTableHdr, COLOR_SECTION_BG, { horizontal: 'center', vertical: 'middle', wrapText: true });
      ws.getRow(r).height = 24;
      r++;

      const roleKey = MEMBER_COMPETENCY_KEY[memberName] || "專案經理_CGL";
      const compTitles = ROLE_COMPETENCIES_MAP[roleKey] || [];

      compTitles.forEach(cT => {
        const cDef = COMPETENCY_DEFINITIONS[cT] || "核心專業職能";
        let selfAns = null;
        if (hasSelf && se.competencies) {
          const item = se.competencies.find(x => x.title === cT || cT.includes(x.title) || x.title.includes(cT));
          if (item) selfAns = item.answer;
        }

        const selfStar = selfAns || "（部屬未填寫自評實例）";
        const selfLvl = "尚不公布"; // 等級統一寫「尚不公布」避免影響主管評分
        const supLvl = SUPERVISOR_EVAL_STATE[memberName]?.[cT]?.level || "尚未評分";
        const supFb = SUPERVISOR_EVAL_STATE[memberName]?.[cT]?.feedback || "【待主管填寫回饋】";

        ws.getCell(r, 1).value = cT;
        ws.getCell(r, 2).value = cDef;
        ws.getCell(r, 3).value = selfStar;
        ws.getCell(r, 4).value = selfLvl;
        ws.getCell(r, 5).value = supLvl;
        ws.getCell(r, 6).value = supFb;

        // Data validation dropdown for Supervisor Lv. 1~5
        ws.getCell(r, 5).dataValidation = {
          type: 'list',
          allowBlank: true,
          formulae: ['"Amazing! (Lv5),Good (L4),Keep (L3),Grow (L2),Start (Lv1),尚未評分"'],
          showErrorMessage: true,
          errorTitle: '評分無效',
          error: '請從下拉選單選取：Amazing! (Lv5)、Good (L4)、Keep (L3)、Grow (L2)、Start (Lv1)',
          promptTitle: '主管評級選單',
          prompt: '請選取評級：Amazing! (Lv5)、Good (L4)、Keep (L3)、Grow (L2)、Start (Lv1)'
        };

        styleRange(ws, r, 1, r, 6, fontBody, COLOR_WHITE, { horizontal: 'left', vertical: 'top', wrapText: true });
        ws.getCell(r, 1).font = fontBodyBold;
        ws.getCell(r, 4).alignment = { horizontal: 'center', vertical: 'middle' };
        ws.getCell(r, 4).font = fontBodyBold;
        ws.getCell(r, 5).alignment = { horizontal: 'center', vertical: 'middle' };
        ws.getCell(r, 5).font = fontBodyBold;
        ws.getCell(r, 5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR_FEEDBACK_BG } };
        ws.getCell(r, 6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR_FEEDBACK_BG } };

        ws.getRow(r).height = Math.max(40, Math.min(140, Math.floor(cDef.length / 25 * 18) + 18));
        r++;
      });

      r++;
      // 3. 評級標準
      addRatingStandardsBlockToSheet(ws, r);
    }

    function buildSubordinatePeerWorksheet(wb, memberName) {
      const ws = wb.addWorksheet(`【${memberName}】同儕評`, { views: [{ showGridLines: true }] });
      const jobRole = JOB_ROLES_MAP[memberName] || "專案經理";
      const supName = MEMBER_SUPERVISOR_MAP[memberName] || "主管";

      const peerRecords = RAW_DATA.filter(e => e.relation === "同事" && e.target === memberName);
      const numPeers = peerRecords.length;

      const selfEntry = RAW_DATA.find(e => e.relation === "自評" && e.target === memberName);
      const hasSelf = Boolean(selfEntry && selfEntry.self_eval);
      const se = hasSelf ? selfEntry.self_eval : null;

      ws.columns = [{ width: 14 }, { width: 20 }, { width: 48 }, { width: 15 }, { width: 15 }, { width: 15 }, { width: 28 }];

      // Title
      ws.mergeCells(1, 1, 1, 7);
      ws.getCell(1, 1).value = `好好星球文化基金會 360 年中成長評估 - 【${memberName}】部屬同儕評估與自評綜合報告（主管專用）`;
      styleRange(ws, 1, 1, 1, 7, fontMainTitle, COLOR_TITLE_BG, { horizontal: 'left', vertical: 'middle', indent: 1 });
      ws.getRow(1).height = 30;

      // Metadata
      ws.mergeCells(2, 1, 2, 7);
      const statusStr = hasSelf ? "已自評" : "尚未自評";
      ws.getCell(2, 1).value = `評估對象：${memberName}（${jobRole}） ｜ 直屬主管：${supName} ｜ 同儕填答樣本：${numPeers} 位 ｜ 自評狀態：${statusStr}`;
      styleRange(ws, 2, 1, 2, 7, fontSubHeader, COLOR_SECTION_BG, { horizontal: 'left', vertical: 'middle', indent: 1 });
      ws.getRow(2).height = 24;

      let r = 3;
      // 一、統計分數
      ws.mergeCells(r, 1, r, 7);
      ws.getCell(r, 1).value = "一、統計分數：同儕量化評估分析（平均分、最高分、最低分）";
      styleRange(ws, r, 1, r, 7, fontSubHeader, COLOR_SECTION_BG, { horizontal: 'left', vertical: 'middle', indent: 1 });
      ws.getRow(r).height = 24;
      r++;

      const headers = ["題號", "評估面向", "評估項目與題目說明", "同儕平均得分", "最高分", "最低分", "給分明細"];
      headers.forEach((h, i) => ws.getCell(r, i + 1).value = h);
      styleRange(ws, r, 1, r, 7, fontTableHdr, COLOR_SECTION_BG, { horizontal: 'center', vertical: 'middle', wrapText: true });
      ws.getRow(r).height = 24;
      r++;

      let collabScores = [];
      let cultureScores = [];

      PEER_QUESTIONS.forEach(([qNo, qCat, qDesc, qKey]) => {
        const scores = peerRecords.map(r_e => r_e.peer_eval ? r_e.peer_eval[qKey] : null).filter(v => v !== null && v !== undefined);
        const avg = scores.length ? (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(1) : "—";
        const best = scores.length ? Math.max(...scores) : "—";
        const worst = scores.length ? Math.min(...scores) : "—";
        const detailStr = scores.length ? scores.join(", ") : "—";

        const isCollab = parseInt(qNo.replace("Q","")) <= 29;
        const catBg = isCollab ? COLOR_COLLAB_BG : (qNo === "Q36" ? COLOR_WHITE : COLOR_CULTURE_BG);

        if (isCollab && scores.length) collabScores.push(...scores);
        else if (!isCollab && qNo !== "Q36" && scores.length) cultureScores.push(...scores);

        ws.getCell(r, 1).value = qNo;
        ws.getCell(r, 2).value = qCat;
        ws.getCell(r, 3).value = qDesc;
        ws.getCell(r, 4).value = avg;
        ws.getCell(r, 5).value = best;
        ws.getCell(r, 6).value = worst;
        ws.getCell(r, 7).value = detailStr;

        styleRange(ws, r, 1, r, 7, fontBody, COLOR_WHITE, { horizontal: 'left', vertical: 'top', wrapText: true });
        ws.getCell(r, 1).alignment = { horizontal: 'center', vertical: 'middle' };
        ws.getCell(r, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: catBg } };
        ws.getCell(r, 2).alignment = { horizontal: 'center', vertical: 'middle' };
        ws.getCell(r, 2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: catBg } };
        ws.getCell(r, 4).alignment = { horizontal: 'center', vertical: 'middle' };
        ws.getCell(r, 4).font = fontBodyBold;
        ws.getCell(r, 4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR_AVG_BG } };
        ws.getCell(r, 5).alignment = { horizontal: 'center', vertical: 'middle' };
        ws.getCell(r, 6).alignment = { horizontal: 'center', vertical: 'middle' };
        ws.getCell(r, 7).alignment = { horizontal: 'center', vertical: 'middle' };
        ws.getRow(r).height = 22;
        r++;
      });

      const collabAvg = collabScores.length ? (collabScores.reduce((a, b) => a + b, 0) / collabScores.length).toFixed(2) : "—";
      const cultureAvg = cultureScores.length ? (cultureScores.reduce((a, b) => a + b, 0) / cultureScores.length).toFixed(2) : "—";

      function getLvlLabel(val) {
        if (val === "—") return "";
        const v = parseFloat(val);
        if (v >= 9.0) return "Amazing!";
        if (v >= 7.0) return "Good";
        if (v >= 5.0) return "Keep";
        if (v >= 3.0) return "Grow";
        return "Start";
      }

      ws.mergeCells(r, 1, r, 3);
      ws.getCell(r, 1).value = "協作狀況（總平均）";
      ws.getCell(r, 4).value = collabAvg;
      ws.getCell(r, 5).value = getLvlLabel(collabAvg);
      styleRange(ws, r, 1, r, 7, fontBodyBold, COLOR_WHITE, { horizontal: 'center', vertical: 'middle' });
      ws.getCell(r, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR_COLLAB_BG } };
      ws.getRow(r).height = 24;
      r++;

      ws.mergeCells(r, 1, r, 3);
      ws.getCell(r, 1).value = "文化實踐（總平均）";
      ws.getCell(r, 4).value = cultureAvg;
      ws.getCell(r, 5).value = getLvlLabel(cultureAvg);
      styleRange(ws, r, 1, r, 7, fontBodyBold, COLOR_WHITE, { horizontal: 'center', vertical: 'middle' });
      ws.getCell(r, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR_CULTURE_BG } };
      ws.getRow(r).height = 24;
      r++;

      // 二、部屬自評細節 (文字全部公布)
      ws.mergeCells(r, 1, r, 7);
      ws.getCell(r, 1).value = "二、部屬自評細節：部屬填答內容與自評完整實例";
      styleRange(ws, r, 1, r, 7, fontSubHeader, COLOR_SECTION_BG, { horizontal: 'left', vertical: 'middle', indent: 1 });
      ws.getRow(r).height = 24;
      r++;

      const selfDetails = [
        ["特質盤點", "最穩定 Top 3", (se?.top3_stable || []).join(", ") || "（未填）"],
        ["特質盤點", "練習中 3 項", (se?.top3_practice || []).join(", ") || "（未填）"],
        ["組織文化", "【信任】獨立行動與決策、主動協作、雙向溝通", se?.values?.['信任'] || "（未填）"],
        ["組織文化", "【多元】尊重差異、多元工作方法、主動表達不同觀點", se?.values?.['多元'] || "（未填）"],
        ["組織文化", "【實驗】透過開放心態嘗試修正與反思，勇於檢討及給予回饋", se?.values?.['實驗'] || "（未填）"],
        ["組織文化", "【可持續】內在韌性、自我照顧、彈性的人際與工作邊界", se?.values?.['可持續'] || "（未填）"],
      ];

      const roleKey = MEMBER_COMPETENCY_KEY[memberName] || "專案經理_CGL";
      const compTitles = ROLE_COMPETENCIES_MAP[roleKey] || [];
      compTitles.forEach(cT => {
        let cAns = "（部屬未填寫）";
        if (hasSelf && se.competencies) {
          const item = se.competencies.find(x => x.title === cT || cT.includes(x.title) || x.title.includes(cT));
          if (item && item.answer) cAns = item.answer;
        }
        selfDetails.push(["專業職能", cT, cAns]);
      });

      selfDetails.forEach(([sec, sub, val]) => {
        ws.getCell(r, 1).value = sec;
        ws.getCell(r, 2).value = sub;
        ws.getCell(r, 3).value = val;
        ws.mergeCells(r, 3, r, 7);

        styleRange(ws, r, 1, r, 7, fontBody, COLOR_WHITE, { horizontal: 'left', vertical: 'top', wrapText: true });
        ws.getCell(r, 1).font = fontBodyBold;
        ws.getCell(r, 1).alignment = { horizontal: 'center', vertical: 'middle' };
        ws.getRow(r).height = Math.max(30, Math.min(120, Math.floor((val || "").length / 35 * 18) + 18));
        r++;
      });

      r++;
      // 三、質化文字回饋
      ws.mergeCells(r, 1, r, 7);
      ws.getCell(r, 1).value = "三、質化文字回饋：同儕文字評價與意見回饋";
      styleRange(ws, r, 1, r, 7, fontSubHeader, COLOR_SECTION_BG, { horizontal: 'left', vertical: 'middle', indent: 1 });
      ws.getRow(r).height = 24;
      r++;

      const fbSections = [
        ["Q37. 工作與文化提升建議（改善與前進方向）", "q37_improvement_advice"],
        ["Q38. 其他補充評價與觀察", "q38_other_comments"],
        ["Q39. 肯定與感謝的話（好好星光大賞）", "q39_starlight_thanks"],
      ];

      fbSections.forEach(([fbTitle, fbKey]) => {
        ws.mergeCells(r, 1, r, 7);
        ws.getCell(r, 1).value = fbTitle;
        styleRange(ws, r, 1, r, 7, fontTableHdr, COLOR_SECTION_BG, { horizontal: 'left', vertical: 'middle', indent: 1 });
        ws.getRow(r).height = 22;
        r++;

        if (numPeers === 0) {
          ws.mergeCells(r, 1, r, 7);
          ws.getCell(r, 1).value = "（目前尚無同儕填答回饋）";
          styleRange(ws, r, 1, r, 7, fontBody, COLOR_WHITE, { horizontal: 'left', vertical: 'top' });
          ws.getRow(r).height = 22;
          r++;
        } else {
          peerRecords.forEach((rEntry, idx) => {
            const peerLabel = `同儕 ${String.fromCharCode(65 + idx)}`;
            const fbText = rEntry.peer_eval ? (rEntry.peer_eval[fbKey] || "（無填寫）") : "（無填寫）";
            ws.getCell(r, 1).value = peerLabel;
            ws.getCell(r, 2).value = fbText;
            ws.mergeCells(r, 2, r, 7);

            styleRange(ws, r, 1, r, 7, fontBody, COLOR_WHITE, { horizontal: 'left', vertical: 'top', wrapText: true });
            ws.getCell(r, 1).font = fontBodyBold;
            ws.getCell(r, 1).alignment = { horizontal: 'center', vertical: 'middle' };
            ws.getCell(r, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR_SECTION_BG } };
            ws.getRow(r).height = Math.max(26, Math.min(120, Math.floor(fbText.length / 35 * 18) + 18));
            r++;
          });
        }
      });

      r++;
      // 四、評級標準
      addRatingStandardsBlockToSheet(ws, r);
    }

    async function exportCurrentSubordinateExcel() {
      const wb = new ExcelJS.Workbook();
      buildSubordinateSupervisorWorksheet(wb, currentSubReviewMember);
      buildSubordinatePeerWorksheet(wb, currentSubReviewMember);
      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      saveAs(blob, `【${currentSubReviewMember}】_【主管評下屬用】.xlsx`);
      showToast(`已成功下載【${currentSubReviewMember}】_【主管評下屬用】.xlsx！`);
    }

    async function exportCurrentSupervisorTeamExcel() {
      exportSupervisorTeamExcel(currentSubReviewTeam);
    }

    async function exportSupervisorTeamExcel(supKey) {
      const wb = new ExcelJS.Workbook();
      const members = SUPERVISOR_TEAMS[supKey] || [];
      members.forEach(m => {
        buildSubordinateSupervisorWorksheet(wb, m);
        buildSubordinatePeerWorksheet(wb, m);
      });
      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      
      let outName = "VVN品牌主管專用_【主管評下屬用】.xlsx";
      if (supKey === "何維安") outName = "VVN品牌主管專用_【主管評下屬用】.xlsx";
      else if (supKey === "姚品瑄") outName = "姚品瑄專案主管專用_【主管評下屬用】.xlsx";
      else if (supKey === "張希慈") outName = "張希慈執行長主管專用_【主管評下屬用】.xlsx";
      else if (supKey === "張希慈_執行長") outName = "張希慈個人自評_【主管評下屬用】.xlsx";
      else outName = `${supKey}主管專用_【主管評下屬用】.xlsx`;

      saveAs(blob, outName);
      showToast(`已成功下載 ${outName}！`);
    }

    async function exportFullMasterWorkbook() {
      const wb = new ExcelJS.Workbook();
      ALL_MEMBERS.forEach(m => {
        buildSubordinateSupervisorWorksheet(wb, m);
        buildSubordinatePeerWorksheet(wb, m);
      });
      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const outName = "好好星球_360年中成長評估_【主管評下屬用】Master總表.xlsx";
      saveAs(blob, outName);
      showToast(`已成功下載 ${outName}！`);
    }

    // =========================================================================
    // GOOGLE APPS SCRIPT SYNC ENGINE
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
              competencies.push({ title: colName, answer: val });
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
            badge.innerText = `🟢 試算表即時連線（${parsed.length}筆）`;
          }
          if (syncText) syncText.innerText = `已同步（${parsed.length}筆）`;
          showToast(`成功優先同步 Google 試算表（表單回覆1）最新 ${parsed.length} 筆資料！`);
          refreshAllViews();
        } else {
          throw new Error("試算表回傳資料格式為空");
        }
      } catch (err) {
        console.warn("GAS Sync fallback to local data:", err);
        if (badge) {
          badge.className = "px-2.5 py-0.5 text-xs font-semibold rounded-md bg-[#F2EEE6] text-[#6E6662] border border-[#E0D7CA]";
          badge.innerText = `⚪ 本地備份資料（${RAW_DATA.length}筆）`;
        }
        if (syncText) syncText.innerText = "同步 Google 試算表";
        if (isUserInitiated) {
          showToast("無法連線至 Google 試算表，系統已切換使用本地備份資料。", true);
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

    window.addEventListener('DOMContentLoaded', () => {
      loadLocalStorageStates();
      refreshAllViews();
      syncFromGoogleAppsScript(false);
    });
  </script>
</body>
</html>
"""

output_file = os.path.join(os.path.dirname(__file__), "..", "index.html")
with open(output_file, "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"Generated {output_file} successfully!")
