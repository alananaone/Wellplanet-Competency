import json

with open("evaluation_data.json", "r", encoding="utf-8") as f:
    data_json_str = f.read()

html_template = """<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>好好星球文化基金會 360 年中成長評估視覺化系統</title>
  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- Chart.js CDN -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <!-- PapaParse for robust client-side CSV parsing -->
  <script src="https://cdn.jsdelivr.net/npm/papaparse@5.4.1/papaparse.min.js"></script>
  <!-- SheetJS for client-side Excel XLSX export -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <!-- Lucide Icons -->
  <script src="https://unpkg.com/lucide@latest"></script>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&family=Noto+Sans+TC:wght@400;500;700&display=swap');
    body {
      font-family: 'Noto Sans TC', 'Plus Jakarta Sans', sans-serif;
      background-color: #F8FAFC;
    }
    .custom-scrollbar::-webkit-scrollbar {
      width: 6px;
      height: 6px;
    }
    .custom-scrollbar::-webkit-scrollbar-thumb {
      background-color: #CBD5E1;
      border-radius: 4px;
    }
    .badge-stable {
      background: #DCFCE7;
      color: #15803D;
      border: 1px solid #BBF7D0;
    }
    .badge-practice {
      background: #FFEDD5;
      color: #C2410C;
      border: 1px solid #FED7AA;
    }
    .drop-zone-active {
      border-color: #4F46E5 !important;
      background-color: #EEF2FF !important;
    }
  </style>
</head>
<body class="text-slate-800 min-h-screen flex flex-col">

  <!-- TOP HEADER -->
  <header class="bg-white border-b border-slate-200 sticky top-0 z-50 shadow-xs">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
      <div class="flex items-center justify-between h-16 gap-4">
        <div class="flex items-center gap-3">
          <div class="w-10 h-10 rounded-xl bg-gradient-to-br from-indigo-500 to-indigo-700 flex items-center justify-center text-white font-bold shadow-md shadow-indigo-100 shrink-0">
            360
          </div>
          <div>
            <h1 class="text-base sm:text-lg font-bold text-slate-900 leading-tight">好好星球文化基金會 360 年中成長評估</h1>
            <p class="text-xs text-slate-500">視覺化篩選與主管部屬專屬報告導出系統</p>
          </div>
        </div>

        <!-- ACTION BUTTONS (UPLOAD CSV + EXPORT EXCEL) -->
        <div class="flex items-center gap-2.5">
          <!-- UPLOAD CSV BUTTON -->
          <label class="cursor-pointer inline-flex items-center gap-1.5 px-3 py-2 text-xs sm:text-sm font-semibold rounded-xl bg-slate-100 text-slate-700 hover:bg-slate-200 hover:text-slate-900 transition shadow-2xs">
            <i data-lucide="upload" class="w-4 h-4 text-indigo-600"></i>
            <span class="hidden sm:inline">上傳最新 CSV</span>
            <span class="sm:hidden">上傳</span>
            <input type="file" id="csvFileInput" accept=".csv" class="hidden" onchange="handleFileUpload(event)">
          </label>

          <!-- EXCEL EXPORT DROPDOWN -->
          <div class="relative inline-block text-left" id="exportDropdown">
            <button onclick="toggleDropdown()" class="inline-flex items-center gap-1.5 sm:gap-2 px-3.5 py-2 text-xs sm:text-sm font-medium rounded-xl bg-indigo-600 text-white hover:bg-indigo-700 transition shadow-sm">
              <i data-lucide="download" class="w-4 h-4"></i>
              <span class="hidden sm:inline">下載主管專用 Excel</span>
              <span class="sm:hidden">下載 Excel</span>
              <i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
            </button>
            <div id="dropdownMenu" class="hidden absolute right-0 mt-2 w-80 origin-top-right rounded-2xl bg-white p-2 shadow-2xl ring-1 ring-black ring-opacity-5 z-50 divide-y divide-slate-100">
              <div class="py-1">
                <div class="px-3 py-1.5 text-[11px] font-bold text-slate-400 uppercase tracking-wider">主管部屬自評專用包 (已排版含題目)</div>
                <button onclick="exportSupervisorExcelClientSide('張希慈')" class="w-full text-left flex items-center gap-2.5 px-3 py-2.5 text-xs sm:text-sm text-slate-700 hover:bg-indigo-50 hover:text-indigo-700 rounded-xl transition">
                  <div class="p-1.5 bg-indigo-50 text-indigo-600 rounded-lg"><i data-lucide="file-spreadsheet" class="w-4 h-4"></i></div>
                  <div>
                    <div class="font-bold">【張希慈】部屬自評彙整表</div>
                    <div class="text-[11px] text-slate-400">何維安、陳泳璇、張芳媐、姚品瑄、胡喻翔</div>
                  </div>
                </button>
                <button onclick="exportSupervisorExcelClientSide('何維安')" class="w-full text-left flex items-center gap-2.5 px-3 py-2.5 text-xs sm:text-sm text-slate-700 hover:bg-indigo-50 hover:text-indigo-700 rounded-xl transition">
                  <div class="p-1.5 bg-indigo-50 text-indigo-600 rounded-lg"><i data-lucide="file-spreadsheet" class="w-4 h-4"></i></div>
                  <div>
                    <div class="font-bold">【何維安】部屬自評彙整表</div>
                    <div class="text-[11px] text-slate-400">林文琇</div>
                  </div>
                </button>
                <button onclick="exportSupervisorExcelClientSide('姚品瑄')" class="w-full text-left flex items-center gap-2.5 px-3 py-2.5 text-xs sm:text-sm text-slate-700 hover:bg-indigo-50 hover:text-indigo-700 rounded-xl transition">
                  <div class="p-1.5 bg-indigo-50 text-indigo-600 rounded-lg"><i data-lucide="file-spreadsheet" class="w-4 h-4"></i></div>
                  <div>
                    <div class="font-bold">【姚品瑄】部屬自評彙整表</div>
                    <div class="text-[11px] text-slate-400">薛筑瑄、戴佑珍</div>
                  </div>
                </button>
                <button onclick="exportSupervisorExcelClientSide('張希慈_執行長')" class="w-full text-left flex items-center gap-2.5 px-3 py-2.5 text-xs sm:text-sm text-slate-700 hover:bg-indigo-50 hover:text-indigo-700 rounded-xl transition">
                  <div class="p-1.5 bg-indigo-50 text-indigo-600 rounded-lg"><i data-lucide="file-spreadsheet" class="w-4 h-4"></i></div>
                  <div>
                    <div class="font-bold">【張希慈】執行長個人自評表</div>
                    <div class="text-[11px] text-slate-400">執行長職位專用</div>
                  </div>
                </button>
              </div>
              <div class="py-1">
                <div class="px-3 py-1.5 text-[11px] font-bold text-slate-400 uppercase tracking-wider">全組織總表</div>
                <button onclick="exportFullWorkbookClientSide()" class="w-full text-left flex items-center gap-2.5 px-3 py-2.5 text-xs sm:text-sm font-bold text-indigo-600 hover:bg-indigo-50 rounded-xl transition">
                  <div class="p-1.5 bg-indigo-100 text-indigo-700 rounded-lg"><i data-lucide="layers" class="w-4 h-4"></i></div>
                  <div>
                    <div>下載完整 Excel (包含所有 Sheets)</div>
                    <div class="text-[11px] text-indigo-400 font-normal">包含評主管、評同事、各主管包</div>
                  </div>
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>

      <!-- MAIN TABS -->
      <nav class="flex space-x-1 border-t border-slate-100 pt-2 -mb-px overflow-x-auto custom-scrollbar">
        <button onclick="switchTab('self')" id="tab-btn-self" class="tab-btn inline-flex items-center gap-2 px-4 py-2.5 text-sm font-semibold border-b-2 border-indigo-600 text-indigo-600 shrink-0">
          <i data-lucide="user-check" class="w-4 h-4"></i>
          自評（依主管分流）
          <span class="px-1.5 py-0.5 text-xs rounded-full bg-indigo-100 text-indigo-700 font-bold" id="badge-self-count">5</span>
        </button>
        <button onclick="switchTab('supervisor')" id="tab-btn-supervisor" class="tab-btn inline-flex items-center gap-2 px-4 py-2.5 text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-slate-700 shrink-0">
          <i data-lucide="award" class="w-4 h-4"></i>
          評主管
          <span class="px-1.5 py-0.5 text-xs rounded-full bg-slate-100 text-slate-600 font-bold" id="badge-sup-count">4</span>
        </button>
        <button onclick="switchTab('peer')" id="tab-btn-peer" class="tab-btn inline-flex items-center gap-2 px-4 py-2.5 text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-slate-700 shrink-0">
          <i data-lucide="users" class="w-4 h-4"></i>
          評同事
          <span class="px-1.5 py-0.5 text-xs rounded-full bg-slate-100 text-slate-600 font-bold" id="badge-peer-count">23</span>
        </button>
        <button onclick="switchTab('overview360')" id="tab-btn-overview360" class="tab-btn inline-flex items-center gap-2 px-4 py-2.5 text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-slate-700 shrink-0">
          <i data-lucide="pie-chart" class="w-4 h-4"></i>
          個人 360 綜合雷達
        </button>
      </nav>
    </div>
  </header>

  <!-- DRAG & DROP UPLOAD NOTIFICATION ZONE (HIDDEN UNTIL DRAG OVER) -->
  <div id="dropZone" ondragover="handleDragOver(event)" ondragleave="handleDragLeave(event)" ondrop="handleDrop(event)" class="transition-all duration-200 border-2 border-dashed border-slate-300 bg-white m-4 rounded-2xl p-4 text-center text-xs text-slate-500 hidden sm:flex items-center justify-center gap-2 hover:border-indigo-400">
    <i data-lucide="file-up" class="w-4 h-4 text-indigo-600"></i>
    <span>💡 支援將最新 Google 表單匯出的 CSV 檔案直接<b>拖曳至此處</b>，即可即時載入並更新全系統數據與 Excel 報表！</span>
  </div>

  <!-- MAIN CONTENT AREA -->
  <main class="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-4 sm:py-6">

    <!-- ======================================================== -->
    <!-- TAB 1: 自評 (依主管分流) -->
    <!-- ======================================================== -->
    <section id="tab-section-self" class="tab-content space-y-6">
      
      <!-- SUB-FILTER FOR SUPERVISOR TEAMS -->
      <div class="bg-white rounded-2xl p-4 border border-slate-200 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div class="flex items-center gap-2 flex-wrap" id="self-supervisor-pills">
          <span class="text-xs font-bold text-slate-400 mr-1 flex items-center gap-1 uppercase tracking-wider">
            <i data-lucide="filter" class="w-3.5 h-3.5"></i> 主管分流：
          </span>
          <button onclick="filterSelfSupervisor('張希慈')" id="self-sup-btn-張希慈" class="self-sup-btn px-3.5 py-2 rounded-xl text-xs sm:text-sm font-semibold bg-indigo-600 text-white shadow-sm transition">
            👑 張希慈 的部屬 (5人)
          </button>
          <button onclick="filterSelfSupervisor('何維安')" id="self-sup-btn-何維安" class="self-sup-btn px-3.5 py-2 rounded-xl text-xs sm:text-sm font-medium bg-slate-100 text-slate-600 hover:bg-slate-200 transition">
            🎨 何維安 的部屬 (1人)
          </button>
          <button onclick="filterSelfSupervisor('姚品瑄')" id="self-sup-btn-姚品瑄" class="self-sup-btn px-3.5 py-2 rounded-xl text-xs sm:text-sm font-medium bg-slate-100 text-slate-600 hover:bg-slate-200 transition">
            🚀 姚品瑄 的部屬 (2人)
          </button>
          <button onclick="filterSelfSupervisor('張希慈_執行長')" id="self-sup-btn-張希慈_執行長" class="self-sup-btn px-3.5 py-2 rounded-xl text-xs sm:text-sm font-medium bg-slate-100 text-slate-600 hover:bg-slate-200 transition">
            ⭐ 執行長個人自評
          </button>
          <button onclick="filterSelfSupervisor('ALL')" id="self-sup-btn-ALL" class="self-sup-btn px-3.5 py-2 rounded-xl text-xs sm:text-sm font-medium bg-slate-100 text-slate-600 hover:bg-slate-200 transition">
            📋 全部已自評 (5人)
          </button>
        </div>

        <div id="self-export-btn-container">
          <!-- Dynamic Export Button for current supervisor -->
        </div>
      </div>

      <!-- TEAM INFO BANNER -->
      <div id="self-team-banner" class="bg-indigo-50/70 border border-indigo-100 rounded-2xl p-4 flex items-center justify-between">
        <div class="flex items-center gap-3">
          <div class="p-2.5 bg-indigo-600 text-white rounded-xl shadow-sm">
            <i data-lucide="users" class="w-5 h-5"></i>
          </div>
          <div>
            <h2 id="self-team-title" class="text-sm font-bold text-indigo-950">張希慈 的部屬自評列表</h2>
            <p id="self-team-subtitle" class="text-xs text-indigo-700">涵蓋部屬：何維安、陳泳璇、張芳媐、姚品瑄、胡喻翔</p>
          </div>
        </div>
        <div class="text-xs px-3 py-1 bg-white rounded-full text-indigo-700 font-bold border border-indigo-100 shadow-sm" id="self-completion-status">
          已填答 2 / 應填 5 人
        </div>
      </div>

      <!-- MEMBER SELF EVALUATION CARDS CONTAINER -->
      <div id="self-eval-cards-container" class="space-y-6">
        <!-- Injected via JavaScript -->
      </div>
    </section>

    <!-- ======================================================== -->
    <!-- TAB 2: 評主管 -->
    <!-- ======================================================== -->
    <section id="tab-section-supervisor" class="tab-content hidden space-y-6">
      <div class="bg-white rounded-2xl p-4 border border-slate-200 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div class="flex items-center gap-2 flex-wrap">
          <span class="text-xs font-bold text-slate-400 mr-1 flex items-center gap-1 uppercase tracking-wider">
            <i data-lucide="user" class="w-3.5 h-3.5"></i> 受評主管：
          </span>
          <button onclick="filterSupervisor('ALL')" id="sup-filter-btn-ALL" class="sup-filter-btn px-3.5 py-2 rounded-xl text-xs sm:text-sm font-semibold bg-indigo-600 text-white shadow-sm transition">
            全部主管 (4筆)
          </button>
          <button onclick="filterSupervisor('張希慈')" id="sup-filter-btn-張希慈" class="sup-filter-btn px-3.5 py-2 rounded-xl text-xs sm:text-sm font-medium bg-slate-100 text-slate-600 hover:bg-slate-200 transition">
            張希慈 (3筆)
          </button>
          <button onclick="filterSupervisor('何維安')" id="sup-filter-btn-何維安" class="sup-filter-btn px-3.5 py-2 rounded-xl text-xs sm:text-sm font-medium bg-slate-100 text-slate-600 hover:bg-slate-200 transition">
            何維安 (1筆)
          </button>
        </div>

        <button onclick="exportFullWorkbookClientSide()" class="inline-flex items-center gap-2 px-3.5 py-2 text-xs sm:text-sm font-medium rounded-xl bg-slate-100 text-slate-700 hover:bg-slate-200 transition">
          <i data-lucide="download" class="w-4 h-4"></i> 下載評主管總表 (XLSX)
        </button>
      </div>

      <!-- CHARTS & SUMMARY CARDS -->
      <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <!-- Radar Chart for Culture -->
        <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
          <h3 class="text-sm font-bold text-slate-800 mb-2 flex items-center gap-2">
            <i data-lucide="compass" class="w-4 h-4 text-indigo-600"></i> 四大文化實踐維度
          </h3>
          <div class="h-64 flex items-center justify-center">
            <canvas id="supervisorRadarChart"></canvas>
          </div>
        </div>

        <!-- Bar Chart for Management Items -->
        <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm lg:col-span-2">
          <h3 class="text-sm font-bold text-slate-800 mb-2 flex items-center gap-2">
            <i data-lucide="bar-chart-3" class="w-4 h-4 text-indigo-600"></i> 管理能力各題平均得分 (滿分10分)
          </h3>
          <div class="h-64">
            <canvas id="supervisorBarChart"></canvas>
          </div>
        </div>
      </div>

      <!-- DETAILED FEEDBACK LIST -->
      <div id="supervisor-feedback-list" class="space-y-4">
        <!-- Injected via JavaScript -->
      </div>
    </section>

    <!-- ======================================================== -->
    <!-- TAB 3: 評同事 -->
    <!-- ======================================================== -->
    <section id="tab-section-peer" class="tab-content hidden space-y-6">
      <div class="bg-white rounded-2xl p-4 border border-slate-200 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div class="flex items-center gap-2 flex-wrap" id="peer-pills-container">
          <span class="text-xs font-bold text-slate-400 mr-1 flex items-center gap-1 uppercase tracking-wider">
            <i data-lucide="user" class="w-3.5 h-3.5"></i> 受評同事：
          </span>
          <!-- Dynamic Peer Buttons Injected -->
        </div>

        <button onclick="exportFullWorkbookClientSide()" class="inline-flex items-center gap-2 px-3.5 py-2 text-xs sm:text-sm font-medium rounded-xl bg-slate-100 text-slate-700 hover:bg-slate-200 transition">
          <i data-lucide="download" class="w-4 h-4"></i> 下載評同事總表 (XLSX)
        </button>
      </div>

      <!-- CHARTS & SUMMARY CARDS -->
      <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <!-- Radar Chart for Culture -->
        <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
          <h3 class="text-sm font-bold text-slate-800 mb-2 flex items-center gap-2">
            <i data-lucide="compass" class="w-4 h-4 text-emerald-600"></i> 文化實踐維度表現
          </h3>
          <div class="h-64 flex items-center justify-center">
            <canvas id="peerRadarChart"></canvas>
          </div>
        </div>

        <!-- Bar Chart for Peer Items -->
        <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm lg:col-span-2">
          <h3 class="text-sm font-bold text-slate-800 mb-2 flex items-center gap-2">
            <i data-lucide="bar-chart-3" class="w-4 h-4 text-emerald-600"></i> 協作、當責與溝通評分 (滿分10分)
          </h3>
          <div class="h-64">
            <canvas id="peerBarChart"></canvas>
          </div>
        </div>
      </div>

      <!-- DETAILED PEER FEEDBACK LIST -->
      <div id="peer-feedback-list" class="space-y-4">
        <!-- Injected via JavaScript -->
      </div>
    </section>

    <!-- ======================================================== -->
    <!-- TAB 4: 個人 360 綜合雷達與對照報告 -->
    <!-- ======================================================== -->
    <section id="tab-section-overview360" class="tab-content hidden space-y-6">
      <div class="bg-white rounded-2xl p-4 border border-slate-200 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div class="flex items-center gap-2 flex-wrap" id="overview-member-pills">
          <span class="text-xs font-bold text-slate-400 mr-1 flex items-center gap-1 uppercase tracking-wider">
            <i data-lucide="sparkles" class="w-3.5 h-3.5"></i> 成員報告：
          </span>
          <!-- Dynamic Member Buttons -->
        </div>
      </div>

      <div id="overview-report-container" class="space-y-6">
        <!-- Dynamic Personal 360 Profile -->
      </div>
    </section>

  </main>

  <!-- TOAST NOTIFICATION -->
  <div id="toast" class="fixed bottom-6 right-6 z-50 transform transition-all duration-300 opacity-0 translate-y-4 pointer-events-none bg-slate-900 text-white px-4 py-3 rounded-2xl shadow-xl flex items-center gap-3 text-sm">
    <div id="toast-icon" class="p-1 rounded-lg bg-emerald-500 text-white">
      <i data-lucide="check" class="w-4 h-4"></i>
    </div>
    <span id="toast-msg">操作成功</span>
  </div>

  <footer class="bg-white border-t border-slate-200 py-6 text-center text-xs text-slate-500 mt-auto">
    好好星球文化基金會 360 年中成長評估視覺化系統 · 支援動態 CSV 載入與客製化 XLSX 匯出
  </footer>

  <script>
    // Embedded Evaluation Dataset
    let RAW_DATA = """ + data_json_str + """;

    // Supervisor Teams Mapping (Updated with 胡喻翔 under 張希慈)
    const SUPERVISOR_TEAMS = {
      "張希慈": ["何維安", "陳泳璇", "張芳媐", "姚品瑄", "胡喻翔"],
      "何維安": ["林文琇"],
      "姚品瑄": ["薛筑瑄", "戴佑珍"],
      "張希慈_執行長": ["張希慈"],
      "ALL": ["何維安", "陳泳璇", "張芳媐", "姚品瑄", "胡喻翔", "林文琇", "薛筑瑄", "戴佑珍", "張希慈"]
    };

    // Role definitions for self-eval mapping in JS
    const JOB_BLOCKS_JS = {
      "執行長": {
        "competencies": [
          [59, "策略決策與組織方向設定"],
          [60, "組織治理與財務風險管理"],
          [61, "關鍵利害關係人關係建立與維繫"],
          [62, "公關、演講與媒體關係"],
          [63, "核心團隊培養與組織文化建立"],
          [64, "其他創造的好事與預防的壞事"]
        ],
        "reflection": [
          [65, "未來一年想創造的價值"],
          [66, "最常感到卡關/掙扎的階段"],
          [67, "卡關具體原因描述"],
          [68, "希望組織當時提供的幫助"],
          [69, "對願景使命理解與實踐的最大轉變"]
        ]
      },
      "行政經理": {
        "competencies": [
          [71, "人事薪資與人資系統管理"],
          [72, "法規與政府公文管理"],
          [73, "董事會與治理作業執行"],
          [74, "財務核銷與內控執行"],
          [75, "總務採購與行政庶務管理"],
          [76, "其他創造的好事與預防的壞事"]
        ],
        "reflection": [
          [77, "未來一年想創造的價值"],
          [78, "最常感到卡關/掙扎的階段"],
          [79, "卡關具體原因描述"],
          [80, "希望組織當時提供的幫助"],
          [81, "對願景使命理解與實踐的最大轉變"]
        ]
      },
      "營運經理": {
        "competencies": [
          [83, "組織營運流程設計與優化"],
          [84, "制度文件與 SOP 建置"],
          [85, "人力資源策略與選用育留"],
          [86, "組織文化落實與制度轉化"],
          [87, "總務採購與行政庶務管理"],
          [88, "其他創造的好事與預防的壞事"]
        ],
        "reflection": [
          [89, "未來一年想創造的價值"],
          [90, "最常感到卡關/掙扎的階段"],
          [91, "卡關具體原因描述"],
          [92, "希望組織當時提供的幫助"],
          [93, "對願景使命理解與實踐的最大轉變"]
        ]
      },
      "專案部門儲備主管": {
        "competencies": [
          [95, "部門策略規劃與專案組合管理"],
          [96, "專案經理管理與培育"],
          [97, "部門預算與資源配置管理"],
          [98, "跨 LAB 專業掌握（CGL、SL、SPL）"],
          [99, "利害關係人管理與衝突協調"],
          [100, "其他創造的好事與預防的壞事"]
        ],
        "reflection": [
          [101, "未來一年想創造的價值"],
          [102, "最常感到卡關/掙扎的階段"],
          [103, "卡關具體原因描述"],
          [104, "希望組織當時提供的幫助"],
          [105, "對願景使命理解與實踐的最大轉變"]
        ]
      },
      "CGL專案經理": {
        "competencies": [
          [107, "專案企劃與現場執行"],
          [108, "專案時程與預算規劃管理"],
          [109, "需求研究與方案迭代"],
          [110, "外部夥伴關係經營"],
          [111, "多元教學設計與現場引導"],
          [112, "其他創造的好事與預防的壞事"]
        ],
        "reflection": [
          [113, "未來一年想創造的價值"],
          [114, "最常感到卡關/掙扎的階段"],
          [115, "卡關具體原因描述"],
          [116, "希望組織當時提供的幫助"],
          [117, "對願景使命理解與實踐的最大轉變"]
        ]
      },
      "SL專案經理": {
        "competencies": [
          [119, "專案企劃與現場執行"],
          [120, "專案時程與預算規劃管理"],
          [121, "需求研究與方案迭代"],
          [122, "外部夥伴關係經營"],
          [123, "社群經營與培力需求回應"],
          [124, "其他創造的好事與預防的壞事"]
        ],
        "reflection": [
          [125, "未來一年想創造的價值"],
          [126, "最常感到卡關/掙扎的階段"],
          [127, "卡關具體原因描述"],
          [128, "希望組織當時提供的幫助"],
          [129, "對願景使命理解與實踐的最大轉變"]
        ]
      },
      "品牌經理": {
        "competencies": [
          [131, "品牌定位與外部溝通一致性"],
          [132, "行銷策略與議題倡議"],
          [133, "品牌活動策劃與策展敘事"],
          [134, "內部品牌管理與雇主品牌"],
          [135, "品牌危機與聲譽風險處理"],
          [136, "其他創造的好事與預防的壞事"]
        ],
        "reflection": [
          [137, "未來一年想創造的價值"],
          [138, "最常感到卡關/掙扎的階段"],
          [139, "卡關具體原因描述"],
          [140, "希望組織當時提供的幫助"],
          [141, "對願景使命理解與實踐的最大轉變"]
        ]
      },
      "視覺設計師": {
        "competencies": [
          [143, "品牌識別系統設計與維護"],
          [144, "視覺設計實務"],
          [145, "需求釐清與創意提案"],
          [146, "其他創造的好事與預防的壞事"]
        ],
        "reflection": [
          [147, "未來一年想創造的價值"],
          [148, "最常感到卡關/掙扎的階段"],
          [149, "卡關具體原因描述"],
          [150, "希望組織當時提供的幫助"],
          [151, "對願景使命理解與實踐的最大轉變"]
        ]
      }
    };

    let currentSelfSupervisor = '張希慈';
    let currentSupervisorFilter = 'ALL';
    let currentPeerFilter = 'ALL';
    let currentOverviewMember = '何維安';

    let supervisorRadar = null;
    let supervisorBar = null;
    let peerRadar = null;
    let peerBar = null;
    let overviewRadar = null;

    function showToast(msg) {
      const toast = document.getElementById('toast');
      document.getElementById('toast-msg').innerText = msg;
      toast.classList.remove('opacity-0', 'translate-y-4', 'pointer-events-none');
      setTimeout(() => {
        toast.classList.add('opacity-0', 'translate-y-4', 'pointer-events-none');
      }, 3500);
    }

    function toggleDropdown() {
      const menu = document.getElementById('dropdownMenu');
      menu.classList.toggle('hidden');
    }

    window.onclick = function(e) {
      if (!e.target.closest('#exportDropdown')) {
        const menu = document.getElementById('dropdownMenu');
        if (menu && !menu.classList.contains('hidden')) {
          menu.classList.add('hidden');
        }
      }
    };

    function switchTab(tabId) {
      document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('border-indigo-600', 'text-indigo-600');
        btn.classList.add('border-transparent', 'text-slate-500');
      });
      document.getElementById('tab-btn-' + tabId).classList.add('border-indigo-600', 'text-indigo-600');
      document.getElementById('tab-btn-' + tabId).classList.remove('border-transparent', 'text-slate-500');

      document.querySelectorAll('.tab-content').forEach(sec => sec.classList.add('hidden'));
      document.getElementById('tab-section-' + tabId).classList.remove('hidden');

      if (tabId === 'supervisor') {
        renderSupervisorSection();
      } else if (tabId === 'peer') {
        renderPeerSection();
      } else if (tabId === 'overview360') {
        renderOverviewSection();
      }
    }

    // ==========================================
    // CLIENT-SIDE CSV PARSING & FILE UPLOAD
    // ==========================================
    function handleDragOver(e) {
      e.preventDefault();
      document.getElementById('dropZone').classList.add('drop-zone-active');
    }
    function handleDragLeave(e) {
      e.preventDefault();
      document.getElementById('dropZone').classList.remove('drop-zone-active');
    }
    function handleDrop(e) {
      e.preventDefault();
      document.getElementById('dropZone').classList.remove('drop-zone-active');
      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        parseCSVFile(e.dataTransfer.files[0]);
      }
    }

    function handleFileUpload(e) {
      if (e.target.files && e.target.files.length > 0) {
        parseCSVFile(e.target.files[0]);
      }
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
          const rawRows = results.data.slice(1);
          const header = results.data[0];

          const newEntries = [];
          rawRows.forEach(row => {
            if (row.length < 4) return;
            const timestamp = (row[0] || "").trim();
            const email = (row[1] || "").trim();
            const target = (row[2] || "").trim();
            const relation = (row[3] || "").trim();
            const job_role = (row[58] || "").trim();

            const entry = {
              timestamp, email, target, relation, job_role
            };

            if (relation === "主管") {
              const scores = {};
              for (let c = 4; c <= 19; c++) {
                const val = (row[c] || "").trim();
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
                q20_vision_mission: (row[20] || "").trim(),
                q21_improvement_advice: (row[21] || "").trim(),
                q22_other_comments: (row[22] || "").trim(),
                q23_starlight_thanks: (row[23] || "").trim()
              };
            } else if (relation === "同事") {
              const scores = {};
              for (let c = 24; c <= 36; c++) {
                const val = (row[c] || "").trim();
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
                q37_improvement_advice: (row[37] || "").trim(),
                q38_other_comments: (row[38] || "").trim(),
                q39_starlight_thanks: (row[39] || "").trim()
              };
            } else if (relation === "自評") {
              const top3_stable = (row[52] || "").split(",").map(s => s.trim()).filter(Boolean);
              const top3_practice = (row[53] || "").split(",").map(s => s.trim()).filter(Boolean);
              const trust_text = (row[54] || "").trim();
              const diversity_text = (row[55] || "").trim();
              const experiment_text = (row[56] || "").trim();
              const sustainability_text = (row[57] || "").trim();

              const competencies = [];
              const reflection = {};

              const job_def = JOB_BLOCKS_JS[job_role];
              if (job_def) {
                job_def.competencies.forEach(([idx, title]) => {
                  competencies.append ? null : competencies.push({ title, answer: (row[idx] || "").trim() });
                });
                job_def.reflection.forEach(([idx, title]) => {
                  reflection[title] = (row[idx] || "").trim();
                });
              }

              entry.self_eval = {
                job_role,
                top3_stable,
                top3_practice,
                values: {
                  "信任": trust_text,
                  "多元": diversity_text,
                  "實驗": experiment_text,
                  "可持續": sustainability_text
                },
                competencies,
                reflection
              };
            }

            newEntries.push(entry);
          });

          RAW_DATA = newEntries;
          updateAllCounts();
          renderSelfSection();
          initPeerPills();
          initOverviewPills();
          showToast(`🎉 成功載入最新 CSV！共更新 ${newEntries.length} 筆填答紀錄。`);
        }
      });
    }

    function updateAllCounts() {
      const selfCount = RAW_DATA.filter(e => e.relation === "自評").length;
      const supCount = RAW_DATA.filter(e => e.relation === "主管").length;
      const peerCount = RAW_DATA.filter(e => e.relation === "同事").length;

      document.getElementById('badge-self-count').innerText = selfCount;
      document.getElementById('badge-sup-count').innerText = supCount;
      document.getElementById('badge-peer-count').innerText = peerCount;
    }

    // ==========================================
    // TAB 1: SELF EVALUATION LOGIC
    // ==========================================
    function filterSelfSupervisor(supKey) {
      currentSelfSupervisor = supKey;
      document.querySelectorAll('.self-sup-btn').forEach(btn => {
        btn.classList.remove('bg-indigo-600', 'text-white');
        btn.classList.add('bg-slate-100', 'text-slate-600');
      });
      const activeBtn = document.getElementById('self-sup-btn-' + supKey);
      if (activeBtn) {
        activeBtn.classList.add('bg-indigo-600', 'text-white');
        activeBtn.classList.remove('bg-slate-100', 'text-slate-600');
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

      if (currentSelfSupervisor === '張希慈') {
        bannerTitle.innerText = "【張希慈】部屬自評彙整";
        bannerSub.innerText = "涵蓋部屬：何維安、陳泳璇、張芳媐、姚品瑄、胡喻翔";
      } else if (currentSelfSupervisor === '何維安') {
        bannerTitle.innerText = "【何維安】部屬自評彙整";
        bannerSub.innerText = "涵蓋部屬：林文琇";
      } else if (currentSelfSupervisor === '姚品瑄') {
        bannerTitle.innerText = "【姚品瑄】部屬自評彙整";
        bannerSub.innerText = "涵蓋部屬：薛筑瑄、戴佑珍";
      } else if (currentSelfSupervisor === '張希慈_執行長') {
        bannerTitle.innerText = "【張希慈】執行長個人自評";
        bannerSub.innerText = "評估職位：執行長";
      } else {
        bannerTitle.innerText = "全組織自評總覽";
        bannerSub.innerText = "查看基金會所有已填寫之自評紀錄";
      }

      exportContainer.innerHTML = `
        <button onclick="exportSupervisorExcelClientSide('${currentSelfSupervisor}')" class="inline-flex items-center gap-2 px-3.5 py-2 rounded-xl text-xs sm:text-sm font-semibold bg-emerald-600 text-white hover:bg-emerald-700 transition shadow-sm">
          <i data-lucide="file-spreadsheet" class="w-4 h-4"></i>
          下載本組主管專用 XLSX 表格
        </button>
      `;

      let completedCount = 0;
      let html = "";

      memberNames.forEach(memName => {
        const entry = RAW_DATA.find(e => e.target === memName && e.relation === "自評");
        if (entry && entry.self_eval) {
          completedCount++;
          const se = entry.self_eval;
          html += `
            <div class="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden transition hover:shadow-md">
              <div class="bg-slate-50/80 border-b border-slate-200 px-6 py-4 flex flex-col md:flex-row md:items-center justify-between gap-3">
                <div class="flex items-center gap-3">
                  <div class="w-11 h-11 rounded-2xl bg-indigo-600 text-white font-bold flex items-center justify-center text-base shadow-sm shadow-indigo-100">
                    ${memName.slice(0, 1)}
                  </div>
                  <div>
                    <div class="flex items-center gap-2">
                      <h3 class="text-base font-bold text-slate-900">${memName}</h3>
                      <span class="px-2.5 py-0.5 text-xs font-semibold rounded-full bg-indigo-100 text-indigo-700">
                        ${se.job_role || '自評'}
                      </span>
                      <span class="px-2 py-0.5 text-xs font-medium rounded-full bg-emerald-100 text-emerald-700 flex items-center gap-1">
                        <i data-lucide="check-circle" class="w-3 h-3"></i> 已填寫自評
                      </span>
                    </div>
                    <p class="text-xs text-slate-400 mt-0.5">${entry.email} · 填寫時間：${entry.timestamp}</p>
                  </div>
                </div>
              </div>

              <div class="p-6 space-y-6">
                <!-- 1. 特質盤點 -->
                <div>
                  <h4 class="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2.5 flex items-center gap-1.5">
                    <i data-lucide="tag" class="w-3.5 h-3.5 text-indigo-500"></i> 工作特質盤點
                  </h4>
                  <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div class="bg-emerald-50/60 border border-emerald-100 rounded-xl p-4">
                      <div class="text-xs font-bold text-emerald-800 mb-2 flex items-center gap-1.5">
                        <i data-lucide="shield-check" class="w-4 h-4"></i> 最穩定、最具代表性 Top 3
                      </div>
                      <div class="flex flex-wrap gap-1.5">
                        ${se.top3_stable.map(t => `<span class="badge-stable px-3 py-1 text-xs font-semibold rounded-lg shadow-2xs">${t}</span>`).join('')}
                      </div>
                    </div>
                    <div class="bg-orange-50/60 border border-orange-100 rounded-xl p-4">
                      <div class="text-xs font-bold text-orange-800 mb-2 flex items-center gap-1.5">
                        <i data-lucide="trending-up" class="w-4 h-4"></i> 目前在練習 / 期望發展 3 項
                      </div>
                      <div class="flex flex-wrap gap-1.5">
                        ${se.top3_practice.map(t => `<span class="badge-practice px-3 py-1 text-xs font-semibold rounded-lg shadow-2xs">${t}</span>`).join('')}
                      </div>
                    </div>
                  </div>
                </div>

                <!-- 2. 四大文化實踐 -->
                <div>
                  <h4 class="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2.5 flex items-center gap-1.5">
                    <i data-lucide="heart" class="w-3.5 h-3.5 text-indigo-500"></i> 四大文化實踐實例 (STAR)
                  </h4>
                  <div class="grid grid-cols-1 md:grid-cols-2 gap-3.5">
                    <div class="bg-slate-50 border border-slate-200/80 rounded-xl p-4">
                      <div class="text-xs font-bold text-indigo-700 mb-1.5 flex items-center gap-1.5">
                        <span class="w-2 h-2 rounded-full bg-indigo-500"></span> 信任 (Trust)
                      </div>
                      <p class="text-xs text-slate-700 leading-relaxed">${se.values['信任'] || '（未填寫）'}</p>
                    </div>
                    <div class="bg-slate-50 border border-slate-200/80 rounded-xl p-4">
                      <div class="text-xs font-bold text-purple-700 mb-1.5 flex items-center gap-1.5">
                        <span class="w-2 h-2 rounded-full bg-purple-500"></span> 多元 (Diversity)
                      </div>
                      <p class="text-xs text-slate-700 leading-relaxed">${se.values['多元'] || '（未填寫）'}</p>
                    </div>
                    <div class="bg-slate-50 border border-slate-200/80 rounded-xl p-4">
                      <div class="text-xs font-bold text-amber-700 mb-1.5 flex items-center gap-1.5">
                        <span class="w-2 h-2 rounded-full bg-amber-500"></span> 實驗 (Experiment)
                      </div>
                      <p class="text-xs text-slate-700 leading-relaxed">${se.values['實驗'] || '（未填寫）'}</p>
                    </div>
                    <div class="bg-slate-50 border border-slate-200/80 rounded-xl p-4">
                      <div class="text-xs font-bold text-teal-700 mb-1.5 flex items-center gap-1.5">
                        <span class="w-2 h-2 rounded-full bg-teal-500"></span> 可持續 (Sustainability)
                      </div>
                      <p class="text-xs text-slate-700 leading-relaxed">${se.values['可持續'] || '（未填寫）'}</p>
                    </div>
                  </div>
                </div>

                <!-- 3. 職位專屬職能展現 -->
                ${se.competencies && se.competencies.length > 0 ? `
                  <div>
                    <h4 class="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2.5 flex items-center gap-1.5">
                      <i data-lucide="briefcase" class="w-3.5 h-3.5 text-indigo-500"></i> ${se.job_role} 專屬職能展現 (STAR 實例)
                    </h4>
                    <div class="space-y-3">
                      ${se.competencies.map(c => `
                        <div class="bg-slate-50 border border-slate-200 rounded-xl p-4">
                          <div class="text-xs font-bold text-slate-800 mb-1.5">${c.title}</div>
                          <p class="text-xs text-slate-600 leading-relaxed whitespace-pre-line">${c.answer || '（未填寫）'}</p>
                        </div>
                      `).join('')}
                    </div>
                  </div>
                ` : ''}

                <!-- 4. 組織卡關點與未來價值展望 -->
                ${se.reflection && Object.keys(se.reflection).length > 0 ? `
                  <div>
                    <h4 class="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2.5 flex items-center gap-1.5">
                      <i data-lucide="help-circle" class="w-3.5 h-3.5 text-indigo-500"></i> 組織卡關點與未來價值展望
                    </h4>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-3.5">
                      ${Object.entries(se.reflection).map(([k, v]) => `
                        <div class="bg-indigo-50/50 border border-indigo-100 rounded-xl p-4">
                          <div class="text-xs font-bold text-indigo-900 mb-1.5">${k}</div>
                          <p class="text-xs text-slate-700 leading-relaxed">${v || '（未填寫）'}</p>
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
            <div class="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 flex items-center justify-between">
              <div class="flex items-center gap-3">
                <div class="w-11 h-11 rounded-2xl bg-slate-100 text-slate-400 font-bold flex items-center justify-center text-base">
                  ${memName.slice(0, 1)}
                </div>
                <div>
                  <div class="flex items-center gap-2">
                    <h3 class="text-base font-bold text-slate-800">${memName}</h3>
                    <span class="px-2 py-0.5 text-xs font-medium rounded-full bg-amber-100 text-amber-700 flex items-center gap-1">
                      <i data-lucide="clock" class="w-3 h-3"></i> 尚未收到自評資料
                    </span>
                  </div>
                  <p class="text-xs text-slate-400 mt-0.5">此成員尚未於表單中送出自評紀錄</p>
                </div>
              </div>
              <div class="text-xs text-slate-400 font-medium">待填寫</div>
            </div>
          `;
        }
      });

      compStatus.innerText = `已填答 ${completedCount} / 應填 ${memberNames.length} 人`;
      container.innerHTML = html;
      lucide.createIcons();
    }

    // ==========================================
    // TAB 2: SUPERVISOR EVALUATION LOGIC
    // ==========================================
    function filterSupervisor(name) {
      currentSupervisorFilter = name;
      document.querySelectorAll('.sup-filter-btn').forEach(btn => {
        btn.classList.remove('bg-indigo-600', 'text-white');
        btn.classList.add('bg-slate-100', 'text-slate-600');
      });
      const activeBtn = document.getElementById('sup-filter-btn-' + name);
      if (activeBtn) {
        activeBtn.classList.add('bg-indigo-600', 'text-white');
        activeBtn.classList.remove('bg-slate-100', 'text-slate-600');
      }
      renderSupervisorSection();
    }

    function renderSupervisorSection() {
      let filtered = RAW_DATA.filter(e => e.relation === "主管");
      if (currentSupervisorFilter !== 'ALL') {
        filtered = filtered.filter(e => e.target === currentSupervisorFilter);
      }

      const avg = (key) => {
        const vals = filtered.map(e => e.supervisor_eval && e.supervisor_eval[key]).filter(v => v !== null && v !== undefined);
        return vals.length ? (vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(1) : 0;
      };

      const cultureLabels = ["信任 (真實表達)", "多元 (聆聽意見)", "實驗 (嘗試創新)", "心理安全 (試錯空間)", "肯定 (讚美認可)", "可持續 (尊重界線)"];
      const cultureData = [
        avg('q12_trust_express'),
        avg('q13_diversity_listen'),
        avg('q14_experiment_try'),
        avg('q15_experiment_psych_safety'),
        avg('q16_sustain_praise'),
        avg('q17_sustain_boundary')
      ];

      const mgmtLabels = ["尋求協助", "具體引導", "改善幅度", "跨部門推動", "資源評估", "失誤回應", "認可表現", "工作成效", "NPS推薦", "整體滿意"];
      const mgmtData = [
        avg('q4_help_easy'),
        avg('q5_guidance_freq'),
        avg('q6_improve_degree'),
        avg('q7_cross_dept'),
        avg('q8_resource_eval'),
        avg('q9_constructive_mistake'),
        avg('q10_recognition'),
        avg('q11_overall_performance'),
        avg('q18_nps_recommend'),
        avg('q19_satisfaction')
      ];

      if (supervisorRadar) supervisorRadar.destroy();
      const ctxRadar = document.getElementById('supervisorRadarChart').getContext('2d');
      supervisorRadar = new Chart(ctxRadar, {
        type: 'radar',
        data: {
          labels: cultureLabels,
          datasets: [{
            label: currentSupervisorFilter === 'ALL' ? '主管群平均' : currentSupervisorFilter,
            data: cultureData,
            backgroundColor: 'rgba(79, 70, 229, 0.2)',
            borderColor: '#4F46E5',
            pointBackgroundColor: '#4F46E5'
          }]
        },
        options: {
          scales: { r: { min: 0, max: 10, ticks: { stepSize: 2 } } },
          plugins: { legend: { display: false } }
        }
      });

      if (supervisorBar) supervisorBar.destroy();
      const ctxBar = document.getElementById('supervisorBarChart').getContext('2d');
      supervisorBar = new Chart(ctxBar, {
        type: 'bar',
        data: {
          labels: mgmtLabels,
          datasets: [{
            data: mgmtData,
            backgroundColor: '#6366F1',
            borderRadius: 6
          }]
        },
        options: {
          scales: { y: { min: 0, max: 10 } },
          plugins: { legend: { display: false } }
        }
      });

      const feedbackContainer = document.getElementById('supervisor-feedback-list');
      feedbackContainer.innerHTML = filtered.map(e => {
        const se = e.supervisor_eval;
        return `
          <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm space-y-4">
            <div class="flex items-center justify-between border-b border-slate-100 pb-3">
              <div class="flex items-center gap-2">
                <span class="font-bold text-slate-800 text-sm">受評主管：${e.target}</span>
                <span class="text-xs text-slate-400">（填答者：${e.email}）</span>
              </div>
              <div class="flex items-center gap-3 text-xs">
                <span class="px-2.5 py-1 rounded-lg bg-indigo-50 text-indigo-700 font-bold">NPS 推薦：${se.q18_nps_recommend || '-'} 分</span>
                <span class="px-2.5 py-1 rounded-lg bg-emerald-50 text-emerald-700 font-bold">滿意度：${se.q19_satisfaction || '-'} 分</span>
              </div>
            </div>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs">
              <div class="bg-slate-50 p-3 rounded-xl">
                <span class="font-bold text-indigo-800">Q20. 願景使命理解之引導：</span>
                <p class="mt-1 text-slate-700">${se.q20_vision_mission || '（無）'}</p>
              </div>
              <div class="bg-slate-50 p-3 rounded-xl">
                <span class="font-bold text-amber-800">Q21. 管理與文化精神建議：</span>
                <p class="mt-1 text-slate-700">${se.q21_improvement_advice || '（無）'}</p>
              </div>
              <div class="bg-slate-50 p-3 rounded-xl">
                <span class="font-bold text-slate-800">Q22. 其他補充評價：</span>
                <p class="mt-1 text-slate-700">${se.q22_other_comments || '（無）'}</p>
              </div>
              <div class="bg-pink-50 p-3 rounded-xl border border-pink-100">
                <span class="font-bold text-pink-700 flex items-center gap-1">
                  <i data-lucide="gift" class="w-3.5 h-3.5"></i> Q23. 肯定與感謝的話（好好星光大賞）：
                </span>
                <p class="mt-1 text-pink-900 font-medium">${se.q23_starlight_thanks || '（無）'}</p>
              </div>
            </div>
          </div>
        `;
      }).join('');

      lucide.createIcons();
    }

    // ==========================================
    // TAB 3: PEER EVALUATION LOGIC
    // ==========================================
    function initPeerPills() {
      const peers = Array.from(new Set(RAW_DATA.filter(e => e.relation === "同事").map(e => e.target)));
      const container = document.getElementById('peer-pills-container');
      let html = `
        <span class="text-xs font-bold text-slate-400 mr-1 flex items-center gap-1 uppercase tracking-wider">
          <i data-lucide="user" class="w-3.5 h-3.5"></i> 受評同事：
        </span>
        <button onclick="filterPeer('ALL')" id="peer-btn-ALL" class="peer-pill-btn px-3 py-1.5 rounded-xl text-xs font-semibold bg-emerald-600 text-white shadow-sm transition">
          全部 (${RAW_DATA.filter(e => e.relation === "同事").length}筆)
        </button>
      `;
      peers.forEach(p => {
        const count = RAW_DATA.filter(e => e.relation === "同事" && e.target === p).length;
        html += `
          <button onclick="filterPeer('${p}')" id="peer-btn-${p}" class="peer-pill-btn px-3 py-1.5 rounded-xl text-xs font-medium bg-slate-100 text-slate-600 hover:bg-slate-200 transition">
            ${p} (${count})
          </button>
        `;
      });
      container.innerHTML = html;
      lucide.createIcons();
    }

    function filterPeer(name) {
      currentPeerFilter = name;
      document.querySelectorAll('.peer-pill-btn').forEach(btn => {
        btn.classList.remove('bg-emerald-600', 'text-white');
        btn.classList.add('bg-slate-100', 'text-slate-600');
      });
      const activeBtn = document.getElementById('peer-btn-' + name);
      if (activeBtn) {
        activeBtn.classList.add('bg-emerald-600', 'text-white');
        activeBtn.classList.remove('bg-slate-100', 'text-slate-600');
      }
      renderPeerSection();
    }

    function renderPeerSection() {
      let filtered = RAW_DATA.filter(e => e.relation === "同事");
      if (currentPeerFilter !== 'ALL') {
        filtered = filtered.filter(e => e.target === currentPeerFilter);
      }

      const avg = (key) => {
        const vals = filtered.map(e => e.peer_eval && e.peer_eval[key]).filter(v => v !== null && v !== undefined);
        return vals.length ? (vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(1) : 0;
      };

      const cultureLabels = ["多元 (接受不同意見)", "多元 (建設性觀點)", "實驗 (開放調整)", "信任 (分享經驗資源)", "肯定 (讚美同事)", "可持續 (尊重界線)"];
      const cultureData = [
        avg('q30_open_to_opposing'),
        avg('q31_constructive_opinions'),
        avg('q32_growth_mindset'),
        avg('q33_share_knowledge'),
        avg('q34_praise_peers'),
        avg('q35_boundary_respect')
      ];

      const workLabels = ["合作狀況", "注重細節", "準時完成", "靈活調整", "追蹤承諾", "說明決策依據", "NPS推薦"];
      const workData = [
        avg('q24_cooperation'),
        avg('q25_detail_oriented'),
        avg('q26_on_time'),
        avg('q27_flexibility'),
        avg('q28_follow_up'),
        avg('q29_transparency'),
        avg('q36_nps_recommend')
      ];

      if (peerRadar) peerRadar.destroy();
      const ctxRadar = document.getElementById('peerRadarChart').getContext('2d');
      peerRadar = new Chart(ctxRadar, {
        type: 'radar',
        data: {
          labels: cultureLabels,
          datasets: [{
            label: currentPeerFilter === 'ALL' ? '全體同事平均' : currentPeerFilter,
            data: cultureData,
            backgroundColor: 'rgba(16, 185, 129, 0.2)',
            borderColor: '#10B981',
            pointBackgroundColor: '#10B981'
          }]
        },
        options: {
          scales: { r: { min: 0, max: 10, ticks: { stepSize: 2 } } },
          plugins: { legend: { display: false } }
        }
      });

      if (peerBar) peerBar.destroy();
      const ctxBar = document.getElementById('peerBarChart').getContext('2d');
      peerBar = new Chart(ctxBar, {
        type: 'bar',
        data: {
          labels: workLabels,
          datasets: [{
            data: workData,
            backgroundColor: '#10B981',
            borderRadius: 6
          }]
        },
        options: {
          scales: { y: { min: 0, max: 10 } },
          plugins: { legend: { display: false } }
        }
      });

      const feedbackContainer = document.getElementById('peer-feedback-list');
      feedbackContainer.innerHTML = filtered.map(e => {
        const pe = e.peer_eval;
        return `
          <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm space-y-4">
            <div class="flex items-center justify-between border-b border-slate-100 pb-3">
              <div class="flex items-center gap-2">
                <span class="font-bold text-slate-800 text-sm">受評同事：${e.target}</span>
                <span class="text-xs text-slate-400">（填答者：${e.email}）</span>
              </div>
              <div class="flex items-center gap-3 text-xs">
                <span class="px-2.5 py-1 rounded-lg bg-emerald-50 text-emerald-700 font-bold">NPS 推薦：${pe.q36_nps_recommend || '-'} 分</span>
                <span class="px-2.5 py-1 rounded-lg bg-slate-100 text-slate-700 font-medium">合作評分：${pe.q24_cooperation || '-'} 分</span>
              </div>
            </div>
            <div class="grid grid-cols-1 md:grid-cols-3 gap-3 text-xs">
              <div class="bg-slate-50 p-3 rounded-xl">
                <span class="font-bold text-slate-800">Q37. 工作與文化提升建議：</span>
                <p class="mt-1 text-slate-700">${pe.q37_improvement_advice || '（無）'}</p>
              </div>
              <div class="bg-slate-50 p-3 rounded-xl">
                <span class="font-bold text-slate-800">Q38. 其他補充評價：</span>
                <p class="mt-1 text-slate-700">${pe.q38_other_comments || '（無）'}</p>
              </div>
              <div class="bg-pink-50 p-3 rounded-xl border border-pink-100">
                <span class="font-bold text-pink-700 flex items-center gap-1">
                  <i data-lucide="gift" class="w-3.5 h-3.5"></i> Q39. 肯定與感謝的話（星光大賞）：
                </span>
                <p class="mt-1 text-pink-900 font-medium">${pe.q39_starlight_thanks || '（無）'}</p>
              </div>
            </div>
          </div>
        `;
      }).join('');

      lucide.createIcons();
    }

    // ==========================================
    // TAB 4: 360 OVERVIEW SECTION
    // ==========================================
    function initOverviewPills() {
      const allTargets = Array.from(new Set(RAW_DATA.map(e => e.target)));
      const members = ["何維安", "張希慈", "戴佑珍", "陳泳璇", "林文琇", "張芳媐", "姚品瑄", "胡喻翔", "薛筑瑄"];
      // merge any extra targets
      allTargets.forEach(t => {
        if (!members.includes(t)) members.push(t);
      });

      const container = document.getElementById('overview-member-pills');
      let html = `
        <span class="text-xs font-bold text-slate-400 mr-1 flex items-center gap-1 uppercase tracking-wider">
          <i data-lucide="sparkles" class="w-3.5 h-3.5"></i> 成員報告：
        </span>
      `;
      members.forEach(m => {
        html += `
          <button onclick="selectOverviewMember('${m}')" id="overview-btn-${m}" class="overview-pill-btn px-3.5 py-1.5 rounded-xl text-xs font-medium bg-slate-100 text-slate-600 hover:bg-slate-200 transition">
            ${m}
          </button>
        `;
      });
      container.innerHTML = html;
      selectOverviewMember(currentOverviewMember);
    }

    function selectOverviewMember(name) {
      currentOverviewMember = name;
      document.querySelectorAll('.overview-pill-btn').forEach(btn => {
        btn.classList.remove('bg-indigo-600', 'text-white');
        btn.classList.add('bg-slate-100', 'text-slate-600');
      });
      const activeBtn = document.getElementById('overview-btn-' + name);
      if (activeBtn) {
        activeBtn.classList.add('bg-indigo-600', 'text-white');
        activeBtn.classList.remove('bg-slate-100', 'text-slate-600');
      }
      renderOverviewSection();
    }

    function renderOverviewSection() {
      const name = currentOverviewMember;
      const selfEntry = RAW_DATA.find(e => e.target === name && e.relation === "自評");
      const supEntries = RAW_DATA.filter(e => e.target === name && e.relation === "主管");
      const peerEntries = RAW_DATA.filter(e => e.target === name && e.relation === "同事");

      const container = document.getElementById('overview-report-container');

      const peerAvg = (k) => {
        const v = peerEntries.map(e => e.peer_eval && e.peer_eval[k]).filter(x => x !== null && x !== undefined);
        return v.length ? (v.reduce((a,b)=>a+b, 0)/v.length).toFixed(1) : "-";
      };
      const supAvg = (k) => {
        const v = supEntries.map(e => e.supervisor_eval && e.supervisor_eval[k]).filter(x => x !== null && x !== undefined);
        return v.length ? (v.reduce((a,b)=>a+b, 0)/v.length).toFixed(1) : "-";
      };

      const se = selfEntry ? selfEntry.self_eval : null;

      container.innerHTML = `
        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-6">
          <div class="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-5">
            <div class="flex items-center gap-3">
              <div class="w-12 h-12 rounded-2xl bg-indigo-600 text-white font-bold flex items-center justify-center text-lg shadow-md shadow-indigo-100">
                ${name.slice(0, 1)}
              </div>
              <div>
                <div class="flex items-center gap-2">
                  <h2 class="text-xl font-bold text-slate-900">${name} 360 綜合評估報告</h2>
                  ${se ? `<span class="px-2.5 py-0.5 text-xs font-semibold rounded-full bg-indigo-100 text-indigo-700">${se.job_role}</span>` : ''}
                </div>
                <p class="text-xs text-slate-500 mt-1">包含自評 (${selfEntry ? '1' : '0'}份)、同儕回饋 (${peerEntries.length}份)、主管回饋 (${supEntries.length}份)</p>
              </div>
            </div>
            <div class="flex items-center gap-2 text-xs flex-wrap">
              <span class="px-3 py-1.5 rounded-xl bg-slate-100 text-slate-700 font-bold">同事 NPS：${peerAvg('q36_nps_recommend')} 分</span>
              ${supEntries.length ? `<span class="px-3 py-1.5 rounded-xl bg-indigo-50 text-indigo-700 font-bold">主管 NPS：${supAvg('q18_nps_recommend')} 分</span>` : ''}
            </div>
          </div>

          <!-- Self Traits Quick View -->
          ${se ? `
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div class="bg-emerald-50/50 border border-emerald-100 rounded-xl p-4">
                <div class="text-xs font-bold text-emerald-800 mb-2 flex items-center gap-1.5">
                  <i data-lucide="shield-check" class="w-4 h-4"></i> 自評最穩定代表特質 Top 3
                </div>
                <div class="flex flex-wrap gap-1.5">
                  ${se.top3_stable.map(t => `<span class="badge-stable px-2.5 py-1 text-xs font-semibold rounded-lg">${t}</span>`).join('')}
                </div>
              </div>
              <div class="bg-orange-50/50 border border-orange-100 rounded-xl p-4">
                <div class="text-xs font-bold text-orange-800 mb-2 flex items-center gap-1.5">
                  <i data-lucide="trending-up" class="w-4 h-4"></i> 自評目前練習中 / 期望發展 3 項
                </div>
                <div class="flex flex-wrap gap-1.5">
                  ${se.top3_practice.map(t => `<span class="badge-practice px-2.5 py-1 text-xs font-semibold rounded-lg">${t}</span>`).join('')}
                </div>
              </div>
            </div>
          ` : `
            <div class="bg-amber-50 border border-amber-100 rounded-xl p-4 text-xs text-amber-800">
              ⚠️ ${name} 尚未在表單中提交自評資料。下方為夥伴與主管之評價。
            </div>
          `}

          <!-- Starlight大会感謝詞 -->
          <div>
            <h3 class="text-sm font-bold text-pink-700 mb-3 flex items-center gap-2">
              <i data-lucide="sparkles" class="w-4 h-4 text-pink-500"></i> 好好星光大賞：夥伴們給 ${name} 的感謝與肯定
            </h3>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
              ${[...supEntries, ...peerEntries].filter(e => (e.supervisor_eval?.q23_starlight_thanks || e.peer_eval?.q39_starlight_thanks)).map(e => {
                const text = e.supervisor_eval?.q23_starlight_thanks || e.peer_eval?.q39_starlight_thanks;
                return `
                  <div class="bg-pink-50/60 border border-pink-100 rounded-xl p-3.5 text-xs text-pink-950 font-medium flex items-start gap-2">
                    <i data-lucide="heart" class="w-4 h-4 text-pink-500 shrink-0 mt-0.5"></i>
                    <div>
                      <p class="leading-relaxed">"${text}"</p>
                      <span class="text-[10px] text-pink-400 mt-1 block">— 來自 ${e.relation} (${e.email.split('@')[0]})</span>
                    </div>
                  </div>
                `;
              }).join('') || '<p class="text-xs text-slate-400">目前尚無星光感謝詞紀錄</p>'}
            </div>
          </div>
        </div>
      `;
      lucide.createIcons();
    }

    // ==========================================
    // CLIENT-SIDE XLSX EXPORT VIA SHEETJS
    // ==========================================
    function exportSupervisorExcelClientSide(supName) {
      const wb = XLSX.utils.book_new();
      const memberNames = SUPERVISOR_TEAMS[supName] || [];

      const rows = [];
      rows.push([`好好星球文化基金會 360 年中成長評估 - 【${supName}】部屬自評專用表`]);
      rows.push([]);

      memberNames.forEach(memName => {
        const memEntry = RAW_DATA.find(e => e.target === memName && e.relation === "自評");
        const role = memEntry && memEntry.self_eval ? memEntry.self_eval.job_role : "尚未填寫自評";
        rows.push([`👤 部屬姓名：${memName} （職位：${role}）`, "", "", ""]);
        rows.push(["評估面向 / 題組", "題目標題與說明", "部屬自評回覆內容", "主管評核與回饋備註"]);

        if (!memEntry || !memEntry.self_eval) {
          rows.push(["狀態說明", "填答紀錄", `目前表單中尚未收到 ${memName} 的自我評估回覆紀錄。`, ""]);
          rows.push([]);
          return;
        }

        const se = memEntry.self_eval;
        rows.push(["基本資訊", "填答時間 / 電子郵件", `${memEntry.timestamp} / ${memEntry.email}`, ""]);
        rows.push(["基本資訊", "評估職位", se.job_role || "未填寫", ""]);
        rows.push(["工作特質盤點", "過去展現最穩定、最具代表性的特質 Top 3", se.top3_stable.join("、"), ""]);
        rows.push(["工作特質盤點", "目前還在練習、或希望未來更穩定發展的特質 3項", se.top3_practice.join("、"), ""]);
        rows.push(["四大文化實踐", "【信任】能獨立行動與決策、主動協作、雙向溝通，在高彈性下給予彼此信任", se.values['信任'] || "（無填寫）", ""]);
        rows.push(["四大文化實踐", "【多元】尊重差異、多元工作方法、主動表達不同觀點與想法", se.values['多元'] || "（無填寫）", ""]);
        rows.push(["四大文化實踐", "【實驗】透過開放的心態不斷嘗試、修正與反思，勇於檢討及給予回饋", se.values['實驗'] || "（無填寫）", ""]);
        rows.push(["四大文化實踐", "【可持續】內在韌性、自我照顧、彈性的人際與工作邊界", se.values['可持續'] || "（無填寫）", ""]);

        if (se.competencies) {
          se.competencies.forEach(c => {
            rows.push([`職能展現 (${se.job_role})`, c.title, c.answer || "（無填寫）", ""]);
          });
        }

        if (se.reflection) {
          Object.entries(se.reflection).forEach(([k, v]) => {
            rows.push(["組織卡關點與展望", k, v || "（無填寫）", ""]);
          });
        }

        rows.push([]);
      });

      const ws = XLSX.utils.aoa_to_sheet(rows);
      ws['!cols'] = [{ wch: 18 }, { wch: 38 }, { wch: 70 }, { wch: 30 }];
      XLSX.utils.book_append_sheet(wb, ws, "部屬自評彙整表");
      XLSX.writeFile(wb, `自評_${supName}主管專用_部屬自評彙整表.xlsx`);
      showToast(`📥 已成功下載【${supName}】部屬自評 XLSX！`);
    }

    function exportFullWorkbookClientSide() {
      const wb = XLSX.utils.book_new();

      // 1. All Supervisor Team Sheets
      ["張希慈", "何維安", "姚品瑄", "張希慈_執行長"].forEach(supName => {
        const memberNames = SUPERVISOR_TEAMS[supName] || [];
        const rows = [];
        rows.push([`好好星球文化基金會 360 年中成長評估 - 【${supName}】部屬自評專用表`]);
        rows.push([]);

        memberNames.forEach(memName => {
          const memEntry = RAW_DATA.find(e => e.target === memName && e.relation === "自評");
          const role = memEntry && memEntry.self_eval ? memEntry.self_eval.job_role : "尚未填寫自評";
          rows.push([`👤 部屬姓名：${memName} （職位：${role}）`, "", "", ""]);
          rows.push(["評估面向 / 題組", "題目標題與說明", "部屬自評回覆內容", "主管評核與回饋備註"]);

          if (!memEntry || !memEntry.self_eval) {
            rows.push(["狀態說明", "填答紀錄", `目前表單中尚未收到 ${memName} 的自我評估回覆紀錄。`, ""]);
            rows.push([]);
            return;
          }

          const se = memEntry.self_eval;
          rows.push(["基本資訊", "填答時間 / 電子郵件", `${memEntry.timestamp} / ${memEntry.email}`, ""]);
          rows.push(["基本資訊", "評估職位", se.job_role || "未填寫", ""]);
          rows.push(["工作特質盤點", "過去展現最穩定、最具代表性的特質 Top 3", se.top3_stable.join("、"), ""]);
          rows.push(["工作特質盤點", "目前還在練習、或希望未來更穩定發展的特質 3項", se.top3_practice.join("、"), ""]);
          rows.push(["四大文化實踐", "【信任】能獨立行動與決策、主動協作、雙向溝通，在高彈性下給予彼此信任", se.values['信任'] || "（無填寫）", ""]);
          rows.push(["四大文化實踐", "【多元】尊重差異、多元工作方法、主動表達不同觀點與想法", se.values['多元'] || "（無填寫）", ""]);
          rows.push(["四大文化實踐", "【實驗】透過開放的心態不斷嘗試、修正與反思，勇於檢討及給予回饋", se.values['實驗'] || "（無填寫）", ""]);
          rows.push(["四大文化實踐", "【可持續】內在韌性、自我照顧、彈性的人際與工作邊界", se.values['可持續'] || "（無填寫）", ""]);

          if (se.competencies) {
            se.competencies.forEach(c => {
              rows.push([`職能展現 (${se.job_role})`, c.title, c.answer || "（無填寫）", ""]);
            });
          }

          if (se.reflection) {
            Object.entries(se.reflection).forEach(([k, v]) => {
              rows.push(["組織卡關點與展望", k, v || "（無填寫）", ""]);
            });
          }

          rows.push([]);
        });

        const ws = XLSX.utils.aoa_to_sheet(rows);
        ws['!cols'] = [{ wch: 18 }, { wch: 38 }, { wch: 70 }, { wch: 30 }];
        const cleanTitle = supName === "張希慈_執行長" ? "自評_執行長" : `自評_${supName}`;
        XLSX.utils.book_append_sheet(wb, ws, cleanTitle);
      });

      // 2. Supervisor Evaluation Sheet
      const supRows = [];
      supRows.push(["時間戳記", "填答者Email", "受評主管", "Q4.尋求協助", "Q5.具體引導", "Q6.改善幅度", "Q7.跨部門推動", "Q8.資源評估", "Q9.失誤回應", "Q10.肯定認可", "Q11.工作成效", "Q12.【信任】表達想法", "Q13.【多元】聆聽意見", "Q14.【實驗】嘗試創新", "Q15.【實驗】試錯空間", "Q16.【可持續】讚美肯定", "Q17.【可持續】尊重界線", "Q18.NPS推薦度", "Q19.滿意度", "Q20.願景使命引導", "Q21.管理提升建議", "Q22.其他補充評價", "Q23.肯定感謝詞(星光大賞)"]);
      RAW_DATA.filter(e => e.relation === "主管").forEach(e => {
        const se = e.supervisor_eval || {};
        supRows.push([
          e.timestamp, e.email, e.target,
          se.q4_help_easy, se.q5_guidance_freq, se.q6_improve_degree, se.q7_cross_dept,
          se.q8_resource_eval, se.q9_constructive_mistake, se.q10_recognition, se.q11_overall_performance,
          se.q12_trust_express, se.q13_diversity_listen, se.q14_experiment_try,
          se.q15_experiment_psych_safety, se.q16_sustain_praise, se.q17_sustain_boundary,
          se.q18_nps_recommend, se.q19_satisfaction,
          se.q20_vision_mission, se.q21_improvement_advice, se.q22_other_comments, se.q23_starlight_thanks
        ]);
      });
      const wsSup = XLSX.utils.aoa_to_sheet(supRows);
      wsSup['!cols'] = Array(19).fill({ wch: 15 }).concat([{ wch: 35 }, { wch: 35 }, { wch: 35 }, { wch: 35 }]);
      XLSX.utils.book_append_sheet(wb, wsSup, "評主管總表");

      // 3. Peer Evaluation Sheet
      const peerRows = [];
      peerRows.push(["時間戳記", "填答者Email", "受評同事", "Q24.合作狀況", "Q25.注重細節", "Q26.準時完成", "Q27.靈活調整", "Q28.追蹤承諾", "Q29.說明依據", "Q30.【多元】接受不同意見", "Q31.【多元】建設性觀點", "Q32.【實驗】開放調整", "Q33.【信任】分享經驗", "Q34.【可持續】讚美同事", "Q35.【可持續】尊重界線", "Q36.NPS推薦度", "Q37.提升建議", "Q38.其他評價", "Q39.肯定感謝詞(星光大賞)"]);
      RAW_DATA.filter(e => e.relation === "同事").forEach(e => {
        const pe = e.peer_eval || {};
        peerRows.push([
          e.timestamp, e.email, e.target,
          pe.q24_cooperation, pe.q25_detail_oriented, pe.q26_on_time, pe.q27_flexibility,
          pe.q28_follow_up, pe.q29_transparency,
          pe.q30_open_to_opposing, pe.q31_constructive_opinions, pe.q32_growth_mindset,
          pe.q33_share_knowledge, pe.q34_praise_peers, pe.q35_boundary_respect,
          pe.q36_nps_recommend,
          pe.q37_improvement_advice, pe.q38_other_comments, pe.q39_starlight_thanks
        ]);
      });
      const wsPeer = XLSX.utils.aoa_to_sheet(peerRows);
      wsPeer['!cols'] = Array(16).fill({ wch: 15 }).concat([{ wch: 35 }, { wch: 35 }, { wch: 35 }]);
      XLSX.utils.book_append_sheet(wb, wsPeer, "評同事總表");

      XLSX.writeFile(wb, "好好星球_360年中成長評估_主管分流與完整彙整表.xlsx");
      showToast("📥 已成功下載全組織完整 Excel 總表！");
    }

    // Initialize everything on load
    window.addEventListener('DOMContentLoaded', () => {
      renderSelfSection();
      initPeerPills();
      initOverviewPills();
      lucide.createIcons();
    });
  </script>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_template)

print("index.html updated with dynamic CSV upload, dynamic XLSX export, and updated hierarchy!")
