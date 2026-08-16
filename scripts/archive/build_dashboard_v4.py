import json

with open("evaluation_data.json", "r", encoding="utf-8") as f:
    data_json_str = f.read()

html_template = """<!DOCTYPE html>
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
  <!-- SheetJS for client-side Excel export -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <!-- Lucide Icons -->
  <script src="https://unpkg.com/lucide@latest"></script>
  
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@500;600;700&family=Plus+Jakarta+Sans:wght@400;500;600;700&family=Noto+Sans+TC:wght@400;500;600;700&display=swap');
    
    :root {
      --bg-page: #FAF7F2;
      --bg-card: #FFFDF9;
      --bg-box: #F9F6EE;
      --bg-box-highlight: #FCE5CD;
      --color-primary: #557A61;
      --color-primary-dark: #3E5A47;
      --color-primary-light: #E4ECD3;
      --color-blush: #F4CCCC;
      --color-apricot: #FCE5CD;
      --color-border: #E8E2D8;
      --color-border-subtle: #EFEAE1;
      --color-text-main: #2E2827;
      --color-text-body: #4A433E;
      --color-text-muted: #8C837C;
    }

    body {
      font-family: 'Noto Sans TC', 'Plus Jakarta Sans', sans-serif;
      background-color: var(--bg-page);
      color: var(--color-text-main);
      letter-spacing: -0.01em;
      line-height: 1.7;
    }

    .font-serif-tc {
      font-family: 'Noto Serif TC', serif;
    }

    .custom-scrollbar::-webkit-scrollbar {
      width: 6px;
      height: 6px;
    }
    .custom-scrollbar::-webkit-scrollbar-thumb {
      background-color: #D9D2C9;
      border-radius: 4px;
    }

    /* Accessibility & Button States */
    button:focus-visible, a:focus-visible, label:focus-visible, select:focus-visible {
      outline: 2px solid var(--color-primary);
      outline-offset: 2px;
    }
    button:active, a:active {
      transform: scale(0.985);
    }

    @media (prefers-reduced-motion: reduce) {
      * {
        animation-duration: 0.01ms !important;
        animation-iteration-count: 1 !important;
        transition-duration: 0.01ms !important;
      }
    }

    /* Badges */
    .badge-stable {
      background-color: #E4ECD3;
      color: #2D5239;
      border: 1px solid #CDE0BC;
    }
    .badge-practice {
      background-color: #FCE5CD;
      color: #783E16;
      border: 1px solid #F5D3B3;
    }

    /* Level Badges */
    .badge-l5 { background-color: #E4ECD3; color: #2D5239; border: 1px solid #CDE0BC; }
    .badge-l4 { background-color: #E2F3F0; color: #235954; border: 1px solid #C4E6E1; }
    .badge-l3 { background-color: #FFF4CD; color: #7A5E12; border: 1px solid #FCE299; }
    .badge-l2 { background-color: #FCE5CD; color: #8C4B1E; border: 1px solid #F5D3B3; }
    .badge-l1 { background-color: #F4CCCC; color: #7A363A; border: 1px solid #E6BDBD; }

    /* Chunk Header styles */
    .chunk-header-banner {
      background-color: var(--color-blush);
      color: #3E2426;
      border-bottom: 1px solid #E6BDBD;
    }
    .chunk-subheader-banner {
      background-color: var(--color-apricot);
      color: #4A2E1C;
      border-bottom: 1px solid #EED1B4;
    }

    .drop-zone-active {
      border-color: #557A61 !important;
      background-color: #E4ECD3 !important;
    }

    .soft-card-shadow {
      box-shadow: 0 4px 20px -2px rgba(120, 100, 80, 0.06), 0 2px 6px -1px rgba(120, 100, 80, 0.04);
    }
  </style>
</head>
<body class="min-h-screen flex flex-col antialiased">

  <!-- TOP HEADER -->
  <header class="bg-[#FFFDF9]/95 backdrop-blur-md border-b border-[#E8E2D8] sticky top-0 z-50 shadow-xs">
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
              <span class="px-2.5 py-0.5 text-xs font-semibold rounded-md bg-[#FCE5CD] text-[#783E16] border border-[#F3D1B0]">
                2026 年中版
              </span>
            </div>
            <p class="text-xs text-[#7A726D] mt-0.5">視覺化篩選與主管部屬專屬 Chunk 報告導出系統</p>
          </div>
        </div>

        <!-- ACTION BUTTONS (UPLOAD CSV + EXCEL DROPDOWN) -->
        <div class="flex items-center gap-3">
          <!-- UPLOAD CSV BUTTON -->
          <label class="cursor-pointer inline-flex items-center gap-2 px-4 py-2.5 text-xs sm:text-sm font-semibold rounded-xl bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] hover:text-[#2E2827] transition border border-[#E0D7CA]">
            <i data-lucide="upload" class="w-4 h-4 text-[#557A61]"></i>
            <span class="hidden sm:inline">上傳最新 CSV</span>
            <span class="sm:hidden">上傳</span>
            <input type="file" id="csvFileInput" accept=".csv" class="hidden" onchange="handleFileUpload(event)">
          </label>

          <!-- EXCEL EXPORT DROPDOWN -->
          <div class="relative inline-block text-left" id="exportDropdown">
            <button onclick="toggleDropdown()" class="inline-flex items-center gap-2 px-4.5 py-2.5 text-xs sm:text-sm font-semibold rounded-xl bg-[#557A61] text-white hover:bg-[#466551] transition shadow-xs">
              <i data-lucide="download" class="w-4 h-4"></i>
              <span class="hidden sm:inline">下載主管專用 Excel (Chunk 格式)</span>
              <span class="sm:hidden">下載 Excel</span>
              <i data-lucide="chevron-down" class="w-4 h-4 opacity-80"></i>
            </button>
            
            <div id="dropdownMenu" class="hidden absolute right-0 mt-2 w-88 origin-top-right rounded-2xl bg-[#FFFDF9] p-3 shadow-2xl ring-1 ring-black/5 z-50 divide-y divide-[#EFEAE1] border border-[#E8E2D8]">
              <div class="py-2">
                <div class="px-3 py-1 text-[11px] font-bold text-[#8C837C] uppercase tracking-wider">
                  主管部屬自評專用包 (Chunk 格式與題目)
                </div>
                <button onclick="exportSupervisorExcelClientSide('張希慈')" class="w-full text-left flex items-center gap-3 px-3 py-2.5 text-xs sm:text-sm text-[#2E2827] hover:bg-[#FCE5CD]/40 rounded-xl transition">
                  <div class="p-2 bg-[#F4CCCC] text-[#4A2426] rounded-xl"><i data-lucide="file-spreadsheet" class="w-4 h-4"></i></div>
                  <div>
                    <div class="font-bold text-[#2E2827]">【張希慈】部屬自評彙整表</div>
                    <div class="text-[11px] text-[#7A726D]">何維安、陳泳璇、張芳媐、姚品瑄、胡喻翔</div>
                  </div>
                </button>
                <button onclick="exportSupervisorExcelClientSide('何維安')" class="w-full text-left flex items-center gap-3 px-3 py-2.5 text-xs sm:text-sm text-[#2E2827] hover:bg-[#FCE5CD]/40 rounded-xl transition">
                  <div class="p-2 bg-[#F4CCCC] text-[#4A2426] rounded-xl"><i data-lucide="file-spreadsheet" class="w-4 h-4"></i></div>
                  <div>
                    <div class="font-bold text-[#2E2827]">【何維安】部屬自評彙整表</div>
                    <div class="text-[11px] text-[#7A726D]">林文琇</div>
                  </div>
                </button>
                <button onclick="exportSupervisorExcelClientSide('姚品瑄')" class="w-full text-left flex items-center gap-3 px-3 py-2.5 text-xs sm:text-sm text-[#2E2827] hover:bg-[#FCE5CD]/40 rounded-xl transition">
                  <div class="p-2 bg-[#F4CCCC] text-[#4A2426] rounded-xl"><i data-lucide="file-spreadsheet" class="w-4 h-4"></i></div>
                  <div>
                    <div class="font-bold text-[#2E2827]">【姚品瑄】部屬自評彙整表</div>
                    <div class="text-[11px] text-[#7A726D]">薛筑瑄、戴佑珍</div>
                  </div>
                </button>
                <button onclick="exportSupervisorExcelClientSide('張希慈_執行長')" class="w-full text-left flex items-center gap-3 px-3 py-2.5 text-xs sm:text-sm text-[#2E2827] hover:bg-[#FCE5CD]/40 rounded-xl transition">
                  <div class="p-2 bg-[#F4CCCC] text-[#4A2426] rounded-xl"><i data-lucide="file-spreadsheet" class="w-4 h-4"></i></div>
                  <div>
                    <div class="font-bold text-[#2E2827]">【張希慈】執行長個人自評表</div>
                    <div class="text-[11px] text-[#7A726D]">執行長職位專用</div>
                  </div>
                </button>
              </div>
              <div class="py-2">
                <div class="px-3 py-1 text-[11px] font-bold text-[#8C837C] uppercase tracking-wider">
                  全組織完整總表
                </div>
                <button onclick="exportFullWorkbookClientSide()" class="w-full text-left flex items-center gap-3 px-3 py-2.5 text-xs sm:text-sm font-bold text-[#557A61] hover:bg-[#E4ECD3]/40 rounded-xl transition">
                  <div class="p-2 bg-[#E4ECD3] text-[#2D5239] rounded-xl"><i data-lucide="layers" class="w-4 h-4"></i></div>
                  <div>
                    <div>下載完整 Excel (全 Sheets 整合)</div>
                    <div class="text-[11px] text-[#557A61] font-normal">包含評主管、評同事、所有主管包</div>
                  </div>
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>

      <!-- MAIN TABS -->
      <nav class="flex space-x-2 border-t border-[#E8E2D8] pt-2.5 -mb-px overflow-x-auto custom-scrollbar">
        <button onclick="switchTab('self')" id="tab-btn-self" class="tab-btn inline-flex items-center gap-2 px-4.5 py-3 text-xs sm:text-sm font-semibold rounded-t-xl border-b-2 border-[#557A61] text-[#2D5239] bg-[#E4ECD3]/30 shrink-0 transition">
          <i data-lucide="user-check" class="w-4 h-4 text-[#557A61]"></i>
          自評（依主管分流）
          <span class="px-2 py-0.5 text-xs rounded-full bg-[#E4ECD3] text-[#2D5239] font-bold" id="badge-self-count">5</span>
        </button>
        <button onclick="switchTab('supervisor')" id="tab-btn-supervisor" class="tab-btn inline-flex items-center gap-2 px-4.5 py-3 text-xs sm:text-sm font-medium rounded-t-xl border-b-2 border-transparent text-[#6E6662] hover:text-[#2E2827] hover:bg-[#F2EEE6]/50 shrink-0 transition">
          <i data-lucide="award" class="w-4 h-4"></i>
          評主管
          <span class="px-2 py-0.5 text-xs rounded-full bg-[#F2EEE6] text-[#6E6662] font-bold" id="badge-sup-count">4</span>
        </button>
        <button onclick="switchTab('peer')" id="tab-btn-peer" class="tab-btn inline-flex items-center gap-2 px-4.5 py-3 text-xs sm:text-sm font-medium rounded-t-xl border-b-2 border-transparent text-[#6E6662] hover:text-[#2E2827] hover:bg-[#F2EEE6]/50 shrink-0 transition">
          <i data-lucide="users" class="w-4 h-4"></i>
          評同事
          <span class="px-2 py-0.5 text-xs rounded-full bg-[#F2EEE6] text-[#6E6662] font-bold" id="badge-peer-count">23</span>
        </button>
        <button onclick="switchTab('overview360')" id="tab-btn-overview360" class="tab-btn inline-flex items-center gap-2 px-4.5 py-3 text-xs sm:text-sm font-medium rounded-t-xl border-b-2 border-transparent text-[#6E6662] hover:text-[#2E2827] hover:bg-[#F2EEE6]/50 shrink-0 transition">
          <i data-lucide="pie-chart" class="w-4 h-4"></i>
          個人 360 綜合雷達
        </button>
      </nav>
    </div>
  </header>

  <!-- DRAG & DROP NOTIFICATION ZONE -->
  <div id="dropZone" ondragover="handleDragOver(event)" ondragleave="handleDragLeave(event)" ondrop="handleDrop(event)" class="transition-all duration-200 border-2 border-dashed border-[#D5CCC0] bg-[#FFFDF9] max-w-7xl mx-auto w-full px-5 py-3.5 my-3.5 rounded-2xl text-center text-xs sm:text-sm text-[#7A726D] hidden sm:flex items-center justify-center gap-2.5 hover:border-[#557A61]">
    <i data-lucide="file-up" class="w-4 h-4 text-[#557A61]"></i>
    <span>支援將最新 Google 表單匯出的 CSV 檔案直接<b>拖曳至此處</b>，即可即時重新整理數據與 Excel 報表。</span>
  </div>

  <!-- MAIN CONTENT AREA -->
  <main class="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-5 sm:py-7">

    <!-- ======================================================== -->
    <!-- TAB 1: 自評 (依主管分流) -->
    <!-- ======================================================== -->
    <section id="tab-section-self" class="tab-content space-y-7">
      
      <!-- SUB-FILTER FOR SUPERVISOR TEAMS -->
      <div class="bg-[#FFFDF9] rounded-2xl p-5 sm:p-6 border border-[#E8E2D8] soft-card-shadow flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div class="flex items-center gap-2.5 flex-wrap" id="self-supervisor-pills">
          <span class="text-xs font-bold text-[#8C837C] mr-1 flex items-center gap-1.5 uppercase tracking-wider">
            <i data-lucide="filter" class="w-3.5 h-3.5 text-[#557A61]"></i> 主管團隊：
          </span>
          <button onclick="filterSelfSupervisor('張希慈')" id="self-sup-btn-張希慈" class="self-sup-btn px-4.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold bg-[#557A61] text-white shadow-xs transition">
            張希慈 的部屬 (5人)
          </button>
          <button onclick="filterSelfSupervisor('何維安')" id="self-sup-btn-何維安" class="self-sup-btn px-4.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
            何維安 的部屬 (1人)
          </button>
          <button onclick="filterSelfSupervisor('姚品瑄')" id="self-sup-btn-姚品瑄" class="self-sup-btn px-4.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
            姚品瑄 的部屬 (2人)
          </button>
          <button onclick="filterSelfSupervisor('張希慈_執行長')" id="self-sup-btn-張希慈_執行長" class="self-sup-btn px-4.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
            執行長個人自評
          </button>
          <button onclick="filterSelfSupervisor('ALL')" id="self-sup-btn-ALL" class="self-sup-btn px-4.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
            全部已自評 (5人)
          </button>
        </div>

        <div id="self-export-btn-container">
          <!-- Dynamic Export Button -->
        </div>
      </div>

      <!-- RATING GUIDE CARD (UNIFIED COHESIVE THEME) -->
      <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] soft-card-shadow overflow-hidden">
        <div class="chunk-subheader-banner px-6 py-4 flex items-center justify-between cursor-pointer hover:bg-[#FCE5CD]/90 transition" onclick="toggleRatingGuide()">
          <div class="flex items-center gap-3">
            <div class="p-1.5 bg-[#4A2E1C] text-white rounded-lg"><i data-lucide="book-open" class="w-4 h-4"></i></div>
            <h3 class="text-sm sm:text-base font-bold text-[#4A2E1C] font-serif-tc">面向分數落點標準說明（Start / Grow / Keep / Good / Amazing!）</h3>
          </div>
          <div class="flex items-center gap-1.5 text-xs sm:text-sm font-semibold text-[#7A4822]">
            <span id="guide-toggle-text">收合說明</span>
            <i data-lucide="chevron-down" id="guide-toggle-icon" class="w-4 h-4 transform rotate-180 transition-transform"></i>
          </div>
        </div>
        
        <div id="rating-guide-content" class="p-6 sm:p-7 overflow-x-auto">
          <table class="w-full text-xs sm:text-sm text-left border-collapse">
            <thead>
              <tr class="border-b border-[#E8E2D8] text-[#8C837C]">
                <th class="py-3 px-4 font-bold text-center w-24">等級 (Level)</th>
                <th class="py-3 px-4 font-bold text-center w-28">落點名稱</th>
                <th class="py-3 px-4 font-bold">定義說明</th>
                <th class="py-3 px-4 font-bold">回饋語氣與後續動作</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-[#EFEAE1]">
              <tr class="bg-[#E4ECD3]/35 hover:bg-[#E4ECD3]/50 transition">
                <td class="py-3.5 px-4 text-center font-bold text-[#2D5239]">L5</td>
                <td class="py-3.5 px-4 text-center font-bold text-[#2D5239]"><span class="badge-l5 px-2.5 py-1 rounded-md shadow-2xs">Amazing!</span></td>
                <td class="py-3.5 px-4 text-[#2E2827] leading-relaxed">遠超職位期待，表現為團隊之標竿與典範。</td>
                <td class="py-3.5 px-4 text-[#4A433E] leading-relaxed">讚賞其突出貢獻，探討經驗複製機制，在成長對話中轉為「可以帶其他人一起做」的方向。</td>
              </tr>
              <tr class="bg-[#E2F3F0]/35 hover:bg-[#E2F3F0]/50 transition">
                <td class="py-3.5 px-4 text-center font-bold text-[#235954]">L4</td>
                <td class="py-3.5 px-4 text-center font-bold text-[#235954]"><span class="badge-l4 px-2.5 py-1 rounded-md shadow-2xs">Good</span></td>
                <td class="py-3.5 px-4 text-[#2E2827] leading-relaxed">優於職位期待，持續展現高標準成果。</td>
                <td class="py-3.5 px-4 text-[#4A433E] leading-relaxed">肯定並具體指出是哪些行為讓它超出標準，設定具挑戰性的下一步目標。</td>
              </tr>
              <tr class="bg-[#FFF4CD]/35 hover:bg-[#FFF4CD]/50 transition">
                <td class="py-3.5 px-4 text-center font-bold text-[#7A5E12]">L3</td>
                <td class="py-3.5 px-4 text-center font-bold text-[#7A5E12]"><span class="badge-l3 px-2.5 py-1 rounded-md shadow-2xs">Keep</span></td>
                <td class="py-3.5 px-4 text-[#2E2827] leading-relaxed">符合職位門檻，展現穩定的工作交付。</td>
                <td class="py-3.5 px-4 text-[#4A433E] leading-relaxed">確認穩定度，指出下一階可以再往前的地方，維持既有節奏並選一到兩項深化。</td>
              </tr>
              <tr class="bg-[#FCE5CD]/35 hover:bg-[#FCE5CD]/50 transition">
                <td class="py-3.5 px-4 text-center font-bold text-[#8C4B1E]">L2</td>
                <td class="py-3.5 px-4 text-center font-bold text-[#8C4B1E]"><span class="badge-l2 px-2.5 py-1 rounded-md shadow-2xs">Grow</span></td>
                <td class="py-3.5 px-4 text-[#2E2827] leading-relaxed">部分符合，部分能力/行為仍在建立階段。</td>
                <td class="py-3.5 px-4 text-[#4A433E] leading-relaxed">說明落差在哪，聚焦一項具體可練習的行為，納入下一期 IDP 設定可觀察的行為指標。</td>
              </tr>
              <tr class="bg-[#F4CCCC]/35 hover:bg-[#F4CCCC]/50 transition">
                <td class="py-3.5 px-4 text-center font-bold text-[#7A363A]">L1</td>
                <td class="py-3.5 px-4 text-center font-bold text-[#7A363A]"><span class="badge-l1 px-2.5 py-1 rounded-md shadow-2xs">Start</span></td>
                <td class="py-3.5 px-4 text-[#2E2827] leading-relaxed">目前的具體事證還看不到這項職能，或事證與職能要求落差明顯。</td>
                <td class="py-3.5 px-4 text-[#4A433E] leading-relaxed">明確對齊職位基本門檻與要求，提供即時支援與改進行動引導。</td>
              </tr>
            </tbody>
          </table>
          <p class="text-xs text-[#8C837C] mt-3.5 italic">註：本落點標準供主管評核與回饋時參照，不對員工直接公布分數與總分，只回饋落點方向與成長指標。</p>
        </div>
      </div>

      <!-- TEAM INFO BANNER -->
      <div id="self-team-banner" class="bg-[#FCE5CD]/40 border border-[#F3D1B0] rounded-2xl p-5 sm:p-6 flex items-center justify-between">
        <div class="flex items-center gap-3.5">
          <div class="p-3 bg-[#F4CCCC] text-[#3E2426] rounded-xl shadow-2xs">
            <i data-lucide="users" class="w-5 h-5"></i>
          </div>
          <div>
            <h2 id="self-team-title" class="text-base sm:text-lg font-bold text-[#3E2426] font-serif-tc">張希慈 的部屬自評列表</h2>
            <p id="self-team-subtitle" class="text-xs sm:text-sm text-[#7A4822] mt-0.5">涵蓋部屬：何維安、陳泳璇、張芳媐、姚品瑄、胡喻翔</p>
          </div>
        </div>
        <div class="text-xs sm:text-sm px-4 py-2 bg-white/90 rounded-full text-[#783E16] font-bold border border-[#F3D1B0] shadow-2xs" id="self-completion-status">
          已填答 2 / 應填 5 人
        </div>
      </div>

      <!-- MEMBER SELF EVALUATION CARDS CONTAINER (CHUNK BY CHUNK) -->
      <div id="self-eval-cards-container" class="space-y-9">
        <!-- Injected via JavaScript -->
      </div>
    </section>

    <!-- ======================================================== -->
    <!-- TAB 2: 評主管 -->
    <!-- ======================================================== -->
    <section id="tab-section-supervisor" class="tab-content hidden space-y-7">
      <div class="bg-[#FFFDF9] rounded-2xl p-5 sm:p-6 border border-[#E8E2D8] soft-card-shadow flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div class="flex items-center gap-2.5 flex-wrap">
          <span class="text-xs font-bold text-[#8C837C] mr-1 flex items-center gap-1.5 uppercase tracking-wider">
            <i data-lucide="user" class="w-3.5 h-3.5 text-[#557A61]"></i> 受評主管：
          </span>
          <button onclick="filterSupervisor('ALL')" id="sup-filter-btn-ALL" class="sup-filter-btn px-4.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold bg-[#557A61] text-white shadow-xs transition">
            全部主管 (4筆)
          </button>
          <button onclick="filterSupervisor('張希慈')" id="sup-filter-btn-張希慈" class="sup-filter-btn px-4.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
            張希慈 (3筆)
          </button>
          <button onclick="filterSupervisor('何維安')" id="sup-filter-btn-何維安" class="sup-filter-btn px-4.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
            何維安 (1筆)
          </button>
        </div>

        <button onclick="exportFullWorkbookClientSide()" class="inline-flex items-center gap-2 px-4.5 py-2.5 text-xs sm:text-sm font-semibold rounded-xl bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition border border-[#E0D7CA]">
          <i data-lucide="download" class="w-4 h-4 text-[#557A61]"></i> 下載評主管總表 (XLSX)
        </button>
      </div>

      <!-- CHARTS & SUMMARY CARDS -->
      <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <!-- Radar Chart for Culture -->
        <div class="bg-[#FFFDF9] p-6 rounded-2xl border border-[#E8E2D8] soft-card-shadow">
          <h3 class="text-sm sm:text-base font-bold text-[#2E2827] mb-3 flex items-center gap-2 font-serif-tc">
            <i data-lucide="compass" class="w-4 h-4 text-[#557A61]"></i> 四大文化實踐維度
          </h3>
          <div class="h-68 flex items-center justify-center">
            <canvas id="supervisorRadarChart"></canvas>
          </div>
        </div>

        <!-- Bar Chart for Management Items -->
        <div class="bg-[#FFFDF9] p-6 rounded-2xl border border-[#E8E2D8] soft-card-shadow lg:col-span-2">
          <h3 class="text-sm sm:text-base font-bold text-[#2E2827] mb-3 flex items-center gap-2 font-serif-tc">
            <i data-lucide="bar-chart-3" class="w-4 h-4 text-[#557A61]"></i> 管理能力各題平均得分 (滿分10分)
          </h3>
          <div class="h-68">
            <canvas id="supervisorBarChart"></canvas>
          </div>
        </div>
      </div>

      <!-- DETAILED FEEDBACK LIST -->
      <div id="supervisor-feedback-list" class="space-y-6">
        <!-- Injected via JavaScript -->
      </div>
    </section>

    <!-- ======================================================== -->
    <!-- TAB 3: 評同事 -->
    <!-- ======================================================== -->
    <section id="tab-section-peer" class="tab-content hidden space-y-7">
      <div class="bg-[#FFFDF9] rounded-2xl p-5 sm:p-6 border border-[#E8E2D8] soft-card-shadow flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div class="flex items-center gap-2.5 flex-wrap" id="peer-pills-container">
          <span class="text-xs font-bold text-[#8C837C] mr-1 flex items-center gap-1.5 uppercase tracking-wider">
            <i data-lucide="user" class="w-3.5 h-3.5 text-[#557A61]"></i> 受評同事：
          </span>
          <!-- Dynamic Peer Buttons Injected -->
        </div>

        <button onclick="exportFullWorkbookClientSide()" class="inline-flex items-center gap-2 px-4.5 py-2.5 text-xs sm:text-sm font-semibold rounded-xl bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition border border-[#E0D7CA]">
          <i data-lucide="download" class="w-4 h-4 text-[#557A61]"></i> 下載評同事總表 (XLSX)
        </button>
      </div>

      <!-- CHARTS & SUMMARY CARDS -->
      <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <!-- Radar Chart for Culture -->
        <div class="bg-[#FFFDF9] p-6 rounded-2xl border border-[#E8E2D8] soft-card-shadow">
          <h3 class="text-sm sm:text-base font-bold text-[#2E2827] mb-3 flex items-center gap-2 font-serif-tc">
            <i data-lucide="compass" class="w-4 h-4 text-[#557A61]"></i> 文化實踐維度表現
          </h3>
          <div class="h-68 flex items-center justify-center">
            <canvas id="peerRadarChart"></canvas>
          </div>
        </div>

        <!-- Bar Chart for Peer Items -->
        <div class="bg-[#FFFDF9] p-6 rounded-2xl border border-[#E8E2D8] soft-card-shadow lg:col-span-2">
          <h3 class="text-sm sm:text-base font-bold text-[#2E2827] mb-3 flex items-center gap-2 font-serif-tc">
            <i data-lucide="bar-chart-3" class="w-4 h-4 text-[#557A61]"></i> 協作、當責與溝通評分 (滿分10分)
          </h3>
          <div class="h-68">
            <canvas id="peerBarChart"></canvas>
          </div>
        </div>
      </div>

      <!-- DETAILED PEER FEEDBACK LIST -->
      <div id="peer-feedback-list" class="space-y-6">
        <!-- Injected via JavaScript -->
      </div>
    </section>

    <!-- ======================================================== -->
    <!-- TAB 4: 個人 360 綜合雷達與對照報告 -->
    <!-- ======================================================== -->
    <section id="tab-section-overview360" class="tab-content hidden space-y-7">
      <div class="bg-[#FFFDF9] rounded-2xl p-5 sm:p-6 border border-[#E8E2D8] soft-card-shadow flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div class="flex items-center gap-2.5 flex-wrap" id="overview-member-pills">
          <span class="text-xs font-bold text-[#8C837C] mr-1 flex items-center gap-1.5 uppercase tracking-wider">
            <i data-lucide="sparkles" class="w-3.5 h-3.5 text-[#557A61]"></i> 成員報告：
          </span>
          <!-- Dynamic Member Buttons -->
        </div>
      </div>

      <div id="overview-report-container" class="space-y-7">
        <!-- Dynamic Personal 360 Profile -->
      </div>
    </section>

  </main>

  <!-- TOAST NOTIFICATION -->
  <div id="toast" class="fixed bottom-6 right-6 z-50 transform transition-all duration-300 opacity-0 translate-y-4 pointer-events-none bg-[#2E2827] text-white px-5 py-4 rounded-2xl shadow-xl flex items-center gap-3 text-xs sm:text-sm border border-[#4A433E]">
    <div id="toast-icon" class="p-1 rounded-lg bg-[#557A61] text-white">
      <i data-lucide="check" class="w-4 h-4"></i>
    </div>
    <span id="toast-msg">操作成功</span>
  </div>

  <footer class="bg-[#FFFDF9] border-t border-[#E8E2D8] py-7 text-center text-xs text-[#7A726D] mt-auto">
    好好星球文化基金會 360 年中成長評估系統 · 採用柔和色彩美學與模組化 Chunk 排版
  </footer>

  <script>
    // Embedded Evaluation Dataset
    let RAW_DATA = """ + data_json_str + """;

    // Supervisor Teams Mapping
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

    function toggleRatingGuide() {
      const content = document.getElementById('rating-guide-content');
      const icon = document.getElementById('guide-toggle-icon');
      const text = document.getElementById('guide-toggle-text');
      content.classList.toggle('hidden');
      if (content.classList.contains('hidden')) {
        icon.classList.remove('rotate-180');
        text.innerText = "展開評分標準說明";
      } else {
        icon.classList.add('rotate-180');
        text.innerText = "收合說明";
      }
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
        btn.classList.remove('border-[#557A61]', 'text-[#2D5239]', 'bg-[#E4ECD3]/30', 'font-semibold');
        btn.classList.add('border-transparent', 'text-[#6E6662]', 'font-medium');
      });
      const activeBtn = document.getElementById('tab-btn-' + tabId);
      activeBtn.classList.add('border-[#557A61]', 'text-[#2D5239]', 'bg-[#E4ECD3]/30', 'font-semibold');
      activeBtn.classList.remove('border-transparent', 'text-[#6E6662]');

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
    // CLIENT-SIDE CSV PARSING
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

            const entry = { timestamp, email, target, relation, job_role };

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
                  competencies.push({ title, answer: (row[idx] || "").trim() });
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
          showToast(`成功載入最新 CSV！共更新 ${newEntries.length} 筆填答紀錄。`);
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
    // TAB 1: SELF EVALUATION (CHUNK BY CHUNK WITH COMFORTABLE CONSISTENT PADDING)
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
        <button onclick="exportSupervisorExcelClientSide('${currentSelfSupervisor}')" class="inline-flex items-center gap-2 px-4.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold bg-[#557A61] text-white hover:bg-[#466551] transition shadow-xs">
          <i data-lucide="file-spreadsheet" class="w-4 h-4"></i>
          下載本組主管專用 XLSX 表格 (Chunk 格式)
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
            <!-- CHUNK BLOCK FOR MEMBER -->
            <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] soft-card-shadow overflow-hidden">
              
              <!-- 1. CHUNK TOP HEADER BANNER (#F4CCCC) -->
              <div class="chunk-header-banner px-6 py-5 flex flex-col md:flex-row md:items-center justify-between gap-3.5">
                <div class="flex items-center gap-4">
                  <div class="w-12 h-12 rounded-2xl bg-[#3E2426] text-white font-bold flex items-center justify-center text-lg font-serif-tc shadow-2xs">
                    ${memName.slice(0, 1)}
                  </div>
                  <div>
                    <div class="flex items-center gap-2.5">
                      <h3 class="text-base sm:text-lg font-bold text-[#3E2426] font-serif-tc">${memName}</h3>
                      <span class="px-3 py-1 text-xs font-semibold rounded-lg bg-[#FFFDF9] text-[#4A2E1C] border border-[#EED1B4]">
                        ${se.job_role || '自評'}
                      </span>
                      <span class="px-2.5 py-0.5 text-xs font-medium rounded-full bg-[#E4ECD3] text-[#2D5239] border border-[#CDE0BC] flex items-center gap-1">
                        <i data-lucide="check-circle" class="w-3.5 h-3.5"></i> 已填寫自評
                      </span>
                    </div>
                    <p class="text-xs text-[#7A4822] mt-1">${entry.email} · 填答時間：${entry.timestamp}</p>
                  </div>
                </div>
              </div>

              <div class="p-6 sm:p-8 space-y-7">
                
                <!-- 2. 工作特質盤點 CHUNK -->
                <div>
                  <div class="text-xs sm:text-sm font-bold text-[#8C837C] uppercase tracking-wider mb-3.5 flex items-center gap-2 font-serif-tc">
                    <i data-lucide="tag" class="w-4 h-4 text-[#557A61]"></i> 一、工作特質盤點
                  </div>
                  <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                    <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                      <div class="text-xs sm:text-sm font-bold text-[#2D5239] flex items-center gap-2">
                        <i data-lucide="shield-check" class="w-4 h-4 text-[#557A61]"></i> 最穩定、最具代表性 Top 3
                      </div>
                      <div class="flex flex-wrap gap-2.5 pt-1">
                        ${se.top3_stable.map(t => `<span class="badge-stable px-3.5 py-1.5 text-xs sm:text-sm font-semibold rounded-xl shadow-2xs">${t}</span>`).join('')}
                      </div>
                    </div>
                    <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                      <div class="text-xs sm:text-sm font-bold text-[#783E16] flex items-center gap-2">
                        <i data-lucide="trending-up" class="w-4 h-4 text-[#C27D38]"></i> 目前在練習 / 期望發展 3 項
                      </div>
                      <div class="flex flex-wrap gap-2.5 pt-1">
                        ${se.top3_practice.map(t => `<span class="badge-practice px-3.5 py-1.5 text-xs sm:text-sm font-semibold rounded-xl shadow-2xs">${t}</span>`).join('')}
                      </div>
                    </div>
                  </div>
                </div>

                <!-- 3. 四大文化實踐 CHUNK (STAR) -->
                <div>
                  <div class="text-xs sm:text-sm font-bold text-[#8C837C] uppercase tracking-wider mb-3.5 flex items-center gap-2 font-serif-tc">
                    <i data-lucide="heart" class="w-4 h-4 text-[#557A61]"></i> 二、四大文化實踐實例 (STAR 敘述)
                  </div>
                  <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                    <!-- 信任 -->
                    <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                      <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2">
                        <span class="w-2.5 h-2.5 rounded-full bg-[#557A61]"></span> 信任 (Trust)
                      </div>
                      <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs">
                        <p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.values['信任'] || '（未填寫）'}</p>
                      </div>
                    </div>
                    <!-- 多元 -->
                    <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                      <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2">
                        <span class="w-2.5 h-2.5 rounded-full bg-[#557A61]"></span> 多元 (Diversity)
                      </div>
                      <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs">
                        <p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.values['多元'] || '（未填寫）'}</p>
                      </div>
                    </div>
                    <!-- 實驗 -->
                    <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                      <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2">
                        <span class="w-2.5 h-2.5 rounded-full bg-[#557A61]"></span> 實驗 (Experiment)
                      </div>
                      <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs">
                        <p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.values['實驗'] || '（未填寫）'}</p>
                      </div>
                    </div>
                    <!-- 可持續 -->
                    <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                      <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2">
                        <span class="w-2.5 h-2.5 rounded-full bg-[#557A61]"></span> 可持續 (Sustainability)
                      </div>
                      <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs">
                        <p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.values['可持續'] || '（未填寫）'}</p>
                      </div>
                    </div>
                  </div>
                </div>

                <!-- 4. 職位專屬職能展現 CHUNK (MATCHING CONSISTENT 2-COLUMN GRID & PADDING) -->
                ${se.competencies && se.competencies.length > 0 ? `
                  <div>
                    <div class="text-xs sm:text-sm font-bold text-[#8C837C] uppercase tracking-wider mb-3.5 flex items-center gap-2 font-serif-tc">
                      <i data-lucide="briefcase" class="w-4 h-4 text-[#557A61]"></i> 三、${se.job_role} 專屬職能展現實例
                    </div>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                      ${se.competencies.map((c, cIdx) => `
                        <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                          <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2 font-serif-tc">
                            <span class="w-2.5 h-2.5 rounded-full bg-[#557A61]"></span> ${c.title}
                          </div>
                          <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs">
                            <p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed whitespace-pre-line">${c.answer || '（未填寫）'}</p>
                          </div>
                        </div>
                      `).join('')}
                    </div>
                  </div>
                ` : ''}

                <!-- 5. 組織卡關點與未來價值展望 CHUNK -->
                ${se.reflection && Object.keys(se.reflection).length > 0 ? `
                  <div>
                    <div class="text-xs sm:text-sm font-bold text-[#8C837C] uppercase tracking-wider mb-3.5 flex items-center gap-2 font-serif-tc">
                      <i data-lucide="help-circle" class="w-4 h-4 text-[#557A61]"></i> 四、組織卡關點反思與未來價值展望
                    </div>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                      ${Object.entries(se.reflection).map(([k, v], rIdx) => `
                        <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                          <div class="text-xs sm:text-sm font-bold text-[#783E16] flex items-center gap-2 font-serif-tc">
                            <i data-lucide="sparkle" class="w-3.5 h-3.5 text-[#C27D38]"></i> ${k}
                          </div>
                          <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs">
                            <p class="text-xs sm:text-sm text-[#4A433E] leading-relaxed">${v || '（未填寫）'}</p>
                          </div>
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
                <div class="w-12 h-12 rounded-2xl bg-[#F2EEE6] text-[#8C837C] font-bold flex items-center justify-center text-lg font-serif-tc">
                  ${memName.slice(0, 1)}
                </div>
                <div>
                  <div class="flex items-center gap-2.5">
                    <h3 class="text-base sm:text-lg font-bold text-[#2E2827] font-serif-tc">${memName}</h3>
                    <span class="px-3 py-1 text-xs font-medium rounded-full bg-[#FFF4CD] text-[#7A5E12] border border-[#FCE299] flex items-center gap-1.5">
                      <i data-lucide="clock" class="w-3.5 h-3.5"></i> 尚未收到自評資料
                    </span>
                  </div>
                  <p class="text-xs sm:text-sm text-[#8C837C] mt-1">此成員尚未於表單中送出自評紀錄，收到後上傳新 CSV 即可同步更新。</p>
                </div>
              </div>
              <div class="text-xs font-semibold text-[#8C837C] px-3.5 py-1.5 bg-[#F2EEE6] rounded-xl">待填寫</div>
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
            backgroundColor: 'rgba(85, 122, 97, 0.2)',
            borderColor: '#557A61',
            pointBackgroundColor: '#557A61'
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
            backgroundColor: '#557A61',
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
          <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] p-6 sm:p-8 soft-card-shadow space-y-6">
            <div class="flex items-center justify-between border-b border-[#E8E2D8] pb-4">
              <div class="flex items-center gap-3">
                <div class="w-10 h-10 rounded-xl bg-[#557A61] text-white font-bold flex items-center justify-center text-sm font-serif-tc shadow-2xs">
                  ${e.target.slice(0, 1)}
                </div>
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

            <!-- 2-COLUMN CONSISTENT PADDING GRID -->
            <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
              <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                <div class="text-xs sm:text-sm font-bold text-[#557A61] flex items-center gap-2">
                  <i data-lucide="compass" class="w-4 h-4"></i> Q20. 願景使命理解之引導
                </div>
                <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs">
                  <p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.q20_vision_mission || '（無）'}</p>
                </div>
              </div>

              <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                <div class="text-xs sm:text-sm font-bold text-[#783E16] flex items-center gap-2">
                  <i data-lucide="trending-up" class="w-4 h-4 text-[#C27D38]"></i> Q21. 管理與文化精神建議
                </div>
                <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs">
                  <p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.q21_improvement_advice || '（無）'}</p>
                </div>
              </div>

              <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2">
                  <i data-lucide="message-square" class="w-4 h-4 text-[#557A61]"></i> Q22. 其他補充評價
                </div>
                <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs">
                  <p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${se.q22_other_comments || '（無）'}</p>
                </div>
              </div>

              <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                <div class="text-xs sm:text-sm font-bold text-[#7A363A] flex items-center gap-2">
                  <i data-lucide="award" class="w-4 h-4 text-[#D48B7B]"></i> Q23. 肯定與感謝詞（好好星光大賞）
                </div>
                <div class="bg-white rounded-xl p-4 my-2 border border-[#F4CCCC] shadow-2xs">
                  <p class="text-xs sm:text-sm text-[#592629] font-medium leading-relaxed">${se.q23_starlight_thanks || '（無）'}</p>
                </div>
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
        <span class="text-xs font-bold text-[#8C837C] mr-1 flex items-center gap-1.5 uppercase tracking-wider">
          <i data-lucide="user" class="w-3.5 h-3.5 text-[#557A61]"></i> 受評同事：
        </span>
        <button onclick="filterPeer('ALL')" id="peer-btn-ALL" class="peer-pill-btn px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold bg-[#557A61] text-white shadow-xs transition">
          全部 (${RAW_DATA.filter(e => e.relation === "同事").length}筆)
        </button>
      `;
      peers.forEach(p => {
        const count = RAW_DATA.filter(e => e.relation === "同事" && e.target === p).length;
        html += `
          <button onclick="filterPeer('${p}')" id="peer-btn-${p}" class="peer-pill-btn px-4 py-2.5 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
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
            backgroundColor: 'rgba(85, 122, 97, 0.2)',
            borderColor: '#557A61',
            pointBackgroundColor: '#557A61'
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
            backgroundColor: '#557A61',
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
          <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] p-6 sm:p-8 soft-card-shadow space-y-6">
            <div class="flex items-center justify-between border-b border-[#E8E2D8] pb-4">
              <div class="flex items-center gap-3">
                <div class="w-10 h-10 rounded-xl bg-[#557A61] text-white font-bold flex items-center justify-center text-sm font-serif-tc shadow-2xs">
                  ${e.target.slice(0, 1)}
                </div>
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

            <!-- 3-COLUMN CONSISTENT PADDING GRID -->
            <div class="grid grid-cols-1 md:grid-cols-3 gap-5">
              <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2">
                  <i data-lucide="trending-up" class="w-4 h-4 text-[#557A61]"></i> Q37. 工作與文化提升建議
                </div>
                <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs">
                  <p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${pe.q37_improvement_advice || '（無）'}</p>
                </div>
              </div>

              <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                <div class="text-xs sm:text-sm font-bold text-[#2E2827] flex items-center gap-2">
                  <i data-lucide="message-square" class="w-4 h-4 text-[#557A61]"></i> Q38. 其他補充評價
                </div>
                <div class="bg-white rounded-xl p-4 my-2 border border-[#E8E2D8] shadow-2xs">
                  <p class="text-xs sm:text-sm text-[#2E2827] leading-relaxed">${pe.q38_other_comments || '（無）'}</p>
                </div>
              </div>

              <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                <div class="text-xs sm:text-sm font-bold text-[#7A363A] flex items-center gap-2">
                  <i data-lucide="award" class="w-4 h-4 text-[#D48B7B]"></i> Q39. 肯定與感謝詞
                </div>
                <div class="bg-white rounded-xl p-4 my-2 border border-[#F4CCCC] shadow-2xs">
                  <p class="text-xs sm:text-sm text-[#592629] font-medium leading-relaxed">${pe.q39_starlight_thanks || '（無）'}</p>
                </div>
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
      allTargets.forEach(t => {
        if (!members.includes(t)) members.push(t);
      });

      const container = document.getElementById('overview-member-pills');
      let html = `
        <span class="text-xs font-bold text-[#8C837C] mr-1 flex items-center gap-1.5 uppercase tracking-wider">
          <i data-lucide="sparkles" class="w-3.5 h-3.5 text-[#557A61]"></i> 成員報告：
        </span>
      `;
      members.forEach(m => {
        html += `
          <button onclick="selectOverviewMember('${m}')" id="overview-btn-${m}" class="overview-pill-btn px-4 py-2 rounded-xl text-xs sm:text-sm font-medium bg-[#F2EEE6] text-[#4A433E] hover:bg-[#EBE4D8] transition">
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
        btn.classList.remove('bg-[#557A61]', 'text-white', 'font-semibold');
        btn.classList.add('bg-[#F2EEE6]', 'text-[#4A433E]', 'font-medium');
      });
      const activeBtn = document.getElementById('overview-btn-' + name);
      if (activeBtn) {
        activeBtn.classList.add('bg-[#557A61]', 'text-white', 'font-semibold');
        activeBtn.classList.remove('bg-[#F2EEE6]', 'text-[#4A433E]', 'font-medium');
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
        <div class="bg-[#FFFDF9] rounded-2xl border border-[#E8E2D8] p-6 sm:p-8 soft-card-shadow space-y-7">
          <div class="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-[#E8E2D8] pb-5">
            <div class="flex items-center gap-4">
              <div class="w-14 h-14 rounded-2xl bg-[#557A61] text-white font-bold flex items-center justify-center text-xl font-serif-tc shadow-xs">
                ${name.slice(0, 1)}
              </div>
              <div>
                <div class="flex items-center gap-2.5">
                  <h2 class="text-xl sm:text-2xl font-bold text-[#2E2827] font-serif-tc">${name} 360 綜合評估報告</h2>
                  ${se ? `<span class="px-3 py-1 text-xs font-semibold rounded-lg bg-[#FCE5CD] text-[#783E16] border border-[#F3D1B0]">${se.job_role}</span>` : ''}
                </div>
                <p class="text-xs text-[#8C837C] mt-1">包含自評 (${selfEntry ? '1' : '0'}份)、同儕回饋 (${peerEntries.length}份)、主管回饋 (${supEntries.length}份)</p>
              </div>
            </div>
            <div class="flex items-center gap-2.5 text-xs sm:text-sm flex-wrap">
              <span class="px-4 py-2 rounded-xl bg-[#E4ECD3] text-[#2D5239] font-bold">同事 NPS：${peerAvg('q36_nps_recommend')} 分</span>
              ${supEntries.length ? `<span class="px-4 py-2 rounded-xl bg-[#FCE5CD] text-[#8C4B1E] font-bold">主管 NPS：${supAvg('q18_nps_recommend')} 分</span>` : ''}
            </div>
          </div>

          <!-- Self Traits Quick View -->
          ${se ? `
            <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
              <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                <div class="text-xs sm:text-sm font-bold text-[#2D5239] flex items-center gap-2">
                  <i data-lucide="shield-check" class="w-4 h-4 text-[#557A61]"></i> 自評最穩定代表特質 Top 3
                </div>
                <div class="flex flex-wrap gap-2.5 pt-1">
                  ${se.top3_stable.map(t => `<span class="badge-stable px-3.5 py-1.5 text-xs sm:text-sm font-semibold rounded-xl">${t}</span>`).join('')}
                </div>
              </div>
              <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 sm:p-6 space-y-3">
                <div class="text-xs sm:text-sm font-bold text-[#783E16] flex items-center gap-2">
                  <i data-lucide="trending-up" class="w-4 h-4 text-[#C27D38]"></i> 自評目前練習中 / 期望發展 3 項
                </div>
                <div class="flex flex-wrap gap-2.5 pt-1">
                  ${se.top3_practice.map(t => `<span class="badge-practice px-3.5 py-1.5 text-xs sm:text-sm font-semibold rounded-xl">${t}</span>`).join('')}
                </div>
              </div>
            </div>
          ` : `
            <div class="bg-[#FFF4CD]/40 border border-[#FCE299] rounded-2xl p-5 sm:p-6 text-xs sm:text-sm text-[#7A5E12]">
              提示：${name} 尚未在表單中提交自評資料。下方為夥伴與主管之評價。
            </div>
          `}

          <!-- Starlight感謝詞 -->
          <div>
            <h3 class="text-sm sm:text-base font-bold text-[#7A363A] mb-3.5 flex items-center gap-2.5 font-serif-tc">
              <i data-lucide="sparkles" class="w-4 h-4 text-[#D48B7B]"></i> 好好星光大賞：夥伴們給 ${name} 的感謝與肯定
            </h3>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
              ${[...supEntries, ...peerEntries].filter(e => (e.supervisor_eval?.q23_starlight_thanks || e.peer_eval?.q39_starlight_thanks)).map(e => {
                const text = e.supervisor_eval?.q23_starlight_thanks || e.peer_eval?.q39_starlight_thanks;
                return `
                  <div class="bg-[#F9F6EE] border border-[#E8E2D8] rounded-2xl p-5 text-xs sm:text-sm text-[#592629] font-medium flex items-start gap-3.5">
                    <i data-lucide="heart" class="w-4 h-4 text-[#D48B7B] shrink-0 mt-0.5"></i>
                    <div>
                      <p class="leading-relaxed">"${text}"</p>
                      <span class="text-xs text-[#8C5558] mt-2 block">— 來自 ${e.relation} (${e.email.split('@')[0]})</span>
                    </div>
                  </div>
                `;
              }).join('') || '<p class="text-xs text-[#8C837C]">目前尚無星光感謝詞紀錄</p>'}
            </div>
          </div>
        </div>
      `;
      lucide.createIcons();
    }

    // ==========================================
    // CLIENT-SIDE XLSX CHUNK EXPORT (MATCHING UPDATED EXCEL STRUCTURE)
    // ==========================================
    function exportSupervisorExcelClientSide(supName) {
      const wb = XLSX.utils.book_new();
      const memberNames = SUPERVISOR_TEAMS[supName] || [];

      const rows = [];
      rows.push([`好好星球文化基金會 360 年中成長評估 - 【${supName}】部屬自評專用表`]);
      rows.push([`職能評分標準與面向分數落點說明（不對員工公布分數與總分，只回饋落點）`]);
      rows.push(["等級 (Level)", "落點名稱", "定義說明", "", "回饋語氣與後續動作"]);
      rows.push(["L5", "Amazing!", "遠超職位期待，表現為團隊之標竿與典範。", "", "讚賞其突出貢獻，探討經驗複製機制，在成長對話中轉為「可以帶其他人一起做」的方向。"]);
      rows.push(["L4", "Good", "優於職位期待，持續展現高標準成果。", "", "肯定並具體指出是哪些行為讓它超出標準，設定具挑戰性的下一步目標。"]);
      rows.push(["L3", "Keep", "符合職位門檻，展現穩定的工作交付。", "", "確認穩定度，指出下一階可以再往前的地方，維持既有節奏並選一到兩項深化。"]);
      rows.push(["L2", "Grow", "部分符合，部分能力/行為仍在建立階段。", "", "說明落差在哪，聚焦一項具體可練習的行為，納入下一期 IDP 設定可觀察的行為指標。"]);
      rows.push(["L1", "Start", "目前的具事證還看不到這項職能，或事證與職能要求落差明顯。", "", "明確對齊職位基本門檻與要求，提供即時支援與改進行動引導。"]);
      rows.push([]);

      const merges = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } },
        { s: { r: 2, c: 2 }, e: { r: 2, c: 3 } },
        { s: { r: 3, c: 2 }, e: { r: 3, c: 3 } },
        { s: { r: 4, c: 2 }, e: { r: 4, c: 3 } },
        { s: { r: 5, c: 2 }, e: { r: 5, c: 3 } },
        { s: { r: 6, c: 2 }, e: { r: 6, c: 3 } },
        { s: { r: 7, c: 2 }, e: { r: 7, c: 3 } }
      ];

      memberNames.forEach(memName => {
        const memEntry = RAW_DATA.find(e => e.target === memName && e.relation === "自評");
        const role = memEntry && memEntry.self_eval ? memEntry.self_eval.job_role : "尚未填寫自評";
        const timeStr = memEntry ? ` ｜ 填答時間：${memEntry.timestamp}` : "";

        const chunkHeaderRow = rows.length;
        rows.push([`部屬姓名：${memName}    ｜    職位：${role}${timeStr}`, "", "", "", ""]);
        merges.push({ s: { r: chunkHeaderRow, c: 0 }, e: { r: chunkHeaderRow, c: 4 } });

        if (!memEntry || !memEntry.self_eval) {
          const unsubmittedRow = rows.length;
          rows.push(["自評狀態", "目前表單中尚未收到該成員的自我評估回覆紀錄。", "", "", ""]);
          merges.push({ s: { r: unsubmittedRow, c: 1 }, e: { r: unsubmittedRow, c: 4 } });
          rows.push([]);
          rows.push([]);
          return;
        }

        const se = memEntry.self_eval;
        
        // 1. 特質盤點
        const traitR1 = rows.length;
        rows.push(["工作特質盤點", "最穩定代表 Top 3", se.top3_stable.join("、") || "（無）", "", ""]);
        merges.push({ s: { r: traitR1, c: 2 }, e: { r: traitR1, c: 4 } });
        
        const traitR2 = rows.length;
        rows.push(["工作特質盤點", "練習中 / 期望發展", se.top3_practice.join("、") || "（無）", "", ""]);
        merges.push({ s: { r: traitR2, c: 2 }, e: { r: traitR2, c: 4 } });
        merges.push({ s: { r: traitR1, c: 0 }, e: { r: traitR2, c: 0 } });

        // 2. 四大文化實踐 (Merged D:E across all 4 rows for Culture Text Feedback; No Level)
        const cultBannerR = rows.length;
        rows.push(["四大文化實踐實例（STAR 敘述）", "", "", "", ""]);
        merges.push({ s: { r: cultBannerR, c: 0 }, e: { r: cultBannerR, c: 4 } });

        const cultHeadR = rows.length;
        rows.push(["評估面向", "四大文化定義與說明", "部屬自評實例 (STAR)", "", "組織文化整體主管回饋 (Feedback)"]);
        merges.push({ s: { r: cultHeadR, c: 3 }, e: { r: cultHeadR, c: 4 } });

        const cultStartR = rows.length;
        rows.push(["組織文化", "【信任】獨立行動與決策、主動協作、雙向溝通，在高彈性下給予彼此信任", se.values['信任'] || "（無）", "", ""]);
        rows.push(["組織文化", "【多元】尊重差異、多元工作方法、主動表達不同觀點與想法", se.values['多元'] || "（無）", "", ""]);
        rows.push(["組織文化", "【實驗】透過開放心態嘗試修正與反思，勇於檢討及給予回饋", se.values['實驗'] || "（無）", "", ""]);
        rows.push(["組織文化", "【可持續】內在韌性、自我照顧、彈性的人際與工作邊界", se.values['可持續'] || "（無）", "", ""]);
        merges.push({ s: { r: cultStartR, c: 0 }, e: { r: cultStartR + 3, c: 0 } });
        merges.push({ s: { r: cultStartR, c: 3 }, e: { r: cultStartR + 3, c: 4 } });

        // 3. 職能展現
        if (se.competencies && se.competencies.length > 0) {
          const compBannerR = rows.length;
          rows.push([`職位專屬職能展現實例【${se.job_role}】`, "", "", "", ""]);
          merges.push({ s: { r: compBannerR, c: 0 }, e: { r: compBannerR, c: 4 } });

          rows.push(["評估面向", "職能項目與題目定義", "部屬自評實例", "評分 (Lv.1-5選單)", "主管回饋 (Feedback)"]);
          const compStartR = rows.length;
          se.competencies.forEach(c => {
            rows.push(["專業職能", c.title, c.answer || "（無填寫）", "", ""]);
          });
          merges.push({ s: { r: compStartR, c: 0 }, e: { r: compStartR + se.competencies.length - 1, c: 0 } });
        }

        // Note: 卡關與展望 is removed from Excel as requested!
        rows.push([]);
        rows.push([]);
      });

      const ws = XLSX.utils.aoa_to_sheet(rows);
      ws['!merges'] = merges;
      ws['!cols'] = [{ wch: 16 }, { wch: 36 }, { wch: 68 }, { wch: 15 }, { wch: 34 }];
      XLSX.utils.book_append_sheet(wb, ws, "部屬自評彙整表");
      XLSX.writeFile(wb, `自評_${supName}主管專用_部屬自評彙整表.xlsx`);
      showToast(`已成功下載【${supName}】部屬自評 XLSX（Chunk 格式）！`);
    }

    function exportFullWorkbookClientSide() {
      const wb = XLSX.utils.book_new();

      // 1. Supervisor Team Sheets in Chunk format
      ["張希慈", "何維安", "姚品瑄", "張希慈_執行長"].forEach(supName => {
        const memberNames = SUPERVISOR_TEAMS[supName] || [];
        const rows = [];
        rows.push([`好好星球文化基金會 360 年中成長評估 - 【${supName}】部屬自評專用表`]);
        rows.push([`職能評分標準與面向分數落點說明（不對員工公布分數與總分，只回饋落點）`]);
        rows.push(["等級 (Level)", "落點名稱", "定義說明", "", "回饋語氣與後續動作"]);
        rows.push(["L5", "Amazing!", "遠超職位期待，表現為團隊之標竿與典範。", "", "讚賞其突出貢獻，探討經驗複製機制，在成長對話中轉為「可以帶其他人一起做」的方向。"]);
        rows.push(["L4", "Good", "優於職位期待，持續展現高標準成果。", "", "肯定並具體指出是哪些行為讓它超出標準，設定具挑戰性的下一步目標。"]);
        rows.push(["L3", "Keep", "符合職位門檻，展現穩定的工作交付。", "", "確認穩定度，指出下一階可以再往前的地方，維持既有節奏並選一到兩項深化。"]);
        rows.push(["L2", "Grow", "部分符合，部分能力/行為仍在建立階段。", "", "說明落差在哪，聚焦一項具體可練習的行為，納入下一期 IDP 設定可觀察的行為指標。"]);
        rows.push(["L1", "Start", "目前的具體事證還看不到這項職能，或事證與職能要求落差明顯。", "", "明確對齊職位基本門檻與要求，提供即時支援與改進行動引導。"]);
        rows.push([]);

        const merges = [
          { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
          { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } },
          { s: { r: 2, c: 2 }, e: { r: 2, c: 3 } },
          { s: { r: 3, c: 2 }, e: { r: 3, c: 3 } },
          { s: { r: 4, c: 2 }, e: { r: 4, c: 3 } },
          { s: { r: 5, c: 2 }, e: { r: 5, c: 3 } },
          { s: { r: 6, c: 2 }, e: { r: 6, c: 3 } },
          { s: { r: 7, c: 2 }, e: { r: 7, c: 3 } }
        ];

        memberNames.forEach(memName => {
          const memEntry = RAW_DATA.find(e => e.target === memName && e.relation === "自評");
          const role = memEntry && memEntry.self_eval ? memEntry.self_eval.job_role : "尚未填寫自評";
          const timeStr = memEntry ? ` ｜ 填答時間：${memEntry.timestamp}` : "";

          const chunkHeaderRow = rows.length;
          rows.push([`部屬姓名：${memName}    ｜    職位：${role}${timeStr}`, "", "", "", ""]);
          merges.push({ s: { r: chunkHeaderRow, c: 0 }, e: { r: chunkHeaderRow, c: 4 } });

          if (!memEntry || !memEntry.self_eval) {
            const unsubmittedRow = rows.length;
            rows.push(["自評狀態", "目前表單中尚未收到該成員的自我評估回覆紀錄。", "", "", ""]);
            merges.push({ s: { r: unsubmittedRow, c: 1 }, e: { r: unsubmittedRow, c: 4 } });
            rows.push([]);
            rows.push([]);
            return;
          }

          const se = memEntry.self_eval;
          const traitR1 = rows.length;
          rows.push(["工作特質盤點", "最穩定代表 Top 3", se.top3_stable.join("、") || "（無）", "", ""]);
          merges.push({ s: { r: traitR1, c: 2 }, e: { r: traitR1, c: 4 } });
          
          const traitR2 = rows.length;
          rows.push(["工作特質盤點", "練習中 / 期望發展", se.top3_practice.join("、") || "（無）", "", ""]);
          merges.push({ s: { r: traitR2, c: 2 }, e: { r: traitR2, c: 4 } });
          merges.push({ s: { r: traitR1, c: 0 }, e: { r: traitR2, c: 0 } });

          const cultBannerR = rows.length;
          rows.push(["四大文化實踐實例（STAR 敘述）", "", "", "", ""]);
          merges.push({ s: { r: cultBannerR, c: 0 }, e: { r: cultBannerR, c: 4 } });

          const cultHeadR = rows.length;
          rows.push(["評估面向", "四大文化定義與說明", "部屬自評實例 (STAR)", "", "組織文化整體主管回饋 (Feedback)"]);
          merges.push({ s: { r: cultHeadR, c: 3 }, e: { r: cultHeadR, c: 4 } });

          const cultStartR = rows.length;
          rows.push(["組織文化", "【信任】獨立行動與決策、主動協作、雙向溝通，在高彈性下給予彼此信任", se.values['信任'] || "（無）", "", ""]);
          rows.push(["組織文化", "【多元】尊重差異、多元工作方法、主動表達不同觀點與想法", se.values['多元'] || "（無）", "", ""]);
          rows.push(["組織文化", "【實驗】透過開放心態嘗試修正與反思，勇於檢討及給予回饋", se.values['實驗'] || "（無）", "", ""]);
          rows.push(["組織文化", "【可持續】內在韌性、自我照顧、彈性的人際與工作邊界", se.values['可持續'] || "（無）", "", ""]);
          merges.push({ s: { r: cultStartR, c: 0 }, e: { r: cultStartR + 3, c: 0 } });
          merges.push({ s: { r: cultStartR, c: 3 }, e: { r: cultStartR + 3, c: 4 } });

          if (se.competencies && se.competencies.length > 0) {
            const compBannerR = rows.length;
            rows.push([`職位專屬職能展現實例【${se.job_role}】`, "", "", "", ""]);
            merges.push({ s: { r: compBannerR, c: 0 }, e: { r: compBannerR, c: 4 } });

            rows.push(["評估面向", "職能項目與題目定義", "部屬自評實例", "評分 (Lv.1-5選單)", "主管回饋 (Feedback)"]);
            const compStartR = rows.length;
            se.competencies.forEach(c => {
              rows.push(["專業職能", c.title, c.answer || "（無填寫）", "", ""]);
            });
            merges.push({ s: { r: compStartR, c: 0 }, e: { r: compStartR + se.competencies.length - 1, c: 0 } });
          }

          rows.push([]);
          rows.push([]);
        });

        const ws = XLSX.utils.aoa_to_sheet(rows);
        ws['!merges'] = merges;
        ws['!cols'] = [{ wch: 16 }, { wch: 36 }, { wch: 68 }, { wch: 15 }, { wch: 34 }];
        const cleanTitle = supName === "張希慈_執行長" ? "自評_執行長" : `自評_${supName}`;
        XLSX.utils.book_append_sheet(wb, ws, cleanTitle);
      });

      // 2. Supervisor Evaluation Sheet
      const supRows = [];
      supRows.push(["好好星球文化基金會 360 年中成長評估 - 【評主管】回覆總表"]);
      supRows.push(["時間戳記", "填答者Email", "受評主管", "Q4.尋求協助", "Q5.具體引導", "Q6.改善幅度", "Q7.跨部門推動", "Q8.資源評估", "Q9.失誤回應", "Q10.肯定認可", "Q11.工作成效", "Q12.【信任】表達想法", "Q13.【多元】聆聽意見", "Q14.【實驗】嘗試創新", "Q15.【實驗】試錯空間", "Q16.【可持續】讚美肯定", "Q17.【可持續】尊重界線", "Q18.NPS推薦度", "Q19.滿意度", "Q20.願景使命引導(質化)", "Q21.管理提升建議(質化)", "Q22.其他補充評價(質化)", "Q23.肯定感謝詞(好好星光大賞)"]);
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
      wsSup['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 22 } }];
      wsSup['!cols'] = Array(19).fill({ wch: 15 }).concat([{ wch: 35 }, { wch: 35 }, { wch: 35 }, { wch: 35 }]);
      XLSX.utils.book_append_sheet(wb, wsSup, "評主管總表");

      // 3. Peer Evaluation Sheet
      const peerRows = [];
      peerRows.push(["好好星球文化基金會 360 年中成長評估 - 【評同事】回覆總表"]);
      peerRows.push(["時間戳記", "填答者Email", "受評同事", "Q24.合作狀況", "Q25.注重細節", "Q26.準時完成", "Q27.靈活調整", "Q28.追蹤承諾", "Q29.說明依據", "Q30.【多元】接受不同意見", "Q31.【多元】建設性觀點", "Q32.【實驗】開放調整", "Q33.【信任】分享經驗", "Q34.【可持續】讚美同事", "Q35.【可持續】尊重界線", "Q36.NPS推薦度", "Q37.提升建議(質化)", "Q38.其他評價(質化)", "Q39.肯定感謝詞(好好星光大賞)"]);
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
      wsPeer['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 18 } }];
      wsPeer['!cols'] = Array(16).fill({ wch: 15 }).concat([{ wch: 35 }, { wch: 35 }, { wch: 35 }]);
      XLSX.utils.book_append_sheet(wb, wsPeer, "評同事總表");

      XLSX.writeFile(wb, "好好星球_360年中成長評估_主管分流與完整彙整表.xlsx");
      showToast("已成功下載全組織完整 Excel 總表（Chunk 格式）！");
    }

    // Initialize on load
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

print("Updated index.html: strictly consistent 2-column padding across ALL cards, unified palette, and clean view without web rating menus!")
