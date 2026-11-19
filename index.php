<!DOCTYPE html>
<html lang="id">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width,initial-scale=1.0" />
    <title>DSS Calculator — Entropy, GRA, MEREC, MARCOS</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <link
      href="https://fonts.googleapis.com/css2?family=Inter+Tight:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap"
      rel="stylesheet"
    />
    <script>
      tailwind.config = {
        theme: {
          extend: {
            fontFamily: { sans: ["Inter Tight", "sans-serif"], mono: ["JetBrains Mono", "monospace"] },
          },
        },
      };
    </script>
    <style>
      * {
        box-sizing: border-box;
      }
      :root {
        --c1: #f2f1ec;
        --c2: #e3ece6;
        --c3: #d3e4da;
        --c4: #1e4d3b;
        --text: #101418;
        --muted: #5c646a;
        --border: #d2d0c6;
        --card: #fbfaf7;
      }
      body {
        font-family: "Inter Tight", "Helvetica Neue", Arial, sans-serif;
        background: var(--c1);
        color: var(--text);
        margin: 0;
        -webkit-font-smoothing: antialiased;
      }
      .glass {
        background: rgba(251, 250, 247, 0.94);
        backdrop-filter: blur(8px);
        border: 1px solid var(--border);
        box-shadow: 0 10px 26px rgba(16, 20, 24, 0.05);
      }
      .glow {
        box-shadow: 0 8px 22px rgba(30, 77, 59, 0.16);
      }
      .badge {
        font-family: "JetBrains Mono", monospace;
      }
      table {
        border-collapse: collapse;
        width: 100%;
      }
      th {
        background: var(--c3);
        padding: 8px 10px;
        font-size: 11px;
        font-weight: 700;
        white-space: nowrap;
        color: var(--text);
        letter-spacing: 0.01em;
      }
      td {
        padding: 7px 10px;
        font-size: 11px;
        border: 1px solid var(--border);
        color: var(--text);
      }
      th {
        border: 1px solid var(--border);
      }
      tr:nth-child(even) td {
        background: rgba(30, 77, 59, 0.035);
      }
      tr:hover td {
        background: var(--c2);
      }
      .fade-in {
        animation: fi 0.3s ease;
      }
      @keyframes fi {
        from {
          opacity: 0;
          transform: translateY(6px);
        }
        to {
          opacity: 1;
          transform: none;
        }
      }
      ::-webkit-scrollbar {
        width: 7px;
        height: 7px;
      }
      ::-webkit-scrollbar-track {
        background: var(--c1);
      }
      ::-webkit-scrollbar-thumb {
        background: #9bb3a6;
        border-radius: 8px;
      }
      input[type="number"]::-webkit-inner-spin-button {
        -webkit-appearance: none;
      }
      .method-tab {
        transition: all 0.2s;
        cursor: pointer;
        border-bottom: 2px solid transparent;
      }
      .method-tab.active {
        border-bottom-color: var(--c4);
        color: var(--text);
        font-weight: 700;
      }
      .tab-panel {
        display: none;
      }
      .tab-panel.active {
        display: block;
      }
      .pill {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        padding: 2px 8px;
        border-radius: 99px;
        font-size: 10px;
        font-family: "JetBrains Mono", monospace;
      }

      /* Override sisa kelas Tailwind/inline dark agar konsisten dengan palet baru */
      .bg-slate-900,
      .bg-slate-800,
      .bg-slate-700 {
        background-color: var(--card) !important;
      }
      .hover\:bg-slate-600:hover,
      .hover\:bg-slate-700\/50:hover {
        background-color: var(--c2) !important;
      }
      .text-slate-500,
      .text-slate-400,
      .text-slate-300,
      .text-slate-200 {
        color: var(--muted) !important;
      }
      .text-white {
        color: var(--text) !important;
      }
      .border-slate-700\/50,
      .border-slate-700,
      .border-slate-600 {
        border-color: var(--border) !important;
      }
      input,
      select,
      textarea {
        background: var(--card) !important;
        color: var(--text) !important;
        border-color: var(--border) !important;
      }
      input:focus,
      select:focus,
      textarea:focus {
        border-color: var(--c4) !important;
        box-shadow: 0 0 0 3px rgba(30, 77, 59, 0.14) !important;
      }
      button {
        color: var(--text) !important;
        border-radius: 7px;
      }
      [style*="background:#3B82F6"],
      [style*="background:#10B981"],
      [style*="background:#8B5CF6"],
      [style*="background:#6366F1"],
      [style*="background:linear-gradient"] {
        background: var(--c4) !important;
        color: #f2f1ec !important;
      }
      [style*="#60A5FA"],
      [style*="#93C5FD"],
      [style*="#34D399"],
      [style*="#A78BFA"],
      [style*="#FCD34D"],
      [style*="#94A3B8"],
      [style*="#EF4444"] {
        color: var(--text) !important;
      }
      [style*="#1E293B"],
      [style*="#334155"],
      [style*="#475569"] {
        border-color: var(--border) !important;
      }

      /* ---------- Editorial refresh ---------- */
      body {
        font-family: "Inter Tight", "Helvetica Neue", Arial, sans-serif;
        letter-spacing: -0.005em;
      }
      h1,
      h2,
      h3 {
        letter-spacing: -0.025em;
      }
      .glass {
        border-radius: 8px;
      }
      header.glass {
        background: #101418 !important;
        border-color: #252c30 !important;
        box-shadow: none !important;
      }
      header.glass h1 {
        color: #f2f1ec !important;
        font-size: 14px;
        letter-spacing: -0.02em;
      }
      header.glass p {
        color: #9ca7ad !important;
      }
      header.glass .text-white {
        color: #f2f1ec !important;
      }
      header.glass .glow {
        background: #1e4d3b !important;
        box-shadow: none !important;
      }
      header.glass button {
        background: #242b30 !important;
        border: 1px solid #3a444b !important;
        color: #e6ebe7 !important;
      }
      aside.glass {
        background: #fbfaf7 !important;
        border-color: #d6d3c9 !important;
        box-shadow: none !important;
        padding: 22px 16px;
      }
      aside .text-slate-500,
      aside .text-slate-400 {
        color: var(--muted) !important;
      }
      aside .text-white {
        color: var(--text) !important;
      }
      aside .bg-slate-800\/60 {
        background: #eef1ec !important;
        border-color: #d6ddd7 !important;
      }
      aside input {
        background: #ffffff !important;
        color: var(--text) !important;
      }
      main {
        padding: 28px !important;
      }
      .glass.rounded-xl,
      .glass.rounded-lg {
        border-radius: 8px !important;
      }
      .shadow-lg {
        box-shadow: none !important;
      }
      /* Buttons: flat, high-contrast, no gradients */
      button[onclick="calculateAll()"],
      button[onclick="exportResults()"],
      button[onclick="buildMatrix()"] {
        background: var(--c4) !important;
        color: #f2f1ec !important;
        border: 1px solid var(--c4) !important;
      }
      button[onclick="calculateAll()"] svg,
      button[onclick="exportResults()"] svg,
      button[onclick="buildMatrix()"] svg {
        color: #f2f1ec !important;
      }
      button[onclick="calculateAll()"]:hover,
      button[onclick="exportResults()"]:hover,
      button[onclick="buildMatrix()"]:hover {
        background: #296a52 !important;
        border-color: #296a52 !important;
      }
      button[onclick="fillSample()"],
      button[onclick="exportTemplate()"] {
        background: #ffffff !important;
        color: var(--text) !important;
        border: 1px solid var(--border) !important;
      }
      button[onclick="fillSample()"]:hover,
      button[onclick="exportTemplate()"]:hover {
        background: var(--c2) !important;
        border-color: #b9c8be !important;
      }
      /* Method tabs: sage active state */
      #btn-gra,
      #btn-merec_marcos,
      #btn-combined {
        border-radius: 7px !important;
      }
      /* Step and result accents */
      .pill {
        border-radius: 5px;
        letter-spacing: 0.02em;
      }
      .fade-in {
        animation-duration: 0.25s;
      }
      /* Table polish */
      th,
      td {
        border-color: #d9d6cc !important;
      }
      th {
        background: #e1ebe4 !important;
        color: #2b342e !important;
        font-size: 10.5px;
        letter-spacing: 0.02em;
      }
      table {
        border-radius: 6px;
        overflow: hidden;
      }
      /* Inputs */
      input[type="number"] {
        font-family: "JetBrains Mono", monospace !important;
        font-size: 12px !important;
      }
      /* Formula / note boxes keep a quiet paper tone */
      .badge {
        letter-spacing: 0.01em;
      }
      /* Remove the old blue glow on the logo mark */
      .glow {
        box-shadow: none;
      }
      /* Better focus visibility */
      button:focus-visible,
      a:focus-visible,
      input:focus-visible,
      select:focus-visible {
        outline: 2px solid var(--c4);
        outline-offset: 2px;
      }
      @media (max-width: 1024px) {
        main {
          padding: 18px !important;
        }
      }

      /* Tailwind blue remnants */
      .bg-blue-600\/20,
      .bg-blue-900\/60,
      .bg-blue-50 {
        background: var(--c3) !important;
      }
      .text-blue-300,
      .text-blue-400,
      .text-blue-500,
      .text-blue-600 {
        color: var(--c4) !important;
      }
      .border-blue-600\/40,
      .border-blue-700\/50,
      .border-blue-500,
      .border-blue-200 {
        border-color: #b9c8be !important;
      }
      /* Dynamic result colours generated in JS */
      [style*="background: #10b981"],
      [style*="background:#10b981"],
      [style*="background: #10B981"],
      [style*="background:#10B981"] {
        background: var(--c4) !important;
        color: #f2f1ec !important;
        border-color: var(--c4) !important;
      }
      [style*="background: #10b981"] svg,
      [style*="background:#10b981"] svg {
        color: #f2f1ec !important;
      }
      [style*="background:#3B82F6"],
      [style*="background:#3b82f6"],
      [style*="background:#8B5CF6"],
      [style*="background:#8b5cf6"],
      [style*="background:#FCD34D"],
      [style*="background:#fcd34d"] {
        background: var(--c4) !important;
        color: #f2f1ec !important;
      }
      [style*="rgba(16,185,129"] {
        background: var(--c2) !important;
        border-color: #c2d6c8 !important;
        color: var(--c4) !important;
      }
      [style*="rgba(245,158,11"] {
        background: #eef2ec !important;
        border-color: #cfdcd3 !important;
        color: #46624f !important;
      }
      [style*="rgba(139,92,246"] {
        background: var(--c3) !important;
        border-color: #b9c8be !important;
        color: var(--c4) !important;
      }
      [style*="rgba(30,41,59"] {
        background: #eef1ec !important;
        border-color: #d6ddd7 !important;
        color: var(--muted) !important;
      }
      [style*="#A78BFA"],
      [style*="#8B5CF6"],
      [style*="#10B981"],
      [style*="#10b981"] {
        color: var(--c4) !important;
      }
      [style*="#FCD34D"],
      [style*="#F59E0B"] {
        color: #46624f !important;
      }
      [style*="#60A5FA"],
      [style*="#94A3B8"] {
        color: var(--muted) !important;
      }
      [style*="background:#0F172A"],
      [style*="background: #0F172A"] {
        background: var(--card) !important;
        color: var(--text) !important;
        border-color: var(--border) !important;
      }
      [style*="#334155"] {
        border-color: var(--border) !important;
      }
      [style*="rgba(59, 130, 246"],
      [style*="rgba(59,130,246"] {
        background: var(--c2) !important;
        border-color: #c2d6c8 !important;
        color: var(--text) !important;
      }
      [style*="rgba(15,23,42,0.6)"],
      [style*="rgba(15, 23, 42, 0.6)"] {
        background: #e8eee9 !important;
      }
      [style*="background: #10b981"],
      [style*="background:#10b981"],
      [style*="background: #10B981"],
      [style*="background:#10B981"] {
        color: #f2f1ec !important;
      }
      [style*="background: #10b981"] svg,
      [style*="background:#10b981"] svg {
        color: #f2f1ec !important;
      }
      [style*="background:#3B82F622"],
      [style*="background:#10B98122"],
      [style*="background:#8B5CF622"],
      [style*="background:#FCD34D22"] {
        background: var(--c2) !important;
        color: var(--c4) !important;
        border-color: #c2d6c8 !important;
      }
      [style*="background:#F59E0B"],
      [style*="background:#f59e0b"] {
        background: var(--c4) !important;
      }
    </style>
  </head>
  <body class="min-h-screen">
    <!-- HEADER -->
    <header
      class="glass border-b border-slate-700/50 sticky top-0 z-50 px-5 py-3 flex items-center justify-between"
    >
      <div class="flex items-center gap-3">
        <div
          class="w-8 h-8 rounded-lg flex items-center justify-center glow"
          style="background: #3b82f6"
        >
          <svg class="w-4 h-4 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path
              stroke-linecap="round"
              stroke-linejoin="round"
              stroke-width="2"
              d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"
            />
          </svg>
        </div>
        <div>
          <h1 class="text-sm font-bold text-white leading-none">DSS Calculator</h1>
          <p class="text-xs text-slate-400 leading-none mt-0.5">Entropy + GRA · MEREC + MARCOS</p>
        </div>
      </div>
      <div class="flex items-center gap-2">
        
        
        <button
          onclick="resetAll()"
          class="text-xs px-3 py-1.5 rounded-lg bg-slate-700 hover:bg-slate-600 text-slate-300 transition"
        >
          ↺ Reset
        </button>
      </div>
    </header>

    <div class="flex" style="min-height: calc(100vh - 52px)">
      <!-- SIDEBAR -->
      <aside
        class="w-60 glass border-r border-slate-700/40 p-4 sticky overflow-y-auto hidden lg:flex flex-col gap-4"
        style="top: 52px; height: calc(100vh - 52px)"
      >
        <div>
          <p class="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-2">Metode</p>
          <div class="space-y-1">
            <button
              onclick="switchMethod('gra')"
              id="btn-gra"
              class="w-full text-left px-3 py-2 rounded-lg text-sm transition bg-blue-600/20 text-blue-300 border border-blue-600/40 font-semibold"
            >
              Entropy + GRA
            </button>
            <button
              onclick="switchMethod('merec_marcos')"
              id="btn-merec_marcos"
              class="w-full text-left px-3 py-2 rounded-lg text-sm transition text-slate-400 hover:bg-slate-700/50"
            >
              MEREC + MARCOS
            </button>
            <button
              onclick="switchMethod('combined')"
              id="btn-combined"
              class="w-full text-left px-3 py-2 rounded-lg text-sm transition text-slate-400 hover:bg-slate-700/50"
            >
              Kombinasi Semua
            </button>
          </div>
        </div>

        <div class="p-3 rounded-lg bg-slate-800/60 border border-slate-700/50">
          <p class="text-xs text-white mb-1 font-semibold">ξ — Koefisien GRA</p>
          <input
            type="number"
            id="xiInput"
            value="0.5"
            min="0.01"
            max="1"
            step="0.05"
            class="w-full bg-slate-900 border border-slate-600 rounded px-2 py-1.5 text-sm text-blue-400 badge focus:outline-none focus:border-blue-500"
          />
          <p class="text-xs text-white mt-1">Default = 0.5</p>
        </div>

        <div id="sidebarSteps" class="space-y-1 text-xs text-slate-400"></div>
      </aside>

      <!-- MAIN -->
      <main class="flex-1 p-5 space-y-6 overflow-x-hidden">
        <!-- INPUT SECTION -->
        <section id="sec-input" class="fade-in">
          <div class="flex items-center gap-3 mb-3">
            <span class="pill bg-blue-900/60 text-blue-300 border border-blue-700/50">INPUT</span>
            <h2 class="text-base font-bold">Matriks Keputusan</h2>
          </div>
          <div class="glass rounded-xl p-5 mb-4">
            <!-- Import + Setup Row -->
            <div class="flex flex-wrap gap-3 items-end mb-4">
              <div>
                <label class="text-xs text-slate-400 mb-1 block">Alternatif (m)</label>
                <input
                  type="number"
                  id="numAlt"
                  value="4"
                  min="2"
                  max="30"
                  class="w-24 bg-slate-900 border border-slate-600 rounded-lg px-3 py-2 text-sm focus:outline-none focus:border-blue-500"
                />
              </div>
              <div>
                <label class="text-xs text-slate-400 mb-1 block">Kriteria (n)</label>
                <input
                  type="number"
                  id="numCrit"
                  value="4"
                  min="2"
                  max="20"
                  class="w-24 bg-slate-900 border border-slate-600 rounded-lg px-3 py-2 text-sm focus:outline-none focus:border-blue-500"
                />
              </div>
              <button
                onclick="buildMatrix()"
                class="px-4 py-2 rounded-lg text-white text-sm font-semibold transition"
                style="background: #3b82f6"
              >
                Buat Tabel
              </button>
              <label class="cursor-pointer">
                <input
                  type="file"
                  id="fileInput"
                  accept=".xlsx,.xls,.csv"
                  class="hidden"
                  onchange="importExcel(event)"
                />
                <div
                  class="px-4 py-2 rounded-lg text-white text-sm font-semibold transition flex items-center gap-2"
                  style="background: #10b981"
                >
                  <svg class="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path
                      stroke-linecap="round"
                      stroke-linejoin="round"
                      stroke-width="2"
                      d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"
                    />
                  </svg>
                  Import Excel
                </div>
              </label>
              <button
                onclick="fillSample()"
                class="px-4 py-2 rounded-lg bg-slate-700 hover:bg-slate-600 text-slate-300 text-sm transition"
              >
                Contoh Data
              </button>
            </div>

            <!-- Format hint -->
            <div
              class="flex items-start gap-2 p-3 rounded-lg mb-4 text-xs"
              style="
                background: rgba(59, 130, 246, 0.08);
                border: 1px solid rgba(59, 130, 246, 0.2);
              "
            >
              <svg
                class="w-4 h-4 text-blue-400 shrink-0 mt-0.5"
                fill="none"
                viewBox="0 0 24 24"
                stroke="currentColor"
              >
                <path
                  stroke-linecap="round"
                  stroke-linejoin="round"
                  stroke-width="2"
                  d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"
                />
              </svg>
              <span class="text-blue-300"
                ><strong>Format Excel:</strong> Baris 1 = nama kriteria (kolom A kosong), Baris 2 =
                tipe (benefit/cost), Baris 3+ = data (kolom A = nama alternatif, kolom B+ = nilai).
                Gunakan template dari tombol di bawah.</span
              >
            </div>

            <div id="matrixArea" class="overflow-x-auto"></div>
          </div>

          <div class="flex flex-wrap gap-3">
            <button
              onclick="calculateAll()"
              class="px-6 py-2.5 rounded-xl text-white font-bold text-sm transition shadow-lg flex items-center gap-2"
              style="background: linear-gradient(135deg, #3b82f6, #6366f1)"
            >
              <svg class="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path
                  stroke-linecap="round"
                  stroke-linejoin="round"
                  stroke-width="2"
                  d="M14.752 11.168l-3.197-2.132A1 1 0 0010 9.87v4.263a1 1 0 001.555.832l3.197-2.132a1 1 0 000-1.664z"
                />
                <path
                  stroke-linecap="round"
                  stroke-linejoin="round"
                  stroke-width="2"
                  d="M21 12a9 9 0 11-18 0 9 9 0 0118 0z"
                />
              </svg>
              Hitung Semua
            </button>
            <button
              onclick="exportTemplate()"
              class="px-4 py-2.5 rounded-xl bg-slate-700 hover:bg-slate-600 text-slate-200 font-semibold text-sm transition flex items-center gap-2"
            >
              <svg class="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path
                  stroke-linecap="round"
                  stroke-linejoin="round"
                  stroke-width="2"
                  d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                />
              </svg>
              Unduh Template Excel
            </button>
          </div>
        </section>

        <!-- RESULTS -->
        <div id="resultsContainer" class="space-y-6"></div>

        <!-- EXPORT -->
        <div id="exportArea" class="hidden fade-in pb-10">
          <button
            onclick="exportResults()"
            class="flex items-center gap-2 px-6 py-3 rounded-xl text-white font-bold text-sm transition shadow-lg"
            style="background: #10b981"
          >
            <svg class="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path
                stroke-linecap="round"
                stroke-linejoin="round"
                stroke-width="2"
                d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
              />
            </svg>
            Export Semua Hasil ke Excel
          </button>
        </div>
      </main>
    </div>

    <script>
      // ═══════════════════════════════════════════════
      //  STATE
      // ═══════════════════════════════════════════════
      let ST = {
        alternatives: [],
        criteria: [],
        types: [],
        matrix: [],
        m: 0,
        n: 0,
        // GRA/Entropy
        normGRA: [],
        pij: [],
        entropy: [],
        degDiv: [],
        weights: [],
        refSeq: [],
        delta: [],
        grc: [],
        grd: [],
        rankGRA: [],
        // MEREC
        normMEREC: [],
        Si: [],
        Sij: [],
        Wi_merec: [],
        rankMEREC: [],
        // MARCOS
        normMARCOS: [],
        wMARCOS: [],
        ideal: [],
        antiIdeal: [],
        sPlus: [],
        sMinus: [],
        kPlus: [],
        kMinus: [],
        kFinal: [],
        rankMARCOS: [],
        activeMethod: "gra",
      };

      const PALETTE = {
        c1: "#f2f1ec",
        c2: "#e3ece6",
        c3: "#d3e4da",
        c4: "#1e4d3b",
        text: "#101418",
        muted: "#5c646a",
        border: "#d2d0c6",
        card: "#fbfaf7",
      };
      const rankColors = ["#1e4d3b", "#5f8b76", "#a9c6b5", "#dbe8e0"];

      function escHTML(v) {
        return String(v ?? "").replace(
          /[&<>'"]/g,
          (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" })[ch],
        );
      }
      function escAttr(v) {
        return escHTML(v);
      }
      function cleanText(v) {
        return String(v ?? "")
          .replace(/\u00A0/g, " ")
          .trim();
      }
      function isBlankCell(v) {
        return v === null || v === undefined || cleanText(v) === "";
      }
      function isNumericLikeCell(v) {
        if (v === null || v === undefined || v === "") return false;
        if (typeof v === "number") return Number.isFinite(v);
        if (v instanceof Date) return false;
        const s = cleanText(v);
        if (s === "") return false;
        // Harus punya digit agar teks seperti nama negara/alternatif tidak terbaca sebagai angka 0.
        if (!/[0-9]/.test(s)) return false;
        return Number.isFinite(parseNumber(s));
      }
      function rowHasData(row) {
        return Array.isArray(row) && row.some((v) => !isBlankCell(v));
      }
      function normalizeLabel(v) {
        return cleanText(v).toLowerCase();
      }
      function parseCriterionType(v) {
        const s = normalizeLabel(v).replace(/[^a-z]/g, "");
        if (["cost", "c", "biaya", "harga", "min", "minimum", "nonbenefit", "costtype"].includes(s))
          return "cost";
        if (["benefit", "b", "beneficial", "profit", "max", "maximum", "keuntungan"].includes(s))
          return "benefit";
        return "benefit";
      }
      function safeDiv(a, b, zeroValue = 0) {
        a = Number(a);
        b = Number(b);
        if (!Number.isFinite(a) || !Number.isFinite(b)) return zeroValue;
        if (Math.abs(b) < 1e-12) return zeroValue;
        const r = a / b;
        return Number.isFinite(r) ? r : zeroValue;
      }

      // ═══════════════════════════════════════════════
      //  BUILD MATRIX (manual)
      // ═══════════════════════════════════════════════
      function buildMatrix(preData) {
        const m = preData ? preData.m : parseInt(document.getElementById("numAlt").value);
        const n = preData ? preData.n : parseInt(document.getElementById("numCrit").value);
        if (preData) {
          document.getElementById("numAlt").value = m;
          document.getElementById("numCrit").value = n;
        }
        let html = `<table>
    <thead>
      <tr style="background:rgba(30,41,59,0.9)">
        <th style="text-align:left;min-width:130px">Alternatif</th>`;
        for (let j = 0; j < n; j++) {
          const cval = preData ? preData.criteria[j] || `C${j + 1}` : `C${j + 1}`;
          html += `<th><input type="text" value="${escAttr(cval)}" data-role="cname" data-j="${j}"
      style="width:90px;background:#0F172A;border:1px solid #334155;border-radius:5px;padding:3px 6px;color:#60A5FA;font-size:11px;text-align:center;outline:none"/></th>`;
        }
        html += `</tr>
      <tr style="background:rgba(15,23,42,0.6)">
        <td style="font-size:11px;color:#94A3B8;font-style:italic;padding:5px 8px">Tipe Kriteria</td>`;
        for (let j = 0; j < n; j++) {
          const tv = preData ? preData.types[j] || "benefit" : "benefit";
          html += `<td style="text-align:center"><select data-role="ctype" data-j="${j}"
      style="background:#0F172A;border:1px solid #334155;border-radius:5px;padding:3px 6px;font-size:11px;color:#10B981;outline:none;width:90px">
      <option value="benefit" ${tv === "benefit" ? "selected" : ""}>benefit</option>
      <option value="cost" ${tv === "cost" ? "selected" : ""}>cost</option>
    </select></td>`;
        }
        html += `</tr></thead><tbody>`;
        for (let i = 0; i < m; i++) {
          const aval = preData ? preData.alternatives[i] || `A${i + 1}` : `A${i + 1}`;
          html += `<tr><td><input type="text" value="${escAttr(aval)}" data-role="aname" data-i="${i}"
      style="width:120px;background:#0F172A;border:1px solid #334155;border-radius:5px;padding:4px 6px;color:#FCD34D;font-size:11px;outline:none"/></td>`;
          for (let j = 0; j < n; j++) {
            const v = preData
              ? preData.matrix?.[i]?.[j] !== undefined
                ? preData.matrix[i][j]
                : ""
              : "";
            html += `<td><input type="number" value="${v}" data-role="val" data-i="${i}" data-j="${j}"
        style="width:90px;background:transparent;border:1px solid #1E293B;border-radius:5px;padding:4px 6px;font-size:11px;text-align:center;color:#F8FAFC;outline:none"
        placeholder="0"
        onfocus="this.style.borderColor='#3B82F6';this.style.background='rgba(30,41,59,0.6)'"
        onblur="this.style.borderColor='#1E293B';this.style.background='transparent'"/></td>`;
          }
          html += `</tr>`;
        }
        html += `</tbody></table>`;
        document.getElementById("matrixArea").innerHTML = html;
      }

      // ═══════════════════════════════════════════════
      //  READ MATRIX FROM DOM
      // ═══════════════════════════════════════════════

      function parseNumber(v) {
        if (v instanceof Date) return 0;
        if (typeof v === "number") return Number.isFinite(v) ? v : 0;
        if (v === null || v === undefined) return 0;

        let s = String(v)
          .trim()
          .replace(/\u00A0/g, "")
          .replace(/\s+/g, "");
        if (s === "") return 0;

        const isPercent = s.endsWith("%");
        s = s.replace(/%$/, "");
        s = s.replace(/[^0-9,\.\-]/g, "");
        if (s === "" || s === "-" || s === "," || s === ".") return 0;

        const lastComma = s.lastIndexOf(",");
        const lastDot = s.lastIndexOf(".");
        if (lastComma > -1 && lastDot > -1) {
          // Pemisah desimal mengikuti separator yang muncul paling akhir.
          if (lastComma > lastDot) s = s.replace(/\./g, "").replace(",", ".");
          else s = s.replace(/,/g, "");
        } else if (lastComma > -1) {
          const parts = s.split(",");
          const looksThousands = parts.length > 1 && parts.slice(1).every((p) => p.length === 3);
          s = looksThousands ? s.replace(/,/g, "") : s.replace(",", ".");
        } else if (lastDot > -1) {
          const parts = s.split(".");
          const looksThousands = parts.length > 1 && parts.slice(1).every((p) => p.length === 3);
          if (looksThousands) s = s.replace(/\./g, "");
        }

        let n = Number(s);
        if (!Number.isFinite(n)) n = 0;
        return isPercent ? n / 100 : n;
      }

      function readMatrix() {
        const m = parseInt(document.getElementById("numAlt").value);
        const n = parseInt(document.getElementById("numCrit").value);
        const cnames = [...document.querySelectorAll('[data-role="cname"]')].map(
          (x) => x.value || `C${parseInt(x.dataset.j) + 1}`,
        );
        const ctypes = [...document.querySelectorAll('[data-role="ctype"]')].map((x) => x.value);
        const anames = [...document.querySelectorAll('[data-role="aname"]')].map(
          (x) => x.value || `A${parseInt(x.dataset.i) + 1}`,
        );
        let matrix = Array.from({ length: m }, () => Array(n).fill(0));
        document.querySelectorAll('[data-role="val"]').forEach((inp) => {
          matrix[parseInt(inp.dataset.i)][parseInt(inp.dataset.j)] = parseNumber(inp.value);
        });
        return { matrix, alternatives: anames, criteria: cnames, types: ctypes, m, n };
      }

      // ═══════════════════════════════════════════════
      //  IMPORT EXCEL — FLEXIBLE & MORE ACCURATE
      // ═══════════════════════════════════════════════
      function normalizeRows(raw) {
        let rows = raw.map((r) => (Array.isArray(r) ? r.slice() : []));
        rows = rows.filter(rowHasData);
        if (!rows.length) return [];

        let maxCols = Math.max(...rows.map((r) => r.length));
        rows = rows.map((r) => Array.from({ length: maxCols }, (_, i) => r[i] ?? ""));

        // Buang kolom kosong di kiri dan kanan akibat formatting Excel.
        while (maxCols > 0 && rows.every((r) => isBlankCell(r[0]))) {
          rows = rows.map((r) => r.slice(1));
          maxCols--;
        }
        while (maxCols > 0 && rows.every((r) => isBlankCell(r[maxCols - 1]))) {
          rows = rows.map((r) => r.slice(0, -1));
          maxCols--;
        }
        return rows;
      }

      function detectExcelStructure(rows) {
        if (!rows.length) return null;
        const maxCols = Math.max(...rows.map((r) => r.length));
        rows = rows.map((r) => Array.from({ length: maxCols }, (_, i) => r[i] ?? ""));

        const typeRowIndex = rows.findIndex((row, idx) => {
          if (idx > 12) return false;
          const vals = row.map(normalizeLabel);
          const compact = vals.map((v) => v.replace(/[^a-z]/g, ""));
          const hasTypeHeader =
            compact.includes("direction") ||
            compact.includes("tipe") ||
            compact.includes("type") ||
            compact.includes("bc");
          const countType = compact.filter(
            (v) =>
              ["benefit", "cost", "biaya", "profit", "max", "min", "minimum", "maximum"].includes(
                v,
              ) ||
              (hasTypeHeader && ["b", "c"].includes(v)),
          ).length;
          return countType >= Math.max(1, Math.ceil((maxCols - 1) * 0.25)) || hasTypeHeader;
        });

        const criteriaRowIndex = typeRowIndex > 0 ? typeRowIndex - 1 : 0;
        const criteriaRow = rows[criteriaRowIndex] || [];
        const typeRow = typeRowIndex >= 0 ? rows[typeRowIndex] : [];

        // Deteksi kolom alternatif. Mendukung format:
        // 1) [kosong, C1, C2, ...]
        // 2) [Alternatif, C1, C2, ...]
        // 3) [No, Alternatif, C1, C2, ...]
        const header0 = normalizeLabel(criteriaRow[0]);
        const header1 = normalizeLabel(criteriaRow[1]);
        const type0 = normalizeLabel(typeRow[0]);
        const type1 = normalizeLabel(typeRow[1]);
        const h0c = header0.replace(/[^a-z]/g, "");
        const h1c = header1.replace(/[^a-z]/g, "");
        const t1c = type1.replace(/[^a-z]/g, "");
        let altCol = 0;
        if (
          ["no", "nomor", "id"].includes(h0c) &&
          [
            "alternatif",
            "alternative",
            "namaalternatif",
            "nama",
            "name",
            "countrycriterion",
            "country",
          ].includes(h1c)
        ) {
          altCol = 1;
        } else if (
          (header0 === "" || ["c"].includes(h0c)) &&
          [
            "alternatif",
            "alternative",
            "namaalternatif",
            "nama",
            "name",
            "countrycriterion",
            "country",
          ].includes(h1c)
        ) {
          altCol = 1;
        } else if (["direction", "tipe", "type", "bc"].includes(t1c) && h1c !== "") {
          altCol = 1;
        }
        const firstCritCol = altCol + 1;

        let candidateCols = [];
        for (let c = firstCritCol; c < maxCols; c++) {
          const hasHeader = !isBlankCell(criteriaRow[c]);
          const hasType = !isBlankCell(typeRow[c]);
          const hasNumeric = rows
            .slice((typeRowIndex >= 0 ? typeRowIndex : criteriaRowIndex) + 1)
            .some((r) => isNumericLikeCell(r[c]));
          if (hasHeader || hasType || hasNumeric) candidateCols.push(c);
        }
        if (!candidateCols.length) return null;

        const dataStart = (typeRowIndex >= 0 ? typeRowIndex : criteriaRowIndex) + 1;
        const dataRows = rows.slice(dataStart).filter((r) => {
          const name = cleanText(r[altCol]);
          const numericCount = candidateCols.filter((c) => isNumericLikeCell(r[c])).length;
          const firstCell = normalizeLabel(r[0]);
          const isInstruction = ["petunjuk", "keterangan", "catatan", "note", "notes"].includes(
            firstCell,
          );
          return !isInstruction && (name !== "" || numericCount > 0) && numericCount > 0;
        });
        if (!dataRows.length) return null;

        const criteria = candidateCols.map((c, i) => {
          const h = cleanText(criteriaRow[c]);
          return h && !["benefit", "cost"].includes(h.toLowerCase()) ? h : `C${i + 1}`;
        });
        const types = candidateCols.map((c) => parseCriterionType(typeRow[c]));
        const alternatives = dataRows.map((r, i) => {
          const name = cleanText(r[altCol]);
          if (name !== "" && !["tipe", "type"].includes(name.toLowerCase())) return name;
          return `A${i + 1}`;
        });
        const matrix = dataRows.map((r) => candidateCols.map((c) => parseNumber(r[c])));

        return {
          m: dataRows.length,
          n: candidateCols.length,
          criteria,
          types,
          alternatives,
          matrix,
        };
      }

      function importExcel(e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function (ev) {
          try {
            const wb = XLSX.read(ev.target.result, { type: "array", cellDates: true, raw: true });
            let parsed = null;

            for (const sheetName of wb.SheetNames) {
              const ws = wb.Sheets[sheetName];
              const raw = XLSX.utils.sheet_to_json(ws, {
                header: 1,
                defval: "",
                blankrows: false,
                raw: true,
              });
              parsed = detectExcelStructure(normalizeRows(raw));
              if (parsed) break;
            }

            if (!parsed) {
              alert(
                "Format Excel belum terbaca. Gunakan format: baris kriteria, baris tipe benefit/cost, lalu data alternatif. Format No + Alternatif juga sudah didukung.",
              );
              return;
            }

            buildMatrix(parsed);
            document.getElementById("resultsContainer").innerHTML = "";
            document.getElementById("exportArea").classList.add("hidden");
            showNotif(`✓ Import berhasil: ${parsed.m} alternatif, ${parsed.n} kriteria`, "green");
          } catch (err) {
            console.error(err);
            alert("Gagal membaca file: " + err.message);
          }
        };
        reader.readAsArrayBuffer(file);
        e.target.value = "";
      }

      // ═══════════════════════════════════════════════
      //  SAMPLE DATA
      // ═══════════════════════════════════════════════
      function fillSample() {
        buildMatrix({
          m: 5,
          n: 5,
          criteria: ["Harga", "Kualitas", "Pengiriman", "Layanan", "Garansi"],
          types: ["cost", "benefit", "benefit", "benefit", "benefit"],
          alternatives: ["Supplier A", "Supplier B", "Supplier C", "Supplier D", "Supplier E"],
          matrix: [
            [80, 85, 90, 88, 3],
            [75, 90, 85, 82, 2],
            [90, 80, 95, 75, 4],
            [70, 95, 80, 92, 5],
            [85, 88, 88, 78, 3],
          ],
        });
      }

      // ═══════════════════════════════════════════════
      //  NOTIF
      // ═══════════════════════════════════════════════
      function showNotif(msg, color = "blue") {
        const el = document.createElement("div");
        el.textContent = msg;
        const bg = PALETTE.c3;
        el.style.cssText = `position:fixed;top:60px;right:16px;z-index:9999;padding:10px 16px;border-radius:8px;font-size:13px;font-weight:700;color:${PALETTE.text};background:${bg};border:1px solid ${PALETTE.border};box-shadow:0 8px 24px rgba(16,20,24,.12);transition:opacity .4s`;
        document.body.appendChild(el);
        setTimeout(() => {
          el.style.opacity = "0";
          setTimeout(() => el.remove(), 400);
        }, 2500);
      }

      // ═══════════════════════════════════════════════
      //  CALCULATE ALL
      // ═══════════════════════════════════════════════
      function calculateAll() {
        const d = readMatrix();
        if (d.matrix.flat().every((v) => v === 0)) {
          alert("Isi data terlebih dahulu!");
          return;
        }
        Object.assign(ST, d);

        calcEntropyGRA();
        calcMEREC();
        calcMARCOS();
        renderResults();
        document.getElementById("exportArea").classList.remove("hidden");
      }

      // ═══════════════════════════════════════════════
      //  1. ENTROPY + GRA
      // ═══════════════════════════════════════════════
      function calcEntropyGRA() {
        const { matrix, types, m, n } = ST;

        // Normalisasi GRA mengikuti sheet GRA_Rumus:
        // Benefit = x/max, Cost = min/x.
        let norm = Array.from({ length: m }, () => Array(n).fill(0));
        for (let j = 0; j < n; j++) {
          const col = matrix.map((r) => Number(r[j]) || 0);
          const mx = Math.max(...col),
            mn = Math.min(...col);
          for (let i = 0; i < m; i++) {
            const x = Number(matrix[i][j]) || 0;
            if (types[j] === "benefit") norm[i][j] = mx === 0 ? 0 : safeDiv(x, mx, 0);
            else norm[i][j] = x === 0 ? (mn === 0 ? 1 : 0) : safeDiv(mn, x, 0);
          }
        }
        ST.normGRA = norm;

        // Bobot Entropy mengikuti sheet Entropy:
        // p_ij = x_ij / jumlah kolom matriks awal, bukan dari matriks normalisasi.
        let pij = Array.from({ length: m }, () => Array(n).fill(0));
        for (let j = 0; j < n; j++) {
          const s = matrix.reduce((acc, r) => acc + (Number(r[j]) || 0), 0);
          for (let i = 0; i < m; i++) pij[i][j] = s === 0 ? 0 : (Number(matrix[i][j]) || 0) / s;
        }
        ST.pij = pij;

        const k = m > 1 ? 1 / Math.log(m) : 0;
        ST.entropy = Array(n).fill(0);
        for (let j = 0; j < n; j++) {
          let s = 0;
          for (let i = 0; i < m; i++) {
            const p = pij[i][j];
            if (p > 0) s += p * Math.log(p);
          }
          ST.entropy[j] = Math.min(1, Math.max(0, -k * s));
        }

        ST.degDiv = ST.entropy.map((e) => Math.max(0, 1 - e));
        const sumD = ST.degDiv.reduce((a, b) => a + b, 0);
        ST.weights = ST.degDiv.map((d) => (sumD === 0 ? 1 / n : d / sumD));

        // Di Excel referensi GRA = 1 untuk setiap kriteria setelah normalisasi x/max atau min/x.
        ST.refSeq = Array(n).fill(1);

        ST.delta = Array.from({ length: m }, () => Array(n).fill(0));
        for (let i = 0; i < m; i++)
          for (let j = 0; j < n; j++) ST.delta[i][j] = Math.abs(ST.refSeq[j] - norm[i][j]);

        const xi = parseFloat(document.getElementById("xiInput").value) || 0.5;
        const allD = ST.delta.flat(),
          dMin = Math.min(...allD),
          dMax = Math.max(...allD);
        ST.grc = Array.from({ length: m }, () => Array(n).fill(0));
        for (let i = 0; i < m; i++)
          for (let j = 0; j < n; j++) {
            ST.grc[i][j] =
              dMax === 0 ? 1 : safeDiv(dMin + xi * dMax, ST.delta[i][j] + xi * dMax, 1);
          }

        ST.grd = Array(m).fill(0);
        for (let i = 0; i < m; i++)
          for (let j = 0; j < n; j++) ST.grd[i] += ST.weights[j] * ST.grc[i][j];

        ST.rankGRA = makeRank(ST.grd, ST.alternatives, "desc");
      }

      // ═══════════════════════════════════════════════
      //  2. MEREC
      // ═══════════════════════════════════════════════
      function calcMEREC() {
        const { matrix, types, m, n } = ST;

        // Normalisasi MEREC disamakan dengan file Excel:
        // Benefit/max = x/max, Cost/min = min/x.
        let norm = Array.from({ length: m }, () => Array(n).fill(0));
        for (let j = 0; j < n; j++) {
          const col = matrix.map((r) => Number(r[j]) || 0),
            mx = Math.max(...col),
            mn = Math.min(...col);
          for (let i = 0; i < m; i++) {
            const x = Number(matrix[i][j]) || 0;
            if (types[j] === "benefit") norm[i][j] = mx === 0 ? 0 : safeDiv(x, mx, 0);
            else norm[i][j] = x === 0 ? (mn === 0 ? 1 : 0) : safeDiv(mn, x, 0);
          }
        }
        ST.normMEREC = norm;

        // Kinerja awal sesuai Excel: Si = LN(1 + SUM(nij)/n)
        ST.Si = Array(m).fill(0);
        for (let i = 0; i < m; i++) {
          const s = norm[i].reduce((a, b) => a + b, 0);
          ST.Si[i] = Math.log(1 + s / n);
        }

        // Penghapusan kriteria mengikuti rumus Excel.
        // Catatan: pada file Excel, hanya term pertama setelah kriteria dihapus yang dikali 1/n.
        // Pola ini dipertahankan agar hasil website sama persis dengan Excel.
        ST.Sij = Array.from({ length: m }, () => Array(n).fill(0));
        for (let i = 0; i < m; i++)
          for (let j = 0; j < n; j++) {
            const terms = [];
            for (let k2 = 0; k2 < n; k2++) {
              if (k2 !== j) terms.push(Math.abs(Math.log(Math.max(norm[i][k2], 1e-300))));
            }
            const val = terms.length ? terms[0] / n + terms.slice(1).reduce((a, b) => a + b, 0) : 0;
            ST.Sij[i][j] = Math.log(1 + val);
          }

        let Hj = Array(n).fill(0);
        for (let j = 0; j < n; j++)
          for (let i = 0; i < m; i++) Hj[j] += Math.abs(ST.Sij[i][j] - ST.Si[i]);

        const sumH = Hj.reduce((a, b) => a + b, 0);
        ST.Wi_merec = Hj.map((h) => (sumH === 0 ? 1 / n : h / sumH));
        ST.Hj = Hj;

        // MEREC hanya dipakai sampai tahap pembobotan.
        // Ranking alternatif dikerjakan pada tahap MARCOS.
        ST.skorMEREC = [];
        ST.rankMEREC = [];
      }

      // ═══════════════════════════════════════════════
      //  3. MARCOS
      // ═══════════════════════════════════════════════
      function calcMARCOS() {
        const { matrix, types, m, n } = ST;

        let ideal = Array(n).fill(0),
          antiIdeal = Array(n).fill(0);
        for (let j = 0; j < n; j++) {
          const col = matrix.map((r) => Number(r[j]) || 0),
            mx = Math.max(...col),
            mn = Math.min(...col);
          ideal[j] = types[j] === "benefit" ? mx : mn;
          antiIdeal[j] = types[j] === "benefit" ? mn : mx;
        }
        ST.ideal = ideal;
        ST.antiIdeal = antiIdeal;

        const extRows = [antiIdeal, ...matrix, ideal];
        const extM = m + 2;

        let normExt = Array.from({ length: extM }, () => Array(n).fill(0));
        for (let j = 0; j < n; j++) {
          const refVal = ideal[j]; // MARCOS: benefit = x/AI, cost = AI/x
          for (let i = 0; i < extM; i++) {
            const x = Number(extRows[i][j]) || 0;
            if (types[j] === "benefit") normExt[i][j] = refVal === 0 ? 0 : safeDiv(x, refVal, 0);
            else normExt[i][j] = x === 0 ? (refVal === 0 ? 1 : 0) : safeDiv(refVal, x, 0);
          }
        }
        ST.normMARCOS = normExt;

        // Pada pasangan MEREC + MARCOS, bobot kriteria yang dipakai adalah bobot hasil MEREC.
        const W = ST.Wi_merec && ST.Wi_merec.length ? ST.Wi_merec : ST.weights;
        let Si = Array(extM).fill(0);
        for (let i = 0; i < extM; i++) for (let j = 0; j < n; j++) Si[i] += W[j] * normExt[i][j];
        ST.Si_marcos = Si;

        const s_ai = Si[extM - 1],
          s_aai = Si[0];
        const Si_alt = Si.slice(1, extM - 1);

        ST.kPlus = Si_alt.map((s) => safeDiv(s, s_ai, 0));
        ST.kMinus = Si_alt.map((s) => safeDiv(s, s_aai, 0));

        const fPlus = ST.kPlus.map((_, i) => safeDiv(ST.kMinus[i], ST.kPlus[i] + ST.kMinus[i], 0));
        const fMinus = ST.kMinus.map((_, i) => safeDiv(ST.kPlus[i], ST.kPlus[i] + ST.kMinus[i], 0));

        ST.kFinal = ST.kPlus.map((_, i) => {
          const fp = fPlus[i],
            fm = fMinus[i];
          if (fp <= 0 || fm <= 0) return 0;
          const denom = 1 + safeDiv(1 - fp, fp, 0) + safeDiv(1 - fm, fm, 0);
          return safeDiv(ST.kPlus[i] + ST.kMinus[i], denom, 0);
        });
        ST.fPlus = fPlus;
        ST.fMinus = fMinus;
        ST.s_ai = s_ai;
        ST.s_aai = s_aai;

        ST.rankMARCOS = makeRank(ST.kFinal, ST.alternatives, "desc");
      }

      // ═══════════════════════════════════════════════
      //  HELPER: makeRank
      // ═══════════════════════════════════════════════
      function makeRank(scores, names, dir = "desc") {
        return scores
          .map((v, i) => ({ idx: i, val: v, name: names[i] }))
          .sort((a, b) => (dir === "desc" ? b.val - a.val : a.val - b.val))
          .map((r, rank) => ({ ...r, rank: rank + 1 }))
          .sort((a, b) => a.idx - b.idx);
      }

      // ═══════════════════════════════════════════════
      //  RENDER
      // ═══════════════════════════════════════════════
      const fmt = (v) => {
        const num = Number(v);
        if (!Number.isFinite(num)) return "-";
        return Math.abs(num) < 0.00001 && num !== 0 ? num.toExponential(3) : num.toFixed(4);
      };
      const pct = (v) => {
        const num = Number(v);
        return Number.isFinite(num) ? (num * 100).toFixed(2) + "%" : "-";
      };

      function sectionHdr(num, title, sub, color = PALETTE.c4) {
        return `<div class="flex items-start gap-3 mb-3">
    <span class="pill shrink-0 mt-0.5" style="background:${color}22;color:${color};border:1px solid ${color}44">STEP ${num}</span>
    <div><p class="font-bold text-sm">${title}</p><p class="text-xs text-slate-400 mt-0.5">${sub}</p></div>
  </div>`;
      }
      function formulaBox(f) {
        return `<div class="mb-3 p-3 rounded-lg text-xs badge" style="background:${PALETTE.c1};border:1px solid ${PALETTE.border};color:${PALETTE.text}">${f}</div>`;
      }
      function matTable(data, rowNames, colNames, highlightCol = -1) {
        let h = `<div class="overflow-x-auto"><table><thead><tr><th style="text-align:left">Alternatif</th>`;
        colNames.forEach(
          (c, j) =>
            (h += `<th style="${j === highlightCol ? "color:var(--text);background:var(--c2)" : ""}">${escHTML(c)}</th>`),
        );
        h += `</tr></thead><tbody>`;
        data.forEach((row, i) => {
          h += `<tr><td style="font-weight:600;color:var(--text);white-space:nowrap">${escHTML(rowNames[i])}</td>`;
          row.forEach((v) => (h += `<td class="badge" style="text-align:center">${fmt(v)}</td>`));
          h += `</tr>`;
        });
        return h + `</tbody></table></div>`;
      }
      function rankCard(ranked, scoreLabel) {
        const sorted = [...ranked].sort((a, b) => a.rank - b.rank);
        const total = Math.max(1, sorted.length);
        return sorted
          .map((r) => {
            const pct2 = total === 1 ? 100 : Math.max(12, 100 - (r.rank - 1) * (78 / (total - 1)));
            const c = rankColors[(r.rank - 1) % rankColors.length];
            return `<div class="p-3 rounded-xl mb-2 flex items-center gap-3" style="background:${PALETTE.card};border:1px solid ${r.rank === 1 ? PALETTE.c4 : PALETTE.border}">
      <span class="w-7 h-7 rounded-full flex items-center justify-center text-xs font-bold shrink-0" style="background:${c};color:${r.rank === 1 ? '#f2f1ec' : PALETTE.text}">${r.rank}</span>
      <div class="flex-1 min-w-0">
        <div class="flex items-center justify-between mb-1">
          <span class="text-sm font-semibold truncate">${escHTML(r.name)}</span>
          <span class="badge text-xs ml-2" style="color:${PALETTE.text}">${fmt(r.val)}</span>
        </div>
        <div class="h-1.5 rounded-full overflow-hidden" style="background:${PALETTE.c1}">
          <div class="h-full rounded-full" style="width:${pct2}%;background:${c};transition:width .5s"></div>
        </div>
      </div>
      ${r.rank === 1 ? `<span class="pill" style="background:${PALETTE.c2};color:${PALETTE.text};border:1px solid ${PALETTE.border}">Terbaik</span>` : ""}
    </div>`;
          })
          .join("");
      }

      // ═══════════════════════════════════════════════
      //  RENDER ALL RESULTS
      // ═══════════════════════════════════════════════
      function renderResults() {
        const {
          alternatives,
          criteria,
          types,
          matrix,
          normGRA,
          pij,
          entropy,
          degDiv,
          weights,
          refSeq,
          delta,
          grc,
          grd,
          rankGRA,
          normMEREC,
          Wi_merec,
          Si,
          Sij,
          Hj,
          normMARCOS,
          ideal,
          antiIdeal,
          kPlus,
          kMinus,
          kFinal,
          fPlus,
          fMinus,
          rankMARCOS,
          s_ai,
          s_aai,
          Si_marcos,
          m,
          n,
        } = ST;

        const xi = parseFloat(document.getElementById("xiInput").value) || 0.5;
        const allD = delta.flat(),
          dMin = Math.min(...allD),
          dMax = Math.max(...allD);
        const extAlt = ["AAI", ...alternatives, "AI"];

        let html = "";

        // ── DIVIDER ──
        html += `<div class="flex items-center gap-3 my-2">
    <div class="flex-1 h-px" style="background:linear-gradient(to right,transparent,#334155)"></div>
    <span class="text-xs text-slate-500 font-semibold tracking-widest uppercase">Hasil Perhitungan</span>
    <div class="flex-1 h-px" style="background:linear-gradient(to left,transparent,#334155)"></div>
  </div>`;

        // ════════ BLOCK A: ENTROPY + GRA ════════
        html += `<div id="block-gra">
  <div class="flex items-center gap-2 mb-4">
    <div class="w-1 h-6 rounded" style="background:#3B82F6"></div>
    <h3 class="text-base font-bold" style="color:#60A5FA">Metode Entropy & GRA</h3>
  </div>`;

        // Step 2: Normalisasi GRA
        html += `<div id="sec-normalize" class="glass rounded-xl p-5 mb-4 fade-in">
    ${sectionHdr(2, "Normalisasi Matriks", "Benefit: xᵢⱼ/max(xⱼ) | Cost: min(xⱼ)/xᵢⱼ")}
    <div class="flex flex-wrap gap-2 mb-3">
      ${types.map((t, j) => `<span class="pill" style="background:${t === "benefit" ? "rgba(16,185,129,0.15)" : "rgba(245,158,11,0.15)"};color:${t === "benefit" ? "#10B981" : "#F59E0B"};border:1px solid ${t === "benefit" ? "rgba(16,185,129,0.3)" : "rgba(245,158,11,0.3)"}">${criteria[j]}: ${t}</span>`).join("")}
    </div>
    ${matTable(normGRA, alternatives, criteria)}
  </div>`;

        // Step 3: Entropy
        html += `<div id="sec-entropy" class="glass rounded-xl p-5 mb-4 fade-in">
    ${sectionHdr(3, "Nilai Entropy (Eⱼ)", "pᵢⱼ = xᵢⱼ/Σxᵢⱼ, Eⱼ = -k·Σ(pᵢⱼ·ln pᵢⱼ), k=1/ln(m)")}
    ${formulaBox(`k = 1/ln(${m}) = <strong>${fmt(1 / Math.log(m))}</strong>`)}
    <p class="text-xs text-slate-400 mb-2">Tabel Probabilitas pᵢⱼ dari matriks awal:</p>
    ${matTable(pij, alternatives, criteria)}
    <p class="text-xs text-slate-400 mt-3 mb-2">Nilai Entropy Eⱼ:</p>
    <div class="overflow-x-auto"><table><thead><tr>${criteria.map((c) => `<th>${c}</th>`).join("")}</tr></thead>
    <tbody><tr>${entropy.map((e) => `<td class="badge" style="text-align:center;color:#93C5FD">${fmt(e)}</td>`).join("")}</tr></tbody></table></div>
  </div>`;

        // Step 4: Bobot
        html += `<div id="sec-weight" class="glass rounded-xl p-5 mb-4 fade-in">
    ${sectionHdr(4, "Bobot Entropy (Wⱼ)", "dⱼ = 1−Eⱼ, Wⱼ = dⱼ/Σdⱼ")}
    <div class="overflow-x-auto mb-3"><table><thead>
      <tr><th style="text-align:left">Parameter</th>${criteria.map((c) => `<th>${c}</th>`).join("")}<th style="color:#FCD34D">Total</th></tr>
    </thead><tbody>
      <tr><td class="badge">Entropy Eⱼ</td>${entropy.map((e) => `<td class="badge" style="text-align:center">${fmt(e)}</td>`).join("")}<td class="badge" style="text-align:center;color:#94A3B8">${fmt(entropy.reduce((a, b) => a + b, 0))}</td></tr>
      <tr><td class="badge">Divergence dⱼ</td>${degDiv.map((d) => `<td class="badge" style="text-align:center">${fmt(d)}</td>`).join("")}<td class="badge" style="text-align:center;color:#FCD34D">${fmt(degDiv.reduce((a, b) => a + b, 0))}</td></tr>
      <tr><td class="badge" style="color:#10B981;font-weight:600">Bobot Wⱼ</td>${weights.map((w) => `<td class="badge" style="text-align:center;color:#10B981;font-weight:600">${fmt(w)}</td>`).join("")}<td class="badge" style="text-align:center;color:#10B981;font-weight:700">${fmt(weights.reduce((a, b) => a + b, 0))}</td></tr>
    </tbody></table></div>
    <div class="flex flex-wrap gap-2">
      ${criteria.map((c, j) => `<span class="pill" style="background:rgba(16,185,129,0.12);color:#10B981;border:1px solid rgba(16,185,129,0.3)">${c}: ${pct(weights[j])}</span>`).join("")}
    </div>
  </div>`;

        // Step 5: Referensi
        html += `<div id="sec-ref" class="glass rounded-xl p-5 mb-4 fade-in">
    ${sectionHdr(5, "Deret Referensi (x₀)", "Nilai max tiap kolom pada matriks ternormalisasi")}
    <div class="overflow-x-auto"><table><thead><tr><th style="text-align:left">Referensi</th>${criteria.map((c) => `<th>${c}</th>`).join("")}</tr></thead>
    <tbody><tr><td style="font-weight:600;color:#60A5FA">x₀</td>${refSeq.map((v) => `<td class="badge" style="text-align:center;color:#60A5FA;font-weight:600">${fmt(v)}</td>`).join("")}</tr>
    </tbody></table></div>
  </div>`;

        // Step 6: Delta
        html += `<div id="sec-delta" class="glass rounded-xl p-5 mb-4 fade-in">
    ${sectionHdr(6, "Selisih Absolut (Δᵢⱼ)", "Δᵢⱼ = |x₀ⱼ − xᵢⱼ|")}
    ${formulaBox(`Δ<sub>min</sub> = <strong>${fmt(dMin)}</strong> &nbsp;|&nbsp; Δ<sub>max</sub> = <strong>${fmt(dMax)}</strong>`)}
    ${matTable(delta, alternatives, criteria)}
  </div>`;

        // Step 7: GRC
        html += `<div id="sec-grc" class="glass rounded-xl p-5 mb-4 fade-in">
    ${sectionHdr(7, "Grey Relational Coefficient (γᵢⱼ)", "γᵢⱼ = (Δmin + ξ·Δmax) / (Δᵢⱼ + ξ·Δmax)")}
    ${formulaBox(`ξ = ${xi}, Δmin = ${fmt(dMin)}, Δmax = ${fmt(dMax)}`)}
    ${matTable(grc, alternatives, criteria)}
  </div>`;

        // Step 8: GRD & Ranking
        html += `<div id="sec-grd" class="glass rounded-xl p-5 mb-4 fade-in">
    ${sectionHdr(8, "GRD & Ranking GRA", "Γᵢ = Σ(Wⱼ·γᵢⱼ) — nilai lebih besar = lebih baik")}
    <div class="grid md:grid-cols-2 gap-5">
      <div class="overflow-x-auto"><table><thead>
        <tr><th style="text-align:left">Alternatif</th>${criteria.map((c) => `<th>${c}</th>`).join("")}<th style="color:#FCD34D">GRD</th><th style="color:#60A5FA">Rank</th></tr>
      </thead><tbody>
        ${rankGRA
          .map(
            (r) => `<tr>
          <td style="font-weight:600;color:#FCD34D;white-space:nowrap">${r.name}</td>
          ${grc[r.idx].map((v, j) => `<td class="badge" style="text-align:center">${fmt(weights[j] * v)}</td>`).join("")}
          <td class="badge" style="text-align:center;font-weight:700;color:#FCD34D">${fmt(r.val)}</td>
          <td style="text-align:center"><span class="pill" style="background:${r.rank === 1 ? "rgba(59,130,246,0.2)" : "rgba(30,41,59,0.5)"};color:${r.rank === 1 ? "#60A5FA" : "#94A3B8"};border:1px solid ${r.rank === 1 ? "rgba(59,130,246,0.4)" : "#334155"}">#${r.rank}</span></td>
        </tr>`,
          )
          .join("")}
      </tbody></table></div>
      <div>${rankCard(rankGRA, "GRD")}</div>
    </div>
  </div></div>`;

        // ════════ BLOCK B: MEREC + MARCOS ════════
        html += `<div id="block-merec">
  <div class="flex items-center gap-2 mb-4 mt-6">
    <div class="w-1 h-6 rounded" style="background:#10B981"></div>
    <h3 class="text-base font-bold" style="color:#34D399">Metode MEREC + MARCOS</h3>
  </div>
  <p class="text-xs text-slate-400 mb-4">Bagian MEREC berhenti sampai pembobotan kriteria, lalu MARCOS memakai bobot tersebut untuk menghitung utility dan ranking akhir.</p>`;

        // Step M1: Normalisasi MEREC
        html += `<div id="sec-merec-norm" class="glass rounded-xl p-5 mb-4 fade-in">
    ${sectionHdr("M1", "Normalisasi MEREC", "Benefit: xᵢⱼ/max(xⱼ) | Cost: min(xⱼ)/xᵢⱼ", "#10B981")}
    ${formulaBox("Benefit: n<sub>ij</sub> = x<sub>ij</sub> / max(x<sub>j</sub>) &nbsp;|&nbsp; Cost: n<sub>ij</sub> = min(x<sub>j</sub>) / x<sub>ij</sub>")}
    ${matTable(normMEREC, alternatives, criteria)}
  </div>`;

        // Step M2: Si
        html += `<div id="sec-merec-si" class="glass rounded-xl p-5 mb-4 fade-in">
    ${sectionHdr("M2", "Nilai Performa Si", "Sᵢ = ln(1 + Σnᵢⱼ/n)", "#10B981")}
    ${formulaBox("S<sub>i</sub> = ln(1 + Σn<sub>ij</sub>/n)")}
    <div class="overflow-x-auto"><table><thead><tr>${alternatives.map((a) => `<th>${a}</th>`).join("")}</tr></thead>
    <tbody><tr>${Si.map((s) => `<td class="badge" style="text-align:center;color:#34D399">${fmt(s)}</td>`).join("")}</tr></tbody></table></div>
  </div>`;

        // Step M3: Sij (exclude)
        html += `<div id="sec-merec-sij" class="glass rounded-xl p-5 mb-4 fade-in">
    ${sectionHdr("M3", "Performa Tanpa Kriteria j (Sᵢⱼ)", "Sᵢⱼ dihitung tanpa menyertakan kriteria j", "#10B981")}
    ${matTable(Sij, alternatives, criteria)}
  </div>`;

        // Step M4: Bobot MEREC
        html += `<div id="sec-merec-w" class="glass rounded-xl p-5 mb-4 fade-in">
    ${sectionHdr("M4", "Bobot MEREC (Wⱼ)", "Hⱼ = Σ|Sᵢⱼ − Sᵢ|, Wⱼ = Hⱼ/ΣHⱼ", "#10B981")}
    <div class="overflow-x-auto mb-3"><table><thead>
      <tr><th style="text-align:left">Parameter</th>${criteria.map((c) => `<th>${c}</th>`).join("")}<th style="color:#FCD34D">Total</th></tr>
    </thead><tbody>
      <tr><td class="badge">Hⱼ</td>${Hj.map((h) => `<td class="badge" style="text-align:center">${fmt(h)}</td>`).join("")}<td class="badge" style="text-align:center;color:#FCD34D">${fmt(Hj.reduce((a, b) => a + b, 0))}</td></tr>
      <tr><td class="badge" style="color:#10B981;font-weight:600">Wⱼ</td>${Wi_merec.map((w) => `<td class="badge" style="text-align:center;color:#10B981;font-weight:600">${fmt(w)}</td>`).join("")}<td class="badge" style="text-align:center;color:#10B981;font-weight:700">${fmt(Wi_merec.reduce((a, b) => a + b, 0))}</td></tr>
    </tbody></table></div>
    <div class="flex flex-wrap gap-2">
      ${criteria.map((c, j) => `<span class="pill" style="background:rgba(16,185,129,0.12);color:#10B981;border:1px solid rgba(16,185,129,0.3)">${c}: ${pct(Wi_merec[j])}</span>`).join("")}
    </div>
  </div>`;

        // MEREC selesai sampai M4: Bobot. Ranking alternatif dilanjutkan di MARCOS.
        html += `</div>`;

        // ════════ BLOCK C: MARCOS, ditampilkan satu paket dengan MEREC ════════
        html += `<div id="block-marcos">
  <div class="flex items-center gap-2 mb-4 mt-6">
    <div class="w-1 h-6 rounded" style="background:#8B5CF6"></div>
    <h3 class="text-base font-bold" style="color:#A78BFA">Lanjutan MARCOS (Measurement of Alternatives and Ranking according to Compromise Solution)</h3>
  </div>`;

        // Step C1: AI & AAI
        html += `<div id="sec-marcos-ref" class="glass rounded-xl p-5 mb-4 fade-in">
    ${sectionHdr("C1", "Ideal (AI) & Anti-Ideal (AAI)", "AI = nilai terbaik per kriteria, AAI = nilai terburuk", "#8B5CF6")}
    <div class="overflow-x-auto"><table><thead>
      <tr><th style="text-align:left">Referensi</th>${criteria.map((c) => `<th>${c}</th>`).join("")}</tr>
    </thead><tbody>
      <tr><td style="color:#EF4444;font-weight:600">AAI</td>${antiIdeal.map((v) => `<td class="badge" style="text-align:center;color:#EF4444">${fmt(v)}</td>`).join("")}</tr>
      <tr><td style="color:#10B981;font-weight:600">AI</td>${ideal.map((v) => `<td class="badge" style="text-align:center;color:#10B981">${fmt(v)}</td>`).join("")}</tr>
    </tbody></table></div>
  </div>`;

        // Step C2: Normalisasi MARCOS
        html += `<div id="sec-marcos-norm" class="glass rounded-xl p-5 mb-4 fade-in">
    ${sectionHdr("C2", "Normalisasi Matriks Diperluas", "Benefit: xᵢⱼ/AI | Cost: AAI/xᵢⱼ (termasuk baris AAI & AI)", "#8B5CF6")}
    ${matTable(normMARCOS, extAlt, criteria)}
  </div>`;

        // Step C3: Si MARCOS
        html += `<div id="sec-marcos-si" class="glass rounded-xl p-5 mb-4 fade-in">
    ${sectionHdr("C3", "Nilai Utility Sᵢ", "Sᵢ = Σ(Wⱼ·nᵢⱼ) menggunakan bobot dari MEREC", "#8B5CF6")}
    ${formulaBox(`S_AI = <strong>${fmt(s_ai)}</strong> &nbsp;|&nbsp; S_AAI = <strong>${fmt(s_aai)}</strong>`)}
    <div class="overflow-x-auto"><table><thead>
      <tr><th style="text-align:left">Alternatif</th><th>Sᵢ</th><th>Kᵢ⁺ (vs AI)</th><th>Kᵢ⁻ (vs AAI)</th><th>f(Kᵢ⁺)</th><th>f(Kᵢ⁻)</th><th style="color:#A78BFA">Kᵢ Final</th></tr>
    </thead><tbody>
      ${rankMARCOS
        .map(
          (r) => `<tr>
        <td style="font-weight:600;color:#FCD34D;white-space:nowrap">${r.name}</td>
        <td class="badge" style="text-align:center">${fmt(ST.Si_marcos[r.idx + 1])}</td>
        <td class="badge" style="text-align:center">${fmt(kPlus[r.idx])}</td>
        <td class="badge" style="text-align:center">${fmt(kMinus[r.idx])}</td>
        <td class="badge" style="text-align:center">${fmt(fPlus[r.idx])}</td>
        <td class="badge" style="text-align:center">${fmt(fMinus[r.idx])}</td>
        <td class="badge" style="text-align:center;color:#A78BFA;font-weight:700">${fmt(r.val)}</td>
      </tr>`,
        )
        .join("")}
    </tbody></table></div>
  </div>`;

        // Step C4: Ranking MARCOS
        html += `<div id="sec-marcos-rank" class="glass rounded-xl p-5 mb-4 fade-in">
    ${sectionHdr("C4", "Ranking MARCOS", "Kᵢ terbesar = alternatif terbaik", "#8B5CF6")}
    <div class="grid md:grid-cols-2 gap-5">
      <div class="overflow-x-auto"><table><thead>
        <tr><th style="text-align:left">Alternatif</th><th>Kᵢ Final</th><th>Rank</th></tr>
      </thead><tbody>
        ${rankMARCOS
          .map(
            (r) => `<tr>
          <td style="font-weight:600;color:#FCD34D;white-space:nowrap">${r.name}</td>
          <td class="badge" style="text-align:center;color:#A78BFA;font-weight:700">${fmt(r.val)}</td>
          <td style="text-align:center"><span class="pill" style="background:${r.rank === 1 ? "rgba(139,92,246,0.2)" : "rgba(30,41,59,0.5)"};color:${r.rank === 1 ? "#A78BFA" : "#94A3B8"};border:1px solid ${r.rank === 1 ? "rgba(139,92,246,0.4)" : "#334155"}">#${r.rank}</span></td>
        </tr>`,
          )
          .join("")}
      </tbody></table></div>
      <div>${rankCard(rankMARCOS, "Kᵢ")}</div>
    </div>
  </div></div>`;

        // ════════ BLOCK D: REKAPITULASI ════════
        html += `<div id="block-recap">
  <div class="flex items-center gap-2 mb-4 mt-6">
    <div class="w-1 h-6 rounded" style="background:#F59E0B"></div>
    <h3 class="text-base font-bold" style="color:#FCD34D">Rekapitulasi Ranking GRA & MARCOS</h3>
  </div>
  <div class="glass rounded-xl p-5 fade-in">
    <div class="overflow-x-auto"><table><thead>
      <tr>
        <th style="text-align:left">Alternatif</th>
        <th style="color:#60A5FA">GRA</th><th style="color:#60A5FA">Rank GRA</th>
        <th style="color:#A78BFA">MARCOS Kᵢ</th><th style="color:#A78BFA">Rank MARCOS</th>
        <th style="color:#FCD34D">Avg Rank</th>
      </tr>
    </thead><tbody>`;

        const avgRanks = alternatives
          .map((_, i) => {
            const rg = rankGRA[i].rank,
              rc = rankMARCOS[i].rank;
            return { i, avg: (rg + rc) / 2 };
          })
          .sort((a, b) => a.avg - b.avg)
          .map((r, rank) => ({ ...r, finalRank: rank + 1 }))
          .sort((a, b) => a.i - b.i);

        alternatives.forEach((_, i) => {
          const ar = avgRanks[i];
          html += `<tr>
      <td style="font-weight:600;color:#FCD34D;white-space:nowrap">${alternatives[i]}</td>
      <td class="badge" style="text-align:center;color:#60A5FA">${fmt(grd[i])}</td>
      <td style="text-align:center"><span class="pill" style="background:${rankGRA[i].rank === 1 ? "rgba(59,130,246,0.2)" : "transparent"};color:${rankGRA[i].rank === 1 ? "#60A5FA" : "#94A3B8"};border:1px solid ${rankGRA[i].rank === 1 ? "rgba(59,130,246,0.4)" : "#334155"}">#${rankGRA[i].rank}</span></td>
      <td class="badge" style="text-align:center;color:#A78BFA">${fmt(kFinal[i])}</td>
      <td style="text-align:center"><span class="pill" style="background:${rankMARCOS[i].rank === 1 ? "rgba(139,92,246,0.2)" : "transparent"};color:${rankMARCOS[i].rank === 1 ? "#A78BFA" : "#94A3B8"};border:1px solid ${rankMARCOS[i].rank === 1 ? "rgba(139,92,246,0.4)" : "#334155"}">#${rankMARCOS[i].rank}</span></td>
      <td style="text-align:center"><span class="pill" style="background:${ar.finalRank === 1 ? "rgba(245,158,11,0.2)" : "rgba(30,41,59,0.5)"};color:${ar.finalRank === 1 ? "#FCD34D" : "#94A3B8"};border:1px solid ${ar.finalRank === 1 ? "rgba(245,158,11,0.4)" : "#334155"};font-weight:${ar.finalRank === 1 ? 700 : 400}">#${ar.finalRank} (avg ${ar.avg.toFixed(2)})</span></td>
    </tr>`;
        });

        html += `</tbody></table></div>
  </div></div>`;

        document.getElementById("resultsContainer").innerHTML = html;
      }

      // ═══════════════════════════════════════════════
      //  EXPORT RESULTS
      // ═══════════════════════════════════════════════
      function exportResults() {
        try {
          if (!ST.matrix || !ST.matrix.length || !ST.criteria || !ST.criteria.length) {
            const d = readMatrix();
            if (!d.matrix || d.matrix.flat().every((v) => Number(v) === 0)) {
              alert(
                "Data belum tersedia. Import/isi data terlebih dahulu, lalu klik Hitung atau Export.",
              );
              return;
            }
            Object.assign(ST, d);
          }

          if (!ST.rankGRA?.length || !ST.Wi_merec?.length || !ST.rankMARCOS?.length) {
            calcEntropyGRA();
            calcMEREC();
            calcMARCOS();
            renderResults();
            document.getElementById("exportArea")?.classList.remove("hidden");
          }

          const {
            alternatives,
            criteria,
            types,
            matrix,
            normGRA,
            pij,
            entropy,
            degDiv,
            weights,
            delta,
            grc,
            grd,
            rankGRA,
            normMEREC,
            Wi_merec,
            Si,
            Sij,
            Hj,
            normMARCOS,
            ideal,
            antiIdeal,
            kPlus,
            kMinus,
            kFinal,
            fPlus,
            fMinus,
            rankMARCOS,
          } = ST;

          const f6 = (v) => {
            const num = Number(v);
            return Number.isFinite(num) ? Number(num.toFixed(6)) : "";
          };
          const rankByIdx = (rankArr, idx) => rankArr?.find((r) => r.idx === idx)?.rank ?? "";

          const sheets = [];
          const addSheet = (name, headers, rows) => {
            sheets.push({
              name: safeSheetName(name),
              aoa: [headers, ...rows].map((row) => row.map(cleanExportCell)),
            });
          };

          addSheet(
            "0_Matriks_Asli",
            ["Alternatif", ...criteria],
            [...matrix.map((r, i) => [alternatives[i], ...r.map(f6)]), ["Tipe", ...types]],
          );

          addSheet(
            "Entropy_1_Pij",
            ["Alternatif", ...criteria],
            pij.map((r, i) => [alternatives[i], ...r.map(f6)]),
          );

          addSheet(
            "Entropy_2_Bobot",
            ["Parameter", ...criteria, "Total"],
            [
              ["Entropy Ej", ...entropy.map(f6), f6(entropy.reduce((a, b) => a + b, 0))],
              ["Divergence dj", ...degDiv.map(f6), f6(degDiv.reduce((a, b) => a + b, 0))],
              ["Bobot Wj", ...weights.map(f6), f6(weights.reduce((a, b) => a + b, 0))],
            ],
          );

          addSheet(
            "GRA_1_Normalisasi",
            ["Alternatif", ...criteria],
            normGRA.map((r, i) => [alternatives[i], ...r.map(f6)]),
          );

          addSheet(
            "GRA_2_Delta",
            ["Alternatif", ...criteria],
            delta.map((r, i) => [alternatives[i], ...r.map(f6)]),
          );

          addSheet(
            "GRA_3_GRC",
            ["Alternatif", ...criteria],
            grc.map((r, i) => [alternatives[i], ...r.map(f6)]),
          );

          addSheet(
            "GRA_4_Ranking",
            ["Ranking", "Alternatif", "GRD"],
            [...rankGRA].sort((a, b) => a.rank - b.rank).map((r) => [r.rank, r.name, f6(r.val)]),
          );

          addSheet(
            "MEREC_1_Normalisasi",
            ["Alternatif", ...criteria],
            normMEREC.map((r, i) => [alternatives[i], ...r.map(f6)]),
          );

          addSheet(
            "MEREC_2_Si",
            ["Alternatif", "Si"],
            Si.map((s, i) => [alternatives[i], f6(s)]),
          );

          addSheet(
            "MEREC_3_Sij",
            ["Alternatif", ...criteria],
            Sij.map((r, i) => [alternatives[i], ...r.map(f6)]),
          );

          addSheet(
            "MEREC_4_Bobot",
            ["Parameter", ...criteria, "Total"],
            [
              ["Hj", ...Hj.map(f6), f6(Hj.reduce((a, b) => a + b, 0))],
              ["Wj", ...Wi_merec.map(f6), f6(Wi_merec.reduce((a, b) => a + b, 0))],
            ],
          );

          addSheet(
            "MARCOS_1_Ideal_AntiIdeal",
            ["Referensi", ...criteria],
            [
              ["AAI", ...antiIdeal.map(f6)],
              ["AI", ...ideal.map(f6)],
            ],
          );

          const marcosLabels = ["AAI", ...alternatives, "AI"];
          addSheet(
            "MARCOS_2_Normalisasi",
            ["Alternatif", ...criteria],
            normMARCOS.map((r, i) => [marcosLabels[i] ?? `R${i + 1}`, ...r.map(f6)]),
          );

          addSheet(
            "MARCOS_3_Ranking",
            ["Ranking", "Alternatif", "Ki+", "Ki-", "f(Ki+)", "f(Ki-)", "Ki Final"],
            [...rankMARCOS]
              .sort((a, b) => a.rank - b.rank)
              .map((r) => [
                r.rank,
                r.name,
                f6(kPlus[r.idx]),
                f6(kMinus[r.idx]),
                f6(fPlus[r.idx]),
                f6(fMinus[r.idx]),
                f6(r.val),
              ]),
          );

          addSheet(
            "REKAPITULASI",
            ["Alternatif", "GRD", "Rank GRA", "Ki MARCOS", "Rank MARCOS", "Avg Rank"],
            alternatives.map((name, i) => {
              const rG = rankByIdx(rankGRA, i);
              const rMa = rankByIdx(rankMARCOS, i);
              const vals = [rG, rMa].filter((v) => v !== "" && Number.isFinite(Number(v)));
              const avg = vals.length ? vals.reduce((a, b) => a + Number(b), 0) / vals.length : NaN;
              return [
                name,
                f6(grd[i]),
                rG,
                f6(kFinal[i]),
                rMa,
                Number.isFinite(avg) ? Number(avg.toFixed(2)) : "",
              ];
            }),
          );

          downloadWorkbook(sheets, "DSS_Entropy_GRA_MEREC_MARCOS");
          showNotif("✓ Export Excel berhasil", "green");
        } catch (err) {
          console.error(err);
          alert("Export gagal: " + err.message);
        }
      }

      function cleanExportCell(v) {
        if (v === null || v === undefined) return "";
        if (typeof v === "number") return Number.isFinite(v) ? v : "";
        const s = String(v);
        return s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "").trim();
      }

      function safeSheetName(name) {
        const cleaned = String(name || "Sheet")
          .replace(/[\\/?*\[\]:]/g, "_")
          .replace(/^'+|'+$/g, "")
          .slice(0, 31);
        return cleaned || "Sheet";
      }

      function downloadWorkbook(sheets, baseName) {
        // Prioritas: file .xlsx asli jika library SheetJS tersedia.
        // Fallback: SpreadsheetML .xls, tetap langsung terbuka di Microsoft Excel tanpa CDN/library tambahan.
        if (window.XLSX && XLSX.utils && XLSX.writeFile) {
          const wb = XLSX.utils.book_new();
          sheets.forEach((sh) => {
            const ws = XLSX.utils.aoa_to_sheet(sh.aoa);
            const colCount = Math.max(...sh.aoa.map((r) => r.length));
            ws["!cols"] = Array.from({ length: colCount }, (_, i) => ({ wch: i === 0 ? 22 : 14 }));
            XLSX.utils.book_append_sheet(wb, ws, sh.name);
          });
          XLSX.writeFile(wb, `${baseName}.xlsx`, { bookType: "xlsx" });
          return;
        }

        const xml = buildSpreadsheetML(sheets);
        downloadBlob(xml, `${baseName}.xls`, "application/vnd.ms-excel;charset=utf-8");
      }

      function buildSpreadsheetML(sheets) {
        const x = xmlEsc;
        const worksheets = sheets
          .map((sh) => {
            const rows = sh.aoa
              .map((row) => {
                const cells = row
                  .map((v) => {
                    const isNum = typeof v === "number" && Number.isFinite(v);
                    return `<Cell><Data ss:Type="${isNum ? "Number" : "String"}">${x(isNum ? v : String(v ?? ""))}</Data></Cell>`;
                  })
                  .join("");
                return `<Row>${cells}</Row>`;
              })
              .join("");
            return `<Worksheet ss:Name="${x(sh.name)}"><Table>${rows}</Table></Worksheet>`;
          })
          .join("");

        return (
          "<" +
          '?xml version="1.0"?>\n' +
          "<" +
          '?mso-application progid="Excel.Sheet"?>\n' +
          `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40"><DocumentProperties xmlns="urn:schemas-microsoft-com:office:office"><Author>DSS Calculator</Author></DocumentProperties><Styles><Style ss:ID="Default" ss:Name="Normal"><Alignment ss:Vertical="Center"/><Font ss:FontName="Calibri" ss:Size="11"/></Style></Styles>${worksheets}</Workbook>`
        );
      }

      function xmlEsc(v) {
        return String(v ?? "").replace(
          /[<>&'"]/g,
          (ch) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", "'": "&apos;", '"': "&quot;" })[ch],
        );
      }

      function downloadBlob(content, filename, mime) {
        const blob = new Blob([content], { type: mime });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = filename;
        a.style.display = "none";
        document.body.appendChild(a);
        a.click();
        setTimeout(() => {
          document.body.removeChild(a);
          URL.revokeObjectURL(url);
        }, 300);
      }

      // ═══════════════════════════════════════════════
      //  EXPORT TEMPLATE
      // ═══════════════════════════════════════════════
      function exportTemplate() {
        const sheets = [
          {
            name: "Template_Data",
            aoa: [
              ["", "C1", "C2", "C3", "C4", "C5"],
              ["Tipe", "benefit", "benefit", "cost", "benefit", "cost"],
              ["Alternatif A", 85, 90, 70, 80, 65],
              ["Alternatif B", 78, 85, 60, 90, 80],
              ["Alternatif C", 92, 80, 85, 70, 55],
              ["Alternatif D", 70, 95, 75, 85, 72],
              ["Alternatif E", 88, 70, 90, 75, 68],
            ],
          },
          {
            name: "Petunjuk",
            aoa: [
              ["PETUNJUK PENGGUNAAN"],
              [""],
              ["Baris 1", "Nama kriteria. Kolom A dikosongkan."],
              ["Baris 2", "Tipe kriteria: ketik benefit atau cost."],
              ["Baris 3+", "Kolom A = nama alternatif, Kolom B+ = nilai numerik."],
              ["Benefit", "Semakin besar nilainya, semakin baik."],
              ["Cost", "Semakin kecil nilainya, semakin baik."],
              ["Import", "Buka website → klik Import Excel → pilih file ini → klik Hitung."],
            ],
          },
        ];
        downloadWorkbook(sheets, "Template_DSS");
        showNotif("✓ Template Excel berhasil diunduh", "green");
      }

      // ═══════════════════════════════════════════════
      //  MISC
      // ═══════════════════════════════════════════════
      function switchMethod(m) {
        ["gra", "merec_marcos", "combined"].forEach((id) => {
          const btn = document.getElementById("btn-" + id);
          if (btn)
            btn.style.cssText =
              id === m
                ? `width:100%;text-align:left;padding:8px 12px;border-radius:8px;font-size:14px;transition:all .2s;background:${PALETTE.c3};color:${PALETTE.text};border:1px solid ${PALETTE.border};font-weight:700`
                : `width:100%;text-align:left;padding:8px 12px;border-radius:8px;font-size:14px;transition:all .2s;color:${PALETTE.muted};background:transparent;border:1px solid transparent`;
        });
        const show = {
          gra: ["block-gra"],
          merec_marcos: ["block-merec", "block-marcos"],
          combined: ["block-gra", "block-merec", "block-marcos", "block-recap"],
        };
        ["block-gra", "block-merec", "block-marcos", "block-recap"].forEach((id) => {
          const el = document.getElementById(id);
          if (el) el.style.display = (show[m] || []).includes(id) ? "block" : "none";
        });
        ST.activeMethod = m;
      }

      function resetAll() {
        ST = {
          alternatives: [],
          criteria: [],
          types: [],
          matrix: [],
          m: 0,
          n: 0,
          normGRA: [],
          pij: [],
          entropy: [],
          degDiv: [],
          weights: [],
          refSeq: [],
          delta: [],
          grc: [],
          grd: [],
          rankGRA: [],
          normMEREC: [],
          Wi_merec: [],
          Si: [],
          Sij: [],
          Hj: [],
          skorMEREC: [],
          rankMEREC: [],
          normMARCOS: [],
          ideal: [],
          antiIdeal: [],
          kPlus: [],
          kMinus: [],
          kFinal: [],
          fPlus: [],
          fMinus: [],
          rankMARCOS: [],
          s_ai: 0,
          s_aai: 0,
          Si_marcos: [],
          activeMethod: "gra",
        };
        document.getElementById("resultsContainer").innerHTML = "";
        document.getElementById("exportArea").classList.add("hidden");
        document.getElementById("matrixArea").innerHTML = "";
        buildMatrix();
      }

      function scrollToSection(id) {
        document.getElementById(id)?.scrollIntoView({ behavior: "smooth", block: "start" });
      }

      // Init
      buildMatrix();
      switchMethod("gra");
    </script>
  </body>
</html>
