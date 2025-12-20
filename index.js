
/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
*/

(() => {
  const FLOORS = [2, 3, 4, 5, 6, 7, 8];
  const LOCAL_STORAGE_KEY = 'schedule-app-state';
  const SUMMARY_TASKS = [
      {id: 'neua-fa', name: 'สรุป QC เหนือฝ้า'},
      {id: 'qc-ww', name: 'สรุป QC WW'},
      {id: 'qc-end', name: 'สรุป QC End'},
  ];
  
  let WEEKS = 0;
  let WEEK_HEADERS = [];
  let MONTH_HEADERS = [];
  const START_DATE = new Date(2025, 8, 15); // 15 Sep 2025

  const SUPABASE_URL = 'https://qunnmzlrsfsgaztqiexf.supabase.co';
  const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InF1bm5temxyc2ZzZ2F6dHFpZXhmIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjEwMzQ3NjgsImV4cCI6MjA3NjYxMDc2OH0.6uUMhDqaq1fGia91r5vTp990amvsiZ_6_eYFJlvgk3c';
  const SUPABASE_TABLE_ID = 'main_schedule';

  const supabaseClient = supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

  const mainTableBody = document.querySelector('#main-schedule-table tbody');
  const popover = document.getElementById('popover');
  const popoverInput = document.getElementById('popover-input');
  const popoverClear = document.getElementById('popover-clear');
  const popoverCancel = document.getElementById('popover-cancel');
  const statusIndicator = document.getElementById('status-indicator');
  const captureButton = document.getElementById('capture-btn');

  let state = {};
  let activeCell = null;

  function generateWeekData() {
    const thaiMonths = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
    const end = new Date(2026, 0, 31);
    
    WEEK_HEADERS = [];
    const current = new Date(START_DATE);

    while (current <= end) {
      const wStart = new Date(current);
      const wEnd = new Date(current);
      wEnd.setDate(wEnd.getDate() + 6);
      WEEK_HEADERS.push(`${wStart.getDate()}-${wEnd.getDate()}/${thaiMonths[wEnd.getMonth()]}`);
      current.setDate(current.getDate() + 7);
    }
    WEEKS = WEEK_HEADERS.length;
    MONTH_HEADERS = [];
    let lastMonth = '';
    for (const h of WEEK_HEADERS) {
        const m = h.split('/')[1];
        if (m !== lastMonth) { MONTH_HEADERS.push({ name: m, span: 1 }); lastMonth = m; }
        else { MONTH_HEADERS[MONTH_HEADERS.length - 1].span++; }
    }
    const today = new Date();
    const dDisplay = document.getElementById('current-date-display');
    if (dDisplay) dDisplay.textContent = `${today.getDate()} ${thaiMonths[today.getMonth()]} ${today.getFullYear() + 543}`;
  }

  document.addEventListener('DOMContentLoaded', async () => {
    generateWeekData();
    await initState();
    createMainTable();
    createSummaryTables();
    addEventListeners();
    renderAll();
  });

  function migrateState(data) {
    for (const floor of FLOORS) {
      if (!data[floor]) data[floor] = {};
      for (let week = 0; week < WEEKS; week++) {
        if (!data[floor][week]) data[floor][week] = { plan: { value: null, task: null }, actual: {} };
      }
    }
    return data;
  }

  async function initState() {
    try {
      updateStatus('Loading...', 'orange', false);
      const { data, error } = await supabaseClient.from('schedules').select('data').eq('id', SUPABASE_TABLE_ID).single();
      if (data && data.data) { state = migrateState(data.data); updateStatus('Online', 'green'); return; }
    } catch (e) { console.error(e); }
    loadFromLocalStorage();
  }

  function loadFromLocalStorage() {
    const saved = localStorage.getItem(LOCAL_STORAGE_KEY);
    if (saved) { state = migrateState(JSON.parse(saved)); } else {
      state = {};
      FLOORS.forEach(f => { state[f] = {}; for (let i = 0; i < WEEKS; i++) state[f][i] = { plan: { value: null, task: null }, actual: {} }; });
    }
    updateStatus('Offline', 'orange');
  }

  async function saveState() {
    localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(state));
    try {
      await supabaseClient.from('schedules').upsert({ id: SUPABASE_TABLE_ID, data: state });
      updateStatus('Online', 'green');
    } catch (e) { updateStatus('Sync Failed', 'red'); }
  }

  function createMainTable() {
    if (!mainTableBody) return;
    const thead = document.querySelector('#main-schedule-table thead');
    let h1 = `<tr class="header-main"><th rowspan="2" class="static-col-1">ชั้น</th><th colspan="2" class="th-neua-fa">เหนือผ้า</th><th colspan="2" class="th-qc-ww">QC WW</th><th colspan="2" class="th-qc-end">QC END</th><th rowspan="2" style="width:40px; border:none; background:transparent;"></th>`;
    MONTH_HEADERS.forEach(m => h1 += `<th colspan="${m.span}">${m.name}</th>`);
    h1 += `</tr>`;
    let h2 = `<tr class="header-sub"><th class="th-neua-fa">ทั้งหมด</th><th class="th-neua-fa">ส่ง</th><th class="th-qc-ww">ทั้งหมด</th><th class="th-qc-ww">ส่ง</th><th class="th-qc-end">ทั้งหมด</th><th class="th-qc-end">ส่ง</th>`;
    WEEK_HEADERS.forEach(w => h2 += `<th>${w.split('/')[0]}</th>`);
    h2 += `</tr>`;
    thead.innerHTML = h1 + h2;

    let tableHTML = '';
    FLOORS.forEach(floor => {
      tableHTML += `<tr>
        <td rowspan="2" class="static-col-1">${floor}</td>
        <td rowspan="2" data-floor="${floor}" data-task-total="neua-fa"></td>
        <td rowspan="2" data-floor="${floor}" data-task-sent="neua-fa"></td>
        <td rowspan="2" data-floor="${floor}" data-task-total="qc-ww"></td>
        <td rowspan="2" data-floor="${floor}" data-task-sent="qc-ww"></td>
        <td rowspan="2" data-floor="${floor}" data-task-total="qc-end"></td>
        <td rowspan="2" data-floor="${floor}" data-task-sent="qc-end"></td>
        <td class="static-col-2">Plan</td>
        ${Array.from({ length: WEEKS }).map((_, i) => `<td class="editable-cell" data-floor="${floor}" data-week="${i}" data-type="plan"></td>`).join('')}
      </tr>
      <tr>
        <td class="static-col-2">Actual</td>
        ${Array.from({ length: WEEKS }).map((_, i) => `<td class="editable-cell" data-floor="${floor}" data-week="${i}" data-type="actual"></td>`).join('')}
      </tr>`;
    });
    mainTableBody.innerHTML = tableHTML;
  }

  function createSummaryTables() {
    const container = document.querySelector('.summary-container');
    if (!container) return;
    container.innerHTML = '';
    SUMMARY_TASKS.forEach(taskInfo => {
      const card = document.createElement('div'); card.className = 'summary-card';
      const wrapper = document.createElement('div'); wrapper.className = 'table-wrapper';
      const table = document.createElement('table'); table.id = `summary-${taskInfo.id}`; table.className = 'summary-table';
      const colorClass = taskInfo.id === 'neua-fa' ? 'th-neua-fa' : taskInfo.id === 'qc-ww' ? 'th-qc-ww' : 'th-qc-end';
      const h = `<thead><tr><th colspan="${WEEKS + 1}" class="${colorClass}">${taskInfo.name}</th></tr><tr><th class="row-label"></th>${WEEK_HEADERS.map(w => `<th>${w}</th>`).join('')}</tr></thead><tbody>` + 
                ['Plan', 'Acc. Plan', 'Actual', 'Acc. Actual'].map(l => `<tr data-row-type="${l.toLowerCase().replace('. ', '-')}"><td class="row-label">${l}</td>${Array.from({ length: WEEKS }).map((_, i) => `<td data-week="${i}">0</td>`).join('')}</tr>`).join('') + `</tbody>`;
      table.innerHTML = h; wrapper.appendChild(table); card.appendChild(wrapper); container.appendChild(card);
    });
  }

  function addEventListeners() {
    mainTableBody.addEventListener('click', (e) => { const target = e.target.closest('.editable-cell'); if (target) showPopover(target); });
    popover.addEventListener('click', (e) => { const task = e.target.dataset.task; if (task) updateCell(task); });
    popoverClear.addEventListener('click', () => updateCell(null));
    popoverCancel.addEventListener('click', hidePopover);
    captureButton.addEventListener('click', captureImage);
    document.getElementById('export-excel-btn').addEventListener('click', () => alert('Excel export feature coming soon.'));
  }

  function showPopover(cell) {
    activeCell = cell; const { floor, week, type } = cell.dataset;
    popoverInput.value = type === 'plan' ? (state[floor][week].plan.value || '') : '';
    const rect = cell.getBoundingClientRect();
    popover.style.display = 'block'; 
    popover.style.top = `${window.scrollY + rect.bottom + 5}px`; 
    popover.style.left = `${Math.min(window.scrollX + rect.left, window.innerWidth - 300)}px`;
    popoverInput.focus(); popoverInput.select();
  }

  function hidePopover() { popover.style.display = 'none'; activeCell = null; }

  function updateCell(task) {
    if (!activeCell) return;
    const { floor, week, type } = activeCell.dataset;
    const val = parseInt(popoverInput.value, 10);
    if (task === null) {
      if (type === 'plan') state[floor][week].plan = { value: null, task: null }; else state[floor][week].actual = {};
    } else {
      if (type === 'plan') state[floor][week].plan = { value: isNaN(val) ? null : val, task };
      else { if (!isNaN(val) && val > 0) state[floor][week].actual[task] = val; else delete state[floor][week].actual[task]; }
    }
    renderAll(); saveState(); hidePopover();
  }

  function renderAll() {
    renderMainTable();
    calculateAndRenderSummaries();
  }

  function renderMainTable() {
    const globalTotals = { 'neua-fa': { p: 0, a: 0 }, 'qc-ww': { p: 0, a: 0 }, 'qc-end': { p: 0, a: 0 } };

    FLOORS.forEach(f => {
      const fT = { 'neua-fa': { p: 0, a: 0 }, 'qc-ww': { p: 0, a: 0 }, 'qc-end': { p: 0, a: 0 } };
      for (let w = 0; w < WEEKS; w++) {
        const d = state[f][w];
        
        // Render Plan
        const pc = mainTableBody.querySelector(`[data-floor="${f}"][data-week="${w}"][data-type="plan"]`);
        pc.innerHTML = ''; pc.className = 'editable-cell';
        if (d.plan.value !== null) {
          const planDiv = document.createElement('div');
          planDiv.textContent = d.plan.value;
          planDiv.className = d.plan.task ? `bg-${d.plan.task}-plan` : '';
          pc.appendChild(planDiv);
        }
        if (d.plan.task && d.plan.value) {
            fT[d.plan.task].p += d.plan.value;
            globalTotals[d.plan.task].p += d.plan.value;
        }

        // Render Actual
        const ac = mainTableBody.querySelector(`[data-floor="${f}"][data-week="${w}"][data-type="actual"]`);
        const ts = Object.keys(d.actual).filter(t => d.actual[t]);
        ac.innerHTML = ''; ac.className = 'editable-cell';
        
        if (ts.length === 1) { 
            const singleDiv = document.createElement('div');
            singleDiv.className = `actual-cell-single bg-${ts[0]}-actual`;
            singleDiv.textContent = d.actual[ts[0]];
            ac.appendChild(singleDiv);
        }
        else if (ts.length > 1) {
            const wrap = document.createElement('div');
            wrap.className = 'actual-cell-split-container';
            ts.forEach(t => { 
                const dv = document.createElement('div');
                dv.className = `actual-cell-split bg-${t}-actual`;
                dv.textContent = d.actual[t];
                wrap.appendChild(dv); 
            });
            ac.appendChild(wrap);
        }
        
        ts.forEach(t => {
            const val = d.actual[t] || 0;
            fT[t].a += val;
            globalTotals[t].a += val;
        });
      }
      SUMMARY_TASKS.forEach(t => {
        const totalEl = mainTableBody.querySelector(`[data-floor="${f}"][data-task-total="${t.id}"]`);
        const sentEl = mainTableBody.querySelector(`[data-floor="${f}"][data-task-sent="${t.id}"]`);
        if (totalEl) totalEl.textContent = fT[t.id].p || '';
        if (sentEl) sentEl.textContent = fT[t.id].a || '';
      });
    });

    SUMMARY_TASKS.forEach(t => {
        const pEl = document.getElementById(`total-${t.id}-plan`);
        const aEl = document.getElementById(`total-${t.id}-sent`);
        if (pEl) pEl.textContent = globalTotals[t.id].p || '0';
        if (aEl) aEl.textContent = globalTotals[t.id].a || '0';
    });
  }

  function calculateAndRenderSummaries() {
    const summaryData = { 
        'neua-fa': { p: Array(WEEKS).fill(0), a: Array(WEEKS).fill(0) }, 
        'qc-ww': { p: Array(WEEKS).fill(0), a: Array(WEEKS).fill(0) }, 
        'qc-end': { p: Array(WEEKS).fill(0), a: Array(WEEKS).fill(0) } 
    };
    
    const today = new Date();
    const msInWeek = 7 * 24 * 60 * 60 * 1000;
    const currentWeekIdx = Math.floor((today - START_DATE) / msInWeek);

    FLOORS.forEach(f => {
      for (let w = 0; w < WEEKS; w++) {
        const d = state[f][w];
        if (d && d.plan && d.plan.task && d.plan.value) {
            if (summaryData[d.plan.task]) summaryData[d.plan.task].p[w] += d.plan.value;
        }
        if (d && d.actual) {
            Object.keys(d.actual).forEach(t => { 
                if (summaryData[t]) summaryData[t].a[w] += d.actual[t] || 0; 
            });
        }
      }
    });

    SUMMARY_TASKS.forEach(task => {
      const table = document.getElementById(`summary-${task.id}`);
      if (!table) return;
      
      let accP = 0, accA = 0;
      for (let w = 0; w < WEEKS; w++) {
        const p = summaryData[task.id].p[w] || 0;
        const a = summaryData[task.id].a[w] || 0;
        
        accP += p;
        const isFuture = w > currentWeekIdx;
        if (!isFuture) {
            accA += a;
        }
        
        const pCell = table.querySelector(`[data-row-type="plan"] [data-week="${w}"]`);
        const apCell = table.querySelector(`[data-row-type="acc-plan"] [data-week="${w}"]`);
        const aCell = table.querySelector(`[data-row-type="actual"] [data-week="${w}"]`);
        const aaCell = table.querySelector(`[data-row-type="acc-actual"] [data-week="${w}"]`);
        
        if (pCell) pCell.textContent = p || '0';
        if (apCell) apCell.textContent = accP || '0';
        if (aCell) aCell.textContent = isFuture ? '' : (a || '0');
        if (aaCell) aaCell.textContent = isFuture ? '' : (accA || '0');
      }
    });
  }

  function updateStatus(msg, color, auto = true) {
      if(statusIndicator) { statusIndicator.textContent = msg; statusIndicator.style.color = color; }
      const dot = document.getElementById('sync-dot');
      if (dot) {
          if (color === 'green') dot.className = 'dot active';
          else if (color === 'orange') dot.className = 'dot';
          else dot.className = 'dot';
      }
  }

  async function captureImage() {
    const el = document.getElementById('capture-area');
    if (!el) return;
    
    const originalBackground = el.style.background;
    el.style.background = 'white';

    const canvas = await html2canvas(el, { 
      scale: 2,
      useCORS: true,
      backgroundColor: '#ffffff'
    });

    el.style.background = originalBackground;

    const link = document.createElement('a');
    link.download = `qc-report-${new Date().toISOString().slice(0,10)}.png`;
    link.href = canvas.toDataURL('image/png');
    link.click();
  }
})();
