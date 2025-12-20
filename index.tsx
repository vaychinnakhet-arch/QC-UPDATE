
/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
*/

declare var ExcelJS: any;
declare var html2canvas: any;
declare var supabase: any;

(() => {
  type Task = 'neua-fa' | 'qc-ww' | 'qc-end';
  type AppState = Record<string, Record<string, { plan: { value: number | null, task: Task | null }, actual: Partial<Record<Task, number>> }>>;

  const FLOORS = [2, 3, 4, 5, 6, 7, 8];
  const START_DATE = new Date(2025, 8, 15);
  const SUPABASE_URL = 'https://qunnmzlrsfsgaztqiexf.supabase.co';
  const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InF1bm5temxyc2ZzZ2F6dHFpZXhmIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjEwMzQ3NjgsImV4cCI6MjA3NjYxMDc2OH0.6uUMhDqaq1fGia91r5vTp990amvsiZ_6_eYFJlvgk3c';
  const TABLE_ID = 'main_schedule';

  const SUMMARY_INFO = [
    { id: 'neua-fa', name: 'สรุป QC เหนือฝ้า' },
    { id: 'qc-ww', name: 'สรุป QC WW' },
    { id: 'qc-end', name: 'สรุป QC End' }
  ];

  let WEEKS = 0;
  let WEEK_HEADERS: string[] = [];
  let MONTH_HEADERS: { name: string, span: number }[] = [];
  
  const supabaseClient = supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

  let state: AppState = {};
  let activeCell: HTMLElement | null = null;
  let isSaving = false;

  let mainBody: HTMLElement | null = null;
  let popover: HTMLElement | null = null;
  let popInput: HTMLInputElement | null = null;
  let syncDot: HTMLElement | null = null;
  let statusTxt: HTMLElement | null = null;

  function generateWeeks() {
    const thaiMonths = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
    const end = new Date(2026, 0, 31);
    const current = new Date(START_DATE);
    WEEK_HEADERS = [];
    while (current <= end) {
      const wS = new Date(current);
      const wE = new Date(current);
      wE.setDate(wE.getDate() + 6);
      WEEK_HEADERS.push(`${wS.getDate()}-${wE.getDate()}/${thaiMonths[wE.getMonth()]}`);
      current.setDate(current.getDate() + 7);
    }
    WEEKS = WEEK_HEADERS.length;
    MONTH_HEADERS = [];
    let lastM = '';
    WEEK_HEADERS.forEach(h => {
      const m = h.split('/')[1];
      if (m !== lastM) { MONTH_HEADERS.push({ name: m, span: 1 }); lastM = m; }
      else { MONTH_HEADERS[MONTH_HEADERS.length - 1].span++; }
    });
    const today = new Date();
    const dDisplay = document.getElementById('current-date-display');
    if (dDisplay) dDisplay.textContent = `${today.getDate()} ${thaiMonths[today.getMonth()]} ${today.getFullYear() + 543}`;
  }

  async function initState() {
    try {
      const { data, error } = await supabaseClient.from('schedules').select('data').eq('id', TABLE_ID).single();
      if (!error && data) state = data.data;
      else FLOORS.forEach(f => { state[f] = {}; for (let w = 0; w < WEEKS; w++) state[f][w] = { plan: { value: null, task: null }, actual: {} }; });
    } catch (e) { console.error(e); }
  }

  function setupRealtime() {
    supabaseClient.channel('db-sync').on('postgres_changes', { event: 'UPDATE', schema: 'public', table: 'schedules', filter: `id=eq.${TABLE_ID}` }, (p: any) => {
      if (!isSaving && p.new?.data) { state = p.new.data; renderAll(); pulse(); }
    }).subscribe();
  }

  function pulse() { if(syncDot) { syncDot.classList.add('pulse'); setTimeout(() => syncDot?.classList.remove('pulse'), 500); } }

  async function save() {
    isSaving = true;
    if (statusTxt) statusTxt.textContent = 'Saving...';
    try {
      const { error } = await supabaseClient.from('schedules').upsert({ id: TABLE_ID, data: state, updated_at: new Date().toISOString() });
      if (!error) { if (statusTxt) statusTxt.textContent = 'Online'; pulse(); }
    } catch (e) { if (statusTxt) statusTxt.textContent = 'Offline'; }
    isSaving = false;
  }

  function createMainTable() {
    const thead = document.querySelector('#main-schedule-table thead');
    if (!thead || !mainBody) return;
    let h1 = `<tr><th rowspan="2" class="static-col-1">ชั้น</th><th colspan="2" class="th-neua-fa">เหนือผ้า</th><th colspan="2" class="th-qc-ww">QC WW</th><th colspan="2" class="th-qc-end">QC End</th><th rowspan="2"></th>`;
    MONTH_HEADERS.forEach(m => h1 += `<th colspan="${m.span}">${m.name}</th>`);
    h1 += `</tr>`;
    let h2 = `<tr><th class="th-neua-fa">ทั้งหมด</th><th class="th-neua-fa">ส่ง</th><th class="th-qc-ww">ทั้งหมด</th><th class="th-qc-ww">ส่ง</th><th class="th-qc-end">ทั้งหมด</th><th class="th-qc-end">ส่ง</th>`;
    WEEK_HEADERS.forEach(w => h2 += `<th>${w.split('-')[0]}</th>`);
    h2 += `</tr>`;
    thead.innerHTML = h1 + h2;

    let bHtml = '';
    FLOORS.forEach(f => {
      bHtml += `<tr><td rowspan="2" class="static-col-1">${f}</td><td rowspan="2" data-total="neua-fa" data-f="${f}" class="floor-total"></td><td rowspan="2" data-sent="neua-fa" data-f="${f}" class="floor-total"></td><td rowspan="2" data-total="qc-ww" data-f="${f}" class="floor-total"></td><td rowspan="2" data-sent="qc-ww" data-f="${f}" class="floor-total"></td><td rowspan="2" data-total="qc-end" data-f="${f}" class="floor-total"></td><td rowspan="2" data-sent="qc-end" data-f="${f}" class="floor-total"></td><td class="static-col-2">Plan</td>${Array.from({ length: WEEKS }).map((_, i) => `<td class="editable-cell" data-f="${f}" data-w="${i}" data-t="plan"></td>`).join('')}</tr>`;
      bHtml += `<tr><td class="static-col-2">Actual</td>${Array.from({ length: WEEKS }).map((_, i) => `<td class="editable-cell" data-f="${f}" data-w="${i}" data-t="actual"></td>`).join('')}</tr>`;
    });
    mainBody.innerHTML = bHtml;
  }

  function createSummaryTables() {
    const container = document.querySelector('.summary-section');
    if (!container) return;
    container.innerHTML = '';
    SUMMARY_INFO.forEach(task => {
      const card = document.createElement('div'); card.className = 'summary-card';
      const wrapper = document.createElement('div'); wrapper.className = 'table-wrapper';
      const table = document.createElement('table'); table.id = `summary-${task.id}`;
      let h = `<thead><tr><th colspan="${WEEKS + 1}" class="th-${task.id.includes('neua') ? 'neua-fa' : task.id.includes('ww') ? 'qc-ww' : 'qc-end'}">${task.name}</th></tr>`;
      h += `<tr><th class="row-label"></th>${WEEK_HEADERS.map(w => `<th>${w}</th>`).join('')}</tr></thead>`;
      h += `<tbody>` + ['Plan', 'Acc. Plan', 'Actual', 'Acc. Actual'].map(r => `<tr data-row="${r.toLowerCase().replace('. ', '-')}"><td class="row-label">${r}</td>${Array.from({ length: WEEKS }).map((_, i) => `<td data-w="${i}">0</td>`).join('')}</tr>`).join('') + `</tbody>`;
      table.innerHTML = h; wrapper.appendChild(table); card.appendChild(wrapper); container.appendChild(card);
    });
  }

  function renderAll() {
    if (!mainBody) return;
    const sumData: any = { 'neua-fa': { p: Array(WEEKS).fill(0), a: Array(WEEKS).fill(0) }, 'qc-ww': { p: Array(WEEKS).fill(0), a: Array(WEEKS).fill(0) }, 'qc-end': { p: Array(WEEKS).fill(0), a: Array(WEEKS).fill(0) } };
    const globalTotals: any = { 'neua-fa': { p: 0, a: 0 }, 'qc-ww': { p: 0, a: 0 }, 'qc-end': { p: 0, a: 0 } };

    const today = new Date();
    const msInWeek = 7 * 24 * 60 * 60 * 1000;
    const currentWeekIdx = Math.floor((today.getTime() - START_DATE.getTime()) / msInWeek);

    FLOORS.forEach(f => {
      const fT: any = { 'neua-fa': { p: 0, a: 0 }, 'qc-ww': { p: 0, a: 0 }, 'qc-end': { p: 0, a: 0 } };
      for (let w = 0; w < WEEKS; w++) {
        const d = state[f][w];
        const pc = mainBody!.querySelector(`[data-f="${f}"][data-w="${w}"][data-t="plan"]`) as HTMLElement;
        if (pc) {
          pc.textContent = d.plan.value?.toString() || '';
          pc.className = 'editable-cell' + (d.plan.task ? ` bg-${d.plan.task}-plan` : '');
        }
        if (d.plan.task && d.plan.value) { 
            sumData[d.plan.task].p[w] += d.plan.value; 
            fT[d.plan.task].p += d.plan.value; 
            globalTotals[d.plan.task].p += d.plan.value;
        }

        const ac = mainBody!.querySelector(`[data-f="${f}"][data-w="${w}"][data-t="actual"]`) as HTMLElement;
        if (ac) {
          const tasks = Object.keys(d.actual) as Task[];
          ac.innerHTML = ''; ac.className = 'editable-cell';
          if (tasks.length === 1) { ac.textContent = d.actual[tasks[0]]?.toString() || ''; ac.classList.add(`bg-${tasks[0]}-actual`); }
          else if (tasks.length > 1) {
            const wrap = document.createElement('div'); wrap.style.display='flex'; wrap.style.height='100%';
            tasks.forEach(t => { const dv = document.createElement('div'); dv.className = `bg-${t}-actual`; dv.style.flex = '1'; dv.textContent = d.actual[t]!.toString(); wrap.appendChild(dv); });
            ac.appendChild(wrap); ac.style.padding = '0';
          }
          tasks.forEach(t => { 
              const val = d.actual[t] || 0;
              sumData[t].a[w] += val; 
              fT[t].a += val; 
              globalTotals[t].a += val;
          });
        }
      }
      SUMMARY_INFO.forEach(t => {
        const totalEl = mainBody!.querySelector(`[data-total="${t.id}"][data-f="${f}"]`);
        const sentEl = mainBody!.querySelector(`[data-sent="${t.id}"][data-f="${f}"]`);
        if (totalEl) totalEl.textContent = fT[t.id].p ? fT[t.id].p.toString() : '';
        if (sentEl) sentEl.textContent = fT[t.id].a ? fT[t.id].a.toString() : '';
      });
    });

    SUMMARY_INFO.forEach(task => {
      const table = document.getElementById(`summary-${task.id}`);
      if (!table) return;
      let accP = 0, accA = 0;
      for (let w = 0; w < WEEKS; w++) {
        const p = sumData[task.id].p[w] || 0;
        const a = sumData[task.id].a[w] || 0;
        const isFuture = w > currentWeekIdx;
        
        accP += p;
        if (!isFuture) {
            accA += a;
        }
        
        const pEl = table.querySelector(`[data-row="plan"] [data-w="${w}"]`);
        const apEl = table.querySelector(`[data-row="acc-plan"] [data-w="${w}"]`);
        const aEl = table.querySelector(`[data-row="actual"] [data-w="${w}"]`);
        const aaEl = table.querySelector(`[data-row="acc-actual"] [data-w="${w}"]`);
        
        if (pEl) pEl.textContent = p.toString();
        if (apEl) apEl.textContent = accP.toString();
        if (aEl) aEl.textContent = isFuture ? '' : a.toString();
        if (aaEl) aaEl.textContent = isFuture ? '' : accA.toString();
      }
    });

    SUMMARY_INFO.forEach(t => {
      const pEl = document.getElementById(`total-${t.id}-plan`);
      const aEl = document.getElementById(`total-${t.id}-sent`);
      if (pEl) pEl.textContent = globalTotals[t.id].p.toString();
      if (aEl) aEl.textContent = globalTotals[t.id].a.toString();
    });
  }

  function addEvents() {
    if (!mainBody || !popover || !popInput) return;

    mainBody.addEventListener('click', e => {
      const c = (e.target as HTMLElement).closest('.editable-cell') as HTMLElement;
      if (c) {
        activeCell = c; const { f, w, t } = c.dataset;
        popInput!.value = t === 'plan' ? (state[f!][w!].plan.value?.toString() || '') : '';
        const rect = c.getBoundingClientRect();
        popover!.style.display = 'block'; popover!.style.top = `${window.scrollY + rect.bottom + 5}px`; popover!.style.left = `${Math.min(window.scrollX + rect.left, window.innerWidth - 220)}px`;
        popInput!.focus(); popInput!.select();
      }
    });

    popover.addEventListener('click', e => {
      const task = (e.target as HTMLElement).dataset.task as Task;
      if (task && activeCell) {
        const val = parseInt(popInput!.value, 10); const final = isNaN(val) ? null : val;
        const { f, w, t } = activeCell!.dataset;
        if (t === 'plan') state[f!][w!].plan = { value: final, task };
        else { if (final === null || final <= 0) delete state[f!][w!].actual[task]; else state[f!][w!].actual[task] = final; }
        renderAll(); save(); popover!.style.display = 'none';
      }
    });

    document.getElementById('popover-clear')?.addEventListener('click', () => {
      if (!activeCell) return;
      const { f, w, t } = activeCell.dataset;
      if (t === 'plan') state[f!][w!].plan = { value: null, task: null }; else state[f!][w!].actual = {};
      renderAll(); save(); popover!.style.display = 'none';
    });
    document.getElementById('popover-cancel')?.addEventListener('click', () => { if(popover) popover.style.display = 'none'; });
    
    document.getElementById('capture-btn')?.addEventListener('click', async () => {
        const el = document.querySelector('.container');
        if (!el) return;
        const canvas = await html2canvas(el, { scale: 2 });
        const link = document.createElement('a'); link.download = `qc-update-${Date.now()}.png`; link.href = canvas.toDataURL(); link.click();
    });

    document.getElementById('export-excel-btn')?.addEventListener('click', exportToExcel);
    document.getElementById('export-json-btn')?.addEventListener('click', () => {
        const blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' });
        const link = document.createElement('a'); link.href = URL.createObjectURL(blob); link.download = `qc-data-${Date.now()}.json`; link.click();
    });
    document.getElementById('import-json-btn')?.addEventListener('click', () => document.getElementById('import-json-input')?.click());
    document.getElementById('import-json-input')?.addEventListener('change', (e: any) => {
        const file = e.target.files?.[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (ev) => {
            try {
                state = JSON.parse(ev.target?.result as string);
                renderAll(); save();
                if(statusTxt) statusTxt.textContent = 'Imported';
            } catch (err) { alert('Invalid file'); }
        };
        reader.readAsText(file);
    });
  }

  async function exportToExcel() {
    if (!ExcelJS) return;
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('QC Update');
    sheet.addRow(['QC Project Update Dashboard']);
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a'); link.href = URL.createObjectURL(blob); link.download = 'qc_update.xlsx'; link.click();
  }

  document.addEventListener('DOMContentLoaded', async () => {
    mainBody = document.querySelector('#main-schedule-table tbody');
    popover = document.getElementById('popover');
    popInput = document.getElementById('popover-input') as HTMLInputElement;
    syncDot = document.getElementById('sync-dot');
    statusTxt = document.getElementById('status-indicator');

    generateWeeks();
    createMainTable();
    createSummaryTables();
    await initState();
    setupRealtime();
    addEvents();
    renderAll();
    if(syncDot) syncDot.classList.add('active');
  });
})();
