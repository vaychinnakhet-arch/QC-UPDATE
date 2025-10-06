/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
*/

// Let TypeScript know that ExcelJS is available on the global scope
declare var ExcelJS: any;

// --- TYPE DEFINITIONS ---
type Task = 'neua-fa' | 'qc-ww' | 'qc-end';
type CellData = {
  value: number | null;
  task: Task | null;
};
type RowType = 'plan' | 'actual';
type AppState = Record<string, Record<string, {
  plan: CellData;
  actual: CellData;
}>>;

// --- CONSTANTS ---
const FLOORS = [2, 3, 4, 5, 6, 7, 8];
const WEEKS = 16;
const WEEK_HEADERS = ["15-21/ก.ย.", "22-28/ก.ย.", "29-5/ต.ค.", "6-12/ต.ค.", "13-19/ต.ค.", "20-26/ต.ค.", "27-2/พ.ย.", "3-9/พ.ย.", "10-16/พ.ย.", "17-23/พ.ย.", "24-30/พ.ย.", "1-7/ธ.ค.", "8-14/ธ.ค.", "15-21/ธ.ค.", "22-28/ธ.ค.", "29-30/ธ.ค."];
const LOCAL_STORAGE_KEY = 'schedule-app-state';

// --- DOM ELEMENTS ---
const mainTableBody = document.querySelector('#main-schedule-table tbody') as HTMLTableSectionElement;
const popover = document.getElementById('popover') as HTMLDivElement;
const popoverInput = document.getElementById('popover-input') as HTMLInputElement;
const popoverClear = document.getElementById('popover-clear') as HTMLButtonElement;
const popoverCancel = document.getElementById('popover-cancel') as HTMLButtonElement;
const statusIndicator = document.getElementById('status-indicator') as HTMLSpanElement;
const exportButton = document.getElementById('export-excel-btn') as HTMLButtonElement;

// --- STATE ---
let state: AppState = {};
let activeCell: HTMLElement | null = null;

// --- INITIALIZATION ---
document.addEventListener('DOMContentLoaded', () => {
  initState();
  createMainTable();
  createSummaryTables();
  addEventListeners();
  renderAll();
  updateStatus('Ready', 'green');
});

function initState() {
  const savedState = localStorage.getItem(LOCAL_STORAGE_KEY);
  if (savedState) {
    state = JSON.parse(savedState);
  } else {
    state = {};
    for (const floor of FLOORS) {
      state[floor] = {};
      for (let i = 0; i < WEEKS; i++) {
        state[floor][i] = {
          plan: { value: null, task: null },
          actual: { value: null, task: null },
        };
      }
    }
  }
}

function saveState() {
  localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(state));
  updateStatus('Saved', 'green');
}

// --- TABLE CREATION ---
function createMainTable() {
  if (!mainTableBody) return;
  let tableHTML = '';
  for (const floor of FLOORS) {
    tableHTML += `
      <tr>
        <td rowspan="2" class="static-col-1">${floor}</td>
        <td rowspan="2" data-floor="${floor}" data-task-total="neua-fa"></td>
        <td rowspan="2" data-floor="${floor}" data-task-sent="neua-fa"></td>
        <td rowspan="2" data-floor="${floor}" data-task-total="qc-ww"></td>
        <td rowspan="2" data-floor="${floor}" data-task-sent="qc-ww"></td>
        <td rowspan="2" data-floor="${floor}" data-task-total="qc-end"></td>
        <td rowspan="2" data-floor="${floor}" data-task-sent="qc-end"></td>
        <td class="row-label static-col-2">Plan</td>
        ${Array.from({ length: WEEKS }).map((_, i) => `<td class="editable-cell" data-floor="${floor}" data-week="${i}" data-type="plan"></td>`).join('')}
        <td rowspan="2"></td>
      </tr>
      <tr>
        <td class="row-label static-col-2">Actual</td>
        ${Array.from({ length: WEEKS }).map((_, i) => `<td class="editable-cell" data-floor="${floor}" data-week="${i}" data-type="actual"></td>`).join('')}
      </tr>
    `;
  }
  mainTableBody.innerHTML = tableHTML;
}

function createSummaryTables() {
    const summaryTasks: {id: Task, name: string}[] = [
        {id: 'neua-fa', name: 'สรุป QC เหนือฝ้า'},
        {id: 'qc-ww', name: 'สรุป QC WW'},
        {id: 'qc-end', name: 'สรุป QC End'},
    ];

    summaryTasks.forEach(taskInfo => {
        const table = document.getElementById(`summary-${taskInfo.id}`) as HTMLTableElement;
        if (!table) return;

        const headerTitleRow = `<tr><th colspan="${WEEKS + 1}">${taskInfo.name}</th></tr>`;
        const headerWeekRow = `
          <tr>
            <th class="row-label"></th>
            ${WEEK_HEADERS.map(h => `<th>${h}</th>`).join('')}
          </tr>
        `;
    
        const bodyRows = ['Plan', 'Acc. Plan', 'Actual', 'Acc. Actual'].map(label => `
          <tr data-row-type="${label.toLowerCase().replace('. ', '-')}">
            <td class="row-label">${label}</td>
            ${Array.from({ length: WEEKS }).map((_, i) => `<td data-week="${i}">0</td>`).join('')}
          </tr>
        `).join('');
        
        table.innerHTML = `<thead>${headerTitleRow}${headerWeekRow}</thead><tbody>${bodyRows}</tbody>`;
      });
}

// --- EVENT LISTENERS ---
function addEventListeners() {
  mainTableBody.addEventListener('click', (e) => {
    const target = e.target as HTMLElement;
    if (target.classList.contains('editable-cell')) {
      showPopover(target);
    }
  });

  popover.addEventListener('click', (e) => {
    e.stopPropagation();
    const target = e.target as HTMLElement;
    const task = target.dataset.task as Task;
    if (task) {
      updateCell(task);
    }
  });

  popoverClear.addEventListener('click', () => updateCell(null));
  popoverCancel.addEventListener('click', hidePopover);
  document.addEventListener('click', (e) => {
    if (popover.style.display === 'block' && !popover.contains(e.target as Node) && e.target !== activeCell) {
      hidePopover();
    }
  });
  exportButton.addEventListener('click', exportToExcel);
}


// --- POPOVER LOGIC ---
function showPopover(cell: HTMLElement) {
  activeCell = cell;
  const { floor, week, type } = cell.dataset;
  if (!floor || !week || !type) return;

  const cellData = state[floor][week][type as RowType];
  popoverInput.value = cellData.value?.toString() || '';
  
  const rect = cell.getBoundingClientRect();
  popover.style.display = 'block';
  popover.style.top = `${window.scrollY + rect.bottom}px`;
  popover.style.left = `${window.scrollX + rect.left}px`;
  popoverInput.focus();
}

function hidePopover() {
  popover.style.display = 'none';
  activeCell = null;
}

// --- DATA & RENDER LOGIC ---
function updateCell(task: Task | null) {
  if (!activeCell) return;
  const { floor, week, type } = activeCell.dataset;
  if (!floor || !week || !type) return;

  const value = popoverInput.value ? parseInt(popoverInput.value, 10) : null;
  state[floor][week][type as RowType] = { value, task };

  renderAll();
  saveState();
  hidePopover();
}

function renderAll() {
  renderMainTable();
  renderFloorTotals();
  calculateAndRenderSummaries();
}

function renderMainTable() {
  for (const floor of FLOORS) {
    for (let week = 0; week < WEEKS; week++) {
      const planCellData = state[floor][week].plan;
      const actualCellData = state[floor][week].actual;

      // Render Plan Cell
      const planCell = mainTableBody.querySelector(`[data-floor="${floor}"][data-week="${week}"][data-type="plan"]`) as HTMLElement;
      planCell.textContent = planCellData.value?.toString() || '';
      planCell.className = 'editable-cell'; // Reset classes
      if (planCellData.task) {
        planCell.classList.add(`bg-${planCellData.task}-plan`);
      }
      
      // Render Actual Cell
      const actualCell = mainTableBody.querySelector(`[data-floor="${floor}"][data-week="${week}"][data-type="actual"]`) as HTMLElement;
      actualCell.textContent = actualCellData.value?.toString() || '';
      actualCell.className = 'editable-cell'; // Reset classes

      // Color the actual cell based on its own task.
      if (actualCellData.task) {
        actualCell.classList.add(`bg-${actualCellData.task}-actual`);
      }
    }
  }
}

function renderFloorTotals() {
    for (const floor of FLOORS) {
        const floorTotals: Record<Task, { plan: number; actual: number }> = {
            'neua-fa': { plan: 0, actual: 0 },
            'qc-ww': { plan: 0, actual: 0 },
            'qc-end': { plan: 0, actual: 0 },
        };

        for (let week = 0; week < WEEKS; week++) {
            const planData = state[floor][week].plan;
            if (planData.task && planData.value) {
                floorTotals[planData.task].plan += planData.value;
            }
            const actualData = state[floor][week].actual;
            if (actualData.task && actualData.value) {
                floorTotals[actualData.task].actual += actualData.value;
            }
        }
        
        // 'neua-fa' total (merged)
        const neuaFaPlanTotal = floorTotals['neua-fa'].plan;
        const neuaFaTotalCell = mainTableBody.querySelector(`[data-floor="${floor}"][data-task-total="neua-fa"]`);
        if (neuaFaTotalCell) neuaFaTotalCell.textContent = neuaFaPlanTotal > 0 ? neuaFaPlanTotal.toString() : '';

        // 'neua-fa' sent
        const neuaFaActualSent = floorTotals['neua-fa'].actual;
        const neuaFaSentCell = mainTableBody.querySelector(`[data-floor="${floor}"][data-task-sent="neua-fa"]`);
        if (neuaFaSentCell) neuaFaSentCell.textContent = neuaFaActualSent > 0 ? neuaFaActualSent.toString() : '';

        // 'qc-ww' total and sent
        const qcWwPlanTotal = floorTotals['qc-ww'].plan;
        const qcWwActualSent = floorTotals['qc-ww'].actual;
        const qcWwTotalCell = mainTableBody.querySelector(`[data-floor="${floor}"][data-task-total="qc-ww"]`);
        if (qcWwTotalCell) qcWwTotalCell.textContent = qcWwPlanTotal > 0 ? qcWwPlanTotal.toString() : '';
        const qcWwSentCell = mainTableBody.querySelector(`[data-floor="${floor}"][data-task-sent="qc-ww"]`);
        if (qcWwSentCell) qcWwSentCell.textContent = qcWwActualSent > 0 ? qcWwActualSent.toString() : '';

        // 'qc-end' total and sent
        const qcEndPlanTotal = floorTotals['qc-end'].plan;
        const qcEndActualSent = floorTotals['qc-end'].actual;
        const qcEndTotalCell = mainTableBody.querySelector(`[data-floor="${floor}"][data-task-total="qc-end"]`);
        if (qcEndTotalCell) qcEndTotalCell.textContent = qcEndPlanTotal > 0 ? qcEndPlanTotal.toString() : '';
        const qcEndSentCell = mainTableBody.querySelector(`[data-floor="${floor}"][data-task-sent="qc-end"]`);
        if (qcEndSentCell) qcEndSentCell.textContent = qcEndActualSent > 0 ? qcEndActualSent.toString() : '';
    }
}


function getSummaryData() {
    const summaryData: Record<Task, Record<RowType, number[]>> = {
      'neua-fa': { plan: Array(WEEKS).fill(0), actual: Array(WEEKS).fill(0) },
      'qc-ww': { plan: Array(WEEKS).fill(0), actual: Array(WEEKS).fill(0) },
      'qc-end': { plan: Array(WEEKS).fill(0), actual: Array(WEEKS).fill(0) },
    };
  
    // Calculate weekly totals from state
    for (const floor of FLOORS) {
      for (let week = 0; week < WEEKS; week++) {
        for (const type of ['plan', 'actual'] as RowType[]) {
          const cellData = state[floor][week][type];
          if (cellData.task && cellData.value) {
            summaryData[cellData.task][type][week] += cellData.value;
          }
        }
      }
    }
    return summaryData;
}

function calculateAndRenderSummaries() {
  const summaryData = getSummaryData();

  // 1. Render summaries and calculate accumulated values
  for (const task of ['neua-fa', 'qc-ww', 'qc-end'] as Task[]) {
    const table = document.getElementById(`summary-${task}`) as HTMLTableElement;
    const lastEnteredWeek = getLastEnteredActualWeekIndex(task);
    let accPlan = 0;
    let accActual = 0;

    for (let week = 0; week < WEEKS; week++) {
      const planValue = summaryData[task].plan[week];
      const actualValue = summaryData[task].actual[week];
      accPlan += planValue;
      
      (table.querySelector(`[data-row-type="plan"] [data-week="${week}"]`) as HTMLElement).textContent = planValue.toString();
      (table.querySelector(`[data-row-type="actual"] [data-week="${week}"]`) as HTMLElement).textContent = actualValue.toString();
      (table.querySelector(`[data-row-type="acc-plan"] [data-week="${week}"]`) as HTMLElement).textContent = accPlan.toString();
      
      const accActualCell = table.querySelector(`[data-row-type="acc-actual"] [data-week="${week}"]`) as HTMLElement;

      if (lastEnteredWeek !== -1 && week > lastEnteredWeek) {
        // This is a future week for actuals
        accActualCell.textContent = '';
        accActualCell.classList.add('future-actual');
      } else {
        // This is a past or current week for actuals
        accActual += actualValue;
        accActualCell.textContent = accActual.toString();
        accActualCell.classList.remove('future-actual');
      }
    }
  }
    
  // 2. Calculate and render main table totals in the footer
  const totalNeuaFaPlan = summaryData['neua-fa'].plan.reduce((a, b) => a + b, 0);
  const totalNeuaFaActual = summaryData['neua-fa'].actual.reduce((a, b) => a + b, 0);
  const totalQcWwPlan = summaryData['qc-ww'].plan.reduce((a, b) => a + b, 0);
  const totalQcWwActual = summaryData['qc-ww'].actual.reduce((a, b) => a + b, 0);
  const totalQcEndPlan = summaryData['qc-end'].plan.reduce((a, b) => a + b, 0);
  const totalQcEndActual = summaryData['qc-end'].actual.reduce((a, b) => a + b, 0);
    
  document.getElementById('total-neua-fa-plan')!.textContent = totalNeuaFaPlan > 0 ? totalNeuaFaPlan.toString() : '';
  document.getElementById('total-neua-fa-sent')!.textContent = totalNeuaFaActual > 0 ? totalNeuaFaActual.toString() : '';
  document.getElementById('total-qc-ww-plan')!.textContent = totalQcWwPlan > 0 ? totalQcWwPlan.toString() : '';
  document.getElementById('total-qc-ww-sent')!.textContent = totalQcWwActual > 0 ? totalQcWwActual.toString() : '';
  document.getElementById('total-qc-end-plan')!.textContent = totalQcEndPlan > 0 ? totalQcEndPlan.toString() : '';
  document.getElementById('total-qc-end-sent')!.textContent = totalQcEndActual > 0 ? totalQcEndActual.toString() : '';
}

// --- UTILITY FUNCTIONS ---
function getLastEnteredActualWeekIndex(task: Task): number {
    let lastWeek = -1;
    for (let week = WEEKS - 1; week >= 0; week--) {
      for (const floor of FLOORS) {
        const cellData = state[floor][week].actual;
        if (cellData.task === task && cellData.value !== null && cellData.value > 0) {
          return week; // Found the last week with data for this task
        }
      }
    }
    return lastWeek; // No actual data found for this task
}

function updateStatus(message: string, color: string, autoClear: boolean = true) {
    if(statusIndicator) {
        statusIndicator.textContent = message;
        statusIndicator.style.color = color;
        if(autoClear) {
            setTimeout(() => {
                if(statusIndicator.textContent === message) {
                    updateStatus('Ready', 'green', false);
                }
            }, 3000);
        }
    }
}

// --- EXCEL EXPORT ---
async function exportToExcel() {
    updateStatus('Exporting...', 'orange', false);
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'VAY CHINNAKHET';
    workbook.created = new Date();
  
    // --- Define Styles ---
    const getFill = (color: string) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: color } });
    const getFont = (color: string, bold: boolean = false) => ({ color: { argb: color }, bold });
    const border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    const centerAlign = { vertical: 'middle', horizontal: 'center' };

    const styles: Record<string, any> = {
        'neua-fa-plan': { fill: getFill('fddca9'), font: getFont('FF212121', true), border, alignment: centerAlign },
        'neua-fa-actual': { fill: getFill('feefd8'), font: getFont('FF212121', true), border, alignment: centerAlign },
        'neua-fa-header': { fill: getFill('f79646'), font: getFont('FF212121', true), border, alignment: centerAlign },
        'qc-ww-plan': { fill: getFill('c9d8e8'), font: getFont('FF212121', true), border, alignment: centerAlign },
        'qc-ww-actual': { fill: getFill('dce6f1'), font: getFont('FF212121', true), border, alignment: centerAlign },
        'qc-ww-header': { fill: getFill('4f81bd'), font: getFont('FFFFFFFF', true), border, alignment: centerAlign },
        'qc-end-plan': { fill: getFill('e2eacb'), font: getFont('FF212121', true), border, alignment: centerAlign },
        'qc-end-actual': { fill: getFill('ebf1de'), font: getFont('FF212121', true), border, alignment: centerAlign },
        'qc-end-header': { fill: getFill('9bbb59'), font: getFont('FF212121', true), border, alignment: centerAlign },
        'future-actual': { fill: getFill('FFF5F5F5'), font: getFont('FFBDBDBD'), border, alignment: centerAlign },
        'default': { border, alignment: centerAlign },
        'header': { font: { bold: true }, border, alignment: centerAlign, fill: getFill('FFE0E0E0') },
        'row-label': { font: { bold: true }, border, alignment: { vertical: 'middle', horizontal: 'left', indent: 1 } },
        'footer': { font: { bold: true }, border, alignment: centerAlign, fill: getFill('FFFFFDE7')},
    };
    
    // --- MAIN SCHEDULE SHEET ---
    const mainSheet = workbook.addWorksheet('Schedule Update');

    mainSheet.columns = [
        { width: 5 }, { width: 8 }, { width: 8 }, { width: 8 },
        { width: 8 }, { width: 8 }, { width: 8 }, { width: 8 },
        ...Array(WEEKS).fill({ width: 8 }),
        { width: 10 }
    ];

    const headerRow1 = mainSheet.addRow([
        'ชั้น', 'เหนือผ้า', '', 'QC WW', '', 'QC End', '', '',
        'ก.ย.', '', 'ต.ค.', '', '', '', 'พ.ย.', '', '', '', '',
        'ธ.ค.', '', '', '', '', 'หมายเหตุ'
    ]);
    headerRow1.eachCell(c => c.style = styles.header!);
    
    const headerRow2 = mainSheet.addRow([
        '', 'ทั้งหมด', 'ส่ง', 'ทั้งหมด', 'ส่ง', 'ทั้งหมด', 'ส่ง', '',
        ...WEEK_HEADERS.map(h => h.split('/')[0]), ''
    ]);
    headerRow2.eachCell(c => c.style = styles.header!);
    
    mainSheet.mergeCells('A1:A2'); mainSheet.mergeCells('H1:H2'); mainSheet.mergeCells('Y1:Y2');
    mainSheet.mergeCells('B1:C1'); mainSheet.mergeCells('D1:E1'); mainSheet.mergeCells('F1:G1');
    mainSheet.mergeCells('I1:J1'); mainSheet.mergeCells('K1:N1'); mainSheet.mergeCells('O1:S1');
    mainSheet.mergeCells('T1:X1');

    ['B1', 'D1', 'F1'].forEach((cell, index) => {
      mainSheet.getCell(cell).style = styles[['neua-fa-header', 'qc-ww-header', 'qc-end-header'][index]]!;
    });
    ['B2', 'D2', 'F2', 'C2', 'E2', 'G2'].forEach((cell, index) => {
        mainSheet.getCell(cell).style = styles[['neua-fa-header', 'qc-ww-header', 'qc-end-header', 'neua-fa-header', 'qc-ww-header', 'qc-end-header'][index]]!;
    });


    // Data Rows
    FLOORS.forEach(floor => {
        const floorTotals = {
            'neua-fa': { plan: 0, actual: 0 }, 'qc-ww': { plan: 0, actual: 0 }, 'qc-end': { plan: 0, actual: 0 }
        };
        for (let week = 0; week < WEEKS; week++) {
            const planData = state[floor][week].plan;
            if (planData.task && planData.value) floorTotals[planData.task].plan += planData.value;
            const actualData = state[floor][week].actual;
            if (actualData.task && actualData.value) floorTotals[actualData.task].actual += actualData.value;
        }

        const planRow = mainSheet.addRow([
            floor,
            floorTotals['neua-fa'].plan > 0 ? floorTotals['neua-fa'].plan : '',
            floorTotals['neua-fa'].actual > 0 ? floorTotals['neua-fa'].actual : '',
            floorTotals['qc-ww'].plan > 0 ? floorTotals['qc-ww'].plan : '',
            floorTotals['qc-ww'].actual > 0 ? floorTotals['qc-ww'].actual : '',
            floorTotals['qc-end'].plan > 0 ? floorTotals['qc-end'].plan : '',
            floorTotals['qc-end'].actual > 0 ? floorTotals['qc-end'].actual : '',
            'Plan',
            ...Array.from({length: WEEKS}, (_, i) => state[floor][i].plan.value || ''),
            ''
        ]);
        const actualRow = mainSheet.addRow([
            '', // for merge
            '', // for merge
            '', // for merge
            '', // for merge
            '', // for merge
            '', // for merge
            '', // for merge
            'Actual',
            ...Array.from({length: WEEKS}, (_, i) => state[floor][i].actual.value || ''),
            ''
        ]);
        
        const startRowNum = planRow.number;
        const endRowNum = actualRow.number;

        // Merge cells
        mainSheet.mergeCells(`A${startRowNum}:A${endRowNum}`);
        mainSheet.mergeCells(`B${startRowNum}:B${endRowNum}`); // neua-fa total
        mainSheet.mergeCells(`C${startRowNum}:C${endRowNum}`); // neua-fa sent
        mainSheet.mergeCells(`D${startRowNum}:D${endRowNum}`); // qc-ww total
        mainSheet.mergeCells(`E${startRowNum}:E${endRowNum}`); // qc-ww sent
        mainSheet.mergeCells(`F${startRowNum}:F${endRowNum}`); // qc-end total
        mainSheet.mergeCells(`G${startRowNum}:G${endRowNum}`); // qc-end sent
        mainSheet.mergeCells(`Y${startRowNum}:Y${endRowNum}`); // Note column
        
        // --- Styling ---
        // Merged Floor Cell
        mainSheet.getCell(`A${startRowNum}`).style = styles.default!;

        // Style Labels
        planRow.getCell(8).style = styles['row-label']!;
        actualRow.getCell(8).style = styles['row-label']!;
        
        // Color task cells (using the lighter 'actual' color for all totals)
        mainSheet.getCell(`B${startRowNum}`).style = styles['neua-fa-actual']!; // neua-fa total (merged)
        mainSheet.getCell(`C${startRowNum}`).style = styles['neua-fa-actual']!; // neua-fa sent (merged)
        mainSheet.getCell(`D${startRowNum}`).style = styles['qc-ww-actual']!;   // qc-ww total (merged)
        mainSheet.getCell(`E${startRowNum}`).style = styles['qc-ww-actual']!;   // qc-ww sent (merged)
        mainSheet.getCell(`F${startRowNum}`).style = styles['qc-end-actual']!;   // qc-end total (merged)
        mainSheet.getCell(`G${startRowNum}`).style = styles['qc-end-actual']!;   // qc-end sent (merged)

        // Style Weekly Data Cells
        for(let week = 0; week < WEEKS; week++) {
            const col = 9 + week;
            const planTask = state[floor][week].plan.task;
            if(planTask) planRow.getCell(col).style = styles[`${planTask}-plan`]!;
            else planRow.getCell(col).style = styles.default!;

            const actualTask = state[floor][week].actual.task;
            if(actualTask) actualRow.getCell(col).style = styles[`${actualTask}-actual`]!;
            else actualRow.getCell(col).style = styles.default!;
        }
    });

    // Footer Row
    const summaryData = getSummaryData();
    const footerValues: (string|number)[] = ['รวม'];
    ['neua-fa', 'qc-ww', 'qc-end'].forEach(task => {
        const totalPlan = summaryData[task as Task].plan.reduce((a, b) => a + b, 0);
        const totalActual = summaryData[task as Task].actual.reduce((a, b) => a + b, 0);
        footerValues.push(totalPlan > 0 ? totalPlan.toString() : '');
        footerValues.push(totalActual > 0 ? totalActual.toString() : '');
    });
    footerValues.push(''); // for Plan/Actual label column
    const footerRow = mainSheet.addRow(footerValues);
    footerRow.eachCell(c => c.style = styles.footer!);

    // --- SUMMARY SHEETS ---
    const summaryTasks: {id: Task, name: string}[] = [
        {id: 'neua-fa', name: 'สรุป QC เหนือฝ้า'},
        {id: 'qc-ww', name: 'สรุป QC WW'},
        {id: 'qc-end', name: 'สรุป QC End'},
    ];
    
    summaryTasks.forEach(taskInfo => {
        const sheet = workbook.addWorksheet(taskInfo.name);
        sheet.columns = [{ width: 12 }, ...Array(WEEKS).fill({ width: 10 })];
        
        const headerRow = sheet.addRow([taskInfo.name]);
        headerRow.getCell(1).style = styles[`${taskInfo.id}-header`]!;
        headerRow.getCell(1).font = { ...styles[`${taskInfo.id}-header`]!.font, size: 14 };
        sheet.mergeCells(1, 1, 1, WEEKS + 1);
        
        const weekHeaderRow = sheet.addRow(['', ...WEEK_HEADERS]);
        weekHeaderRow.getCell(1).style = styles['row-label']!;
        weekHeaderRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            if (colNumber > 1) cell.style = styles.header!;
        });

        const rows = [
            { label: 'Plan', type: 'plan', acc: false },
            { label: 'Acc. Plan', type: 'plan', acc: true },
            { label: 'Actual', type: 'actual', acc: false },
            { label: 'Acc. Actual', type: 'actual', acc: true }
        ];

        let accPlan = 0;
        let accActual = 0;
        const lastEnteredWeek = getLastEnteredActualWeekIndex(taskInfo.id);

        rows.forEach(rowData => {
            const row = sheet.addRow([rowData.label]);
            row.getCell(1).style = styles['row-label']!;
            for(let week = 0; week < WEEKS; week++) {
                let value: number | string = 0;
                let style = styles.default;
                if(rowData.type === 'plan') {
                    const weeklyValue = summaryData[taskInfo.id].plan[week];
                    if(rowData.acc) {
                        accPlan += weeklyValue;
                        value = accPlan;
                    } else {
                        value = weeklyValue;
                    }
                } else { // actual
                    const weeklyValue = summaryData[taskInfo.id].actual[week];
                    if(rowData.acc) {
                        if(lastEnteredWeek !== -1 && week > lastEnteredWeek) {
                            value = '';
                            style = styles['future-actual'];
                        } else {
                            accActual += weeklyValue;
                            value = accActual;
                        }
                    } else {
                        value = weeklyValue;
                    }
                }
                
                const cell = row.getCell(week + 2);
                if (value !== 0 && value !== '') cell.value = value;
                else cell.value = null;
                cell.style = style!;
            }
        });
    });

    // --- DOWNLOAD ---
    try {
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'schedule_update.xlsx';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        updateStatus('Export Successful!', 'green');
    } catch (error) {
        console.error('Failed to export to Excel:', error);
        updateStatus('Export Failed!', 'red');
    }
}