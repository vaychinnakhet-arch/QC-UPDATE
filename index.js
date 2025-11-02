/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
*/

// FIX: Wrapped the code in an IIFE to prevent global scope pollution.
(() => {
  // --- CONSTANTS ---
  const FLOORS = [2, 3, 4, 5, 6, 7, 8];
  const WEEKS = 16;
  const WEEK_HEADERS = ["15-21/ก.ย.", "22-28/ก.ย.", "29-5/ต.ค.", "6-12/ต.ค.", "13-19/ต.ค.", "20-26/ต.ค.", "27-2/พ.ย.", "3-9/พ.ย.", "10-16/พ.ย.", "17-23/พ.ย.", "24-30/พ.ย.", "1-7/ธ.ค.", "8-14/ธ.ค.", "15-21/ธ.ค.", "22-28/ธ.ค.", "29-30/ธ.ค."];
  const LOCAL_STORAGE_KEY = 'schedule-app-state';

  // --- SUPABASE CONSTANTS (USER-CONFIGURABLE) ---
  // IMPORTANT: Replace with your Supabase project URL and anon key.
  const SUPABASE_URL = 'https://qunnmzlrsfsgaztqiexf.supabase.co'; // <-- REPLACE
  const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InF1bm5temxyc2ZzZ2F6dHFpZXhmIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjEwMzQ3NjgsImV4cCI6MjA3NjYxMDc2OH0.6uUMhDqaq1fGia91r5vTp990amvsiZ_6_eYFJlvgk3c'; // <-- REPLACE
  const SUPABASE_TABLE_ID = 'main_schedule';

  // --- Supabase Client ---
  const supabaseClient = SUPABASE_URL.includes('your-project-id') ? null : supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

  // --- DOM ELEMENTS ---
  const mainTableBody = document.querySelector('#main-schedule-table tbody');
  const popover = document.getElementById('popover');
  const popoverInput = document.getElementById('popover-input');
  const popoverClear = document.getElementById('popover-clear');
  const popoverCancel = document.getElementById('popover-cancel');
  const statusIndicator = document.getElementById('status-indicator');
  const exportButton = document.getElementById('export-excel-btn');
  const exportJsonButton = document.getElementById('export-json-btn');
  const importJsonButton = document.getElementById('import-json-btn');
  const importJsonInput = document.getElementById('import-json-input');
  const captureButton = document.getElementById('capture-btn');

  // --- STATE ---
  let state = {};
  let activeCell = null;

  // --- INITIALIZATION ---
  document.addEventListener('DOMContentLoaded', async () => {
    await initState();
    createMainTable();
    createSummaryTables();
    addEventListeners();
    renderAll();
  });

  function migrateState(data) {
    // Migration logic for old data structure
    for (const floor of FLOORS) {
      if (!data[floor]) continue;
      for (let week = 0; week < WEEKS; week++) {
        if (!data[floor][week]) continue;
        const actualCell = data[floor][week].actual;
        // Check for old format: { task: '...', value: ... }
        if (actualCell && actualCell.task) {
          const { task, value } = actualCell;
          data[floor][week].actual = {};
          if (task && value) {
            data[floor][week].actual[task] = value;
          }
        } else if (!actualCell || typeof actualCell !== 'object') {
          // Ensure it's an object if it doesn't exist or is wrong type
          data[floor][week].actual = {};
        }
      }
    }
    return data;
  }

  async function initState() {
    // If Supabase is not configured, show a warning and load from local storage
    if (!supabaseClient) {
      console.warn('Supabase not configured. Loading from local storage.');
      const warningEl = document.getElementById('supabase-warning');
      if (warningEl) warningEl.style.display = 'block';
      updateStatus('Supabase not configured', 'orange', false);
      loadFromLocalStorage();
      return;
    }

    // Try loading from Supabase first
    try {
      updateStatus('Loading from cloud...', 'orange', false);
      const { data, error } = await supabaseClient
          .from('schedules')
          .select('data')
          .eq('id', SUPABASE_TABLE_ID)
          .single();

      if (error && error.code !== 'PGRST116') { // PGRST116: "The result contains 0 rows"
          throw error;
      }

      if (data && data.data) {
          state = migrateState(data.data);
          updateStatus('Loaded from cloud', 'green');
          return;
      }
    } catch (error) {
      console.error('Failed to load from Supabase:', error);
      updateStatus('Cloud load failed', 'red', false);
    }
    
    // Fallback to local storage if Supabase fails or has no data
    console.log('No cloud data found. Falling back to local storage.');
    loadFromLocalStorage();
  }

  function loadFromLocalStorage() {
    const savedState = localStorage.getItem(LOCAL_STORAGE_KEY);
    if (savedState) {
      const parsedState = JSON.parse(savedState);
      state = migrateState(parsedState);
      updateStatus('Loaded locally', 'green');
    } else {
      // Initialize fresh state if nothing is saved
      state = {};
      for (const floor of FLOORS) {
        state[floor] = {};
        for (let i = 0; i < WEEKS; i++) {
          state[floor][i] = {
            plan: { value: null, task: null },
            actual: {},
          };
        }
      }
      updateStatus('Ready', 'green');
    }
  }

  async function saveState() {
    // Always save to local storage for offline access and speed
    localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(state));
    updateStatus('Saved locally', 'green');
    
    // If Supabase is not configured, stop here
    if (!supabaseClient) {
      return;
    }

    // Attempt to save to Supabase
    try {
      updateStatus('Saving to cloud...', 'orange', false);
      const { error } = await supabaseClient
        .from('schedules')
        .upsert({ id: SUPABASE_TABLE_ID, data: state });
      
      if (error) {
          throw error;
      }
      updateStatus('Saved to cloud', 'green');
    } catch (error) {
      console.error('Failed to save to Supabase:', error);
      updateStatus('Cloud save failed!', 'red');
    }
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
      const summaryTasks = [
          {id: 'neua-fa', name: 'สรุป QC เหนือฝ้า'},
          {id: 'qc-ww', name: 'สรุป QC WW'},
          {id: 'qc-end', name: 'สรุป QC End'},
      ];

      summaryTasks.forEach(taskInfo => {
          const table = document.getElementById(`summary-${taskInfo.id}`);
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
      const target = e.target;
      if (target.classList.contains('editable-cell')) {
        showPopover(target);
      } else if (target.parentElement?.classList.contains('editable-cell')) {
        // Handle clicks on inner divs of split cells
        showPopover(target.parentElement);
      }
    });

    popover.addEventListener('click', (e) => {
      e.stopPropagation();
      const target = e.target;
      const task = target.dataset.task;
      if (task) {
        updateCell(task);
      }
    });

    popoverClear.addEventListener('click', () => updateCell(null));
    popoverCancel.addEventListener('click', hidePopover);
    document.addEventListener('click', (e) => {
      if (popover.style.display === 'block' && !popover.contains(e.target) && e.target !== activeCell) {
        hidePopover();
      }
    });
    exportButton.addEventListener('click', exportToExcel);
    exportJsonButton.addEventListener('click', exportStateToJson);
    importJsonButton.addEventListener('click', () => importJsonInput.click());
    importJsonInput.addEventListener('change', importStateFromJson);
    captureButton.addEventListener('click', captureImage);
  }


  // --- POPOVER LOGIC ---
  function showPopover(cell) {
    activeCell = cell;
    const { floor, week, type } = cell.dataset;
    if (!floor || !week || !type) return;

    if (type === 'plan') {
      const cellData = state[floor][week].plan;
      popoverInput.value = cellData.value?.toString() || '';
    } else {
      // For actual cells, don't pre-fill. User types a value and picks a task to add/update.
      popoverInput.value = '';
    }
  
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
  function updateCell(task) {
    if (!activeCell) return;
    const { floor, week, type } = activeCell.dataset;
    if (!floor || !week || !type) return;

    // If task is null, it's a 'clear' operation.
    if (task === null) {
        if (type === 'plan') {
            state[floor][week].plan = { value: null, task: null };
        } else {
            state[floor][week].actual = {}; // Clear all actual tasks
        }
    } else {
        const value = popoverInput.value ? parseInt(popoverInput.value, 10) : null;
        if (type === 'plan') {
            state[floor][week].plan = { value, task };
        } else { // 'actual'
            if (!state[floor][week].actual) {
                state[floor][week].actual = {};
            }
            if (value === null || value <= 0) {
                // If value is empty or zero, remove that task
                delete state[floor][week].actual[task];
            } else {
                // Add or update the task with the new value
                state[floor][week].actual[task] = value;
            }
        }
    }

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

        // Render Plan Cell
        const planCell = mainTableBody.querySelector(`[data-floor="${floor}"][data-week="${week}"][data-type="plan"]`);
        planCell.textContent = planCellData.value?.toString() || '';
        planCell.className = 'editable-cell'; // Reset classes
        if (planCellData.task) {
          planCell.classList.add(`bg-${planCellData.task}-plan`);
        }
      
        // Render Actual Cell
        const actualCell = mainTableBody.querySelector(`[data-floor="${floor}"][data-week="${week}"][data-type="actual"]`);
        const actualCellData = state[floor][week].actual;
        const tasks = actualCellData ? Object.keys(actualCellData).filter(t => actualCellData[t]) : [];

        // Cleanup cell before rendering
        actualCell.innerHTML = '';
        actualCell.className = 'editable-cell';
        actualCell.style.padding = '';

        if (tasks.length === 0) {
            // Cell is empty, do nothing
        } else if (tasks.length === 1) {
            const task = tasks[0];
            actualCell.textContent = actualCellData[task]?.toString() || '';
            actualCell.classList.add(`bg-${task}-actual`);
        } else { // 2 or more tasks
            actualCell.style.padding = '0';
            const container = document.createElement('div');
            container.className = 'actual-cell-split-container';

            // Display all tasks
            tasks.forEach(task => {
                const value = actualCellData[task];
                const splitDiv = document.createElement('div');
                splitDiv.className = `actual-cell-split bg-${task}-actual`;
                splitDiv.textContent = value?.toString() || '';
                container.appendChild(splitDiv);
            });
            actualCell.appendChild(container);
        }
      }
    }
  }

  function renderFloorTotals() {
      for (const floor of FLOORS) {
          const floorTotals = {
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
              if (actualData) {
                  for (const task of Object.keys(actualData)) {
                      const value = actualData[task];
                      if (value) {
                          floorTotals[task].actual += value;
                      }
                  }
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
      const summaryData = {
        'neua-fa': { plan: Array(WEEKS).fill(0), actual: Array(WEEKS).fill(0) },
        'qc-ww': { plan: Array(WEEKS).fill(0), actual: Array(WEEKS).fill(0) },
        'qc-end': { plan: Array(WEEKS).fill(0), actual: Array(WEEKS).fill(0) },
      };
  
      // Calculate weekly totals from state
      for (const floor of FLOORS) {
        for (let week = 0; week < WEEKS; week++) {
          // Plan
          const planData = state[floor][week].plan;
          if (planData.task && planData.value) {
            summaryData[planData.task].plan[week] += planData.value;
          }
          // Actual
          const actualData = state[floor][week].actual;
          if (actualData) {
            for (const task of Object.keys(actualData)) {
              const value = actualData[task];
              if (value) {
                summaryData[task].actual[week] += value;
              }
            }
          }
        }
      }
      return summaryData;
  }

  function calculateAndRenderSummaries() {
    const summaryData = getSummaryData();

    // 1. Render summaries and calculate accumulated values
    for (const task of ['neua-fa', 'qc-ww', 'qc-end']) {
        const table = document.getElementById(`summary-${task}`);
        
        const planRange = getTaskWeekRange(task, 'plan');
        const actualRange = getTaskWeekRange(task, 'actual');

        let accPlan = 0;
        let accActual = 0;

        for (let week = 0; week < WEEKS; week++) {
            const planValue = summaryData[task].plan[week];
            const actualValue = summaryData[task].actual[week];
            
            // Always calculate accumulated values
            accPlan += planValue;
            accActual += actualValue;

            const planCell = table.querySelector(`[data-row-type="plan"] [data-week="${week}"]`);
            const accPlanCell = table.querySelector(`[data-row-type="acc-plan"] [data-week="${week}"]`);
            const actualCell = table.querySelector(`[data-row-type="actual"] [data-week="${week}"]`);
            const accActualCell = table.querySelector(`[data-row-type="acc-actual"] [data-week="${week}"]`);
            
            const allCells = [planCell, accPlanCell, actualCell, accActualCell];
            allCells.forEach(cell => cell.classList.remove('inactive-cell'));

            // --- Plan and Acc. Plan Rows ---
            if (planRange.first !== -1 && week >= planRange.first && week <= planRange.last) {
                planCell.textContent = planValue > 0 ? planValue.toString() : '0';
                accPlanCell.textContent = accPlan > 0 ? accPlan.toString() : '0';
            } else {
                planCell.textContent = '';
                accPlanCell.textContent = '';
                planCell.classList.add('inactive-cell');
                accPlanCell.classList.add('inactive-cell');
            }

            // --- Actual and Acc. Actual Rows ---
            if (actualRange.first !== -1 && week >= actualRange.first && week <= actualRange.last) {
                actualCell.textContent = actualValue > 0 ? actualValue.toString() : '0';
                accActualCell.textContent = accActual > 0 ? accActual.toString() : '0';
            } else {
                actualCell.textContent = '';
                accActualCell.textContent = '';
                actualCell.classList.add('inactive-cell');
                accActualCell.classList.add('inactive-cell');
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
    
    document.getElementById('total-neua-fa-plan').textContent = totalNeuaFaPlan > 0 ? totalNeuaFaPlan.toString() : '';
    document.getElementById('total-neua-fa-sent').textContent = totalNeuaFaActual > 0 ? totalNeuaFaActual.toString() : '';
    document.getElementById('total-qc-ww-plan').textContent = totalQcWwPlan > 0 ? totalQcWwPlan.toString() : '';
    document.getElementById('total-qc-ww-sent').textContent = totalQcWwActual > 0 ? totalQcWwActual.toString() : '';
    document.getElementById('total-qc-end-plan').textContent = totalQcEndPlan > 0 ? totalQcEndPlan.toString() : '';
    document.getElementById('total-qc-end-sent').textContent = totalQcEndActual > 0 ? totalQcEndActual.toString() : '';
  }

  // --- UTILITY FUNCTIONS ---
  function getTaskWeekRange(task, type) {
    let first = -1;
    let last = -1;
    
    // Find first week with data
    for (let week = 0; week < WEEKS; week++) {
      for (const floor of FLOORS) {
        if (type === 'plan') {
            const cellData = state[floor][week].plan;
            if (cellData.task === task && cellData.value !== null && cellData.value > 0) {
              first = week;
              break; 
            }
        } else { // actual
            const cellData = state[floor][week].actual;
            if (cellData && cellData[task] && cellData[task] > 0) {
                first = week;
                break;
            }
        }
      }
      if (first !== -1) break;
    }

    // Find last week with data
    for (let week = WEEKS - 1; week >= 0; week--) {
      for (const floor of FLOORS) {
        if (type === 'plan') {
            const cellData = state[floor][week].plan;
            if (cellData.task === task && cellData.value !== null && cellData.value > 0) {
              last = week;
              break;
            }
        } else { // actual
            const cellData = state[floor][week].actual;
            if (cellData && cellData[task] && cellData[task] > 0) {
                last = week;
                break;
            }
        }
      }
      if (last !== -1) break;
    }

    return { first, last };
  }

  function updateStatus(message, color, autoClear = true) {
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

  // --- IMAGE CAPTURE ---
  async function captureImage() {
    // The element we want to capture is the whole container.
    const captureElement = document.querySelector('.container');
    // The elements inside we want to hide during capture.
    const actionsElement = document.querySelector('.header-actions');

    if (!captureElement || !html2canvas) {
        console.error('Capture element or html2canvas not found');
        return;
    }

    updateStatus('Capturing...', 'orange', false);

    // Store original style to restore it later
    const originalActionsDisplay = actionsElement.style.display;
    
    try {
        // Hide the action buttons so they don't appear in the screenshot
        actionsElement.style.display = 'none';

        const canvas = await html2canvas(captureElement, {
            scale: 2, // Higher scale for better quality
            backgroundColor: '#ffffff', // Set a white background
            useCORS: true, 
            // These options help capture the entire element, not just the visible part
            width: captureElement.scrollWidth,
            height: captureElement.scrollHeight,
            windowWidth: document.documentElement.offsetWidth,
            windowHeight: document.documentElement.offsetHeight,
        });

        // Create a link to download the image
        const link = document.createElement('a');
        link.download = 'schedule-capture.png';
        link.href = canvas.toDataURL('image/png');
        link.click();

        updateStatus('Capture successful!', 'green');

    } catch (error) {
        console.error('Failed to capture image:', error);
        updateStatus('Capture failed!', 'red');
    } finally {
        // IMPORTANT: Always restore the original style, even if capture fails
        actionsElement.style.display = originalActionsDisplay;
    }
  }

  // --- DATA IMPORT/EXPORT ---
  function exportStateToJson() {
    try {
      updateStatus('Exporting JSON...', 'orange', false);
      const jsonString = JSON.stringify(state, null, 2);
      const blob = new Blob([jsonString], { type: 'application/json' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'schedule-data.json';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      updateStatus('JSON Export Successful!', 'green');
    } catch (error) {
      console.error('Failed to export state to JSON:', error);
      updateStatus('JSON Export Failed!', 'red');
    }
  }

  function importStateFromJson(event) {
    const input = event.target;
    if (!input.files || input.files.length === 0) {
      return;
    }
    const file = input.files[0];
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const text = e.target?.result;
        if (typeof text !== 'string') {
          throw new Error("File could not be read.");
        }
        const importedState = JSON.parse(text);

        // Basic validation: Check if the imported object has keys that are in FLOORS
        const firstFloorKey = FLOORS[0].toString();
        if (!importedState[firstFloorKey] || !importedState[firstFloorKey]['0']) {
           throw new Error('Invalid or corrupted data file.');
        }

        state = importedState;
        saveState();
        renderAll();
        updateStatus('Import Successful!', 'green');
      } catch (error) {
        console.error('Failed to import state from JSON:', error);
        const message = error instanceof Error ? error.message : 'An unknown error occurred.';
        updateStatus(`Import Failed: ${message}`, 'red', false);
      } finally {
          // Reset the input value to allow importing the same file again
          input.value = '';
      }
    };

    reader.onerror = () => {
       updateStatus('Failed to read file!', 'red');
       input.value = '';
    };
    
    reader.readAsText(file);
  }

  // --- EXCEL EXPORT ---
  async function exportToExcel() {
      updateStatus('Exporting...', 'orange', false);
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'VAY CHINNAKHET';
      workbook.created = new Date();
  
      // --- Define Styles ---
      const getFill = (color) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: color } });
      const getFont = (color, bold = false) => ({ color: { argb: color }, bold });
      const border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      const centerAlign = { vertical: 'middle', horizontal: 'center' };

      const styles = {
          'neua-fa-plan': { fill: getFill('fddca9'), font: getFont('FF212121', true), border, alignment: centerAlign },
          'neua-fa-actual': { fill: getFill('feefd8'), font: getFont('FF212121', true), border, alignment: centerAlign },
          'neua-fa-header': { fill: getFill('f79646'), font: getFont('FF212121', true), border, alignment: centerAlign },
          'qc-ww-plan': { fill: getFill('c9d8e8'), font: getFont('FF212121', true), border, alignment: centerAlign },
          'qc-ww-actual': { fill: getFill('dce6f1'), font: getFont('FF212121', true), border, alignment: centerAlign },
          'qc-ww-header': { fill: getFill('4f81bd'), font: getFont('FFFFFFFF', true), border, alignment: centerAlign },
          'qc-end-plan': { fill: getFill('e2eacb'), font: getFont('FF212121', true), border, alignment: centerAlign },
          'qc-end-actual': { fill: getFill('ebf1de'), font: getFont('FF212121', true), border, alignment: centerAlign },
          'qc-end-header': { fill: getFill('9bbb59'), font: getFont('FF212121', true), border, alignment: centerAlign },
          'inactive-cell': { fill: getFill('FFF5F5F5'), font: getFont('FFBDBDBD'), border, alignment: centerAlign },
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
      headerRow1.eachCell(c => c.style = styles.header);
    
      const headerRow2 = mainSheet.addRow([
          '', 'ทั้งหมด', 'ส่ง', 'ทั้งหมด', 'ส่ง', 'ทั้งหมด', 'ส่ง', '',
          ...WEEK_HEADERS.map(h => h.split('/')[0]), ''
      ]);
      headerRow2.eachCell(c => c.style = styles.header);
    
      mainSheet.mergeCells('A1:A2'); mainSheet.mergeCells('H1:H2'); mainSheet.mergeCells('Y1:Y2');
      mainSheet.mergeCells('B1:C1'); mainSheet.mergeCells('D1:E1'); mainSheet.mergeCells('F1:G1');
      mainSheet.mergeCells('I1:J1'); mainSheet.mergeCells('K1:N1'); mainSheet.mergeCells('O1:S1');
      mainSheet.mergeCells('T1:X1');

      ['B1', 'D1', 'F1'].forEach((cell, index) => {
        mainSheet.getCell(cell).style = styles[['neua-fa-header', 'qc-ww-header', 'qc-end-header'][index]];
      });
      ['B2', 'D2', 'F2', 'C2', 'E2', 'G2'].forEach((cell, index) => {
          mainSheet.getCell(cell).style = styles[['neua-fa-header', 'qc-ww-header', 'qc-end-header', 'neua-fa-header', 'qc-ww-header', 'qc-end-header'][index]];
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
              if (actualData) {
                  for (const task of Object.keys(actualData)) {
                      if(actualData[task]) floorTotals[task].actual += actualData[task];
                  }
              }
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
          
          const actualRowValues = Array.from({length: WEEKS}, (_, i) => {
            const actualCellData = state[floor][i].actual;
            if (!actualCellData) return '';
            const totalValue = Object.values(actualCellData).reduce((sum, val) => sum + (val || 0), 0);
            return totalValue > 0 ? totalValue : '';
          });

          const actualRow = mainSheet.addRow([
              '', '', '', '', '', '', '', // for merge
              'Actual',
              ...actualRowValues,
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
          mainSheet.getCell(`A${startRowNum}`).style = styles.default;
          planRow.getCell(8).style = styles['row-label'];
          actualRow.getCell(8).style = styles['row-label'];
        
          mainSheet.getCell(`B${startRowNum}`).style = styles['neua-fa-plan'];
          mainSheet.getCell(`C${startRowNum}`).style = styles['neua-fa-plan'];
          mainSheet.getCell(`D${startRowNum}`).style = styles['qc-ww-plan'];
          mainSheet.getCell(`E${startRowNum}`).style = styles['qc-ww-plan'];
          mainSheet.getCell(`F${startRowNum}`).style = styles['qc-end-plan'];
          mainSheet.getCell(`G${startRowNum}`).style = styles['qc-end-plan'];

          // Style Weekly Data Cells
          for(let week = 0; week < WEEKS; week++) {
              const col = 9 + week;
              const planTask = state[floor][week].plan.task;
              if(planTask) planRow.getCell(col).style = styles[`${planTask}-plan`];
              else planRow.getCell(col).style = styles.default;

              const actualData = state[floor][week].actual;
              const actualTasks = actualData ? Object.keys(actualData).filter(t => actualData[t]) : [];

              if(actualTasks.length > 0) {
                  // Style with the first task's color for simplicity in Excel
                  actualRow.getCell(col).style = styles[`${actualTasks[0]}-actual`];
              }
              else actualRow.getCell(col).style = styles.default;
          }
      });

      // Footer Row
      const summaryData = getSummaryData();
      const footerValues = ['รวม'];
      ['neua-fa', 'qc-ww', 'qc-end'].forEach(task => {
          const totalPlan = summaryData[task].plan.reduce((a, b) => a + b, 0);
          const totalActual = summaryData[task].actual.reduce((a, b) => a + b, 0);
          footerValues.push(totalPlan > 0 ? totalPlan.toString() : '');
          footerValues.push(totalActual > 0 ? totalActual.toString() : '');
      });
      footerValues.push(''); // for Plan/Actual label column
      const footerRow = mainSheet.addRow(footerValues);
      
      // Style footer cells individually
      footerRow.getCell(1).style = styles.footer; // 'รวม'
      footerRow.getCell(2).style = styles['neua-fa-plan'];
      footerRow.getCell(3).style = styles['neua-fa-plan'];
      footerRow.getCell(4).style = styles['qc-ww-plan'];
      footerRow.getCell(5).style = styles['qc-ww-plan'];
      footerRow.getCell(6).style = styles['qc-end-plan'];
      footerRow.getCell(7).style = styles['qc-end-plan'];
      footerRow.getCell(8).style = styles.footer; // Empty cell for Plan/Actual label

      // --- SUMMARY SHEETS ---
      const summaryTasks = [
          {id: 'neua-fa', name: 'สรุป QC เหนือฝ้า'},
          {id: 'qc-ww', name: 'สรุป QC WW'},
          {id: 'qc-end', name: 'สรุป QC End'},
      ];
    
      summaryTasks.forEach(taskInfo => {
          const sheet = workbook.addWorksheet(taskInfo.name);
          sheet.columns = [{ width: 12 }, ...Array(WEEKS).fill({ width: 10 })];
        
          const headerRow = sheet.addRow([taskInfo.name]);
          headerRow.getCell(1).style = styles[`${taskInfo.id}-header`];
          headerRow.getCell(1).font = { ...styles[`${taskInfo.id}-header`].font, size: 14 };
          sheet.mergeCells(1, 1, 1, WEEKS + 1);
        
          const weekHeaderRow = sheet.addRow(['', ...WEEK_HEADERS]);
          weekHeaderRow.getCell(1).style = styles['row-label'];
          weekHeaderRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
              if (colNumber > 1) cell.style = styles.header;
          });
          
          const planRange = getTaskWeekRange(taskInfo.id, 'plan');
          const actualRange = getTaskWeekRange(taskInfo.id, 'actual');

          // Row 1: Plan
          const planRow = sheet.addRow(['Plan']);
          planRow.getCell(1).style = styles['row-label'];
          for (let week = 0; week < WEEKS; week++) {
              const cell = planRow.getCell(week + 2);
              if (planRange.first !== -1 && week >= planRange.first && week <= planRange.last) {
                  const val = summaryData[taskInfo.id].plan[week];
                  cell.value = val > 0 ? val : null;
                  cell.style = styles.default;
              } else {
                  cell.value = null;
                  cell.style = styles['inactive-cell'];
              }
          }
          
          // Row 2: Acc. Plan
          const accPlanRow = sheet.addRow(['Acc. Plan']);
          accPlanRow.getCell(1).style = styles['row-label'];
          let accPlan = 0;
          for (let week = 0; week < WEEKS; week++) {
              accPlan += summaryData[taskInfo.id].plan[week];
              const cell = accPlanRow.getCell(week + 2);
              if (planRange.first !== -1 && week >= planRange.first && week <= planRange.last) {
                  cell.value = accPlan > 0 ? accPlan : null;
                  cell.style = styles.default;
              } else {
                  cell.value = null;
                  cell.style = styles['inactive-cell'];
              }
          }

          // Row 3: Actual
          const actualRowSheet = sheet.addRow(['Actual']);
          actualRowSheet.getCell(1).style = styles['row-label'];
          for (let week = 0; week < WEEKS; week++) {
              const cell = actualRowSheet.getCell(week + 2);
              if (actualRange.first !== -1 && week >= actualRange.first && week <= actualRange.last) {
                  const val = summaryData[taskInfo.id].actual[week];
                  cell.value = val > 0 ? val : null;
                  cell.style = styles.default;
              } else {
                  cell.value = null;
                  cell.style = styles['inactive-cell'];
              }
          }

          // Row 4: Acc. Actual
          const accActualRow = sheet.addRow(['Acc. Actual']);
          accActualRow.getCell(1).style = styles['row-label'];
          let accActual = 0;
          for (let week = 0; week < WEEKS; week++) {
              accActual += summaryData[taskInfo.id].actual[week];
              const cell = accActualRow.getCell(week + 2);
              if (actualRange.first !== -1 && week >= actualRange.first && week <= actualRange.last) {
                  cell.value = accActual > 0 ? accActual : null;
                  cell.style = styles.default;
              } else {
                  cell.value = null;
                  cell.style = styles['inactive-cell'];
              }
          }
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
})();