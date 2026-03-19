/**
 * Lift Dashboard — Final edits
 * - Removed dark mode
 * - Back-to-top button and header navigation restored
 * - Sortable tables
 * - Collapsible Current Lifts
 * - Scrollable Repair Logs
 * - Weekday checkboxes for recurrence
 * - Timetable supports 5-min slots and Week view (Mon-Sun)
 * - Orange theme retained
 *
 * Data persisted in localStorage.
 */

/* Storage keys */
const STORAGE = {
  LIFTS: 'ld_lifts',
  REPAIRS: 'ld_repairs',
  SCHEDULES: 'ld_schedules'
};

/* App state */
let lifts = [];
let repairs = [];
let schedules = [];

/* UI cache */
const UI = {};

document.addEventListener('DOMContentLoaded', init);

function init(){
  cacheUI();
  bindTopNav();
  bindBackToTop();
  bindInventoryForm();
  bindCSVButtons();
  bindRepairForm();
  bindMaintenanceForm();
  bindTimetableControls();
  bindCollapsibleLifts();
  bindSortableTables();
  ensureStorageDefaults();
  loadAll();
  renderAll();
}

/* ---------------- Cache UI ---------------- */
function cacheUI(){
  const $ = id => document.getElementById(id);
  UI.tabButtons = document.querySelectorAll('.tabbtn');

  UI.inventoryForm = $('inventoryForm');
  UI.blockInput = $('blockInput');
  UI.liftNoInput = $('liftNoInput');
  UI.inventorySaveBtn = $('inventorySaveBtn');
  UI.inventoryCancelEdit = $('inventoryCancelEdit');
  UI.editingLiftId = $('editingLiftId');
  UI.filterBlock = $('filterBlock');
  UI.importLiftsFile = $('importLiftsFile');
  UI.importLiftsBtn = $('importLiftsBtn');
  UI.exportLiftsBtn = $('exportLiftsBtn');
  UI.liftTableBody = document.querySelector('#liftTable tbody');
  UI.currentLiftsCard = $('currentLiftsCard');
  UI.currentLiftsBody = $('currentLiftsBody');
  UI.toggleLiftsBtn = $('toggleLiftsBtn');
  UI.toggleLiftsIcon = $('toggleLiftsIcon');
  UI.refreshLiftsBtn = $('refreshLiftsBtn');

  UI.repairForm = $('repairForm');
  UI.rBlock = $('rBlock');
  UI.rLift = $('rLift');
  UI.rStart = $('rStart');
  UI.rEnd = $('rEnd');
  UI.rReason = $('rReason');
  UI.rFollow = $('rFollow');
  UI.rRemarks = $('rRemarks');
  UI.repairTableBody = document.querySelector('#repairTable tbody');
  UI.repairTableWrap = $('repairTableWrap');
  UI.exportRepairs = $('exportRepairs');

  UI.maintenanceForm = $('maintenanceForm');
  UI.mBlock = $('mBlock');
  UI.mLift = $('mLift');
  UI.mType = $('mType');
  UI.mStart = $('mStart');
  UI.mEnd = $('mEnd');
  UI.mRecurrence = $('mRecurrence');
  UI.mInterval = $('mInterval');
  UI.mWeekdays = document.querySelectorAll('.weekday');
  UI.mNote = $('mNote');
  UI.scheduleTableBody = document.querySelector('#scheduleTable tbody');
  UI.clearSchedules = $('clearSchedules');
  UI.exportSchedulesBtn = $('exportSchedulesBtn');
  UI.importSchedulesFile = $('importSchedulesFile');
  UI.importSchedulesBtn = $('importSchedulesBtn');
  UI.scheduleFilterBlock = $('scheduleFilterBlock');

  UI.ttDate = $('ttDate');
  UI.viewMode = $('viewMode');
  UI.ttStartTime = $('ttStartTime');
  UI.ttEndTime = $('ttEndTime');
  UI.ttSlotMins = $('ttSlotMins');
  UI.renderTimetable = $('renderTimetable');
  UI.timetableWrap = $('timetableWrap');
  UI.ttFilterBlock = $('ttFilterBlock');

  UI.statTotalLifts = $('statTotalLifts');
  UI.statSuspended = $('statSuspended');
  UI.statMaintenance = $('statMaintenance');
  UI.statOutstandingRepairs = $('statOutstandingRepairs');
  UI.outstandingTable = document.querySelector('#outstandingTable tbody');
  UI.repeatedTable = document.querySelector('#repeatedTable tbody');
  UI.suspensionSummary = $('suspensionSummary');

  UI.backToTopBtn = $('backToTopBtn');

  UI.toast = $('toast');
  UI.toastMsg = $('toastMsg');
  UI.toastClose = $('toastClose');
  UI.toastClose.addEventListener('click', hideToast);
}

/* ---------------- Top navigation & back-to-top ---------------- */
function bindTopNav(){
  document.querySelectorAll('.tabbtn').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      const panelId = btn.dataset.panel;
      showPanel(panelId);
      // show back-to-top button after jump
      UI.backToTopBtn.style.display = 'inline-flex';
    });
  });
  // default
  showPanel('dashboardPanel');
}

function showPanel(panelId){
  document.querySelectorAll('.panel').forEach(p=>p.hidden = p.id !== panelId);
  document.querySelectorAll('.tabbtn').forEach(b=>b.classList.toggle('active', b.dataset.panel === panelId));
  window.scrollTo({top:0, behavior:'smooth'});
  const panel = document.getElementById(panelId);
  if(panel){ panel.setAttribute('tabindex','-1'); panel.focus(); }
}

function bindBackToTop(){
  UI.backToTopBtn.addEventListener('click', ()=>{ window.scrollTo({top:0, behavior:'smooth'}); UI.backToTopBtn.style.display='none'; });
  // hide initially
  UI.backToTopBtn.style.display = 'none';
}

/* ---------------- Storage ---------------- */
function ensureStorageDefaults(){
  if(!localStorage.getItem(STORAGE.LIFTS)) save(STORAGE.LIFTS, []);
  if(!localStorage.getItem(STORAGE.REPAIRS)) save(STORAGE.REPAIRS, []);
  if(!localStorage.getItem(STORAGE.SCHEDULES)) save(STORAGE.SCHEDULES, []);
}
function save(key, value){ localStorage.setItem(key, JSON.stringify(value)); }
function load(key){ try{ return JSON.parse(localStorage.getItem(key)); }catch(e){ return null; } }
function loadAll(){ lifts = load(STORAGE.LIFTS) || []; repairs = load(STORAGE.REPAIRS) || []; schedules = load(STORAGE.SCHEDULES) || []; }

/* ---------------- Inventory add/edit/delete ---------------- */
function bindInventoryForm(){
  UI.inventoryForm.addEventListener('submit', e=>{
    e.preventDefault();
    const block = UI.blockInput.value.trim();
    const liftNo = UI.liftNoInput.value.trim();
    if(!block || !liftNo) return;
    const editingId = UI.editingLiftId.value;
    if(editingId){
      const item = lifts.find(l=>l.id===editingId);
      if(item){ item.block = block; item.liftNo = liftNo; save(STORAGE.LIFTS, lifts); UI.editingLiftId.value=''; UI.inventorySaveBtn.innerHTML = `<svg class="icon"><use href="#icon-add"></use></svg> Add Lift`; showToast('Lift updated'); }
    } else {
      if(lifts.some(l=>l.block===block && l.liftNo===liftNo)){ showToast('Lift already exists'); return; }
      lifts.push({id:genId(), block, liftNo});
      save(STORAGE.LIFTS, lifts);
      showToast('Lift added');
    }
    renderAll();
    UI.blockInput.value=''; UI.liftNoInput.value='';
  });

  UI.inventoryCancelEdit.addEventListener('click', ()=>{
    UI.editingLiftId.value=''; UI.inventorySaveBtn.innerHTML = `<svg class="icon"><use href="#icon-add"></use></svg> Add Lift`; UI.inventoryCancelEdit.hidden=true; UI.blockInput.value=''; UI.liftNoInput.value='';
  });

  UI.filterBlock.addEventListener('change', renderLifts);
  UI.refreshLiftsBtn.addEventListener('click', ()=>{ renderLifts(); showToast('Lifts refreshed'); });
}

/* ---------------- Repair form (show immediately) ---------------- */
function bindRepairForm(){
  UI.repairForm.addEventListener('submit', e=>{
    e.preventDefault();
    const entry = {
      id: genId(),
      block: UI.rBlock.value,
      liftNo: UI.rLift.value,
      start: UI.rStart.value,
      end: UI.rEnd.value || '',
      reason: UI.rReason.value,
      followUp: UI.rFollow.value,
      remarks: UI.rRemarks.value || ''
    };
    repairs.push(entry);
    save(STORAGE.REPAIRS, repairs);
    renderAll();
    UI.repairForm.reset();
    showToast('Repair log added');
    showPanel('repairPanel');
    highlightLatestRepair();
  });

  UI.exportRepairs.addEventListener('click', ()=>{
    const blob = new Blob([JSON.stringify(repairs, null, 2)], {type:'application/json'});
    downloadBlob(blob, 'repair-logs.json', 'application/json');
  });
}

/* ---------------- Maintenance form with weekday checkboxes ---------------- */
function bindMaintenanceForm(){
  UI.maintenanceForm.addEventListener('submit', e=>{
    e.preventDefault();
    const start = UI.mStart.value;
    const end = UI.mEnd.value;
    if(!start || !end || new Date(start) >= new Date(end)){ showToast('Invalid start/end'); return; }
    const recType = UI.mRecurrence.value;
    const interval = Math.max(1, parseInt(UI.mInterval.value||'1',10));
    const weekdays = Array.from(UI.mWeekdays).filter(cb=>cb.checked).map(cb=>cb.value);
    const entry = {
      id: genId(),
      block: UI.mBlock.value,
      liftNo: UI.mLift.value,
      type: UI.mType.value,
      start,
      end,
      note: UI.mNote.value || '',
      recurrence: recType === 'none' ? null : { type: recType, interval, weekdays }
    };
    schedules.push(entry);
    save(STORAGE.SCHEDULES, schedules);
    renderAll();
    UI.maintenanceForm.reset();
    UI.mInterval.value='1';
    showToast('Schedule added');
  });

  UI.clearSchedules.addEventListener('click', ()=>{
    if(confirm('Clear all schedules?')){ schedules=[]; save(STORAGE.SCHEDULES, schedules); renderAll(); showToast('Schedules cleared'); }
  });

  UI.exportSchedulesBtn.addEventListener('click', ()=>{
    const rows = schedules.map(s=>({
      block:s.block, liftNo:s.liftNo, type:s.type, start:s.start, end:s.end,
      recurrenceType: s.recurrence ? s.recurrence.type : '',
      recurrenceInterval: s.recurrence ? s.recurrence.interval : '',
      recurrenceWeekdays: s.recurrence ? (s.recurrence.weekdays||[]).join(',') : '',
      note: s.note || ''
    }));
    const csv = toCSV(rows, ['block','liftNo','type','start','end','recurrenceType','recurrenceInterval','recurrenceWeekdays','note']);
    downloadBlob(new Blob([csv], {type:'text/csv'}), 'schedules.csv', 'text/csv');
  });

  UI.importSchedulesBtn.addEventListener('click', ()=>{
    const file = UI.importSchedulesFile.files[0];
    if(!file){ showToast('Choose CSV first'); return; }
    readCSVFile(file, rows=>{
      rows.forEach(r=>{
        if(r.block && r.liftNo && r.start && r.end){
          const rec = r.recurrenceType ? { type:r.recurrenceType, interval: parseInt(r.recurrenceInterval||'1',10)||1, weekdays: r.recurrenceWeekdays ? r.recurrenceWeekdays.split(',').map(x=>x.trim().toUpperCase()) : [] } : null;
          schedules.push({ id:genId(), block:r.block, liftNo:r.liftNo, type:r.type||'maintenance', start:r.start, end:r.end, note:r.note||'', recurrence:rec });
        }
      });
      save(STORAGE.SCHEDULES, schedules);
      renderAll();
      showToast('Schedules imported');
    });
  });

  UI.scheduleFilterBlock.addEventListener('change', renderSchedules);
}

/* ---------------- Timetable controls (day/week, 5-min slots) ---------------- */
function bindTimetableControls(){
  UI.renderTimetable.addEventListener('click', renderTimetable);
  const today = new Date().toISOString().slice(0,10);
  UI.ttDate.value = today;
  UI.viewMode.addEventListener('change', ()=>{ /* no-op; render when user clicks Render */ });
  UI.ttFilterBlock.addEventListener('change', renderTimetable);
}

/* ---------------- CSV import/export for lifts ---------------- */
function bindCSVButtons(){
  UI.importLiftsBtn.addEventListener('click', ()=>{
    const file = UI.importLiftsFile.files[0];
    if(!file){ showToast('Choose CSV first'); return; }
    readCSVFile(file, rows=>{
      rows.forEach(r=>{ if(r.block && r.liftNo && !lifts.some(l=>l.block===r.block && l.liftNo===r.liftNo)) lifts.push({id:genId(), block:r.block, liftNo:r.liftNo}); });
      save(STORAGE.LIFTS, lifts); renderAll(); showToast('Lifts imported');
    });
  });
  UI.exportLiftsBtn.addEventListener('click', ()=>{ const csv = toCSV(lifts, ['block','liftNo']); downloadBlob(new Blob([csv], {type:'text/csv'}), 'lifts.csv', 'text/csv'); });
}

/* ---------------- Collapsible Current Lifts ---------------- */
function bindCollapsibleLifts(){
  UI.toggleLiftsBtn.addEventListener('click', ()=>{
    const body = UI.currentLiftsBody;
    const icon = UI.toggleLiftsIcon;
    if(body.style.display === 'none' || getComputedStyle(body).display === 'none'){
      body.style.display = ''; icon.innerHTML = '<use href="#icon-collapse"></use>';
    } else {
      body.style.display = 'none'; icon.innerHTML = '<use href="#icon-expand"></use>';
    }
  });
}

/* ---------------- Sortable tables ---------------- */
function bindSortableTables(){
  document.querySelectorAll('table.sortable thead th[data-key]').forEach(th=>{
    th.addEventListener('click', ()=>{
      const table = th.closest('table');
      const key = th.dataset.key;
      const current = th.classList.contains('sort-asc') ? 'asc' : th.classList.contains('sort-desc') ? 'desc' : null;
      // clear other headers
      table.querySelectorAll('thead th').forEach(h=>h.classList.remove('sort-asc','sort-desc'));
      const dir = current === 'asc' ? 'desc' : 'asc';
      th.classList.add(dir === 'asc' ? 'sort-asc' : 'sort-desc');
      sortTable(table.id, key, dir);
    });
  });
}

/* Generic table sorter (works with our data arrays) */
function sortTable(tableId, key, dir){
  // map table id to data source and re-render
  const dirMul = dir === 'asc' ? 1 : -1;
  if(tableId === 'liftTable'){
    lifts.sort((a,b)=> compare(a[key], b[key]) * dirMul);
    save(STORAGE.LIFTS, lifts); renderLifts();
  } else if(tableId === 'repairTable'){
    repairs.sort((a,b)=> compare(a[key], b[key]) * dirMul);
    save(STORAGE.REPAIRS, repairs); renderRepairs();
  } else if(tableId === 'scheduleTable'){
    schedules.sort((a,b)=> compare(a[key], b[key]) * dirMul);
    save(STORAGE.SCHEDULES, schedules); renderSchedules();
  } else if(tableId === 'outstandingTable'){
    // outstanding is derived; sort repairs then re-render outstanding
    repairs.sort((a,b)=> compare(a[key], b[key]) * dirMul);
    save(STORAGE.REPAIRS, repairs); renderDashboard();
  } else if(tableId === 'repeatedTable'){
    // repeated is derived; recompute and sort in renderRepeatedTable
    renderDashboard();
  }
}
function compare(a,b){
  if(a==null) return -1;
  if(b==null) return 1;
  if(!isNaN(Date.parse(a)) && !isNaN(Date.parse(b))) return new Date(a) - new Date(b);
  if(typeof a === 'number' || typeof b === 'number') return (a||0) - (b||0);
  return String(a).localeCompare(String(b));
}

/* ---------------- Render all ---------------- */
function renderAll(){
  loadAll();
  renderDashboard();
  renderLifts();
  populateSelects();
  renderRepairs();
  renderSchedules();
}

/* ---------------- Dashboard ---------------- */
function renderDashboard(){
  UI.statTotalLifts.querySelector('.stat-value').textContent = lifts.length;
  const now = new Date();
  const currentSuspended = countActiveOccurrences(now);
  const currentMaintenance = countActiveOccurrences(now, 'maintenance');
  UI.statSuspended.querySelector('.stat-value').textContent = currentSuspended;
  UI.statMaintenance.querySelector('.stat-value').textContent = currentMaintenance;

  const outstanding = repairs.filter(r => !r.end || new Date(r.end) > now);
  UI.statOutstandingRepairs.querySelector('.stat-value').textContent = outstanding.length;
  renderOutstandingTable(outstanding);

  const repeated = computeRepeatedRepairs(90);
  renderRepeatedTable(repeated);

  renderSuspensionSummary();
}

/* Outstanding table */
function renderOutstandingTable(list){
  UI.outstandingTable.innerHTML = '';
  list.slice().sort((a,b)=> new Date(a.start)-new Date(b.start)).forEach(r=>{
    const tr = document.createElement('tr');
    const age = humanDuration(new Date(r.start), new Date());
    tr.innerHTML = `<td>${escapeHtml(r.block)}</td><td>${escapeHtml(r.liftNo)}</td><td>${formatDateTime(r.start)}</td><td>${escapeHtml(r.reason)}</td><td>${escapeHtml(r.followUp)}</td><td>${age}</td>`;
    UI.outstandingTable.appendChild(tr);
  });
}

/* Repeated repairs */
function computeRepeatedRepairs(daysWindow=90){
  const cutoff = new Date(Date.now() - daysWindow*24*60*60*1000);
  const map = {};
  repairs.forEach(r=>{
    const key = `${r.block}||${r.liftNo}`;
    const d = new Date(r.start);
    if(d >= cutoff){
      map[key] = map[key] || {block:r.block, liftNo:r.liftNo, count:0, last:null};
      map[key].count += 1;
      if(!map[key].last || new Date(r.start) > new Date(map[key].last)) map[key].last = r.start;
    }
  });
  return Object.values(map).filter(x=>x.count>1).sort((a,b)=> b.count - a.count);
}
function renderRepeatedTable(list){
  UI.repeatedTable.innerHTML = '';
  list.forEach(r=>{
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${escapeHtml(r.block)}</td><td>${escapeHtml(r.liftNo)}</td><td>${r.count}</td><td>${formatDateTime(r.last)}</td>`;
    UI.repeatedTable.appendChild(tr);
  });
}

/* Suspension summary */
function renderSuspensionSummary(){
  UI.suspensionSummary.innerHTML = '';
  const now = new Date();
  const rows = lifts.slice().sort((a,b)=> a.block.localeCompare(b.block) || a.liftNo.localeCompare(b.liftNo));
  rows.forEach(l=>{
    const card = document.createElement('div'); card.className='suspension-card';
    const title = document.createElement('div'); title.style.fontWeight='700'; title.textContent = `${l.block} • ${l.liftNo}`;
    const activeOcc = expandOccurrencesForLift(l.block, l.liftNo, now, now).filter(o=> new Date(o.start) <= now && new Date(o.end) > now);
    let status='Active', statusClass='', nextResume=null;
    if(activeOcc.length){
      if(activeOcc.some(a=>a.type==='lock')){ status='Locked'; statusClass='lock'; }
      else if(activeOcc.some(a=>a.type==='maintenance')){ status='Maintenance'; statusClass='maintenance'; }
      else { status='Suspended'; statusClass='suspended'; }
      nextResume = activeOcc.reduce((acc,s)=>{ const e=new Date(s.end); return !acc||e<acc?e:acc; }, null);
    } else {
      const futureOcc = expandOccurrencesForLift(l.block, l.liftNo, now, addDays(now,365));
      if(futureOcc.length){ nextResume = futureOcc.reduce((acc,s)=>{ const st=new Date(s.start); return !acc||st<acc?st:acc; }, null); }
    }
    const statusEl = document.createElement('div'); statusEl.textContent = `Status: ${status}`; statusEl.style.marginTop='8px'; statusEl.style.fontWeight='600';
    if(statusClass) statusEl.className = statusClass;
    const resumeEl = document.createElement('div'); resumeEl.className='muted'; resumeEl.style.marginTop='6px';
    resumeEl.textContent = nextResume ? `Next resume: ${formatDateTime(nextResume)}` : 'No upcoming schedule';
    card.appendChild(title); card.appendChild(statusEl); card.appendChild(resumeEl);
    UI.suspensionSummary.appendChild(card);
  });

  const suspendedCount = countActiveOccurrences(new Date());
  const maintenanceCount = countActiveOccurrences(new Date(), 'maintenance');
  UI.statSuspended.querySelector('.stat-value').textContent = suspendedCount;
  UI.statMaintenance.querySelector('.stat-value').textContent = maintenanceCount;
}

/* ---------------- Lifts list ---------------- */
function renderLifts(){
  UI.liftTableBody.innerHTML = '';
  const blockFilter = UI.filterBlock.value || '';
  lifts.forEach(l=>{
    if(blockFilter && l.block !== blockFilter) return;
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${escapeHtml(l.block)}</td><td>${escapeHtml(l.liftNo)}</td><td>
      <button class="btn" data-id="${l.id}" data-action="edit"><svg class="icon"><use href="#icon-edit"></use></svg> Edit</button>
      <button class="btn" data-id="${l.id}" data-action="delete"><svg class="icon"><use href="#icon-delete"></use></svg> Delete</button>
    </td>`;
    UI.liftTableBody.appendChild(tr);
  });

  UI.liftTableBody.querySelectorAll('button[data-action="delete"]').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      const id = btn.dataset.id;
      if(!confirm('Delete lift?')) return;
      const removed = lifts.find(x=>x.id===id);
      lifts = lifts.filter(x=>x.id!==id);
      save(STORAGE.LIFTS, lifts);
      schedules = schedules.filter(s=>!(s.block===removed.block && s.liftNo===removed.liftNo));
      repairs = repairs.filter(r=>!(r.block===removed.block && r.liftNo===removed.liftNo));
      save(STORAGE.SCHEDULES, schedules); save(STORAGE.REPAIRS, repairs);
      renderAll(); showToast('Lift deleted');
    });
  });

  UI.liftTableBody.querySelectorAll('button[data-action="edit"]').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      const id = btn.dataset.id;
      const item = lifts.find(x=>x.id===id);
      if(!item) return;
      UI.blockInput.value = item.block; UI.liftNoInput.value = item.liftNo; UI.editingLiftId.value = item.id;
      UI.inventorySaveBtn.innerHTML = `<svg class="icon"><use href="#icon-edit"></use></svg> Save Changes`; UI.inventoryCancelEdit.hidden=false;
      showPanel('inventoryPanel'); UI.blockInput.focus();
    });
  });

  populateFilterOptions();
}

/* ---------------- Repairs ---------------- */
function renderRepairs(){
  UI.repairTableBody.innerHTML = '';
  repairs.slice().reverse().forEach(r=>{
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${escapeHtml(r.block)}</td><td>${escapeHtml(r.liftNo)}</td><td>${formatDateTime(r.start)}</td><td>${r.end?formatDateTime(r.end):'<em>Open</em>'}</td><td>${escapeHtml(r.reason)}</td><td>${escapeHtml(r.followUp)}</td><td>${escapeHtml(r.remarks||'')}</td><td><button class="btn" data-id="${r.id}" data-action="delr"><svg class="icon"><use href="#icon-delete"></use></svg> Delete</button></td>`;
    UI.repairTableBody.appendChild(tr);
  });

  UI.repairTableBody.querySelectorAll('button[data-action="delr"]').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      const id = btn.dataset.id;
      if(!confirm('Delete repair log?')) return;
      repairs = repairs.filter(x=>x.id!==id);
      save(STORAGE.REPAIRS, repairs);
      renderAll(); showToast('Repair deleted');
    });
  });
}

/* Highlight latest repair */
function highlightLatestRepair(){
  const tbody = UI.repairTableBody;
  const firstRow = tbody.querySelector('tr');
  if(!firstRow) return;
  firstRow.style.transition = 'background-color 0.4s';
  firstRow.style.backgroundColor = 'rgba(250,240,230,0.9)';
  setTimeout(()=>{ firstRow.style.backgroundColor = ''; }, 1400);
}

/* ---------------- Schedules ---------------- */
function renderSchedules(){
  UI.scheduleTableBody.innerHTML = '';
  const blockFilter = UI.scheduleFilterBlock.value || '';
  schedules.slice().reverse().forEach(s=>{
    if(blockFilter && s.block !== blockFilter) return;
    const recText = s.recurrence ? `${s.recurrence.type} every ${s.recurrence.interval}` + (s.recurrence.weekdays && s.recurrence.weekdays.length ? ` (${s.recurrence.weekdays.join(',')})` : '') : 'None';
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${escapeHtml(s.block)}</td><td>${escapeHtml(s.liftNo)}</td><td>${escapeHtml(s.type)}</td><td>${formatDateTime(s.start)}</td><td>${formatDateTime(s.end)}</td><td>${escapeHtml(recText)}</td><td>${escapeHtml(s.note||'')}</td><td><button class="btn" data-id="${s.id}" data-action="dels"><svg class="icon"><use href="#icon-delete"></use></svg> Delete</button></td>`;
    UI.scheduleTableBody.appendChild(tr);
  });

  UI.scheduleTableBody.querySelectorAll('button[data-action="dels"]').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      const id = btn.dataset.id;
      if(!confirm('Delete schedule entry?')) return;
      schedules = schedules.filter(x=>x.id!==id);
      save(STORAGE.SCHEDULES, schedules);
      renderAll(); showToast('Schedule deleted');
    });
  });
}

/* ---------------- Populate selects (fix selection) ---------------- */
function populateSelects(){
  const blocks = [...new Set(lifts.map(l=>l.block))].sort();

  ['rBlock','mBlock','filterBlock','scheduleFilterBlock','ttFilterBlock'].forEach(id=>{
    const el = document.getElementById(id); if(!el) return;
    const prev = el.value || '';
    el.innerHTML = '<option value="">All</option>';
    blocks.forEach(b=>{ const opt = document.createElement('option'); opt.value=b; opt.textContent=b; el.appendChild(opt); });
    if(prev) el.value = prev;
  });

  ['rLift','mLift'].forEach(id=>{
    const el = document.getElementById(id); if(!el) return;
    const blockSelectId = id === 'rLift' ? 'rBlock' : 'mBlock';
    const blockVal = document.getElementById(blockSelectId) ? document.getElementById(blockSelectId).value : '';
    const prev = el.value || '';
    el.innerHTML = '';
    lifts.filter(l=>!blockVal || l.block===blockVal).forEach(l=>{ const opt = document.createElement('option'); opt.value=l.liftNo; opt.textContent=l.liftNo; el.appendChild(opt); });
    if(prev) el.value = prev;
  });

  const rBlockEl = document.getElementById('rBlock'); const mBlockEl = document.getElementById('mBlock');
  if(rBlockEl) rBlockEl.onchange = ()=> filterLiftsForBlock('rBlock','rLift');
  if(mBlockEl) mBlockEl.onchange = ()=> filterLiftsForBlock('mBlock','mLift');

  if(document.getElementById('rBlock')) filterLiftsForBlock('rBlock','rLift');
  if(document.getElementById('mBlock')) filterLiftsForBlock('mBlock','mLift');
}

/* Filter lifts for block */
function filterLiftsForBlock(blockId, liftId){
  const block = document.getElementById(blockId).value;
  const sel = document.getElementById(liftId);
  if(!sel) return;
  const prev = sel.value || '';
  sel.innerHTML = '';
  lifts.filter(l=>!block || l.block===block).forEach(l=>{ const opt = document.createElement('option'); opt.value=l.liftNo; opt.textContent=l.liftNo; sel.appendChild(opt); });
  if(prev) sel.value = prev;
}

/* ---------------- Timetable rendering (day/week) ---------------- */
function renderTimetable(){
  const view = UI.viewMode.value; // 'day' or 'week'
  const date = UI.ttDate.value;
  if(!date){ showToast('Select a date'); return; }
  const startTime = UI.ttStartTime.value || '00:00';
  const endTime = UI.ttEndTime.value || '23:59';
  const slotMins = parseInt(UI.ttSlotMins.value,10) || 15;
  const startDate = new Date(`${date}T${startTime}`);
  const endDate = new Date(`${date}T${endTime}`);
  if(startDate >= endDate){ showToast('Start must be before end'); return; }

  UI.timetableWrap.innerHTML = '';

  if(view === 'day'){
    renderDayTimetable(startDate, endDate, slotMins);
  } else {
    // week view: compute Monday of the week containing date
    const d = new Date(date);
    const day = d.getDay(); // 0 Sun .. 6 Sat
    const monday = addDays(d, (day === 0 ? -6 : 1 - day)); // shift to Monday
    const weekStart = new Date(monday); weekStart.setHours(0,0,0,0);
    const weekEnd = addDays(weekStart, 7);
    renderWeekTimetable(weekStart, weekEnd, slotMins);
  }
}

/* Day timetable (columns: time slots) */
function renderDayTimetable(startDate, endDate, slotMins){
  const slots = [];
  let cur = new Date(startDate);
  while(cur < endDate){
    const next = new Date(cur.getTime() + slotMins*60000);
    slots.push({start: new Date(cur), end: new Date(Math.min(next, endDate))});
    cur = next;
  }

  const blockFilter = UI.ttFilterBlock.value || '';
  const rows = lifts.slice().sort((a,b)=> a.block.localeCompare(b.block) || a.liftNo.localeCompare(b.liftNo)).filter(r=> !blockFilter || r.block===blockFilter);

  const table = document.createElement('table'); table.className='timetable';
  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  headRow.innerHTML = `<th class="liftCell">Lift</th>` + slots.map(s=>`<th>${formatTime(s.start)}<br><small>${formatTime(s.end)}</small></th>`).join('');
  thead.appendChild(headRow); table.appendChild(thead);

  const tbody = document.createElement('tbody');
  rows.forEach(r=>{
    const tr = document.createElement('tr');
    const liftCell = document.createElement('td'); liftCell.className='liftCell'; liftCell.textContent = `${r.block} • ${r.liftNo}`; tr.appendChild(liftCell);
    const occs = expandOccurrencesForLift(r.block, r.liftNo, startDate, endDate);
    slots.forEach(slot=>{
      const td = document.createElement('td');
      const overlapping = occs.some(o=> new Date(o.start) < slot.end && new Date(o.end) > slot.start);
      if(overlapping){
        const types = occs.filter(o=> new Date(o.start) < slot.end && new Date(o.end) > slot.start).map(x=>x.type);
        if(types.includes('lock')) td.classList.add('lock');
        else if(types.includes('maintenance')) td.classList.add('maintenance');
        else td.classList.add('suspended');
        td.textContent = 'Suspended';
        const notes = occs.filter(o=> new Date(o.start) < slot.end && new Date(o.end) > slot.start).map(x=>`${x.type}${x.note? ' • ' + x.note : ''}`);
        td.title = notes.join('; ');
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  UI.timetableWrap.appendChild(table);
  renderOverlapSummary(slots, rows);
}

/* Week timetable (columns: Mon..Sun, rows: time slots) */
function renderWeekTimetable(weekStart, weekEnd, slotMins){
  // build time slots for a single day (use start/end from controls)
  const dayStart = new Date(`${weekStart.toISOString().slice(0,10)}T${UI.ttStartTime.value || '00:00'}`);
  const dayEnd = new Date(`${weekStart.toISOString().slice(0,10)}T${UI.ttEndTime.value || '23:59'}`);
  if(dayStart >= dayEnd){ showToast('Start must be before end'); return; }

  const slots = [];
  let cur = new Date(dayStart);
  while(cur < dayEnd){
    const next = new Date(cur.getTime() + slotMins*60000);
    slots.push({start: new Date(cur), end: new Date(Math.min(next, dayEnd))});
    cur = next;
  }

  const blockFilter = UI.ttFilterBlock.value || '';
  const rows = lifts.slice().sort((a,b)=> a.block.localeCompare(b.block) || a.liftNo.localeCompare(b.liftNo)).filter(r=> !blockFilter || r.block===blockFilter);

  const table = document.createElement('table'); table.className='timetable';
  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  headRow.innerHTML = `<th class="liftCell">Time</th>` + ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'].map(d=>`<th>${d}</th>`).join('');
  thead.appendChild(headRow); table.appendChild(thead);

  const tbody = document.createElement('tbody');
  slots.forEach(slot=>{
    const tr = document.createElement('tr');
    const timeCell = document.createElement('td'); timeCell.className='liftCell'; timeCell.textContent = `${formatTime(slot.start)} - ${formatTime(slot.end)}`; tr.appendChild(timeCell);
    // for each day, check overlapping occurrences across that day's datetime
    for(let i=0;i<7;i++){
      const dayStartDate = addDays(weekStart, i);
      const slotStart = new Date(dayStartDate); slotStart.setHours(slot.start.getHours(), slot.start.getMinutes(), 0, 0);
      const slotEnd = new Date(dayStartDate); slotEnd.setHours(slot.end.getHours(), slot.end.getMinutes(), 0, 0);
      const td = document.createElement('td');
      // check if any lift has schedule overlapping this slot; we will show per-lift rows below instead of aggregated
      td.textContent = '';
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  });

  // For week view we will render per-lift blocks below the time grid for clarity
  UI.timetableWrap.appendChild(table);

  // Now render per-lift week rows (compact)
  rows.forEach(r=>{
    const card = document.createElement('div'); card.className='card'; card.style.marginTop='8px';
    const title = document.createElement('div'); title.style.fontWeight='700'; title.textContent = `${r.block} • ${r.liftNo}`;
    card.appendChild(title);
    const grid = document.createElement('div'); grid.style.display='grid'; grid.style.gridTemplateColumns='repeat(7,1fr)'; grid.style.gap='6px'; grid.style.marginTop='8px';
    for(let i=0;i<7;i++){
      const dayStartDate = addDays(weekStart, i);
      const dayEndDate = addDays(dayStartDate,1);
      const occs = expandOccurrencesForLift(r.block, r.liftNo, dayStartDate, dayEndDate);
      const cell = document.createElement('div'); cell.style.minHeight='36px'; cell.style.border='1px solid rgba(11,18,32,0.04)'; cell.style.borderRadius='6px'; cell.style.padding='6px'; cell.style.background='#fff';
      if(occs.length){
        occs.forEach(o=>{
          const el = document.createElement('div'); el.style.fontSize='0.85rem'; el.style.padding='4px'; el.style.borderRadius='4px'; el.style.marginBottom='4px';
          if(o.type==='lock'){ el.style.background='linear-gradient(90deg, rgba(234,88,12,0.12), rgba(234,88,12,0.06))'; el.style.color='#4b2b10'; }
          else if(o.type==='maintenance'){ el.style.background='linear-gradient(90deg, rgba(255,193,7,0.12), rgba(255,193,7,0.06))'; el.style.color='#4b2b10'; }
          else { el.style.background='linear-gradient(90deg, rgba(249,115,22,0.12), rgba(249,115,22,0.06))'; el.style.color='#4b2b10'; }
          el.textContent = `${formatTime(new Date(o.start))} - ${formatTime(new Date(o.end))} ${o.note? ' • ' + o.note : ''}`;
          cell.appendChild(el);
        });
      } else {
        cell.textContent = '—';
        cell.style.color = 'var(--muted)';
      }
      grid.appendChild(cell);
    }
    card.appendChild(grid);
    UI.timetableWrap.appendChild(card);
  });

  // Add overlap summary for week
  renderOverlapSummaryWeek(weekStart, slots, rows);
}

/* Overlap summary for day/week */
function renderOverlapSummary(slots, rows){
  const summary = document.createElement('div'); summary.className='card'; const h = document.createElement('h3'); h.textContent='Overlap Summary'; summary.appendChild(h);
  const list = document.createElement('div'); list.style.display='grid'; list.style.gridTemplateColumns='repeat(auto-fit,minmax(220px,1fr))'; list.style.gap='8px';
  slots.forEach(slot=>{
    const sdiv = document.createElement('div'); sdiv.style.background='rgba(255,255,255,0.95)'; sdiv.style.padding='8px'; sdiv.style.borderRadius='8px';
    const title = document.createElement('strong'); title.textContent = `${formatTime(slot.start)} - ${formatTime(slot.end)}`; sdiv.appendChild(title);
    const ul = document.createElement('ul'); ul.style.margin='8px 0 0 0'; ul.style.padding='0 0 0 16px';
    rows.forEach(r=>{
      const occs = expandOccurrencesForLift(r.block, r.liftNo, slot.start, slot.end);
      if(occs.some(o=> new Date(o.start) < slot.end && new Date(o.end) > slot.start)){
        const li = document.createElement('li'); li.textContent = `${r.block} • ${r.liftNo}`; ul.appendChild(li);
      }
    });
    if(!ul.children.length){ const none = document.createElement('div'); none.style.color='var(--muted)'; none.textContent='No suspensions'; sdiv.appendChild(none); } else sdiv.appendChild(ul);
    list.appendChild(sdiv);
  });
  summary.appendChild(list); UI.timetableWrap.appendChild(summary);
}

/* Week overlap summary (compact) */
function renderOverlapSummaryWeek(weekStart, slots, rows){
  const summary = document.createElement('div'); summary.className='card'; const h = document.createElement('h3'); h.textContent='Weekly Overlap Summary'; summary.appendChild(h);
  const grid = document.createElement('div'); grid.style.display='grid'; grid.style.gridTemplateColumns='repeat(7,1fr)'; grid.style.gap='8px';
  for(let i=0;i<7;i++){
    const day = addDays(weekStart, i);
    const sdiv = document.createElement('div'); sdiv.style.background='rgba(255,255,255,0.95)'; sdiv.style.padding='8px'; sdiv.style.borderRadius='8px';
    const title = document.createElement('strong'); title.textContent = day.toLocaleDateString(undefined, {weekday:'short', month:'short', day:'numeric'}); sdiv.appendChild(title);
    const ul = document.createElement('ul'); ul.style.margin='8px 0 0 0'; ul.style.padding='0 0 0 16px';
    rows.forEach(r=>{
      const occs = expandOccurrencesForLift(r.block, r.liftNo, day, addDays(day,1));
      if(occs.length){ const li = document.createElement('li'); li.textContent = `${r.block} • ${r.liftNo}`; ul.appendChild(li); }
    });
    if(!ul.children.length){ const none = document.createElement('div'); none.style.color='var(--muted)'; none.textContent='No suspensions'; sdiv.appendChild(none); } else sdiv.appendChild(ul);
    grid.appendChild(sdiv);
  }
  summary.appendChild(grid); UI.timetableWrap.appendChild(summary);
}

/* ---------------- Expand occurrences (recurrence support) ---------------- */
function expandOccurrencesForLift(block, liftNo, rangeStart, rangeEnd){
  const occs = [];
  schedules.forEach(s=>{
    if(s.block !== block || s.liftNo !== liftNo) return;
    if(!s.recurrence){
      const st = new Date(s.start), en = new Date(s.end);
      if(st < rangeEnd && en > rangeStart) occs.push({start:s.start, end:s.end, type:s.type, note:s.note, source:s.id});
    } else {
      const rec = s.recurrence;
      const capEnd = addDays(rangeEnd, 365);
      const baseStart = new Date(s.start), baseEnd = new Date(s.end);
      let iterStart = new Date(baseStart);
      if(rec.type === 'daily'){
        const daysDiff = Math.floor((rangeStart - baseStart)/(24*60*60*1000));
        const steps = Math.max(0, Math.floor(daysDiff/rec.interval)-1);
        iterStart = addDays(baseStart, steps*rec.interval);
      } else if(rec.type === 'weekly'){
        const weeksDiff = Math.floor((rangeStart - baseStart)/(7*24*60*60*1000));
        const steps = Math.max(0, Math.floor(weeksDiff/rec.interval)-1);
        iterStart = addDays(baseStart, steps*rec.interval*7);
      } else if(rec.type === 'monthly'){
        const monthsDiff = (rangeStart.getFullYear()-baseStart.getFullYear())*12 + (rangeStart.getMonth()-baseStart.getMonth());
        const steps = Math.max(0, Math.floor(monthsDiff/rec.interval)-1);
        iterStart = addMonths(baseStart, steps*rec.interval);
      }
      let safety=0; let curStart = new Date(iterStart);
      while(curStart < capEnd && safety < 5000){
        if(rec.type === 'daily'){
          const occStart = new Date(curStart);
          const occEnd = new Date(occStart.getTime() + (baseEnd - baseStart));
          if(occStart < rangeEnd && occEnd > rangeStart) occs.push({start:occStart.toISOString(), end:occEnd.toISOString(), type:s.type, note:s.note, source:s.id});
          curStart = addDays(curStart, rec.interval);
        } else if(rec.type === 'weekly'){
          const days = (rec.weekdays && rec.weekdays.length) ? rec.weekdays : [weekdayCode(baseStart)];
          const weekStart = new Date(curStart);
          for(let d=0; d<7; d++){
            const candidate = addDays(weekStart, d);
            const code = weekdayCode(candidate);
            if(days.includes(code)){
              const occStart = new Date(candidate); occStart.setHours(baseStart.getHours(), baseStart.getMinutes(), baseStart.getSeconds(), baseStart.getMilliseconds());
              const occEnd = new Date(occStart.getTime() + (baseEnd - baseStart));
              if(occStart < rangeEnd && occEnd > rangeStart) occs.push({start:occStart.toISOString(), end:occEnd.toISOString(), type:s.type, note:s.note, source:s.id});
            }
          }
          curStart = addDays(curStart, rec.interval*7);
        } else if(rec.type === 'monthly'){
          const day = baseStart.getDate();
          const occStart = new Date(curStart); occStart.setDate(day); occStart.setHours(baseStart.getHours(), baseStart.getMinutes(), baseStart.getSeconds(), baseStart.getMilliseconds());
          if(occStart.getDate() === day){
            const occEnd = new Date(occStart.getTime() + (baseEnd - baseStart));
            if(occStart < rangeEnd && occEnd > rangeStart) occs.push({start:occStart.toISOString(), end:occEnd.toISOString(), type:s.type, note:s.note, source:s.id});
          }
          curStart = addMonths(curStart, rec.interval);
        } else break;
        safety++;
      }
    }
  });
  occs.sort((a,b)=> new Date(a.start) - new Date(b.start));
  return occs;
}

/* Expand all occurrences */
function expandAllOccurrences(rangeStart, rangeEnd){
  const all = [];
  schedules.forEach(s=>{
    const occs = expandOccurrencesForLift(s.block, s.liftNo, rangeStart, rangeEnd);
    occs.forEach(o=> all.push(Object.assign({}, o, {block:s.block, liftNo:s.liftNo})));
  });
  return all;
}

/* ---------------- Utilities ---------------- */
function genId(){ return 'id_' + Math.random().toString(36).slice(2,9); }
function formatDateTime(v){ if(!v) return ''; const d = new Date(v); if(isNaN(d)) return v; return d.toLocaleString(); }
function formatTime(d){ if(!(d instanceof Date)) d = new Date(d); return d.toLocaleTimeString([], {hour:'2-digit', minute:'2-digit'}); }
function escapeHtml(s){ if(!s) return ''; return String(s).replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function humanDuration(start,end){ const ms = Math.max(0, end - start); const days = Math.floor(ms/(24*60*60*1000)); const hrs = Math.floor((ms%(24*60*60*1000))/(60*60*1000)); const mins = Math.floor((ms%(60*60*1000))/(60*1000)); if(days) return `${days}d ${hrs}h`; if(hrs) return `${hrs}h ${mins}m`; return `${mins}m`; }
function addDays(d,n){ const x=new Date(d); x.setDate(x.getDate()+n); return x; }
function addMonths(d,n){ const x=new Date(d); const m = x.getMonth()+n; x.setMonth(m); return x; }
function weekdayCode(date){ const codes=['SU','MO','TU','WE','TH','FR','SA']; return codes[date.getDay()]; }

/* ---------------- Count active occurrences ---------------- */
function countActiveOccurrences(atDate, typeFilter=null){
  const occs = expandAllOccurrences(atDate, atDate);
  return occs.filter(o=> (!typeFilter || o.type===typeFilter) && new Date(o.start) <= atDate && new Date(o.end) > atDate).length;
}

/* ---------------- CSV helpers ---------------- */
function toCSV(rows, columns){
  const esc = v=>{ if(v==null) return ''; const s=String(v); if(s.includes(',')||s.includes('"')||s.includes('\n')) return `"${s.replace(/"/g,'""')}"`; return s; };
  const header = columns.join(',') + '\n';
  const body = rows.map(r=> columns.map(c=> esc(r[c])).join(',')).join('\n');
  return header + body;
}
function readCSVFile(file, callback){
  const reader = new FileReader();
  reader.onload = e=>{ const text = e.target.result; const rows = parseCSV(text); callback(rows); };
  reader.readAsText(file);
}
function parseCSV(text){
  const lines = text.split(/\r?\n/).filter(l=>l.trim()!=='');
  if(!lines.length) return [];
  const header = splitCSVLine(lines.shift()).map(h=>h.trim());
  return lines.map(line=>{ const values = splitCSVLine(line); const obj={}; header.forEach((h,i)=> obj[h] = (values[i]||'').trim()); return obj; });
}
function splitCSVLine(line){
  const values=[]; let cur='', inQuotes=false;
  for(let i=0;i<line.length;i++){ const ch=line[i];
    if(ch==='"'){ if(inQuotes && line[i+1]==='"'){ cur+='"'; i++; continue; } inQuotes=!inQuotes; continue; }
    if(ch===',' && !inQuotes){ values.push(cur); cur=''; continue; }
    cur+=ch;
  }
  values.push(cur); return values;
}
function downloadBlob(blob, filename, mime){ const url = URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download=filename; a.click(); URL.revokeObjectURL(url); }

/* ---------------- Toast ---------------- */
function showToast(msg, timeout=2200){ UI.toastMsg.textContent = msg; UI.toast.hidden = false; clearTimeout(UI._toastTimer); UI._toastTimer = setTimeout(hideToast, timeout); }
function hideToast(){ UI.toast.hidden = true; clearTimeout(UI._toastTimer); }

/* ---------------- Helpers: populate filters ---------------- */
function populateFilterOptions(){
  const blocks = [...new Set(lifts.map(l=>l.block))].sort();
  const blockSelects = [UI.filterBlock, UI.scheduleFilterBlock, UI.ttFilterBlock];
  blockSelects.forEach(sel=>{ if(!sel) return; const prev = sel.value || ''; sel.innerHTML = '<option value="">All</option>'; blocks.forEach(b=>{ const opt=document.createElement('option'); opt.value=b; opt.textContent=b; sel.appendChild(opt); }); if(prev) sel.value = prev; });
}

/* ---------------- Initial render helpers ---------------- */
if(!localStorage.getItem(STORAGE.LIFTS)) save(STORAGE.LIFTS, []);
if(!localStorage.getItem(STORAGE.REPAIRS)) save(STORAGE.REPAIRS, []);
if(!localStorage.getItem(STORAGE.SCHEDULES)) save(STORAGE.SCHEDULES, []);
