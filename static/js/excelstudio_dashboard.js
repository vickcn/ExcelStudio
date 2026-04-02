(() => {
  const state = {
    parserResult: null,
    left: { hot: null, sheetName: null, workbook: null, file: null, workbookName: '', serverPath: '', annotRowIndex: null, annotColIndex: null },
    right: { hot: null, sheetName: null, workbook: null, file: null, workbookName: '', serverPath: '', annotRowIndex: null, annotColIndex: null },
    rulesFilePath: '',
    suspect: { sheetName: null, cells: [], index: -1 },
    syncScroll: { enabled: false, leftHandler: null, rightHandler: null, leftHolder: null, rightHolder: null },
  };

  const $ = (id) => document.getElementById(id);

  const els = {
    fileInputExcel: $('fileInputExcel'),
    rulesFileInput: $('rulesFileInput'),
    uploadDropzone: $('uploadDropzone'),
    uploadedFileInfo: $('uploadedFileInfo'),
    rulesFileInfo: $('rulesFileInfo'),
    healthDot: $('healthDot'),
    parserMetadata: $('parserMetadata'),
    parserPreview: $('parserPreview'),
    metricSheets: $('metricSheets'),
    metricTables: $('metricTables'),
    metricImages: $('metricImages'),
    metricFormulas: $('metricFormulas'),
    leftSheetSelect: $('leftSheetSelect'),
    rightSheetSelect: $('rightSheetSelect'),
    leftAnnotRowSelect: $('leftAnnotRowSelect'),
    leftAnnotColSelect: $('leftAnnotColSelect'),
    rightAnnotRowSelect: $('rightAnnotRowSelect'),
    rightAnnotColSelect: $('rightAnnotColSelect'),
    leftAnnotTip: $('leftAnnotTip'),
    rightAnnotTip: $('rightAnnotTip'),
    leftTableOverlay: $('leftTableOverlay'),
    rightTableOverlay: $('rightTableOverlay'),
    syncProgress: $('syncProgress'),
    apiResponseBox: $('apiResponseBox'),
    runLog: $('runLog'),
    rulesJsonModal: $('rulesJsonModal'),
    rulesJsonModalContent: $('rulesJsonModalContent'),
    rulesJsonModalClose: $('rulesJsonModalClose'),
    rulesJsonModalBackdrop: $('rulesJsonModalBackdrop'),
    actionButtons: [
      $('btnRunRuleDiscovery'),
      $('btnRunAudit'),
      $('btnRunFastMark'),
      $('btnRunRulesOnly'),
      $('btnRunFullFlow'),
    ],
    btnPickFileLeft: $('btnPickFileLeft'),
    btnPickFileRight: $('btnPickFileRight'),
    btnPickRulesFile: $('btnPickRulesFile'),
    btnDownloadRulesFile: $('btnDownloadRulesFile'),
  };

  function log(message, level = 'INFO') {
    const wrap = document.createElement('div');
    wrap.className = 'log-entry';
    const now = new Date();
    const time = now.toLocaleTimeString('zh-TW', { hour12: false });
    wrap.innerHTML = `<span class="log-time">${time}</span><span class="log-level">${level}</span><span>${escapeHtml(message)}</span>`;
    els.runLog.prepend(wrap);
  }

  function setApiResponse(obj) {
    els.apiResponseBox.textContent = typeof obj === 'string' ? obj : JSON.stringify(obj, null, 2);
  }

  function truncateText(text, limit = 1000) {
    if (text == null) return '';
    const raw = String(text);
    if (raw.length <= limit) return raw;
    return `${raw.slice(0, limit)}
... (truncated, see modal for full content)`;
  }

  function openRulesJsonModal(text) {
    if (!els.rulesJsonModal || !els.rulesJsonModalContent) return;
    els.rulesJsonModalContent.textContent = text ?? '';
    els.rulesJsonModal.classList.add('is-active');
    els.rulesJsonModal.setAttribute('aria-hidden', 'false');
  }

  function closeRulesJsonModal() {
    if (!els.rulesJsonModal) return;
    els.rulesJsonModal.classList.remove('is-active');
    els.rulesJsonModal.setAttribute('aria-hidden', 'true');
  }

  function escapeHtml(text) {
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function setWorkInProgress(flag) {
    if (els.syncProgress) {
      els.syncProgress.classList.toggle('is-active', flag);
      els.syncProgress.setAttribute('aria-hidden', String(!flag));
    }
    [els.leftTableOverlay, els.rightTableOverlay].forEach((overlay) => {
      if (!overlay) return;
      overlay.classList.toggle('is-active', flag);
      overlay.setAttribute('aria-hidden', String(!flag));
    });
  }

  async function checkHealth() {
    try {
      const data = await window.ExcelStudioApi.health();
      els.healthDot.className = 'status-dot ok';
      setApiResponse(data);
      log('API health 檢查成功');
    } catch (err) {
      els.healthDot.className = 'status-dot fail';
      setApiResponse(err.message || String(err));
      log(`API health 檢查失敗: ${err.message || err}`, 'ERROR');
    }
  }

  function setPickSideOnInput(side) {
    const inp = els.fileInputExcel;
    if (inp) inp.dataset.pickSide = side;
  }

  function bindFileUi() {
    const armPick = (side) => {
      setPickSideOnInput(side);
    };

    els.btnPickFileLeft?.addEventListener(
      'pointerdown',
      (e) => {
        e?.stopPropagation();
        armPick('left');
      },
      true,
    );
    els.btnPickFileRight?.addEventListener(
      'pointerdown',
      (e) => {
        e?.stopPropagation();
        armPick('right');
      },
      true,
    );

    els.btnPickFileLeft?.addEventListener('click', (e) => {
      e?.preventDefault();
      e?.stopPropagation();
      armPick('left');
      els.fileInputExcel?.click();
    });
    els.btnPickFileRight?.addEventListener('click', (e) => {
      e?.preventDefault();
      e?.stopPropagation();
      armPick('right');
      els.fileInputExcel?.click();
    });

    els.fileInputExcel?.addEventListener('change', async (e) => {
      const inp = e.target;
      const raw = inp?.dataset?.pickSide;
      const side = raw === 'right' ? 'right' : 'left';
      const file = inp.files?.[0];
      if (inp && inp.dataset) delete inp.dataset.pickSide;
      if (file) {
        log(side === 'left' ? `已選擇左表檔案：${file.name}` : `已選擇右表檔案：${file.name}`);
        await loadWorkbook(file, side);
        inp.value = '';
      }
    });

    els.btnPickRulesFile?.addEventListener('click', (e) => {
      e?.preventDefault();
      e?.stopPropagation();
      els.rulesFileInput?.click();
    });
    els.btnDownloadRulesFile?.addEventListener('click', async (e) => {
      e?.preventDefault();
      e?.stopPropagation();
      await downloadRulesFile();
    });
    els.rulesFileInput?.addEventListener('change', async (e) => {
      const file = e.target?.files?.[0];
      if (!file) return;
      log(`已選擇規則檔：${file.name}`);
      try {
        const result = await window.ExcelStudioApi.uploadRulesFile(file);
        if (result?.path) {
          state.rulesFilePath = result.path;
          updateRulesFileInfo();
        }
        log(`規則檔已上傳：${result?.path || file.name}`);
      } catch (err) {
        log(`規則檔上傳失敗: ${err.message || err}`, 'ERROR');
      } finally {
        e.target.value = '';
      }
    });

    // 勿在空白上傳區 click 開檔：使用者第二次常點標題/圖示，舊版會強制當左表。右表請按「載入右表」；左表請按「載入左表」或拖曳至此。

    ['dragenter', 'dragover'].forEach(eventName => {
      els.uploadDropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        els.uploadDropzone.classList.add('drag-over');
      });
    });
    ['dragleave', 'drop'].forEach(eventName => {
      els.uploadDropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        els.uploadDropzone.classList.remove('drag-over');
      });
    });
    els.uploadDropzone.addEventListener('drop', async (e) => {
      const file = e.dataTransfer?.files?.[0];
      if (file) await loadWorkbook(file, 'left');
    });

    updateRulesFileInfo();
  }

  function isEmptyValue(value) {
    if (value === null || value === undefined) return true;
    if (typeof value === 'string') return value.trim() === '';
    return false;
  }

  function toExcelColLabel(index) {
    let n = index + 1;
    let label = '';
    while (n > 0) {
      const rem = (n - 1) % 26;
      label = String.fromCharCode(65 + rem) + label;
      n = Math.floor((n - 1) / 26);
    }
    return label;
  }

  function getNonEmptyRowIndices(data) {
    if (!Array.isArray(data)) return [];
    const indices = [];
    data.forEach((row, r) => {
      if (!Array.isArray(row)) return;
      const hasValue = row.some(cell => !isEmptyValue(cell));
      if (hasValue) indices.push(r);
    });
    return indices;
  }

  function getNonEmptyColIndices(data) {
    if (!Array.isArray(data) || data.length === 0) return [];
    const maxCols = Math.max(...data.map(row => (Array.isArray(row) ? row.length : 0)), 0);
    const indices = [];
    for (let c = 0; c < maxCols; c += 1) {
      let hasValue = false;
      for (let r = 0; r < data.length; r += 1) {
        const row = data[r];
        if (Array.isArray(row) && !isEmptyValue(row[c])) {
          hasValue = true;
          break;
        }
      }
      if (hasValue) indices.push(c);
    }
    return indices;
  }

  function setSelectOptions(select, indices, labelFn) {
    if (!select) return;
    const prev = select.value;
    select.innerHTML = '';
    const emptyOpt = document.createElement('option');
    emptyOpt.value = '';
    emptyOpt.textContent = select.dataset?.placeholder || '';
    select.appendChild(emptyOpt);
    indices.forEach((idx) => {
      const opt = document.createElement('option');
      opt.value = String(idx);
      opt.textContent = labelFn(idx);
      select.appendChild(opt);
    });
    if (prev && indices.includes(Number(prev))) {
      select.value = prev;
    } else {
      select.value = '';
    }
  }

  function updateAnnotationOptions(side, data) {
    const paneState = state[side];
    const rowSelect = side === 'left' ? els.leftAnnotRowSelect : els.rightAnnotRowSelect;
    const colSelect = side === 'left' ? els.leftAnnotColSelect : els.rightAnnotColSelect;
    if (rowSelect) rowSelect.dataset.placeholder = '列索引';
    if (colSelect) colSelect.dataset.placeholder = '欄索引';
    const rowIndices = getNonEmptyRowIndices(data);
    const colIndices = getNonEmptyColIndices(data);
    setSelectOptions(rowSelect, rowIndices, (r) => `列 ${r + 1}`);
    setSelectOptions(colSelect, colIndices, (c) => `欄 ${toExcelColLabel(c)}`);

    const rowValue = rowSelect?.value;
    const colValue = colSelect?.value;
    paneState.annotRowIndex = rowValue === '' ? null : Number(rowValue);
    paneState.annotColIndex = colValue === '' ? null : Number(colValue);
  }

  function bindAnnotationControls() {
    const bindSide = (side) => {
      const paneState = state[side];
      const rowSelect = side === 'left' ? els.leftAnnotRowSelect : els.rightAnnotRowSelect;
      const colSelect = side === 'left' ? els.leftAnnotColSelect : els.rightAnnotColSelect;
      rowSelect?.addEventListener('change', () => {
        paneState.annotRowIndex = rowSelect.value === '' ? null : Number(rowSelect.value);
      });
      colSelect?.addEventListener('change', () => {
        paneState.annotColIndex = colSelect.value === '' ? null : Number(colSelect.value);
      });
    };
    bindSide('left');
    bindSide('right');
  }

  function hideAnnotationTip(side) {
    const tip = side === 'left' ? els.leftAnnotTip : els.rightAnnotTip;
    if (!tip) return;
    tip.classList.remove('is-active');
    tip.setAttribute('aria-hidden', 'true');
  }

  function showAnnotationTip(side, row, col, td) {
    const paneState = state[side];
    const tip = side === 'left' ? els.leftAnnotTip : els.rightAnnotTip;
    if (!tip || !paneState?.hot) return;

    const rowIdx = paneState.annotRowIndex;
    const colIdx = paneState.annotColIndex;
    if (rowIdx === null && colIdx === null) {
      hideAnnotationTip(side);
      return;
    }

    const data = paneState.hot.getData() || [];
    const lines = [];
    if (rowIdx !== null) {
      const value = data?.[rowIdx]?.[col];
      const text = isEmptyValue(value) ? '' : String(value);
      if (text !== '') lines.push(text);
    }
    if (colIdx !== null) {
      const value = data?.[row]?.[colIdx];
      const text = isEmptyValue(value) ? '' : String(value);
      if (text !== '') lines.push(text);
    }
    if (!lines.length) {
      hideAnnotationTip(side);
      return;
    }

    tip.innerHTML = lines.map(line => `<div>${escapeHtml(line)}</div>`).join('');
    const tdRect = td.getBoundingClientRect();
    const panelRect = paneState.hot.rootElement?.closest('.table-panel')?.getBoundingClientRect();
    if (panelRect) {
      const left = tdRect.left - panelRect.left + 12;
      const top = tdRect.top - panelRect.top + 12;
      tip.style.left = `${Math.max(8, left)}px`;
      tip.style.top = `${Math.max(8, top)}px`;
    }
    tip.classList.add('is-active');
    tip.setAttribute('aria-hidden', 'false');
  }

  async function downloadRulesFile() {
    const baseUrl = window.ExcelStudioApi.getBaseUrl();
    if (!baseUrl) {
      log('API Base URL 未設定', 'WARN');
      return;
    }
    const url = `${baseUrl}/api/rules/download`;
    try {
      const resp = await fetch(url);
      if (!resp.ok) {
        let detail = '';
        try { detail = await resp.text(); } catch {}
        throw new Error(detail || `下載規則檔失敗: ${resp.status}`);
      }
      const blob = await resp.blob();
      const disposition = resp.headers.get('content-disposition') || '';
      let filename = 'discovered_rules.json';
      const match = /filename\\*=UTF-8''([^;]+)|filename=\"?([^\";]+)\"?/i.exec(disposition);
      if (match) filename = decodeURIComponent(match[1] || match[2] || filename);
      const link = document.createElement('a');
      const href = URL.createObjectURL(blob);
      link.href = href;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(href);
      log(`規則檔已下載：${filename}`);
    } catch (err) {
      log(`規則檔下載失敗: ${err.message || err}`, 'ERROR');
    }
  }

  function updateUploadedFileInfo() {
    const leftFile = state.left.file;
    const rightFile = state.right.file;
    if (!leftFile && !rightFile) {
      els.uploadedFileInfo.classList.add('hidden');
      els.uploadedFileInfo.innerHTML = '';
      return;
    }
    const leftName = leftFile ? escapeHtml(leftFile.name) : '尚未載入';
    const rightName = rightFile ? escapeHtml(rightFile.name) : '尚未載入';
    const leftSize = leftFile ? `${(leftFile.size / 1024).toFixed(1)} KB` : '-';
    const rightSize = rightFile ? `${(rightFile.size / 1024).toFixed(1)} KB` : '-';
    const leftSheets = state.left.workbook?.SheetNames?.length ?? 0;
    const rightSheets = state.right.workbook?.SheetNames?.length ?? 0;
    els.uploadedFileInfo.classList.remove('hidden');
    els.uploadedFileInfo.innerHTML = [
      `左表: <b>${leftName}</b>（${leftSize}，工作表數: ${leftSheets}）`,
      `右表: <b>${rightName}</b>（${rightSize}，工作表數: ${rightSheets}）`,
    ].join('<br>');
  }

  function updateRulesFileInfo() {
    if (!els.rulesFileInfo) return;
    const defaultPath = els.rulesFileInfo.dataset?.defaultPath || '';
    if (!state.rulesFilePath && defaultPath) {
      state.rulesFilePath = defaultPath;
    }
    const displayPath = state.rulesFilePath || defaultPath || '尚未設定';
    els.rulesFileInfo.classList.remove('hidden');
    els.rulesFileInfo.textContent = `規則檔: ${displayPath}`;
  }

  function getSheetPayload(side) {
    const paneState = state[side];
    const sheetName = paneState.sheetName;
    if (!paneState.workbook || !sheetName) {
      return null;
    }
    const data = sheetToAoa(side, sheetName);
    return { sheet_name: sheetName, table: data };
  }

  async function loadWorkbook(file, side) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellStyles: true, cellFormula: true });
    const paneState = state[side];
    paneState.workbook = wb;
    paneState.file = file;
    paneState.workbookName = file.name;
    paneState.serverPath = '';
    paneState.sheetName = null;

    const names = wb.SheetNames || [];
    updateUploadedFileInfo();
    fillSheetSelect(side, names);
    if (names[0]) {
      renderSheetToPane(side, names[0]);
      log(`已載入${side === 'left' ? '左表' : '右表'} Excel: ${file.name}，工作表數=${names.length}`);
    }
  }

  function fillSheetSelect(side, sheetNames) {
    const select = side === 'left' ? els.leftSheetSelect : els.rightSheetSelect;
    select.innerHTML = '';
    if (!sheetNames || sheetNames.length === 0) {
      select.disabled = true;
      return;
    }
    select.disabled = false;
    sheetNames.forEach((name) => {
      const opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      select.appendChild(opt);
    });
    if (side === 'left') {
      els.leftSheetSelect.onchange = () => renderSheetToPane('left', els.leftSheetSelect.value);
    } else {
      els.rightSheetSelect.onchange = () => renderSheetToPane('right', els.rightSheetSelect.value);
    }
  }

  function sheetToAoa(side, sheetName) {
    const paneState = state[side];
    const sheet = paneState.workbook?.Sheets?.[sheetName];
    if (!sheet) return [[]];
    return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: true });
  }

  function renderSheetToPane(side, sheetName) {
    const paneState = state[side];
    if (!paneState.workbook) {
      log(`尚未載入${side === 'left' ? '左表' : '右表'} Excel`, 'WARN');
      return;
    }
    const data = sheetToAoa(side, sheetName);
    const targetId = side === 'left' ? 'leftHot' : 'rightHot';
    const target = $(targetId);
    const dataClone = data.map(row => Array.isArray(row) ? [...row] : [row]);

    if (paneState.hot) paneState.hot.destroy();
    paneState.hot = new Handsontable(target, {
      data: dataClone,
      rowHeaders: true,
      colHeaders: true,
      height: '100%',
      width: '100%',
      licenseKey: 'non-commercial-and-evaluation',
      stretchH: 'all',
      manualColumnResize: true,
      manualRowResize: true,
      filters: true,
      dropdownMenu: true,
      columnSorting: true,
      contextMenu: true,
      formulas: false,
      autoWrapRow: true,
      autoWrapCol: true,
      afterChange: (changes, source) => {
        if (!changes || source === 'loadData') return;
        syncPaneBackToWorkbook(side);
      },
      afterOnCellMouseOver: (event, coords, td) => {
        if (!td) return;
        if (coords.row < 0 || coords.col < 0) return;
        showAnnotationTip(side, coords.row, coords.col, td);
      },
      afterOnCellMouseOut: () => {
        hideAnnotationTip(side);
      },
    });
    paneState.sheetName = sheetName;
    if (side === 'left') els.leftSheetSelect.value = sheetName;
    if (side === 'right') els.rightSheetSelect.value = sheetName;
    updateAnnotationOptions(side, dataClone);
    applySuspectToPane(paneState, sheetName);
    if (state.syncScroll.enabled) {
      disableScrollSync();
      enableScrollSync();
    }
  }

  bindAnnotationControls();

  function clearSuspectFromPane(paneState) {
    if (!paneState?.hot || !paneState.suspectCells?.length) return;
    paneState.suspectCells.forEach(({ r, c }) => {
      paneState.hot.setCellMeta(r, c, 'className', null);
    });
    paneState.suspectCells = [];
    paneState.hot.render();
  }

  function applySuspectToPane(paneState, sheetName) {
    if (!paneState?.hot) return;
    clearSuspectFromPane(paneState);
    const cells = state.suspect.cells || [];
    paneState.suspectCells = [];
    cells.forEach((cell) => {
      const r = Number(cell.row) - 1;
      const c = Number(cell.col) - 1;
      if (Number.isNaN(r) || Number.isNaN(c) || r < 0 || c < 0) return;
      paneState.hot.setCellMeta(r, c, 'className', 'hot-suspect');
      paneState.suspectCells.push({ r, c });
    });
    paneState.hot.render();
  }

    function getScrollHolders() {
    const leftHot = state.left.hot;
    const rightHot = state.right.hot;
    if (!leftHot || !rightHot) return null;
    const leftHolder = leftHot.rootElement?.querySelector('.wtHolder');
    const rightHolder = rightHot.rootElement?.querySelector('.wtHolder');
    if (!leftHolder || !rightHolder) return null;
    return { leftHolder, rightHolder };
  }

  function updateSyncButton() {
    const btn = $('btnSyncPanes');
    if (!btn) return;
    btn.textContent = state.syncScroll.enabled ? '解除同步' : '同步欄列';
  }

  function enableScrollSync() {
    const holders = getScrollHolders();
    if (!holders) {
      log('左右表尚未載入或無法取得捲動容器', 'WARN');
      return;
    }
    const { leftHolder, rightHolder } = holders;
    const lock = { active: false };
    const syncFromLeft = () => {
      if (lock.active) return;
      lock.active = true;
      rightHolder.scrollTop = leftHolder.scrollTop;
      rightHolder.scrollLeft = leftHolder.scrollLeft;
      lock.active = false;
    };
    const syncFromRight = () => {
      if (lock.active) return;
      lock.active = true;
      leftHolder.scrollTop = rightHolder.scrollTop;
      leftHolder.scrollLeft = rightHolder.scrollLeft;
      lock.active = false;
    };
    leftHolder.addEventListener('scroll', syncFromLeft);
    rightHolder.addEventListener('scroll', syncFromRight);
    state.syncScroll.enabled = true;
    state.syncScroll.leftHandler = syncFromLeft;
    state.syncScroll.rightHandler = syncFromRight;
    state.syncScroll.leftHolder = leftHolder;
    state.syncScroll.rightHolder = rightHolder;
    updateSyncButton();
    syncFromLeft();
    log('已啟用欄列同步');
  }

  function disableScrollSync() {
    if (!state.syncScroll.enabled) {
      updateSyncButton();
      return;
    }
    const { leftHolder, rightHolder, leftHandler, rightHandler } = state.syncScroll;
    if (leftHolder && leftHandler) leftHolder.removeEventListener('scroll', leftHandler);
    if (rightHolder && rightHandler) rightHolder.removeEventListener('scroll', rightHandler);
    state.syncScroll.enabled = false;
    state.syncScroll.leftHandler = null;
    state.syncScroll.rightHandler = null;
    state.syncScroll.leftHolder = null;
    state.syncScroll.rightHolder = null;
    updateSyncButton();
    log('已解除欄列同步');
  }

  function syncPaneViewport() {
    if (state.syncScroll.enabled) {
      disableScrollSync();
      return;
    }
    enableScrollSync();
  }

  function centerCellInPane(pane, r, c) {
    if (!pane?.hot) return;
    pane.hot.scrollViewportTo(r, c);
    pane.hot.selectCell(r, c, r, c, true);
    requestAnimationFrame(() => {
      const holder = pane.hot.rootElement?.querySelector('.wtHolder');
      const cell = pane.hot.getCell(r, c);
      if (!holder || !cell) return;
      const holderRect = holder.getBoundingClientRect();
      const cellRect = cell.getBoundingClientRect();
      const offsetTop = cellRect.top - holderRect.top + holder.scrollTop;
      const offsetLeft = cellRect.left - holderRect.left + holder.scrollLeft;
      const targetTop = offsetTop - holder.clientHeight / 2 + cellRect.height / 2;
      const targetLeft = offsetLeft - holder.clientWidth / 2 + cellRect.width / 2;
      holder.scrollTop = Math.max(0, targetTop);
      holder.scrollLeft = Math.max(0, targetLeft);
    });
  }

  function jumpToNextSuspect() {
    const cells = state.suspect.cells || [];
    const sheetName = state.suspect.sheetName;
    if (!sheetName || cells.length === 0) {
      window.alert('沒有疑慮點');
      return;
    }
    let nextIndex = state.suspect.index + 1;
    if (nextIndex >= cells.length) {
      window.alert('已經沒有其他疑慮點');
      nextIndex = 0;
    }
    state.suspect.index = nextIndex;
    const cell = cells[nextIndex] || {};
    const r = Number(cell.row) - 1;
    const c = Number(cell.col) - 1;
    if (Number.isNaN(r) || Number.isNaN(c) || r < 0 || c < 0) {
      return;
    }

    const panes = [];
    if (state.left.sheetName === sheetName) panes.push(state.left);
    if (state.right.sheetName === sheetName) panes.push(state.right);
    if (panes.length === 0) {
      renderSheetToPane('left', sheetName);
      panes.push(state.left);
    }

    panes.forEach((pane) => centerCellInPane(pane, r, c));
  }

  function syncPaneBackToWorkbook(side) {
    const paneState = state[side];
    if (!paneState.hot || !paneState.sheetName) return;
    const aoa = paneState.hot.getData();
    const newSheet = XLSX.utils.aoa_to_sheet(aoa);
    if (paneState.workbook) {
      paneState.workbook.Sheets[paneState.sheetName] = newSheet;
    }
  }

  function exportPaneXlsx(side) {
    const paneState = state[side];
    if (!paneState.hot || !paneState.sheetName) return;
    const aoa = paneState.hot.getData();
    const sheet = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    const sheetName = paneState.sheetName || (side === 'left' ? '左表' : '右表');
    XLSX.utils.book_append_sheet(wb, sheet, sheetName);
    const wbArray = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbArray], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const filename = side === 'left' ? '左表.xlsx' : '右表.xlsx';
    saveAs(blob, filename);
    log(`已匯出${side === 'left' ? '左表' : '右表'} Excel: ${sheetName}`);
  }

  async function runParserPreview() {
    const sourceSide = state.left.file ? 'left' : (state.right.file ? 'right' : null);
    if (!sourceSide) {
      log('尚未載入左表或右表 Excel', 'WARN');
      return;
    }
    const sourceState = state[sourceSide];
    const basic = {
      file_name: sourceState.file.name,
      size_kb: +(sourceState.file.size / 1024).toFixed(1),
      sheets: sourceState.workbook?.SheetNames || [],
    };
    const previewData = {};
    (sourceState.workbook?.SheetNames || []).slice(0, 3).forEach((name) => {
      const aoa = sheetToAoa(sourceSide, name).slice(0, 8).map(row => row.slice(0, 8));
      previewData[name] = aoa;
    });

    els.parserMetadata.textContent = JSON.stringify({
      note: `此處顯示${sourceSide === 'left' ? '左表' : '右表'}可讀取的 workbook 結構；正式串接時請改呼叫 xlsx_parser process/preview API 回傳內容`,
      ...basic,
    }, null, 2);
    els.parserPreview.textContent = JSON.stringify(previewData, null, 2);

    const sheetCount = sourceState.workbook?.SheetNames?.length || 0;
    els.metricSheets.textContent = String(sheetCount);
    els.metricTables.textContent = sheetCount > 0 ? String(sheetCount) : '0';
    els.metricImages.textContent = '-';
    els.metricFormulas.textContent = '-';

    log(`已完成${sourceSide === 'left' ? '左表' : '右表'} workbook 初步分析（尚未串接 xlsx_parser API）`);
  }

  function getCommonParams() {
    return {
      row_name: $('rowName').value.trim() || null,
      window_height: Number($('windowHeight').value || 3),
      window_width: Number($('windowWidth').value || 1),
      tolerance: Number($('tolerance').value || 0.01),
      strict_row_match: $('strictRowMatch').checked,
      consistency_threshold: Number($('consistencyThreshold').value || 0.8),
      quick_scan_threshold: Number($('quickScanThreshold').value || 3),
    };
  }

  function guessOutputPath(name) {
    return `outputs/${name}`;
  }

  async function runAction(action) {
    const setBusy = (flag) => {
      els.actionButtons?.forEach(btn => btn && (btn.disabled = flag));
      setWorkInProgress(flag);
    };
    setBusy(true);

    const endpointMap = {
      ruleDiscovery: '/api/rules/discover-json',
      audit: '/api/audit-json',
      markFast: '/api/audit-json',
      rulesOnly: '/api/audit-json',
      fullFlow: '/api/full-flow-json',
    };
    const endpoint = endpointMap[action] || '(unknown)';
    const baseUrl = window.ExcelStudioApi?.getBaseUrl?.() || '';

    const params = getCommonParams();
    const leftPayload = getSheetPayload('left');
    const rightPayload = getSheetPayload('right');
    if (action === 'ruleDiscovery' && !leftPayload) {
      log('規則偵測需要左表檔案，請先載入左表。', 'WARN');
      setBusy(false);
      return;
    }
    if (['audit', 'markFast', 'rulesOnly'].includes(action) && !rightPayload) {
      log('稽核需要右表檔案，請先載入右表。', 'WARN');
      setBusy(false);
      return;
    }
    if (['audit', 'markFast', 'rulesOnly'].includes(action) && !state.rulesFilePath) {
      log('尚未設定規則檔，請先上傳規則檔或使用預設路徑。', 'WARN');
      setBusy(false);
      return;
    }
    if (action === 'fullFlow' && (!leftPayload || !rightPayload)) {
      log('完整流程需要左表（規則偵測）與右表（稽核）。', 'WARN');
      setBusy(false);
      return;
    }

    let payload;
    let exec;
    if (action === 'ruleDiscovery') {
      payload = {
        baseline_table: leftPayload?.table || [],
        baseline_sheet_name: leftPayload?.sheet_name || null,
        start_loc_row_name: params.row_name,
        window_height: params.window_height,
        window_width: params.window_width,
        consistency_threshold: params.consistency_threshold,
        quick_scan_threshold: params.quick_scan_threshold,
      };
      exec = () => window.ExcelStudioApi.ruleDiscoveryJson(payload);
    } else if (action === 'audit') {
      payload = {
        detect_rules_file: state.rulesFilePath,
        target_table: rightPayload?.table || [],
        target_sheet_name: rightPayload?.sheet_name || null,
        row_name: params.row_name,
        window_height: params.window_height,
        window_width: params.window_width,
        tolerance: params.tolerance,
        strict_row_match: params.strict_row_match,
      };
      exec = () => window.ExcelStudioApi.auditJson(payload);
    } else if (action === 'markFast') {
      payload = {
        detect_rules_file: state.rulesFilePath,
        target_table: rightPayload?.table || [],
        target_sheet_name: rightPayload?.sheet_name || null,
        row_name: params.row_name,
        window_height: params.window_height,
        window_width: params.window_width,
        tolerance: params.tolerance,
        strict_row_match: params.strict_row_match,
      };
      exec = () => window.ExcelStudioApi.auditJson(payload);
    } else if (action === 'rulesOnly') {
      payload = {
        detect_rules_file: state.rulesFilePath,
        target_table: rightPayload?.table || [],
        target_sheet_name: rightPayload?.sheet_name || null,
        row_name: params.row_name,
        window_height: params.window_height,
        window_width: params.window_width,
        tolerance: params.tolerance,
        strict_row_match: params.strict_row_match,
      };
      exec = () => window.ExcelStudioApi.auditJson(payload);
    } else if (action === 'fullFlow') {
      payload = {
        baseline_table: leftPayload?.table || [],
        baseline_sheet_name: leftPayload?.sheet_name || null,
        target_table: rightPayload?.table || [],
        target_sheet_name: rightPayload?.sheet_name || null,
        row_name: params.row_name,
        window_height: params.window_height,
        window_width: params.window_width,
        tolerance: params.tolerance,
        strict_row_match: params.strict_row_match,
        consistency_threshold: params.consistency_threshold,
        quick_scan_threshold: params.quick_scan_threshold,
      };
      exec = () => window.ExcelStudioApi.fullFlowJson(payload);
    } else {
      throw new Error(`未知的 action: ${action}`);
    }

    setApiResponse(payload);
    const leftName = state.left.file?.name || '-';
    const rightName = state.right.file?.name || '-';
    log(`執行 ${action} API 呼叫 -> ${endpoint} ${baseUrl ? `(${baseUrl})` : ''} | 左檔=${leftName} 右檔=${rightName}`);
    try {
      const result = await exec();
      setApiResponse(result);
      if ((action === 'ruleDiscovery' || action === 'fullFlow') && result?.detect_rules_file) {
        state.rulesFilePath = result.detect_rules_file;
        updateRulesFileInfo();
        log(`規則檔已更新：${result.detect_rules_file}`);
      } else if (action === 'fullFlow' && result?.discovery?.detect_rules_file) {
        state.rulesFilePath = result.discovery.detect_rules_file;
        updateRulesFileInfo();
        log(`規則檔已更新：${result.discovery.detect_rules_file}`);
      }
      log(`${action} 完成`);
      if (action === 'audit' || action === 'fullFlow') {
        const sheetName = result?.suspect_sheet || result?.mark_summary?.suspect_sheet || null;
        const cells = result?.suspect_cells || result?.mark_summary?.suspect_cells || [];
        if (sheetName && Array.isArray(cells)) {
          state.suspect.sheetName = sheetName;
          state.suspect.cells = cells;
          state.suspect.index = -1;
          applySuspectToPane(state.left, state.left.sheetName);
          applySuspectToPane(state.right, state.right.sheetName);
          log(`已標記疑似儲存格：${cells.length}`);
        } else {
          state.suspect.sheetName = null;
          state.suspect.cells = [];
          state.suspect.index = -1;
          clearSuspectFromPane(state.left);
          clearSuspectFromPane(state.right);
        }
      }
        } catch (err) {
      setApiResponse(err.message || String(err));
      log(`${action} failed: ${err.message || err}`, 'ERROR');
    } finally {
      setBusy(false);
    }
  }

  function bindActions() {
    $('btnCheckHealth').addEventListener('click', checkHealth);
    $('btnParserPreview').addEventListener('click', runParserPreview);
    $('btnShowRulesJson').addEventListener('click', async () => {
      try {
        const data = await window.ExcelStudioApi.rulesJson();
        const fullText = typeof data === 'string' ? data : JSON.stringify(data?.data ?? data, null, 2);
        setApiResponse(truncateText(fullText, 1000));
        openRulesJsonModal(fullText);
        log('Rules JSON loaded (modal opened).');
      } catch (err) {
        setApiResponse(err.message || String(err));
        log(`Rules JSON failed: ${err.message || err}`, 'ERROR');
      }
    });
    $('btnClearApiResponse').addEventListener('click', async () => {
      try {
        const result = await window.ExcelStudioApi.clearOutputs();
        setApiResponse('Cleared.');
        log(`Outputs cleared: ${(result?.removed || []).length}`);
      } catch (err) {
        setApiResponse(err.message || String(err));
        log(`Clear outputs failed: ${err.message || err}`, 'ERROR');
      }
    });
    $('btnReloadLeft').addEventListener('click', () => state.left.sheetName && renderSheetToPane('left', state.left.sheetName));
    $('btnReloadRight').addEventListener('click', () => state.right.sheetName && renderSheetToPane('right', state.right.sheetName));
    $('btnSyncPanes').addEventListener('click', syncPaneViewport);
    $('btnNextSuspect').addEventListener('click', jumpToNextSuspect);
    $('btnExportLeftCsv').addEventListener('click', () => exportPaneXlsx('left'));
    $('btnExportRightCsv').addEventListener('click', () => exportPaneXlsx('right'));

    $('btnRunRuleDiscovery').addEventListener('click', () => runAction('ruleDiscovery'));
    $('btnRunAudit').addEventListener('click', () => runAction('audit'));
    $('btnRunFastMark').addEventListener('click', () => runAction('markFast'));
    $('btnRunRulesOnly').addEventListener('click', () => runAction('rulesOnly'));
    $('btnRunFullFlow').addEventListener('click', () => runAction('fullFlow'));

    els.rulesJsonModalClose?.addEventListener('click', closeRulesJsonModal);
    els.rulesJsonModalBackdrop?.addEventListener('click', closeRulesJsonModal);
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') closeRulesJsonModal();
    });
  }

  function initParserPanelToggle() {
    const panel = document.getElementById('parserPanel');
    const toggle = document.getElementById('toggleParserPanel');
    if (!panel || !toggle) return;
    const updateLabel = () => {
      const collapsed = panel.classList.contains('is-collapsed');
      toggle.textContent = collapsed ? '>' : 'v';
      toggle.setAttribute('aria-expanded', String(!collapsed));
    };
    toggle.addEventListener('click', () => {
      panel.classList.toggle('is-collapsed');
      updateLabel();
    });
    updateLabel();
  }

  function initQuickPanelToggle() {
    const panel = document.getElementById('quickPanel');
    const toggle = document.getElementById('toggleQuickPanel');
    if (!panel || !toggle) return;
    const updateLabel = () => {
      const collapsed = panel.classList.contains('is-collapsed');
      toggle.textContent = collapsed ? '>' : 'v';
      toggle.setAttribute('aria-expanded', String(!collapsed));
    };
    toggle.addEventListener('click', () => {
      panel.classList.toggle('is-collapsed');
      updateLabel();
      // const collapsed = panel.classList.contains('is-collapsed');
      // log(`快速參數${collapsed ? '已收合' : '已展開'}`);
    });
    updateLabel();
  }

  function initApiPanelToggle() {
    const panel = document.getElementById('apiPanel');
    const toggle = document.getElementById('toggleApiPanel');
    if (!panel || !toggle) return;
    const updateLabel = () => {
      const collapsed = panel.classList.contains('is-collapsed');
      toggle.textContent = collapsed ? '>' : 'v';
      toggle.setAttribute('aria-expanded', String(!collapsed));
    };
    toggle.addEventListener('click', () => {
      panel.classList.toggle('is-collapsed');
      updateLabel();
    });
    updateLabel();
  }

  function initLogPanelToggle() {
    const panel = document.getElementById('logPanel');
    const toggle = document.getElementById('toggleLogPanel');
    if (!panel || !toggle) return;
    const updateLabel = () => {
      const collapsed = panel.classList.contains('is-collapsed');
      toggle.textContent = collapsed ? '>' : 'v';
      toggle.setAttribute('aria-expanded', String(!collapsed));
    };
    toggle.addEventListener('click', () => {
      panel.classList.toggle('is-collapsed');
      updateLabel();
    });
    updateLabel();
  }

  bindFileUi();
  bindActions();
  initQuickPanelToggle();
  initParserPanelToggle();
  initApiPanelToggle();
  initLogPanelToggle();
  checkHealth();
})();





