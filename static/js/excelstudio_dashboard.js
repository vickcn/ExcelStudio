(() => {
  const state = {
    parserResult: null,
    left: { hot: null, sheetName: null, workbook: null, file: null, workbookName: '', serverPath: '', annotRowIndex: null, annotColIndex: null },
    right: { hot: null, sheetName: null, workbook: null, file: null, workbookName: '', serverPath: '', annotRowIndex: null, annotColIndex: null },
    rulesFilePath: '',
    suspect: { sheetName: null, cells: [], index: -1 },
    syncScroll: { enabled: false, leftHandler: null, rightHandler: null, leftHolder: null, rightHolder: null },
    task: { id: null, action: null, polling: null, status: 'idle' },
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
    syncProgressText: $('syncProgressText'),
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
    btnStopTask: $('btnStopTask'),
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

  function setActionButtonsDisabled(flag) {
    els.actionButtons?.forEach(btn => btn && (btn.disabled = flag));
  }

  function setWorkInProgress(flag, options = {}) {
    const { message = '??銝?, percent = 0, canStop = false } = options;
    if (els.syncProgress) {
      els.syncProgress.classList.toggle('is-active', flag);
      els.syncProgress.setAttribute('aria-hidden', String(!flag));
    }
    if (els.syncProgressText) {
      if (flag) {
        const pct = Number.isFinite(percent) ? Math.round(percent) : 0;
        els.syncProgressText.textContent = `${message} ${pct}%`;
      } else {
        els.syncProgressText.textContent = '??銝?0%';
      }
    }
    if (els.btnStopTask) {
      els.btnStopTask.disabled = !canStop;
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
      log('API health 瑼Ｘ??');
    } catch (err) {
      els.healthDot.className = 'status-dot fail';
      setApiResponse(err.message || String(err));
      log(`API health 瑼Ｘ憭望?: ${err.message || err}`, 'ERROR');
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
        log(side === 'left' ? `撌脤?椰銵冽?獢?${file.name}` : `撌脤?銵冽?獢?${file.name}`);
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
      log(`撌脤????嚗?{file.name}`);
      try {
        const result = await window.ExcelStudioApi.uploadRulesFile(file);
        if (result?.path) {
          state.rulesFilePath = result.path;
          updateRulesFileInfo();
        }
        log(`閬?瑼歇銝嚗?{result?.path || file.name}`);
      } catch (err) {
        log(`閬?瑼??喳仃?? ${err.message || err}`, 'ERROR');
      } finally {
        e.target.value = '';
      }
    });

    // ?踹蝛箇銝? click ??嚗蝙?刻洵鈭活撣賊?璅?/?內嚗???撘瑕?嗅椰銵具銵刻????亙銵具?撌西”隢????亙椰銵具???單迨??

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
    if (rowSelect) rowSelect.dataset.placeholder = '?揣撘?;
    if (colSelect) colSelect.dataset.placeholder = '甈揣撘?;
    const rowIndices = getNonEmptyRowIndices(data);
    const colIndices = getNonEmptyColIndices(data);
    setSelectOptions(rowSelect, rowIndices, (r) => `??${r + 1}`);
    setSelectOptions(colSelect, colIndices, (c) => `甈?${toExcelColLabel(c)}`);

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
      log('API Base URL ?芾身摰?, 'WARN');
      return;
    }
    const url = `${baseUrl}/api/rules/download`;
    try {
      const resp = await fetch(url);
      if (!resp.ok) {
        let detail = '';
        try { detail = await resp.text(); } catch {}
        throw new Error(detail || `銝?閬?瑼仃?? ${resp.status}`);
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
      log(`閬?瑼歇銝?嚗?{filename}`);
    } catch (err) {
      log(`閬?瑼?頛仃?? ${err.message || err}`, 'ERROR');
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
    const leftName = leftFile ? escapeHtml(leftFile.name) : '撠頛';
    const rightName = rightFile ? escapeHtml(rightFile.name) : '撠頛';
    const leftSize = leftFile ? `${(leftFile.size / 1024).toFixed(1)} KB` : '-';
    const rightSize = rightFile ? `${(rightFile.size / 1024).toFixed(1)} KB` : '-';
    const leftSheets = state.left.workbook?.SheetNames?.length ?? 0;
    const rightSheets = state.right.workbook?.SheetNames?.length ?? 0;
    els.uploadedFileInfo.classList.remove('hidden');
    els.uploadedFileInfo.innerHTML = [
      `撌西”: <b>${leftName}</b>嚗?{leftSize}嚗極雿”?? ${leftSheets}嚗,
      `?唾”: <b>${rightName}</b>嚗?{rightSize}嚗極雿”?? ${rightSheets}嚗,
    ].join('<br>');
  }

  function updateRulesFileInfo() {
    if (!els.rulesFileInfo) return;
    const defaultPath = els.rulesFileInfo.dataset?.defaultPath || '';
    if (!state.rulesFilePath && defaultPath) {
      state.rulesFilePath = defaultPath;
    }
    const displayPath = state.rulesFilePath || defaultPath || '撠閮剖?';
    els.rulesFileInfo.classList.remove('hidden');
    els.rulesFileInfo.textContent = `閬?瑼? ${displayPath}`;
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
      log(`撌脰???{side === 'left' ? '撌西”' : '?唾”'} Excel: ${file.name}嚗極雿”??${names.length}`);
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
      log(`撠頛${side === 'left' ? '撌西”' : '?唾”'} Excel`, 'WARN');
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
    btn.textContent = state.syncScroll.enabled ? '閫??郊' : '?郊甈?';
  }

  function enableScrollSync() {
    const holders = getScrollHolders();
    if (!holders) {
      log('撌血銵典??芾??交??⊥????脣?摰孵', 'WARN');
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
    log('撌脣??冽???甇?);
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
    log('撌脰圾?斗???甇?);
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
      window.alert('瘝??暺?);
      return;
    }
    let nextIndex = state.suspect.index + 1;
    if (nextIndex >= cells.length) {
      window.alert('撌脩?瘝??嗡??暺?);
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
    const sheetName = paneState.sheetName || (side === 'left' ? '撌西”' : '?唾”');
    XLSX.utils.book_append_sheet(wb, sheet, sheetName);
    const wbArray = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbArray], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const filename = side === 'left' ? '撌西”.xlsx' : '?唾”.xlsx';
    saveAs(blob, filename);
    log(`撌脣??{side === 'left' ? '撌西”' : '?唾”'} Excel: ${sheetName}`);
  }

  async function runParserPreview() {
    const sourceSide = state.left.file ? 'left' : (state.right.file ? 'right' : null);
    if (!sourceSide) {
      log('撠頛撌西”?銵?Excel', 'WARN');
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
      note: `甇方?憿舐內${sourceSide === 'left' ? '撌西”' : '?唾”'}?航??? workbook 蝯?嚗迤撘葡?交?隢?澆 xlsx_parser process/preview API ??批捆`,
      ...basic,
    }, null, 2);
    els.parserPreview.textContent = JSON.stringify(previewData, null, 2);

    const sheetCount = sourceState.workbook?.SheetNames?.length || 0;
    els.metricSheets.textContent = String(sheetCount);
    els.metricTables.textContent = sheetCount > 0 ? String(sheetCount) : '0';
    els.metricImages.textContent = '-';
    els.metricFormulas.textContent = '-';

    log(`撌脣???{sourceSide === 'left' ? '撌西”' : '?唾”'} workbook ?郊??嚗??芯葡??xlsx_parser API嚗);
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

  function clearTaskPolling() {
    if (state.task.polling) {
      clearInterval(state.task.polling);
      state.task.polling = null;
    }
  }

  function resetTaskState() {
    state.task.id = null;
    state.task.action = null;
    state.task.status = 'idle';
  }

  async function pollTaskProgress() {
    if (!state.task.id) return;
    try {
      const data = await window.ExcelStudioApi.taskProgress(state.task.id);
      if (!data?.success) return;
      const status = data.status || 'running';
      const percent = Number.isFinite(data.progress) ? data.progress : 0;
      const message = status === 'cancel_requested' ? '銝剜迫銝? : (data.message || '??銝?);
      const canStop = status === 'running';
      setWorkInProgress(true, { message, percent, canStop });
      state.task.status = status;

      if (status === 'done') {
        clearTaskPolling();
        setWorkInProgress(false);
        setActionButtonsDisabled(false);
        applyActionResult(state.task.action, data.result);
        resetTaskState();
      } else if (status === 'error') {
        clearTaskPolling();
        setWorkInProgress(false);
        setActionButtonsDisabled(false);
        setApiResponse(data.error || data.message || '隞餃?憭望?');
        log(`${state.task.action} 憭望?: ${data.message || 'unknown error'}`, 'ERROR');
        resetTaskState();
      } else if (status === 'cancelled') {
        clearTaskPolling();
        setWorkInProgress(false);
        setActionButtonsDisabled(false);
        log('隞餃?撌脖葉甇?, 'WARN');
        resetTaskState();
      }
    } catch (err) {
      log(`?脣漲?亥岷憭望?: ${err.message || err}`, 'ERROR');
    }
  }

  async function startTask(action, payload, endpoint, baseUrl, leftName, rightName) {
    setApiResponse(payload);
    log(`?瑁? ${action} API ?澆 -> ${endpoint} ${baseUrl ? `(${baseUrl})` : ''} | 撌行?=${leftName} ?單?=${rightName}`);
    try {
      const result = await window.ExcelStudioApi.taskStart(action, payload);
      if (!result?.task_id) {
        throw new Error('task_id 蝻箏仃嚗瘜蕭頩日脣漲');
      }
      state.task.id = result.task_id;
      state.task.action = action;
      state.task.status = 'running';
      setWorkInProgress(true, { message: '??銝?, percent: 1, canStop: true });
      clearTaskPolling();
      state.task.polling = setInterval(pollTaskProgress, 1000);
      await pollTaskProgress();
    } catch (err) {
      setApiResponse(err.message || String(err));
      log(`${action} ??憭望?: ${err.message || err}`, 'ERROR');
      setWorkInProgress(false);
      setActionButtonsDisabled(false);
      resetTaskState();
    }
  }

  async function stopCurrentTask() {
    if (!state.task.id) {
      log('?桀?瘝??瑁?銝剔?隞餃?', 'WARN');
      return;
    }
    try {
      await window.ExcelStudioApi.taskStop(state.task.id);
      log('撌脤銝剜迫隢?');
    } catch (err) {
      log(`銝剜迫憭望?: ${err.message || err}`, 'ERROR');
    }
  }

  function applyActionResult(action, result) {
    if (result !== undefined) {
      setApiResponse(result);
    }
    if ((action === 'ruleDiscovery' || action === 'fullFlow') && result?.detect_rules_file) {
      state.rulesFilePath = result.detect_rules_file;
      updateRulesFileInfo();
      log(`閬?瑼歇?湔: ${result.detect_rules_file}`);
    } else if (action === 'fullFlow' && result?.discovery?.detect_rules_file) {
      state.rulesFilePath = result.discovery.detect_rules_file;
      updateRulesFileInfo();
      log(`閬?瑼歇?湔: ${result.discovery.detect_rules_file}`);
    }
    log(`${action} 摰?`);
    if (['audit', 'fullFlow', 'markFast', 'rulesOnly'].includes(action)) {
      const sheetName = result?.suspect_sheet || result?.mark_summary?.suspect_sheet || null;
      const cells = result?.suspect_cells || result?.mark_summary?.suspect_cells || [];
      if (sheetName && Array.isArray(cells)) {
        state.suspect.sheetName = sheetName;
        state.suspect.cells = cells;
        state.suspect.index = -1;
        applySuspectToPane(state.left, state.left.sheetName);
        applySuspectToPane(state.right, state.right.sheetName);
        log(`??脣??? ${cells.length}`);
      } else {
        state.suspect.sheetName = null;
        state.suspect.cells = [];
        state.suspect.index = -1;
        clearSuspectFromPane(state.left);
        clearSuspectFromPane(state.right);
      }
    }
  }

  async function runAction(action) {
    if (state.task.id) {
      log('已有任務執行中，請先完成或中止', 'WARN');
      return;
    }
    setActionButtonsDisabled(true);

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
      log('規則偵測需要左表資料', 'WARN');
      setActionButtonsDisabled(false);
      return;
    }
    if (['audit', 'markFast', 'rulesOnly'].includes(action) && !rightPayload) {
      log('稽核需要右表資料', 'WARN');
      setActionButtonsDisabled(false);
      return;
    }
    if (['audit', 'markFast', 'rulesOnly'].includes(action) && !state.rulesFilePath) {
      log('尚未指定規則檔，請先上傳或選用預設規則檔', 'WARN');
      setActionButtonsDisabled(false);
      return;
    }
    if (action === 'fullFlow' && (!leftPayload || !rightPayload)) {
      log('完整流程需要左表與右表', 'WARN');
      setActionButtonsDisabled(false);
      return;
    }

    let payload;
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
    } else {
      throw new Error(`未知 action: ${action}`);
    }

    const leftName = state.left.file?.name || '-';
    const rightName = state.right.file?.name || '-';
    await startTask(action, payload, endpoint, baseUrl, leftName, rightName);
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
    els.btnStopTask?.addEventListener('click', stopCurrentTask);

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
      // log(`敹恍???{collapsed ? '撌脫?? : '撌脣???}`);
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







