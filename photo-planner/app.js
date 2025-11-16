/*
  合影拍照站位规划工具
  - 上传解析：Excel/CSV/TXT（姓名, 身高）
  - 排列：按身高降序，首排可预留领导位，整体中间高两侧低对称
  - 行列：手动或自动分配（每排人数相差不超过2）
  - 可视化：SVG网格，悬停显示身高，滚轮缩放
  - 交互：拖拽微调（交换），撤销/重做，重置
  - 导出：PNG、CSV、Excel，打印
*/

(() => {
  /** @typedef {{name: string, height: number}} Person */
  /** @typedef {{ type: 'person'|'leader'|'empty', label: string, name?: string, height?: number }} Seat */

  const el = {
    fileInput: document.getElementById('fileInput'),
    btnUseDemo: document.getElementById('btnUseDemo'),
    uploadMsg: document.getElementById('uploadMsg'),

    rowsInput: document.getElementById('rowsInput'),
    colsInput: document.getElementById('colsInput'),
    leaderCountInput: document.getElementById('leaderCountInput'),
    leaderNamesInput: document.getElementById('leaderNamesInput'),
    btnGenerate: document.getElementById('btnGenerate'),
    btnReset: document.getElementById('btnReset'),
    settingsMsg: document.getElementById('settingsMsg'),

    canvasWrapper: document.getElementById('canvasWrapper'),
    stage: document.getElementById('stage'),
    stageContent: document.getElementById('stageContent'),
    canvasMsg: document.getElementById('canvasMsg'),

    btnZoomIn: document.getElementById('btnZoomIn'),
    btnZoomOut: document.getElementById('btnZoomOut'),
    btnFit: document.getElementById('btnFit'),

    btnExportPNG: document.getElementById('btnExportPNG'),
    btnExportCSV: document.getElementById('btnExportCSV'),
    btnExportXLSX: document.getElementById('btnExportXLSX'),
    btnPrint: document.getElementById('btnPrint'),

    btnUndo: document.getElementById('btnUndo'),
    btnRedo: document.getElementById('btnRedo'),
  };

  const state = {
    people: /** @type {Person[]} */ ([]),
    initialSeats: /** @type {Seat[][]} */ ([]),
    seats: /** @type {Seat[][]} */ ([]),
    history: /** @type {Seat[][][]} */ ([]),
    future: /** @type {Seat[][][]} */ ([]),
    zoom: 1,
    panX: 0,
    panY: 0,
  };

  // ---------- Utils ----------
  const msg = (elOut, text, isError = false) => {
    elOut.textContent = text || '';
    elOut.style.color = isError ? '#f87171' : '#9ca3af';
  };

  const deepCloneSeats = (seats) => seats.map(row => row.map(s => ({...s})));

  const saveHistory = () => {
    state.history.push(deepCloneSeats(state.seats));
    if (state.history.length > 100) state.history.shift();
    state.future = [];
    updateUndoRedoButtons();
  };

  const restoreFrom = (stackFrom, stackTo) => {
    if (stackFrom.length === 0) return;
    stackTo.push(deepCloneSeats(state.seats));
    state.seats = deepCloneSeats(stackFrom.pop());
    render();
    updateUndoRedoButtons();
  };

  const updateUndoRedoButtons = () => {
    el.btnUndo.disabled = state.history.length === 0;
    el.btnRedo.disabled = state.future.length === 0;
  };

  const parseNumber = (s) => {
    const n = Number(String(s).toString().trim());
    return Number.isFinite(n) ? n : NaN;
  };

  const downloadBlob = (blob, filename) => {
    saveAs(blob, filename);
  };

  const autoDistributeRows = (total, rows) => {
    // Distribute so that counts differ by at most 2
    const base = Math.floor(total / rows);
    const remainder = total % rows; // first `remainder` rows get +1
    const arr = new Array(rows).fill(base).map((v, i) => v + (i < remainder ? 1 : 0));
    // ensure difference <= 2 by spreading further if needed (edge cases)
    let changed = true;
    while (changed) {
      changed = false;
      let maxI = 0, minI = 0;
      for (let i = 1; i < arr.length; i++) {
        if (arr[i] > arr[maxI]) maxI = i;
        if (arr[i] < arr[minI]) minI = i;
      }
      if (arr[maxI] - arr[minI] > 2) {
        arr[maxI]--; arr[minI]++;
        changed = true;
      }
    }
    return arr;
  };

  const parseColsInput = (text, rows, total) => {
    const trimmed = (text || '').trim();
    if (!trimmed) return autoDistributeRows(total, rows);
    if (/^\d+(\s*,\s*\d+)*$/.test(trimmed)) {
      const arr = trimmed.split(',').map(s => parseInt(s.trim(), 10));
      if (arr.length !== rows) throw new Error(`每排人数个数应为 ${rows} 个`);
      const sum = arr.reduce((a,b) => a+b, 0);
      if (sum !== total) throw new Error(`每排人数之和需等于总人数（${total}）`);
      const max = Math.max(...arr), min = Math.min(...arr);
      if (max - min > 2) throw new Error('每排人数相差不应超过 2');
      return arr;
    }
    if (/^\d+$/.test(trimmed)) {
      const per = parseInt(trimmed, 10);
      const fullRows = Math.floor(total / per);
      const remainder = total % per;
      const arr = new Array(rows).fill(0).map((_, i) => i < fullRows ? per : 0);
      if (remainder > 0) {
        for (let i = 0; i < rows; i++) {
          if (arr[i] === 0 && remainder - arr[i] > 0) {
            const give = Math.min(per, total - arr.reduce((a,b)=>a+b,0));
            arr[i] = give;
          }
        }
      }
      // cleanup: if still not sum==total, auto distribute
      const sum = arr.reduce((a,b)=>a+b,0);
      if (sum !== total) return autoDistributeRows(total, rows);
      const max = Math.max(...arr), min = Math.min(...arr);
      if (max - min > 2) return autoDistributeRows(total, rows);
      return arr;
    }
    throw new Error('每排人数输入格式有误');
  };

  // Middle-high symmetric ordering for a row
  const symmetricOrder = (people) => {
    const arr = [...people];
    arr.sort((a, b) => b.height - a.height);
    const result = [];
    let left = true;
    result.push(arr.shift()); // tallest center
    while (arr.length) {
      const p = arr.shift();
      if (left) result.unshift(p); else result.push(p);
      left = !left;
    }
    return result;
  };

  const buildSeats = (people, rows, perRowCounts, leaderCount, leaderNames) => {
    // First row (front) index 0.
    const seats = [];
    const sorted = [...people].sort((a,b)=>b.height - a.height);
    let idx = 0;

    for (let r = 0; r < rows; r++) {
      const count = perRowCounts[r];
      const rowSeats = [];

      // leader placeholders for row 0
      if (r === 0 && leaderCount > 0) {
        // Build positions with symmetric center allocation for leaders
        const head = Math.min(leaderCount, 50);
        const leaderPositions = new Array(head).fill(0).map((_,i)=>i);
        // We want leaders centered; represent as seat objects first
        // We'll mix leaders and persons later at render time but store explicitly now
        const leaderLabels = (leaderNames || []).slice(0, head);
        for (let i = 0; i < head; i++) {
          const label = leaderLabels[i] || '领导位';
          rowSeats.push({ type: 'leader', label });
        }
      }

      const remaining = count - rowSeats.length;
      const take = sorted.slice(idx, idx + remaining);
      idx += remaining;
      const ordered = symmetricOrder(take);
      // merge: leaders should be centered; interleave: leaders in middle, others symmetrically around
      if (rowSeats.length > 0) {
        // Place leaders centered
        const finalRow = [];
        const L = rowSeats.length;
        const total = L + ordered.length;
        const center = Math.floor(total / 2);
        // Build array with empty
        for (let i = 0; i < total; i++) finalRow[i] = null;
        // Place leaders: center outward
        const leaderOrdered = [];
        // create symmetric positions indexes for L leaders around center
        const leaderIndices = [];
        if (L > 0) {
          // Start with center pos
          if (total % 2 === 1) {
            leaderIndices.push(center);
            let step = 1;
            while (leaderIndices.length < L) {
              leaderIndices.push(center - step);
              if (leaderIndices.length < L) leaderIndices.push(center + step);
              step++;
            }
          } else {
            // even total: two centers
            const c1 = center - 1, c2 = center;
            leaderIndices.push(c1, c2);
            let step = 1;
            while (leaderIndices.length < L) {
              leaderIndices.push(c1 - step);
              if (leaderIndices.length < L) leaderIndices.push(c2 + step);
              step++;
            }
          }
        }
        for (let i = 0; i < L; i++) finalRow[leaderIndices[i]] = rowSeats[i];
        // Fill remaining with ordered persons from center outward gaps
        const gaps = [];
        for (let i = 0; i < total; i++) if (!finalRow[i]) gaps.push(i);
        // Arrange ordered people into gaps from middle outward
        const gapOrder = [];
        const mid = Math.floor(gaps.length/2);
        // build symmetric index sequence for gaps using middle-high pattern
        const gapsSorted = [...gaps].sort((a,b)=> Math.abs(a - center) - Math.abs(b - center));
        for (let i = 0; i < gapsSorted.length; i++) gapOrder.push(gapsSorted[i]);
        for (let i = 0; i < ordered.length; i++) finalRow[gapOrder[i]] = { type: 'person', label: ordered[i].name, name: ordered[i].name, height: ordered[i].height };
        seats.push(finalRow.map(s => s || { type: 'empty', label: '' }));
      } else {
        seats.push(ordered.map(p => ({ type: 'person', label: p.name, name: p.name, height: p.height })));
      }
    }

    return seats;
  };

  const getAllSeatsLinear = (seats) => seats.flat();

  const toCSV = (seats) => {
    const rows = [['row','col','type','name','height']];
    seats.forEach((row, rIdx) => {
      row.forEach((s, cIdx) => {
        rows.push([
          String(rIdx + 1),
          String(cIdx + 1),
          s.type,
          s.type === 'leader' ? s.label : (s.name || ''),
          s.type === 'person' && Number.isFinite(s.height) ? String(s.height) : ''
        ]);
      });
    });
    return rows.map(r => r.map(cell => /[",\n]/.test(cell) ? '"' + cell.replace(/"/g,'""') + '"' : cell).join(',')).join('\n');
  };

  const toXLSX = (seats) => {
    const data = [['行','列','类型','姓名','身高']];
    seats.forEach((row, rIdx) => {
      row.forEach((s, cIdx) => {
        data.push([
          rIdx + 1,
          cIdx + 1,
          s.type === 'leader' ? '领导位' : (s.type === 'person' ? '人员' : '空'),
          s.type === 'leader' ? s.label : (s.name || ''),
          s.type === 'person' && Number.isFinite(s.height) ? s.height : ''
        ]);
      });
    });
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '站位');
    const wbout = XLSX.write(wb, { bookType:'xlsx', type:'array' });
    return new Blob([wbout], { type: 'application/octet-stream' });
  };

  // ---------- Rendering ----------
  const CARD_W = 110, CARD_H = 60, GAP_X = 12, GAP_Y = 26, PAD = 40;

  const render = () => {
    const g = el.stageContent;
    while (g.firstChild) g.removeChild(g.firstChild);

    // compute stage size
    const rows = state.seats.length;
    const cols = Math.max(0, ...state.seats.map(r => r.length));
    const width = PAD*2 + cols * (CARD_W + GAP_X) - GAP_X;
    const height = PAD*2 + rows * (CARD_H + GAP_Y) - GAP_Y + 30; // row label room

    el.stage.setAttribute('viewBox', `${-state.panX} ${-state.panY} ${width/ state.zoom} ${height/ state.zoom}`);

    // row labels
    for (let r = 0; r < rows; r++) {
      const y = PAD + r * (CARD_H + GAP_Y) - 6;
      const text = svgEl('text', { x: 8, y, class: 'row-label' }, `第 ${r+1} 排`);
      g.appendChild(text);
    }

    for (let r = 0; r < rows; r++) {
      const row = state.seats[r];
      for (let c = 0; c < row.length; c++) {
        const x = PAD + c * (CARD_W + GAP_X);
        const y = PAD + r * (CARD_H + GAP_Y);
        const seat = row[c];
        const isLeader = seat.type === 'leader';
        const isEmpty = seat.type === 'empty';
        const group = svgEl('g', { 'data-r': String(r), 'data-c': String(c), cursor: isEmpty ? 'default' : 'grab' });

        const rect = svgEl('rect', {
          x, y, rx: 10, ry: 10, width: CARD_W, height: CARD_H,
          class: `card ${isLeader ? 'leader' : ''} ${isEmpty ? 'empty' : ''}`
        });
        group.appendChild(rect);

        const label = seat.type === 'leader' ? seat.label : (seat.name || '');
        const sub = seat.type === 'person' ? `${seat.height ?? ''}` : (isEmpty ? '' : '');
        const text1 = svgEl('text', { x: x + CARD_W/2, y: y + 26, class:'card-label', 'text-anchor':'middle' }, label);
        const text2 = svgEl('text', { x: x + CARD_W/2, y: y + 44, class:'card-sub', 'text-anchor':'middle' }, sub);
        group.appendChild(text1);
        group.appendChild(text2);

        // hover tooltip for height
        if (seat.type === 'person') {
          group.addEventListener('mousemove', (ev) => showTooltip(`${seat.name}：${seat.height} cm`, ev));
          group.addEventListener('mouseleave', hideTooltip);
        } else if (seat.type === 'leader') {
          group.addEventListener('mousemove', (ev) => showTooltip(`${seat.label}`, ev));
          group.addEventListener('mouseleave', hideTooltip);
        }

        // drag behavior
        if (!isEmpty) enableDrag(group, r, c);

        g.appendChild(group);
      }
    }
  };

  const svgEl = (name, attrs = {}, text = '') => {
    const e = document.createElementNS('http://www.w3.org/2000/svg', name);
    for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, String(v));
    if (text) e.textContent = text;
    return e;
  };

  // ---------- Tooltip ----------
  const tooltip = document.createElement('div');
  tooltip.className = 'tooltip';
  document.body.appendChild(tooltip);
  const showTooltip = (text, ev) => {
    tooltip.textContent = text;
    tooltip.style.display = 'block';
    const pad = 10;
    tooltip.style.left = ev.clientX + pad + 'px';
    tooltip.style.top = ev.clientY + pad + 'px';
  };
  const hideTooltip = () => { tooltip.style.display = 'none'; };

  // ---------- Drag & drop (swap) ----------
  let dragInfo = null; // { r: number, c: number, pointerId: number, group: SVGGElement }

  function enableDrag(group, r, c) {
    group.addEventListener('pointerdown', (ev) => {
      if (ev.button !== 0) return;
      ev.preventDefault();
      dragInfo = { r, c, pointerId: ev.pointerId, group };
      try { group.setPointerCapture(ev.pointerId); } catch {}
      group.style.cursor = 'grabbing';
    });
  }

  el.stage.addEventListener('pointerup', (ev) => {
    if (!dragInfo) return;
    const { r, c, pointerId, group } = dragInfo;
    try { group.releasePointerCapture(pointerId); } catch {}
    group.style.cursor = 'grab';
    const target = getSeatFromEvent(ev);
    if (target && (target.r !== r || target.c !== c)) {
      saveHistory();
      swapSeats(r, c, target.r, target.c);
    }
    dragInfo = null;
  });

  el.stage.addEventListener('pointercancel', () => {
    if (!dragInfo) return;
    const { pointerId, group } = dragInfo;
    try { group.releasePointerCapture(pointerId); } catch {}
    group.style.cursor = 'grab';
    dragInfo = null;
  });

  function getSeatFromEvent(ev) {
    const pt = el.stage.createSVGPoint();
    pt.x = ev.clientX; pt.y = ev.clientY;
    const ctm = el.stage.getScreenCTM();
    if (!ctm) return null;
    const p = pt.matrixTransform(ctm.inverse());

    const r = Math.round((p.y - PAD - CARD_H/2) / (CARD_H + GAP_Y));
    const c = Math.round((p.x - PAD - CARD_W/2) / (CARD_W + GAP_X));
    if (r < 0 || r >= state.seats.length) return null;
    if (c < 0 || c >= state.seats[r].length) return null;
    return { r, c };
  }

  function swapSeats(r1, c1, r2, c2) {
    const a = state.seats[r1][c1];
    const b = state.seats[r2][c2];
    state.seats[r1][c1] = b;
    state.seats[r2][c2] = a;
    render();
  }

  // ---------- Zoom & Pan ----------
  el.canvasWrapper.addEventListener('wheel', (ev) => {
    ev.preventDefault();
    const delta = -Math.sign(ev.deltaY) * 0.1;
    const newZoom = Math.min(4, Math.max(0.3, state.zoom + delta));

    // zoom to mouse point
    const rect = el.stage.getBoundingClientRect();
    const mx = ev.clientX - rect.left;
    const my = ev.clientY - rect.top;
    const wx = (mx / rect.width) * (rect.width / state.zoom);
    const wy = (my / rect.height) * (rect.height / state.zoom);

    state.panX += wx * (1/state.zoom - 1/newZoom);
    state.panY += wy * (1/state.zoom - 1/newZoom);
    state.zoom = newZoom;
    render();
  }, { passive: false });

  el.btnZoomIn.addEventListener('click', () => { state.zoom = Math.min(4, state.zoom + 0.1); render(); });
  el.btnZoomOut.addEventListener('click', () => { state.zoom = Math.max(0.3, state.zoom - 0.1); render(); });
  el.btnFit.addEventListener('click', () => { state.zoom = 1; state.panX = 0; state.panY = 0; render(); });

  // ---------- Upload parsing ----------
  el.fileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const people = await parseFile(file);
      if (!people || people.length === 0) throw new Error('未读取到有效人员数据');
      state.people = people;
      msg(el.uploadMsg, `已加载 ${people.length} 人`);
    } catch (err) {
      console.error(err);
      msg(el.uploadMsg, '文件解析失败：请确保文件包含姓名和身高列', true);
    }
  });

  el.btnUseDemo.addEventListener('click', () => {
    const demo = [
      { name: '张一', height: 186 }, { name: '李二', height: 183 }, { name: '王三', height: 182 },
      { name: '赵四', height: 180 }, { name: '陈五', height: 179 }, { name: '吴六', height: 178 },
      { name: '郑七', height: 177 }, { name: '王八', height: 175 }, { name: '钱九', height: 174 },
      { name: '孙十', height: 173 }, { name: '周十一', height: 172 }, { name: '吴十二', height: 171 },
      { name: '郑十三', height: 170 }, { name: '王十四', height: 169 }, { name: '钱十五', height: 168 },
      { name: '孙十六', height: 167 }, { name: '周十七', height: 166 }, { name: '吴十八', height: 165 },
      { name: '郑十九', height: 164 }, { name: '王二十', height: 163 }
    ];
    state.people = demo;
    msg(el.uploadMsg, `已加载示例 ${demo.length} 人`);
  });

  async function parseFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'csv' || ext === 'txt') {
      const text = await file.text();
      const result = Papa.parse(text, { header: true, skipEmptyLines: true });
      const rows = result.data;
      return normalizePeople(rows);
    }
    if (ext === 'xls' || ext === 'xlsx') {
      const data = new Uint8Array(await file.arrayBuffer());
      const wb = XLSX.read(data, { type: 'array' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      return normalizePeople(rows);
    }
    throw new Error('不支持的文件类型');
  }

  function normalizePeople(rows) {
    // find name and height keys by fuzzy matching
    const candidates = rows && rows.length ? Object.keys(rows[0]) : [];
    const findKey = (names) => candidates.find(k => names.some(n => String(k).toLowerCase().includes(n)));
    const nameKey = findKey(['name','姓名']);
    const heightKey = findKey(['height','身高']);
    if (!nameKey || !heightKey) throw new Error('缺少姓名或身高列');

    const people = [];
    for (const r of rows) {
      const name = String(r[nameKey]).trim();
      const h = parseNumber(r[heightKey]);
      if (!name) continue;
      if (!Number.isFinite(h)) continue;
      people.push({ name, height: h });
    }
    return people;
  }

  // ---------- Generate seating ----------
  el.btnGenerate.addEventListener('click', () => {
    if (state.people.length === 0) {
      msg(el.settingsMsg, '请先上传人员文件或使用示例数据', true); return;
    }
    const rows = Math.max(1, parseInt(el.rowsInput.value, 10) || 1);
    const leaderCount = Math.min(50, Math.max(0, parseInt(el.leaderCountInput.value, 10) || 0));
    const leaderNames = (el.leaderNamesInput.value || '').split(',').map(s=>s.trim()).filter(Boolean).slice(0, 50);

    // first row must have capacity >= leaderCount
    const total = state.people.length + leaderCount; // reserve slots

    let perRow;
    try {
      perRow = parseColsInput(el.colsInput.value, rows, total);
    } catch (e) {
      msg(el.settingsMsg, String(e.message || e), true); return;
    }
    if (perRow[0] < leaderCount) {
      msg(el.settingsMsg, '第1排人数需不少于领导席位数量', true); return;
    }

    const seats = buildSeats(state.people, rows, perRow, leaderCount, leaderNames);
    state.seats = seats;
    state.initialSeats = deepCloneSeats(seats);
    state.history = []; state.future = [];
    updateUndoRedoButtons();
    state.zoom = 1; state.panX = 0; state.panY = 0;
    msg(el.settingsMsg, '已生成站位图');
    render();
  });

  // ---------- Reset ----------
  el.btnReset.addEventListener('click', () => {
    if (state.initialSeats.length) {
      state.seats = deepCloneSeats(state.initialSeats);
      state.history = []; state.future = [];
      updateUndoRedoButtons();
      render();
      msg(el.canvasMsg, '已恢复初始排序');
    }
  });

  // ---------- Undo/Redo ----------
  el.btnUndo.addEventListener('click', () => restoreFrom(state.history, state.future));
  el.btnRedo.addEventListener('click', () => restoreFrom(state.future, state.history));

  // ---------- Export ----------
  el.btnExportCSV.addEventListener('click', () => {
    if (!state.seats.length) return;
    const csv = toCSV(state.seats);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    downloadBlob(blob, '站位.csv');
  });

  el.btnExportXLSX.addEventListener('click', () => {
    if (!state.seats.length) return;
    const blob = toXLSX(state.seats);
    downloadBlob(blob, '站位.xlsx');
  });

  el.btnExportPNG.addEventListener('click', async () => {
    if (!state.seats.length) return;
    // Use html-to-image on a wrapper containing the SVG
    const node = el.canvasWrapper;
    try {
      const dataUrl = await htmlToImage.toPng(node, { pixelRatio: 2, backgroundColor: '#0b1220' });
      const res = await fetch(dataUrl);
      const blob = await res.blob();
      downloadBlob(blob, '站位.png');
    } catch (e) {
      console.error(e);
      msg(el.canvasMsg, '导出PNG失败', true);
    }
  });

  el.btnPrint.addEventListener('click', () => {
    window.print();
  });

})();
