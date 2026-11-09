/* Excel A→B 转换 · 操作台
 * 浏览器端可直接完成转换（SheetJS）；若由 serve.py 本地服务打开，
 * 会自动切换到「本地引擎模式」，由 Python 引擎执行以完整保留模板格式。
 */
(function () {
  'use strict';

  const $ = (id) => document.getElementById(id);

  const state = {
    mode: 'browser',           // browser | local
    A: null,                   // {name, buf, wb, sheet, headerRow, headers, matrix}
    B: null,
    rules: [],                 // {target, type, source, transforms, def}
    started: false,
    filterTouched: false,      // 用户是否手动改过「行过滤」（改过则不再自动套用默认）
    filterRules: [],           // 多条件：[{column, match, mode}]
    batchA: []                 // 批量模式：除主文件外的其余 A 文件（File 对象）
  };

  /* ---- 变换预设：常用公式（下拉选项）----
   * value 是「变换」输入框的可编辑文本，必须能被 parseTransforms 正确解析，
   * 且与 Python 引擎 TRANSFORM_PRESETS 中的 @别名 链一一对应。            */
  const TRANSFORM_PRESETS = [
    { label: '去空格', value: 'trim' },
    { label: '转数字', value: 'trim,number' },
    { label: '金额两位', value: 'number,round:2' },
    { label: '取整', value: 'number,int' },
    { label: '转大写', value: 'trim,upper' },
    { label: '转小写', value: 'trim,lower' },
    { label: '转文本', value: 'text' },
    { label: '标准日期', value: 'date:"%Y-%m-%d"' },
    { label: '日期+30天', value: 'date_offset:30' },
    { label: '首行', value: 'first_line' },
    { label: '补零4位', value: 'pad:["0",4]' },
    { label: '去SH前缀', value: 'regex_replace:["^SH",""]' },
    { label: '是/否', value: 'bool' },
    { label: '加前缀P-', value: 'prefix:"P-"' },
    { label: '清空变换', value: '' }
  ];
  /* ---- /变换预设 ---- */

  /* ------------------------------------------------------------------ */
  /* 工具                                                                */
  /* ------------------------------------------------------------------ */
  const isBlank = (v) => v === null || v === undefined || (typeof v === 'string' && v.trim() === '');

  function toText(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return fmtDate(v, '%Y-%m-%d');
    if (typeof v === 'number' && Number.isInteger(v)) return String(v);
    return String(v);
  }

  function fmtDate(d, fmt) {
    const p = (n, w) => String(n).padStart(w || 2, '0');
    return fmt
      .replace(/%Y/g, d.getFullYear())
      .replace(/%m/g, p(d.getMonth() + 1))
      .replace(/%d/g, p(d.getDate()))
      .replace(/%H/g, p(d.getHours()))
      .replace(/%M/g, p(d.getMinutes()))
      .replace(/%S/g, p(d.getSeconds()));
  }

  function parseDate(v) {
    if (v instanceof Date) return v;
    if (typeof v === 'number') return new Date(Date.UTC(1899, 11, 30) + v * 86400000);
    const s = toText(v).trim();
    let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
    m = s.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日$/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }

  /* 变换链：与 Python 引擎保持一致 */
  function applyTransform(value, spec) {
    let name, arg;
    if (typeof spec === 'string') { name = spec; arg = null; }
    else { name = Object.keys(spec)[0]; arg = spec[name]; }
    name = String(name).trim().toLowerCase();

    switch (name) {
      case 'trim': case 'strip': return toText(value).trim();
      case 'upper': return toText(value).toUpperCase();
      case 'lower': return toText(value).toLowerCase();
      case 'text': return toText(value);
      case 'int': { const n = parseInt(toText(value).replace(/,/g, ''), 10); return isNaN(n) ? value : n; }
      case 'number': {
        const f = parseFloat(toText(value).replace(/,/g, ''));
        if (isNaN(f)) return value;
        return Number.isInteger(f) ? f : f;
      }
      case 'round': { const f = parseFloat(toText(value)); return isNaN(f) ? value : Math.round(f * Math.pow(10, +(arg || 0))) / Math.pow(10, +(arg || 0)); }
      case 'bool': return ['', '0', 'false', 'False', '否'].indexOf(toText(value).trim()) >= 0 ? 0 : 1;
      case 'date': { const d = parseDate(value); return d ? fmtDate(d, arg || '%Y-%m-%d') : toText(value); }
      case 'date_offset': {
        const d = parseDate(value); if (!d) return value;
        const nd = new Date(d.getTime() + (+(arg || 0)) * 86400000);
        return fmtDate(nd, '%Y-%m-%d');
      }
      case 'prefix': return String(arg) + toText(value);
      case 'suffix': return toText(value) + String(arg);
      case 'replace': case 'regex_replace': {
        try { return toText(value).replace(new RegExp(arg[0], 'g'), arg[1]); }
        catch (e) { return toText(value); }
      }
      case 'split': {
        const parts = toText(value).split(arg[0]);
        const i = +arg[1];
        const p = i < 0 ? parts[parts.length + i] : parts[i];
        return p === undefined ? '' : p.trim();
      }
      case 'first_line': {
        const s = toText(value);
        return s ? s.split(/\r\n|\r|\n/)[0].trim() : '';
      }
      case 'pad': return toText(value).padStart(+arg[1], arg[0]);
      default: return value;
    }
  }

  /* 解析变换输入框文本 -> transforms 数组 */
  function parseTransforms(text) {
    text = (text || '').trim();
    if (!text) return [];
    // 顶层按逗号切分（忽略 [] 与 "" 内的逗号）
    const tokens = [];
    let cur = '', depth = 0, inStr = false, quote = '';
    for (const ch of text) {
      if (inStr) { cur += ch; if (ch === quote) inStr = false; continue; }
      if (ch === '"' || ch === "'") { inStr = true; quote = ch; cur += ch; continue; }
      if (ch === '[') depth++;
      if (ch === ']') depth--;
      if (ch === ',' && depth === 0) { tokens.push(cur.trim()); cur = ''; continue; }
      cur += ch;
    }
    if (cur.trim()) tokens.push(cur.trim());

    return tokens.filter(Boolean).map((tk) => {
      const idx = tk.indexOf(':');
      if (idx < 0) return tk.trim();
      const name = tk.slice(0, idx).trim();
      const rawArg = tk.slice(idx + 1).trim();
      let arg;
      try { arg = JSON.parse(rawArg); }
      catch (e) { arg = rawArg.replace(/^["']|["']$/g, ''); }
      const o = {}; o[name] = arg; return o;
    });
  }

  function transformsToText(ts) {
    // 配置里 transforms 可能写成字符串（单个变换名，或引擎的 "@预设别名"）。
    // 内置别名在此还原成等价变换链文本，保证浏览器模式导入后也能正确执行。
    if (typeof ts === 'string') {
      const hit = TRANSFORM_PRESETS.find((p) => '@' + p.label === ts.trim());
      return hit ? hit.value : ts;
    }
    return (ts || []).map((t) => {
      if (typeof t === 'string') return t;
      const k = Object.keys(t)[0];
      const v = t[k];
      if (v === null || v === undefined) return k;
      return k + ':' + (typeof v === 'string' ? v : JSON.stringify(v));
    }).join(',');
  }

  /* 表头行识别：前 8 行中非空最多者，多个并列取最后一个 */
  function detectHeaderRow(matrix) {
    const n = Math.min(8, matrix.length);
    let best = 1, bestCount = -1;
    for (let r = 0; r < n; r++) {
      const c = (matrix[r] || []).filter((v) => !isBlank(v)).length;
      if (c >= bestCount) { bestCount = c; best = r + 1; }
    }
    return best;
  }

  function buildHeaders(matrix, headerRow) {
    const raw = matrix[headerRow - 1] || [];
    const seen = {}, out = [];
    raw.forEach((v, i) => {
      let name = toText(v).trim();
      if (!name) name = '__COL' + (i + 1) + '__';
      if (seen[name] !== undefined) { seen[name] += 1; name = name + '#' + seen[name]; }
      else seen[name] = 0;
      out.push(name);
    });
    return out;
  }

  /* ------------------------------------------------------------------ */
  /* 日志 / 状态                                                          */
  /* ------------------------------------------------------------------ */
  function log(msg, cls) {
    const el = $('log');
    el.hidden = false;
    const span = document.createElement('span');
    span.className = cls || '';
    span.textContent = (cls === 'err' ? '✖ ' : cls === 'ok' ? '✔ ' : cls === 'warn' ? '⚠ ' : '· ') + msg + '\n';
    el.appendChild(span);
    el.scrollTop = el.scrollHeight;
  }
  function setStatus(t) { $('status').textContent = t; }

  /* 通用：给输入控件打「非法」标红 */
  function markInvalid(el, bad, msg) {
    el.classList.toggle('invalid', !!bad);
    el.title = bad ? (msg || '输入无效') : '';
  }

  /* ------------------------------------------------------------------ */
  /* 配置记忆：按 A 表表头指纹记住上次配置，加载同类文件时自动套用          */
  /* ------------------------------------------------------------------ */
  const MEM_KEY = 'excel-a2b:configs';
  const fpOf = (st) => st.headers.join('|');
  function memAll() { try { return JSON.parse(localStorage.getItem(MEM_KEY)) || {}; } catch (e) { return {}; } }
  function memWrite(all) {
    const ks = Object.keys(all);
    if (ks.length > 20) {  // LRU：最多保留 20 套
      ks.sort((a, b) => (all[a].savedAt || 0) - (all[b].savedAt || 0));
      ks.slice(0, ks.length - 20).forEach((k) => delete all[k]);
    }
    try { localStorage.setItem(MEM_KEY, JSON.stringify(all)); } catch (e) { /* 存储满等情况忽略 */ }
  }
  function memSave() {
    if (!state.A) return;
    const all = memAll();
    all[fpOf(state.A)] = {
      savedAt: Date.now(),
      aName: state.A.name,
      rules: state.rules,
      filterChecked: $('optFilter').checked,
      combine: $('fltCombine').value,
      filterRules: state.filterRules,
      split: readSplit(),
      agg: readAgg(),
      options: {
        skipEmpty: $('optSkipEmpty').checked,
        clear: $('optClear').checked,
        createCol: $('optCreateCol').checked,
        onMissing: $('optOnMissing').value
      },
      outName: $('outName').value
    };
    memWrite(all);
  }
  /* A 表加载后调用：若存在同结构历史配置且当前映射为空，自动套用 */
  function memApply() {
    if (!state.A) return false;
    const blank = state.rules.every((r) => !r.source && !r.def && r.type === 'column');
    if (!blank) return false;
    const rec = memAll()[fpOf(state.A)];
    if (!rec) return false;
    state.rules = (rec.rules || []).map((r) => Object.assign({}, r));
    state.filterRules = (rec.filterRules || []).map((r) => Object.assign({}, r));
    state.filterTouched = true;
    $('optFilter').checked = !!rec.filterChecked;
    $('fltCombine').value = rec.combine || 'and';
    if (rec.options) {
      $('optSkipEmpty').checked = rec.options.skipEmpty !== false;
      $('optClear').checked = rec.options.clear !== false;
      $('optCreateCol').checked = !!rec.options.createCol;
      if (rec.options.onMissing) $('optOnMissing').value = rec.options.onMissing;
    }
    if (rec.outName) $('outName').value = rec.outName;
    applySplitUI(rec.split);
    applyAggUI(rec.agg);
    renderFilterRules();
    const when = rec.savedAt ? new Date(rec.savedAt).toLocaleString() : '';
    log('识别到该表结构的历史配置（来自 ' + (rec.aName || '?') + ' · ' + when + '），已自动套用；点「清空映射」可重新开始', 'ok');
    return true;
  }

  /* ------------------------------------------------------------------ */
  /* 主题                                                                */
  /* ------------------------------------------------------------------ */
  $('btnTheme').addEventListener('click', () => {
    const cur = document.documentElement.getAttribute('data-theme');
    document.documentElement.setAttribute('data-theme', cur === 'dark' ? 'light' : 'dark');
  });

  /* ------------------------------------------------------------------ */
  /* 文件加载                                                            */
  /* ------------------------------------------------------------------ */
  function wireDrop(dropId, inputId, clearId, key) {
    const drop = $(dropId), input = $(inputId);
    drop.addEventListener('click', (e) => { if (e.target.id !== clearId) input.click(); });
    drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('over'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('over'));
    drop.addEventListener('drop', (e) => {
      e.preventDefault(); drop.classList.remove('over');
      if (!e.dataTransfer.files.length) return;
      if (key === 'A') loadFilesA(e.dataTransfer.files);
      else loadFile(key, e.dataTransfer.files[0]);
    });
    input.addEventListener('change', () => {
      if (!input.files.length) return;
      if (key === 'A') loadFilesA(input.files);
      else loadFile(key, input.files[0]);
      input.value = '';
    });
    $(clearId).addEventListener('click', (e) => { e.stopPropagation(); clearFile(key); });
  }

  /* A 表支持多文件：首个为主文件，其余进入批量队列（套用同一配置逐个转换） */
  function loadFilesA(files) {
    files = Array.prototype.slice.call(files);
    if (!files.length) return;
    state.batchA = files.slice(1);
    renderBatchBar();
    if (state.batchA.length) log('批量模式：共 ' + files.length + ' 个 A 文件将套用同一配置', 'ok');
    loadFile('A', files[0]);
  }

  function renderBatchBar() {
    const bar = $('batchBar');
    if (!bar) return;
    const n = state.batchA.length;
    bar.hidden = n === 0;
    if (n) {
      $('batchCount').textContent = String(n + 1);
      const names = state.batchA.slice(0, 4).map((f) => f.name).join('、');
      $('batchNames').textContent = names + (n > 4 ? ' 等' : '');
    }
  }

  const stem = (name) => String(name || '').replace(/\.[^.]+$/, '');

  /* 读取批量文件为与 state.A 同构的对象（沿用主文件的工作表名与表头行） */
  function readFileState(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const buf = e.target.result;
          const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true, cellStyles: true });
          let sheet = state.A.sheet;
          if (wb.SheetNames.indexOf(sheet) < 0) sheet = wb.SheetNames[0];
          const matrix = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, raw: true, defval: null, blankrows: true });
          const headerRow = state.A.headerRow;
          resolve({ name: file.name, buf: buf, wb: wb, sheet: sheet, headerRow: headerRow, headers: buildHeaders(matrix, headerRow), matrix: matrix });
        } catch (err) { reject(err); }
      };
      reader.onerror = () => reject(new Error('文件读取失败'));
      reader.readAsArrayBuffer(file);
    });
  }

  function loadFile(key, file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const buf = e.target.result;
        const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true, cellStyles: true });
        const sheet = wb.SheetNames[0];
        const matrix = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, raw: true, defval: null, blankrows: true });
        const headerRow = detectHeaderRow(matrix);
        state[key] = { name: file.name, buf: buf, wb: wb, sheet: sheet, headerRow: headerRow, headers: buildHeaders(matrix, headerRow), matrix: matrix };

        // 健壮性提示：大文件 / 合并单元格
        if (matrix.length > 20000) log('提示：' + key + ' 表行数较多（' + matrix.length + ' 行），浏览器模式可能较慢，建议改用本地引擎模式', 'warn');
        const merges = wb.Sheets[sheet]['!merges'];
        if (merges && merges.length) log('提示：' + key + ' 表含 ' + merges.length + ' 处合并单元格，取值以左上角为准、其余视为空', 'warn');

        // 加载 B 后：按 B 的列重建映射行；若之前已有同名目标列，保留其配置
        if (key === 'B') {
          const prev = {};
          state.rules.forEach((r) => { if (r.target) prev[r.target] = r; });
          state.rules = state[key].headers.map((h) => prev[h] || newRule(h));
        }

        log('已加载 ' + key + ' 表：' + file.name + '（' + wb.SheetNames.length + ' 个工作表）', 'ok');
        renderFile(key);
        if (key === 'A') memApply();   // 同结构历史配置自动套用
        afterLoad();
      } catch (err) {
        log('读取失败：' + err.message, 'err');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function newRule(target) {
    return { target: target || '', type: 'column', source: '', transforms: '', def: '' };
  }

  function clearFile(key) {
    state[key] = null;
    const info = $(key === 'A' ? 'infoA' : 'infoB');
    const drop = $(key === 'A' ? 'dropA' : 'dropB');
    info.className = 'empty';
    info.innerHTML = key === 'A'
      ? '拖入或点击选择<small>支持 .xls / .xlsx / .xlsm / .csv</small>'
      : '拖入或点击选择<small>给了模板 → 保留其表头与格式</small>';
    drop.classList.remove('loaded');
    $(key === 'A' ? 'sheetA' : 'sheetB').innerHTML = '<option>—</option>';
    $(key === 'A' ? 'sheetA' : 'sheetB').disabled = true;
    $(key === 'A' ? 'headA' : 'headB').disabled = true;
    if (key === 'A') {
      state.filterTouched = false;
      state.filterRules = [];
      state.batchA = [];
      $('optFilter').checked = false;
      populateFilterColumns();
      renderBatchBar();
    }
    afterLoad();
  }

  function renderFile(key) {
    const st = state[key];
    const info = $(key === 'A' ? 'infoA' : 'infoB');
    const drop = $(key === 'A' ? 'dropA' : 'dropB');
    if (!st) return;
    drop.classList.add('loaded');
    info.className = '';
    const dataRows = st.matrix.slice(st.headerRow).filter((r) => (r || []).some((v) => !isBlank(v))).length;
    info.innerHTML = '<div class="fname">' + esc(st.name) + '</div>' +
      '<div class="fmeta">' + st.matrix.length + ' 行（有效数据 ' + dataRows + ' 行）· ' + st.headers.length + ' 列 · 表头第 ' + st.headerRow + ' 行</div>';

    const sheetSel = $(key === 'A' ? 'sheetA' : 'sheetB');
    sheetSel.innerHTML = '';
    st.wb.SheetNames.forEach((n) => {
      const o = document.createElement('option'); o.value = n; o.textContent = n;
      if (n === st.sheet) o.selected = true;
      sheetSel.appendChild(o);
    });
    sheetSel.disabled = false;
    const headInput = $(key === 'A' ? 'headA' : 'headB');
    headInput.value = st.headerRow;
    headInput.disabled = false;

    populateFilterColumns();
  }

  /* 重建列下拉选项，保留当前选中值；值不在列表时自动补一个选项 */
  function rebuildColOptions(sel, headers) {
    const cur = sel.value;
    sel.innerHTML = '<option value="">— 选列 —</option>';
    headers.forEach((h) => {
      const o = document.createElement('option'); o.value = h; o.textContent = h;
      if (h === cur) o.selected = true;
      sel.appendChild(o);
    });
  }
  function ensureOption(sel, val) {
    if (!val) return;
    const exists = Array.prototype.some.call(sel.options, (o) => o.value === val);
    if (!exists) { const o = document.createElement('option'); o.value = val; o.textContent = val; sel.appendChild(o); }
  }

  /* 填充「行过滤」：A 表含货号列且用户未手动改过时，自动加一条货号默认正则条件 */
  function populateFilterColumns() {
    $('optFilter').disabled = !state.A;
    const headers = state.A ? state.A.headers : [];
    rebuildColOptions($('splitCol'), headers);
    rebuildColOptions($('aggCol'), headers);
    if (state.A && !state.filterTouched && headers.indexOf('货号') >= 0 && state.filterRules.length === 0) {
      $('optFilter').checked = true;
      state.filterRules.push({ column: '货号', match: '^SH\\d+C-\\d+$', mode: 'include' });
    }
    renderFilterRules();
  }

  /* 渲染多条过滤条件 */
  function renderFilterRules() {
    const box = $('filterRules');
    box.innerHTML = '';
    const headers = state.A ? state.A.headers : [];
    state.filterRules.forEach((rule, i) => {
      const row = document.createElement('div');
      row.className = 'filter-rule';

      const sel = document.createElement('select');
      sel.className = 'fr-col';
      sel.innerHTML = '<option value="">— 选列 —</option>';
      headers.forEach((h) => {
        const o = document.createElement('option'); o.value = h; o.textContent = h;
        if (h === rule.column) o.selected = true;
        sel.appendChild(o);
      });
      sel.addEventListener('change', () => {
        rule.column = sel.value; state.filterTouched = true;
        markInvalid(sel, !!sel.value && headers.indexOf(sel.value) < 0, 'A 表中未找到该列');
      });

      const inp = document.createElement('input');
      inp.type = 'text'; inp.className = 'fr-match'; inp.placeholder = '^SH\\d+C-\\d+$';
      inp.value = rule.match || '';
      const checkRe = () => {
        let bad = false;
        if (inp.value.trim()) { try { new RegExp(inp.value); } catch (e) { bad = true; } }
        markInvalid(inp, bad, '正则表达式非法');
      };
      inp.addEventListener('input', () => { rule.match = inp.value; state.filterTouched = true; checkRe(); });
      checkRe();

      const mode = document.createElement('select');
      mode.className = 'fr-mode';
      [['include', '保留匹配的行'], ['exclude', '跳过匹配的行']].forEach(([v, t]) => {
        const o = document.createElement('option'); o.value = v; o.textContent = t;
        if (rule.mode === v) o.selected = true;
        mode.appendChild(o);
      });
      mode.addEventListener('change', () => { rule.mode = mode.value; state.filterTouched = true; });

      const del = document.createElement('button');
      del.className = 'btn-del'; del.textContent = '×'; del.title = '删除条件';
      del.addEventListener('click', () => { state.filterRules.splice(i, 1); renderFilterRules(); });

      row.append(sel, inp, mode, del);
      box.appendChild(row);
    });
    toggleFilterFields();
  }

  /* 某行是否通过多条件过滤（rf = {combine, rules}） */
  function rowPassesFilter(rf, dict) {
    const rules = rf.rules || [];
    if (!rules.length) return true;
    const res = rules.map((r) => {
      const val = toText(dict[r.column] || '');
      let isMatch = true;
      try { isMatch = new RegExp(r.match).test(val); } catch (e) { isMatch = true; }
      return r.mode === 'exclude' ? !isMatch : isMatch;
    });
    return rf.combine === 'or' ? res.some(Boolean) : res.every(Boolean);
  }

  /* 勾选「按条件过滤行」后，启用子控件；否则置灰 */
  function toggleFilterFields() {
    const on = $('optFilter').checked && !$('optFilter').disabled;
    document.getElementById('filterBody').classList.toggle('off', !on);
  }

  function esc(s) { return String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

  /* 切换工作表 / 表头行后刷新 */
  function reparse(key) {
    const st = state[key];
    if (!st) return;
    st.sheet = $(key === 'A' ? 'sheetA' : 'sheetB').value;
    st.matrix = XLSX.utils.sheet_to_json(st.wb.Sheets[st.sheet], { header: 1, raw: true, defval: null, blankrows: true });
    st.headerRow = Math.max(1, parseInt($(key === 'A' ? 'headA' : 'headB').value, 10) || 1);
    st.headers = buildHeaders(st.matrix, st.headerRow);
    renderFile(key);
    afterLoad();
  }

  function afterLoad() {
    const hasA = !!state.A, hasB = !!state.B;
    $('btnDetect').disabled = !(hasA || hasB);
    $('btnReload').disabled = !(hasA || hasB);
    $('btnAuto').disabled = !hasA;
    $('btnAdd').disabled = !(hasA || hasB);
    $('btnClearMap').disabled = state.rules.length === 0;
    $('btnExport').disabled = state.rules.length === 0;

    // 全新打开、什么都没加载时，给一行空规则占位
    if (!hasA && !hasB && state.rules.length === 0) state.rules = [newRule()];

    // 双表就绪且所有来源都还空着时，自动按同名匹配一次
    if (hasA && hasB && state.rules.length && state.rules.every((r) => !r.source && r.type === 'column' && r.target)) {
      autoMatch(true);
    }

    renderMapping();
    updateHint();
    updateReady();
    toggleFilterFields();
  }

  function updateHint() {
    const parts = [];
    if (state.A) parts.push('A：表头第 ' + state.A.headerRow + ' 行，' + state.A.headers.length + ' 列');
    if (state.B) parts.push('B：表头第 ' + state.B.headerRow + ' 行，' + state.B.headers.length + ' 列');
    $('structHint').textContent = parts.join('　|　');
  }

  function dataRowCount(st) {
    return st.matrix.slice(st.headerRow).filter((r) => (r || []).some((v) => !isBlank(v))).length;
  }

  function updateReady() {
    const mapped = state.rules.filter((r) => r.target && (r.type === 'const' || r.type === 'index' || String(r.source).trim()));
    const ok = !!state.A && mapped.length > 0;
    $('btnPreview').disabled = !ok;
    $('btnRun').disabled = !ok;
    $('mapCount').textContent = state.rules.length ? ('已配置 ' + state.rules.length + ' 列，其中 ' + mapped.length + ' 列有来源') : '';
    if (ok) {
      let t = '就绪：将写入 ' + mapped.length + ' 列，源表 ' + dataRowCount(state.A) + ' 行数据';
      if (state.batchA.length) t += '（批量 ' + (state.batchA.length + 1) + ' 个文件）';
      setStatus(t);
    }
  }

  /* ------------------------------------------------------------------ */
  /* 映射表渲染                                                          */
  /* ------------------------------------------------------------------ */
  function renderMapping() {
    const body = $('mapBody');
    body.innerHTML = '';
    $('mapEmpty').hidden = state.rules.length > 0;
    $('mapTable').hidden = state.rules.length === 0;

    state.rules.forEach((rule, i) => {
      const tr = document.createElement('tr');

      // 目标列
      const tdT = document.createElement('td');
      const inT = document.createElement('input');
      inT.type = 'text'; inT.value = rule.target; inT.placeholder = 'B 的列名';
      inT.setAttribute('list', 'dlTarget');
      inT.addEventListener('input', () => { rule.target = inT.value; updateReady(); });
      tdT.appendChild(inT);

      // 类型
      const tdTy = document.createElement('td');
      const sel = document.createElement('select');
      [['column', '列'], ['const', '常量'], ['template', '模板'], ['expr', '表达式'], ['index', '行号']].forEach(([v, t]) => {
        const o = document.createElement('option'); o.value = v; o.textContent = t;
        if (rule.type === v) o.selected = true;
        sel.appendChild(o);
      });
      sel.addEventListener('change', () => { rule.type = sel.value; renderMapping(); updateReady(); });
      tdTy.appendChild(sel);

      // 来源 / 值
      const tdS = document.createElement('td');
      const inS = document.createElement('input');
      inS.type = 'text';
      if (rule.type === 'column') {
        inS.value = rule.source; inS.placeholder = 'A 的列名';
        inS.setAttribute('list', 'dlSource');
      } else if (rule.type === 'const') {
        inS.value = rule.def === '__const__' ? rule.source : rule.source; inS.placeholder = '固定值';
      } else if (rule.type === 'template') {
        inS.value = rule.source; inS.placeholder = '{款式}-{货号}';
        inS.setAttribute('list', 'dlSource');
      } else if (rule.type === 'expr') {
        inS.value = rule.source; inS.placeholder = '如 num(数量)*num(单价)';
      } else {
        inS.value = ''; inS.placeholder = '从 1 开始自增'; inS.disabled = true; inS.setAttribute('list', 'dlSource');
      }
      const checkSrc = () => {
        // 列来源在 A 表中找不到时标红提示（模板/表达式不校验）
        if (!state.A || rule.type !== 'column' || !rule.source) return markInvalid(inS, false);
        markInvalid(inS, colIndexOf(state.A, rule.source) < 0, 'A 表中未找到列「' + rule.source + '」');
      };
      inS.addEventListener('input', () => { rule.source = inS.value; checkSrc(); updateReady(); });
      checkSrc();
      tdS.appendChild(inS);

      // 变换：自由输入框 + 「预设」下拉（点选即套用，套用后仍可手动编辑）
      const tdF = document.createElement('td');
      const wrapF = document.createElement('div');
      wrapF.className = 'tf-wrap';
      const inF = document.createElement('input');
      inF.type = 'text'; inF.className = 'rule-tf';
      inF.value = rule.transforms; inF.placeholder = 'trim,number';
      inF.addEventListener('input', () => { rule.transforms = inF.value; });
      const selF = document.createElement('select');
      selF.className = 'tf-preset';
      selF.title = '选择预设会替换当前变换内容，选中后仍可手动编辑';
      const phF = document.createElement('option');
      phF.value = ''; phF.textContent = '预设 ▾';
      selF.appendChild(phF);
      TRANSFORM_PRESETS.forEach((p, idx) => {
        const o = document.createElement('option');
        // option 值用下标，避免「清空变换」的空串与占位项冲突
        o.value = String(idx); o.textContent = p.label;
        o.title = p.value || '清空变换';
        selF.appendChild(o);
      });
      selF.addEventListener('change', () => {
        if (selF.value === '') return;                 // 占位项，不触发套用
        const preset = TRANSFORM_PRESETS[Number(selF.value)];
        if (!preset) return;
        inF.value = preset.value;                     // 写入输入框（覆盖已有内容）
        rule.transforms = preset.value;               // 同步到规则对象
        selF.value = '';                              // 复位占位，同一预设可重复选
        updateReady();                                // 刷新就绪状态
      });
      wrapF.append(inF, selF);
      tdF.appendChild(wrapF);

      // 默认值
      const tdD = document.createElement('td');
      const inD = document.createElement('input');
      inD.type = 'text'; inD.value = rule.def; inD.placeholder = '空值兜底';
      inD.addEventListener('input', () => { rule.def = inD.value; });
      tdD.appendChild(inD);

      // 删除
      const tdX = document.createElement('td');
      const del = document.createElement('button');
      del.className = 'btn-del'; del.textContent = '×'; del.title = '删除该行';
      del.addEventListener('click', () => { state.rules.splice(i, 1); renderMapping(); afterLoad(); });
      tdX.appendChild(del);

      tr.append(tdT, tdTy, tdS, tdF, tdD, tdX);
      body.appendChild(tr);
    });

    renderDatalists();
    updateReady();
  }

  function renderDatalists() {
    let dlT = $('dlTarget');
    if (!dlT) { dlT = document.createElement('datalist'); dlT.id = 'dlTarget'; document.body.appendChild(dlT); }
    dlT.innerHTML = state.B ? state.B.headers.map((h) => '<option value="' + esc(h) + '">').join('') : '';

    let dlS = $('dlSource');
    if (!dlS) { dlS = document.createElement('datalist'); dlS.id = 'dlSource'; document.body.appendChild(dlS); }
    dlS.innerHTML = state.A ? state.A.headers.map((h) => '<option value="' + esc(h) + '">').join('') : '';
  }

  /* 同名自动匹配：先精确，再"唯一且不歧义"的包含匹配 */
  function autoMatch(silent) {
    if (!state.A) return;
    const H = state.A.headers;
    const norm = (s) => String(s).replace(/\s+/g, '').replace(/^.*[.．]/, '');
    const used = new Set();
    let n = 0;

    const pend = state.rules.filter((r) => r.type === 'column' && !r.source && r.target);
    if (!pend.length) { if (!silent) log('没有待匹配的空来源列', 'warn'); return; }

    // 1) 归一化后完全相同
    pend.forEach((r) => {
      const t = norm(r.target);
      if (!t) return;
      const hit = H.find((h) => norm(h) === t && !used.has(h));
      if (hit) { r.source = hit; used.add(hit); n++; }
    });

    // 2) 包含匹配：候选唯一，且不被多个目标列同时争抢
    const claim = new Map();
    pend.filter((r) => !r.source).forEach((r) => {
      const t = norm(r.target);
      if (!t) return;
      const hits = H.filter((h) => !used.has(h) && norm(h) && (t.indexOf(norm(h)) >= 0 || norm(h).indexOf(t) >= 0));
      if (hits.length === 1) claim.set(hits[0], (claim.get(hits[0]) || []).concat(r));
    });
    claim.forEach((rs, col) => {
      if (rs.length === 1) { rs[0].source = col; used.add(col); n++; }
    });

    if (!silent || n) log('自动匹配完成：' + n + ' 列已填充来源（其余请手动指定）', 'ok');
    renderMapping();
  }

  /* ------------------------------------------------------------------ */
  /* 计算                                                               */
  /* ------------------------------------------------------------------ */
  function rowDict(st, row) {
    const d = {};
    st.headers.forEach((h, i) => { d[h] = row[i]; });
    return d;
  }

  function colIndexOf(st, spec) {
    if (spec === null || spec === undefined || spec === '') return -1;
    if (typeof spec === 'number') return spec >= 1 ? spec - 1 : spec;
    const s = String(spec).trim();
    if (/^\d+$/.test(s)) return parseInt(s, 10) - 1;
    if (/^[A-Za-z]{1,3}$/.test(s)) {
      let n = 0;
      for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
      return n - 1;
    }
    let i = st.headers.indexOf(s);
    if (i >= 0) return i;
    const hits = st.headers.map((h, k) => h.indexOf(s) >= 0 ? k : -1).filter((k) => k >= 0);
    return hits.length === 1 ? hits[0] : -1;
  }

  const TPL = /\{([^{}]+)\}/g;

  /* expr 表达式求值：列名（合法标识符，中文可用）直接作变量，
   * 任意列可用 row["列名"] 访问；num(x) 转数值、text(x) 转文本。结果按表达式缓存编译。 */
  const _exprCache = new Map();
  const _numH = (v) => { const f = parseFloat(toText(v).replace(/,/g, '')); return isNaN(f) ? 0 : f; };
  function evalExpr(expr, dict) {
    let ent = _exprCache.get(expr);
    if (!ent) {
      const names = Object.keys(dict).filter((k) =>
        /^[A-Za-z_一-龥][\w一-龥]*$/.test(k) && ['row', 'num', 'text'].indexOf(k) < 0);
      ent = { names: names, f: new Function('row', 'num', 'text', ...names, '"use strict"; return (' + expr + ');') };
      _exprCache.set(expr, ent);
    }
    return ent.f(dict, _numH, toText, ...ent.names.map((k) => dict[k]));
  }

  function resolveValue(rule, dict, st, rowIndex, onMissing) {
    if (rule.type === 'const') return rule.source;
    if (rule.type === 'index') return rowIndex;
    if (rule.type === 'expr') return evalExpr(rule.source, dict);

    const src = rule.source;
    if (rule.type === 'template' || (typeof src === 'string' && TPL.test(src))) {
      TPL.lastIndex = 0;
      return String(src).replace(TPL, (m, k) => {
        const key = String(k).trim();
        if (key === '#index' || key === 'index' || key === '序号') return String(rowIndex);
        return toText(dict[key]);
      });
    }
    const ci = colIndexOf(st, src);
    if (ci >= 0) return dict[st.headers[ci]];
    if (typeof src === 'string' && src in dict) return dict[src];
    if (onMissing === 'error') throw new Error('源列不存在：' + src);
    if (onMissing === 'keep') return String(src);
    return null;
  }

  function computeRows(limit, st) {
    st = st || state.A;
    const onMissing = $('optOnMissing').value;
    const skipEmpty = $('optSkipEmpty').checked;
    const rf = readFilter();
    const sp = readSplit();
    const ag = readAgg();
    let dicts = [];
    const start = st.headerRow; // 0 基下标 = headerRow
    for (let i = start; i < st.matrix.length; i++) {
      const row = st.matrix[i] || [];
      if (skipEmpty && row.every((v) => isBlank(v))) continue;
      const dict = rowDict(st, row);
      // 行过滤：按多条件（且/或）跳过（如末尾 TOTAL / 客户要求）
      if (rf && !rowPassesFilter(rf, dict)) continue;
      dicts.push(dict);
    }
    // 拆行（一行→多行）→ 聚合（多行→一行）
    dicts = applyAggregate(applySplit(dicts, sp), ag);
    const out = [];
    for (const dict of dicts) {
      const rec = [];
      state.rules.forEach((r) => {
        if (!r.target) return;
        let v;
        try {
          v = resolveValue(r, dict, st, out.length + 1, onMissing);
        } catch (e) {
          throw new Error('目标列「' + r.target + '」取值失败：' + e.message);
        }
        if (isBlank(v) && r.def !== '' && r.def !== undefined && r.def !== null) v = r.def;
        parseTransforms(r.transforms).forEach((t) => { v = applyTransform(v, t); });
        rec.push(v);
      });
      out.push(rec);
      if (limit && out.length >= limit) break;
    }
    return out;
  }

  /* 预览用：按源行逐一计算，并标注「空行 / 被过滤」跳过状态（用于灰显） */
  function previewData(limit) {
    const st = state.A;
    const onMissing = $('optOnMissing').value;
    const skipEmpty = $('optSkipEmpty').checked;
    const rf = readFilter();
    const items = [];
    let kept = 0, nEmpty = 0, nFiltered = 0;
    const start = st.headerRow;
    for (let i = start; i < st.matrix.length; i++) {
      const row = st.matrix[i] || [];
      let skip = null;
      if (skipEmpty && row.every((v) => isBlank(v))) skip = '空行';
      const dict = rowDict(st, row);
      if (!skip && rf && !rowPassesFilter(rf, dict)) skip = '被过滤';
      const rec = [];
      state.rules.forEach((r) => {
        if (!r.target) return;
        let v;
        try {
          v = resolveValue(r, dict, st, kept + 1, onMissing);
        } catch (e) {
          v = '⚠ ' + e.message;
        }
        if (isBlank(v) && r.def !== '' && r.def !== undefined && r.def !== null) v = r.def;
        parseTransforms(r.transforms).forEach((t) => { v = applyTransform(v, t); });
        rec.push(v);
      });
      if (skip) { if (skip === '空行') nEmpty++; else nFiltered++; } else kept++;
      items.push({ rec: rec, skip: skip });
      if (limit && items.length >= limit) break;
    }
    return { items: items, kept: kept, nEmpty: nEmpty, nFiltered: nFiltered };
  }

  function activeTargets() { return state.rules.filter((r) => r.target).map((r) => r.target); }

  /* 读取行过滤配置（多条条件，且/或组合）；无有效条件返回 null */
  function readFilter() {
    if (!$('optFilter').checked) return null;
    const rules = state.filterRules.filter((r) => r.column && r.match && r.match.trim());
    if (!rules.length) return null;
    return { combine: $('fltCombine').value, rules: rules };
  }

  /* 拆行 / 聚合：读取 UI 配置 */
  function readSplit() {
    if (!$('optSplit').checked) return null;
    const col = $('splitCol').value;
    if (!col) return null;
    return { column: col, delimiter: $('splitDelim').value || '/' };
  }
  function readAgg() {
    if (!$('optAgg').checked) return null;
    const col = $('aggCol').value;
    if (!col) return null;
    const sum = $('aggSum').value.split(/[,，]/).map((s) => s.trim()).filter(Boolean);
    return { group_by: [col], sum: sum };
  }
  function applySplitUI(sp) {
    $('optSplit').checked = !!(sp && sp.column);
    if (sp && sp.column) {
      ensureOption($('splitCol'), sp.column);
      $('splitCol').value = sp.column;
      $('splitDelim').value = sp.delimiter || '/';
    }
  }
  function applyAggUI(ag) {
    const gb = ag && ag.group_by ? (Array.isArray(ag.group_by) ? ag.group_by[0] : ag.group_by) : '';
    $('optAgg').checked = !!gb;
    if (gb) {
      ensureOption($('aggCol'), gb);
      $('aggCol').value = gb;
      $('aggSum').value = (ag.sum || []).join(',');
    }
  }

  /* 拆行：一行 → 多行（空片段跳过） */
  function applySplit(dicts, sp) {
    if (!sp) return dicts;
    const out = [];
    dicts.forEach((d) => {
      toText(d[sp.column]).split(sp.delimiter || '/').forEach((p) => {
        const v = p.trim();
        if (v === '') return;
        const d2 = Object.assign({}, d);
        d2[sp.column] = v;
        out.push(d2);
      });
    });
    return out;
  }

  /* 聚合：按分组列汇总，sum 列求和，其余列取组内第一行 */
  function applyAggregate(dicts, ag) {
    if (!ag) return dicts;
    const groups = new Map();
    dicts.forEach((d) => {
      const key = ag.group_by.map((g) => toText(d[g])).join(' ');
      let g = groups.get(key);
      if (!g) {
        g = Object.assign({}, d);
        ag.sum.forEach((sc) => { g[sc] = _numH(d[sc]); });
        groups.set(key, g);
      } else {
        ag.sum.forEach((sc) => { g[sc] = _numH(g[sc]) + _numH(d[sc]); });
      }
    });
    return Array.from(groups.values());
  }

  /* ------------------------------------------------------------------ */
  /* 预览                                                               */
  /* ------------------------------------------------------------------ */
  $('btnPreview').addEventListener('click', () => {
    try {
      const heads = activeTargets();
      const data = previewData(12);
      $('previewEmpty').hidden = true;
      $('previewWrap').hidden = false;
      $('pvHead').innerHTML = '<tr><th class="pv-flag"></th>' + heads.map((h) => '<th>' + esc(h) + '</th>').join('') + '</tr>';
      $('pvBody').innerHTML = data.items.map((it) =>
        '<tr' + (it.skip ? ' class="pv-skip"' : '') + '>' +
        '<td class="pv-flag">' + (it.skip || '') + '</td>' +
        it.rec.map((v) => '<td>' + esc(toText(v)) + '</td>').join('') + '</tr>').join('');
      const total = computeRows(0).length;
      const skipped = data.nEmpty + data.nFiltered;
      $('previewHint').textContent = '预览前 ' + data.items.length + ' 行（灰显 ' + skipped + ' 行将被跳过）· 拆行/聚合后实际写入 ' + total + ' 行';
      log('预览完成：写入 ' + total + ' 行 × ' + heads.length + ' 列' + (skipped ? '，灰显跳过 ' + skipped + ' 行' : ''), 'ok');
    } catch (e) {
      log('预览出错：' + e.message, 'err');
    }
  });

  /* ------------------------------------------------------------------ */
  /* 转换                                                               */
  /* ------------------------------------------------------------------ */
  function buildConfig() {
    const opts = {
      skip_empty_source_rows: $('optSkipEmpty').checked,
      clear_existing_data: $('optClear').checked,
      on_missing_source: $('optOnMissing').value,
      create_missing_columns: $('optCreateCol').checked
    };
    const rf = readFilter();
    if (rf) opts.row_filter = rf;
    const sp = readSplit();
    if (sp) opts.split = sp;
    const ag = readAgg();
    if (ag) opts.aggregate = ag;
    const mapping = state.rules.filter((r) => r.target).map((r) => {
      const m = { target: r.target, type: r.type };
      if (r.type === 'const') { m.value = r.source; }
      else if (r.type === 'index') { m.source = '@index'; }
      else { m.source = r.source; }
      const ts = parseTransforms(r.transforms);
      if (ts.length) m.transforms = ts;
      if (r.def !== '' && r.def !== undefined && r.def !== null) m.default = r.def;
      return m;
    });
    return {
      source: { path: state.A ? state.A.name : '', sheet: state.A ? state.A.sheet : null, header_row: state.A ? state.A.headerRow : 1 },
      target: { template: state.B ? state.B.name : null, path: $('outName').value || '转换结果.xlsx', sheet: state.B ? state.B.sheet : null, header_row: state.B ? state.B.headerRow : 1 },
      mapping: mapping,
      options: opts
    };
  }

  function colLetter(n) { // 1 基 -> A
    let s = '';
    while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  /* 浏览器端执行（st 默认为主 A 文件，批量时传入各文件） */
  function runBrowser(cfg, st) {
    const rows = computeRows(0, st || state.A);
    let wb, ws, dataStart;

    if (state.B) {
      wb = XLSX.read(new Uint8Array(state.B.buf), { type: 'array', cellDates: true, cellStyles: true });
      ws = wb.Sheets[state.B.sheet];
      const headerRow = state.B.headerRow;
      dataStart = headerRow + 1;
      const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
      const bHeaders = buildHeaders(XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true }), headerRow);

      // 目标列定位
      const colIdx = cfg.mapping.map((m) => mapTargetCol(ws, bHeaders, headerRow, m.target, cfg.options.create_missing_columns));

      if (cfg.options.clear_existing_data) {
        for (let R = dataStart; R <= range.e.r + 1; R++) {
          for (let C = 1; C <= range.e.c + 1; C++) delete ws[colLetter(C) + R];
        }
      }
      rows.forEach((rec, ri) => {
        const R = dataStart + ri;
        rec.forEach((v, ci) => {
          ws[colLetter(colIdx[ci]) + R] = (v === null || v === undefined) ? { t: 's', v: '' } : toCell(v);
        });
      });
      // 裁剪合并单元格：丢弃起始行落在数据区内的合并
      if (ws['!merges']) ws['!merges'] = ws['!merges'].filter((m) => m.s.r < dataStart - 1);
      // 修正范围
      const lastCol = Math.max(range.e.c, Math.max.apply(null, colIdx.map((c) => c - 1)));
      ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(range.e.r, dataStart - 1 + rows.length - 1), c: lastCol } });
    } else {
      const heads = activeTargets();
      const aoa = [heads].concat(rows.map((r) => r.map((v) => (v === null || v === undefined) ? '' : v)));
      ws = XLSX.utils.aoa_to_sheet(aoa);
      wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    }
    XLSX.writeFile(wb, cfg.target.path, { cellDates: true });
    return rows.length;
  }

  function toCell(v) {
    if (v instanceof Date) return { t: 'd', v: v };
    if (typeof v === 'number') return { t: 'n', v: v };
    if (typeof v === 'boolean') return { t: 'b', v: v };
    return { t: 's', v: toText(v) };
  }

  function mapTargetCol(ws, bHeaders, headerRow, spec, createMissing) {
    const s = String(spec).trim();
    if (/^\d+$/.test(s)) return parseInt(s, 10);
    if (/^[A-Za-z]{1,3}$/.test(s)) { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
    let i = bHeaders.indexOf(s);
    if (i >= 0) return i + 1;
    const hits = bHeaders.map((h, k) => h.indexOf(s) >= 0 ? k : -1).filter((k) => k >= 0);
    if (hits.length === 1) return hits[0] + 1;
    if (createMissing) {
      const n = bHeaders.length + 1;
      ws[colLetter(n) + headerRow] = { t: 's', v: s };
      return n;
    }
    throw new Error('B 表表头未找到列：' + spec);
  }

  /* 本地引擎执行（st 默认为主 A 文件，批量时传入各文件） */
  async function runLocal(cfg, st) {
    st = st || state.A;
    const toB64 = (buf) => {
      const bytes = new Uint8Array(buf); let bin = '';
      const CH = 0x8000;
      for (let i = 0; i < bytes.length; i += CH) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CH));
      return btoa(bin);
    };
    const payload = {
      source: { b64: toB64(st.buf), name: st.name, sheet: st.sheet, header_row: st.headerRow },
      target: state.B ? { b64: toB64(state.B.buf), name: state.B.name, sheet: state.B.sheet, header_row: state.B.headerRow } : null,
      mapping: cfg.mapping, options: cfg.options, out_name: cfg.target.path
    };
    const resp = await fetch('/api/convert', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    const data = await resp.json();
    if (!data.ok) throw new Error(data.error || '本地引擎执行失败');
    const bin = atob(data.b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = data.filename || cfg.target.path; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    return data.stats;
  }

  /* 批量转换：主文件 + batchA 队列逐个套用同一配置，输出名 = 源文件名_输出名 */
  async function runBatch(cfg) {
    const items = [{ st: state.A }].concat(state.batchA.map((f) => ({ file: f })));
    let ok = 0, fail = 0;
    for (const it of items) {
      let st = it.st;
      try {
        if (!st) st = await readFileState(it.file);
        const outName = stem(st.name) + '_' + cfg.target.path;
        const cfg2 = Object.assign({}, cfg, { target: Object.assign({}, cfg.target, { path: outName }) });
        if (state.mode === 'local') await runLocal(cfg2, st);
        else runBrowser(cfg2, st);
        ok++;
        log('批量 ✔ ' + st.name + ' → ' + outName, 'ok');
      } catch (e) {
        fail++;
        log('批量 ✖ ' + (st ? st.name : it.file.name) + '：' + e.message, 'err');
      }
    }
    return { ok: ok, fail: fail };
  }

  $('btnRun').addEventListener('click', async () => {
    const cfg = buildConfig();
    if (!cfg.mapping.length) { log('请先配置至少一条映射', 'err'); return; }
    $('btnRun').disabled = true;
    setStatus('正在转换…');
    try {
      const heads = activeTargets();
      if (state.batchA.length) {
        const r = await runBatch(cfg);
        log('批量转换完成：成功 ' + r.ok + ' 个，失败 ' + r.fail + ' 个', r.fail ? 'warn' : 'ok');
        setStatus('批量完成：成功 ' + r.ok + ' / 共 ' + (r.ok + r.fail) + ' 个文件');
      } else if (state.mode === 'local') {
        const stats = await runLocal(cfg);
        const skipInfo = stats.skipped_by_filter ? '，过滤跳过 ' + stats.skipped_by_filter + ' 行' : '';
        log('转换完成（本地引擎）：写入 ' + stats.written_rows + ' 行 × ' + stats.target_columns + ' 列' + skipInfo, 'ok');
        setStatus('完成：' + stats.written_rows + ' 行已写入，模板格式已保留');
      } else {
        const n = runBrowser(cfg);
        log('转换完成（浏览器）：写入 ' + n + ' 行 × ' + heads.length + ' 列', 'ok');
        setStatus('完成：' + n + ' 行已写入并下载');
      }
      memSave();
      $('previewHint').textContent = '最近一次：' + new Date().toLocaleTimeString();
    } catch (e) {
      log('转换失败：' + e.message, 'err');
      setStatus('转换失败');
    } finally {
      $('btnRun').disabled = false;
      updateReady();
    }
  });

  /* ------------------------------------------------------------------ */
  /* 导入 / 导出                                                         */
  /* ------------------------------------------------------------------ */
  $('btnExport').addEventListener('click', () => {
    const cfg = buildConfig();
    const blob = new Blob([JSON.stringify(cfg, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = 'mapping.json'; a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 3000);
    log('已导出 mapping.json（可直接用于命令行引擎）', 'ok');
  });

  $('btnImport').addEventListener('click', () => $('fileJson').click());
  $('fileJson').addEventListener('change', (e) => {
    const f = e.target.files[0]; if (!f) return;
    const r = new FileReader();
    r.onload = (ev) => {
      try {
        const cfg = JSON.parse(ev.target.result);
        state.rules = (cfg.mapping || []).map((m) => ({
          target: m.target,
          type: m.type || (m.value !== undefined ? 'const' : (m.source === '@index' ? 'index' : (String(m.source || '').indexOf('{') >= 0 ? 'template' : 'column'))),
          source: m.value !== undefined ? m.value : (m.source === '@index' ? '' : (m.source || '')),
          transforms: transformsToText(m.transforms),
          def: m.default !== undefined ? m.default : ''
        }));
        if (cfg.source && cfg.source.header_row) $('headA').value = cfg.source.header_row;
        if (cfg.target) {
          if (cfg.target.header_row) $('headB').value = cfg.target.header_row;
          if (cfg.target.path) $('outName').value = cfg.target.path;
        }
        if (cfg.options) {
          $('optSkipEmpty').checked = cfg.options.skip_empty_source_rows !== false;
          $('optClear').checked = cfg.options.clear_existing_data !== false;
          $('optCreateCol').checked = !!cfg.options.create_missing_columns;
          if (cfg.options.on_missing_source) $('optOnMissing').value = cfg.options.on_missing_source;
          const rf = cfg.options.row_filter;
          if (rf) {
            $('optFilter').checked = true;
            let combine = 'and', rules = [];
            if (Array.isArray(rf)) { rules = rf; }
            else if (rf.rules) { combine = rf.combine || 'and'; rules = rf.rules; }
            else { rules = [rf]; }
            $('fltCombine').value = combine;
            state.filterRules = rules.map((r) => ({
              column: r.column || r.source || '',
              match: r.match || r.pattern || r.regex || '',
              mode: r.mode || 'include'
            }));
            state.filterTouched = true;
            renderFilterRules();
          }
          if (cfg.options.split) applySplitUI(cfg.options.split);
          if (cfg.options.aggregate) applyAggUI(cfg.options.aggregate);
        }
        log('已导入配置：' + state.rules.length + ' 条映射（请确认已加载对应文件）', 'ok');
        renderMapping(); updateReady();
      } catch (err) { log('导入失败：' + err.message, 'err'); }
    };
    r.readAsText(f);
    e.target.value = '';
  });

  /* ------------------------------------------------------------------ */
  /* 其它按钮                                                            */
  /* ------------------------------------------------------------------ */
  $('btnDetect').addEventListener('click', () => {
    let changed = 0;
    ['A', 'B'].forEach((k) => {
      const st = state[k]; if (!st) return;
      const hr = detectHeaderRow(st.matrix);
      if (hr !== st.headerRow) { st.headerRow = hr; $(k === 'A' ? 'headA' : 'headB').value = hr; changed++; }
    });
    ['A', 'B'].forEach(reparse);
    log(changed ? '已重新识别表头行（' + changed + ' 个文件调整）' : '表头行识别结果未变化', 'ok');
  });

  $('btnReload').addEventListener('click', () => { reparse('A'); reparse('B'); log('已按当前设置重新读取', 'ok'); });
  $('sheetA').addEventListener('change', () => reparse('A'));
  $('sheetB').addEventListener('change', () => reparse('B'));
  $('headA').addEventListener('change', () => reparse('A'));
  $('headB').addEventListener('change', () => reparse('B'));
  $('btnAuto').addEventListener('click', () => autoMatch(false));
  $('btnAdd').addEventListener('click', () => {
    state.rules.push(newRule());
    renderMapping(); afterLoad();
  });
  $('btnClearMap').addEventListener('click', () => {
    state.rules.forEach((r) => { r.source = ''; r.transforms = ''; r.def = ''; });
    renderMapping(); updateReady();
    log('已清空所有来源（保留目标列）', 'ok');
  });

  wireDrop('dropA', 'fileA', 'clearA', 'A');
  wireDrop('dropB', 'fileB', 'clearB', 'B');
  $('optFilter').addEventListener('change', () => { state.filterTouched = true; toggleFilterFields(); });
  $('btnAddFilter').addEventListener('click', () => {
    state.filterRules.push({ column: '', match: '', mode: 'include' });
    state.filterTouched = true;
    renderFilterRules();
  });
  $('btnBatchClear').addEventListener('click', () => {
    state.batchA = [];
    renderBatchBar();
    updateReady();
    log('已取消批量模式', 'ok');
  });

  /* ------------------------------------------------------------------ */
  /* 启动：探测本地引擎                                                   */
  /* ------------------------------------------------------------------ */
  (async function boot() {
    // file:// 打开时不做服务探测（避免无谓的跨域报错），直接用浏览器引擎
    if (location.protocol === 'file:') {
      log('以本地文件方式打开，使用浏览器端引擎', 'ok');
      state.started = true;
      return afterLoad();
    }
    try {
      const ctl = new AbortController();
      const t = setTimeout(() => ctl.abort(), 1200);
      const r = await fetch('/api/health', { signal: ctl.signal });
      clearTimeout(t);
      const j = await r.json();
      if (j && j.ok) {
        state.mode = 'local';
        $('modeBadge').classList.add('local');
        $('modeText').textContent = '本地引擎模式';
        log('检测到本地引擎，将使用 Python 引擎转换（完整保留模板格式）', 'ok');
      }
    } catch (e) { /* 保持浏览器模式 */ }
    state.started = true;
    afterLoad();
  })();

})();
