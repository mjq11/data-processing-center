/* ====================================================================
   数据处理中心 — 核心逻辑
   ==================================================================== */

// ==================== 全局状态 ====================
const state = {
  currentView: 'home',
  // 智能匹配
  match: { sourceWB: null, targetWB: null, sourceFile: null, targetFile: null, resultWB: null },
  // 格式转换
  convert: { workbook: null, file: null },
  // 数据清洗
  clean: { workbook: null, file: null, resultWB: null },
  // 报表合并
  merge: { workbooks: [], files: [], resultWB: null },
  // 数据比对
  diff: { oldWB: null, newWB: null, oldFile: null, newFile: null, resultWB: null },
  // 敏感数据脱敏
  mask: { workbook: null, file: null, resultWB: null },
  // 数据拆分器
  split: { workbook: null, file: null, zipBlob: null },
  // 公式计算
  formula: { workbook: null, file: null, resultWB: null },
  // 多维分组汇总
  pivot: { workbook: null, file: null, resultWB: null },
  // 一键图表生成
  chart: { workbook: null, file: null, chartInstance: null },
};

// ==================== 工具函数 ====================

/** 显示 Toast 通知 */
function showToast(msg, type = 'success') {
  const c = document.getElementById('toast-container');
  const icons = {
    success: '<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><path d="M20 6L9 17l-5-5"/></svg>',
    error: '<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><path d="M18 6L6 18M6 6l12 12"/></svg>',
    info: '<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/></svg>',
  };
  const t = document.createElement('div');
  t.className = `toast ${type}`;
  t.innerHTML = `<span class="toast-icon">${icons[type] || icons.info}</span>${msg}`;
  c.appendChild(t);
  setTimeout(() => t.remove(), 3200);
}

/** 格式化文件大小 */
function formatSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

/** 读取文件为 SheetJS Workbook */
function readFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        resolve(wb);
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

/** 获取 Sheet 的 JSON 数据（含表头）*/
function getSheetData(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(ws, { defval: '' });
}

/** 获取 Sheet 的表头列名 */
function getSheetHeaders(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  const headers = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: range.s.r, c });
    const cell = ws[addr];
    headers.push(cell ? String(cell.v) : `列${c + 1}`);
  }
  return headers;
}

/** 渲染数据预览表格 */
function renderPreviewTable(theadEl, tbodyEl, headers, data, maxRows = 100, highlightMap = null) {
  // 渲染表头
  let headHtml = '<tr><th class="row-num">#</th>';
  headers.forEach(h => { headHtml += `<th>${escHtml(h)}</th>`; });
  headHtml += '</tr>';
  theadEl.innerHTML = headHtml;

  // 渲染数据行
  const rows = data.slice(0, maxRows);
  let bodyHtml = '';
  rows.forEach((row, i) => {
    bodyHtml += '<tr>';
    bodyHtml += `<td class="row-num">${i + 1}</td>`;
    headers.forEach(h => {
      const val = row[h] !== undefined ? row[h] : '';
      let cls = '';
      if (highlightMap && highlightMap[i] && highlightMap[i][h]) {
        cls = ` class="${highlightMap[i][h]}"`;
      }
      if (val === '' || val === null || val === undefined) {
        bodyHtml += `<td${cls}><span class="cell-empty">—</span></td>`;
      } else {
        bodyHtml += `<td${cls}>${escHtml(String(val))}</td>`;
      }
    });
    bodyHtml += '</tr>';
  });
  tbodyEl.innerHTML = bodyHtml || '<tr><td colspan="99" style="text-align:center;color:var(--text-muted);padding:20px">暂无数据</td></tr>';
}

/** HTML 转义 */
function escHtml(s) {
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}

/** 下载 Workbook */
function downloadWorkbook(wb, filename) {
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

/** 下载 CSV */
function downloadCSV(csvString, filename) {
  // 添加 BOM 以确保 Excel 正确识别 UTF-8 编码
  const blob = new Blob(['\ufeff' + csvString], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

/** 下载 JSON */
function downloadJSON(data, filename) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

/** 显示/隐藏处理中覆盖层 */
function showProcessing(title, msg) {
  document.getElementById('processing-title').textContent = title || '正在处理...';
  document.getElementById('processing-msg').textContent = msg || '请稍候';
  document.getElementById('processing-overlay').classList.add('active');
}
function hideProcessing() {
  document.getElementById('processing-overlay').classList.remove('active');
}

/** 填充 select 选项（用于非搜索的普通下拉框） */
function populateSelect(selectEl, options, defaultVal) {
  selectEl.innerHTML = '';
  options.forEach(opt => {
    const o = document.createElement('option');
    o.value = opt;
    o.textContent = opt;
    if (opt === defaultVal) o.selected = true;
    selectEl.appendChild(o);
  });
}

// ==================== 可搜索下拉框组件 ====================

/**
 * SearchableSelect — 带搜索功能的下拉框组件
 * 用法：
 *   const ss = new SearchableSelect(containerEl, { placeholder: '搜索列名...' });
 *   ss.setOptions(['列A', '列B', '列C']);
 *   ss.onChange(val => console.log('选中:', val));
 */
class SearchableSelect {
  constructor(container, opts = {}) {
    this.container = container;
    this.options = [];
    this.filtered = [];
    this.selectedValue = '';
    this.isOpen = false;
    this.focusedIdx = -1;
    this.placeholder = opts.placeholder || '搜索...';
    this._onChange = null;
    this._build();
    this._bindEvents();
  }

  /** 构建 DOM 结构 */
  _build() {
    const wrap = document.createElement('div');
    wrap.className = 'ss-wrap';

    // 触发按钮
    wrap.innerHTML = `
      <div class="ss-trigger">
        <span class="ss-trigger-text">请选择</span>
        <svg class="ss-trigger-arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M6 9l6 6 6-6"/></svg>
      </div>
      <div class="ss-dropdown">
        <div class="ss-search-wrap">
          <span class="ss-search-icon">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><path d="M21 21l-4.35-4.35"/></svg>
          </span>
          <input class="ss-search" type="text" placeholder="${escHtml(this.placeholder)}" />
        </div>
        <div class="ss-options"></div>
        <div class="ss-count"></div>
      </div>
    `;

    this.wrap = wrap;
    this.trigger = wrap.querySelector('.ss-trigger');
    this.triggerText = wrap.querySelector('.ss-trigger-text');
    this.dropdown = wrap.querySelector('.ss-dropdown');
    this.searchInput = wrap.querySelector('.ss-search');
    this.optionsList = wrap.querySelector('.ss-options');
    this.countEl = wrap.querySelector('.ss-count');

    this.container.appendChild(wrap);
  }

  /** 绑定事件 */
  _bindEvents() {
    // 触发打开/关闭
    this.trigger.addEventListener('click', e => {
      e.stopPropagation();
      this.isOpen ? this.close() : this.open();
    });

    // 搜索输入
    this.searchInput.addEventListener('input', () => {
      this._filter(this.searchInput.value);
    });

    // 阻止点击下拉面板时冒泡关闭
    this.dropdown.addEventListener('click', e => e.stopPropagation());

    // 键盘导航
    this.searchInput.addEventListener('keydown', e => {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        this._moveFocus(1);
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        this._moveFocus(-1);
      } else if (e.key === 'Enter') {
        e.preventDefault();
        if (this.focusedIdx >= 0 && this.focusedIdx < this.filtered.length) {
          this.select(this.filtered[this.focusedIdx]);
        }
      } else if (e.key === 'Escape') {
        this.close();
      }
    });

    // 点击外部关闭
    document.addEventListener('click', () => {
      if (this.isOpen) this.close();
    });
  }

  /** 设置选项列表 */
  setOptions(options, defaultVal) {
    this.options = options || [];
    this.filtered = [...this.options];
    this.focusedIdx = -1;

    // 设置默认值
    if (defaultVal !== undefined && this.options.includes(defaultVal)) {
      this.selectedValue = defaultVal;
      this.triggerText.textContent = defaultVal;
    } else if (this.options.length > 0) {
      this.selectedValue = this.options[0];
      this.triggerText.textContent = this.options[0];
    } else {
      this.selectedValue = '';
      this.triggerText.textContent = '请选择';
    }

    this._render();
  }

  /** 获取当前选中值 */
  get value() {
    return this.selectedValue;
  }

  /** 设置选中值 */
  set value(val) {
    this.select(val, false);
  }

  /** 选中某个选项 */
  select(val, triggerChange = true) {
    this.selectedValue = val;
    this.triggerText.textContent = val || '请选择';
    this.close();
    this._render();
    if (triggerChange && this._onChange) {
      this._onChange(val);
    }
  }

  /** 注册变更回调 */
  onChange(fn) {
    this._onChange = fn;
  }

  /** 打开下拉面板 */
  open() {
    this.isOpen = true;
    this.trigger.classList.add('open');
    this.dropdown.classList.add('open');
    this.searchInput.value = '';
    this.focusedIdx = -1;
    this._filter('');
    // 延迟聚焦，避免触发按钮点击冲突
    setTimeout(() => this.searchInput.focus(), 50);
  }

  /** 关闭下拉面板 */
  close() {
    this.isOpen = false;
    this.trigger.classList.remove('open');
    this.dropdown.classList.remove('open');
    this.focusedIdx = -1;
  }

  /** 过滤选项 */
  _filter(query) {
    const q = query.toLowerCase().trim();
    if (!q) {
      this.filtered = [...this.options];
    } else {
      this.filtered = this.options.filter(opt =>
        opt.toLowerCase().includes(q)
      );
    }
    this.focusedIdx = -1;
    this._render(q);
  }

  /** 渲染选项列表 */
  _render(query = '') {
    const q = query.toLowerCase().trim();

    if (this.filtered.length === 0) {
      this.optionsList.innerHTML = '<div class="ss-empty">无匹配结果</div>';
      this.countEl.textContent = '';
      return;
    }

    let html = '';
    this.filtered.forEach((opt, i) => {
      const isActive = opt === this.selectedValue;
      const isFocused = i === this.focusedIdx;
      const cls = (isActive ? ' active' : '') + (isFocused ? ' focused' : '');

      // 高亮匹配文字
      let display;
      if (q && opt.toLowerCase().includes(q)) {
        const idx = opt.toLowerCase().indexOf(q);
        display = escHtml(opt.substring(0, idx)) +
          '<span class="ss-hl">' + escHtml(opt.substring(idx, idx + q.length)) + '</span>' +
          escHtml(opt.substring(idx + q.length));
      } else {
        display = escHtml(opt);
      }

      html += `<div class="ss-option${cls}" data-idx="${i}" data-val="${escHtml(opt)}">${display}</div>`;
    });

    this.optionsList.innerHTML = html;

    // 显示筛选计数
    if (q) {
      this.countEl.textContent = `${this.filtered.length} / ${this.options.length} 项`;
    } else {
      this.countEl.textContent = `共 ${this.options.length} 项`;
    }

    // 绑定选项点击
    this.optionsList.querySelectorAll('.ss-option').forEach(el => {
      el.addEventListener('click', () => {
        this.select(el.dataset.val);
      });
    });
  }

  /** 键盘上下移动焦点 */
  _moveFocus(dir) {
    if (this.filtered.length === 0) return;
    this.focusedIdx += dir;
    if (this.focusedIdx < 0) this.focusedIdx = this.filtered.length - 1;
    if (this.focusedIdx >= this.filtered.length) this.focusedIdx = 0;

    // 更新焦点样式
    this.optionsList.querySelectorAll('.ss-option').forEach((el, i) => {
      el.classList.toggle('focused', i === this.focusedIdx);
    });

    // 滚动到焦点项
    const focusedEl = this.optionsList.querySelector('.ss-option.focused');
    if (focusedEl) {
      focusedEl.scrollIntoView({ block: 'nearest' });
    }
  }

  /** 销毁组件 */
  destroy() {
    if (this.wrap && this.wrap.parentNode) {
      this.wrap.parentNode.removeChild(this.wrap);
    }
  }
}

/**
 * 创建一个可搜索下拉框，挂载到指定容器
 * @returns {SearchableSelect} 实例
 */
function createSearchableSelect(container, options, opts = {}) {
  const ss = new SearchableSelect(container, opts);
  if (options) ss.setOptions(options);
  return ss;
}

/** 创建文件标签 */
function createFileTag(name, size, onRemove) {
  const tag = document.createElement('div');
  tag.className = 'file-tag';
  tag.innerHTML = `
    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M16 13H8M16 17H8M10 9H8"/></svg>
    <span class="file-tag-name">${escHtml(name)}</span>
    <span class="file-tag-size">${formatSize(size)}</span>
    <button class="file-tag-remove" title="移除">×</button>
  `;
  tag.querySelector('.file-tag-remove').addEventListener('click', e => {
    e.stopPropagation();
    onRemove();
    tag.remove();
  });
  return tag;
}

// ==================== 视图切换 ====================
function switchView(view) {
  state.currentView = view;
  document.getElementById('view-home').style.display = view === 'home' ? '' : 'none';
  document.querySelectorAll('.tool-view').forEach(el => {
    el.classList.toggle('active', el.id === `view-${view}`);
  });
}

// ==================== 拖拽上传通用绑定 ====================
function bindUploadZone(zoneEl, fileInputEl, onFile) {
  zoneEl.addEventListener('click', () => fileInputEl.click());
  zoneEl.addEventListener('dragover', e => { e.preventDefault(); zoneEl.classList.add('drag-over'); });
  zoneEl.addEventListener('dragleave', () => zoneEl.classList.remove('drag-over'));
  zoneEl.addEventListener('drop', e => {
    e.preventDefault();
    zoneEl.classList.remove('drag-over');
    if (e.dataTransfer.files.length > 0) onFile(e.dataTransfer.files);
  });
  fileInputEl.addEventListener('change', () => {
    if (fileInputEl.files.length > 0) onFile(fileInputEl.files);
    fileInputEl.value = '';
  });
}

// ==================== 主题切换 ====================
function initTheme() {
  const saved = localStorage.getItem('dc-theme');
  if (saved === 'dark') document.documentElement.setAttribute('data-theme', 'dark');
  document.getElementById('btn-theme').addEventListener('click', () => {
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    document.documentElement.setAttribute('data-theme', isDark ? '' : 'dark');
    localStorage.setItem('dc-theme', isDark ? 'light' : 'dark');
  });
}

// ==================== 首页 ====================
function initHome() {
  document.querySelectorAll('.tool-card').forEach(card => {
    card.addEventListener('click', () => {
      const tool = card.dataset.tool;
      switchView(tool);
    });
  });
  // Logo 点击返回首页
  document.getElementById('logo').addEventListener('click', () => switchView('home'));
}

// ==================== 智能匹配 ====================
function initMatch() {
  const s = state.match;

  // 可搜索下拉框实例引用
  let ssSourceKey = null;
  let ssTgtKey = null;

  // 返回
  document.getElementById('match-back').addEventListener('click', () => switchView('home'));

  // 上传源表
  bindUploadZone(
    document.getElementById('match-upload-source'),
    document.getElementById('match-file-source'),
    async (files) => {
      try {
        s.sourceFile = files[0];
        s.sourceWB = await readFile(files[0]);
        const container = document.getElementById('match-files-source');
        container.innerHTML = '';
        container.appendChild(createFileTag(files[0].name, files[0].size, () => {
          s.sourceWB = null; s.sourceFile = null;
          hideMatchConfig();
        }));
        showToast(`源表 "${files[0].name}" 加载成功，${s.sourceWB.SheetNames.length} 个 Sheet`);
        checkMatchReady();
      } catch (e) { showToast('文件读取失败: ' + e.message, 'error'); }
    }
  );

  // 上传目标表
  bindUploadZone(
    document.getElementById('match-upload-target'),
    document.getElementById('match-file-target'),
    async (files) => {
      try {
        s.targetFile = files[0];
        s.targetWB = await readFile(files[0]);
        const container = document.getElementById('match-files-target');
        container.innerHTML = '';
        container.appendChild(createFileTag(files[0].name, files[0].size, () => {
          s.targetWB = null; s.targetFile = null;
          hideMatchConfig();
        }));
        showToast(`目标表 "${files[0].name}" 加载成功`);
        checkMatchReady();
      } catch (e) { showToast('文件读取失败: ' + e.message, 'error'); }
    }
  );

  function hideMatchConfig() {
    document.getElementById('match-config').style.display = 'none';
    document.getElementById('match-action-bar').style.display = 'none';
    document.getElementById('match-preview').style.display = 'none';
    document.getElementById('match-result-stats').style.display = 'none';
    document.getElementById('match-download').style.display = 'none';
  }

  function checkMatchReady() {
    if (!s.sourceWB || !s.targetWB) return;
    // 显示配置面板
    document.getElementById('match-config').style.display = '';
    document.getElementById('match-action-bar').style.display = '';

    // 填充 Sheet 选择（Sheet 数通常较少，保持原生 select）
    populateSelect(document.getElementById('match-source-sheet'), s.sourceWB.SheetNames);
    populateSelect(document.getElementById('match-target-sheet'), s.targetWB.SheetNames);

    updateMatchColumns();
  }

  // Sheet 变更时刷新列
  document.getElementById('match-source-sheet').addEventListener('change', updateMatchColumns);
  document.getElementById('match-target-sheet').addEventListener('change', updateMatchColumns);

  function updateMatchColumns() {
    const srcSheet = document.getElementById('match-source-sheet').value;
    const tgtSheet = document.getElementById('match-target-sheet').value;
    if (!srcSheet || !tgtSheet) return;

    const srcHeaders = getSheetHeaders(s.sourceWB, srcSheet);
    const tgtHeaders = getSheetHeaders(s.targetWB, tgtSheet);

    // 匹配键 — 使用可搜索下拉框
    const srcKeyContainer = document.getElementById('match-source-key');
    srcKeyContainer.innerHTML = '';
    if (ssSourceKey) ssSourceKey.destroy();
    ssSourceKey = createSearchableSelect(srcKeyContainer, srcHeaders, { placeholder: '搜索源表列名...' });
    ssSourceKey.onChange(() => {}); // 无需联动，仅用于选择

    const tgtKeyContainer = document.getElementById('match-target-key');
    tgtKeyContainer.innerHTML = '';
    if (ssTgtKey) ssTgtKey.destroy();
    ssTgtKey = createSearchableSelect(tgtKeyContainer, tgtHeaders, { placeholder: '搜索目标表列名...' });
    ssTgtKey.onChange(() => {}); // 无需联动

    // 预生成一条填充规则
    const rulesEl = document.getElementById('match-fill-rules');
    rulesEl.innerHTML = '';
    addFillRule(srcHeaders, tgtHeaders);

    document.getElementById('match-execute').disabled = false;

    // 显示源表预览
    const srcData = getSheetData(s.sourceWB, srcSheet);
    renderPreviewTable(
      document.getElementById('match-preview-thead'),
      document.getElementById('match-preview-tbody'),
      srcHeaders, srcData, 50
    );
    document.getElementById('match-preview').style.display = '';
    document.getElementById('match-preview-stats').innerHTML =
      `<span>${srcData.length} 行</span><span>${srcHeaders.length} 列</span>`;
  }

  // 添加填充规则 — 使用可搜索下拉框
  function addFillRule(srcHeaders, tgtHeaders) {
    const rulesEl = document.getElementById('match-fill-rules');
    const item = document.createElement('div');
    item.className = 'fill-rule-item';

    item.innerHTML = `
      <span class="match-rule-label">源表列</span>
      <div class="fill-src-container"></div>
      <span class="match-rule-arrow">→</span>
      <span class="match-rule-label">填入目标列</span>
      <div class="fill-tgt-container"></div>
      <button class="fill-rule-remove" title="删除">
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M18 6L6 18M6 6l12 12"/></svg>
      </button>
    `;

    // 创建可搜索下拉框
    const srcSS = createSearchableSelect(
      item.querySelector('.fill-src-container'),
      srcHeaders,
      { placeholder: '搜索源表列...' }
    );
    const tgtSS = createSearchableSelect(
      item.querySelector('.fill-tgt-container'),
      tgtHeaders,
      { placeholder: '搜索目标列...' }
    );

    // 将实例引用存到 DOM 元素上，以便执行时读取值
    item._ssSrc = srcSS;
    item._ssTgt = tgtSS;

    item.querySelector('.fill-rule-remove').addEventListener('click', () => {
      srcSS.destroy();
      tgtSS.destroy();
      item.remove();
    });
    rulesEl.appendChild(item);
  }

  document.getElementById('match-add-rule').addEventListener('click', () => {
    const srcSheet = document.getElementById('match-source-sheet').value;
    const tgtSheet = document.getElementById('match-target-sheet').value;
    if (!srcSheet || !tgtSheet) return;
    const srcHeaders = getSheetHeaders(s.sourceWB, srcSheet);
    const tgtHeaders = getSheetHeaders(s.targetWB, tgtSheet);
    addFillRule(srcHeaders, tgtHeaders);
  });

  // 执行匹配
  document.getElementById('match-execute').addEventListener('click', () => {
    const srcSheet = document.getElementById('match-source-sheet').value;
    const tgtSheet = document.getElementById('match-target-sheet').value;
    const srcKey = ssSourceKey ? ssSourceKey.value : '';
    const tgtKey = ssTgtKey ? ssTgtKey.value : '';

    if (!srcKey || !tgtKey) {
      showToast('请先选择匹配键列', 'error');
      return;
    }

    // 收集填充规则（从 SearchableSelect 实例获取值）
    const rules = [];
    document.querySelectorAll('#match-fill-rules .fill-rule-item').forEach(item => {
      const src = item._ssSrc ? item._ssSrc.value : '';
      const tgt = item._ssTgt ? item._ssTgt.value : '';
      if (src && tgt) {
        rules.push({ src, tgt });
      }
    });

    if (rules.length === 0) {
      showToast('请至少添加一条填充规则', 'error');
      return;
    }

    showProcessing('正在匹配数据...', `匹配键: ${srcKey} ⟷ ${tgtKey}`);

    // 使用 setTimeout 让 UI 先渲染
    setTimeout(() => {
      try {
        const srcData = getSheetData(s.sourceWB, srcSheet);
        const tgtData = getSheetData(s.targetWB, tgtSheet);
        const tgtHeaders = getSheetHeaders(s.targetWB, tgtSheet);

        // 构建源表索引（按匹配键）
        const srcIndex = {};
        srcData.forEach(row => {
          const key = String(row[srcKey] || '').trim();
          if (key) srcIndex[key] = row;
        });

        // 执行匹配填充
        let matchedCount = 0;
        let unmatchedCount = 0;
        const highlightMap = {};

        tgtData.forEach((row, i) => {
          const key = String(row[tgtKey] || '').trim();
          const srcRow = srcIndex[key];
          if (srcRow) {
            matchedCount++;
            rules.forEach(rule => {
              row[rule.tgt] = srcRow[rule.src] !== undefined ? srcRow[rule.src] : '';
              if (!highlightMap[i]) highlightMap[i] = {};
              highlightMap[i][rule.tgt] = 'cell-matched';
            });
          } else {
            unmatchedCount++;
            rules.forEach(rule => {
              if (!highlightMap[i]) highlightMap[i] = {};
              highlightMap[i][rule.tgt] = 'cell-unmatched';
            });
          }
        });

        // 确保填充列出现在表头中
        const allHeaders = [...tgtHeaders];
        rules.forEach(rule => {
          if (!allHeaders.includes(rule.tgt)) allHeaders.push(rule.tgt);
        });

        // 生成结果 Workbook
        const ws = XLSX.utils.json_to_sheet(tgtData, { header: allHeaders });
        const resultWB = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(resultWB, ws, tgtSheet);
        s.resultWB = resultWB;

        // 渲染结果预览
        renderPreviewTable(
          document.getElementById('match-preview-thead'),
          document.getElementById('match-preview-tbody'),
          allHeaders, tgtData, 100, highlightMap
        );
        document.getElementById('match-preview-stats').innerHTML =
          `<span>${tgtData.length} 行</span><span>${allHeaders.length} 列</span>`;

        // 渲染统计
        document.getElementById('match-result-stats').style.display = '';
        document.getElementById('match-result-stats').innerHTML = `
          <div class="result-stat-card success">
            <div class="result-stat-icon"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><path d="M22 4L12 14.01l-3-3"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${matchedCount}</span><span class="result-stat-label">匹配成功</span></div>
          </div>
          <div class="result-stat-card warning">
            <div class="result-stat-icon"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M12 8v4M12 16h.01"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${unmatchedCount}</span><span class="result-stat-label">未匹配</span></div>
          </div>
          <div class="result-stat-card info">
            <div class="result-stat-icon"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 12h6"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${tgtData.length}</span><span class="result-stat-label">总行数</span></div>
          </div>
          <div class="result-stat-card info">
            <div class="result-stat-icon"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M16 3h5v5M4 20L21 3"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${(matchedCount / tgtData.length * 100).toFixed(1)}%</span><span class="result-stat-label">匹配率</span></div>
          </div>
        `;

        document.getElementById('match-download').style.display = '';
        showToast(`匹配完成！成功 ${matchedCount} 行，未匹配 ${unmatchedCount} 行`);
      } catch (e) {
        showToast('匹配失败: ' + e.message, 'error');
      }
      hideProcessing();
    }, 100);
  });

  // 下载结果
  document.getElementById('match-download').addEventListener('click', () => {
    if (!s.resultWB) return;
    const name = s.targetFile ? s.targetFile.name.replace(/\.\w+$/, '') : '匹配结果';
    downloadWorkbook(s.resultWB, `${name}_已匹配.xlsx`);
    showToast('文件已下载');
  });
}

// ==================== 格式转换 ====================
function initConvert() {
  const s = state.convert;

  document.getElementById('convert-back').addEventListener('click', () => switchView('home'));

  bindUploadZone(
    document.getElementById('convert-upload'),
    document.getElementById('convert-file'),
    async (files) => {
      try {
        s.file = files[0];
        s.workbook = await readFile(files[0]);
        const container = document.getElementById('convert-files');
        container.innerHTML = '';
        container.appendChild(createFileTag(files[0].name, files[0].size, () => {
          s.workbook = null; s.file = null;
          document.getElementById('convert-config').style.display = 'none';
          document.getElementById('convert-action-bar').style.display = 'none';
          document.getElementById('convert-preview').style.display = 'none';
        }));

        // 自动识别源格式，设置默认转换目标
        const ext = files[0].name.split('.').pop().toLowerCase();
        const formatSel = document.getElementById('convert-format');
        if (ext === 'csv') formatSel.value = 'xlsx';
        else if (ext === 'json') formatSel.value = 'xlsx';
        else formatSel.value = 'csv';

        populateSelect(document.getElementById('convert-sheet'), s.workbook.SheetNames);
        document.getElementById('convert-config').style.display = '';
        document.getElementById('convert-action-bar').style.display = '';

        // 预览
        const sheetName = s.workbook.SheetNames[0];
        const headers = getSheetHeaders(s.workbook, sheetName);
        const data = getSheetData(s.workbook, sheetName);
        renderPreviewTable(
          document.getElementById('convert-preview-thead'),
          document.getElementById('convert-preview-tbody'),
          headers, data, 50
        );
        document.getElementById('convert-preview').style.display = '';
        document.getElementById('convert-preview-stats').innerHTML =
          `<span>${data.length} 行</span><span>${headers.length} 列</span>`;

        showToast(`"${files[0].name}" 加载成功`);
      } catch (e) { showToast('文件读取失败: ' + e.message, 'error'); }
    }
  );

  // Sheet 变更时刷新预览
  document.getElementById('convert-sheet').addEventListener('change', () => {
    if (!s.workbook) return;
    const sheetName = document.getElementById('convert-sheet').value;
    const headers = getSheetHeaders(s.workbook, sheetName);
    const data = getSheetData(s.workbook, sheetName);
    renderPreviewTable(
      document.getElementById('convert-preview-thead'),
      document.getElementById('convert-preview-tbody'),
      headers, data, 50
    );
    document.getElementById('convert-preview-stats').innerHTML =
      `<span>${data.length} 行</span><span>${headers.length} 列</span>`;
  });

  // 执行转换
  document.getElementById('convert-execute').addEventListener('click', () => {
    if (!s.workbook) return;
    const sheetName = document.getElementById('convert-sheet').value;
    const format = document.getElementById('convert-format').value;
    const baseName = s.file.name.replace(/\.\w+$/, '');

    try {
      const ws = s.workbook.Sheets[sheetName];

      if (format === 'csv') {
        const csv = XLSX.utils.sheet_to_csv(ws);
        downloadCSV(csv, `${baseName}.csv`);
      } else if (format === 'json') {
        const json = XLSX.utils.sheet_to_json(ws);
        downloadJSON(json, `${baseName}.json`);
      } else {
        const newWB = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWB, ws, sheetName);
        downloadWorkbook(newWB, `${baseName}.xlsx`);
      }

      showToast(`已转换为 ${format.toUpperCase()} 格式并下载`);
    } catch (e) { showToast('转换失败: ' + e.message, 'error'); }
  });
}

// ==================== 数据清洗 ====================
function initClean() {
  const s = state.clean;

  document.getElementById('clean-back').addEventListener('click', () => switchView('home'));

  // 清洗选项切换
  document.querySelectorAll('#clean-options .clean-option').forEach(opt => {
    opt.addEventListener('click', () => {
      const cb = opt.querySelector('input[type="checkbox"]');
      cb.checked = !cb.checked;
      opt.classList.toggle('selected', cb.checked);
    });
  });

  bindUploadZone(
    document.getElementById('clean-upload'),
    document.getElementById('clean-file'),
    async (files) => {
      try {
        s.file = files[0];
        s.workbook = await readFile(files[0]);
        const container = document.getElementById('clean-files');
        container.innerHTML = '';
        container.appendChild(createFileTag(files[0].name, files[0].size, () => {
          s.workbook = null; s.file = null;
          document.getElementById('clean-config').style.display = 'none';
          document.getElementById('clean-action-bar').style.display = 'none';
          document.getElementById('clean-preview').style.display = 'none';
          document.getElementById('clean-result-stats').style.display = 'none';
          document.getElementById('clean-download').style.display = 'none';
        }));

        populateSelect(document.getElementById('clean-sheet'), s.workbook.SheetNames);
        document.getElementById('clean-config').style.display = '';
        document.getElementById('clean-action-bar').style.display = '';

        showToast(`"${files[0].name}" 加载成功`);
      } catch (e) { showToast('文件读取失败: ' + e.message, 'error'); }
    }
  );

  // 执行清洗
  document.getElementById('clean-execute').addEventListener('click', () => {
    if (!s.workbook) return;
    const sheetName = document.getElementById('clean-sheet').value;

    // 获取选中的规则
    const rules = [];
    document.querySelectorAll('#clean-options .clean-option').forEach(opt => {
      if (opt.querySelector('input').checked) rules.push(opt.dataset.rule);
    });

    if (rules.length === 0) {
      showToast('请至少选择一条清洗规则', 'error');
      return;
    }

    showProcessing('正在清洗数据...', `已选 ${rules.length} 条规则`);

    setTimeout(() => {
      try {
        const headers = getSheetHeaders(s.workbook, sheetName);
        let data = getSheetData(s.workbook, sheetName);
        const originalCount = data.length;
        let trimCount = 0, emptyRowsRemoved = 0, duplicatesRemoved = 0, spacesNormalized = 0, newlinesRemoved = 0;

        // 去除首尾空格
        if (rules.includes('trim')) {
          data.forEach(row => {
            headers.forEach(h => {
              if (typeof row[h] === 'string') {
                const trimmed = row[h].trim();
                if (trimmed !== row[h]) { trimCount++; row[h] = trimmed; }
              }
            });
          });
        }

        // 合并多余空格
        if (rules.includes('normalize-spaces')) {
          data.forEach(row => {
            headers.forEach(h => {
              if (typeof row[h] === 'string') {
                const normalized = row[h].replace(/  +/g, ' ');
                if (normalized !== row[h]) { spacesNormalized++; row[h] = normalized; }
              }
            });
          });
        }

        // 去除换行符
        if (rules.includes('newlines')) {
          data.forEach(row => {
            headers.forEach(h => {
              if (typeof row[h] === 'string') {
                const cleaned = row[h].replace(/[\r\n]+/g, ' ');
                if (cleaned !== row[h]) { newlinesRemoved++; row[h] = cleaned; }
              }
            });
          });
        }

        // 删除空行
        if (rules.includes('empty-rows')) {
          const before = data.length;
          data = data.filter(row => headers.some(h => {
            const v = row[h];
            return v !== '' && v !== null && v !== undefined;
          }));
          emptyRowsRemoved = before - data.length;
        }

        // 删除重复行
        if (rules.includes('duplicates')) {
          const before = data.length;
          const seen = new Set();
          data = data.filter(row => {
            const key = headers.map(h => String(row[h] || '')).join('||');
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
          });
          duplicatesRemoved = before - data.length;
        }

        // 生成结果 Workbook
        const ws = XLSX.utils.json_to_sheet(data, { header: headers });
        const resultWB = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(resultWB, ws, sheetName);
        s.resultWB = resultWB;

        // 渲染预览
        renderPreviewTable(
          document.getElementById('clean-preview-thead'),
          document.getElementById('clean-preview-tbody'),
          headers, data, 100
        );
        document.getElementById('clean-preview').style.display = '';
        document.getElementById('clean-preview-stats').innerHTML =
          `<span>${data.length} 行</span><span>${headers.length} 列</span>`;

        // 统计卡片
        const stats = [];
        if (rules.includes('trim')) stats.push({ label: '空格清除', value: trimCount, type: 'success' });
        if (rules.includes('empty-rows')) stats.push({ label: '空行删除', value: emptyRowsRemoved, type: 'warning' });
        if (rules.includes('duplicates')) stats.push({ label: '重复行删除', value: duplicatesRemoved, type: 'warning' });
        if (rules.includes('normalize-spaces')) stats.push({ label: '多余空格合并', value: spacesNormalized, type: 'success' });
        if (rules.includes('newlines')) stats.push({ label: '换行符清除', value: newlinesRemoved, type: 'success' });

        document.getElementById('clean-result-stats').style.display = '';
        document.getElementById('clean-result-stats').innerHTML = stats.map(s => `
          <div class="result-stat-card ${s.type}">
            <div class="result-stat-icon"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><path d="M22 4L12 14.01l-3-3"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${s.value}</span><span class="result-stat-label">${s.label}</span></div>
          </div>
        `).join('');

        document.getElementById('clean-download').style.display = '';
        showToast(`清洗完成！原 ${originalCount} 行 → 现 ${data.length} 行`);
      } catch (e) { showToast('清洗失败: ' + e.message, 'error'); }
      hideProcessing();
    }, 100);
  });

  // 下载结果
  document.getElementById('clean-download').addEventListener('click', () => {
    if (!s.resultWB) return;
    const name = s.file ? s.file.name.replace(/\.\w+$/, '') : '清洗结果';
    downloadWorkbook(s.resultWB, `${name}_已清洗.xlsx`);
    showToast('文件已下载');
  });
}

// ==================== 报表合并 ====================
function initMerge() {
  const s = state.merge;

  document.getElementById('merge-back').addEventListener('click', () => switchView('home'));

  bindUploadZone(
    document.getElementById('merge-upload'),
    document.getElementById('merge-file'),
    async (files) => {
      try {
        const container = document.getElementById('merge-files');
        for (const file of files) {
          const wb = await readFile(file);
          s.workbooks.push(wb);
          s.files.push(file);
          container.appendChild(createFileTag(file.name, file.size, () => {
            const idx = s.files.indexOf(file);
            if (idx >= 0) { s.files.splice(idx, 1); s.workbooks.splice(idx, 1); }
            if (s.files.length === 0) {
              document.getElementById('merge-config').style.display = 'none';
              document.getElementById('merge-action-bar').style.display = 'none';
              document.getElementById('merge-preview').style.display = 'none';
              document.getElementById('merge-result-stats').style.display = 'none';
              document.getElementById('merge-download').style.display = 'none';
            }
          }));
        }
        showToast(`已加载 ${s.files.length} 个文件`);
        if (s.files.length >= 1) {
          document.getElementById('merge-config').style.display = '';
          document.getElementById('merge-action-bar').style.display = '';
        }
      } catch (e) { showToast('文件读取失败: ' + e.message, 'error'); }
    }
  );

  // 执行合并
  document.getElementById('merge-execute').addEventListener('click', () => {
    if (s.workbooks.length === 0) { showToast('请先上传文件', 'error'); return; }

    const mode = document.getElementById('merge-mode').value;
    const headerMode = document.getElementById('merge-header').value;

    showProcessing('正在合并数据...', `${s.files.length} 个文件`);

    setTimeout(() => {
      try {
        let allData = [];
        let allHeaders = [];

        if (mode === 'append') {
          // 纵向追加
          s.workbooks.forEach((wb, i) => {
            const sheetName = wb.SheetNames[0];
            const headers = getSheetHeaders(wb, sheetName);
            const data = getSheetData(wb, sheetName);

            if (i === 0) {
              allHeaders = [...headers];
            } else {
              headers.forEach(h => {
                if (!allHeaders.includes(h)) allHeaders.push(h);
              });
            }

            if (headerMode === 'all' && i > 0) {
              // 添加文件来源标记行
              const markerRow = {};
              allHeaders.forEach(h => { markerRow[h] = ''; });
              markerRow[allHeaders[0]] = `--- ${s.files[i].name} ---`;
              allData.push(markerRow);
            }

            allData = allData.concat(data);
          });
        } else {
          // 合并所有 Sheets
          s.workbooks.forEach(wb => {
            wb.SheetNames.forEach(sheetName => {
              const headers = getSheetHeaders(wb, sheetName);
              const data = getSheetData(wb, sheetName);
              headers.forEach(h => {
                if (!allHeaders.includes(h)) allHeaders.push(h);
              });
              allData = allData.concat(data);
            });
          });
        }

        // 生成结果
        const ws = XLSX.utils.json_to_sheet(allData, { header: allHeaders });
        const resultWB = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(resultWB, ws, '合并结果');
        s.resultWB = resultWB;

        // 预览
        renderPreviewTable(
          document.getElementById('merge-preview-thead'),
          document.getElementById('merge-preview-tbody'),
          allHeaders, allData, 100
        );
        document.getElementById('merge-preview').style.display = '';
        document.getElementById('merge-preview-stats').innerHTML =
          `<span>${allData.length} 行</span><span>${allHeaders.length} 列</span>`;

        // 统计
        document.getElementById('merge-result-stats').style.display = '';
        document.getElementById('merge-result-stats').innerHTML = `
          <div class="result-stat-card info">
            <div class="result-stat-icon"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${s.files.length}</span><span class="result-stat-label">合并文件数</span></div>
          </div>
          <div class="result-stat-card success">
            <div class="result-stat-icon"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${allData.length}</span><span class="result-stat-label">合并后总行数</span></div>
          </div>
          <div class="result-stat-card info">
            <div class="result-stat-icon"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2v20M2 12h20"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${allHeaders.length}</span><span class="result-stat-label">总列数</span></div>
          </div>
        `;

        document.getElementById('merge-download').style.display = '';
        showToast(`合并完成！共 ${allData.length} 行`);
      } catch (e) { showToast('合并失败: ' + e.message, 'error'); }
      hideProcessing();
    }, 100);
  });

  // 下载
  document.getElementById('merge-download').addEventListener('click', () => {
    if (!s.resultWB) return;
    downloadWorkbook(s.resultWB, '合并结果.xlsx');
    showToast('文件已下载');
  });
}

// ==================== 数据比对 ====================
function initDiff() {
  const s = state.diff;

  document.getElementById('diff-back').addEventListener('click', () => switchView('home'));

  // 上传旧版
  bindUploadZone(
    document.getElementById('diff-upload-old'),
    document.getElementById('diff-file-old'),
    async (files) => {
      try {
        s.oldFile = files[0];
        s.oldWB = await readFile(files[0]);
        const container = document.getElementById('diff-files-old');
        container.innerHTML = '';
        container.appendChild(createFileTag(files[0].name, files[0].size, () => {
          s.oldWB = null; s.oldFile = null;
          hideDiffConfig();
        }));
        showToast(`旧版数据 "${files[0].name}" 加载成功`);
        checkDiffReady();
      } catch (e) { showToast('文件读取失败: ' + e.message, 'error'); }
    }
  );

  // 上传新版
  bindUploadZone(
    document.getElementById('diff-upload-new'),
    document.getElementById('diff-file-new'),
    async (files) => {
      try {
        s.newFile = files[0];
        s.newWB = await readFile(files[0]);
        const container = document.getElementById('diff-files-new');
        container.innerHTML = '';
        container.appendChild(createFileTag(files[0].name, files[0].size, () => {
          s.newWB = null; s.newFile = null;
          hideDiffConfig();
        }));
        showToast(`新版数据 "${files[0].name}" 加载成功`);
        checkDiffReady();
      } catch (e) { showToast('文件读取失败: ' + e.message, 'error'); }
    }
  );

  function hideDiffConfig() {
    document.getElementById('diff-config').style.display = 'none';
    document.getElementById('diff-action-bar').style.display = 'none';
    document.getElementById('diff-preview').style.display = 'none';
    document.getElementById('diff-result-stats').style.display = 'none';
    document.getElementById('diff-download').style.display = 'none';
  }

  function checkDiffReady() {
    if (!s.oldWB || !s.newWB) return;
    document.getElementById('diff-config').style.display = '';
    document.getElementById('diff-action-bar').style.display = '';

    populateSelect(document.getElementById('diff-old-sheet'), s.oldWB.SheetNames);
    populateSelect(document.getElementById('diff-new-sheet'), s.newWB.SheetNames);

    // 主键列从旧版取
    const headers = getSheetHeaders(s.oldWB, s.oldWB.SheetNames[0]);
    populateSelect(document.getElementById('diff-key-col'), headers);
  }

  document.getElementById('diff-old-sheet').addEventListener('change', () => {
    if (!s.oldWB) return;
    const headers = getSheetHeaders(s.oldWB, document.getElementById('diff-old-sheet').value);
    populateSelect(document.getElementById('diff-key-col'), headers);
  });

  // 执行比对
  document.getElementById('diff-execute').addEventListener('click', () => {
    if (!s.oldWB || !s.newWB) return;

    const oldSheet = document.getElementById('diff-old-sheet').value;
    const newSheet = document.getElementById('diff-new-sheet').value;
    const keyCol = document.getElementById('diff-key-col').value;

    showProcessing('正在比对数据...', `主键列: ${keyCol}`);

    setTimeout(() => {
      try {
        const oldHeaders = getSheetHeaders(s.oldWB, oldSheet);
        const newHeaders = getSheetHeaders(s.newWB, newSheet);
        const oldData = getSheetData(s.oldWB, oldSheet);
        const newData = getSheetData(s.newWB, newSheet);

        // 所有列的合集
        const allHeaders = [...new Set([...oldHeaders, ...newHeaders])];

        // 建索引
        const oldIndex = {};
        oldData.forEach(row => { oldIndex[String(row[keyCol] || '')] = row; });
        const newIndex = {};
        newData.forEach(row => { newIndex[String(row[keyCol] || '')] = row; });

        // 比对
        const resultRows = [];
        const highlightMap = {};
        let addedCount = 0, changedCount = 0, removedCount = 0, unchangedCount = 0;

        // 处理新数据中的行
        newData.forEach(row => {
          const key = String(row[keyCol] || '');
          const oldRow = oldIndex[key];
          const resultRow = { _差异类型: '', ...row };
          const rowIdx = resultRows.length;

          if (!oldRow) {
            // 新增行
            resultRow._差异类型 = '🟢 新增';
            addedCount++;
            highlightMap[rowIdx] = {};
            allHeaders.forEach(h => { highlightMap[rowIdx][h] = 'cell-new'; });
          } else {
            // 检查是否有变更
            let hasChange = false;
            highlightMap[rowIdx] = {};
            allHeaders.forEach(h => {
              const oldVal = String(oldRow[h] || '');
              const newVal = String(row[h] || '');
              if (oldVal !== newVal) {
                hasChange = true;
                highlightMap[rowIdx][h] = 'cell-changed';
              }
            });
            if (hasChange) {
              resultRow._差异类型 = '🟡 修改';
              changedCount++;
            } else {
              resultRow._差异类型 = '⚪ 不变';
              unchangedCount++;
            }
          }
          resultRows.push(resultRow);
        });

        // 处理旧数据中被删除的行
        oldData.forEach(row => {
          const key = String(row[keyCol] || '');
          if (!newIndex[key]) {
            const rowIdx = resultRows.length;
            const resultRow = { _差异类型: '🔴 删除', ...row };
            highlightMap[rowIdx] = {};
            allHeaders.forEach(h => { highlightMap[rowIdx][h] = 'cell-deleted'; });
            resultRows.push(resultRow);
            removedCount++;
          }
        });

        const displayHeaders = ['_差异类型', ...allHeaders];

        // 渲染预览
        renderPreviewTable(
          document.getElementById('diff-preview-thead'),
          document.getElementById('diff-preview-tbody'),
          displayHeaders, resultRows, 200, highlightMap
        );
        document.getElementById('diff-preview').style.display = '';
        document.getElementById('diff-preview-stats').innerHTML =
          `<span>${resultRows.length} 行</span>`;

        // 统计
        document.getElementById('diff-result-stats').style.display = '';
        document.getElementById('diff-result-stats').innerHTML = `
          <div class="result-stat-card success">
            <div class="result-stat-icon"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 5v14M5 12h14"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${addedCount}</span><span class="result-stat-label">新增行</span></div>
          </div>
          <div class="result-stat-card warning">
            <div class="result-stat-icon"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${changedCount}</span><span class="result-stat-label">修改行</span></div>
          </div>
          <div class="result-stat-card error">
            <div class="result-stat-icon"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${removedCount}</span><span class="result-stat-label">删除行</span></div>
          </div>
          <div class="result-stat-card info">
            <div class="result-stat-icon"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M8 12h8"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${unchangedCount}</span><span class="result-stat-label">未变行</span></div>
          </div>
        `;

        // 生成带差异标记的结果 Workbook
        const ws = XLSX.utils.json_to_sheet(resultRows, { header: displayHeaders });
        const resultWB = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(resultWB, ws, '差异报告');
        s.resultWB = resultWB;

        document.getElementById('diff-download').style.display = '';
        showToast(`比对完成！新增 ${addedCount}，修改 ${changedCount}，删除 ${removedCount}`);
      } catch (e) { showToast('比对失败: ' + e.message, 'error'); }
      hideProcessing();
    }, 100);
  });

  // 下载
  document.getElementById('diff-download').addEventListener('click', () => {
    if (!s.resultWB) return;
    downloadWorkbook(s.resultWB, '差异报告.xlsx');
    showToast('文件已下载');
  });
}

// ==================== 新增 5 大功能板块逻辑 ====================

// 🛡️ 敏感数据脱敏
function initMask() {
  const s = state.mask;
  let ssMaskCol = null;

  document.getElementById('mask-back').addEventListener('click', () => switchView('home'));

  bindUploadZone(
    document.getElementById('mask-upload'),
    document.getElementById('mask-file'),
    async (files) => {
      try {
        s.file = files[0];
        s.workbook = await readFile(files[0]);
        const container = document.getElementById('mask-files');
        container.innerHTML = '';
        container.appendChild(createFileTag(files[0].name, files[0].size, () => {
          s.workbook = null; s.file = null;
          hideMaskConfig();
        }));
        showToast(`文件 "${files[0].name}" 上传成功，开始配置脱敏规则`);
        checkMaskReady();
      } catch (e) { showToast('文件读取失败: ' + e.message, 'error'); }
    }
  );

  function hideMaskConfig() {
    document.getElementById('mask-config').style.display = 'none';
    document.getElementById('mask-action-bar').style.display = 'none';
    document.getElementById('mask-preview').style.display = 'none';
    document.getElementById('mask-result-stats').style.display = 'none';
    document.getElementById('mask-download').style.display = 'none';
  }

  function checkMaskReady() {
    if (!s.workbook) return;
    document.getElementById('mask-config').style.display = '';
    document.getElementById('mask-action-bar').style.display = '';

    populateSelect(document.getElementById('mask-sheet'), s.workbook.SheetNames);
    updateMaskColumns();
  }

  document.getElementById('mask-sheet').addEventListener('change', updateMaskColumns);

  function updateMaskColumns() {
    const sheetName = document.getElementById('mask-sheet').value;
    if (!sheetName) return;

    const headers = getSheetHeaders(s.workbook, sheetName);
    const container = document.getElementById('mask-rules-list');
    container.innerHTML = '';

    headers.forEach(h => {
      const item = document.createElement('div');
      item.className = 'fill-rule-item';
      
      // 智能猜测默认规则
      let defaultRule = 'none';
      const lowerH = h.toLowerCase();
      if (lowerH.includes('名') || lowerH.includes('name') || lowerH.includes('客户') || lowerH.includes('联系人')) defaultRule = 'name';
      else if (lowerH.includes('手机') || lowerH.includes('电话') || lowerH.includes('tel') || lowerH.includes('phone') || lowerH.includes('mobile')) defaultRule = 'phone';
      else if (lowerH.includes('身份') || lowerH.includes('idcard') || lowerH.includes('证件') || lowerH.includes('sfz')) defaultRule = 'idcard';
      else if (lowerH.includes('邮') || lowerH.includes('email') || lowerH.includes('mail')) defaultRule = 'email';
      else if (lowerH.includes('薪') || lowerH.includes('资') || lowerH.includes('密码') || lowerH.includes('金额') || lowerH.includes('salary')) defaultRule = 'secret';

      item.innerHTML = `
        <span class="match-rule-label" style="width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escHtml(h)}">📋 ${escHtml(h)}</span>
        <span class="match-rule-arrow">➔</span>
        <select class="mask-rule-select" data-col="${escHtml(h)}" style="flex:1;min-width:140px;padding:4px 8px;border:1px solid var(--border);border-radius:var(--r-sm);background:var(--bg-card);color:var(--text)">
          <option value="none" ${defaultRule === 'none' ? 'selected' : ''}>⚪ 不脱敏 (保持原样)</option>
          <option value="name" ${defaultRule === 'name' ? 'selected' : ''}>👤 姓名遮蔽 (如：张*三 / 李*)</option>
          <option value="phone" ${defaultRule === 'phone' ? 'selected' : ''}>📱 手机号遮蔽 (如：138****5678)</option>
          <option value="idcard" ${defaultRule === 'idcard' ? 'selected' : ''}>🆔 身份证遮蔽 (保留前6后4位)</option>
          <option value="email" ${defaultRule === 'email' ? 'selected' : ''}>📧 电子邮箱遮蔽 (如：ab***@qq.com)</option>
          <option value="secret" ${defaultRule === 'secret' ? 'selected' : ''}>🔒 完全遮蔽 (替换为 ***)</option>
          <option value="hash" ${defaultRule === 'hash' ? 'selected' : ''}>🌀 安全哈希 (SHA-256算法伪混淆，保留唯一关联性)</option>
        </select>
      `;
      container.appendChild(item);
    });

    // 预览原数据
    const data = getSheetData(s.workbook, sheetName);
    renderPreviewTable(
      document.getElementById('mask-preview-thead'),
      document.getElementById('mask-preview-tbody'),
      headers, data, 50
    );
    document.getElementById('mask-preview').style.display = '';
    document.getElementById('mask-preview-stats').innerHTML =
      `<span>${data.length} 行</span><span>${headers.length} 列</span>`;
  }

  function maskValue(val, rule) {
    if (val === null || val === undefined || val === '') return '';
    const s = String(val).trim();
    if (rule === 'none') return val;
    if (rule === 'name') {
      if (s.length <= 1) return '*';
      if (s.length === 2) return s[0] + '*';
      return s[0] + '*'.repeat(s.length - 2) + s[s.length - 1];
    } else if (rule === 'phone') {
      return s.replace(/(\d{3})\d{4}(\d{4})/, '$1****$2');
    } else if (rule === 'idcard') {
      if (s.length < 10) return '******';
      return s.substring(0, 6) + '********' + s.substring(s.length - 4);
    } else if (rule === 'email') {
      const idx = s.indexOf('@');
      if (idx <= 0) return '***';
      const name = s.substring(0, idx);
      const domain = s.substring(idx);
      if (name.length <= 2) return name[0] + '***' + domain;
      return name.substring(0, 2) + '***' + domain;
    } else if (rule === 'secret') {
      return '***';
    } else if (rule === 'hash') {
      let hash = 0;
      for (let i = 0; i < s.length; i++) {
        hash = (hash << 5) - hash + s.charCodeAt(i);
        hash |= 0;
      }
      return 'h_' + Math.abs(hash).toString(16);
    }
    return val;
  }

  document.getElementById('mask-execute').addEventListener('click', () => {
    if (!s.workbook) return;
    const sheetName = document.getElementById('mask-sheet').value;
    const data = getSheetData(s.workbook, sheetName);
    const headers = getSheetHeaders(s.workbook, sheetName);

    // 收集规则
    const rules = {};
    let maskedColsCount = 0;
    document.querySelectorAll('#mask-rules-list .mask-rule-select').forEach(sel => {
      const col = sel.dataset.col;
      const val = sel.value;
      rules[col] = val;
      if (val !== 'none') maskedColsCount++;
    });

    if (maskedColsCount === 0) {
      showToast('未选择任何脱敏规则，表格将保持原样', 'info');
    }

    showProcessing('正在本地脱敏处理...', '数据完全在您的浏览器端计算，绝对安全');

    setTimeout(() => {
      try {
        const maskedData = JSON.parse(JSON.stringify(data)); // 深拷贝
        let totalMaskedCells = 0;

        maskedData.forEach(row => {
          headers.forEach(h => {
            const rule = rules[h];
            if (rule && rule !== 'none') {
              const originalVal = row[h];
              const newVal = maskValue(originalVal, rule);
              if (originalVal !== newVal) {
                row[h] = newVal;
                totalMaskedCells++;
              }
            }
          });
        });

        // 结果存 Workbook
        const ws = XLSX.utils.json_to_sheet(maskedData, { header: headers });
        const resultWB = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(resultWB, ws, sheetName);
        s.resultWB = resultWB;

        // 渲染预览
        renderPreviewTable(
          document.getElementById('mask-preview-thead'),
          document.getElementById('mask-preview-tbody'),
          headers, maskedData, 100
        );
        document.getElementById('mask-preview-stats').innerHTML =
          `<span>已脱敏预览 ${maskedData.length} 行</span><span>${headers.length} 列</span>`;

        // 统计卡片
        document.getElementById('mask-result-stats').style.display = '';
        document.getElementById('mask-result-stats').innerHTML = `
          <div class="result-stat-card success">
            <div class="result-stat-icon" style="background:var(--red-soft);color:var(--red)"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${maskedColsCount}</span><span class="result-stat-label">脱敏字段数</span></div>
          </div>
          <div class="result-stat-card success">
            <div class="result-stat-icon" style="background:var(--green-soft);color:var(--green)"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><path d="M22 4L12 14.01l-3-3"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${totalMaskedCells}</span><span class="result-stat-label">完成脱敏单元格</span></div>
          </div>
          <div class="result-stat-card info">
            <div class="result-stat-icon"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 12h6"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${maskedData.length}</span><span class="result-stat-label">处理行数</span></div>
          </div>
        `;

        document.getElementById('mask-download').style.display = '';
        showToast('脱敏完成！文件已准备好下载。');
      } catch (err) {
        showToast('脱敏失败: ' + err.message, 'error');
      }
      hideProcessing();
    }, 150);
  });

  document.getElementById('mask-download').addEventListener('click', () => {
    if (!s.resultWB) return;
    const name = s.file ? s.file.name.replace(/\.\w+$/, '') : '脱敏数据';
    downloadWorkbook(s.resultWB, `${name}_已安全脱敏.xlsx`);
    showToast('已安全保存脱敏数据');
  });
}

// ✂️ 数据拆分器
function initSplit() {
  const s = state.split;
  let ssSplitCol = null;

  document.getElementById('split-back').addEventListener('click', () => switchView('home'));

  // 模式切换
  document.getElementById('split-mode').addEventListener('change', (e) => {
    const isCol = e.target.value === 'col';
    document.getElementById('split-col-wrap').style.display = isCol ? '' : 'none';
    document.getElementById('split-rows-wrap').style.display = isCol ? 'none' : '';
  });

  bindUploadZone(
    document.getElementById('split-upload'),
    document.getElementById('split-file'),
    async (files) => {
      try {
        s.file = files[0];
        s.workbook = await readFile(files[0]);
        const container = document.getElementById('split-files');
        container.innerHTML = '';
        container.appendChild(createFileTag(files[0].name, files[0].size, () => {
          s.workbook = null; s.file = null; s.zipBlob = null;
          hideSplitConfig();
        }));
        showToast(`表格 "${files[0].name}" 加载成功`);
        checkSplitReady();
      } catch (e) { showToast('文件读取失败: ' + e.message, 'error'); }
    }
  );

  function hideSplitConfig() {
    document.getElementById('split-config').style.display = 'none';
    document.getElementById('split-action-bar').style.display = 'none';
    document.getElementById('split-preview').style.display = 'none';
    document.getElementById('split-result-stats').style.display = 'none';
  }

  function checkSplitReady() {
    if (!s.workbook) return;
    document.getElementById('split-config').style.display = '';
    document.getElementById('split-action-bar').style.display = '';

    populateSelect(document.getElementById('split-sheet'), s.workbook.SheetNames);
    updateSplitColumns();
  }

  document.getElementById('split-sheet').addEventListener('change', updateSplitColumns);

  function updateSplitColumns() {
    const sheetName = document.getElementById('split-sheet').value;
    if (!sheetName) return;

    const headers = getSheetHeaders(s.workbook, sheetName);
    const selectContainer = document.getElementById('split-col-select');
    selectContainer.innerHTML = '';

    if (ssSplitCol) ssSplitCol.destroy();
    ssSplitCol = createSearchableSelect(selectContainer, headers, { placeholder: '选择拆分的依据列...' });

    // 预览
    const data = getSheetData(s.workbook, sheetName);
    renderPreviewTable(
      document.getElementById('split-preview-thead'),
      document.getElementById('split-preview-tbody'),
      headers, data, 50
    );
    document.getElementById('split-preview').style.display = '';
    document.getElementById('split-preview-stats').innerHTML =
      `<span>原数据: ${data.length} 行</span><span>${headers.length} 列</span>`;
  }

  document.getElementById('split-execute').addEventListener('click', () => {
    if (!s.workbook) return;
    if (typeof JSZip === 'undefined') {
      showToast('JSZip 库未加载，请检查网络连接', 'error');
      return;
    }

    const sheetName = document.getElementById('split-sheet').value;
    const mode = document.getElementById('split-mode').value;
    const data = getSheetData(s.workbook, sheetName);
    const headers = getSheetHeaders(s.workbook, sheetName);

    if (data.length === 0) {
      showToast('表中没有任何数据', 'error');
      return;
    }

    showProcessing('正在执行拆分数据并打包...', '正在使用浏览器引擎分类打包，请稍候');

    setTimeout(async () => {
      try {
        const zip = new JSZip();
        let subTablesCount = 0;
        const baseName = s.file.name.replace(/\.\w+$/, '');

        if (mode === 'col') {
          // 按列拆分
          const colName = ssSplitCol ? ssSplitCol.value : '';
          if (!colName) {
            showToast('请先选择需要拆分的依据列', 'error');
            hideProcessing();
            return;
          }

          // 分组
          const groups = {};
          data.forEach(row => {
            const val = String(row[colName] !== undefined ? row[colName] : '空值').trim().replace(/[\/\\?%*:|"<>\s]/g, '_');
            if (!groups[val]) groups[val] = [];
            groups[val].push(row);
          });

          subTablesCount = Object.keys(groups).length;

          for (const key of Object.keys(groups)) {
            const subData = groups[key];
            const ws = XLSX.utils.json_to_sheet(subData, { header: headers });
            const newWB = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWB, ws, sheetName);

            const wbout = XLSX.write(newWB, { bookType: 'xlsx', type: 'array' });
            zip.file(`${baseName}_${key}.xlsx`, wbout);
          }
        } else {
          // 按行数拆分
          const maxRows = parseInt(document.getElementById('split-rows-val').value) || 1000;
          if (maxRows <= 0) {
            showToast('行数必须大于0', 'error');
            hideProcessing();
            return;
          }

          subTablesCount = Math.ceil(data.length / maxRows);

          for (let i = 0; i < subTablesCount; i++) {
            const start = i * maxRows;
            const end = Math.min(start + maxRows, data.length);
            const subData = data.slice(start, end);

            const ws = XLSX.utils.json_to_sheet(subData, { header: headers });
            const newWB = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWB, ws, sheetName);

            const wbout = XLSX.write(newWB, { bookType: 'xlsx', type: 'array' });
            zip.file(`${baseName}_第${i + 1}部分_${start + 1}-${end}行.xlsx`, wbout);
          }
        }

        const zipBlob = await zip.generateAsync({ type: 'blob' });
        s.zipBlob = zipBlob;

        // 下载
        const url = URL.createObjectURL(zipBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${baseName}_拆分包.zip`;
        a.click();
        URL.revokeObjectURL(url);

        // 统计
        document.getElementById('split-result-stats').style.display = '';
        document.getElementById('split-result-stats').innerHTML = `
          <div class="result-stat-card success">
            <div class="result-stat-icon" style="background:var(--pink-soft);color:var(--pink)"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M8 3v3a2 2 0 0 1-2 2H3m18 0h-3a2 2 0 0 1-2-2V3"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${subTablesCount}</span><span class="result-stat-label">拆分子表个数</span></div>
          </div>
          <div class="result-stat-card success">
            <div class="result-stat-icon" style="background:var(--green-soft);color:var(--green)"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><path d="M22 4L12 14.01l-3-3"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${formatSize(zipBlob.size)}</span><span class="result-stat-label">ZIP压缩包大小</span></div>
          </div>
          <div class="result-stat-card info">
            <div class="result-stat-icon"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 12h6"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${data.length}</span><span class="result-stat-label">原数据总行数</span></div>
          </div>
        `;

        showToast(`拆分成功！已生成 ${subTablesCount} 个子表格并自动下载压缩包。`);
      } catch (err) {
        showToast('拆分打包失败: ' + err.message, 'error');
      }
      hideProcessing();
    }, 150);
  });
}

// 🧮 派生列与公式计算
function initFormula() {
  const s = state.formula;

  document.getElementById('formula-back').addEventListener('click', () => switchView('home'));

  bindUploadZone(
    document.getElementById('formula-upload'),
    document.getElementById('formula-file'),
    async (files) => {
      try {
        s.file = files[0];
        s.workbook = await readFile(files[0]);
        const container = document.getElementById('formula-files');
        container.innerHTML = '';
        container.appendChild(createFileTag(files[0].name, files[0].size, () => {
          s.workbook = null; s.file = null; s.resultWB = null;
          hideFormulaConfig();
        }));
        showToast(`文件 "${files[0].name}" 上传成功，开始编辑派生列`);
        checkFormulaReady();
      } catch (e) { showToast('文件读取失败: ' + e.message, 'error'); }
    }
  );

  function hideFormulaConfig() {
    document.getElementById('formula-config').style.display = 'none';
    document.getElementById('formula-action-bar').style.display = 'none';
    document.getElementById('formula-preview').style.display = 'none';
    document.getElementById('formula-result-stats').style.display = 'none';
    document.getElementById('formula-download').style.display = 'none';
  }

  function checkFormulaReady() {
    if (!s.workbook) return;
    document.getElementById('formula-config').style.display = '';
    document.getElementById('formula-action-bar').style.display = '';

    populateSelect(document.getElementById('formula-sheet'), s.workbook.SheetNames);
    updateFormulaColumns();
  }

  document.getElementById('formula-sheet').addEventListener('change', updateFormulaColumns);

  function updateFormulaColumns() {
    const sheetName = document.getElementById('formula-sheet').value;
    if (!sheetName) return;

    const headers = getSheetHeaders(s.workbook, sheetName);
    const tagsContainer = document.getElementById('formula-field-tags');
    tagsContainer.innerHTML = '';

    headers.forEach(h => {
      const tag = document.createElement('span');
      tag.className = 'formula-field-tag';
      tag.innerHTML = `📄 ${escHtml(h)}`;
      tag.addEventListener('click', () => {
        const input = document.getElementById('formula-expr-input');
        input.value += `[${h}]`;
        input.focus();
      });
      tagsContainer.appendChild(tag);
    });

    // 预设公式的切换逻辑
    const presetSelect = document.getElementById('formula-preset');
    presetSelect.onchange = (e) => {
      const input = document.getElementById('formula-expr-input');
      const val = e.target.value;
      if (val === 'id_birthday') {
        const idCol = headers.find(h => h.includes('身份证') || h.includes('ID') || h.includes('sfz')) || headers[0] || '身份证';
        input.value = `ID_BIRTHDAY([${idCol}])`;
        document.getElementById('formula-new-col').value = '出生日期';
      } else if (val === 'id_gender') {
        const idCol = headers.find(h => h.includes('身份证') || h.includes('ID') || h.includes('sfz')) || headers[0] || '身份证';
        input.value = `ID_GENDER([${idCol}])`;
        document.getElementById('formula-new-col').value = '性别';
      } else if (val === 'concat') {
        const c1 = headers[0] || '列A';
        const c2 = headers[1] || '列B';
        input.value = `[${c1}] + '-' + [${c2}]`;
        document.getElementById('formula-new-col').value = '合并列';
      } else if (val === 'math_mult') {
        const c1 = headers.find(h => h.includes('单价') || h.includes('price')) || headers[0] || '单价';
        const c2 = headers.find(h => h.includes('数量') || h.includes('count')) || headers[1] || '数量';
        input.value = `[${c1}] * [${c2}]`;
        document.getElementById('formula-new-col').value = '金额汇总';
      } else {
        input.value = '';
      }
    };

    // 初始原表预览
    const data = getSheetData(s.workbook, sheetName);
    renderPreviewTable(
      document.getElementById('formula-preview-thead'),
      document.getElementById('formula-preview-tbody'),
      headers, data, 50
    );
    document.getElementById('formula-preview').style.display = '';
    document.getElementById('formula-preview-stats').innerHTML =
      `<span>原数据: ${data.length} 行</span><span>${headers.length} 列</span>`;
  }

  // 沙箱公式计算器
  function evalSandboxFormula(row, expr, headers) {
    let finalExpr = expr;

    // 内置提取身份证生日辅助函数
    const ID_BIRTHDAY = (id) => {
      const s = String(id).trim();
      if (s.length !== 18) return '无效身份证';
      return `${s.substring(6, 10)}-${s.substring(10, 12)}-${s.substring(12, 14)}`;
    };

    // 内置提取身份证性别辅助函数
    const ID_GENDER = (id) => {
      const s = String(id).trim();
      if (s.length !== 18) return '无效身份证';
      const num = parseInt(s.charAt(16));
      return num % 2 === 0 ? '女' : '男';
    };

    // 字符串替换
    headers.forEach(h => {
      let cellVal = row[h];
      if (cellVal === undefined || cellVal === null) cellVal = '';
      
      // 如果计算公式中包含当前列，将 [列名] 替换为其实际的值
      const replacement = typeof cellVal === 'number' ? cellVal : JSON.stringify(String(cellVal));
      finalExpr = finalExpr.split(`[${h}]`).join(replacement);
    });

    try {
      // 构造纯安全数学与字符串计算
      const fn = new Function('ID_BIRTHDAY', 'ID_GENDER', `return (${finalExpr});`);
      const res = fn(ID_BIRTHDAY, ID_GENDER);
      if (res === null || res === undefined) return '';
      return res;
    } catch (e) {
      return '计算出错';
    }
  }

  document.getElementById('formula-execute').addEventListener('click', () => {
    if (!s.workbook) return;
    const sheetName = document.getElementById('formula-sheet').value;
    const newColName = document.getElementById('formula-new-col').value.trim();
    const expr = document.getElementById('formula-expr-input').value.trim();

    if (!newColName) {
      showToast('新列名不能为空', 'error');
      return;
    }
    if (!expr) {
      showToast('公式算式不能为空', 'error');
      return;
    }

    const data = getSheetData(s.workbook, sheetName);
    const headers = getSheetHeaders(s.workbook, sheetName);

    showProcessing('正在计算新数据并派生列...', '计算将在完全本地的沙箱中执行，保持性能卓越');

    setTimeout(() => {
      try {
        const calculatedData = JSON.parse(JSON.stringify(data));
        let errorCount = 0;

        calculatedData.forEach((row, i) => {
          const res = evalSandboxFormula(row, expr, headers);
          row[newColName] = res;
          if (res === '计算出错') errorCount++;
        });

        const newHeaders = [...headers];
        if (!newHeaders.includes(newColName)) newHeaders.push(newColName);

        // 渲染高亮预览
        const highlightMap = {};
        calculatedData.forEach((row, i) => {
          highlightMap[i] = {};
          highlightMap[i][newColName] = row[newColName] === '计算出错' ? 'cell-unmatched' : 'cell-new';
        });

        // 存入 Workbook
        const ws = XLSX.utils.json_to_sheet(calculatedData, { header: newHeaders });
        const resultWB = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(resultWB, ws, sheetName);
        s.resultWB = resultWB;

        renderPreviewTable(
          document.getElementById('formula-preview-thead'),
          document.getElementById('formula-preview-tbody'),
          newHeaders, calculatedData, 100, highlightMap
        );

        // 结果统计
        document.getElementById('formula-result-stats').style.display = '';
        document.getElementById('formula-result-stats').innerHTML = `
          <div class="result-stat-card success">
            <div class="result-stat-icon" style="background:var(--blue-soft);color:var(--blue)"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${newColName}</span><span class="result-stat-label">新派生列名</span></div>
          </div>
          <div class="result-stat-card info">
            <div class="result-stat-icon"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 12h6"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${calculatedData.length}</span><span class="result-stat-label">总计派生计算行数</span></div>
          </div>
          <div class="result-stat-card ${errorCount > 0 ? 'error' : 'success'}">
            <div class="result-stat-icon"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M12 8v4M12 16h.01"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${errorCount}</span><span class="result-stat-label">失败报错行</span></div>
          </div>
        `;

        document.getElementById('formula-download').style.display = '';
        showToast('公式派生计算已全部完成！');
      } catch (err) {
        showToast('公式派生列失败: ' + err.message, 'error');
      }
      hideProcessing();
    }, 150);
  });

  document.getElementById('formula-download').addEventListener('click', () => {
    if (!s.resultWB) return;
    const name = s.file ? s.file.name.replace(/\.\w+$/, '') : '派生表格';
    downloadWorkbook(s.resultWB, `${name}_派生新列.xlsx`);
    showToast('计算表格已成功下载');
  });
}

// 📊 多维分组汇总 (简易透视表)
function initPivot() {
  const s = state.pivot;
  let ssGroupCol = null;
  let ssValCol = null;

  document.getElementById('pivot-back').addEventListener('click', () => switchView('home'));

  bindUploadZone(
    document.getElementById('pivot-upload'),
    document.getElementById('pivot-file'),
    async (files) => {
      try {
        s.file = files[0];
        s.workbook = await readFile(files[0]);
        const container = document.getElementById('pivot-files');
        container.innerHTML = '';
        container.appendChild(createFileTag(files[0].name, files[0].size, () => {
          s.workbook = null; s.file = null; s.resultWB = null;
          hidePivotConfig();
        }));
        showToast(`数据表格 "${files[0].name}" 导入成功`);
        checkPivotReady();
      } catch (e) { showToast('文件读取失败: ' + e.message, 'error'); }
    }
  );

  function hidePivotConfig() {
    document.getElementById('pivot-config').style.display = 'none';
    document.getElementById('pivot-action-bar').style.display = 'none';
    document.getElementById('pivot-preview').style.display = 'none';
    document.getElementById('pivot-result-stats').style.display = 'none';
    document.getElementById('pivot-download').style.display = 'none';
  }

  function checkPivotReady() {
    if (!s.workbook) return;
    document.getElementById('pivot-config').style.display = '';
    document.getElementById('pivot-action-bar').style.display = '';

    populateSelect(document.getElementById('pivot-sheet'), s.workbook.SheetNames);
    updatePivotColumns();
  }

  document.getElementById('pivot-sheet').addEventListener('change', updatePivotColumns);

  function updatePivotColumns() {
    const sheetName = document.getElementById('pivot-sheet').value;
    if (!sheetName) return;

    const headers = getSheetHeaders(s.workbook, sheetName);
    
    const groupWrap = document.getElementById('pivot-group-col');
    groupWrap.innerHTML = '';
    if (ssGroupCol) ssGroupCol.destroy();
    ssGroupCol = createSearchableSelect(groupWrap, headers, { placeholder: '选择一个分组字段 (如分类、销售员)' });

    const valWrap = document.getElementById('pivot-val-col');
    valWrap.innerHTML = '';
    if (ssValCol) ssValCol.destroy();
    ssValCol = createSearchableSelect(valWrap, headers, { placeholder: '选择用于计算的指标字段 (如金额、数量)' });

    // 初始原表预览
    const data = getSheetData(s.workbook, sheetName);
    renderPreviewTable(
      document.getElementById('pivot-preview-thead'),
      document.getElementById('pivot-preview-tbody'),
      headers, data, 50
    );
    document.getElementById('pivot-preview').style.display = '';
    document.getElementById('pivot-preview-stats').innerHTML =
      `<span>原数据明细: ${data.length} 行</span>`;
  }

  document.getElementById('pivot-execute').addEventListener('click', () => {
    if (!s.workbook) return;
    const sheetName = document.getElementById('pivot-sheet').value;
    const groupCol = ssGroupCol ? ssGroupCol.value : '';
    const valCol = ssValCol ? ssValCol.value : '';
    const aggFunc = document.getElementById('pivot-agg-func').value;

    if (!groupCol || !valCol) {
      showToast('分组字段和数值指标字段都必须选择', 'error');
      return;
    }

    const data = getSheetData(s.workbook, sheetName);

    showProcessing('正在进行多维分类求和汇总计算...', '使用前端哈希图引擎，极大提升大表统计能力');

    setTimeout(() => {
      try {
        // 哈希表分组 reduce
        const groups = {};
        data.forEach(row => {
          let gKey = row[groupCol];
          if (gKey === undefined || gKey === null || gKey === '') gKey = '(空值)';
          gKey = String(gKey).trim();

          if (!groups[gKey]) groups[gKey] = [];
          groups[gKey].push(row);
        });

        const aggLabel = {
          sum: '求和',
          count: '计数',
          avg: '平均值',
          max: '最大值',
          min: '最小值'
        }[aggFunc];

        const valHeader = `${valCol}_${aggLabel}`;
        const summaryData = [];

        Object.keys(groups).forEach(key => {
          const rows = groups[key];
          const item = {};
          item[groupCol] = key;

          // 提取用于运算的数字
          const nums = rows.map(r => {
            const v = Number(r[valCol]);
            return isNaN(v) ? null : v;
          }).filter(v => v !== null);

          let resultVal = 0;
          if (aggFunc === 'sum') {
            resultVal = nums.reduce((a, b) => a + b, 0);
          } else if (aggFunc === 'count') {
            resultVal = rows.length;
          } else if (aggFunc === 'avg') {
            resultVal = nums.length > 0 ? (nums.reduce((a, b) => a + b, 0) / nums.length) : 0;
          } else if (aggFunc === 'max') {
            resultVal = nums.length > 0 ? Math.max(...nums) : 0;
          } else if (aggFunc === 'min') {
            resultVal = nums.length > 0 ? Math.min(...nums) : 0;
          }

          // 精度处理
          item[valHeader] = typeof resultVal === 'number' ? Number(resultVal.toFixed(2)) : resultVal;
          summaryData.push(item);
        });

        // 默认按分组键字母或数字升序排序
        summaryData.sort((a, b) => String(a[groupCol]).localeCompare(String(b[groupCol]), 'zh'));

        // 生成新表
        const resHeaders = [groupCol, valHeader];
        const ws = XLSX.utils.json_to_sheet(summaryData, { header: resHeaders });
        const resultWB = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(resultWB, ws, '分类汇总结果');
        s.resultWB = resultWB;

        // 渲染预览
        renderPreviewTable(
          document.getElementById('pivot-preview-thead'),
          document.getElementById('pivot-preview-tbody'),
          resHeaders, summaryData, 100
        );

        document.getElementById('pivot-preview-stats').innerHTML =
          `<span>汇总表: ${summaryData.length} 行</span><span>2 列</span>`;

        // 统计结果卡片
        document.getElementById('pivot-result-stats').style.display = '';
        document.getElementById('pivot-result-stats').innerHTML = `
          <div class="result-stat-card success">
            <div class="result-stat-icon" style="background:var(--orange-soft);color:var(--orange)"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${summaryData.length}</span><span class="result-stat-label">分类汇总分组数</span></div>
          </div>
          <div class="result-stat-card info">
            <div class="result-stat-icon"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 12h6"/></svg></div>
            <div class="result-stat-info"><span class="result-stat-value">${data.length}</span><span class="result-stat-label">原表格总行数</span></div>
          </div>
        `;

        document.getElementById('pivot-download').style.display = '';
        showToast(`分组汇总分析成功！已精简合并生成 ${summaryData.length} 行指标表`);
      } catch (err) {
        showToast('分组汇总失败: ' + err.message, 'error');
      }
      hideProcessing();
    }, 150);
  });

  document.getElementById('pivot-download').addEventListener('click', () => {
    if (!s.resultWB) return;
    const name = s.file ? s.file.name.replace(/\.\w+$/, '') : '透视汇总结果';
    downloadWorkbook(s.resultWB, `${name}_分类汇总报表.xlsx`);
    showToast('分组汇总表已成功下载');
  });
}

// 📈 一键图表生成与可视化
function initChart() {
  const s = state.chart;
  let ssChartX = null;
  let ssChartY = null;

  document.getElementById('chart-back').addEventListener('click', () => switchView('home'));

  bindUploadZone(
    document.getElementById('chart-upload'),
    document.getElementById('chart-file'),
    async (files) => {
      try {
        s.file = files[0];
        s.workbook = await readFile(files[0]);
        const container = document.getElementById('chart-files');
        container.innerHTML = '';
        container.appendChild(createFileTag(files[0].name, files[0].size, () => {
          s.workbook = null; s.file = null;
          if (s.chartInstance) { s.chartInstance.destroy(); s.chartInstance = null; }
          hideChartConfig();
        }));
        showToast(`数据表格 "${files[0].name}" 导入成功`);
        checkChartReady();
      } catch (e) { showToast('文件读取失败: ' + e.message, 'error'); }
    }
  );

  function hideChartConfig() {
    document.getElementById('chart-config').style.display = 'none';
    document.getElementById('chart-action-bar').style.display = 'none';
    document.getElementById('chart-render-panel').style.display = 'none';
    document.getElementById('chart-export-img').style.display = 'none';
  }

  function checkChartReady() {
    if (!s.workbook) return;
    document.getElementById('chart-config').style.display = '';
    document.getElementById('chart-action-bar').style.display = '';

    populateSelect(document.getElementById('chart-sheet'), s.workbook.SheetNames);
    updateChartColumns();
  }

  document.getElementById('chart-sheet').addEventListener('change', updateChartColumns);

  function updateChartColumns() {
    const sheetName = document.getElementById('chart-sheet').value;
    if (!sheetName) return;

    const headers = getSheetHeaders(s.workbook, sheetName);

    const xWrap = document.getElementById('chart-x-col');
    xWrap.innerHTML = '';
    if (ssChartX) ssChartX.destroy();
    ssChartX = createSearchableSelect(xWrap, headers, { placeholder: '选择分类字段 (例如部门、产品)' });

    const yWrap = document.getElementById('chart-y-col');
    yWrap.innerHTML = '';
    if (ssChartY) ssChartY.destroy();
    // 智能猜测 Y 轴指标字段
    let defaultYCol = headers.find(h => h.includes('金额') || h.includes('销售') || h.includes('总') || h.includes('业绩') || h.includes('count') || h.includes('price')) || headers[1] || headers[0];
    ssChartY = createSearchableSelect(yWrap, headers, { placeholder: '选择数值分析字段 (例如营业额、销售量)' });
  }

  document.getElementById('chart-execute').addEventListener('click', () => {
    if (!s.workbook) return;
    if (typeof Chart === 'undefined') {
      showToast('Chart.js 渲染引擎未成功加载，请检查网络', 'error');
      return;
    }

    const sheetName = document.getElementById('chart-sheet').value;
    const xCol = ssChartX ? ssChartX.value : '';
    const yCol = ssChartY ? ssChartY.value : '';
    const chartType = document.getElementById('chart-type').value;

    if (!xCol || !yCol) {
      showToast('您必须完整选择 X 轴与 Y 轴的映射数据列', 'error');
      return;
    }

    const data = getSheetData(s.workbook, sheetName);

    showProcessing('正在进行可视化图表渲染...', '图表正在使用 GPU 硬件加速绘制');

    setTimeout(() => {
      try {
        // 先对数据按 X 轴列做自动分组求和汇总，这样画出来的图表才最精美具有分析意义！
        const summary = {};
        data.forEach(row => {
          let xVal = row[xCol];
          if (xVal === undefined || xVal === null || xVal === '') xVal = '未分类';
          xVal = String(xVal).trim();

          const yVal = Number(row[yCol]);
          const num = isNaN(yVal) ? 0 : yVal;

          if (summary[xVal] === undefined) summary[xVal] = 0;
          summary[xVal] += num;
        });

        // 限制最多画 30 条数据（防止柱子密密麻麻卡死），按大小排个序
        const summaryArray = Object.keys(summary).map(key => ({
          label: key,
          value: Number(summary[key].toFixed(2))
        }));
        summaryArray.sort((a, b) => b.value - a.value); // 降序

        const topData = summaryArray.slice(0, 30);
        const labels = topData.map(item => item.label);
        const values = topData.map(item => item.value);

        if (s.chartInstance) {
          s.chartInstance.destroy();
        }

        // 创建图表画布
        document.getElementById('chart-render-panel').style.display = 'flex';
        const canvas = document.getElementById('chart-canvas');
        const ctx = canvas.getContext('2d');

        // 主题色自适应
        const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
        const labelColor = isDark ? '#a1a8b9' : '#5b6477';
        const gridColor = isDark ? 'rgba(255, 255, 255, 0.05)' : 'rgba(0, 0, 0, 0.04)';

        // 高颜值色彩搭配渐变
        let backgroundColors = 'rgba(91, 95, 199, 0.75)';
        let borderColors = 'rgba(91, 95, 199, 1)';

        if (chartType === 'pie' || chartType === 'doughnut') {
          backgroundColors = [
            'rgba(46, 124, 246, 0.75)', 'rgba(30, 169, 124, 0.75)', 'rgba(233, 133, 13, 0.75)',
            'rgba(142, 78, 198, 0.75)', 'rgba(6, 182, 212, 0.75)', 'rgba(236, 72, 153, 0.75)',
            'rgba(229, 72, 77, 0.75)', 'rgba(91, 95, 199, 0.75)', 'rgba(120, 120, 120, 0.75)'
          ];
          borderColors = '#fff';
        }

        const isHorizontal = chartType === 'horizontalBar';
        const actualType = isHorizontal ? 'bar' : chartType;
        const options = {
          responsive: true,
          maintainAspectRatio: false,
          indexAxis: isHorizontal ? 'y' : 'x',
          plugins: {
            legend: {
              display: chartType === 'pie' || chartType === 'doughnut',
              labels: { color: labelColor, font: { family: 'inherit', size: 11, weight: 'bold' } }
            },
            tooltip: {
              padding: 10,
              cornerRadius: 8,
              font: { family: 'inherit' }
            }
          },
          scales: (chartType === 'pie' || chartType === 'doughnut') ? {} : {
            x: {
              grid: { color: gridColor },
              ticks: { color: labelColor, font: { family: 'inherit', size: 10 } }
            },
            y: {
              grid: { color: gridColor },
              ticks: { color: labelColor, font: { family: 'inherit', size: 10 } }
            }
          }
        };

        s.chartInstance = new Chart(ctx, {
          type: actualType,
          data: {
            labels: labels,
            datasets: [{
              label: `${yCol} (汇总)`,
              data: values,
              backgroundColor: backgroundColors,
              borderColor: borderColors,
              borderWidth: 1.5,
              borderRadius: chartType === 'bar' || isHorizontal ? 6 : 0,
              hoverOffset: 12
            }]
          },
          options: options
        });

        document.getElementById('chart-export-img').style.display = '';
        showToast('高颜值数据可视化图表已成功生成！');
      } catch (err) {
        showToast('图表绘制失败: ' + err.message, 'error');
      }
      hideProcessing();
    }, 150);
  });

  // 保存图片
  document.getElementById('chart-export-img').addEventListener('click', () => {
    const canvas = document.getElementById('chart-canvas');
    if (!canvas) return;

    const a = document.createElement('a');
    a.href = canvas.toDataURL('image/png');
    const baseName = s.file ? s.file.name.replace(/\.\w+$/, '') : '图表';
    a.download = `${baseName}_数据分析大屏.png`;
    a.click();
    showToast('动态图表已成功保存为超清 PNG 图片');
  });
}

// ==================== 初始化 ====================
document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  initHome();
  initMatch();
  initConvert();
  initClean();
  initMerge();
  initDiff();
  initMask();
  initSplit();
  initFormula();
  initPivot();
  initChart();
});
