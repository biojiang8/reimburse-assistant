const ICONS = {
  inspecting: '../assets/icons/loader-circle.svg',
  ready: '../assets/icons/circle-check.svg',
  processing: '../assets/icons/loader-circle.svg',
  done: '../assets/icons/file-check-2.svg',
  error: '../assets/icons/triangle-alert.svg'
};

const state = {
  items: [],
  selectedInput: null,
  outputDirectory: '',
  signaturePath: '',
  engineReady: false,
  busy: false,
  progress: { completed: 0, total: 0 },
  activeTab: 'invoices',
  verifyResults: null,
  projectName: '',
  projectCode: '',
  projects: []
};

const elements = {
  engineStatus: document.querySelector('#engine-status'),
  addFilesButton: document.querySelector('#add-files-button'),
  dropSelectButton: document.querySelector('#drop-select-button'),
  dropZone: document.querySelector('#drop-zone'),
  queueList: document.querySelector('#queue-list'),
  queueSummary: document.querySelector('#queue-summary'),
  verifyButton: document.querySelector('#verify-button'),
  verifyPanel: document.querySelector('#verify-panel'),
  verifySummary: document.querySelector('#verify-summary'),
  verifyList: document.querySelector('#verify-list'),
  verifyRefreshButton: document.querySelector('#verify-refresh-button'),
  clearButton: document.querySelector('#clear-button'),
  signatureButton: document.querySelector('#signature-button'),
  signatureLabel: document.querySelector('#signature-label'),
  projectNameInput: document.querySelector('#project-name-input'),
  projectCodeInput: document.querySelector('#project-code-input'),
  projectSelect: document.querySelector('#project-select'),
  outputDirectoryButton: document.querySelector('#output-directory-button'),
  outputDirectoryLabel: document.querySelector('#output-directory-label'),
  openOutputButton: document.querySelector('#open-output-button'),
  overwriteToggle: document.querySelector('#overwrite-toggle'),
  processButton: document.querySelector('#process-button'),
  processCount: document.querySelector('#process-count'),
  previewFilename: document.querySelector('#preview-filename'),
  pdfPreview: document.querySelector('#pdf-preview'),
  previewEmpty: document.querySelector('#preview-empty'),
  revealFileButton: document.querySelector('#reveal-file-button'),
  statusDot: document.querySelector('#status-dot'),
  statusMessage: document.querySelector('#status-message'),
  progressGroup: document.querySelector('#progress-group'),
  progressFill: document.querySelector('#progress-fill'),
  progressLabel: document.querySelector('#progress-label'),
  toastRegion: document.querySelector('#toast-region'),
  tabInvoices: document.querySelector('#tab-invoices'),
  tabInvoicesButton: document.querySelector('#tab-invoices-button')
};

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function basename(value) {
  return String(value || '').split(/[\\/]/).filter(Boolean).at(-1) || '';
}

function compactPath(value) {
  const parts = String(value || '').split(/[\\/]/).filter(Boolean);
  if (parts.length <= 2) return parts.join('/');
  return `…/${parts.slice(-2).join('/')}`;
}

function readyItems() {
  return state.items.filter((item) => item.status === 'ready' || item.status === 'error');
}

function setStatus(message, type = '') {
  elements.statusMessage.textContent = message;
  elements.statusDot.className = `status-dot${type ? ` ${type}` : ''}`;
}

function showToast(message, type = 'info') {
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  const icon = type === 'error' ? ICONS.error : ICONS.done;
  toast.innerHTML = `<img src="${icon}" alt="" /><span>${escapeHtml(message)}</span>`;
  elements.toastRegion.appendChild(toast);
  window.setTimeout(() => toast.remove(), 4200);
}

function statusLabel(item) {
  if (item.status === 'inspecting') return '识别中';
  if (item.status === 'processing') return '处理中';
  if (item.status === 'done') return '已完成';
  if (item.status === 'error') return '可重试';
  return '待处理';
}

function renderQueue() {
  elements.queueList.innerHTML = state.items
    .map((item) => {
      const selected = item.input === state.selectedInput;
      const displayName = item.seller || basename(item.input) || '正在识别';
      const detail = item.error
        ? item.error
        : item.invoice_number
          ? `dzfp${item.invoice_number}`
          : basename(item.input);
      const amount = item.total ? `¥${item.total}` : '—';
      const photoCount = (item.photos || []).length;
      const verifiedBadge = verificationBadge(item);
      const evidenceChip = item.word
        ? `<button class="file-chip" type="button" data-open-file="${escapeHtml(item.word)}" title="打开实物证据 Word">
             <img src="../assets/icons/file-check-2.svg" alt="" />
             <span>实物证据</span>
           </button>`
        : '';
      const excelChip = item.excel
        ? `<button class="file-chip" type="button" data-open-file="${escapeHtml(item.excel)}" title="打开发票明细 Excel">
             <img src="../assets/icons/file-spreadsheet.svg" alt="" />
             <span>明细</span>
           </button>`
        : '';
      // 已处理但缺实物证据时才显示补生成按钮（主流程由一键处理自动完成）
      const needsEvidence =
        item.status === 'done' && photoCount > 0 && !item.word;
      const evidenceButton = needsEvidence
        ? `<button
             class="button button-small"
             type="button"
             data-gen-evidence="${escapeHtml(item.input)}"
             ${state.busy ? 'disabled' : ''}
             title="生成实物证据 Word 并更新明细 Excel"
           >生成实物</button>`
        : '';
      return `
        <div
          class="queue-row invoice-row"
          role="option"
          tabindex="0"
          aria-selected="${selected}"
          data-input="${escapeHtml(item.input)}"
        >
          <span class="status-icon ${item.status}">
            <img src="${ICONS[item.status] || ICONS.ready}" alt="" />
          </span>
          <span class="invoice-main">
            <strong title="${escapeHtml(displayName)}">${escapeHtml(displayName)}</strong>
            <span title="${escapeHtml(detail)}">${escapeHtml(detail)}</span>
          </span>
          <span class="invoice-amount">
            <strong>${escapeHtml(amount)}</strong>
            <span>${statusLabel(item)}</span>
          </span>
          <span class="photo-slot" data-photo-slot="${escapeHtml(item.input)}" title="拖入或添加实物照片">
            <img src="../assets/icons/image.svg" alt="" />
            <span>实物照片</span>
            <em>${photoCount} 张</em>
            <button class="photo-add" type="button" data-add-photos="${escapeHtml(item.input)}" title="添加实物照片">＋</button>
          </span>
          <span class="row-actions">
            ${excelChip}
            ${evidenceChip}
            ${evidenceButton}
          </span>
          <span class="verify-badge" data-verify-badge="${escapeHtml(item.input)}">${verifiedBadge}</span>
          <button
            class="icon-button row-remove"
            type="button"
            title="移除"
            aria-label="移除发票"
            data-remove="${escapeHtml(item.input)}"
            ${state.busy ? 'disabled' : ''}
          >
            <img src="../assets/icons/x.svg" alt="" />
          </button>
        </div>`;
    })
    .join('');

  elements.queueList.querySelectorAll('.queue-row').forEach((row) => {
    const activate = () => selectItem(row.dataset.input);
    row.addEventListener('click', (event) => {
      const target = event.target;
      if (target.closest('[data-remove]')) return;
      if (target.closest('[data-add-photos]')) {
        event.stopPropagation();
        addPhotosViaPicker(row.dataset.input).catch((error) => showToast(error.message, 'error'));
        return;
      }
      if (target.closest('[data-gen-evidence]')) {
        event.stopPropagation();
        generateEvidenceFor(row.dataset.input).catch((error) => showToast(error.message, 'error'));
        return;
      }
      if (target.closest('[data-open-file]')) {
        event.stopPropagation();
        const path = target.closest('[data-open-file]').dataset.openFile;
        window.invoiceApp.showItem(path).catch((error) => showToast(error.message, 'error'));
        return;
      }
      activate();
    });
    row.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        activate();
      }
    });
  });

  elements.queueList.querySelectorAll('[data-remove]').forEach((button) => {
    button.addEventListener('click', (event) => {
      event.stopPropagation();
      removeItem(button.dataset.remove);
    });
  });
}

function verificationBadge(item) {
  if (!state.verifyResults || state.verifyResults.length === 0) return '';
  const result = state.verifyResults.find((r) => r.invoice_number === item.invoice_number);
  if (!result) return '';
  const passed = Object.values(result.slots).filter((slot) => slot.ok).length;
  const total = Object.values(result.slots).length;
  if (result.ok) return `<span class="verify-ok" title="校验通过">✓ 通过</span>`;
  return `<span class="verify-fail" title="${escapeHtml(
    Object.entries(result.slots)
      .filter(([, slot]) => !slot.ok)
      .map(([, slot]) => slot.message)
      .join('；')
  )}">✕ ${passed}/${total}</span>`;
}

function renderSummary() {
  const completed = state.items.filter((item) => item.status === 'done').length;
  const failed = state.items.filter((item) => item.status === 'error').length;
  const ready = state.items.filter((item) => item.status === 'ready').length;
  const parts = [`${state.items.length} 张发票`];
  if (ready) parts.push(`${ready} 待处理`);
  if (completed) parts.push(`${completed} 已完成`);
  if (failed) parts.push(`${failed} 异常`);
  elements.queueSummary.textContent = parts.join(' · ');
  elements.clearButton.disabled = state.busy || state.items.length === 0;
  elements.processCount.textContent = String(readyItems().length);
  elements.processButton.disabled = state.busy || !state.engineReady || readyItems().length === 0;
  const processed = state.items.filter((item) => item.status === 'done').length;
  elements.verifyButton.disabled = state.busy || processed === 0;
}

function findItem(input) {
  return state.items.find((item) => item.input === input);
}

async function addPhotosViaPicker(input) {
  const item = findItem(input);
  if (!item) return;
  const selected = await window.invoiceApp.selectPhotos();
  if (!selected || selected.length === 0) return;
  addPhotosToItem(input, selected);
}

function addPhotosToItem(input, paths) {
  const item = findItem(input);
  if (!item) return;
  const images = (paths || []).filter(
    (path) => path && /\.(jpe?g|png)$/i.test(path)
  );
  if (images.length === 0) {
    showToast('拖入的文件中没有图片（支持 jpg/png）', 'error');
    return;
  }
  item.photos = [...new Set([...(item.photos || []), ...images])];
  showToast(`已添加 ${images.length} 张实物照片`);
  render();
  // 已处理过的发票：拖入照片后自动补齐实物证据 Word 并重新校验
  if (item.status === 'done' && !state.busy && item.invoice_number) {
    setStatus('自动生成实物证据 Word', 'busy');
    generateEvidenceCore(item, { verify: true, silent: true })
      .then(() => showToast('实物证据 Word 已自动生成并更新明细 Excel'))
      .catch((error) => {
        showToast(error.message, 'error');
        setStatus('自动生成失败', 'error');
        render();
      });
  }
}

async function generateEvidenceCore(item, { verify = true, silent = false } = {}) {
  if (!item.invoice_number) throw new Error('该发票尚未识别完成');
  if ((item.photos || []).length === 0) throw new Error('请先拖入实物照片');
  const response = await window.invoiceApp.generateEvidence({
    invoice: item,
    photos: item.photos,
    outputDirectory: state.outputDirectory,
    excelPath: item.excel || null,
    projectName: state.projectName || elements.projectNameInput.value.trim(),
    projectCode: state.projectCode || elements.projectCodeInput.value.trim()
  });
  item.word = response.output;
  if (response.excel) item.excel = response.excel;
  if (verify) await verifyAll();
  if (!silent) {
    showToast('实物证据 Word 已生成，明细 Excel 已更新');
    setStatus('实物证据已生成', 'ready');
  }
  return response;
}

async function generateEvidenceFor(input) {
  const item = findItem(input);
  if (!item) return;
  setStatus('正在生成实物证据 Word 并更新明细 Excel', 'busy');
  try {
    await generateEvidenceCore(item, { verify: true });
  } catch (error) {
    showToast(error.message, 'error');
    setStatus('生成失败', 'error');
  }
  render();
}

async function verifyAll() {
  const processed = state.items.filter((item) => item.status === 'done');
  if (processed.length === 0) {
    showToast('请先处理发票', 'error');
    return;
  }
  setStatus('正在进行 AI 校验', 'busy');
  try {
    const response = await window.invoiceApp.verifyPackage({
      items: processed.map((item) => ({
        invoiceNumber: item.invoice_number || '',
        pdf: item.output || null,
        excel: item.excel || null,
        word: item.word || null,
        photos: item.photos || []
      }))
    });
    state.verifyResults = response.results || [];
    renderVerifyPanel();
    const passed = state.verifyResults.filter((result) => result.ok).length;
    if (passed === state.verifyResults.length) {
      showToast(`校验通过：${passed}/${state.verifyResults.length} 张发票槽位齐全`);
      setStatus('AI 校验通过', 'ready');
    } else {
      showToast(`校验未完全通过：${passed}/${state.verifyResults.length} 张通过`, 'error');
      setStatus('AI 校验存在缺口', 'error');
    }
  } catch (error) {
    showToast(error.message, 'error');
    setStatus('校验失败', 'error');
  }
  render();
}

function renderVerifyPanel() {
  const results = state.verifyResults || [];
  if (results.length === 0) {
    elements.verifyPanel.hidden = true;
    elements.verifySummary.textContent = '尚未校验';
    return;
  }
  elements.verifyPanel.hidden = false;
  const passed = results.filter((result) => result.ok).length;
  elements.verifySummary.textContent = `${passed}/${results.length} 张发票槽位齐全`;
  const slotLabels = { pdf: 'PDF', excel: 'Excel', word: 'Word', photos: '照片' };
  elements.verifyList.innerHTML = results
    .map((result) => {
      const chips = Object.entries(result.slots)
        .map(([key, slot]) => {
          const label = slotLabels[key] || key;
          return `<span class="verify-chip ${slot.ok ? 'pass' : 'fail'}" title="${escapeHtml(slot.message)}">
            ${slot.ok ? '✓' : '✕'} ${label}
          </span>`;
        })
        .join('');
      return `
        <div class="verify-row ${result.ok ? 'pass' : 'fail'}">
          <span class="verify-invoice">dzfp${escapeHtml(result.invoice_number)}</span>
          <span class="verify-slots">${chips}</span>
        </div>`;
    })
    .join('');
}

function handleFilesDropped(payload) {
  const paths = payload.paths || [];
  if (paths.length === 0) return;
  // 坐标转 CSS 像素（Tauri 给的是物理像素），定位到实物照片槽位
  const dpr = window.devicePixelRatio || 1;
  let targetInput = null;
  if (typeof payload.x === 'number' && typeof payload.y === 'number') {
    const element = document.elementFromPoint(payload.x / dpr, payload.y / dpr);
    const slot = element?.closest('[data-photo-slot]');
    if (slot) targetInput = slot.dataset.photoSlot;
  }
  if (!targetInput) {
    // 兜底：最近 hover 的照片槽位
    const hovered = document.querySelector('.photo-slot.drag-over');
    if (hovered) targetInput = hovered.dataset.photoSlot;
  }
  if (targetInput) {
    addPhotosToItem(targetInput, paths);
    return;
  }
  const pdfs = paths.filter((path) => path.toLowerCase().endsWith('.pdf'));
  if (pdfs.length) {
    addFiles(pdfs).catch((error) => showToast(error.message, 'error'));
  } else {
    showToast('请把照片拖到某张发票的「实物照片」列，发票 PDF 拖到上方区域', 'error');
  }
}

function renderPreview() {
  const selected = state.items.find((item) => item.input === state.selectedInput);
  if (!selected) {
    elements.previewFilename.textContent = '未选择发票';
    elements.pdfPreview.classList.remove('visible');
    elements.pdfPreview.removeAttribute('src');
    elements.previewEmpty.hidden = false;
    elements.revealFileButton.disabled = true;
    return;
  }

  elements.previewFilename.textContent = selected.suggested_filename || basename(selected.input);
  const previewUrl = selected.outputPreviewUrl || selected.previewUrl;
  if (previewUrl) {
    if (window.__TAURI__ && previewUrl.startsWith('file://')) {
      // Tauri 的 WKWebView 不允许直接加载 file:// 图片，转 data URL
      const filePath = decodeURIComponent(previewUrl.replace(/^file:\/\//, ''));
      window.invoiceApp
        .readImageBase64(filePath)
        .then((dataUrl) => {
          if (elements.pdfPreview.src !== dataUrl) elements.pdfPreview.src = dataUrl;
          elements.pdfPreview.classList.add('visible');
          elements.previewEmpty.hidden = true;
        })
        .catch(() => {
          elements.pdfPreview.classList.remove('visible');
          elements.previewEmpty.hidden = false;
        });
    } else {
      if (elements.pdfPreview.src !== previewUrl) elements.pdfPreview.src = previewUrl;
      elements.pdfPreview.classList.add('visible');
      elements.previewEmpty.hidden = true;
    }
  }
  elements.revealFileButton.disabled = false;
}

function render() {
  renderQueue();
  renderSummary();
  renderPreview();
  elements.addFilesButton.disabled = state.busy;
  elements.dropSelectButton.disabled = state.busy;
  elements.signatureButton.disabled = state.busy;
  elements.outputDirectoryButton.disabled = state.busy;
  elements.overwriteToggle.disabled = state.busy;
}

function selectItem(input) {
  state.selectedInput = input;
  render();
}

function removeItem(input) {
  const index = state.items.findIndex((item) => item.input === input);
  if (index < 0) return;
  state.items.splice(index, 1);
  if (state.selectedInput === input) {
    state.selectedInput = state.items[Math.min(index, state.items.length - 1)]?.input || null;
  }
  render();
}

async function addFiles(filePaths) {
  if (!Array.isArray(filePaths) || filePaths.length === 0) return;
  const existing = new Set(state.items.map((item) => item.input));
  const unique = filePaths.filter((value) => value && !existing.has(value));
  if (unique.length === 0) {
    showToast('所选发票已在队列中');
    return;
  }

  unique.forEach((input) => {
    state.items.push({ input, status: 'inspecting' });
  });
  state.selectedInput ||= unique[0];
  setStatus(`正在识别 ${unique.length} 张发票`, 'busy');
  render();

  try {
    const results = await window.invoiceApp.inspectInvoices(unique);
    results.forEach((result) => {
      const index = state.items.findIndex((item) => item.input === result.input);
      if (index < 0) return;
      state.items[index] = {
        ...state.items[index],
        ...result,
        status: result.ok ? 'ready' : 'error'
      };
    });
    const failures = results.filter((item) => !item.ok).length;
    if (failures) {
      showToast(`${failures} 张发票未能识别`, 'error');
      setStatus('部分发票需要检查', 'error');
    } else {
      setStatus(`${results.length} 张发票已识别`, 'ready');
    }
  } catch (error) {
    unique.forEach((input) => {
      const item = state.items.find((entry) => entry.input === input);
      if (item) {
        item.status = 'error';
        item.error = error.message;
      }
    });
    showToast(error.message, 'error');
    setStatus('识别失败', 'error');
  }
  render();
}

async function chooseFiles() {
  const files = await window.invoiceApp.selectInvoices();
  await addFiles(files);
}

async function chooseSignature() {
  const selected = await window.invoiceApp.selectSignature();
  if (!selected) return;
  state.signaturePath = selected;
  elements.signatureLabel.textContent = basename(selected);
}

async function chooseOutputDirectory() {
  const selected = await window.invoiceApp.selectOutputDirectory(state.outputDirectory);
  if (!selected) return;
  state.outputDirectory = selected;
  elements.outputDirectoryLabel.textContent = compactPath(selected);
  elements.outputDirectoryLabel.title = selected;
}

async function loadProjects() {
  try {
    const data = await window.invoiceApp.loadProjects();
    state.projects = Array.isArray(data.projects) ? data.projects : [];
  } catch {
    state.projects = [];
  }
  renderProjectSelect();
}

function renderProjectSelect() {
  const select = elements.projectSelect;
  const previous = select.value;
  select.innerHTML = '';
  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = state.projects.length ? '选择项目…' : '暂无项目';
  select.appendChild(placeholder);
  state.projects.forEach((project, index) => {
    const option = document.createElement('option');
    option.value = String(index);
    option.textContent = `${project.name} · ${project.code}`;
    select.appendChild(option);
  });
  const addNew = document.createElement('option');
  addNew.value = '__new__';
  addNew.textContent = '＋ 新增项目…';
  select.appendChild(addNew);

  if (previous && select.querySelector(`option[value="${previous}"]`)) {
    select.value = previous;
  } else if (state.projects.length > 0) {
    select.value = '0'; // 默认选中最近使用的项目
  } else {
    select.value = '__new__';
  }
  syncProjectInputs();
}

function syncProjectInputs() {
  const value = elements.projectSelect.value;
  const isNew = value === '__new__' || value === '';
  elements.projectNameInput.hidden = !isNew;
  elements.projectCodeInput.hidden = !isNew;
  if (isNew) {
    if (value === '__new__' && !state.projects.length) {
      elements.projectNameInput.placeholder = '项目名称（必填，自动记忆）';
      elements.projectCodeInput.placeholder = '经费代码（必填，自动记忆）';
    }
    return;
  }
  const project = state.projects[Number(value)];
  state.projectName = project?.name || '';
  state.projectCode = project?.code || '';
  elements.projectNameInput.value = state.projectName;
  elements.projectCodeInput.value = state.projectCode;
}

function upsertProject(name, code) {
  state.projects = state.projects.filter(
    (project) => !(project.name === name && project.code === code)
  );
  state.projects.unshift({ name, code });
  state.projects = state.projects.slice(0, 50);
  window.invoiceApp.saveProjects(state.projects).catch(() => {});
  renderProjectSelect();
}

async function processQueue() {
  const targets = readyItems();
  if (targets.length === 0 || state.busy) return;

  const isNew = elements.projectSelect.value === '__new__' || elements.projectSelect.value === '';
  if (isNew) {
    state.projectName = elements.projectNameInput.value.trim();
    state.projectCode = elements.projectCodeInput.value.trim();
  } else {
    syncProjectInputs();
  }
  if (!state.projectName) {
    showToast('请先选择或填写项目名称（写入发票明细 Excel）', 'error');
    elements.projectSelect.focus();
    return;
  }
  if (!state.projectCode) {
    showToast('请先选择或填写经费代码（写入发票明细 Excel）', 'error');
    elements.projectSelect.focus();
    return;
  }
  if (!state.signaturePath) {
    showToast('请先选择电子签图片（点击上方「电子签」按钮）', 'error');
    elements.signatureButton.focus();
    return;
  }
  // 自动记忆当前项目（输入一次，以后下拉选择）
  upsertProject(state.projectName, state.projectCode);

  state.busy = true;
  state.progress = { completed: 0, total: targets.length };
  targets.forEach((item) => {
    item.status = 'ready';
    delete item.error;
  });
  elements.progressGroup.hidden = false;
  elements.progressFill.style.width = '0%';
  elements.progressLabel.textContent = `0 / ${targets.length}`;
  setStatus('正在处理发票', 'busy');
  render();

  try {
    const response = await window.invoiceApp.processInvoices({
      files: targets.map((item) => item.input),
      outputDirectory: state.outputDirectory,
      signaturePath: state.signaturePath,
      overwrite: elements.overwriteToggle.checked,
      projectName: state.projectName,
      projectCode: state.projectCode
    });
    (response.results || []).forEach((result) => {
      const item = findItem(result.input);
      if (!item) return;
      Object.assign(item, result);
      item.status = result.ok ? 'done' : 'error';
      if (result.error) item.error = result.error;
    });
    const failures = response.results.filter((item) => !item.ok).length;
    if (failures) {
      showToast(`${response.results.length - failures} 张完成，${failures} 张失败`, 'error');
      setStatus('处理完成，存在异常', 'error');
    } else {
      showToast(`${response.results.length} 张签字完成（PDF + 明细 Excel）`);
      setStatus('正在生成实物证据', 'busy');
    }
    // 对有照片的已处理发票自动生成实物证据 Word（最后统一校验一次）
    const doneWithPhotos = state.items.filter(
      (item) => item.status === 'done' && item.invoice_number && (item.photos || []).length > 0
    );
    for (const item of doneWithPhotos) {
      try {
        await generateEvidenceCore(item, { verify: false, silent: true });
      } catch (error) {
        showToast(error.message, 'error');
      }
    }
    render();
    await verifyAll();
    if (doneWithPhotos.length > 0) {
      showToast(`已自动生成 ${doneWithPhotos.length} 份实物证据 Word`);
    }
  } catch (error) {
    showToast(error.message, 'error');
    setStatus('处理失败', 'error');
  } finally {
    state.busy = false;
    window.setTimeout(() => {
      elements.progressGroup.hidden = true;
    }, 1200);
    render();
  }
}

function handleProgress(payload) {
  const item = state.items.find((entry) => entry.input === payload.input);
  if (!item) return;
  item.status = payload.status;
  if (payload.result) {
    Object.assign(item, payload.result);
    if (payload.result.error) item.error = payload.result.error;
  }

  if (payload.status === 'done' || payload.status === 'error') {
    state.progress.completed += 1;
  }
  const total = state.progress.total || payload.total || 1;
  const completed = Math.min(state.progress.completed, total);
  elements.progressFill.style.width = `${Math.round((completed / total) * 100)}%`;
  elements.progressLabel.textContent = `${completed} / ${total}`;
  render();
}

async function revealSelected() {
  const selected = state.items.find((item) => item.input === state.selectedInput);
  if (!selected) return;
  await window.invoiceApp.showItem(selected.output || selected.input);
}

function registerEvents() {
  elements.addFilesButton.addEventListener('click', chooseFiles);
  elements.dropSelectButton.addEventListener('click', chooseFiles);
  elements.signatureButton.addEventListener('click', chooseSignature);
  elements.outputDirectoryButton.addEventListener('click', chooseOutputDirectory);
  elements.projectSelect.addEventListener('change', () => {
    const value = elements.projectSelect.value;
    if (value === '__new__' || value === '') {
      state.projectName = '';
      state.projectCode = '';
      elements.projectNameInput.value = '';
      elements.projectCodeInput.value = '';
    }
    syncProjectInputs();
    if (value === '__new__') elements.projectNameInput.focus();
  });
  elements.processButton.addEventListener('click', processQueue);
  elements.clearButton.addEventListener('click', () => {
    state.items = [];
    state.selectedInput = null;
    state.verifyResults = null;
    elements.verifyPanel.hidden = true;
    elements.verifySummary.textContent = '尚未校验';
    setStatus('等待添加发票');
    render();
  });
  elements.verifyButton.addEventListener('click', () => {
    verifyAll().catch((error) => showToast(error.message, 'error'));
  });
  elements.verifyRefreshButton.addEventListener('click', () => {
    verifyAll().catch((error) => showToast(error.message, 'error'));
  });
  elements.openOutputButton.addEventListener('click', async () => {
    try {
      await window.invoiceApp.openPath(state.outputDirectory);
    } catch (error) {
      showToast(error.message, 'error');
    }
  });
  elements.revealFileButton.addEventListener('click', () => {
    revealSelected().catch((error) => showToast(error.message, 'error'));
  });

  ['dragenter', 'dragover'].forEach((name) => {
    elements.dropZone.addEventListener(name, (event) => {
      event.preventDefault();
      elements.dropZone.classList.add('drag-active');
    });
  });
  ['dragleave', 'drop'].forEach((name) => {
    elements.dropZone.addEventListener(name, (event) => {
      event.preventDefault();
      elements.dropZone.classList.remove('drag-active');
    });
  });
  elements.dropZone.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      chooseFiles();
    }
  });

  // 实物照片槽位悬停高亮（兜底定位：优先用坐标映射，其次取悬停中的槽位）
  elements.queueList.addEventListener('dragover', (event) => {
    const slot = event.target.closest('[data-photo-slot]');
    elements.queueList
      .querySelectorAll('.photo-slot.drag-over')
      .forEach((el) => el.classList.remove('drag-over'));
    if (slot) slot.classList.add('drag-over');
  });
  elements.queueList.addEventListener('dragleave', (event) => {
    const slot = event.target.closest('[data-photo-slot]');
    if (slot) slot.classList.remove('drag-over');
  });

  // 统一的文件投放处理（发票 PDF → 队列；图片 → 对应发票的实物照片槽位）
  if (window.invoiceApp.onFilesDropped) {
    window.invoiceApp.onFilesDropped(handleFilesDropped);
  } else {
    elements.dropZone.addEventListener('drop', (event) => {
      const files = window.invoiceApp.resolveDroppedFiles
        ? window.invoiceApp.resolveDroppedFiles(event)
        : Array.from(event.dataTransfer.files)
            .filter((file) => file.name.toLowerCase().endsWith('.pdf'))
            .map((file) => window.invoiceApp.getPathForFile(file))
            .filter(Boolean);
      addFiles(files).catch((error) => showToast(error.message, 'error'));
    });
  }

  window.invoiceApp.onProgress(handleProgress);
}

async function initialize() {
  registerEvents();
  try {
    await loadProjects();
  } catch {
    /* 项目台账读取失败不影响使用 */
  }
  try {
    const defaults = await window.invoiceApp.getDefaults();
    state.outputDirectory = defaults.outputDirectory;
    state.signaturePath = defaults.signaturePath;
    state.engineReady = defaults.engineReady;
    elements.outputDirectoryLabel.textContent = compactPath(defaults.outputDirectory);
    elements.outputDirectoryLabel.title = defaults.outputDirectory;
    elements.signatureLabel.textContent = basename(defaults.signaturePath) || '未选择（必选）';
    if (defaults.engineReady) {
      elements.engineStatus.textContent = `本地处理 · ${defaults.engineLabel}`;
      setStatus('等待添加发票', 'ready');
    } else {
      elements.engineStatus.textContent = '处理引擎不可用';
      setStatus(defaults.engineError || '处理引擎不可用', 'error');
      showToast(defaults.engineError || '处理引擎不可用', 'error');
    }
  } catch (error) {
    elements.engineStatus.textContent = '初始化失败';
    setStatus(error.message, 'error');
    showToast(error.message, 'error');
  }
  render();
}

initialize();
