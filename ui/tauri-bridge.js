/* Tauri 适配层：把 app.js 使用的 window.invoiceApp API 映射到 Tauri IPC。
   与 Electron 版 preload.js 保持同一接口，前端逻辑零改动。 */
(function () {
  if (!window.__TAURI__) {
    console.warn('Tauri 运行时未注入，应用可能无法工作。');
    return;
  }

  const { invoke } = window.__TAURI__.core;
  const { listen } = window.__TAURI__.event;

  // 拖拽事件回调队列（Tauri 的 drag-drop 事件是异步到达的，且带坐标）
  const dropCallbacks = [];
  let pendingDrops = [];

  if (window.__TAURI__.webview) {
    const { getCurrentWebview } = window.__TAURI__.webview;
    getCurrentWebview().onDragDropEvent((event) => {
      if (event.payload.type !== 'drop') return;
      const paths = event.payload.paths || [];
      pendingDrops = paths;
      const payload = {
        paths,
        x: event.payload.position ? event.payload.position.x : null,
        y: event.payload.position ? event.payload.position.y : null
      };
      dropCallbacks.forEach((callback) => callback(payload));
    });
  }

  window.invoiceApp = {
    getDefaults: () => invoke('app_defaults'),
    selectInvoices: () => invoke('select_files', { kind: 'invoices' }),
    selectPhotos: () => invoke('select_files', { kind: 'photos' }),
    selectSignature: async () => {
      const files = await invoke('select_files', { kind: 'signature' });
      return files[0] || null;
    },
    selectOutputDirectory: (currentPath) =>
      invoke('select_directory', { current: currentPath || null }),
    inspectInvoices: async (filePaths) => {
      const results = await invoke('inspect_invoices', { paths: filePaths });
      // Python JSON 字段 snake_case → 前端 camelCase 约定（与 Electron 版一致）。
      // preview_path 是裸路径，统一加 file:// 前缀，renderPreview 里检测到
      // file:// 后会走 read_image_base64 转 data URL（WKWebView 不能直接加载 file:// 图片）。
      return results.map((item) => ({
        ...item,
        previewUrl: item.preview_path ? `file://${item.preview_path}` : null
      }));
    },
    processInvoices: async (payload) => {
      const response = await invoke('process_invoices', { payload });
      response.results = (response.results || []).map((item) => ({
        ...item,
        outputPreviewUrl: item.preview_path ? `file://${item.preview_path}` : null
      }));
      return response;
    },
    generateEvidence: (payload) => invoke('generate_evidence', { payload }),
    verifyPackage: (payload) => invoke('verify_package', { payload }),
    loadProjects: () => invoke('load_projects'),
    saveProjects: (projects) => invoke('save_projects', { projects }),
    openPath: (value) => invoke('open_path', { value }),
    showItem: (value) => invoke('show_item', { value }),
    readImageBase64: (filePath) => invoke('read_image_base64', { path: filePath }),
    getPathForFile: () => null,
    resolveDroppedFiles: (event) => {
      // Tauri 模式：返回异步收集到的路径（仅 PDF）
      const paths = pendingDrops.filter((p) => p.toLowerCase().endsWith('.pdf'));
      pendingDrops = [];
      return paths;
    },
    onFilesDropped: (callback) => {
      dropCallbacks.push(callback);
      return () => {
        const index = dropCallbacks.indexOf(callback);
        if (index >= 0) dropCallbacks.splice(index, 1);
      };
    },
    onProgress: (callback) =>
      listen('invoice:progress', (event) => callback(event.payload))
  };
})();
