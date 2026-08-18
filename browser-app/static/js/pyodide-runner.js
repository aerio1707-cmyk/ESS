/**
 * 初始化 Pyodide、把 core/*.py 寫進虛擬檔案系統、包裝 core/browser_bridge.py 的
 * 三個函式供 app.js 呼叫。全程資料只在瀏覽器記憶體裡，不經過任何伺服器。
 */
window.PyodideRunner = (() => {
  "use strict";

  const CORE_FILES = [
    "core/__init__.py",
    "core/models.py",
    "core/excel_io.py",
    "core/field_mapping.py",
    "core/profiling.py",
    "core/sanitize.py",
    "core/matching.py",
    "core/stats.py",
    "core/excel_writer.py",
    "core/pipeline.py",
    "core/browser_bridge.py",
  ];
  const CONFIG_FILES = ["config/field_keywords.json"];

  let pyodide = null;
  let bridge = null;

  async function fetchText(relativePath) {
    const res = await fetch(`../${relativePath}`);
    if (!res.ok) throw new Error(`載入 ${relativePath} 失敗（HTTP ${res.status}）`);
    return res.text();
  }

  function mkdirSafe(FS, path) {
    try {
      FS.mkdir(path);
    } catch (err) {
      // 目錄已存在則忽略
    }
  }

  function ensureParentDirs(FS, vfsPath) {
    const parts = vfsPath.split("/").filter(Boolean);
    let current = "";
    for (let i = 0; i < parts.length - 1; i++) {
      current += "/" + parts[i];
      mkdirSafe(FS, current);
    }
  }

  async function init(onProgress) {
    onProgress?.("載入 Pyodide 執行環境…");
    pyodide = await loadPyodide();

    onProgress?.("安裝 openpyxl…");
    await pyodide.loadPackage("micropip");
    const micropip = pyodide.pyimport("micropip");
    await micropip.install("openpyxl");

    onProgress?.("載入比對邏輯程式碼…");
    const FS = pyodide.FS;
    for (const relPath of [...CORE_FILES, ...CONFIG_FILES]) {
      const text = await fetchText(relPath);
      const vfsPath = `/app/${relPath}`;
      ensureParentDirs(FS, vfsPath);
      FS.writeFile(vfsPath, text);
    }
    mkdirSafe(FS, "/app/uploads");
    mkdirSafe(FS, "/app/outputs");

    await pyodide.runPythonAsync(`
import sys
if "/app" not in sys.path:
    sys.path.insert(0, "/app")
import core.browser_bridge
`);
    bridge = pyodide.pyimport("core.browser_bridge");
  }

  async function writeUploadedFile(file, vfsFilename) {
    const buf = await file.arrayBuffer();
    const bytes = new Uint8Array(buf);
    const vfsPath = `/app/uploads/${vfsFilename}`;
    pyodide.FS.writeFile(vfsPath, bytes);
    return vfsPath;
  }

  function listSheets(payload) {
    return JSON.parse(bridge.list_sheets_json(JSON.stringify(payload)));
  }

  function mappingPreview(payload) {
    return JSON.parse(bridge.mapping_preview_json(JSON.stringify(payload)));
  }

  function run(payload) {
    return JSON.parse(bridge.run_json(JSON.stringify(payload)));
  }

  function downloadUrlFromVfs(vfsPath) {
    const bytes = pyodide.FS.readFile(vfsPath);
    const blob = new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    return URL.createObjectURL(blob);
  }

  return { init, writeUploadedFile, listSheets, mappingPreview, run, downloadUrlFromVfs };
})();
