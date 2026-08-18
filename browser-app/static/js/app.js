(() => {
  "use strict";

  const state = {
    files: { target: null, roster: null },
    targetPath: null,
    rosterPath: null,
    targetStem: null,
    targetSheets: [],
    rosterSheets: [],
    targetDetection: null,
    rosterDetection: null,
    maxUnlockedStep: 1,
  };

  const PASS_CLASS = {
    "統編比對成功": "tax_id",
    "名稱比對成功": "customer_name",
    "GroupName模糊比對成功": "group_name",
    "無法匹配": "unmatched",
  };

  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

  // ---------------------------------------------------------------------
  // Toast
  // ---------------------------------------------------------------------
  let toastTimer = null;
  function showToast(message, isError = false) {
    const el = $("#toast");
    el.textContent = message;
    el.classList.toggle("is-error", isError);
    el.classList.remove("hidden");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => el.classList.add("hidden"), 4200);
  }

  // ---------------------------------------------------------------------
  // 步驟切換
  // ---------------------------------------------------------------------
  function unlockStep(n) {
    state.maxUnlockedStep = Math.max(state.maxUnlockedStep, n);
    $$(".step-tab").forEach((tab) => {
      const step = Number(tab.dataset.step);
      tab.disabled = step > state.maxUnlockedStep;
    });
  }

  function showStep(n) {
    $$(".step-page").forEach((page) => page.classList.remove("is-active"));
    $$(".step-tab").forEach((tab) => tab.classList.remove("is-active"));
    $(`#step${n}`).classList.add("is-active");
    $(`.step-tab[data-step="${n}"]`).classList.add("is-active");
  }

  $$(".step-tab").forEach((tab) => {
    tab.addEventListener("click", () => {
      if (!tab.disabled) showStep(Number(tab.dataset.step));
    });
  });

  // ---------------------------------------------------------------------
  // RoughJS 手繪裝飾（僅 4 處：標頭底線／兩個上傳框／結果頁圖章）
  // ---------------------------------------------------------------------
  function sizeCanvasToParent(canvas, parent) {
    const rect = (parent || canvas.parentElement).getBoundingClientRect();
    canvas.width = Math.max(1, Math.round(rect.width));
    canvas.height = Math.max(1, Math.round(rect.height));
  }

  function drawHeaderDivider() {
    const canvas = $("#headerDivider");
    canvas.width = canvas.clientWidth;
    canvas.height = 14;
    const rc = rough.canvas(canvas);
    rc.line(2, 7, canvas.width - 2, 7, { stroke: "#2f6f62", strokeWidth: 2.5, roughness: 1.8 });
  }

  function drawUploadBoxes() {
    $$(".upload-slot").forEach((slot) => {
      const canvas = $(".rough-box", slot);
      sizeCanvasToParent(canvas, slot);
      const rc = rough.canvas(canvas);
      const hasFile = slot.classList.contains("has-file");
      rc.rectangle(4, 4, canvas.width - 8, canvas.height - 8, {
        stroke: hasFile ? "#2f6f62" : "#b8863b",
        strokeWidth: 2,
        roughness: 2.2,
        fill: hasFile ? "rgba(47,111,98,0.05)" : undefined,
        fillStyle: "solid",
      });
    });
  }

  function drawSeal() {
    const canvas = $("#sealCanvas");
    const ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    const rc = rough.canvas(canvas);
    const cx = canvas.width / 2;
    const cy = canvas.height / 2;
    rc.circle(cx, cy, canvas.width * 0.86, {
      stroke: "#b23a2e",
      strokeWidth: 3,
      roughness: 2.3,
      fill: "rgba(178,58,46,0.05)",
      fillStyle: "solid",
    });
    rc.circle(cx, cy, canvas.width * 0.66, { stroke: "#b23a2e", strokeWidth: 1.4, roughness: 2 });
    ctx.save();
    ctx.translate(cx, cy);
    ctx.rotate((-9 * Math.PI) / 180);
    ctx.fillStyle = "#b23a2e";
    ctx.textAlign = "center";
    ctx.font = "700 16px 'Noto Serif TC', serif";
    ctx.fillText("比對完成", 0, -4);
    ctx.font = "600 10px 'IBM Plex Mono', monospace";
    const today = new Date();
    const dateStr = `${today.getFullYear()}.${String(today.getMonth() + 1).padStart(2, "0")}.${String(today.getDate()).padStart(2, "0")}`;
    ctx.fillText(dateStr, 0, 14);
    ctx.restore();
  }

  window.addEventListener("resize", () => {
    if (!$("#desk").classList.contains("hidden")) {
      drawHeaderDivider();
      drawUploadBoxes();
    }
  });

  // ---------------------------------------------------------------------
  // Step 1：上傳
  // ---------------------------------------------------------------------
  function stripExtension(filename) {
    const idx = filename.lastIndexOf(".");
    return idx > 0 ? filename.slice(0, idx) : filename;
  }

  $$("input[type=file]").forEach((input) => {
    input.addEventListener("change", () => {
      const kind = input.dataset.input;
      const file = input.files[0] || null;
      state.files[kind] = file;
      const slot = input.closest(".upload-slot");
      const nameEl = $(`.upload-filename[data-filename="${kind}"]`);
      if (file) {
        nameEl.textContent = file.name;
        slot.classList.add("has-file");
      } else {
        nameEl.textContent = "尚未選擇檔案";
        slot.classList.remove("has-file");
      }
      drawUploadBoxes();
      $("#btnUpload").disabled = !(state.files.target && state.files.roster);
    });
  });

  $("#btnUpload").addEventListener("click", async () => {
    const btn = $("#btnUpload");
    btn.disabled = true;
    try {
      state.targetPath = await PyodideRunner.writeUploadedFile(state.files.target, "target.xlsx");
      state.rosterPath = await PyodideRunner.writeUploadedFile(state.files.roster, "roster.xlsx");
      state.targetStem = stripExtension(state.files.target.name);

      const data = PyodideRunner.listSheets({ target_path: state.targetPath, roster_path: state.rosterPath });
      if (data.error) throw new Error(data.error);

      state.targetSheets = data.target_sheets;
      state.rosterSheets = data.roster_sheets;

      fillSelect($("#targetSheetSelect"), state.targetSheets);
      fillSelect($("#rosterSheetSelect"), state.rosterSheets);
      $("#sheetPanel").classList.remove("hidden");
      showToast("讀取成功，請確認工作表");
    } catch (err) {
      showToast(err.message, true);
    } finally {
      btn.disabled = false;
    }
  });

  function fillSelect(selectEl, options) {
    selectEl.innerHTML = "";
    options.forEach((name) => {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      selectEl.appendChild(opt);
    });
  }

  $("#btnMappingPreview").addEventListener("click", async () => {
    const btn = $("#btnMappingPreview");
    btn.disabled = true;
    try {
      const payload = {
        target_path: state.targetPath,
        target_sheet: $("#targetSheetSelect").value,
        roster_path: state.rosterPath,
        roster_sheet: $("#rosterSheetSelect").value,
      };
      const data = PyodideRunner.mappingPreview(payload);
      if (data.error) throw new Error(data.error);

      state.targetDetection = data.target;
      state.rosterDetection = data.roster;
      renderMappingGrid($("#targetMappingGrid"), data.target, ["tax_id", "customer_name", "acc_se"], "target");
      renderMappingGrid(
        $("#rosterMappingGrid"),
        data.roster,
        ["tax_id", "customer_name", "group_name", "acc_se"],
        "roster"
      );

      unlockStep(2);
      showStep(2);
    } catch (err) {
      showToast(err.message, true);
    } finally {
      btn.disabled = false;
    }
  });

  // ---------------------------------------------------------------------
  // Step 2：欄位對應確認
  // ---------------------------------------------------------------------
  const FIELD_LABEL = {
    tax_id: "統編欄",
    customer_name: "客戶名稱欄",
    group_name: "Group Name 欄",
    acc_se: "Acc. SE 欄",
  };

  function renderMappingGrid(container, detection, fields, prefix) {
    container.innerHTML = "";
    const saved = detection.saved_mapping;
    fields.forEach((key) => {
      const card = document.createElement("div");
      card.className = "mapping-card field";

      const label = document.createElement("span");
      label.textContent = FIELD_LABEL[key];
      card.appendChild(label);

      const select = document.createElement("select");
      select.id = `${prefix}-${key}`;
      if (key !== "group_name") {
        const placeholder = document.createElement("option");
        placeholder.value = "";
        placeholder.textContent = "— 請選擇 —";
        select.appendChild(placeholder);
      } else {
        const none = document.createElement("option");
        none.value = "";
        none.textContent = "（不使用）";
        select.appendChild(none);
      }
      detection.header_options.forEach((opt) => {
        const el = document.createElement("option");
        el.value = opt.col;
        const displayText = opt.text.replace("__", " / ");
        el.textContent = `${displayText}（第 ${opt.col} 欄）`;
        select.appendChild(el);
      });

      const savedKey = key === "acc_se" ? "acc_se_col" : `${key}_col`;
      const preselect = saved ? saved[savedKey] : detection.best_guess[key];
      if (preselect != null) select.value = String(preselect);

      card.appendChild(select);
      container.appendChild(card);
    });
  }

  function readMappingForm(detection, fields, prefix) {
    const out = {
      sheet_name: detection.sheet_name,
      header_row: detection.header_row,
      data_start_row: detection.data_start_row,
      signature: detection.signature,
    };
    fields.forEach((key) => {
      const select = $(`#${prefix}-${key}`);
      const val = select.value;
      const outKey = key === "acc_se" ? "acc_se_col" : `${key}_col`;
      out[outKey] = val === "" ? null : Number(val);
    });
    return out;
  }

  $("#btnRun").addEventListener("click", async () => {
    const btn = $("#btnRun");
    btn.disabled = true;
    try {
      const targetForm = readMappingForm(state.targetDetection, ["tax_id", "customer_name", "acc_se"], "target");
      const rosterForm = readMappingForm(
        state.rosterDetection,
        ["tax_id", "customer_name", "group_name", "acc_se"],
        "roster"
      );
      const formulaColRaw = $("#formulaColInput").value.trim();

      const payload = {
        target_path: state.targetPath,
        roster_path: state.rosterPath,
        target: targetForm,
        roster: rosterForm,
        formula_col: formulaColRaw ? columnLetterToIndex(formulaColRaw) : null,
        output_dir: "/app/outputs",
        source_stem: state.targetStem,
      };

      // 讓瀏覽器先畫出按鈕停用狀態，再進入會鎖住主執行緒的 Python 運算
      await new Promise((resolve) => setTimeout(resolve, 30));
      const data = PyodideRunner.run(payload);
      if (data.error) throw new Error(data.error);

      renderResults(data);
      unlockStep(3);
      showStep(3);
      drawSeal();
      showToast("比對完成");
    } catch (err) {
      showToast(err.message, true);
    } finally {
      btn.disabled = false;
    }
  });

  function columnLetterToIndex(letters) {
    let n = 0;
    for (const ch of letters.toUpperCase()) {
      const code = ch.charCodeAt(0) - 64;
      if (code < 1 || code > 26) return null;
      n = n * 26 + code;
    }
    return n || null;
  }

  // ---------------------------------------------------------------------
  // Step 3：結果
  // ---------------------------------------------------------------------
  function renderResults(data) {
    const cards = $("#statCards");
    cards.innerHTML = "";

    addStatCard(cards, data.stats.total, "客戶資料總數", "neutral");
    Object.entries(data.stats.pass_counts).forEach(([label, count]) => {
      const isWarn = label === "無法匹配";
      addStatCard(cards, count, label, isWarn ? "warn" : "");
    });
    if (data.stats.formula_na_baseline != null) {
      const na = data.stats.pass_counts["無法匹配"] || 0;
      const improved = data.stats.formula_na_baseline - na;
      addStatCard(cards, `${improved >= 0 ? "+" : ""}${improved}`, "較現行公式改善筆數", "neutral");
    }

    const tbody = $("#previewTable tbody");
    tbody.innerHTML = "";
    data.preview_rows.forEach((row) => {
      const tr = document.createElement("tr");

      const passClass = PASS_CLASS[row.match_pass] || "unmatched";
      const anomaliesHtml = row.anomalies.map((a) => `<span class="anomaly-tag">${escapeHtml(a)}</span>`).join("");

      tr.innerHTML = `
        <td class="num">${row.row_index}</td>
        <td>${escapeHtml(row.customer_name)}</td>
        <td class="mono">${escapeHtml(row.tax_id_clean || row.tax_id_raw)}</td>
        <td>${escapeHtml(row.matched_acc_se)}</td>
        <td><span class="status-stamp pass-${passClass}">${escapeHtml(row.match_pass)}</span></td>
        <td>${anomaliesHtml}</td>
      `;
      tbody.appendChild(tr);
    });

    $("#previewHint").textContent = data.preview_truncated
      ? `僅預覽前 ${data.preview_rows.length} 列，完整結果請下載主檔查看`
      : `共 ${data.preview_rows.length} 列`;

    setDownloadLink($("#downloadPrimary"), data.downloads.primary);
    setDownloadLink($("#downloadUnmatched"), data.downloads.unmatched);
  }

  function setDownloadLink(anchorEl, vfsPath) {
    const filename = vfsPath.split("/").pop();
    anchorEl.href = PyodideRunner.downloadUrlFromVfs(vfsPath);
    anchorEl.download = filename;
  }

  function addStatCard(container, num, label, variant) {
    const card = document.createElement("div");
    card.className = `stat-card${variant === "warn" ? " is-warn" : variant === "neutral" ? " is-neutral" : ""}`;
    card.innerHTML = `<span class="num">${num}</span><span class="label">${escapeHtml(label)}</span>`;
    container.appendChild(card);
  }

  function escapeHtml(str) {
    const div = document.createElement("div");
    div.textContent = str == null ? "" : String(str);
    return div.innerHTML;
  }

  // ---------------------------------------------------------------------
  // 初始化：先跑 Pyodide 開機流程，成功後才顯示主畫面
  // ---------------------------------------------------------------------
  document.addEventListener("DOMContentLoaded", async () => {
    if (window.lucide) lucide.createIcons();

    try {
      await PyodideRunner.init((message) => {
        $("#bootHint").textContent = message;
      });
      $("#bootScreen").classList.add("hidden");
      $("#desk").classList.remove("hidden");
      if (window.lucide) lucide.createIcons();
      drawHeaderDivider();
      drawUploadBoxes();
    } catch (err) {
      $("#bootHint").textContent = `載入失敗：${err.message}（請重新整理頁面再試一次）`;
    }
  });
})();
