// Excel Image Spreadsheet & SQL Studio Application
// Supports both:
// 1. Python FastAPI Backend + SQLite (when running on localhost:8000)
// 2. Client-Side SQLite in WebAssembly via sql.js + SheetJS (when hosted on GitHub Pages)

document.addEventListener("DOMContentLoaded", async () => {
  // State
  let records = [];
  let currentActiveCell = { rowId: null, colType: null, coord: "A1" };
  let pendingUpload = { rowId: null, imageType: null };
  let zoomLevel = 1.0;
  let rotationDeg = 0;
  let isClientSqlMode = false;
  let clientDb = null; // sql.js Database instance for GitHub Pages mode
  let clientObjectUrls = {}; // Cache object URLs for blob images in client mode

  // DOM Elements
  const tableBody = document.getElementById("tableBody");
  const btnAddRow = document.getElementById("btnAddRow");
  const btnAddMultipleRows = document.getElementById("btnAddMultipleRows");
  const btnBottomAddRow = document.getElementById("btnBottomAddRow");
  const btnExportExcel = document.getElementById("btnExportExcel");
  const btnSqlManager = document.getElementById("btnSqlManager");
  const searchInput = document.getElementById("searchInput");
  const themeToggle = document.getElementById("themeToggle");
  const themeIconSun = document.getElementById("themeIconSun");
  const themeIconMoon = document.getElementById("themeIconMoon");
  const globalFileInput = document.getElementById("globalFileInput");
  const activeCellAddress = document.getElementById("activeCellAddress");
  const saveIndicator = document.getElementById("saveIndicator");
  const statusModeTag = document.getElementById("statusModeTag");

  // Stats elements
  const statRowCount = document.getElementById("statRowCount");
  const statFullCount = document.getElementById("statFullCount");
  const statHalfCount = document.getElementById("statHalfCount");
  const statTotalImages = document.getElementById("statTotalImages");
  const statDbSize = document.getElementById("statDbSize");
  const statServerInfo = document.getElementById("statServerInfo");

  // Lightbox elements
  const lightboxModal = document.getElementById("lightboxModal");
  const lightboxImg = document.getElementById("lightboxImg");
  const lightboxTitle = document.getElementById("lightboxTitle");
  const lightboxBadge = document.getElementById("lightboxBadge");
  const btnZoomIn = document.getElementById("btnZoomIn");
  const btnZoomOut = document.getElementById("btnZoomOut");
  const btnRotate = document.getElementById("btnRotate");
  const btnResetZoom = document.getElementById("btnResetZoom");
  const btnDownloadImage = document.getElementById("btnDownloadImage");
  const btnCloseLightbox = document.getElementById("btnCloseLightbox");
  const zoomLevelDisplay = document.getElementById("zoomLevelDisplay");

  // SQL Modal elements
  const sqlModal = document.getElementById("sqlModal");
  const btnCloseSqlModal = document.getElementById("btnCloseSqlModal");
  const dbEngineVal = document.getElementById("dbEngineVal");
  const dbPathVal = document.getElementById("dbPathVal");
  const dbSizeVal = document.getElementById("dbSizeVal");
  const dbImagesVal = document.getElementById("dbImagesVal");
  const schemaTableBody = document.getElementById("schemaTableBody");
  const sqlQueryInput = document.getElementById("sqlQueryInput");
  const btnRunSql = document.getElementById("btnRunSql");
  const btnPresetSelect = document.getElementById("btnPresetSelect");
  const btnPresetImages = document.getElementById("btnPresetImages");
  const btnDownloadSqlDb = document.getElementById("btnDownloadSqlDb");
  const sqlResultContainer = document.getElementById("sqlResultContainer");

  // Toast container
  const toastContainer = document.getElementById("toastContainer");

  // -------------------------------------------------------------
  // Theme Management
  // -------------------------------------------------------------
  const savedTheme = localStorage.getItem("excel_app_theme") || "light";
  document.documentElement.setAttribute("data-theme", savedTheme);
  updateThemeIcons(savedTheme);

  themeToggle.addEventListener("click", () => {
    const currentTheme = document.documentElement.getAttribute("data-theme") || "light";
    const newTheme = currentTheme === "light" ? "dark" : "light";
    document.documentElement.setAttribute("data-theme", newTheme);
    localStorage.setItem("excel_app_theme", newTheme);
    updateThemeIcons(newTheme);
  });

  function updateThemeIcons(theme) {
    if (theme === "dark") {
      themeIconSun.classList.add("hidden");
      themeIconMoon.classList.remove("hidden");
    } else {
      themeIconSun.classList.remove("hidden");
      themeIconMoon.classList.add("hidden");
    }
  }

  // -------------------------------------------------------------
  // Toast Notifications
  // -------------------------------------------------------------
  function showToast(message, type = "success") {
    const toast = document.createElement("div");
    toast.className = `toast toast-${type}`;
    
    let iconSvg = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#107C41" stroke-width="2"><polyline points="20 6 9 17 4 12"></polyline></svg>`;
    if (type === "error") {
      iconSvg = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#dc2626" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>`;
    } else if (type === "info") {
      iconSvg = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>`;
    }
    
    toast.innerHTML = `${iconSvg}<span>${message}</span>`;
    toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = "0";
      toast.style.transform = "translateX(100%)";
      toast.style.transition = "all 0.2s ease";
      setTimeout(() => toast.remove(), 200);
    }, 3200);
  }

  function setSaveStatus(saving = false) {
    if (saving) {
      saveIndicator.innerHTML = `
        <span class="indicator-dot" style="background-color: var(--warning)"></span>
        <span class="indicator-text" style="color: var(--warning)">Đang lưu vào SQL...</span>
      `;
    } else {
      saveIndicator.innerHTML = `
        <span class="indicator-dot"></span>
        <span class="indicator-text">${isClientSqlMode ? "Đã lưu vào WebAssembly SQL" : "Đã đồng bộ SQL"}</span>
      `;
    }
  }

  function formatFileSize(bytes) {
    if (!bytes || bytes === 0) return "0 B";
    const k = 1024;
    const sizes = ["B", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
  }

  // -------------------------------------------------------------
  // Mode Detection: Backend API vs. Client-side SQLite (GitHub Pages)
  // -------------------------------------------------------------
  async function detectModeAndInit() {
    const isGitHub = window.location.hostname.includes("github.io") || window.location.protocol === "file:";
    if (!isGitHub) {
      try {
        const testRes = await fetch("/api/records");
        if (testRes.ok) {
          isClientSqlMode = false;
          statusModeTag.textContent = "SẴN SÀNG";
          statServerInfo.textContent = "Server: Localhost:8000";
          await fetchRecords();
          return;
        }
      } catch (e) {
        console.warn("Backend API không phản hồi, chuyển sang chế độ Client SQLite (sql.js):", e);
      }
    }

    // Initialize Client-side SQLite in WebAssembly (GitHub Pages mode)
    await initClientSql();
  }

  // -------------------------------------------------------------
  // Client-Side SQLite Implementation (sql.js for GitHub Pages)
  // -------------------------------------------------------------
  async function initClientSql() {
    isClientSqlMode = true;
    statusModeTag.textContent = "GITHUB PAGES (SQLITE)";
    statusModeTag.style.backgroundColor = "#2563eb";
    statServerInfo.textContent = "Chế độ: WebAssembly SQLite";

    try {
      showToast("Khởi tạo SQLite trong trình duyệt (GitHub Pages)...", "info");
      const SQL = await window.initSqlJs({
        locateFile: file => `https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.8.0/${file}`
      });

      // Try load from localStorage
      const savedDb = localStorage.getItem("proexcel_sqlite_db");
      if (savedDb) {
        const u8 = Uint8Array.from(atob(savedDb), c => c.charCodeAt(0));
        clientDb = new SQL.Database(u8);
      } else {
        clientDb = new SQL.Database();
      }

      // Initialize table
      clientDb.run(`
        CREATE TABLE IF NOT EXISTS records (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          row_order INTEGER DEFAULT 0,
          name TEXT NOT NULL DEFAULT '',
          full_image_data BLOB,
          full_image_mime TEXT,
          full_image_name TEXT,
          full_image_size INTEGER DEFAULT 0,
          half_image_data BLOB,
          half_image_mime TEXT,
          half_image_name TEXT,
          half_image_size INTEGER DEFAULT 0,
          created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
          updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
      `);

      // Seed if empty
      const check = clientDb.exec("SELECT COUNT(*) FROM records");
      if (check[0].values[0][0] === 0) {
        clientDb.run("INSERT INTO records (row_order, name) VALUES (1, 'Dự án Alpha (Demo trên GitHub)')");
        clientDb.run("INSERT INTO records (row_order, name) VALUES (2, 'Kiểm tra giao diện Excel')");
        clientDb.run("INSERT INTO records (row_order, name) VALUES (3, 'Báo cáo số liệu')");
        saveClientDb();
      }

      await fetchRecords();
      showToast("Đã tải cơ sở dữ liệu SQLite trong trình duyệt!");
    } catch (err) {
      console.error("Lỗi khởi tạo SQLite WebAssembly:", err);
      showToast("Không thể nạp SQLite WebAssembly", "error");
    }
  }

  function saveClientDb() {
    if (!clientDb) return;
    try {
      const data = clientDb.export();
      let binary = "";
      const len = data.byteLength;
      for (let i = 0; i < len; i++) {
        binary += String.fromCharCode(data[i]);
      }
      localStorage.setItem("proexcel_sqlite_db", btoa(binary));
    } catch (e) {
      console.warn("Không thể lưu CSDL vào localStorage:", e);
    }
  }

  // -------------------------------------------------------------
  // Data Fetching & Rendering
  // -------------------------------------------------------------
  async function fetchRecords() {
    if (isClientSqlMode) {
      loadRecordsFromClientDb();
      return;
    }

    try {
      const res = await fetch("/api/records");
      const data = await res.json();
      records = data.records || [];
      renderTable();
      updateStats();
    } catch (err) {
      console.error("Lỗi tải dữ liệu:", err);
      showToast("Không thể kết nối đến máy chủ CSDL", "error");
    }
  }

  function loadRecordsFromClientDb() {
    if (!clientDb) return;
    try {
      const res = clientDb.exec(`
        SELECT 
          id, row_order, name, 
          full_image_name, full_image_size, full_image_mime,
          (full_image_data IS NOT NULL AND LENGTH(full_image_data) > 0) AS has_full,
          half_image_name, half_image_size, half_image_mime,
          (half_image_data IS NOT NULL AND LENGTH(half_image_data) > 0) AS has_half,
          created_at, updated_at
        FROM records ORDER BY row_order ASC, id ASC
      `);

      records = [];
      if (res.length > 0) {
        const cols = res[0].columns;
        const vals = res[0].values;
        vals.forEach(row => {
          const r = {};
          cols.forEach((col, idx) => r[col] = row[idx]);

          records.push({
            id: r.id,
            row_order: r.row_order,
            name: r.name,
            full_image: {
              exists: Boolean(r.has_full),
              name: r.full_image_name,
              size: r.full_image_size || 0,
              mime: r.full_image_mime,
              url: r.has_full ? getClientImageUrl(r.id, "full") : null
            },
            half_image: {
              exists: Boolean(r.has_half),
              name: r.half_image_name,
              size: r.half_image_size || 0,
              mime: r.half_image_mime,
              url: r.has_half ? getClientImageUrl(r.id, "half") : null
            },
            created_at: r.created_at,
            updated_at: r.updated_at
          });
        });
      }

      renderTable();
      updateStats();
    } catch (e) {
      console.error("Lỗi đọc CSDL client:", e);
    }
  }

  function getClientImageUrl(recordId, imageType) {
    const key = `${recordId}_${imageType}`;
    if (clientObjectUrls[key]) return clientObjectUrls[key];

    try {
      const col = imageType === "full" ? "full_image_data" : "half_image_data";
      const mimeCol = imageType === "full" ? "full_image_mime" : "half_image_mime";
      const stmt = clientDb.prepare(`SELECT ${col}, ${mimeCol} FROM records WHERE id = :id`);
      stmt.bind({ ":id": recordId });
      if (stmt.step()) {
        const row = stmt.get();
        const blobData = row[0];
        const mime = row[1] || "image/png";
        if (blobData && blobData.length > 0) {
          const blob = new Blob([blobData], { type: mime });
          const url = URL.createObjectURL(blob);
          clientObjectUrls[key] = url;
          stmt.free();
          return url;
        }
      }
      stmt.free();
    } catch (e) {
      console.error("Lỗi trích xuất ảnh BLOB client:", e);
    }
    return null;
  }

  function renderTable() {
    const filter = (searchInput.value || "").trim().toLowerCase();
    const filtered = records.filter(r => !filter || (r.name || "").toLowerCase().includes(filter));

    tableBody.innerHTML = "";

    if (filtered.length === 0) {
      const emptyTr = document.createElement("tr");
      emptyTr.innerHTML = `
        <td colspan="5" style="text-align: center; padding: 40px; color: var(--text-muted)">
          <div style="font-size: 14px; font-weight: 600; margin-bottom: 6px;">Chưa có dữ liệu phù hợp</div>
          <div style="font-size: 12px;">Bấm nút <b>Thêm hàng</b> phía trên để tạo dòng mới</div>
        </td>
      `;
      tableBody.appendChild(emptyTr);
      return;
    }

    filtered.forEach((record, index) => {
      const tr = document.createElement("tr");
      tr.id = `row-${record.id}`;
      tr.setAttribute("data-row-id", record.id);
      tr.setAttribute("data-row-index", index + 1);

      // Cell 1: STT
      const tdIndex = document.createElement("td");
      tdIndex.className = "cell-index";
      tdIndex.textContent = index + 1;
      tdIndex.title = `ID bản ghi SQL: ${record.id}`;

      // Cell 2: Tên (Column A)
      const tdName = document.createElement("td");
      tdName.className = "cell-name";
      const nameWrapper = document.createElement("div");
      nameWrapper.className = "name-cell-wrapper";
      const textarea = document.createElement("textarea");
      textarea.className = "cell-name-input";
      textarea.value = record.name || "";
      textarea.placeholder = "Nhập tên / tiêu đề...";
      textarea.rows = 2;

      textarea.addEventListener("focus", () => {
        setActiveCell(record.id, "name", `A${index + 1}`);
        tr.classList.add("row-selected");
      });

      textarea.addEventListener("blur", () => {
        tr.classList.remove("row-selected");
        if (textarea.value !== record.name) {
          updateRecordName(record.id, textarea.value);
        }
      });

      textarea.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) {
          e.preventDefault();
          textarea.blur();
          addNewRow();
        }
      });

      nameWrapper.appendChild(textarea);
      tdName.appendChild(nameWrapper);

      // Cell 3: Ảnh chụp full (Column B)
      const tdFull = document.createElement("td");
      tdFull.className = "cell-image";
      const fullDropzone = createImageDropzone(record, "full", index + 1);
      tdFull.appendChild(fullDropzone);

      // Cell 4: Ảnh chụp 1 nửa (Column C)
      const tdHalf = document.createElement("td");
      tdHalf.className = "cell-image";
      const halfDropzone = createImageDropzone(record, "half", index + 1);
      tdHalf.appendChild(halfDropzone);

      // Cell 5: Actions
      const tdActions = document.createElement("td");
      tdActions.className = "cell-actions";
      const btnDelete = document.createElement("button");
      btnDelete.className = "btn-delete-row";
      btnDelete.title = "Xóa hàng này khỏi SQL";
      btnDelete.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>`;
      btnDelete.addEventListener("click", () => deleteRow(record.id));
      tdActions.appendChild(btnDelete);

      tr.appendChild(tdIndex);
      tr.appendChild(tdName);
      tr.appendChild(tdFull);
      tr.appendChild(tdHalf);
      tr.appendChild(tdActions);

      tableBody.appendChild(tr);
    });
  }

  // -------------------------------------------------------------
  // Image Dropzone Component
  // -------------------------------------------------------------
  function createImageDropzone(record, imageType, rowIndex) {
    const colLetter = imageType === "full" ? "B" : "C";
    const cellCoord = `${colLetter}${rowIndex}`;
    const imgData = record[`${imageType}_image`];

    const dropzone = document.createElement("div");
    dropzone.className = "cell-image-dropzone";
    dropzone.tabIndex = 0;
    dropzone.setAttribute("data-row-id", record.id);
    dropzone.setAttribute("data-image-type", imageType);
    dropzone.setAttribute("data-coord", cellCoord);

    dropzone.addEventListener("focus", () => {
      document.querySelectorAll(".cell-image-dropzone").forEach(el => el.classList.remove("cell-focused"));
      dropzone.classList.add("cell-focused");
      setActiveCell(record.id, imageType, cellCoord);
    });

    dropzone.addEventListener("dragover", (e) => {
      e.preventDefault();
      dropzone.classList.add("drag-over");
    });
    dropzone.addEventListener("dragleave", () => {
      dropzone.classList.remove("drag-over");
    });
    dropzone.addEventListener("drop", async (e) => {
      e.preventDefault();
      dropzone.classList.remove("drag-over");
      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        const file = e.dataTransfer.files[0];
        if (file.type.startsWith("image/")) {
          await uploadImageFile(record.id, imageType, file);
        } else {
          showToast("Vui lòng chọn tệp hình ảnh hợp lệ", "error");
        }
      }
    });

    if (imgData && imgData.exists && imgData.url) {
      const card = document.createElement("div");
      card.className = "image-preview-card";

      const thumbWrap = document.createElement("div");
      thumbWrap.className = "img-thumb-wrap";
      const img = document.createElement("img");
      img.className = "img-thumb";
      img.src = imgData.url;
      img.alt = imgData.name || "Ảnh";
      img.loading = "lazy";
      thumbWrap.appendChild(img);

      const meta = document.createElement("div");
      meta.className = "img-meta-info";
      const title = document.createElement("div");
      title.className = "img-file-name";
      title.textContent = imgData.name || (imageType === "full" ? "Ảnh chụp Full" : "Ảnh chụp 1 nửa");
      title.title = imgData.name || "";

      const badges = document.createElement("div");
      badges.className = "img-meta-badges";
      const sizeBadge = document.createElement("span");
      sizeBadge.className = "img-size-badge";
      sizeBadge.textContent = formatFileSize(imgData.size);

      const typeBadge = document.createElement("span");
      typeBadge.className = "img-type-badge";
      typeBadge.textContent = imageType === "full" ? "FULL" : "1/2";

      badges.appendChild(sizeBadge);
      badges.appendChild(typeBadge);
      meta.appendChild(title);
      meta.appendChild(badges);

      const overlay = document.createElement("div");
      overlay.className = "img-actions-overlay";

      const btnZoom = document.createElement("button");
      btnZoom.className = "btn-cell-act";
      btnZoom.title = "Phóng to / Xem chi tiết";
      btnZoom.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line><line x1="11" y1="8" x2="11" y2="14"></line><line x1="8" y1="11" x2="14" y2="11"></line></svg>`;
      btnZoom.addEventListener("click", (e) => {
        e.stopPropagation();
        openLightbox(imgData.url, record.name || title.textContent, imageType);
      });

      const btnReplace = document.createElement("button");
      btnReplace.className = "btn-cell-act";
      btnReplace.title = "Thay thế ảnh khác";
      btnReplace.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"/></svg>`;
      btnReplace.addEventListener("click", (e) => {
        e.stopPropagation();
        triggerFilePicker(record.id, imageType);
      });

      const btnDownload = document.createElement("a");
      btnDownload.className = "btn-cell-act";
      btnDownload.title = "Tải ảnh về máy";
      btnDownload.href = imgData.url;
      btnDownload.download = imgData.name || `${imageType}_${record.id}.png`;
      btnDownload.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>`;
      btnDownload.addEventListener("click", (e) => e.stopPropagation());

      const btnDeleteImg = document.createElement("button");
      btnDeleteImg.className = "btn-cell-act btn-delete-img";
      btnDeleteImg.title = "Xóa ảnh này";
      btnDeleteImg.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"></path></svg>`;
      btnDeleteImg.addEventListener("click", (e) => {
        e.stopPropagation();
        deleteImage(record.id, imageType);
      });

      overlay.appendChild(btnZoom);
      overlay.appendChild(btnReplace);
      overlay.appendChild(btnDownload);
      overlay.appendChild(btnDeleteImg);

      card.appendChild(thumbWrap);
      card.appendChild(meta);
      card.appendChild(overlay);

      card.addEventListener("click", () => {
        openLightbox(imgData.url, record.name || title.textContent, imageType);
      });

      dropzone.appendChild(card);
    } else {
      const emptyWrap = document.createElement("div");
      emptyWrap.className = "dropzone-empty";
      emptyWrap.innerHTML = `
        <div class="dropzone-icon">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8">
            <rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect>
            <circle cx="8.5" cy="8.5" r="1.5"></circle>
            <polyline points="21 15 16 10 5 21"></polyline>
          </svg>
        </div>
        <div class="dropzone-title">Chọn ảnh / Kéo thả vào đây</div>
        <div class="dropzone-shortcuts">
          <span class="kbd-badge">Ctrl + V</span>
          <span class="dropzone-sub">để dán ảnh chụp</span>
        </div>
      `;

      emptyWrap.addEventListener("click", () => {
        triggerFilePicker(record.id, imageType);
      });

      dropzone.appendChild(emptyWrap);
    }

    return dropzone;
  }

  function setActiveCell(rowId, colType, coord) {
    currentActiveCell = { rowId, colType, coord };
    activeCellAddress.textContent = coord;
  }

  function triggerFilePicker(rowId, imageType) {
    pendingUpload = { rowId, imageType };
    globalFileInput.value = "";
    globalFileInput.click();
  }

  globalFileInput.addEventListener("change", async (e) => {
    if (e.target.files && e.target.files.length > 0 && pendingUpload.rowId) {
      const file = e.target.files[0];
      await uploadImageFile(pendingUpload.rowId, pendingUpload.imageType, file);
      pendingUpload = { rowId: null, imageType: null };
    }
  });

  // -------------------------------------------------------------
  // Clipboard Paste Handler (Ctrl + V)
  // -------------------------------------------------------------
  document.addEventListener("paste", async (e) => {
    const focusedDropzone = document.querySelector(".cell-image-dropzone:focus, .cell-image-dropzone.cell-focused");
    let targetRowId = null;
    let targetImageType = null;

    if (focusedDropzone) {
      targetRowId = parseInt(focusedDropzone.getAttribute("data-row-id"));
      targetImageType = focusedDropzone.getAttribute("data-image-type");
    } else if (currentActiveCell.rowId && (currentActiveCell.colType === "full" || currentActiveCell.colType === "half")) {
      targetRowId = currentActiveCell.rowId;
      targetImageType = currentActiveCell.colType;
    }

    if (!targetRowId || !targetImageType) return;

    const items = (e.clipboardData || e.originalEvent.clipboardData).items;
    for (let i = 0; i < items.length; i++) {
      if (items[i].type.indexOf("image") !== -1) {
        e.preventDefault();
        const blob = items[i].getAsFile();
        const file = new File([blob], `pasted_${targetImageType}_${Date.now()}.png`, { type: blob.type });
        await uploadImageFile(targetRowId, targetImageType, file);
        break;
      }
    }
  });

  // -------------------------------------------------------------
  // CRUD Actions
  // -------------------------------------------------------------
  async function addNewRow(name = "") {
    setSaveStatus(true);
    if (isClientSqlMode) {
      try {
        const max = clientDb.exec("SELECT MAX(row_order) FROM records");
        const nextOrder = (max[0].values[0][0] || 0) + 1;
        clientDb.run("INSERT INTO records (row_order, name) VALUES (?, ?)", [nextOrder, name]);
        saveClientDb();
        showToast("Đã thêm 1 hàng mới vào SQLite");
        fetchRecords();
      } finally {
        setSaveStatus(false);
      }
      return;
    }

    try {
      const res = await fetch("/api/records", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name })
      });
      const data = await res.json();
      if (data.success) {
        showToast("Đã thêm 1 hàng mới vào CSDL SQL");
        await fetchRecords();
        setTimeout(() => {
          const newRow = document.getElementById(`row-${data.record.id}`);
          if (newRow) {
            const input = newRow.querySelector(".cell-name-input");
            if (input) {
              input.focus();
              newRow.scrollIntoView({ behavior: "smooth", block: "nearest" });
            }
          }
        }, 50);
      }
    } catch (err) {
      console.error("Lỗi thêm hàng:", err);
      showToast("Không thể thêm hàng", "error");
    } finally {
      setSaveStatus(false);
    }
  }

  async function addMultipleRows(count = 5) {
    setSaveStatus(true);
    if (isClientSqlMode) {
      for (let i = 0; i < count; i++) {
        const max = clientDb.exec("SELECT MAX(row_order) FROM records");
        const nextOrder = (max[0].values[0][0] || 0) + 1;
        clientDb.run("INSERT INTO records (row_order, name) VALUES (?, '')", [nextOrder]);
      }
      saveClientDb();
      showToast(`Đã thêm nhanh ${count} hàng mới vào SQLite`);
      fetchRecords();
      setSaveStatus(false);
      return;
    }

    try {
      for (let i = 0; i < count; i++) {
        await fetch("/api/records", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ name: "" })
        });
      }
      showToast(`Đã thêm nhanh ${count} hàng mới vào SQL`);
      await fetchRecords();
    } catch (err) {
      showToast("Lỗi khi thêm nhiều hàng", "error");
    } finally {
      setSaveStatus(false);
    }
  }

  async function updateRecordName(recordId, newName) {
    setSaveStatus(true);
    if (isClientSqlMode) {
      clientDb.run("UPDATE records SET name = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", [newName, recordId]);
      saveClientDb();
      const item = records.find(r => r.id === recordId);
      if (item) item.name = newName;
      setSaveStatus(false);
      return;
    }

    try {
      const res = await fetch(`/api/records/${recordId}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: newName })
      });
      const data = await res.json();
      if (data.success) {
        const item = records.find(r => r.id === recordId);
        if (item) item.name = newName;
      }
    } catch (err) {
      console.error("Lỗi cập nhật tên:", err);
      showToast("Không thể lưu tên hàng", "error");
    } finally {
      setSaveStatus(false);
    }
  }

  async function deleteRow(recordId) {
    if (!confirm("Bạn có chắc chắn muốn xóa hàng này và các ảnh liên quan khỏi SQL?")) return;
    setSaveStatus(true);
    if (isClientSqlMode) {
      clientDb.run("DELETE FROM records WHERE id = ?", [recordId]);
      delete clientObjectUrls[`${recordId}_full`];
      delete clientObjectUrls[`${recordId}_half`];
      saveClientDb();
      showToast("Đã xóa hàng khỏi SQLite");
      fetchRecords();
      setSaveStatus(false);
      return;
    }

    try {
      const res = await fetch(`/api/records/${recordId}`, { method: "DELETE" });
      const data = await res.json();
      if (data.success) {
        showToast("Đã xóa hàng khỏi CSDL SQL");
        await fetchRecords();
      }
    } catch (err) {
      console.error("Lỗi xóa hàng:", err);
      showToast("Không thể xóa hàng", "error");
    } finally {
      setSaveStatus(false);
    }
  }

  async function uploadImageFile(recordId, imageType, file) {
    setSaveStatus(true);

    if (isClientSqlMode) {
      try {
        const arrayBuffer = await file.arrayBuffer();
        const u8 = new Uint8Array(arrayBuffer);
        const col = imageType === "full" ? "full_image_data" : "half_image_data";
        const mimeCol = imageType === "full" ? "full_image_mime" : "half_image_mime";
        const nameCol = imageType === "full" ? "full_image_name" : "half_image_name";
        const sizeCol = imageType === "full" ? "full_image_size" : "half_image_size";

        clientDb.run(
          `UPDATE records SET ${col} = ?, ${mimeCol} = ?, ${nameCol} = ?, ${sizeCol} = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?`,
          [u8, file.type || "image/png", file.name, u8.length, recordId]
        );
        delete clientObjectUrls[`${recordId}_${imageType}`];
        saveClientDb();
        showToast(`Đã lưu ảnh ${imageType === "full" ? "Full" : "1 Nửa"} vào SQLite BLOB!`);
        fetchRecords();
      } catch (err) {
        console.error("Lỗi lưu ảnh client:", err);
        showToast("Lỗi xử lý ảnh", "error");
      } finally {
        setSaveStatus(false);
      }
      return;
    }

    const formData = new FormData();
    formData.append("file", file);
    try {
      const res = await fetch(`/api/records/${recordId}/upload-image/${imageType}`, {
        method: "POST",
        body: formData
      });
      const data = await res.json();
      if (data.success) {
        showToast(`Đã lưu ảnh ${imageType === "full" ? "Full" : "1 Nửa"} vào SQL BLOB!`);
        await fetchRecords();
      } else {
        showToast(data.detail || "Không thể tải ảnh", "error");
      }
    } catch (err) {
      console.error("Lỗi upload ảnh:", err);
      showToast("Lỗi kết nối khi tải ảnh", "error");
    } finally {
      setSaveStatus(false);
    }
  }

  async function deleteImage(recordId, imageType) {
    if (!confirm("Bạn có chắc muốn xóa ảnh này khỏi CSDL?")) return;
    setSaveStatus(true);

    if (isClientSqlMode) {
      const col = imageType === "full" ? "full_image_data" : "half_image_data";
      const mimeCol = imageType === "full" ? "full_image_mime" : "half_image_mime";
      const nameCol = imageType === "full" ? "full_image_name" : "half_image_name";
      const sizeCol = imageType === "full" ? "full_image_size" : "half_image_size";

      clientDb.run(
        `UPDATE records SET ${col} = NULL, ${mimeCol} = NULL, ${nameCol} = NULL, ${sizeCol} = 0, updated_at = CURRENT_TIMESTAMP WHERE id = ?`,
        [recordId]
      );
      delete clientObjectUrls[`${recordId}_${imageType}`];
      saveClientDb();
      showToast("Đã xóa ảnh khỏi SQLite");
      fetchRecords();
      setSaveStatus(false);
      return;
    }

    try {
      const res = await fetch(`/api/records/${recordId}/image/${imageType}`, {
        method: "DELETE"
      });
      const data = await res.json();
      if (data.success) {
        showToast("Đã xóa ảnh khỏi SQL");
        await fetchRecords();
      }
    } catch (err) {
      console.error("Lỗi xóa ảnh:", err);
      showToast("Không thể xóa ảnh", "error");
    } finally {
      setSaveStatus(false);
    }
  }

  // -------------------------------------------------------------
  // Stats & Status Bar
  // -------------------------------------------------------------
  async function updateStats() {
    if (isClientSqlMode) {
      try {
        const rowsCount = clientDb.exec("SELECT COUNT(*) FROM records")[0].values[0][0];
        const fullCount = clientDb.exec("SELECT COUNT(*) FROM records WHERE full_image_data IS NOT NULL")[0].values[0][0];
        const halfCount = clientDb.exec("SELECT COUNT(*) FROM records WHERE half_image_data IS NOT NULL")[0].values[0][0];
        const totalImg = fullCount + halfCount;
        const dbExport = clientDb.export();
        const dbSize = dbExport.byteLength;

        statRowCount.textContent = `Hàng: ${rowsCount}`;
        statFullCount.textContent = `Ảnh Full: ${fullCount}`;
        statHalfCount.textContent = `Ảnh 1 Nửa: ${halfCount}`;
        statTotalImages.textContent = `Tổng ảnh BLOB: ${totalImg}`;
        statDbSize.textContent = `SQLite: ${formatFileSize(dbSize)}`;

        dbEngineVal.textContent = "SQLite 3 (WebAssembly / In-Browser)";
        dbPathVal.textContent = "GitHub Pages • Trình duyệt (IndexedDB/RAM)";
        dbPathVal.title = "CSDL chạy trực tiếp trên trình duyệt";
        dbSizeVal.textContent = formatFileSize(dbSize);
        dbImagesVal.textContent = `${totalImg} ảnh (${fullCount} Full + ${halfCount} Nửa)`;

        schemaTableBody.innerHTML = `
          <tr><td><code>id</code></td><td><span class="kbd-badge">INTEGER</span></td><td>🔑 Khóa chính</td><td><code>NULL</code></td></tr>
          <tr><td><code>row_order</code></td><td><span class="kbd-badge">INTEGER</span></td><td>-</td><td><code>0</code></td></tr>
          <tr><td><code>name</code></td><td><span class="kbd-badge">TEXT</span></td><td>-</td><td><code>''</code></td></tr>
          <tr><td><code>full_image_data</code></td><td><span class="kbd-badge">BLOB</span></td><td>-</td><td><code>NULL</code></td></tr>
          <tr><td><code>full_image_mime</code></td><td><span class="kbd-badge">TEXT</span></td><td>-</td><td><code>NULL</code></td></tr>
          <tr><td><code>full_image_name</code></td><td><span class="kbd-badge">TEXT</span></td><td>-</td><td><code>NULL</code></td></tr>
          <tr><td><code>full_image_size</code></td><td><span class="kbd-badge">INTEGER</span></td><td>-</td><td><code>0</code></td></tr>
          <tr><td><code>half_image_data</code></td><td><span class="kbd-badge">BLOB</span></td><td>-</td><td><code>NULL</code></td></tr>
          <tr><td><code>half_image_mime</code></td><td><span class="kbd-badge">TEXT</span></td><td>-</td><td><code>NULL</code></td></tr>
          <tr><td><code>half_image_name</code></td><td><span class="kbd-badge">TEXT</span></td><td>-</td><td><code>NULL</code></td></tr>
          <tr><td><code>half_image_size</code></td><td><span class="kbd-badge">INTEGER</span></td><td>-</td><td><code>0</code></td></tr>
          <tr><td><code>created_at</code></td><td><span class="kbd-badge">TIMESTAMP</span></td><td>-</td><td><code>CURRENT_TIMESTAMP</code></td></tr>
          <tr><td><code>updated_at</code></td><td><span class="kbd-badge">TIMESTAMP</span></td><td>-</td><td><code>CURRENT_TIMESTAMP</code></td></tr>
        `;
      } catch (e) {
        console.error("Lỗi tính stats client:", e);
      }
      return;
    }

    try {
      const res = await fetch("/api/db/stats");
      const stats = await res.json();

      statRowCount.textContent = `Hàng: ${stats.total_rows}`;
      statFullCount.textContent = `Ảnh Full: ${stats.full_images_count}`;
      statHalfCount.textContent = `Ảnh 1 Nửa: ${stats.half_images_count}`;
      statTotalImages.textContent = `Tổng ảnh BLOB: ${stats.total_images_stored}`;
      statDbSize.textContent = `SQLite: ${stats.db_size_formatted}`;

      dbEngineVal.textContent = stats.engine;
      dbPathVal.textContent = stats.db_path;
      dbPathVal.title = stats.db_path;
      dbSizeVal.textContent = stats.db_size_formatted;
      dbImagesVal.textContent = `${stats.total_images_stored} ảnh (${stats.full_images_count} Full + ${stats.half_images_count} Nửa)`;

      schemaTableBody.innerHTML = "";
      if (stats.columns) {
        stats.columns.forEach(col => {
          const tr = document.createElement("tr");
          tr.innerHTML = `
            <td><code>${col.name}</code></td>
            <td><span class="kbd-badge">${col.type}</span></td>
            <td>${col.pk ? "🔑 Khóa chính" : "-"}</td>
            <td><code>${col.dflt_value || "NULL"}</code></td>
          `;
          schemaTableBody.appendChild(tr);
        });
      }
    } catch (err) {
      console.error("Lỗi lấy thông tin CSDL:", err);
    }
  }

  // -------------------------------------------------------------
  // Lightbox Modal Handling
  // -------------------------------------------------------------
  function openLightbox(url, title, type) {
    zoomLevel = 1.0;
    rotationDeg = 0;
    applyImageTransform();
    lightboxImg.src = url;
    lightboxTitle.textContent = title || "Xem ảnh chi tiết";
    lightboxBadge.textContent = type === "full" ? "ẢNH CHỤP FULL" : "ẢNH CHỤP 1 NỬA";
    btnDownloadImage.href = url;
    btnDownloadImage.download = `${type}_image.png`;
    lightboxModal.classList.remove("hidden");
  }

  function closeLightbox() {
    lightboxModal.classList.add("hidden");
    lightboxImg.src = "";
  }

  function applyImageTransform() {
    lightboxImg.style.transform = `scale(${zoomLevel}) rotate(${rotationDeg}deg)`;
    zoomLevelDisplay.textContent = `${Math.round(zoomLevel * 100)}%`;
  }

  btnZoomIn.addEventListener("click", () => {
    zoomLevel = Math.min(zoomLevel + 0.25, 3.5);
    applyImageTransform();
  });

  btnZoomOut.addEventListener("click", () => {
    zoomLevel = Math.max(zoomLevel - 0.25, 0.4);
    applyImageTransform();
  });

  btnRotate.addEventListener("click", () => {
    rotationDeg = (rotationDeg + 90) % 360;
    applyImageTransform();
  });

  btnResetZoom.addEventListener("click", () => {
    zoomLevel = 1.0;
    rotationDeg = 0;
    applyImageTransform();
  });

  btnCloseLightbox.addEventListener("click", closeLightbox);
  lightboxModal.addEventListener("click", (e) => {
    if (e.target === lightboxModal) closeLightbox();
  });

  // -------------------------------------------------------------
  // SQL Manager Modal & Custom Queries
  // -------------------------------------------------------------
  btnSqlManager.addEventListener("click", () => {
    updateStats();
    sqlModal.classList.remove("hidden");
  });

  btnCloseSqlModal.addEventListener("click", () => {
    sqlModal.classList.add("hidden");
  });

  sqlModal.addEventListener("click", (e) => {
    if (e.target === sqlModal) sqlModal.classList.add("hidden");
  });

  btnPresetSelect.addEventListener("click", () => {
    sqlQueryInput.value = "SELECT id, row_order, name, created_at, updated_at FROM records ORDER BY id ASC;";
    runCustomSql();
  });

  btnPresetImages.addEventListener("click", () => {
    sqlQueryInput.value = "SELECT id, name, full_image_name, full_image_size, full_image_mime, half_image_name, half_image_size FROM records WHERE full_image_size > 0 OR half_image_size > 0;";
    runCustomSql();
  });

  btnRunSql.addEventListener("click", runCustomSql);

  // Download SQLite .db file button
  btnDownloadSqlDb.addEventListener("click", () => {
    if (isClientSqlMode) {
      if (!clientDb) return;
      const data = clientDb.export();
      const blob = new Blob([data], { type: "application/x-sqlite3" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "excel_app.db";
      a.click();
      URL.revokeObjectURL(url);
      showToast("Đã tải tệp cơ sở dữ liệu SQLite (.db)!");
    } else {
      window.location.href = "/api/db/download";
    }
  });

  async function runCustomSql() {
    const query = sqlQueryInput.value.trim();
    if (!query) return;

    btnRunSql.disabled = true;
    btnRunSql.innerHTML = "Đang chạy...";
    sqlResultContainer.classList.remove("hidden");
    sqlResultContainer.innerHTML = `<div style="padding: 16px; color: var(--text-muted); text-align: center;">Đang thực thi câu lệnh SQL...</div>`;

    if (isClientSqlMode) {
      try {
        const results = clientDb.exec(query);
        if (results.length === 0) {
          sqlResultContainer.innerHTML = `<div style="padding: 14px; color: var(--primary);">✅ Câu lệnh SQL đã thực thi thành công (0 dòng trả về).</div>`;
        } else {
          const cols = results[0].columns;
          const rows = results[0].values;
          let tableHtml = `<table class="sql-result-table"><thead><tr>`;
          cols.forEach(col => tableHtml += `<th>${col}</th>`);
          tableHtml += `</tr></thead><tbody>`;

          rows.forEach(row => {
            tableHtml += `<tr>`;
            row.forEach(val => {
              let displayVal = val;
              if (val instanceof Uint8Array) {
                displayVal = `&lt;BLOB ${val.length} bytes&gt;`;
              } else if (val === null) {
                displayVal = `<span style="color:var(--text-light)">NULL</span>`;
              }
              tableHtml += `<td>${displayVal}</td>`;
            });
            tableHtml += `</tr>`;
          });
          tableHtml += `</tbody></table>`;
          sqlResultContainer.innerHTML = tableHtml;
        }
        saveClientDb();
        fetchRecords();
      } catch (err) {
        sqlResultContainer.innerHTML = `<div style="padding: 14px; color: var(--danger);">❌ Lỗi SQL: ${err.message}</div>`;
      } finally {
        btnRunSql.disabled = false;
        btnRunSql.innerHTML = `
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polygon points="5 3 19 12 5 21 5 3"></polygon></svg>
          Chạy Truy Vấn SQL
        `;
      }
      return;
    }

    try {
      const res = await fetch("/api/db/query", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ query })
      });
      const result = await res.json();

      if (!result.success) {
        sqlResultContainer.innerHTML = `
          <div style="padding: 14px; color: var(--danger); font-family: var(--font-mono); font-size: 12px; background: var(--danger-light);">
            ❌ Lỗi SQL: ${result.error}
          </div>
        `;
        return;
      }

      if (result.columns && result.rows) {
        let tableHtml = `<table class="sql-result-table"><thead><tr>`;
        result.columns.forEach(col => {
          tableHtml += `<th>${col}</th>`;
        });
        tableHtml += `</tr></thead><tbody>`;

        if (result.rows.length === 0) {
          tableHtml += `<tr><td colspan="${result.columns.length}" style="text-align:center; padding:16px; color:var(--text-muted)">Không có hàng nào thỏa mãn điều kiện</td></tr>`;
        } else {
          result.rows.forEach(row => {
            tableHtml += `<tr>`;
            result.columns.forEach(col => {
              const val = row[col] === null ? `<span style="color:var(--text-light)">NULL</span>` : String(row[col]);
              tableHtml += `<td>${val}</td>`;
            });
            tableHtml += `</tr>`;
          });
        }
        tableHtml += `</tbody></table>`;
        sqlResultContainer.innerHTML = tableHtml;
      } else {
        sqlResultContainer.innerHTML = `
          <div style="padding: 14px; color: var(--primary); font-family: var(--font-mono); font-size: 12px;">
            ✅ ${result.message}
          </div>
        `;
        await fetchRecords();
      }
    } catch (err) {
      console.error("SQL Run error:", err);
      sqlResultContainer.innerHTML = `<div style="padding: 14px; color: var(--danger);">Lỗi khi gọi API truy vấn SQL</div>`;
    } finally {
      btnRunSql.disabled = false;
      btnRunSql.innerHTML = `
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polygon points="5 3 19 12 5 21 5 3"></polygon></svg>
        Chạy Truy Vấn SQL
      `;
    }
  }

  // -------------------------------------------------------------
  // Toolbar Buttons
  // -------------------------------------------------------------
  btnAddRow.addEventListener("click", () => addNewRow());
  btnAddMultipleRows.addEventListener("click", () => addMultipleRows(5));
  btnBottomAddRow.addEventListener("click", () => addNewRow());

  btnExportExcel.addEventListener("click", () => {
    if (isClientSqlMode) {
      if (typeof window.XLSX === "undefined") {
        showToast("Đang nạp thư viện Excel...", "info");
        return;
      }
      showToast("Đang xuất tệp Excel (.xlsx)...", "info");
      const exportData = [
        ["STT", "Tên", "Ảnh chụp full", "Ảnh chụp 1 nửa", "Ngày tạo"]
      ];
      records.forEach((r, idx) => {
        exportData.push([
          idx + 1,
          r.name || "",
          r.full_image.exists ? (r.full_image.name || "Có ảnh") : "(Trống)",
          r.half_image.exists ? (r.half_image.name || "Có ảnh") : "(Trống)",
          r.created_at || ""
        ]);
      });

      const ws = window.XLSX.utils.aoa_to_sheet(exportData);
      ws['!cols'] = [{ wch: 8 }, { wch: 30 }, { wch: 25 }, { wch: 25 }, { wch: 22 }];
      const wb = window.XLSX.utils.book_new();
      window.XLSX.utils.book_append_sheet(wb, ws, "Dữ Liệu Ảnh");
      window.XLSX.writeFile(wb, "bang_du_lieu_anh.xlsx");
      showToast("Đã tải file Excel thành công!");
    } else {
      showToast("Đang tạo và tải xuống tệp Excel kèm ảnh nhúng...", "info");
      window.location.href = "/api/export/excel";
    }
  });

  searchInput.addEventListener("input", () => {
    renderTable();
  });

  // Global Keyboard shortcuts
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      closeLightbox();
      sqlModal.classList.add("hidden");
    }
    if ((e.ctrlKey || e.altKey) && e.key.toLowerCase() === "n") {
      e.preventDefault();
      addNewRow();
    }
  });

  // Initial Launch
  detectModeAndInit();
});
