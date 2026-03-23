(function () {
  const api = window.inventoryDashboard;
  if (!api) return;

  const STORAGE_KEY = "inventory-dashboard-uploaded-payload-v1";
  const PRODUCTS = ["바나나칩(70g)", "바나나칩(15g)", "밀크 초코칩", "다크 초코칩", "화이트 초코칩"];
  const ALIASES = {
    "밀크초코칩": "밀크 초코칩",
    "다크초코칩": "다크 초코칩",
    "화이트초코칩": "화이트 초코칩",
  };
  const PRODUCT_KEYS = ["p1", "p2", "p3", "p4", "p5"];

  const fileInput = document.getElementById("inventoryFileInput");
  const updateBtn = document.getElementById("inventoryUpdateBtn");
  const resetBtn = document.getElementById("inventoryResetBtn");
  const statusEl = document.getElementById("uploadStatus");
  const fileLabelEl = document.getElementById("uploadFileLabel");
  const sourceLabelEl = document.getElementById("uploadSourceLabel");

  if (!fileInput || !updateBtn || !resetBtn || !statusEl || !fileLabelEl || !sourceLabelEl) return;

  function setStatus(message, type) {
    statusEl.textContent = message;
    statusEl.className = "upload-status";
    if (type) statusEl.classList.add(type);
  }

  function clone(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function updatePanelMeta(selectedFiles, sourceText) {
    const names = selectedFiles && selectedFiles.length ? selectedFiles.map((file) => file.name).join(", ") : "없음";
    fileLabelEl.textContent = `선택 파일: ${names}`;
    const current = api.getData();
    sourceLabelEl.textContent = `데이터 원본: ${sourceText || current.file_name || "기본 내장 데이터"}`;
  }

  function canonProduct(name) {
    if (typeof name !== "string") return null;
    const normalized = ALIASES[name.trim()] || name.trim();
    return PRODUCTS.includes(normalized) ? normalized : null;
  }

  function pad2(value) {
    return String(value).padStart(2, "0");
  }

  function toLocalYmd(date) {
    return api.localYmd(date);
  }

  function excelNumberToDate(value) {
    if (typeof value !== "number" || !window.XLSX || !XLSX.SSF) return null;
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    return new Date(parsed.y, parsed.m - 1, parsed.d);
  }

  function toDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    if (typeof value === "number") return excelNumberToDate(value);
    if (typeof value === "string") {
      const parsed = new Date(value);
      if (!Number.isNaN(parsed.getTime())) return parsed;
    }
    return null;
  }

  function readCell(sheet, row, col) {
    return sheet[XLSX.utils.encode_cell({ r: row - 1, c: col - 1 })];
  }

  function cellValue(sheet, row, col) {
    const cell = readCell(sheet, row, col);
    return cell ? cell.v : undefined;
  }

  function cellText(sheet, row, col) {
    const cell = readCell(sheet, row, col);
    if (!cell) return "";
    return String(cell.w != null ? cell.w : cell.v != null ? cell.v : "").trim();
  }

  function getMaxColumn(sheet) {
    const ref = sheet["!ref"];
    if (!ref) return 0;
    return XLSX.utils.decode_range(ref).e.c + 1;
  }

  function ensureDayBucket(map, dateStr) {
    if (!map.has(dateStr)) {
      const products = {};
      PRODUCTS.forEach((product) => {
        products[product] = { total: 0, lots: [] };
      });
      map.set(dateStr, products);
    }
    return map.get(dateStr);
  }

  function buildPayloadFromFiles(files, workbooks) {
    const dailyData = new Map();

    workbooks.forEach((workbook) => {
      let yearHint = 2025;
      let prevMonth = null;

      workbook.SheetNames.forEach((sheetName) => {
        const match = /^(\d{2})-(\d{2})$/.exec(sheetName);
        if (!match) return;

        const mm = Number(match[1]);
        const dd = Number(match[2]);
        if (prevMonth !== null && mm < prevMonth) yearHint += 1;
        prevMonth = mm;

        const sheet = workbook.Sheets[sheetName];
        const rawDate = toDate(cellValue(sheet, 2, 1));
        const dateStr = rawDate ? toLocalYmd(rawDate) : `${yearHint}-${pad2(mm)}-${pad2(dd)}`;
        const bucket = ensureDayBucket(dailyData, dateStr);

        [5, 10, 15, 20, 25].forEach((startRow) => {
          const product = canonProduct(cellValue(sheet, startRow, 2));
          if (!product) return;

          let ttlCol = null;
          const maxColumn = getMaxColumn(sheet);
          for (let col = 1; col <= maxColumn; col += 1) {
            if (cellText(sheet, startRow, col).toUpperCase().includes("TTL")) {
              ttlCol = col;
              break;
            }
          }
          if (!ttlCol) return;

          const expRow = startRow + 1;
          const qtyRow = startRow + 2;
          const lots = [];

          for (let col = 2; col < ttlCol; col += 1) {
            const qty = Number(cellValue(sheet, qtyRow, col) || 0);
            if (!Number.isFinite(qty) || qty <= 0) continue;
            const expiryDate = toDate(cellValue(sheet, expRow, col));
            lots.push({
              expiry: expiryDate ? toLocalYmd(expiryDate) : null,
              qty: Math.trunc(qty),
            });
          }

          const ttlValue = Number(cellValue(sheet, qtyRow, ttlCol) || 0);
          const total = Number.isFinite(ttlValue) && ttlValue > 0
            ? Math.trunc(ttlValue)
            : lots.reduce((sum, lot) => sum + lot.qty, 0);

          bucket[product] = { total, lots };
        });
      });
    });

    const allDates = [...dailyData.keys()].sort();
    if (!allDates.length) {
      throw new Error("유효한 시트 데이터가 없습니다. 시트명이 `MM-DD` 형식인지 확인해주세요.");
    }

    const dailyRows = allDates.map((date) => {
      const row = { date };
      let sum = 0;
      PRODUCTS.forEach((product, index) => {
        const total = dailyData.get(date)[product].total || 0;
        row[PRODUCT_KEYS[index]] = total;
        sum += total;
      });
      row.sum = sum;
      return row;
    });

    const monthLast = new Map();
    allDates.forEach((date) => {
      const month = date.slice(0, 7);
      if (!monthLast.has(month) || date > monthLast.get(month)) monthLast.set(month, date);
    });

    const monthlyRows = [...monthLast.keys()].sort().map((month) => {
      const latestDate = monthLast.get(month);
      const row = { month, latest_date: latestDate };
      let sum = 0;
      PRODUCTS.forEach((product, index) => {
        const total = dailyData.get(latestDate)[product].total || 0;
        row[PRODUCT_KEYS[index]] = total;
        sum += total;
      });
      row.sum = sum;
      return row;
    });

    const weeklyLast = new Map();
    allDates.forEach((dateStr) => {
      const date = new Date(`${dateStr}T00:00:00`);
      const jan4 = new Date(date.getFullYear(), 0, 4);
      const jan4Day = jan4.getDay() || 7;
      const week1 = new Date(jan4);
      week1.setDate(jan4.getDate() - (jan4Day - 1));
      const diffDays = Math.floor((date - week1) / 86400000);
      const week = Math.floor(diffDays / 7) + 1;
      const key = `${date.getFullYear()}-${pad2(week)}`;
      if (!weeklyLast.has(key) || dateStr > weeklyLast.get(key)) weeklyLast.set(key, dateStr);
    });

    const weeklyRows = [...weeklyLast.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([key, latestDate]) => {
      const [year, week] = key.split("-");
      const latestTotal = PRODUCTS.reduce((sum, product) => sum + (dailyData.get(latestDate)[product].total || 0), 0);
      return {
        week_label: `${year}년 ${week}주`,
        latest_date: latestDate,
        latest_total: latestTotal,
      };
    });

    const latestDate = allDates[allDates.length - 1];
    const latestBlock = dailyData.get(latestDate);
    const today = toLocalYmd(new Date());
    const todayOrd = Math.floor(new Date(`${today}T00:00:00`).getTime() / 86400000);
    const nearLimit = todayOrd + 180;
    const bestBeforeLimit = todayOrd + 240;

    const expiryRows = PRODUCTS.map((product) => {
      let nearLots = 0;
      let nearQty = 0;
      let bestBeforeLots = 0;
      let bestBeforeQty = 0;
      let expiredLots = 0;

      latestBlock[product].lots.forEach((lot) => {
        if (!lot.expiry || lot.qty <= 0) return;
        const expiryOrd = Math.floor(new Date(`${lot.expiry}T00:00:00`).getTime() / 86400000);
        if (todayOrd <= expiryOrd && expiryOrd <= nearLimit) {
          nearLots += 1;
          nearQty += lot.qty;
        }
        if (todayOrd <= expiryOrd && expiryOrd <= bestBeforeLimit) {
          bestBeforeLots += 1;
          bestBeforeQty += lot.qty;
        }
        if (expiryOrd < todayOrd) expiredLots += 1;
      });

      let status = "정상";
      if (expiredLots > 0) status = "폐기 필요";
      else if (nearQty > 0) status = "임박 관리";
      else if (bestBeforeQty > 0) status = "상미기한 관리";

      return {
        product,
        near_lots: nearLots,
        near_qty: nearQty,
        best_before_lots: bestBeforeLots,
        best_before_qty: bestBeforeQty,
        status,
      };
    });

    const payload = {
      products: PRODUCTS,
      dates: allDates,
      daily_rows: dailyRows,
      monthly_rows: monthlyRows,
      weekly_rows: weeklyRows,
      expiry_rows: expiryRows,
      latest_lots: Object.fromEntries(PRODUCTS.map((product) => [product, latestBlock[product].lots])),
      generated_on: today,
      latest_date: latestDate,
      file_name: files.length === 1 ? files[0].name : `업로드 ${files.length}개 통합`,
    };

    PRODUCTS.forEach((product, index) => {
      payload[PRODUCT_KEYS[index]] = allDates.map((date) => dailyData.get(date)[product].total || 0);
    });

    return payload;
  }

  async function buildPayloadFromSelection(files) {
    if (!window.XLSX) throw new Error("엑셀 파서 로딩에 실패했습니다. 네트워크 상태를 확인해주세요.");
    const workbooks = [];
    for (const file of files) {
      const buffer = await file.arrayBuffer();
      workbooks.push(XLSX.read(buffer, { type: "array", cellDates: true }));
    }
    return buildPayloadFromFiles(files, workbooks);
  }

  function savePayload(payload) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
  }

  function loadPayload() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch (error) {
      console.warn("stored payload parse failed", error);
      return null;
    }
  }

  function applyStoredPayload() {
    const stored = loadPayload();
    if (!stored) {
      updatePanelMeta([], "기본 내장 데이터");
      return;
    }
    api.setData(clone(stored));
    updatePanelMeta([], stored.file_name || "브라우저 저장 데이터");
    setStatus("이 브라우저에 저장된 업로드 데이터로 대시보드를 복원했습니다.", "success");
  }

  fileInput.addEventListener("change", () => {
    updatePanelMeta([...fileInput.files], "업로드 대기");
    if (fileInput.files.length) {
      setStatus("파일을 확인했습니다. `업데이트`를 누르면 화면 데이터가 다시 계산됩니다.");
    } else {
      setStatus("업로드할 파일을 선택해주세요.");
    }
  });

  updateBtn.addEventListener("click", async () => {
    const files = [...fileInput.files];
    if (!files.length) {
      setStatus("업데이트할 엑셀 파일을 먼저 선택해주세요.", "error");
      return;
    }

    updateBtn.disabled = true;
    setStatus("재고 파일을 읽고 대시보드 데이터를 계산하는 중입니다...");

    try {
      const payload = await buildPayloadFromSelection(files);
      api.setData(clone(payload));
      savePayload(payload);
      updatePanelMeta(files, payload.file_name);
      setStatus(`업데이트 완료: 최신 데이터 기준일은 ${payload.latest_date} 입니다.`, "success");
    } catch (error) {
      console.error(error);
      setStatus(error.message || "업데이트 중 오류가 발생했습니다.", "error");
    } finally {
      updateBtn.disabled = false;
    }
  });

  resetBtn.addEventListener("click", () => {
    localStorage.removeItem(STORAGE_KEY);
    fileInput.value = "";
    api.setData(api.getEmbeddedData());
    updatePanelMeta([], "기본 내장 데이터");
    setStatus("기본 내장 데이터로 복원했습니다.", "success");
  });

  applyStoredPayload();
})();
