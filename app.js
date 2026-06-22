/* ============================================================
   校園樹木碳匯量測記錄表 產生器
   buildWorkbook() 為純函式（瀏覽器 / Node 皆可用），
   依範例 xlsx 的格式產生完全相同版面的工作簿。
   ============================================================ */

const KAI = "標楷體";        // 量測表字型
const MING = "新細明體";      // 地點代碼表字型

function fnt(size, bold, name) {
  return { name: name || KAI, size: size, bold: !!bold };
}
function fill(argb) {
  return { type: "pattern", pattern: "solid", fgColor: { argb: argb } };
}
function thin() { return { style: "thin", color: { argb: "FF000000" } }; }
function bd(sides) { // sides: string like "tlrb" — each side gets its own object
  const b = {};                // (ExcelJS drops sides that share one object reference)
  if (sides.indexOf("t") >= 0) b.top = thin();
  if (sides.indexOf("l") >= 0) b.left = thin();
  if (sides.indexOf("r") >= 0) b.right = thin();
  if (sides.indexOf("b") >= 0) b.bottom = thin();
  return b;
}
// 顏色（對應範例 indexed 13/11/47 與 theme0）
const C_YELLOW = "FFFFFF00";
const C_GREEN  = "FF00FF00";
const C_PEACH  = "FFFFCC99";
const C_WHITE  = "FFFFFFFF";

function fit(natW, natH, maxW, maxH) {
  if (!natW || !natH) return { width: maxW, height: maxH };
  const s = Math.min(maxW / natW, maxH / natH);
  return { width: Math.round(natW * s), height: Math.round(natH * s) };
}

/* ---------- 主產生函式 ---------- */
function buildWorkbook(d, imgs) {
  const wb = new ExcelJS.Workbook();
  buildMeasureSheet(wb, d, imgs);
  buildSpeciesSheet(wb);
  buildLocationSheet(wb);
  return wb;
}

function buildMeasureSheet(wb, d, imgs) {
  const ws = wb.addWorksheet("量測表", {
    views: [{ showGridLines: false }],
    pageSetup: {
      orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0,
      margins: { left: 0.62, right: 0.53, top: 0.41, bottom: 0.46, header: 0.3, footer: 0.3 }
    }
  });

  // 欄寬：A=8、B=17、C~J=11.9（與範例相同）
  ws.properties.defaultRowHeight = 32.4;
  ws.getColumn(1).width = 8;   // A
  ws.getColumn(2).width = 17;  // B
  for (let c = 3; c <= 10; c++) ws.getColumn(c).width = 11.88671875; // C–J

  // 列高
  const H = { 1:32.4, 2:6.6, 3:32.4, 4:32.4, 5:32.4, 6:32.4, 7:32.4,
              8:195, 9:32.4, 10:32.4, 11:47.4, 12:32.4, 13:32.4, 14:32.4, 15:32.4 };
  Object.keys(H).forEach(r => ws.getRow(+r).height = H[r]);

  // 合併儲存格
  ["A1:J1","C4:J4","C5:J5","C6:E6","G6:H6","I6:J6","C7:F7","G7:J7",
   "C8:F8","G8:J8","C9:D9","E9:F9","G9:H9","I9:J9"].forEach(m => ws.mergeCells(m));

  const set = (addr, v, o) => {
    o = o || {};
    const c = ws.getCell(addr);
    if (v !== undefined && v !== null) c.value = v;
    if (o.font) c.font = o.font;
    if (o.fill) c.fill = o.fill;
    if (o.border) c.border = o.border;
    if (o.numFmt) c.numFmt = o.numFmt;
    if (o.note) c.note = o.note;
    c.alignment = {
      horizontal: o.h || "left",
      vertical: o.v || "middle",
      wrapText: !!o.wrap
    };
    return c;
  };

  // ── 標題
  set("A1", "國立高雄師範大學 校園樹木碳匯量測記錄表", { font: fnt(20, true), h: "center", v: "middle" });

  // ── 一、基本資料
  set("A3", "一、", { font: fnt(18, true), v: "middle" });
  set("B3", "基本資料", { font: fnt(18, true), v: "middle" });

  set("B4", "樹木編號：", { font: fnt(16), h: "right", v: "middle" });
  set("C4", d.treeNo, { font: fnt(16), h: "left", v: "middle", numFmt: "@" });

  set("B5", "樹種：", { font: fnt(16), h: "right", v: "middle" });
  set("C5", d.species, { font: fnt(16), h: "left", v: "middle", numFmt: "@" });

  set("B6", "位置：", { font: fnt(16), h: "right", v: "middle" });
  set("C6", d.position || "", { font: fnt(16), h: "left", v: "middle", numFmt: "@" });
  set("F6", "GPS定位：", { font: fnt(16), v: "middle", numFmt: "@" });
  set("G6", d.gpsX !== "" && d.gpsX != null ? "TWD97/TM2-X " + d.gpsX : "TWD97/TM2-X",
      { font: fnt(12), h: "center", v: "middle", numFmt: "@" });
  set("I6", d.gpsY !== "" && d.gpsY != null ? "TWD97/TM2-Y " + d.gpsY : "TWD97/TM2-Y",
      { font: fnt(12), h: "center", v: "middle", numFmt: "@" });

  set("B7", "植生狀況：", { font: fnt(16), h: "right", v: "middle" });
  // 合併儲存格的框線只設在主儲存格上，會自動延伸到整個合併範圍的外框
  set("C7", "(近照)", { font: fnt(16), h: "center", v: "middle", border: bd("b") });   // C7:F7
  set("G7", "(大範圍)", { font: fnt(16), h: "center", v: "middle", border: bd("b") });  // G7:J7

  // ── 照片框（第 8 列）外框 = 上+左+右（下緣由第 9 列上框線形成）
  set("C8", null, { border: bd("tlr") });  // C8:F8
  set("G8", null, { border: bd("tlr") });  // G8:J8

  // ── 拍攝日期列（第 9 列）
  set("C9", "拍攝日期：", { font: fnt(16), h: "right", v: "middle", wrap: true, border: bd("tlb") }); // C9:D9
  set("E9", d.dateRoc, { font: fnt(16), h: "left", v: "middle", wrap: true, border: bd("tb") });      // E9:F9
  set("G9", "拍攝日期：", { font: fnt(16), h: "right", v: "middle", wrap: true, border: bd("tlb") }); // G9:H9 (左框=兩日期框分隔線)
  set("I9", d.dateRoc, { font: fnt(16), h: "left", v: "middle", wrap: true, border: bd("trb") });     // I9:J9

  // ── 二、量測記錄
  set("A10", "二、", { font: fnt(16), v: "middle" });
  set("B10", "量測記錄", { font: fnt(16), v: "middle" });

  // 表頭（第 11 列）
  const allb = bd("tlrb");
  set("A11", "項次", { font: fnt(16), h: "center", v: "middle", border: allb });
  set("B11", "量測日期", { font: fnt(16), h: "center", v: "middle", wrap: true, border: allb, numFmt: "yyyy/m/d;@" });
  set("C11", "直徑(cm)", { font: fnt(14), h: "center", v: "middle", wrap: true, border: allb });
  set("D11", "樹全高(m)", { font: fnt(14), h: "center", v: "middle", wrap: true, border: allb });
  // 註：原範例在此兩格有公式說明的「註解」，但 ExcelJS 的註解(VML)與圖片並存時
  // 會讓 Excel 開檔時「修復內容」而把框線清掉、照片亂掉，故移除註解以確保相容性。
  set("E11", "幹材材積 m³ ", { font: fnt(14), h: "center", v: "middle", wrap: true, border: allb,
      fill: fill(C_YELLOW) });
  set("F11", "碳匯量(ton)", { font: fnt(14), h: "center", v: "middle", wrap: true, border: allb,
      fill: fill(C_GREEN) });

  // 資料列 12-15（項次 1-4），第一列填入本次量測
  for (let i = 0; i < 4; i++) {
    const r = 12 + i;
    const first = i === 0;
    set("A" + r, i + 1, { font: fnt(16), h: "center", v: "middle", border: allb });
    set("B" + r, first ? d.dateRoc : null,
        { font: fnt(16), h: "center", v: "middle", wrap: true, border: allb, fill: fill(C_WHITE), numFmt: "@" });
    set("C" + r, first && d.diameter != null ? d.diameter : null,
        { font: fnt(16), h: "center", v: "middle", border: allb, fill: fill(C_PEACH) });
    set("D" + r, first && d.height != null ? d.height : null,
        { font: fnt(16), h: "center", v: "middle", border: allb, fill: fill(C_PEACH) });
    set("E" + r, { formula: "SUMPRODUCT(C"+r+"/100/2,C"+r+"/100/2,PI(),D"+r+",植栽種類!$D$2)" },
        { font: fnt(16), v: "middle", border: allb, fill: fill(C_YELLOW), numFmt: "0.0000_ " });
    set("F" + r, { formula: "SUMPRODUCT(E"+r+",植栽種類!$E$2,植栽種類!$H$2)" },
        { font: fnt(16), h: "center", v: "middle", border: allb, fill: fill(C_GREEN),
          numFmt: "0.0000_);[Red]\\(0.0000\\)" });
  }

  // ── 照片
  if (imgs && imgs.close) {
    const id = wb.addImage({ base64: imgs.close.base64, extension: imgs.close.ext });
    const e = fit(imgs.close.w, imgs.close.h, 170, 250);
    ws.addImage(id, { tl: { col: 2.15, row: 7.06 }, ext: e, editAs: "oneCell" });
  }
  if (imgs && imgs.wide) {
    const id = wb.addImage({ base64: imgs.wide.base64, extension: imgs.wide.ext });
    const e = fit(imgs.wide.w, imgs.wide.h, 380, 250);
    ws.addImage(id, { tl: { col: 6.1, row: 7.06 }, ext: e, editAs: "oneCell" });
  }
}

function buildSpeciesSheet(wb) {
  const ws = wb.addWorksheet("植栽種類");
  ws.getColumn(1).width = 9;
  ws.getColumn(2).width = 9.4;
  ws.getColumn(3).width = 9;
  ws.getRow(1).height = 46.2;
  SPECIES_GRID.forEach(([r, row]) => {
    row.forEach((val, ci) => {
      if (val === null || val === undefined) return;
      const cell = ws.getCell(r, ci + 1);
      if (typeof val === "string" && val.charAt(0) === "=") {
        cell.value = { formula: val.slice(1) };
      } else {
        cell.value = val;
      }
      cell.font = fnt(r === 1 ? 12 : 11);
      if (r === 1) cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    });
  });
}

function buildLocationSheet(wb) {
  const ws = wb.addWorksheet("地點代碼表");
  ws.getColumn(2).width = 21.4;
  LOC_GRID.forEach(([r, row]) => {
    row.forEach((val, ci) => {
      if (val === null || val === undefined) return;
      const cell = ws.getCell(r, ci + 1);
      cell.value = val;
      cell.font = fnt(12, false, MING);
    });
  });
}

/* ============================================================
   以下為瀏覽器介面（Node 測試時會略過）
   ============================================================ */
if (typeof document !== "undefined") {
  const $ = id => document.getElementById(id);

  // 填入地點下拉
  LOC_OPTIONS.forEach(o => {
    const opt = document.createElement("option");
    opt.value = o.code;
    opt.textContent = `${o.campus} / ${o.place}（${o.code}）`;
    if (o.code === "YE") opt.selected = true;
    $("loc").appendChild(opt);
  });

  // 填入樹種下拉（依類別分組）
  const groups = {};
  TREE_OPTIONS.forEach(o => { (groups[o.cat] = groups[o.cat] || []).push(o); });
  Object.keys(groups).forEach(cat => {
    const og = document.createElement("optgroup");
    og.label = cat;
    groups[cat].forEach(o => {
      const opt = document.createElement("option");
      opt.value = o.name;
      opt.textContent = `${o.name}（${o.code}）`;
      if (o.name === "黑板樹") opt.selected = true;
      og.appendChild(opt);
    });
    $("species").appendChild(og);
  });
  // 其他選項
  const otherOpt = document.createElement("option");
  otherOpt.value = "__other__";
  otherOpt.textContent = "其他（自行輸入）";
  $("species").appendChild(otherOpt);

  // 預設位置與日期
  const today = new Date();
  $("date").value = today.toISOString().slice(0, 10);

  function rocDate(iso) {
    if (!iso) return "";
    const [y, m, dd] = iso.split("-");
    return (parseInt(y, 10) - 1911) + "." + m + "." + dd;
  }
  function pad3(n) { return String(n).padStart(3, "0"); }
  function currentTreeNo() {
    const ov = $("noOverride").value.trim();
    if (ov) return ov;
    return $("loc").value + "-" + pad3($("seq").value || "0");
  }
  function currentSpecies() {
    return $("species").value === "__other__"
      ? $("speciesOther").value.trim()
      : $("species").value;
  }

  function refresh() {
    $("treeNoPreview").textContent = currentTreeNo();
    const c = parseFloat($("circ").value);
    $("diaPreview").textContent = isNaN(c) ? "—" : (Math.round(c / Math.PI * 10) / 10).toFixed(1);
    $("rocPreview").textContent = rocDate($("date").value) || "—";
  }
  ["loc","seq","noOverride","circ","date"].forEach(id => {
    $(id).addEventListener("input", refresh);
    $(id).addEventListener("change", refresh);
  });
  $("species").addEventListener("change", () => {
    $("speciesOtherWrap").style.display = $("species").value === "__other__" ? "block" : "none";
  });

  // 預設位置敘述跟著地點走
  function suggestPos() {
    const o = LOC_OPTIONS.find(x => x.code === $("loc").value);
    if (o && !$("position").dataset.touched) {
      $("position").value = o.place.replace(/周圍$/, "旁");
    }
  }
  $("loc").addEventListener("change", suggestPos);
  $("position").addEventListener("input", () => { $("position").dataset.touched = "1"; });
  suggestPos();
  refresh();

  // 照片預覽 + 暫存
  // 重新編碼：把手機照片的 EXIF 旋轉「燒進」影像、去掉方向標籤，並縮到合理大小。
  // 這樣嵌進 Excel 時方向正確、比例不變（Excel 不會讀 EXIF，所以一定要先轉正）。
  const photos = { close: null, wide: null };
  const MAXDIM = 1400;
  function encode(src, w, h, cb) {
    let dw = w, dh = h;
    if (Math.max(w, h) > MAXDIM) {
      const s = MAXDIM / Math.max(w, h);
      dw = Math.round(w * s); dh = Math.round(h * s);
    }
    const cv = document.createElement("canvas");
    cv.width = dw; cv.height = dh;
    cv.getContext("2d").drawImage(src, 0, 0, dw, dh);
    const url = cv.toDataURL("image/jpeg", 0.9);
    cb({ base64: url.split(",")[1], ext: "jpeg", w: dw, h: dh, dataUrl: url });
  }
  function loadPhoto(file, cb) {
    if (window.createImageBitmap) {
      createImageBitmap(file, { imageOrientation: "from-image" })
        .then(bm => { encode(bm, bm.width, bm.height, cb); bm.close && bm.close(); })
        .catch(useImg);
    } else { useImg(); }
    function useImg() {
      const img = new Image();
      img.onload = () => encode(img, img.naturalWidth, img.naturalHeight, cb);
      img.src = URL.createObjectURL(file);
    }
  }
  function hookPhoto(inputId, prevId, key) {
    $(inputId).addEventListener("change", e => {
      const f = e.target.files[0];
      if (!f) return;
      const pv = $(prevId), im = pv.querySelector("img");
      pv.style.display = "block"; im.removeAttribute("src");
      photos[key] = null;
      loadPhoto(f, res => { photos[key] = res; im.src = res.dataUrl; });
    });
  }
  hookPhoto("closePhoto", "closePrev", "close");
  hookPhoto("widePhoto", "widePrev", "wide");

  // 產生
  $("gen").addEventListener("click", async () => {
    const msg = $("msg");
    const treeNo = currentTreeNo();
    const species = currentSpecies();
    if (!species) { msg.className = "msg err"; msg.textContent = "請選擇或輸入樹種。"; return; }

    const circ = parseFloat($("circ").value);
    const height = parseFloat($("height").value);
    const d = {
      treeNo,
      species,
      position: $("position").value.trim(),
      gpsX: $("gpsX").value.trim(),
      gpsY: $("gpsY").value.trim(),
      diameter: isNaN(circ) ? null : Math.round(circ / Math.PI * 10) / 10,
      height: isNaN(height) ? null : height,
      dateRoc: rocDate($("date").value)
    };

    try {
      $("gen").disabled = true;
      $("gen").textContent = "產生中…";
      const wb = buildWorkbook(d, photos);
      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `${treeNo}-${species}.xlsx`;
      a.click();
      URL.revokeObjectURL(a.href);
      msg.className = "msg ok";
      msg.textContent = `已產生：${treeNo}-${species}.xlsx`;
    } catch (err) {
      console.error(err);
      msg.className = "msg err";
      msg.textContent = "產生失敗：" + err.message;
    } finally {
      $("gen").disabled = false;
      $("gen").textContent = "⬇ 產生並下載 Excel";
    }
  });
}

// Node 測試用
if (typeof module !== "undefined" && module.exports) {
  module.exports = { buildWorkbook };
}
