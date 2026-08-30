/* global pdfjsLib, XLSX */

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const pdfInput = document.getElementById("pdfInput");
const dropZone = document.getElementById("dropZone");
const fileInfo = document.getElementById("fileInfo");
const extractBtn = document.getElementById("extractBtn");
const downloadBtn = document.getElementById("downloadBtn");
const removeQ = document.getElementById("removeQ");
const dedupe = document.getElementById("dedupe");
const progressWrap = document.getElementById("progressWrap");
const progressText = document.getElementById("progressText");
const progressPct = document.getElementById("progressPct");
const progressBar = document.getElementById("progressBar");
const statusBox = document.getElementById("statusBox");
const resultCard = document.getElementById("resultCard");
const resultCount = document.getElementById("resultCount");
const resultBody = document.getElementById("resultBody");

let selectedFile = null;
let extractedAWBs = [];

pdfInput.addEventListener("change", (event) => {
  const [file] = event.target.files;
  setFile(file || null);
});

["dragenter", "dragover"].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.add("dragging");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.remove("dragging");
  });
});

dropZone.addEventListener("drop", (event) => {
  const file = event.dataTransfer.files?.[0];
  if (!file) return;
  setFile(file);
});

extractBtn.addEventListener("click", processPdf);
downloadBtn.addEventListener("click", exportExcel);

function setFile(file) {
  resetResult();

  if (!file) {
    selectedFile = null;
    extractBtn.disabled = true;
    fileInfo.classList.add("hidden");
    return;
  }

  if (file.type !== "application/pdf" && !file.name.toLowerCase().endsWith(".pdf")) {
    selectedFile = null;
    extractBtn.disabled = true;
    showStatus("File harus berformat PDF.", "error");
    return;
  }

  selectedFile = file;
  extractBtn.disabled = false;
  fileInfo.textContent = `${file.name} • ${formatBytes(file.size)}`;
  fileInfo.classList.remove("hidden");
  hideStatus();
}

async function processPdf() {
  if (!selectedFile) return;

  resetResult();
  extractBtn.disabled = true;
  progressWrap.classList.remove("hidden");
  hideStatus();
  setProgress(0, "Membaca PDF...");

  try {
    const buffer = await selectedFile.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buffer) }).promise;

    if (!pdf.numPages) throw new Error("PDF tidak memiliki halaman.");

    const allFound = [];

    for (let pageNo = 1; pageNo <= pdf.numPages; pageNo += 1) {
      setProgress(
        Math.round(((pageNo - 1) / pdf.numPages) * 100),
        `Mengekstrak halaman ${pageNo}/${pdf.numPages}...`
      );

      const page = await pdf.getPage(pageNo);

      // Ekstrak langsung kode AWB dari struktur digital PDF (Instant 0.1 Detik)
      const pageFound = await extractAwbDirect(page);
      allFound.push(...pageFound);

      const pct = Math.round((pageNo / pdf.numPages) * 100);
      setProgress(pct, `Selesai halaman ${pageNo}/${pdf.numPages}`);
    }

    let normalized = allFound
      .map(cleanCandidate)
      .filter(Boolean);

    if (removeQ.checked) {
      normalized = normalized.map((value) => value.replace(/^Q(?=[A-Z0-9])/i, ""));
    }

    // Filter kandidat AWB (Tanpa batas minimum, maks 18 karakter)
    normalized = normalized.filter(isLikelyAwb);

    if (dedupe.checked) {
      normalized = [...new Set(normalized)];
    }

    extractedAWBs = normalized;
    renderResult();

    if (extractedAWBs.length) {
      showStatus(`Selesai. ${extractedAWBs.length} AWB ditemukan.`, "success");
    } else {
      showStatus(
        "Belum ada AWB yang terbaca. Pastikan PDF berformat export Trace & Tracking CORESYS.",
        "error"
      );
    }
  } catch (error) {
    console.error(error);
    showStatus(`Gagal memproses PDF: ${error.message || error}`, "error");
  } finally {
    extractBtn.disabled = !selectedFile;
  }
}

// Ekstraksi langsung kode AWB dari PDF (Hyperlink Link & Teks Digital PDF.js)
async function extractAwbDirect(page) {
  const found = [];

  // A. Sumber Utama: Ambil dari Hyperlink / Annotations URL (https://.../detail_awb/<KODE_AWB>)
  try {
    const annotations = await page.getAnnotations();
    for (const ann of annotations) {
      if (ann.url) {
        const match = ann.url.match(/detail[_\-\s]*awb[_\-\s\/\:\-]*([A-Z0-9]+)/i);
        if (match) found.push(match[1]);
      }
    }
  } catch (e) {
    console.warn("Gagal membaca annotations halaman:", e);
  }

  // B. Sumber Kedua: Ambil dari Teks Digital PDF
  try {
    const textContent = await page.getTextContent();
    const items = textContent.items.map((item) => item.str);
    
    // B1. Gabung TANPA spasi (menangani token terpisah seperti detail_ awb / 8161056...)
    const rawNoSpaces = items.map((s) => s.trim()).join("");
    const regexNoSpaces = /detail[_\-\s]*awb[_\-\s\/\:\-]*([A-Z0-9]{3,20})/gi;
    let m;
    while ((m = regexNoSpaces.exec(rawNoSpaces)) !== null) {
      found.push(m[1]);
    }

    // B2. Gabung DENGAN spasi
    const rawWithSpaces = items.join(" ");
    const regexWithSpaces = /detail[_\-\s]*awb[_\-\s\/\:\-]*([A-Z0-9]{3,20})/gi;
    while ((m = regexWithSpaces.exec(rawWithSpaces)) !== null) {
      found.push(m[1]);
    }

    // B3. Ambil kode berawalan Q di sel tabel (Q8161056107366 / QCGK8161070900011)
    const qAwbRegex = /\bQ([A-Z0-9]{3,18})\b/g;
    while ((m = qAwbRegex.exec(rawWithSpaces)) !== null) {
      found.push(m[1]);
    }
  } catch (e) {
    console.warn("Gagal membaca text content halaman:", e);
  }

  return found;
}

function cleanCandidate(value) {
  return String(value || "")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "")
    .trim();
}

function isLikelyAwb(value) {
  if (!/^[A-Z0-9]+$/.test(value)) return false;
  
  // TANPA BATAS MINIMUM (bisa 4, 5, 13 digit dst).
  // Hanya membatasi maksimal 18 karakter agar No. Referensi (20 digit) tidak ikut masuk.
  if (value.length > 18) return false;

  // Daftar kata kunci teks UI / header tabel yang diabaikan
  const ignoreList = [
    "NO", "AWB", "POS", "PM", "AM", "PDF", "PAGE",
    "HTTP", "HTTPS", "ONLINE", "CORESYS", "CORESYSSAP",
    "REFERENCE", "REFERANCE", "DETAIL", "TRACKING",
    "SATRIA", "ANTARAN", "PRIMA"
  ];

  if (ignoreList.includes(value)) return false;

  return true;
}

function renderResult() {
  resultBody.innerHTML = "";
  resultCount.textContent = String(extractedAWBs.length);

  extractedAWBs.forEach((awb, index) => {
    const tr = document.createElement("tr");
    const no = document.createElement("td");
    const value = document.createElement("td");
    no.textContent = String(index + 1);
    value.textContent = awb;
    tr.append(no, value);
    resultBody.appendChild(tr);
  });

  resultCard.classList.toggle("hidden", !extractedAWBs.length);
}

function exportExcel() {
  if (!extractedAWBs.length) return;

  const rows = extractedAWBs.map((awb) => ({ "No. AWB": awb }));
  const ws = XLSX.utils.json_to_sheet(rows, { header: ["No. AWB"] });

  // Paksa format text agar AWB tidak berubah menjadi scientific notation / angka.
  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let row = range.s.r + 1; row <= range.e.r; row += 1) {
    const cell = ws[XLSX.utils.encode_cell({ r: row, c: 0 })];
    if (cell) cell.z = "@";
  }

  ws["!cols"] = [{ wch: 24 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "AWB");

  const baseName = selectedFile?.name?.replace(/\.pdf$/i, "") || "hasil_awb";
  XLSX.writeFile(wb, `${baseName}_AWB.xlsx`);
}

function setProgress(percent, text) {
  const safePercent = Math.max(0, Math.min(100, Math.round(percent)));
  progressBar.style.width = `${safePercent}%`;
  progressPct.textContent = `${safePercent}%`;
  progressText.textContent = text;
}

function showStatus(message, type = "info") {
  statusBox.textContent = message;
  statusBox.className = `status ${type}`;
}

function hideStatus() {
  statusBox.className = "status hidden";
  statusBox.textContent = "";
}

function resetResult() {
  extractedAWBs = [];
  resultBody.innerHTML = "";
  resultCount.textContent = "0";
  resultCard.classList.add("hidden");
  progressWrap.classList.add("hidden");
  progressBar.style.width = "0%";
  progressPct.textContent = "0%";
}

function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / 1024 ** index;
  return `${value.toFixed(index === 0 ? 0 : 1)} ${units[index]}`;
}