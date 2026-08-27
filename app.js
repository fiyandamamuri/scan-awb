/* global pdfjsLib, Tesseract, XLSX */

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
      const page = await pdf.getPage(pageNo);
      const viewport = page.getViewport({ scale: 2.25 });
      const canvas = document.createElement("canvas");
      const ctx = canvas.getContext("2d", { willReadFrequently: true });

      canvas.width = Math.ceil(viewport.width);
      canvas.height = Math.ceil(viewport.height);

      await page.render({ canvasContext: ctx, viewport }).promise;

      // OCR satu halaman penuh. Format CORESYS menampilkan AWB dan URL /detail_awb/<kode>,
      // sehingga URL dipakai sebagai sumber utama karena paling aman dari kolom No. Reference.
      const result = await Tesseract.recognize(canvas, "eng", {
        logger: (m) => {
          if (m.status === "recognizing text" && typeof m.progress === "number") {
            const pageBase = (pageNo - 1) / pdf.numPages;
            const pagePart = m.progress / pdf.numPages;
            const pct = Math.min(98, Math.round((pageBase + pagePart) * 100));
            setProgress(pct, `OCR halaman ${pageNo}/${pdf.numPages}...`);
          }
        },
      });

      const pageFound = extractAwbFromOcr(result.data.text || "");
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

    // Jaga agar yang diekspor benar-benar tampak seperti kode AWB.
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
        "Belum ada AWB yang terbaca. Coba PDF dengan kualitas lebih jelas atau hasil export yang sama seperti template CORESYS.",
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

function extractAwbFromOcr(rawText) {
  const text = String(rawText || "")
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .toUpperCase();

  const found = [];

  // 1) Jalur utama: kode setelah /detail_awb/ pada hyperlink di kolom No. AWB.
  const urlRegex = /DETAIL\s*[_\- ]?\s*AWB\s*[\\/]\s*([A-Z0-9]{12,22})/g;
  let match;
  while ((match = urlRegex.exec(text)) !== null) {
    found.push(match[1]);
  }

  // 2) OCR kadang memecah '_' atau '/' menjadi spasi/tanda lain.
  const looseUrlRegex = /DETAIL\s*[_\- ]?\s*AWB[^A-Z0-9]{0,8}([A-Z0-9]{12,22})/g;
  while ((match = looseUrlRegex.exec(text)) !== null) {
    found.push(match[1]);
  }

  // 3) Fallback: hanya baca bagian sesudah header tabel "No. AWB".
  // Ini mengurangi risiko ikut mengambil daftar input "Masukan No." pada halaman pertama.
  if (!found.length) {
    const tableText = sliceAfterAwbHeader(text);
    const lines = tableText.split(/\r?\n/).map((x) => x.trim()).filter(Boolean);

    for (const line of lines) {
      // Prioritaskan baris yang berisi URL/detail_awb.
      if (/DETAIL|CORESYSSAP|ONLINE\./.test(line)) {
        const candidates = line.match(/\bQ?[A-Z0-9]{14,18}\b/g) || [];
        if (candidates.length) found.push(candidates[0]);
      }
    }
  }

  return found;
}

function sliceAfterAwbHeader(text) {
  const variants = ["NO. AWB", "NO AWB", "NO, AWB", "NO.AWB"];
  let bestIndex = -1;

  for (const variant of variants) {
    const index = text.lastIndexOf(variant);
    if (index > bestIndex) bestIndex = index;
  }

  return bestIndex >= 0 ? text.slice(bestIndex) : text;
}

function cleanCandidate(value) {
  return String(value || "")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "")
    .trim();
}

function isLikelyAwb(value) {
  if (!/^[A-Z0-9]+$/.test(value)) return false;
  if (value.length < 14 || value.length > 18) return false;
  if (/^(HTTPS|ONLINE|CORESYS|REFERENCE|DETAIL|TRACKING)/.test(value)) return false;
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
