import imageCompression from 'browser-image-compression';
import Cropper from 'cropperjs';
import 'cropperjs/dist/cropper.css';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import pptxgen from 'pptxgenjs';
import Tesseract from 'tesseract.js';
import * as XLSX from 'xlsx';
import mammoth from 'mammoth';
import './style.css';

const widget = document.querySelector('#toolWidget');
const eyebrow = document.querySelector('#toolEyebrow');

let pendingAction = null;
let currentFile = null;
let currentImg = null;
let cropperInstance = null;

const MIN_ANIM_MS = 5000;

const formatMime = { jpg: 'image/jpeg', png: 'image/png', webp: 'image/webp' };

const toolLabels = {
  resize: 'Resize', compress: 'Compress', crop: 'Crop', pdf: 'Convert to PDF',
  texttoppt: 'Text to PPT', imagetoexcel: 'Image to Excel',
  wordtoexcel: 'Word to Excel', exceltopdf: 'Excel to PDF',
};

const acceptMap = {
  resize: () => formatMime[pendingAction.format] || 'image/*',
  compress: () => formatMime[pendingAction.format] || 'image/*',
  crop: () => formatMime[pendingAction.format] || 'image/*',
  pdf: () => formatMime[pendingAction.format] || 'image/*',
  imagetoexcel: () => 'image/*',
  wordtoexcel: () => '.docx',
  exceltopdf: () => '.xlsx,.xls,.csv',
};

// ---------- Hero video overlay ----------
function ensureOverlay() {
  let overlay = document.querySelector('#heroOverlay');
  if (!overlay) {
    overlay = document.createElement('div');
    overlay.id = 'heroOverlay';
    overlay.className = 'hero-overlay hidden';
    overlay.innerHTML = `
      <video id="heroVideo" class="hero-video" src="/hero-video/hero-process-sound.mp4" loop playsinline></video>
      <div class="progress-bar-track"><div class="progress-bar-fill"></div></div>
      <p class="overlay-caption" id="overlayCaption">Working...</p>
    `;
    document.body.appendChild(overlay);
  }
  return overlay;
}

function playSprite(setName, fps, caption) {
  const overlay = ensureOverlay();
  overlay.classList.remove('hidden');
  document.querySelector('#overlayCaption').textContent = caption;

  const video = document.querySelector('#heroVideo');
  video.currentTime = 0;
  video.muted = false;
  video.play().catch(() => {
    // Browser blocked autoplay-with-sound — fall back to muted so it still plays visually
    video.muted = true;
    video.play();
  });
}

function stopSprite() {
  const overlay = document.querySelector('#heroOverlay');
  if (overlay) {
    overlay.classList.add('hidden');
    const video = document.querySelector('#heroVideo');
    if (video) video.pause();
  }
}

async function waitOutMinimum(startTime) {
  const elapsed = Date.now() - startTime;
  const remaining = MIN_ANIM_MS - elapsed;
  if (remaining > 0) {
    await new Promise((resolve) => setTimeout(resolve, remaining));
  }
}

// ---------- Widget shell ----------
function renderWidget() {
  widget.innerHTML = `
    <div class="drop-zone" id="masterDropZone">
      <input type="file" id="masterFileInput" accept="*/*" />
      <p class="drop-text" id="dropText">Choose a tool above, then drop your file here</p>
    </div>

    <div id="pptPanel" class="panel hidden">
      <textarea id="pptText" rows="8" placeholder="Slide 1 title
Bullet point one
Bullet point two

Slide 2 title
Another bullet"></textarea>
      <button id="genPpt" class="doc-action-btn">Generate PPTX</button>
    </div>

    <div id="status"></div>

    <div id="resizePanel" class="panel hidden">
      <label>Width (px): <input type="number" id="resizeWidth" placeholder="e.g. 800" /></label>
      <label>Height (px): <input type="number" id="resizeHeight" placeholder="e.g. 600" /></label>
      <button id="applyResize">Apply Resize</button>
    </div>

    <div id="imageArea"></div>

    <div id="cropControls" class="hidden">
      <button id="applyCrop">Apply Crop</button>
      <button id="cancelCrop">Cancel</button>
    </div>

    <div id="results"></div>
  `;
  wireStaticControls();
}

renderWidget();

document.querySelectorAll('.dropdown a[data-tool], .dropdown a[data-doctool]').forEach((link) => {
  link.addEventListener('click', (e) => {
    e.preventDefault();

    const isImageTool = !!link.dataset.tool;
    pendingAction = isImageTool
      ? { kind: 'image', tool: link.dataset.tool, format: link.dataset.format }
      : { kind: 'doc', tool: link.dataset.doctool };

    resetWidgetState();
    eyebrow.textContent = `SELECTED / ${toolLabels[pendingAction.tool].toUpperCase()}`;
    widget.scrollIntoView({ behavior: 'smooth', block: 'center' });

    if (pendingAction.tool === 'texttoppt') {
      document.querySelector('#masterDropZone').classList.add('hidden');
      document.querySelector('#pptPanel').classList.remove('hidden');
    } else {
      document.querySelector('#pptPanel').classList.add('hidden');
      document.querySelector('#masterDropZone').classList.remove('hidden');
      const input = document.querySelector('#masterFileInput');
      input.accept = acceptMap[pendingAction.tool] ? acceptMap[pendingAction.tool]() : '*/*';
      document.querySelector('#dropText').textContent = `Drop a file for "${toolLabels[pendingAction.tool]}", or click to browse`;
    }
  });
});

function resetWidgetState() {
  document.querySelector('#status').textContent = '';
  document.querySelector('#results').innerHTML = '';
  document.querySelector('#imageArea').innerHTML = '';
  document.querySelector('#resizePanel').classList.add('hidden');
  document.querySelector('#cropControls').classList.add('hidden');
  stopSprite();
  if (cropperInstance) { cropperInstance.destroy(); cropperInstance = null; }
  currentFile = null;
  currentImg = null;
}

function wireStaticControls() {
  const dropZone = document.querySelector('#masterDropZone');
  const input = document.querySelector('#masterFileInput');

  dropZone.addEventListener('click', (e) => {
    if (e.target !== input) input.click();
  });
  dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-active'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-active'));
  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-active');
    if (e.dataTransfer.files.length > 0) {
      input.files = e.dataTransfer.files;
      input.dispatchEvent(new Event('change'));
    }
  });

  input.addEventListener('change', handleFileSelected);
  document.querySelector('#genPpt').addEventListener('click', generatePpt);
  document.querySelector('#applyResize').addEventListener('click', applyResize);
  document.querySelector('#applyCrop').addEventListener('click', applyCrop);
  document.querySelector('#cancelCrop').addEventListener('click', cancelCrop);
}

function handleFileSelected(e) {
  const file = e.target.files[0];
  if (!file) return;

  const status = document.querySelector('#status');

  if (!pendingAction) {
    status.textContent = 'Please choose a tool from the menu above first.';
    e.target.value = '';
    return;
  }

  currentFile = file;
  status.textContent = `Loaded: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;

  if (pendingAction.kind === 'image') {
    showImage(file);
    if (pendingAction.tool === 'resize') runResize();
    if (pendingAction.tool === 'compress') runCompress();
    if (pendingAction.tool === 'crop') runCrop();
    if (pendingAction.tool === 'pdf') runPdf();
  } else {
    if (pendingAction.tool === 'imagetoexcel') runImageToExcel(file);
    if (pendingAction.tool === 'wordtoexcel') runWordToExcel(file);
    if (pendingAction.tool === 'exceltopdf') runExcelToPdf(file);
  }
}

function showImage(fileOrBlob) {
  const imageArea = document.querySelector('#imageArea');
  imageArea.innerHTML = '';
  const url = URL.createObjectURL(fileOrBlob);
  const img = document.createElement('img');
  img.src = url;
  img.className = 'preview';
  imageArea.appendChild(img);
  currentImg = img;
}

function showDownload(blobOrFile, filename) {
  const results = document.querySelector('#results');
  const url = URL.createObjectURL(blobOrFile);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.textContent = `Download ${filename}`;
  link.className = 'download-btn';
  results.innerHTML = '';
  results.appendChild(link);
}

// ================= IMAGE TOOLS =================
function runResize() {
  document.querySelector('#resizePanel').classList.remove('hidden');
  document.querySelector('#cropControls').classList.add('hidden');
}

function applyResize() {
  const width = parseInt(document.querySelector('#resizeWidth').value);
  const height = parseInt(document.querySelector('#resizeHeight').value);
  const status = document.querySelector('#status');
  if (!width || !height) {
    status.textContent = 'Please enter both width and height.';
    return;
  }
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  canvas.getContext('2d').drawImage(currentImg, 0, 0, width, height);
  canvas.toBlob((blob) => {
    currentFile = new File([blob], currentFile.name, { type: currentFile.type });
    showImage(currentFile);
    status.textContent = `Resized to ${width}x${height}`;
    document.querySelector('#resizePanel').classList.add('hidden');
    showDownload(currentFile, `resized-${currentFile.name}`);
  }, currentFile.type);
}

async function runCompress() {
  const status = document.querySelector('#status');
  const originalSizeKB = (currentFile.size / 1024).toFixed(1);
  status.textContent = `Compressing... (current: ${originalSizeKB} KB)`;
  const startTime = Date.now();
  playSprite('run', 16, 'Compressing your image...');
  try {
    const compressedFile = await imageCompression(currentFile, { maxSizeMB: 1, maxWidthOrHeight: 1920, useWebWorker: true });
    const compressedSizeKB = (compressedFile.size / 1024).toFixed(1);
    const reduction = (100 - (compressedFile.size / currentFile.size) * 100).toFixed(1);
    currentFile = compressedFile;
    showImage(currentFile);
    status.textContent = `Compressed: ${originalSizeKB} KB → ${compressedSizeKB} KB (${reduction}% smaller)`;
    await waitOutMinimum(startTime);
    stopSprite();
    showDownload(currentFile, `compressed-${currentFile.name}`);
  } catch (error) {
    await waitOutMinimum(startTime);
    stopSprite();
    status.textContent = `Error: ${error.message}`;
  }
}

function runCrop() {
  document.querySelector('#resizePanel').classList.add('hidden');
  document.querySelector('#cropControls').classList.remove('hidden');
  if (cropperInstance) cropperInstance.destroy();
  cropperInstance = new Cropper(currentImg, { viewMode: 1, autoCropArea: 0.8 });
}

async function applyCrop() {
  if (!cropperInstance) return;
  const startTime = Date.now();
  playSprite('walk', 12, 'Crop applied!');

  cropperInstance.getCroppedCanvas().toBlob(async (blob) => {
    currentFile = new File([blob], currentFile.name, { type: currentFile.type });
    cropperInstance.destroy();
    cropperInstance = null;
    document.querySelector('#cropControls').classList.add('hidden');
    showImage(currentFile);
    document.querySelector('#status').textContent = 'Crop applied.';
    await waitOutMinimum(startTime);
    stopSprite();
    showDownload(currentFile, `cropped-${currentFile.name}`);
  }, currentFile.type);
}

function cancelCrop() {
  if (cropperInstance) { cropperInstance.destroy(); cropperInstance = null; }
  document.querySelector('#cropControls').classList.add('hidden');
  showImage(currentFile);
}

function runPdf() {
  const startTime = Date.now();
  playSprite('fly', 14, 'Converting to PDF...');
  const img = new Image();
  img.onload = async () => {
    const pdf = new jsPDF({ orientation: img.width > img.height ? 'landscape' : 'portrait', unit: 'px', format: [img.width, img.height] });
    pdf.addImage(img, 'JPEG', 0, 0, img.width, img.height);
    document.querySelector('#status').textContent = 'Converted to PDF.';
    await waitOutMinimum(startTime);
    stopSprite();
    showDownload(pdf.output('blob'), `${currentFile.name.split('.')[0]}.pdf`);
  };
  img.src = URL.createObjectURL(currentFile);
}

// ================= DOCUMENT TOOLS =================
async function generatePpt() {
  const text = document.querySelector('#pptText').value.trim();
  const status = document.querySelector('#status');
  if (!text) {
    status.textContent = 'Please paste some text first.';
    return;
  }
  status.textContent = 'Generating PPTX...';
  const blocks = text.split(/\n\s*\n/);
  const pptx = new pptxgen();
  blocks.forEach((block) => {
    const lines = block.split('\n').filter((l) => l.trim());
    const slide = pptx.addSlide();
    slide.addText(lines[0] || 'Untitled', { x: 0.5, y: 0.4, w: 9, h: 1, fontSize: 28, bold: true });
    if (lines.length > 1) {
      slide.addText(lines.slice(1).map((l) => ({ text: l, options: { bullet: true, breakLine: true } })), { x: 0.5, y: 1.5, w: 9, h: 4.5, fontSize: 18 });
    }
  });
  const blob = await pptx.write('blob');
  status.textContent = `Generated ${blocks.length} slide(s).`;
  showDownload(blob, 'presentation.pptx');
}

async function runImageToExcel(file) {
  const status = document.querySelector('#status');
  status.textContent = 'Reading image (this can take a moment)...';
  try {
    const { data: { text } } = await Tesseract.recognize(file, 'eng', {
      logger: (m) => {
        if (m.status === 'recognizing text') status.textContent = `Reading image... ${Math.round(m.progress * 100)}%`;
      },
    });
    const rows = text.split('\n').filter((l) => l.trim()).map((line) => line.trim().split(/\s{2,}|\t/));
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    status.textContent = `Extracted ${rows.length} row(s).`;
    showDownload(new Blob([wbout], { type: 'application/octet-stream' }), `${file.name.split('.')[0]}.xlsx`);
  } catch (error) {
    status.textContent = `Error: ${error.message}`;
  }
}

async function runWordToExcel(file) {
  const status = document.querySelector('#status');
  status.textContent = 'Reading Word document...';
  try {
    const arrayBuffer = await file.arrayBuffer();
    const { value: html } = await mammoth.convertToHtml({ arrayBuffer });
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const tables = doc.querySelectorAll('table');
    const workbook = XLSX.utils.book_new();

    if (tables.length > 0) {
      tables.forEach((table, i) => {
        const rows = Array.from(table.querySelectorAll('tr')).map((tr) =>
          Array.from(tr.querySelectorAll('td, th')).map((cell) => cell.textContent.trim())
        );
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), `Table${i + 1}`);
      });
      status.textContent = `Found ${tables.length} table(s).`;
    } else {
      const paragraphs = Array.from(doc.querySelectorAll('p')).map((p) => [p.textContent.trim()]).filter((r) => r[0]);
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(paragraphs), 'Sheet1');
      status.textContent = 'No tables found — imported paragraphs as single column.';
    }

    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    showDownload(new Blob([wbout], { type: 'application/octet-stream' }), `${file.name.split('.')[0]}.xlsx`);
  } catch (error) {
    status.textContent = `Error: ${error.message}`;
  }
}

async function runExcelToPdf(file) {
  const status = document.querySelector('#status');
  status.textContent = 'Reading spreadsheet...';
  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

    const pdf = new jsPDF({ orientation: rows[0] && rows[0].length > 8 ? 'landscape' : 'portrait' });
    autoTable(pdf, {
      head: [rows[0]],
      body: rows.slice(1),
      styles: { fontSize: 8 },
      headStyles: { fillColor: [37, 99, 235] },
    });

    status.textContent = `Converted sheet "${sheetName}" (${rows.length} rows).`;
    showDownload(pdf.output('blob'), `${file.name.split('.')[0]}.pdf`);
  } catch (error) {
    status.textContent = `Error: ${error.message}`;
  }
}