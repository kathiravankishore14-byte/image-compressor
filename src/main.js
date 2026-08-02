import imageCompression from 'browser-image-compression';
import Cropper from 'cropperjs';
import 'cropperjs/dist/cropper.css';
import { jsPDF } from 'jspdf';
import 'jspdf-autotable';
import pptxgen from 'pptxgenjs';
import Tesseract from 'tesseract.js';
import * as XLSX from 'xlsx';
import mammoth from 'mammoth';
import './style.css';

// ================= IMAGE TOOLS =================
document.querySelector('#toolWidget').innerHTML = `
  <input type="file" id="fileInput" accept="image/*" />
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

const fileInput = document.querySelector('#fileInput');
const status = document.querySelector('#status');
const resizePanel = document.querySelector('#resizePanel');
const imageArea = document.querySelector('#imageArea');
const cropControls = document.querySelector('#cropControls');
const results = document.querySelector('#results');

let currentFile = null;
let currentImg = null;
let cropperInstance = null;
let pendingTool = null;

const formatMime = { jpg: 'image/jpeg', png: 'image/png', webp: 'image/webp' };

document.querySelectorAll('.dropdown a[data-tool]').forEach((link) => {
  link.addEventListener('click', (e) => {
    e.preventDefault();
    pendingTool = link.dataset.tool;
    fileInput.accept = formatMime[link.dataset.format] || 'image/*';
    document.querySelector('#toolWidget').scrollIntoView({ behavior: 'smooth', block: 'center' });
    fileInput.click();
  });
});

fileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;

  currentFile = file;
  results.innerHTML = '';
  resizePanel.classList.add('hidden');
  cropControls.classList.add('hidden');
  status.textContent = `Loaded: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;

  showImage(file);

  if (pendingTool) {
    const tool = pendingTool;
    pendingTool = null;
    runTool(tool);
  }
});

function showImage(fileOrBlob) {
  imageArea.innerHTML = '';
  const url = URL.createObjectURL(fileOrBlob);
  const img = document.createElement('img');
  img.src = url;
  img.id = 'currentImage';
  img.className = 'preview';
  imageArea.appendChild(img);
  currentImg = img;
}

function runTool(tool) {
  if (tool === 'resize') runResize();
  if (tool === 'compress') runCompress();
  if (tool === 'crop') runCrop();
  if (tool === 'pdf') runPdf();
}

function runResize() {
  resizePanel.classList.remove('hidden');
  cropControls.classList.add('hidden');
}

document.querySelector('#applyResize').addEventListener('click', () => {
  const width = parseInt(document.querySelector('#resizeWidth').value);
  const height = parseInt(document.querySelector('#resizeHeight').value);
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
    resizePanel.classList.add('hidden');
    showDownload(currentFile, `resized-${currentFile.name}`, results);
  }, currentFile.type);
});

async function runCompress() {
  resizePanel.classList.add('hidden');
  cropControls.classList.add('hidden');
  const originalSizeKB = (currentFile.size / 1024).toFixed(1);
  status.textContent = `Compressing... (current: ${originalSizeKB} KB)`;
  try {
    const compressedFile = await imageCompression(currentFile, { maxSizeMB: 1, maxWidthOrHeight: 1920, useWebWorker: true });
    const compressedSizeKB = (compressedFile.size / 1024).toFixed(1);
    const reduction = (100 - (compressedFile.size / currentFile.size) * 100).toFixed(1);
    currentFile = compressedFile;
    showImage(currentFile);
    status.textContent = `Compressed: ${originalSizeKB} KB → ${compressedSizeKB} KB (${reduction}% smaller)`;
    showDownload(currentFile, `compressed-${currentFile.name}`, results);
  } catch (error) {
    status.textContent = `Error: ${error.message}`;
  }
}

function runCrop() {
  resizePanel.classList.add('hidden');
  cropControls.classList.remove('hidden');
  if (cropperInstance) cropperInstance.destroy();
  cropperInstance = new Cropper(currentImg, { viewMode: 1, autoCropArea: 0.8 });
}

document.querySelector('#applyCrop').addEventListener('click', () => {
  if (!cropperInstance) return;
  cropperInstance.getCroppedCanvas().toBlob((blob) => {
    currentFile = new File([blob], currentFile.name, { type: currentFile.type });
    cropperInstance.destroy();
    cropperInstance = null;
    cropControls.classList.add('hidden');
    showImage(currentFile);
    status.textContent = 'Crop applied.';
    showDownload(currentFile, `cropped-${currentFile.name}`, results);
  }, currentFile.type);
});

document.querySelector('#cancelCrop').addEventListener('click', () => {
  if (cropperInstance) { cropperInstance.destroy(); cropperInstance = null; }
  cropControls.classList.add('hidden');
  showImage(currentFile);
});

function runPdf() {
  resizePanel.classList.add('hidden');
  cropControls.classList.add('hidden');
  const img = new Image();
  img.onload = () => {
    const pdf = new jsPDF({ orientation: img.width > img.height ? 'landscape' : 'portrait', unit: 'px', format: [img.width, img.height] });
    pdf.addImage(img, 'JPEG', 0, 0, img.width, img.height);
    status.textContent = 'Converted to PDF.';
    showDownload(pdf.output('blob'), `${currentFile.name.split('.')[0]}.pdf`, results);
  };
  img.src = URL.createObjectURL(currentFile);
}

function showDownload(blobOrFile, filename, container) {
  const url = URL.createObjectURL(blobOrFile);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.textContent = `Download ${filename}`;
  link.className = 'download-btn';
  container.innerHTML = '';
  container.appendChild(link);
}

// ================= DOCUMENT TOOLS =================
const docWidget = document.querySelector('#docToolWidget');
let activeDocTool = null;

document.querySelectorAll('.dropdown a[data-doctool]').forEach((link) => {
  link.addEventListener('click', (e) => {
    e.preventDefault();
    activeDocTool = link.dataset.doctool;
    docWidget.scrollIntoView({ behavior: 'smooth', block: 'center' });
    renderDocTool(activeDocTool);
  });
});

function renderDocTool(tool) {
  if (tool === 'texttoppt') {
    docWidget.innerHTML = `
      <p class="doc-label">Paste your text. Separate slides with a blank line — first line of each block becomes the slide title.</p>
      <textarea id="pptText" rows="8" placeholder="Slide 1 title
Bullet point one
Bullet point two

Slide 2 title
Another bullet"></textarea>
      <button id="genPpt" class="doc-action-btn">Generate PPTX</button>
      <div id="docStatus"></div>
      <div id="docResults"></div>
    `;
    document.querySelector('#genPpt').addEventListener('click', generatePpt);
  }

  if (tool === 'imagetoexcel') {
    docWidget.innerHTML = `
      <div class="drop-zone" id="dropZone">
        <input type="file" id="imgToExcelInput" accept="image/*" />
        <p class="drop-text">Drag &amp; drop an image here, or click to browse</p>
      </div>
      <p class="doc-label">Works best on clear, printed tables — not handwriting.</p>
      <div id="docStatus"></div>
      <div id="docResults"></div>
    `;
    document.querySelector('#imgToExcelInput').addEventListener('change', imageToExcel);
    setupDropZone('dropZone', 'imgToExcelInput');
  }

  if (tool === 'wordtoexcel') {
    docWidget.innerHTML = `
      <div class="drop-zone" id="dropZone">
        <input type="file" id="wordToExcelInput" accept=".docx" />
        <p class="drop-text">Drag &amp; drop a .docx file here, or click to browse</p>
      </div>
      <p class="doc-label">Works best on Word docs containing tables.</p>
      <div id="docStatus"></div>
      <div id="docResults"></div>
    `;
    document.querySelector('#wordToExcelInput').addEventListener('change', wordToExcel);
    setupDropZone('dropZone', 'wordToExcelInput');
  }

  if (tool === 'exceltopdf') {
    docWidget.innerHTML = `
      <div class="drop-zone" id="dropZone">
        <input type="file" id="excelToPdfInput" accept=".xlsx,.xls,.csv" />
        <p class="drop-text">Drag &amp; drop a spreadsheet here, or click to browse</p>
      </div>
      <div id="docStatus"></div>
      <div id="docResults"></div>
    `;
    document.querySelector('#excelToPdfInput').addEventListener('change', excelToPdf);
    setupDropZone('dropZone', 'excelToPdfInput');
  }
}

function setupDropZone(zoneId, inputId) {
  const zone = document.getElementById(zoneId);
  const input = document.getElementById(inputId);

  zone.addEventListener('click', (e) => {
    if (e.target !== input) input.click();
  });

  zone.addEventListener('dragover', (e) => {
    e.preventDefault();
    zone.classList.add('drag-active');
  });

  zone.addEventListener('dragleave', () => {
    zone.classList.remove('drag-active');
  });

  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    zone.classList.remove('drag-active');
    if (e.dataTransfer.files.length > 0) {
      input.files = e.dataTransfer.files;
      input.dispatchEvent(new Event('change'));
    }
  });
}

async function generatePpt() {
  const text = document.querySelector('#pptText').value.trim();
  const docStatus = document.querySelector('#docStatus');
  const docResults = document.querySelector('#docResults');

  if (!text) {
    docStatus.textContent = 'Please paste some text first.';
    return;
  }

  docStatus.textContent = 'Generating PPTX...';
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
  docStatus.textContent = `Generated ${blocks.length} slide(s).`;
  showDownload(blob, 'presentation.pptx', docResults);
}

async function imageToExcel(e) {
  const file = e.target.files[0];
  if (!file) return;
  const docStatus = document.querySelector('#docStatus');
  const docResults = document.querySelector('#docResults');

  docStatus.textContent = 'Reading image (this can take a moment)...';
  try {
    const { data: { text } } = await Tesseract.recognize(file, 'eng', {
      logger: (m) => {
        if (m.status === 'recognizing text') {
          docStatus.textContent = `Reading image... ${Math.round(m.progress * 100)}%`;
        }
      },
    });

    const rows = text.split('\n').filter((l) => l.trim()).map((line) => line.trim().split(/\s{2,}|\t/));
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });

    docStatus.textContent = `Extracted ${rows.length} row(s).`;
    showDownload(blob, `${file.name.split('.')[0]}.xlsx`, docResults);
  } catch (error) {
    docStatus.textContent = `Error: ${error.message}`;
  }
}

async function wordToExcel(e) {
  const file = e.target.files[0];
  if (!file) return;
  const docStatus = document.querySelector('#docStatus');
  const docResults = document.querySelector('#docResults');

  docStatus.textContent = 'Reading Word document...';
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
        const worksheet = XLSX.utils.aoa_to_sheet(rows);
        XLSX.utils.book_append_sheet(workbook, worksheet, `Table${i + 1}`);
      });
      docStatus.textContent = `Found ${tables.length} table(s).`;
    } else {
      const paragraphs = Array.from(doc.querySelectorAll('p')).map((p) => [p.textContent.trim()]).filter((r) => r[0]);
      const worksheet = XLSX.utils.aoa_to_sheet(paragraphs);
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
      docStatus.textContent = 'No tables found — imported paragraphs as single column.';
    }

    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    showDownload(blob, `${file.name.split('.')[0]}.xlsx`, docResults);
  } catch (error) {
    docStatus.textContent = `Error: ${error.message}`;
  }
}

async function excelToPdf(e) {
  const file = e.target.files[0];
  if (!file) return;
  const docStatus = document.querySelector('#docStatus');
  const docResults = document.querySelector('#docResults');

  docStatus.textContent = 'Reading spreadsheet...';
  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

    const pdf = new jsPDF({ orientation: rows[0] && rows[0].length > 8 ? 'landscape' : 'portrait' });
    pdf.autoTable({
      head: [rows[0]],
      body: rows.slice(1),
      styles: { fontSize: 8 },
      headStyles: { fillColor: [37, 99, 235] },
    });

    docStatus.textContent = `Converted sheet "${sheetName}" (${rows.length} rows).`;
    showDownload(pdf.output('blob'), `${file.name.split('.')[0]}.pdf`, docResults);
  } catch (error) {
    docStatus.textContent = `Error: ${error.message}`;
  }
}