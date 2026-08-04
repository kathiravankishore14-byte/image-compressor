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

// ================= STATE =================
let originalFile = null;
let currentFile = null;
let fileCategory = null;
let usedTools = new Set();
let cropperInstance = null;
let currentImg = null;

const categoryTools = {
  image: ['resize', 'compress', 'crop', 'pdf', 'imagetoexcel', 'imagetoppt'],
  word: ['wordtoexcel'],
  excel: ['exceltopdf'],
  text: ['texttoppt'],
};

const toolMeta = {
  resize: { label: 'Resize Image', needsConfig: true, accept: 'image/*', category: 'image' },
  compress: { label: 'Compress Image', needsConfig: false, accept: 'image/*', category: 'image' },
  crop: { label: 'Crop Image', needsConfig: true, accept: 'image/*', category: 'image' },
  pdf: { label: 'Convert to PDF', needsConfig: false, accept: 'image/*', category: 'image' },
  imagetoexcel: { label: 'Image to Excel', needsConfig: false, accept: 'image/*', category: 'image' },
  imagetoppt: { label: 'Image to PPT', needsConfig: true, accept: 'image/*', category: 'image' },
  wordtoexcel: { label: 'Word to Excel', needsConfig: false, accept: '.docx', category: 'word' },
  exceltopdf: { label: 'Excel to PDF', needsConfig: false, accept: '.xlsx,.xls,.csv', category: 'excel' },
  texttoppt: { label: 'Text to PPT', needsConfig: true, accept: null, category: 'text' },
};

function detectCategory(file) {
  const name = file.name.toLowerCase();
  if (file.type.startsWith('image/')) return 'image';
  if (name.endsWith('.docx')) return 'word';
  if (name.endsWith('.xlsx') || name.endsWith('.xls') || name.endsWith('.csv')) return 'excel';
  return null;
}

// ================= SCREEN NAV =================
function showScreen(id) {
  document.querySelectorAll('.screen').forEach((s) => s.classList.add('hidden'));
  document.getElementById(id).classList.remove('hidden');
  if (id !== 'screenResult') {
    document.getElementById(id).scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

// ================= SCREEN 1: Landing =================
const landingDropZone = document.querySelector('#landingDropZone');
const landingFileInput = document.querySelector('#landingFileInput');

landingDropZone.addEventListener('click', (e) => { if (e.target !== landingFileInput) landingFileInput.click(); });
landingDropZone.addEventListener('dragover', (e) => { e.preventDefault(); landingDropZone.classList.add('drag-active'); });
landingDropZone.addEventListener('dragleave', () => landingDropZone.classList.remove('drag-active'));
landingDropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  landingDropZone.classList.remove('drag-active');
  if (e.dataTransfer.files.length > 0) {
    landingFileInput.files = e.dataTransfer.files;
    landingFileInput.dispatchEvent(new Event('change'));
  }
});
landingFileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const cat = detectCategory(file);
  if (!cat) return;
  originalFile = file;
  currentFile = file;
  fileCategory = cat;
  usedTools = new Set();
  renderToolsScreen();
  showScreen('screenTools');
});

document.querySelector('#getStartedBtn').addEventListener('click', () => {
  originalFile = null;
  currentFile = null;
  fileCategory = null;
  usedTools = new Set();
  renderCategoryPicker();
  showScreen('screenTools');
});

// ================= SCREEN 2: Category picker + Tool selection =================
const screenToolsContent = document.querySelector('#screenToolsContent');

function renderCategoryPicker() {
  screenToolsContent.innerHTML = `
    <p class="eyebrow">02 / WHAT KIND OF FILE?</p>
    <h2 class="screen-title">Pick a category</h2>
    <div class="tool-select-grid" id="categoryGrid">
      <button class="tool-select-btn" data-cat="image">Image (JPG/PNG)</button>
      <button class="tool-select-btn" data-cat="word">Word Document</button>
      <button class="tool-select-btn" data-cat="excel">Excel Spreadsheet</button>
      <button class="tool-select-btn" data-cat="text">Text (paste in)</button>
    </div>
    <button class="back-link" id="backToLanding1">← Back</button>
  `;
  document.querySelectorAll('#categoryGrid .tool-select-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      fileCategory = btn.dataset.cat;
      renderToolsScreen();
    });
  });
  document.querySelector('#backToLanding1').addEventListener('click', () => showScreen('screenLanding'));
}

function renderToolsScreen(excludeUsed = false) {
  let tools = categoryTools[fileCategory] || [];
  if (excludeUsed) tools = tools.filter((t) => !usedTools.has(t));

  screenToolsContent.innerHTML = `
    <p class="eyebrow">02 / CHOOSE A TOOL</p>
    <h2 class="screen-title">What do you want to do?</h2>
    <div class="tool-select-grid" id="toolsGrid">
      ${tools.map((t) => `<button class="tool-select-btn" data-tool="${t}">${toolMeta[t].label}</button>`).join('')}
    </div>
    <button class="back-link" id="backToLanding2">← Back</button>
  `;
  document.querySelectorAll('#toolsGrid .tool-select-btn').forEach((btn) => {
    btn.addEventListener('click', () => pickTool(btn.dataset.tool));
  });
  document.querySelector('#backToLanding2').addEventListener('click', () => showScreen('screenLanding'));
}

// ================= SCREEN 3: Configuration =================
const screenConfigContent = document.querySelector('#screenConfigContent');

function pickTool(toolKey) {
  const meta = toolMeta[toolKey];

  if (toolKey === 'texttoppt') {
    renderTextToPptConfig();
    showScreen('screenConfig');
    return;
  }

  if (!currentFile) {
    renderFileDropConfig(toolKey);
    showScreen('screenConfig');
    return;
  }

  if (meta.needsConfig) {
    renderToolConfig(toolKey, currentFile);
    showScreen('screenConfig');
  } else {
    runTool(toolKey, currentFile);
  }
}

function renderFileDropConfig(toolKey) {
  const meta = toolMeta[toolKey];
  screenConfigContent.innerHTML = `
    <p class="eyebrow">03 / ADD YOUR FILE</p>
    <h2 class="screen-title">${meta.label}</h2>
    <div class="drop-zone" id="configDropZone">
      <input type="file" id="configFileInput" accept="${meta.accept || '*/*'}" />
      <p class="drop-text">Drop your file here, or click to browse</p>
    </div>
    <button class="back-link" id="backToTools1">← Choose a different tool</button>
  `;
  const dz = document.querySelector('#configDropZone');
  const input = document.querySelector('#configFileInput');
  dz.addEventListener('click', (e) => { if (e.target !== input) input.click(); });
  dz.addEventListener('dragover', (e) => { e.preventDefault(); dz.classList.add('drag-active'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('drag-active'));
  dz.addEventListener('drop', (e) => {
    e.preventDefault();
    dz.classList.remove('drag-active');
    if (e.dataTransfer.files.length > 0) {
      input.files = e.dataTransfer.files;
      input.dispatchEvent(new Event('change'));
    }
  });
  input.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    originalFile = file;
    currentFile = file;
    if (meta.needsConfig) {
      renderToolConfig(toolKey, file);
    } else {
      runTool(toolKey, file);
    }
  });
  document.querySelector('#backToTools1').addEventListener('click', () => showScreen('screenTools'));
}

function renderToolConfig(toolKey, file) {
  if (toolKey === 'resize') renderResizeConfig(file);
  if (toolKey === 'crop') renderCropConfig(file);
  if (toolKey === 'imagetoppt') renderImageToPptConfig(file);
}

function renderResizeConfig(file) {
  screenConfigContent.innerHTML = `
    <p class="eyebrow">03 / RESIZE</p>
    <h2 class="screen-title">Set new dimensions</h2>
    <div id="configImageArea"></div>
    <div class="panel">
      <label>Width (px): <input type="number" id="resizeWidth" placeholder="e.g. 800" /></label>
      <label>Height (px): <input type="number" id="resizeHeight" placeholder="e.g. 600" /></label>
      <button id="applyResize">Apply Resize</button>
    </div>
    <button class="back-link" id="backToTools2">← Choose a different tool</button>
  `;
  showConfigImage(file);
  document.querySelector('#applyResize').addEventListener('click', () => {
    const width = parseInt(document.querySelector('#resizeWidth').value);
    const height = parseInt(document.querySelector('#resizeHeight').value);
    if (!width || !height) return;
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    canvas.getContext('2d').drawImage(currentImg, 0, 0, width, height);
    canvas.toBlob((blob) => {
      const resized = new File([blob], file.name, { type: file.type });
      runTool('resize', file, resized);
    }, file.type);
  });
  document.querySelector('#backToTools2').addEventListener('click', () => showScreen('screenTools'));
}

function renderCropConfig(file) {
  screenConfigContent.innerHTML = `
    <p class="eyebrow">03 / CROP</p>
    <h2 class="screen-title">Drag to select an area</h2>
    <div id="configImageArea"></div>
    <div id="cropControls">
      <button id="applyCrop">Apply Crop</button>
    </div>
    <button class="back-link" id="backToTools3">← Choose a different tool</button>
  `;
  showConfigImage(file);
  setTimeout(() => {
    if (cropperInstance) cropperInstance.destroy();
    cropperInstance = new Cropper(currentImg, { viewMode: 1, autoCropArea: 0.8 });
  }, 50);
  document.querySelector('#applyCrop').addEventListener('click', () => {
    if (!cropperInstance) return;
    cropperInstance.getCroppedCanvas().toBlob((blob) => {
      const cropped = new File([blob], file.name, { type: file.type });
      cropperInstance.destroy();
      cropperInstance = null;
      runTool('crop', file, cropped);
    }, file.type);
  });
  document.querySelector('#backToTools3').addEventListener('click', () => showScreen('screenTools'));
}

function renderImageToPptConfig(file) {
  screenConfigContent.innerHTML = `
    <p class="eyebrow">03 / IMAGE TO PPT</p>
    <h2 class="screen-title">Customize your slide</h2>
    <div id="configImageArea"></div>
    <div class="panel">
      <label>Image position:
        <select id="pptPosition">
          <option value="full">Full slide</option>
          <option value="centered">Centered with margin</option>
          <option value="left">Left half</option>
          <option value="right">Right half</option>
        </select>
      </label>
      <label>Slide title (optional): <input type="text" id="pptTitle" placeholder="e.g. Product Photo" /></label>
      <button id="applyImgToPpt">Generate PPTX</button>
    </div>
    <button class="back-link" id="backToTools4">← Choose a different tool</button>
  `;
  showConfigImage(file);
  document.querySelector('#applyImgToPpt').addEventListener('click', async () => {
    const position = document.querySelector('#pptPosition').value;
    const title = document.querySelector('#pptTitle').value.trim();
    const pptx = new pptxgen();
    const slide = pptx.addSlide();

    if (title) {
      slide.addText(title, { x: 0.5, y: 0.3, w: 9, h: 0.8, fontSize: 24, bold: true });
    }

    const imgUrl = URL.createObjectURL(file);
    const positions = {
      full: { x: 0, y: 0, w: 10, h: 5.63 },
      centered: { x: 2, y: title ? 1.3 : 0.8, w: 6, h: 3.8 },
      left: { x: 0.3, y: 1, w: 4.5, h: 3.8 },
      right: { x: 5.2, y: 1, w: 4.5, h: 3.8 },
    };
    slide.addImage({ path: imgUrl, ...positions[position] });

    const blob = await pptx.write('blob');
    startProcessingScreen();
    await runVideoOnly();
    finishAndShowResult('imagetoppt', blob, 'image-slide.pptx');
  });
  document.querySelector('#backToTools4').addEventListener('click', () => showScreen('screenTools'));
}

function renderTextToPptConfig() {
  screenConfigContent.innerHTML = `
    <p class="eyebrow">03 / TEXT TO PPT</p>
    <h2 class="screen-title">Paste your text</h2>
    <p class="doc-label">Separate slides with a blank line — first line of each block becomes the slide title.</p>
    <textarea id="pptText" rows="8" placeholder="Slide 1 title
Bullet point one
Bullet point two

Slide 2 title
Another bullet"></textarea>
    <button id="genPpt" class="doc-action-btn">Generate PPTX</button>
    <button class="back-link" id="backToTools5">← Back</button>
  `;
  document.querySelector('#genPpt').addEventListener('click', async () => {
    const text = document.querySelector('#pptText').value.trim();
    if (!text) return;
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
    startProcessingScreen();
    await runVideoOnly();
    finishAndShowResult('texttoppt', blob, 'presentation.pptx');
  });
  document.querySelector('#backToTools5').addEventListener('click', () => showScreen('screenTools'));
}

function showConfigImage(file) {
  const area = document.querySelector('#configImageArea');
  area.innerHTML = '';
  const url = URL.createObjectURL(file);
  const img = document.createElement('img');
  img.src = url;
  img.className = 'preview';
  area.appendChild(img);
  currentImg = img;
}

// ================= RUN TOOLS (no-config path) =================
async function runTool(toolKey, sourceFile, precomputedResult = null) {
  startProcessingScreen();

  let resultBlob, resultName;

  try {
    if (precomputedResult) {
      resultBlob = precomputedResult;
      resultName = `${toolKey}-${sourceFile.name}`;
    } else if (toolKey === 'compress') {
      resultBlob = await imageCompression(sourceFile, { maxSizeMB: 1, maxWidthOrHeight: 1920, useWebWorker: true });
      resultName = `compressed-${sourceFile.name}`;
    } else if (toolKey === 'pdf') {
      resultBlob = await imageToPdfBlob(sourceFile);
      resultName = `${sourceFile.name.split('.')[0]}.pdf`;
    } else if (toolKey === 'imagetoexcel') {
      resultBlob = await imageToExcelBlob(sourceFile);
      resultName = `${sourceFile.name.split('.')[0]}.xlsx`;
    } else if (toolKey === 'wordtoexcel') {
      resultBlob = await wordToExcelBlob(sourceFile);
      resultName = `${sourceFile.name.split('.')[0]}.xlsx`;
    } else if (toolKey === 'exceltopdf') {
      resultBlob = await excelToPdfBlob(sourceFile);
      resultName = `${sourceFile.name.split('.')[0]}.pdf`;
    }

    await runVideoOnly();
    finishAndShowResult(toolKey, resultBlob, resultName);
  } catch (error) {
    await runVideoOnly();
    document.querySelector('#screenResultContent').innerHTML = `<p class="eyebrow">ERROR</p><p style="color:#fff;">${error.message}</p>`;
  }
}

function imageToPdfBlob(file) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const pdf = new jsPDF({ orientation: img.width > img.height ? 'landscape' : 'portrait', unit: 'px', format: [img.width, img.height] });
      pdf.addImage(img, 'JPEG', 0, 0, img.width, img.height);
      resolve(pdf.output('blob'));
    };
    img.src = URL.createObjectURL(file);
  });
}

async function imageToExcelBlob(file) {
  const { data: { text } } = await Tesseract.recognize(file, 'eng');
  const rows = text.split('\n').filter((l) => l.trim()).map((line) => line.trim().split(/\s{2,}|\t/));
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  return new Blob([wbout], { type: 'application/octet-stream' });
}

async function wordToExcelBlob(file) {
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
  } else {
    const paragraphs = Array.from(doc.querySelectorAll('p')).map((p) => [p.textContent.trim()]).filter((r) => r[0]);
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(paragraphs), 'Sheet1');
  }
  const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  return new Blob([wbout], { type: 'application/octet-stream' });
}

async function excelToPdfBlob(file) {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
  const pdf = new jsPDF({ orientation: rows[0] && rows[0].length > 8 ? 'landscape' : 'portrait' });
  autoTable(pdf, { head: [rows[0]], body: rows.slice(1), styles: { fontSize: 8 }, headStyles: { fillColor: [37, 99, 235] } });
  return pdf.output('blob');
}

// ================= SCREEN 4: Processing + Result (full viewport) =================
const screenResultContent = document.querySelector('#screenResultContent');

function startProcessingScreen() {
  showScreen('screenResult');
  screenResultContent.innerHTML = `
    <video id="resultVideo" class="hero-video-full" src="/hero-video/hero-process-sound.mp4" playsinline></video>
    <div id="resultBelow" class="hidden"></div>
  `;
  const video = document.querySelector('#resultVideo');
  video.muted = false;
  video.play().catch(() => { video.muted = true; video.play(); });
}

function runVideoOnly() {
  return new Promise((resolve) => {
    const video = document.querySelector('#resultVideo');
    if (!video) { resolve(); return; }
    const onEnded = () => { video.removeEventListener('ended', onEnded); resolve(); };
    video.addEventListener('ended', onEnded);
  });
}

function finishAndShowResult(toolKey, blob, filename) {
  usedTools.add(toolKey);
  if (toolKey !== 'imagetoppt' && toolKey !== 'texttoppt') {
    currentFile = new File([blob], filename, { type: blob.type || 'application/octet-stream' });
  }

  const url = URL.createObjectURL(blob);
  const belowArea = document.querySelector('#resultBelow');
  belowArea.classList.remove('hidden');
  belowArea.innerHTML = `
    <a href="${url}" download="${filename}" class="download-btn">Download ${filename}</a>
    <div id="furtherOptionsArea"></div>
  `;

  renderFurtherOptions();
}

function renderFurtherOptions() {
  const remaining = (categoryTools[fileCategory] || []).filter((t) => !usedTools.has(t));
  const area = document.querySelector('#furtherOptionsArea');
  if (!area) return;

  if (remaining.length === 0 || fileCategory === 'text') {
    area.innerHTML = `<p class="doc-label" style="margin-top:20px;">That's every tool for this file. <button class="back-link" id="startOverBtn">Start Over</button></p>`;
    document.querySelector('#startOverBtn').addEventListener('click', resetToLanding);
    return;
  }

  area.innerHTML = `
    <p class="doc-label" style="margin-top:24px;">Want to do more with this file?</p>
    <div class="tool-select-grid">
      ${remaining.map((t) => `<button class="tool-select-btn" data-further="${t}">${toolMeta[t].label}</button>`).join('')}
    </div>
    <button class="back-link" id="startOverBtn2">Start Over</button>
  `;
  document.querySelector('#startOverBtn2').addEventListener('click', resetToLanding);
  document.querySelectorAll('[data-further]').forEach((btn) => {
    btn.addEventListener('click', () => showFileChooser(btn.dataset.further));
  });
}

function showFileChooser(toolKey) {
  const area = document.querySelector('#furtherOptionsArea');
  let secondsLeft = 5;
  area.innerHTML = `
    <p class="doc-label">Use which file for "${toolMeta[toolKey].label}"?</p>
    <div class="tool-select-grid">
      <button class="tool-select-btn" id="useOriginal">Original file</button>
      <button class="tool-select-btn" id="useProcessed">Processed file (default in <span id="countdown">${secondsLeft}</span>s)</button>
    </div>
  `;
  const timer = setInterval(() => {
    secondsLeft -= 1;
    const el = document.querySelector('#countdown');
    if (el) el.textContent = secondsLeft;
    if (secondsLeft <= 0) {
      clearInterval(timer);
      proceedWithChain(toolKey, currentFile);
    }
  }, 1000);

  document.querySelector('#useOriginal').addEventListener('click', () => {
    clearInterval(timer);
    proceedWithChain(toolKey, originalFile);
  });
  document.querySelector('#useProcessed').addEventListener('click', () => {
    clearInterval(timer);
    proceedWithChain(toolKey, currentFile);
  });
}

function proceedWithChain(toolKey, fileToUse) {
  currentFile = fileToUse;
  const meta = toolMeta[toolKey];
  if (meta.needsConfig) {
    renderToolConfig(toolKey, fileToUse);
    showScreen('screenConfig');
  } else {
    runTool(toolKey, fileToUse);
  }
}

function resetToLanding() {
  originalFile = null;
  currentFile = null;
  fileCategory = null;
  usedTools = new Set();
  showScreen('screenLanding');
}

// ================= HEADER NAV =================
document.querySelectorAll('[data-navtool]').forEach((link) => {
  link.addEventListener('click', (e) => {
    e.preventDefault();
    originalFile = null;
    currentFile = null;
    fileCategory = link.dataset.navcat;
    usedTools = new Set();
    pickTool(link.dataset.navtool);
  });
});