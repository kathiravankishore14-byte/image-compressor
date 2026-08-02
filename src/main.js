import imageCompression from 'browser-image-compression';
import Cropper from 'cropperjs';
import 'cropperjs/dist/cropper.css';
import { jsPDF } from 'jspdf';
import './style.css';

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

// ---------- Header nav wiring ----------
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

// ---------- Step 1: Handle upload ----------
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

// ---------- Tool: Resize ----------
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
  const ctx = canvas.getContext('2d');
  ctx.drawImage(currentImg, 0, 0, width, height);

  canvas.toBlob((blob) => {
    currentFile = new File([blob], currentFile.name, { type: currentFile.type });
    showImage(currentFile);
    status.textContent = `Resized to ${width}x${height}`;
    resizePanel.classList.add('hidden');
    showDownload(currentFile, `resized-${currentFile.name}`);
  }, currentFile.type);
});

// ---------- Tool: Compress ----------
async function runCompress() {
  resizePanel.classList.add('hidden');
  cropControls.classList.add('hidden');

  const originalSizeKB = (currentFile.size / 1024).toFixed(1);
  status.textContent = `Compressing... (current: ${originalSizeKB} KB)`;

  const options = { maxSizeMB: 1, maxWidthOrHeight: 1920, useWebWorker: true };

  try {
    const compressedFile = await imageCompression(currentFile, options);
    const compressedSizeKB = (compressedFile.size / 1024).toFixed(1);
    const reduction = (100 - (compressedFile.size / currentFile.size) * 100).toFixed(1);

    currentFile = compressedFile;
    showImage(currentFile);
    status.textContent = `Compressed: ${originalSizeKB} KB → ${compressedSizeKB} KB (${reduction}% smaller)`;
    showDownload(currentFile, `compressed-${currentFile.name}`);
  } catch (error) {
    status.textContent = `Error: ${error.message}`;
  }
}

// ---------- Tool: Crop ----------
function runCrop() {
  resizePanel.classList.add('hidden');
  cropControls.classList.remove('hidden');

  if (cropperInstance) cropperInstance.destroy();
  cropperInstance = new Cropper(currentImg, {
    viewMode: 1,
    autoCropArea: 0.8,
  });
}

document.querySelector('#applyCrop').addEventListener('click', () => {
  if (!cropperInstance) return;

  const canvas = cropperInstance.getCroppedCanvas();
  canvas.toBlob((blob) => {
    currentFile = new File([blob], currentFile.name, { type: currentFile.type });
    cropperInstance.destroy();
    cropperInstance = null;
    cropControls.classList.add('hidden');
    showImage(currentFile);
    status.textContent = 'Crop applied.';
    showDownload(currentFile, `cropped-${currentFile.name}`);
  }, currentFile.type);
});

document.querySelector('#cancelCrop').addEventListener('click', () => {
  if (cropperInstance) {
    cropperInstance.destroy();
    cropperInstance = null;
  }
  cropControls.classList.add('hidden');
  showImage(currentFile);
});

// ---------- Tool: Convert to PDF ----------
function runPdf() {
  resizePanel.classList.add('hidden');
  cropControls.classList.add('hidden');

  const img = new Image();
  img.onload = () => {
    const pdf = new jsPDF({
      orientation: img.width > img.height ? 'landscape' : 'portrait',
      unit: 'px',
      format: [img.width, img.height],
    });
    pdf.addImage(img, 'JPEG', 0, 0, img.width, img.height);
    const pdfBlob = pdf.output('blob');

    status.textContent = 'Converted to PDF.';
    showDownload(pdfBlob, `${currentFile.name.split('.')[0]}.pdf`);
  };
  img.src = URL.createObjectURL(currentFile);
}

// ---------- Shared: show download link ----------
function showDownload(blobOrFile, filename) {
  const url = URL.createObjectURL(blobOrFile);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.textContent = `Download ${filename}`;
  link.className = 'download-btn';

  results.innerHTML = '';
  results.appendChild(link);
}