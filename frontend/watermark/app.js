const IMAGE_EXTS = new Set(['.jpg', '.jpeg', '.png', '.webp', '.bmp', '.gif']);
const FORMAT_MAP = { 'image/jpeg': 'image/jpeg', 'image/png': 'image/png' };

const dropZone = document.getElementById('dropZone');
const folderInput = document.getElementById('folderInput');
const villageInput = document.getElementById('villageInput');
const fileListEl = document.getElementById('fileList');
const fileListArea = document.getElementById('fileListArea');
const statusEl = document.getElementById('status');
const processBtn = document.getElementById('processBtn');
const progressArea = document.getElementById('progressArea');
const progressLabel = document.getElementById('progressLabel');
const progressPercent = document.getElementById('progressPercent');
const progressBar = document.getElementById('progressBar');
const resultArea = document.getElementById('resultArea');
const successCount = document.getElementById('successCount');
const skipCount = document.getElementById('skipCount');
const downloadBtn = document.getElementById('downloadBtn');

let images = [];
let processedBlobs = [];

function isImageFile(name) {
  const ext = name.slice(name.lastIndexOf('.')).toLowerCase();
  return IMAGE_EXTS.has(ext);
}

function extractFolder(relativePath) {
  const parts = relativePath.split('/');
  return parts.length > 1 ? parts[0] : '未分组';
}

function mergeNewFiles(newFiles) {
  const existing = new Set(images.map(i => i.relativePath));
  const added = [];
  for (const img of newFiles) {
    if (!existing.has(img.relativePath)) {
      images.push(img);
      added.push(img);
    }
  }
  return added;
}

// ─── File loading ───────────────────────────────────────────

function loadFromInput(input) {
  const files = Array.from(input.files || []);
  if (!files.length) return;
  const newImages = files
    .filter(f => isImageFile(f.name))
    .map(f => ({
      file: f,
      relativePath: f.webkitRelativePath || f.name,
      folder: extractFolder(f.webkitRelativePath || f.name),
      filename: f.name,
      hasGPS: null,
      gps: null,
      date: null,
      skipped: false,
    }));
  const added = mergeNewFiles(newImages);
  if (added.length) {
    readAllExif();
  } else {
    statusEl.textContent = '没有新的图片文件（已全部在列表中）';
  }
}

async function loadFromDrop(items) {
  const newImages = [];
  const entries = [];
  for (const item of items) {
    const entry = item.webkitGetAsEntry();
    if (entry) entries.push(entry);
  }
  for (const entry of entries) {
    await walkEntry(entry, '', newImages);
  }
  if (!newImages.length) {
    statusEl.textContent = '未找到图片文件';
    return;
  }
  const added = mergeNewFiles(newImages);
  if (added.length) {
    readAllExif();
  } else {
    statusEl.textContent = '没有新的图片文件（已全部在列表中）';
  }
}

async function walkEntry(entry, path, collector) {
  if (entry.isFile) {
    if (isImageFile(entry.name)) {
      const file = await new Promise((resolve, reject) => entry.file(resolve, reject));
      collector.push({
        file,
        relativePath: path ? path + '/' + entry.name : entry.name,
        folder: extractFolder(path ? path + '/' + entry.name : entry.name),
        filename: entry.name,
        hasGPS: null,
        gps: null,
        date: null,
        skipped: false,
      });
    }
  } else if (entry.isDirectory) {
    const reader = entry.createReader();
    const entries = await readDirEntries(reader);
    for (const child of entries) {
      await walkEntry(child, path ? path + '/' + entry.name : entry.name, collector);
    }
  }
}

function readDirEntries(reader) {
  return new Promise((resolve) => {
    const all = [];
    function read() {
      reader.readEntries((results) => {
        if (results.length) {
          all.push(...results);
          read();
        } else {
          resolve(all);
        }
      });
    }
    read();
  });
}

// ─── EXIF reading ───────────────────────────────────────────

async function readAllExif() {
  statusEl.textContent = `正在读取 ${images.length} 张图片的 EXIF 信息...`;
  for (const img of images) {
    if (img.hasGPS != null) continue; // already read
    try {
      const gps = await exifr.gps(img.file);
      if (gps && gps.latitude != null && gps.longitude != null) {
        img.hasGPS = true;
        img.gps = gps;
      } else {
        img.hasGPS = false;
      }
    } catch {
      img.hasGPS = false;
    }
    img.date = new Date(img.file.lastModified);
  }
  renderFileList();
}

// ─── Render file list ───────────────────────────────────────

function renderFileList() {
  const folders = new Map();
  for (const img of images) {
    if (!folders.has(img.folder)) folders.set(img.folder, []);
    folders.get(img.folder).push(img);
  }

  fileListEl.innerHTML = '';
  for (const [folder, items] of folders) {
    const gpsOk = items.filter(i => i.hasGPS).length;
    const details = document.createElement('details');
    details.className = 'folder-group';
    details.open = true;

    const summary = document.createElement('summary');
    summary.className = 'folder-group-header';
    summary.textContent = `📁 ${folder}  (${items.length} 张，${gpsOk} 张有 GPS)`;
    details.appendChild(summary);

    for (const img of items) {
      const item = document.createElement('div');
      item.className = 'file-item';
      const icon = img.hasGPS ? '✅' : '⚠️';
      item.innerHTML = `
        <span class="status-icon">${icon}</span>
        <span class="name">${img.filename}</span>
        <span class="status-text">${img.hasGPS ? '有 GPS' : '无 GPS，将跳过'}</span>
      `;
      details.appendChild(item);
    }
    fileListEl.appendChild(details);
  }

  fileListArea.classList.remove('hidden');
  processBtn.classList.remove('hidden');
  statusEl.textContent = `共 ${images.length} 张图片，${images.filter(i => i.hasGPS).length} 张可处理`;
}

// ─── Process ────────────────────────────────────────────────

async function processImages() {
  const village = villageInput.value.trim();
  const toProcess = images.filter(i => i.hasGPS);
  const skipped = images.filter(i => !i.hasGPS);

  if (!toProcess.length) {
    statusEl.textContent = '没有包含 GPS 信息的图片可供处理';
    return;
  }

  processBtn.disabled = true;
  processBtn.textContent = '处理中...';
  progressArea.classList.remove('hidden');
  resultArea.classList.add('hidden');
  processedBlobs = [];

  let done = 0;
  const total = toProcess.length;

  for (const img of toProcess) {
    progressLabel.textContent = `${img.filename}`;
    const percent = Math.round((done / total) * 100);
    progressBar.style.width = percent + '%';
    progressPercent.textContent = percent + '%';

    try {
      const blob = await renderWatermark(img, village);
      processedBlobs.push({ blob, folder: img.folder, filename: img.filename });
      done++;
    } catch {
      done++;
    }
  }

  progressBar.style.width = '100%';
  progressPercent.textContent = '100%';
  progressLabel.textContent = '处理完成';

  const success = processedBlobs.length;
  successCount.textContent = success;
  skipCount.textContent = skipped.length;

  processBtn.disabled = false;
  processBtn.textContent = '开始处理';
  resultArea.classList.remove('hidden');
}

// ─── Watermark rendering ────────────────────────────────────

function renderWatermark(img, village) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => {
      const w = image.naturalWidth;
      const h = image.naturalHeight;
      const fontSize = Math.max(12, Math.min(w, h) * 0.02);
      const lineHeight = fontSize * 1.5;
      const padding = fontSize * 0.4;
      const margin = fontSize * 1.0;
      const radius = fontSize * 0.3;

      const lines = [
        `拍摄时间：${formatDate(img.date)}`,
        `经度：${img.gps.longitude.toFixed(6)}°E`,
        `纬度：${img.gps.latitude.toFixed(6)}°N`,
        `地块编码：${img.folder}`,
        `村名：${village || ''}`,
      ];

      const canvas = document.createElement('canvas');
      canvas.width = w;
      canvas.height = h;
      const ctx = canvas.getContext('2d');

      ctx.drawImage(image, 0, 0);

      ctx.font = `${fontSize}px "Microsoft YaHei", "PingFang SC", sans-serif`;
      ctx.textBaseline = 'top';
      ctx.textAlign = 'left';

      const maxWidth = lines.reduce((max, l) => Math.max(max, ctx.measureText(l).width), 0);
      const bgW = maxWidth + padding * 2;
      const bgH = lines.length * lineHeight + padding * 2;
      const bgX = margin;
      const bgY = h - margin - bgH;

      ctx.beginPath();
      roundRect(ctx, bgX, bgY, bgW, bgH, radius);
      ctx.fillStyle = 'rgba(0, 0, 0, 0.55)';
      ctx.fill();

      ctx.fillStyle = '#ffffff';
      for (let i = 0; i < lines.length; i++) {
        ctx.fillText(lines[i], bgX + padding, bgY + padding + i * lineHeight);
      }

      const mime = FORMAT_MAP[img.file.type] || 'image/jpeg';
      const quality = mime === 'image/png' ? undefined : 1.0;
      canvas.toBlob((blob) => {
        if (blob) resolve(blob);
        else reject(new Error('Canvas toBlob failed'));
      }, mime, quality);
    };
    image.onerror = () => reject(new Error('Image load failed'));
    image.src = URL.createObjectURL(img.file);
  });
}

function roundRect(ctx, x, y, w, h, r) {
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + w - r, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + r);
  ctx.lineTo(x + w, y + h - r);
  ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
  ctx.lineTo(x + r, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - r);
  ctx.lineTo(x, y + r);
  ctx.quadraticCurveTo(x, y, x + r, y);
  ctx.closePath();
}

function formatDate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  const hh = String(date.getHours()).padStart(2, '0');
  const mm = String(date.getMinutes()).padStart(2, '0');
  return `${y}.${m}.${d} ${hh}:${mm}`;
}

// ─── ZIP download ───────────────────────────────────────────

async function downloadZip() {
  if (!processedBlobs.length) return;
  const zip = new JSZip();

  const groups = new Map();
  for (const item of processedBlobs) {
    if (!groups.has(item.folder)) groups.set(item.folder, []);
    groups.get(item.folder).push(item);
  }

  for (const [folder, items] of groups) {
    const folderEl = zip.folder(folder);
    for (const item of items) {
      folderEl.file(item.filename, item.blob);
    }
  }

  downloadBtn.disabled = true;
  downloadBtn.textContent = '打包中...';
  const blob = await zip.generateAsync({ type: 'blob' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = '水印图片.zip';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  downloadBtn.disabled = false;
  downloadBtn.textContent = '下载全部 ZIP';
}

// ─── UI events ──────────────────────────────────────────────

dropZone.addEventListener('click', () => folderInput.click());

dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const items = e.dataTransfer.items;
  if (items && items.length) {
    loadFromDrop(items);
  }
});

folderInput.addEventListener('change', () => {
  if (folderInput.files.length) {
    loadFromInput(folderInput);
  }
  folderInput.value = '';
});

processBtn.addEventListener('click', processImages);
downloadBtn.addEventListener('click', downloadZip);
