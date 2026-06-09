const requiredFilesEl = document.querySelector('#requiredFiles');
const fileInput = document.querySelector('#fileInput');
const outputFilenameInput = document.querySelector('#outputFilename');
const buildButton = document.querySelector('#buildButton');
const downloadButton = document.querySelector('#downloadButton');
const logsEl = document.querySelector('#logs');

function setLogs(lines) {
  logsEl.textContent = Array.isArray(lines) && lines.length ? lines.join('\n') : '暂无日志';
}

function disableDownload() {
  downloadButton.href = '#';
  downloadButton.className = 'pointer-events-none rounded-xl bg-slate-300 px-5 py-3 text-sm font-semibold text-white';
  downloadButton.setAttribute('aria-disabled', 'true');
}

function enableDownload(url) {
  downloadButton.href = url;
  downloadButton.className = 'rounded-xl bg-emerald-600 px-5 py-3 text-sm font-semibold text-white shadow-sm hover:bg-emerald-500';
  downloadButton.setAttribute('aria-disabled', 'false');
}

async function loadRequiredFiles() {
  const response = await fetch('/api/required-files');
  const data = await response.json();
  requiredFilesEl.innerHTML = '';
  data.requiredFiles.forEach((filename) => {
    const item = document.createElement('li');
    item.className = 'rounded-lg bg-white px-3 py-2 ring-1 ring-slate-200';
    item.textContent = `${filename.replace('.xlsx', '')}*.xlsx`;
    requiredFilesEl.appendChild(item);
  });
}

async function buildWorkbook() {
  disableDownload();
  const files = Array.from(fileInput.files || []);
  if (files.length === 0) {
    setLogs(['WARN: 请先选择要上传的源工作簿。']);
    return;
  }

  const formData = new FormData();
  formData.append('output_filename', outputFilenameInput.value || '清查表（已填充）.xlsx');
  files.forEach((file) => formData.append('files', file));

  buildButton.disabled = true;
  buildButton.textContent = '生成中...';
  setLogs(['INFO: 正在上传并生成，请稍候...']);

  try {
    const response = await fetch('/api/build', { method: 'POST', body: formData });
    const data = await response.json();
    setLogs(data.logs || []);
    if (response.ok && data.downloadUrl) {
      enableDownload(data.downloadUrl);
    }
  } catch (error) {
    setLogs([`WARN: 请求失败：${error}`]);
  } finally {
    buildButton.disabled = false;
    buildButton.textContent = '生成清查表';
  }
}

buildButton.addEventListener('click', buildWorkbook);
loadRequiredFiles().catch((error) => setLogs([`WARN: 无法加载必需文件清单：${error}`]));
