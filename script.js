pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

let pdfFiles = [];
let cleanResults = [];
let latestResult = null;
let currentPageResults = [];
let sensitivity = 2;
let totalCreditsUsed = 0;
let stopRequested = false;
let processing = false;
let metadataSources = [];
let metadataRows = [];

let googleAccessToken = '';
let googleTokenClient = null;
let activeGoogleClientId = '';
let driveFolderId = '';
const DRIVE_SCOPE = 'https://www.googleapis.com/auth/drive';

let automateCancelRequested = false;
let automateRunning = false;
let automateResults = [];
let automateCsvText = '';

let excelData = [];
let excelColumns = [];
let excelMappings = { title: '', code: '', year: '', dept: '' };

const MODELS = {
  anthropic: [
    { id: 'claude-opus-4-5', label: 'Claude Opus 4.5' },
    { id: 'claude-sonnet-4-5', label: 'Claude Sonnet 4.5' },
    { id: 'claude-haiku-4-5-20251001', label: 'Claude Haiku 4.5' },
  ],
  openai: [
    { id: 'gpt-4o-mini', label: 'ChatGPT 4o mini' },
    { id: 'gpt-5-nano', label: 'ChatGPT 5 nano' },
    { id: 'gpt-5-mini', label: 'ChatGPT 5 mini' },
    { id: 'gpt-5.4-nano', label: 'ChatGPT 5.4 nano' },
    { id: 'gpt-5.4-mini', label: 'ChatGPT 5.4 mini' },
  ],
  openrouter: [
    { id: 'openai/gpt-4o-mini', label: 'OpenRouter · GPT-4o mini' },
    { id: 'openai/gpt-5-mini', label: 'OpenRouter · GPT-5 mini' },
    { id: 'anthropic/claude-3.5-sonnet', label: 'OpenRouter · Claude 3.5 Sonnet' },
    { id: 'google/gemini-2.5-flash', label: 'OpenRouter · Gemini 2.5 Flash' },
  ],
  groq: [
    { id: 'llama-3.3-70b-versatile', label: 'Groq · Llama 3.3 70B Versatile' },
    { id: 'llama-3.1-8b-instant', label: 'Groq · Llama 3.1 8B Instant' },
    { id: 'qwen/qwen3-32b', label: 'Groq · Qwen3 32B' },
  ],
  gemini: [
    { id: 'gemini-3-flash', label: 'Gemini 3 Flash' },
    { id: 'gemini-2.5-flash', label: 'Gemini 2.5 Flash' },
    { id: 'gemini-3.1-flash-lite', label: 'Gemini 3.1 Flash Lite' },
  ],
  local: [
    { id: 'local-vision-v2', label: 'Local Vision v2 (non-AI)' }
  ]
};

const PLACEHOLDERS = {
  anthropic: 'sk-ant-...',
  openai: 'sk-...',
  openrouter: 'sk-or-...',
  groq: 'gsk_...',
  gemini: 'AIza...',
  local: 'No key needed for local mode'
};

function switchView(view) {
  const manual = document.getElementById('manual-view');
  const automate = document.getElementById('automate-view');
  const mBtn = document.getElementById('view-manual-btn');
  const aBtn = document.getElementById('view-automate-btn');
  const isAuto = view === 'automate';

  manual.classList.toggle('active', !isAuto);
  automate.classList.toggle('active', isAuto);
  mBtn.classList.toggle('active', !isAuto);
  aBtn.classList.toggle('active', isAuto);
  if (isAuto) window.scrollTo({ top: 0, behavior: 'smooth' });
}

function onProviderChange() {
  const p = document.getElementById('provider-select').value;
  const ms = document.getElementById('model-select');
  ms.innerHTML = '';
  MODELS[p].forEach(m => {
    const o = document.createElement('option');
    o.value = m.id; o.textContent = m.label;
    ms.appendChild(o);
  });
  document.getElementById('api-key').placeholder = PLACEHOLDERS[p];
  document.getElementById('api-key').disabled = p === 'local';
}

function onAutoProviderChange(phase) {
  const pSel = document.getElementById(`auto-${phase}-provider`);
  const mSel = document.getElementById(`auto-${phase}-model`);
  if (!pSel || !mSel) return;
  
  const provider = pSel.value;
  mSel.innerHTML = '';
  
  if (provider === 'local') {
    const opt = document.createElement('option');
    opt.value = 'local-v1';
    opt.textContent = (phase === 'clean') ? 'Local Detector' : 'Local NLP + OCR';
    mSel.appendChild(opt);
    return;
  }
  
  const models = MODELS[provider] || [];
  models.forEach(m => {
    const opt = document.createElement('option');
    opt.value = m.id;
    opt.textContent = m.label;
    mSel.appendChild(opt);
  });
}

function initAutomateProviderSelects() {
  ['clean', 'meta'].forEach(phase => {
    const pSel = document.getElementById(`auto-${phase}-provider`);
    if (!pSel) return;
    pSel.innerHTML = '';
    Object.keys(MODELS).forEach(key => {
      const op = document.createElement('option');
      op.value = key;
      op.textContent = key.charAt(0).toUpperCase() + key.slice(1);
      pSel.appendChild(op);
    });
    pSel.value = 'openai';
    onAutoProviderChange(phase);
  });
}

function setAutoStatus(text, state) {
  const el = document.getElementById('auto-drive-status');
  if (!el) return;
  el.textContent = text || '';
  el.className = 'drive-status' + (state ? ` ${state}` : '');
}

function autoLog(message, state) {
  const box = document.getElementById('auto-log');
  if (!box) return;
  const line = document.createElement('div');
  const stamp = new Date().toLocaleTimeString();
  line.className = 'auto-log-line' + (state ? ` ${state}` : '');
  line.textContent = `[${stamp}] ${message}`;
  box.appendChild(line);
  box.scrollTop = box.scrollHeight;
}

function resetAutoSteps() {
  for (let i = 1; i <= 5; i++) {
    const el = document.getElementById(`auto-step-${i}`);
    if (el) el.classList.remove('active', 'done', 'fail');
  }
}

function setAutoStepState(step, state) {
  const el = document.getElementById(`auto-step-${step}`);
  if (!el) return;
  el.classList.remove('active', 'done', 'fail');
  if (state) el.classList.add(state);
}

function getActiveGoogleClientId() {
  const autoClient = document.getElementById('auto-google-client-id')?.value.trim();
  if (autoClient) return autoClient;
  const mainClient = document.getElementById('google-client-id')?.value.trim();
  return mainClient;
}

function getActiveDriveFolderId() {
  const autoDest = document.getElementById('auto-destination-folder-id')?.value.trim();
  if (autoDest) return autoDest;
  return document.getElementById('drive-folder-id')?.value.trim();
}

function toggleKeyVis() {
  const inp = document.getElementById('api-key');
  const btn = document.querySelector('.toggle-vis');
  if (inp.type === 'password') { inp.type = 'text'; btn.textContent = 'Hide'; }
  else { inp.type = 'password'; btn.textContent = 'Show'; }
}

function setSens(v) {
  sensitivity = v;
  document.querySelectorAll('.sens-btn').forEach(b => b.classList.toggle('active', +b.dataset.v === v));
}

function handleDrop(e) {
  e.preventDefault();
  document.getElementById('drop-zone').classList.remove('drag-over');
  handleFile(e.dataTransfer.files);
}

function toPdfArray(input) {
  if (!input) return [];
  if (input instanceof FileList) return Array.from(input).filter(f => f.type === 'application/pdf' || /\.pdf$/i.test(f.name));
  if (Array.isArray(input)) return input.filter(f => f && (f.type === 'application/pdf' || /\.pdf$/i.test(f.name)));
  if (input instanceof File) return (input.type === 'application/pdf' || /\.pdf$/i.test(input.name)) ? [input] : [];
  return [];
}

function handleFile(input) {
  const picked = toPdfArray(input);
  if (!picked.length) return showError('Please upload one or more PDF files.');

  const keyed = new Map(pdfFiles.map(f => [f.name + '|' + f.size + '|' + f.lastModified, f]));
  picked.forEach(f => keyed.set(f.name + '|' + f.size + '|' + f.lastModified, f));
  pdfFiles = Array.from(keyed.values());

  const totalBytes = pdfFiles.reduce((acc, f) => acc + f.size, 0);
  document.getElementById('file-name-display').textContent = `${pdfFiles.length} PDF${pdfFiles.length > 1 ? 's' : ''} selected`;
  document.getElementById('file-size-display').textContent = `${(totalBytes / 1024).toFixed(0)} KB total`;

  const list = document.getElementById('file-list');
  list.innerHTML = '';
  pdfFiles.forEach(f => {
    const li = document.createElement('li');
    li.textContent = `• ${f.name}`;
    list.appendChild(li);
  });

  document.getElementById('file-info').style.display = 'flex';
  const first = pdfFiles[0];
  document.getElementById('output-name').value = pdfFiles.length === 1
    ? first.name.replace(/\.pdf$/i, '') + '_cleaned'
    : 'folio_batch_cleaned';
  document.getElementById('run-btn').disabled = false;
  hideError();
}

function resetAll() {
  pdfFiles = [];
  cleanResults = [];
  latestResult = null;
  currentPageResults = [];
  totalCreditsUsed = 0;
  ['file-info','progress-section','result-section'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = 'none';
  });
  document.getElementById('page-chips').innerHTML = '';
  document.getElementById('result-list').innerHTML = '';
  document.getElementById('file-list').innerHTML = '';
  document.getElementById('run-btn').disabled = true;
  document.getElementById('file-input').value = '';
  metadataSources = [];
  metadataRows = [];
  document.getElementById('metadata-file-input').value = '';
  document.getElementById('metadata-stage').style.display = 'none';
  document.getElementById('meta-papers').innerHTML = '';
  document.getElementById('meta-table-body').innerHTML = '<tr><td colspan="7" class="meta-empty">Run extraction to populate this table.</td></tr>';
  setMetaStatus('No papers selected yet.');
  const ds = document.getElementById('drive-status');
  if (ds) { ds.textContent = ''; ds.className = 'drive-status'; }
  hideError();
}

function showError(m) { const e = document.getElementById('error-box'); e.textContent = m; e.style.display = 'block'; }
function hideError() { document.getElementById('error-box').style.display = 'none'; }

function setProgress(pct, label, count) {
  document.getElementById('progress-fill').style.width = pct + '%';
  if (label) document.getElementById('progress-label').textContent = label;
  if (count !== undefined) document.getElementById('progress-count').textContent = count;
}

function addChip(i) {
  const el = document.createElement('span');
  el.className = 'chip pending';
  el.textContent = i + 1;
  el.id = 'chip-' + i;
  document.getElementById('page-chips').appendChild(el);
}

function updateChip(i, status) {
  const el = document.getElementById('chip-' + i);
  if (!el) return;
  el.className = 'chip ' + status;
  if (status === 'empty') el.textContent = '✕ ' + (i + 1);
  else if (status === 'content') el.textContent = '✓ ' + (i + 1);
  else el.textContent = i + 1;
}

function sanitizeBaseName(name) {
  return name.replace(/\.pdf$/i, '').replace(/[^a-zA-Z0-9-_]+/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '');
}

function getOutputName(file, fileIndex) {
  const base = document.getElementById('output-name').value.trim() || 'cleaned-document';
  if (pdfFiles.length === 1) return `${base}.pdf`;
  const src = sanitizeBaseName(file.name) || `file_${fileIndex + 1}`;
  return `${base}_${String(fileIndex + 1).padStart(2, '0')}_${src}.pdf`;
}

function buildResultCard(result, idx) {
  const wrap = document.createElement('div');
  wrap.className = 'result-item';

  const head = document.createElement('div');
  head.className = 'result-item-head';

  const nm = document.createElement('div');
  nm.className = 'result-item-name';
  nm.textContent = result.outputName;

  const meta = document.createElement('div');
  meta.className = 'result-item-meta';
  meta.textContent = `${result.kept}/${result.totalPages} kept · ${result.removed} removed · ${result.creditsUsed} credits`;

  head.appendChild(nm);
  head.appendChild(meta);
  wrap.appendChild(head);

  const actions = document.createElement('div');
  actions.className = 'result-actions';

  const dBtn = document.createElement('button');
  dBtn.className = 'btn-mini';
  dBtn.type = 'button';
  dBtn.textContent = 'Download';
  dBtn.addEventListener('click', () => downloadResultByIndex(idx));

  const uBtn = document.createElement('button');
  uBtn.className = 'btn-mini';
  uBtn.type = 'button';
  uBtn.textContent = 'Upload to Drive';
  uBtn.addEventListener('click', () => uploadResultByIndex(idx));

  actions.appendChild(dBtn);
  actions.appendChild(uBtn);
  wrap.appendChild(actions);

  return wrap;
}

function refreshResultCards() {
  const list = document.getElementById('result-list');
  list.innerHTML = '';
  cleanResults.forEach((r, idx) => {
    const card = buildResultCard(r, idx);
    card.style.animationDelay = `${Math.min(idx * 55, 380)}ms`;
    list.appendChild(card);
  });
}

function updateSummaryStats() {
  const total = cleanResults.reduce((a, r) => a + r.totalPages, 0);
  const removed = cleanResults.reduce((a, r) => a + r.removed, 0);
  const kept = cleanResults.reduce((a, r) => a + r.kept, 0);
  document.getElementById('stat-total').textContent = total;
  document.getElementById('stat-removed').textContent = removed;
  document.getElementById('stat-kept').textContent = kept;
  document.getElementById('stat-credits').textContent = totalCreditsUsed;
}

function setMetaStatus(text, state) {
  const el = document.getElementById('meta-status');
  if (!el) return;
  el.textContent = text;
  el.className = 'meta-status' + (state ? ` ${state}` : '');
}

function renderMetadataSourceList() {
  const holder = document.getElementById('meta-papers');
  if (!holder) return;
  holder.innerHTML = '';
  if (!metadataSources.length) {
    holder.innerHTML = '<div class="meta-paper-item">No files selected.</div>';
    return;
  }
  metadataSources.forEach((s, idx) => {
    const item = document.createElement('div');
    item.className = 'meta-paper-item';
    item.textContent = `${idx + 1}. ${s.name}`;
    holder.appendChild(item);
  });
}

function openMetadataStage(useCleaned) {
  const stage = document.getElementById('metadata-stage');
  if (!stage) return;
  stage.style.display = 'block';
  stage.classList.remove('reveal');
  void stage.offsetWidth;
  stage.classList.add('reveal');
  if (useCleaned) {
    if (!cleanResults.length) {
      setMetaStatus('No cleaned PDFs available yet. Run step 03 first.', 'err');
      stage.scrollIntoView({ behavior: 'smooth', block: 'start' });
      return;
    }
    metadataSources = cleanResults.map(r => ({ name: r.outputName, bytes: r.cleanPdfBytes }));
    renderMetadataSourceList();
    setMetaStatus(`${metadataSources.length} cleaned PDF(s) ready for extraction.`, 'ok');
  }
  stage.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function handleMetadataUpload(files) {
  const picked = toPdfArray(files);
  if (!picked.length) {
    setMetaStatus('Upload one or more PDF files.', 'err');
    return;
  }
  metadataSources = picked.map(f => ({ name: f.name, file: f }));
  const stage = document.getElementById('metadata-stage');
  if (!stage) return;
  stage.style.display = 'block';
  stage.classList.remove('reveal');
  void stage.offsetWidth;
  stage.classList.add('reveal');
  renderMetadataSourceList();
  setMetaStatus(`${metadataSources.length} uploaded PDF(s) ready for extraction.`, 'ok');
}

async function getSourceBytes(source) {
  if (source.bytes) {
    if (source.bytes instanceof Uint8Array) return source.bytes;
    if (source.bytes instanceof ArrayBuffer) return new Uint8Array(source.bytes);
  }
  if (source.file) {
    const ab = await source.file.arrayBuffer();
    return new Uint8Array(ab);
  }
  throw new Error('Invalid metadata source bytes.');
}

async function renderPdfPageToBase64(doc, pageNumber, scale = 1.6, quality = 0.8) {
  const page = await doc.getPage(pageNumber);
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement('canvas');
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
  return canvas.toDataURL('image/jpeg', quality).split(',')[1];
}

function parseFirstJsonObject(text) {
  const cleaned = (text || '').trim().replace(/^```json/i, '').replace(/^```/, '').replace(/```$/, '').trim();
  const start = cleaned.indexOf('{');
  const end = cleaned.lastIndexOf('}');
  if (start === -1 || end === -1 || end <= start) throw new Error('No JSON object found in model response.');
  return JSON.parse(cleaned.slice(start, end + 1));
}

function normalizeMetadata(obj, sourceFile) {
  const rawDepartments = Array.isArray(obj.departments)
    ? obj.departments
    : String(obj.departments || '').split(/[;,|]/g);

  const departments = rawDepartments
    .map(v => String(v).trim())
    .filter(Boolean)
    .join(', ');

  return {
    school: String(obj.school || '').trim(),
    departments,
    subjectCode: String(obj.subject_code || obj.subjectCode || '').trim(),
    subject: String(obj.subject || '').trim(),
    month: String(obj.month || '').trim(),
    year: String(obj.year || '').trim(),
    isFirstPage: Boolean(obj.is_first_page),
    confidence: Number(obj.confidence || 0),
    sourceFile
  };
}

async function extractQuestionPaperFields(b64, provider, model, apiKey) {
  const prompt = "Decide if this page is the FIRST PAGE of a question paper/exam paper. Indicators include school/institute name, department/program, subject code, subject name, exam month/year, instructions, time, max marks. Return ONLY JSON with keys: is_first_page (boolean), school, departments (array of strings), subject_code, subject, month, year, confidence (number 0-1). If not first page, return is_first_page false and empty fields.";

  async function openAiCompatibleMeta(endpoint, authHeader, tokenField, maxTokens) {
    const body = {
      model,
      messages: [{ role: 'user', content: [
        { type: 'image_url', image_url: { url: 'data:image/jpeg;base64,' + b64, detail: 'high' } },
        { type: 'text', text: prompt }
      ] }]
    };
    body[tokenField] = maxTokens;

    const res = await fetch(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        ...authHeader
      },
      body: JSON.stringify(body)
    });
    if (!res.ok) {
      const e = await res.json().catch(() => ({}));
      throw new Error(e.error?.message || `Metadata API error ${res.status}`);
    }
    const d = await res.json();
    return parseFirstJsonObject(d.choices?.[0]?.message?.content || '');
  }

  if (provider === 'anthropic') {
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({
        model,
        max_tokens: 220,
        messages: [{ role: 'user', content: [
          { type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: b64 } },
          { type: 'text', text: prompt }
        ] }]
      })
    });
    if (!res.ok) { const e = await res.json().catch(() => ({})); throw new Error(e.error?.message || 'Anthropic metadata error ' + res.status); }
    const d = await res.json();
    return parseFirstJsonObject(d.content?.[0]?.text || '');
  }

  if (provider === 'openai') {
    return openAiCompatibleMeta(
      'https://api.openai.com/v1/chat/completions',
      { 'Authorization': 'Bearer ' + apiKey },
      'max_completion_tokens',
      220
    );
  }

  if (provider === 'openrouter') {
    return openAiCompatibleMeta(
      'https://openrouter.ai/api/v1/chat/completions',
      {
        'Authorization': 'Bearer ' + apiKey,
        'HTTP-Referer': window.location.origin || 'http://localhost',
        'X-Title': 'Folio PDF Cleaner'
      },
      'max_tokens',
      220
    );
  }

  if (provider === 'groq') {
    return openAiCompatibleMeta(
      'https://api.groq.com/openai/v1/chat/completions',
      { 'Authorization': 'Bearer ' + apiKey },
      'max_tokens',
      220
    );
  }

  if (provider === 'gemini') {
    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        contents: [{ parts: [
          { inline_data: { mime_type: 'image/jpeg', data: b64 } },
          { text: prompt }
        ] }],
        generationConfig: { maxOutputTokens: 220 }
      })
    });
    if (!res.ok) { const e = await res.json().catch(() => ({})); throw new Error(e.error?.message || 'Gemini metadata error ' + res.status); }
    const d = await res.json();
    return parseFirstJsonObject(d.candidates?.[0]?.content?.parts?.[0]?.text || '');
  }

  throw new Error('Unknown provider for metadata extraction.');
}

function renderMetadataRows() {
  const body = document.getElementById('meta-table-body');
  if (!body) return;
  body.innerHTML = '';
  if (!metadataRows.length) {
    body.innerHTML = '<tr><td colspan="7" class="meta-empty">No rows extracted yet.</td></tr>';
    return;
  }
  metadataRows.forEach((r, idx) => {
    const tr = document.createElement('tr');
    tr.style.animationDelay = `${Math.min(idx * 45, 320)}ms`;
    tr.innerHTML = `
      <td>${r.school || '—'}</td>
      <td>${r.departments || '—'}</td>
      <td>${r.subjectCode || '—'}</td>
      <td>${r.subject || '—'}</td>
      <td>${r.month || '—'}</td>
      <td>${r.year || '—'}</td>
      <td>${r.sourceFile || '—'}</td>
    `;
    body.appendChild(tr);
  });
}

async function extractMetadataTable() {
  const provider = document.getElementById('provider-select').value;
  const model = document.getElementById('model-select').value;
  const apiKey = document.getElementById('api-key').value.trim();
  if (!metadataSources.length) return setMetaStatus('Select cleaned or uploaded PDFs first.', 'err');
  if (provider !== 'local' && !apiKey) return setMetaStatus('Enter API key in step 02 before extraction.', 'err');

  metadataRows = [];
  setMetaStatus(`Scanning all pages to detect question-paper starts across ${metadataSources.length} file(s)…`);

  for (let i = 0; i < metadataSources.length; i++) {
    const src = metadataSources[i];
    setMetaStatus(`Scanning ${i + 1}/${metadataSources.length}: ${src.name}`);
    const bytes = await getSourceBytes(src);
    const doc = await pdfjsLib.getDocument({ data: bytes }).promise;

    for (let p = 1; p <= doc.numPages; p++) {
      setMetaStatus(`Scanning ${src.name} · page ${p}/${doc.numPages}`);
      let raw;
      if (provider === 'local') {
        const page = await doc.getPage(p);
        raw = await extractQuestionPaperFieldsLocalFromPage(page, doc);
      } else {
        const b64 = await renderPdfPageToBase64(doc, p, 1.6, 0.8);
        raw = await extractQuestionPaperFields(b64, provider, model, apiKey);
      }
      const normalized = normalizeMetadata(raw, `${src.name} (p.${p})`);

      if (normalized.isFirstPage && normalized.confidence >= 0.5) {
        metadataRows.push(normalized);
      }
    }
  }

  renderMetadataRows();
  setMetaStatus(`Complete. Found ${metadataRows.length} detected first-page record(s).`, 'ok');
}

function downloadMetadataCsv() {
  if (!metadataRows.length) return setMetaStatus('No metadata rows to export yet.', 'err');
  const header = 'school,department,subject_code,subject,month,year,source_file';
  const rows = metadataRows.map(r => [
    r.school,
    r.departments,
    r.subjectCode,
    r.subject,
    r.month,
    r.year,
    r.sourceFile
  ].map(v => `"${String(v || '').replace(/"/g, '""')}"`).join(','));
  const csv = [header, ...rows].join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'question_paper_metadata.csv';
  a.click();
  URL.revokeObjectURL(a.href);
}

async function pageToBase64(pdfDoc, i) {
  const page = await pdfDoc.getPage(i + 1);
  const vp = page.getViewport({ scale: 1.5 });
  const canvas = document.createElement('canvas');
  canvas.width = vp.width; canvas.height = vp.height;
  await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;
  return canvas.toDataURL('image/jpeg', 0.75).split(',')[1];
}

async function getPageTextLines(page) {
  const textContent = await page.getTextContent();
  const grouped = new Map();
  textContent.items.forEach(it => {
    const t = String(it.str || '').replace(/\s+/g, ' ').trim();
    if (!t) return;
    const y = Math.round((it.transform?.[5] || 0) / 3);
    if (!grouped.has(y)) grouped.set(y, []);
    grouped.get(y).push(t);
  });

  return Array.from(grouped.keys())
    .sort((a, b) => b - a)
    .map(y => grouped.get(y).join(' ').replace(/\s+/g, ' ').trim())
    .filter(Boolean);
}

async function extractQuestionPaperFieldsLocalFromPage(page, pdfDoc) {
  let lines = await getPageTextLines(page);
  let text = lines.join(' \n ');
  
  // OCR Fallback if text layer is sparse
  if (text.trim().length < 50 && pdfDoc) {
    const pageNum = page.pageNumber;
    const b64 = await renderPdfPageToBase64(pdfDoc, pageNum, 2.0, 0.9);
    const ocrResult = await Tesseract.recognize('data:image/jpeg;base64,' + b64, 'eng');
    text = ocrResult.data.text;
    lines = text.split('\n');
    autoLog(`Local OCR performed on page ${pageNum}.`, 'ok');
  }

  return extractPaperMetadataLocal(text, lines);
}

function extractPaperMetadataLocal(text, lines) {
  if (!text) return { is_first_page: false, confidence: 0 };
  
  const doc = nlp(text);
  
  // Extract Schools/Universities using NLP
  let school = doc.organizations().filter(o => /university|college|school|institute|academy|polytechnic/i.test(o.text())).first().text();
  if (!school) {
    school = lines.find(l => /(university|college|institute|school|polytechnic|academy)/i.test(l)) || (lines[0] || '');
  }

  const deptLines = lines
    .filter(l => /(department|dept\.?|programme|program|faculty|school of)/i.test(l))
    .slice(0, 3)
    .map(l => l.replace(/^(department|dept\.?|faculty|school of)\s*(of)?\s*/i, '').trim())
    .filter(Boolean);

  const codeMatch = text.match(/\b[A-Z]{2,}[\s\/-]?[A-Z0-9]{2,}[\s\/-]?\d{1,4}[A-Z0-9]*\b/);
  
  let subjectLine = lines.find(l => /(subject|course|paper)\s*[:\-]/i.test(l)) || '';
  if (!subjectLine) {
    // Try to find a line that looks like a title (Title Case, no numbers)
    const candidates = lines.filter(l => l.length > 10 && l.length < 60 && !/\d/.test(l));
    subjectLine = candidates[0] || '';
  }

  const monthMatch = text.match(/\b(January|February|March|April|May|June|July|August|September|October|November|December)\b/i);
  const yearMatch = text.match(/\b(19|20)\d{2}\b/);

  let score = 0;
  if (school && school.length > 5) score += 1;
  if (deptLines.length) score += 1;
  if (codeMatch) score += 1;
  if (subjectLine) score += 1;
  if (monthMatch) score += 1;
  if (yearMatch) score += 1;
  if (/\b(time|max\s*marks|duration|instructions?)\b/i.test(text)) score += 1;

  return {
    is_first_page: score >= 3,
    school: school,
    departments: deptLines,
    subject_code: codeMatch ? codeMatch[0] : '',
    subject: subjectLine.replace(/^(subject|course|paper)\s*[:\-]?\s*/i, '').trim(),
    month: monthMatch ? monthMatch[0] : '',
    year: yearMatch ? yearMatch[0] : '',
    confidence: Math.min(1, score / 7)
  };
}

async function classifyPageLocal(pdfDoc, pageIndex) {
  const page = await pdfDoc.getPage(pageIndex + 1);
  const textContent = await page.getTextContent();
  const textChars = textContent.items
    .map(it => String(it.str || ''))
    .join('')
    .replace(/\s+/g, '')
    .length;

  const viewport = page.getViewport({ scale: 1.5 });
  const canvas = document.createElement('canvas');
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  const ctx = canvas.getContext('2d', { willReadFrequently: true });
  await page.render({ canvasContext: ctx, viewport }).promise;

  const { data, width, height } = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const marginX = Math.floor(width * 0.03);
  const marginY = Math.floor(height * 0.03);
  const x0 = Math.max(0, marginX);
  const y0 = Math.max(0, marginY);
  const x1 = Math.max(x0 + 1, width - marginX);
  const y1 = Math.max(y0 + 1, height - marginY);

  let sum = 0;
  let sumSq = 0;
  let pix = 0;
  for (let y = y0; y < y1; y++) {
    for (let x = x0; x < x1; x++) {
      const i = (y * width + x) * 4;
      const g = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
      sum += g;
      sumSq += g * g;
      pix++;
    }
  }

  const mean = sum / Math.max(1, pix);
  const variance = Math.max(0, sumSq / Math.max(1, pix) - mean * mean);
  const stdDev = Math.sqrt(variance);
  const threshold = Math.min(235, Math.max(185, mean - stdDev * 0.65));

  let darkCount = 0;
  let minX = x1;
  let minY = y1;
  let maxX = x0;
  let maxY = y0;

  const dw = Math.max(1, Math.floor((x1 - x0) / 2));
  const dh = Math.max(1, Math.floor((y1 - y0) / 2));
  const grid = new Uint8Array(dw * dh);

  for (let gy = 0; gy < dh; gy++) {
    for (let gx = 0; gx < dw; gx++) {
      const sx = x0 + gx * 2;
      const sy = y0 + gy * 2;
      const i = (sy * width + sx) * 4;
      const g = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
      if (g < threshold - 8) {
        grid[gy * dw + gx] = 1;
      }
    }
  }

  for (let y = y0; y < y1; y++) {
    for (let x = x0; x < x1; x++) {
      const i = (y * width + x) * 4;
      const g = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
      if (g < threshold) {
        darkCount++;
        if (x < minX) minX = x;
        if (y < minY) minY = y;
        if (x > maxX) maxX = x;
        if (y > maxY) maxY = y;
      }
    }
  }

  const visited = new Uint8Array(grid.length);
  let largestComp = 0;
  let significantInk = 0;
  const minComp = 26;
  const stack = [];

  for (let idx = 0; idx < grid.length; idx++) {
    if (!grid[idx] || visited[idx]) continue;
    visited[idx] = 1;
    stack.push(idx);
    let comp = 0;

    while (stack.length) {
      const cur = stack.pop();
      comp++;
      const x = cur % dw;
      const y = Math.floor(cur / dw);

      const nbs = [
        [x - 1, y], [x + 1, y], [x, y - 1], [x, y + 1],
        [x - 1, y - 1], [x + 1, y - 1], [x - 1, y + 1], [x + 1, y + 1]
      ];

      for (let n = 0; n < nbs.length; n++) {
        const nx = nbs[n][0];
        const ny = nbs[n][1];
        if (nx < 0 || ny < 0 || nx >= dw || ny >= dh) continue;
        const ni = ny * dw + nx;
        if (!grid[ni] || visited[ni]) continue;
        visited[ni] = 1;
        stack.push(ni);
      }
    }

    if (comp > largestComp) largestComp = comp;
    if (comp >= minComp) significantInk += comp;
  }

  const roiArea = Math.max(1, (x1 - x0) * (y1 - y0));
  const inkRatio = darkCount / roiArea;
  const bboxArea = darkCount ? (maxX - minX + 1) * (maxY - minY + 1) : 0;
  const bboxRatio = bboxArea / roiArea;

  if (sensitivity === 1) {
    if (textChars >= 1) return 'content';
    if (significantInk > 210 && inkRatio > 0.0012) return 'content';
    if (largestComp > 150 && bboxRatio > 0.009) return 'content';
    return 'empty';
  }

  if (sensitivity === 2) {
    if (textChars >= 5) return 'content';
    if (significantInk > 320 && inkRatio > 0.0019 && bboxRatio > 0.013) return 'content';
    if (largestComp > 280 && stdDev > 8) return 'content';
    return 'empty';
  }

  if (textChars >= 28) return 'content';
  if (significantInk > 760 && inkRatio > 0.0045 && bboxRatio > 0.03) return 'content';
  if (largestComp > 680 && stdDev > 13) return 'content';
  return 'empty';
}

const PROMPTS = {
  1: "Analyze this scanned document page. Reply ONLY with the single word 'empty' or 'content'. Reply 'empty' ONLY if the page is completely blank — pure white or near-white with at most faint scanner dust or grain, zero readable text, drawings, or marks. Reply 'content' for everything else.",
  2: "Analyze this scanned document page. Reply ONLY with the single word 'empty' or 'content'. Reply 'empty' if the page has no meaningful content — only scanner noise, dust, smudges, shadow at edges, or faint texture with no readable text, diagrams, or actual marks. Reply 'content' if there is any readable text, drawings, tables, stamps, or meaningful marks.",
  3: "Analyze this scanned document page. Reply ONLY with the single word 'empty' or 'content'. Reply 'empty' if the page has no significant content — includes pages with only a lone page number, an isolated header or footer line, faint watermarks, scanner noise, or sparse marks conveying no real information. Reply 'content' only if the page contains substantial readable content: paragraphs, data, diagrams, or important annotations."
};

async function classifyPage(b64, provider, model, apiKey) {
  const prompt = PROMPTS[sensitivity];

  async function openAiCompatibleRequest(endpoint, authHeader, tokenField, maxTokens) {
    const body = {
      model,
      messages: [{ role: 'user', content: [
        { type: 'image_url', image_url: { url: 'data:image/jpeg;base64,' + b64, detail: 'low' } },
        { type: 'text', text: prompt }
      ] }]
    };
    body[tokenField] = maxTokens;

    const res = await fetch(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        ...authHeader
      },
      body: JSON.stringify(body)
    });
    if (!res.ok) {
      const e = await res.json().catch(() => ({}));
      throw new Error(e.error?.message || `API error ${res.status}`);
    }
    const d = await res.json();
    return (d.choices?.[0]?.message?.content || '').toLowerCase().includes('empty') ? 'empty' : 'content';
  }

  if (provider === 'anthropic') {
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({ model, max_tokens: 10, messages: [{ role: 'user', content: [
        { type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: b64 } },
        { type: 'text', text: prompt }
      ]}]})
    });
    if (!res.ok) { const e = await res.json().catch(()=>({})); throw new Error(e.error?.message || 'Anthropic error ' + res.status); }
    const d = await res.json();
    return (d.content?.[0]?.text || '').toLowerCase().includes('empty') ? 'empty' : 'content';
  }

  if (provider === 'openai') {
    return openAiCompatibleRequest(
      'https://api.openai.com/v1/chat/completions',
      { 'Authorization': 'Bearer ' + apiKey },
      'max_completion_tokens',
      10
    );
  }

  if (provider === 'openrouter') {
    return openAiCompatibleRequest(
      'https://openrouter.ai/api/v1/chat/completions',
      {
        'Authorization': 'Bearer ' + apiKey,
        'HTTP-Referer': window.location.origin || 'http://localhost',
        'X-Title': 'Folio PDF Cleaner'
      },
      'max_tokens',
      10
    );
  }

  if (provider === 'groq') {
    return openAiCompatibleRequest(
      'https://api.groq.com/openai/v1/chat/completions',
      { 'Authorization': 'Bearer ' + apiKey },
      'max_tokens',
      10
    );
  }

  if (provider === 'gemini') {
    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ contents: [{ parts: [
        { inline_data: { mime_type: 'image/jpeg', data: b64 } },
        { text: prompt }
      ]}], generationConfig: { maxOutputTokens: 10 }})
    });
    if (!res.ok) { const e = await res.json().catch(()=>({})); throw new Error(e.error?.message || 'Gemini error ' + res.status); }
    const d = await res.json();
    return (d.candidates?.[0]?.content?.parts?.[0]?.text || '').toLowerCase().includes('empty') ? 'empty' : 'content';
  }

  throw new Error('Unknown provider');
}

async function classifyPageWithRetry(b64, provider, model, apiKey, retries = 1) {
  let lastErr = null;
  for (let attempt = 0; attempt <= retries; attempt++) {
    try {
      return await classifyPage(b64, provider, model, apiKey);
    } catch (err) {
      lastErr = err;
      if (attempt >= retries) break;
      await new Promise(resolve => setTimeout(resolve, 700 * (attempt + 1)));
    }
  }
  throw lastErr;
}

function syncDriveFolderId() {
  const el = document.getElementById('drive-folder-id');
  if (el) driveFolderId = el.value.trim();
}

function cancelProcessing() {
  if (!processing) return;
  stopRequested = true;
  setProgress(undefined, 'Cancelling… finishing current request safely.');
}

function setDriveAuthStatus(text, state) {
  const el = document.getElementById('drive-auth-status');
  if (!el) return;
  el.textContent = text || '';
  el.className = 'drive-status' + (state ? ` ${state}` : '');
}

function ensureGoogleTokenClient() {
  const clientId = getActiveGoogleClientId();
  if (!clientId) throw new Error('Enter your Google OAuth Client ID first.');
  if (!window.google || !google.accounts || !google.accounts.oauth2) {
    throw new Error('Google Identity Services is not available. Reload and try again.');
  }
  if (!googleTokenClient || activeGoogleClientId !== clientId) {
    googleTokenClient = google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: DRIVE_SCOPE,
      callback: () => {}
    });
    activeGoogleClientId = clientId;
  }
}

function requestGoogleToken(promptMode) {
  return new Promise((resolve, reject) => {
    try {
      ensureGoogleTokenClient();
      googleTokenClient.callback = (resp) => {
        if (resp && resp.error) return reject(new Error(resp.error));
        if (!resp || !resp.access_token) return reject(new Error('No access token returned by Google.'));
        googleAccessToken = resp.access_token;
        resolve(resp.access_token);
      };
      googleTokenClient.requestAccessToken({ prompt: promptMode });
    } catch (err) {
      reject(err);
    }
  });
}

async function connectGoogleAccount() {
  try {
    setDriveAuthStatus('Connecting to Google…');
    await requestGoogleToken('consent');
    setDriveAuthStatus('Google account connected.', 'ok');
  } catch (err) {
    setDriveAuthStatus('Google connect failed: ' + err.message, 'err');
  }
}

async function connectGoogleAccountAuto() {
  try {
    setAutoStatus('Connecting to Google…');
    await requestGoogleToken('consent');
    setAutoStatus('Google account connected.', 'ok');
    autoLog('Google account authorization successful.', 'ok');
  } catch (err) {
    setAutoStatus('Google connect failed: ' + err.message, 'err');
    autoLog('Google connect failed: ' + err.message, 'err');
  }
}

async function createDriveFolder() {
  const folderInput = document.getElementById('drive-folder-id');
  const desiredName = `Folio Cleaned PDFs ${new Date().toISOString().slice(0, 10)}`;
  try {
    setDriveAuthStatus('Creating Drive folder…');
    if (!googleAccessToken) await requestGoogleToken('consent');

    const res = await fetch('https://www.googleapis.com/drive/v3/files', {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + googleAccessToken,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        name: desiredName,
        mimeType: 'application/vnd.google-apps.folder'
      })
    });

    if (!res.ok) {
      const errBody = await res.json().catch(() => ({}));
      throw new Error(errBody.error?.message || `Drive error ${res.status}`);
    }

    const data = await res.json();
    driveFolderId = data.id;
    folderInput.value = driveFolderId;
    setDriveAuthStatus(`Folder ready: ${desiredName} (${driveFolderId})`, 'ok');
  } catch (err) {
    setDriveAuthStatus('Folder creation failed: ' + err.message, 'err');
  }
}

async function createAutoDestinationFolder() {
  const input = document.getElementById('auto-destination-folder-id');
  const desiredName = `Folio Automated Output ${new Date().toISOString().slice(0, 10)}`;
  try {
    setAutoStatus('Creating destination folder…');
    if (!googleAccessToken) await requestGoogleToken('consent');

    const res = await fetch('https://www.googleapis.com/drive/v3/files', {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + googleAccessToken,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        name: desiredName,
        mimeType: 'application/vnd.google-apps.folder'
      })
    });

    if (!res.ok) {
      const errBody = await res.json().catch(() => ({}));
      throw new Error(errBody.error?.message || `Drive error ${res.status}`);
    }

    const data = await res.json();
    input.value = data.id;
    setAutoStatus(`Destination ready: ${desiredName} (${data.id})`, 'ok');
    autoLog(`Destination folder created (${data.id}).`, 'ok');
  } catch (err) {
    setAutoStatus('Destination create failed: ' + err.message, 'err');
    autoLog('Destination folder creation failed: ' + err.message, 'err');
  }
}

async function listPdfsInDriveFolder(folderId) {
  if (!googleAccessToken) await requestGoogleToken('consent');
  const q = `'${folderId}' in parents and mimeType='application/pdf' and trashed=false`;
  const url = `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(q)}&fields=files(id,name,size)&pageSize=1000`;
  const res = await fetch(url, {
    headers: { 'Authorization': 'Bearer ' + googleAccessToken }
  });
  if (!res.ok) {
    const errBody = await res.json().catch(() => ({}));
    throw new Error(errBody.error?.message || `Drive list error ${res.status}`);
  }
  const data = await res.json();
  return data.files || [];
}

async function downloadDriveFileBytes(fileId) {
  if (!googleAccessToken) await requestGoogleToken('consent');
  let res = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`, {
    headers: { 'Authorization': 'Bearer ' + googleAccessToken }
  });

  if (res.status === 401) {
    await requestGoogleToken('');
    res = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`, {
      headers: { 'Authorization': 'Bearer ' + googleAccessToken }
    });
  }

  if (!res.ok) {
    let msg = `Drive download error ${res.status}`;
    try {
      const t = await res.text();
      if (t) msg = t;
    } catch (_) {}
    throw new Error(msg);
  }

  return new Uint8Array(await res.arrayBuffer());
}

async function previewSourceFolder() {
  const sourceFolderId = document.getElementById('auto-source-folder-id').value.trim();
  if (!sourceFolderId) {
    setAutoStatus('Enter source folder ID first.', 'err');
    return;
  }
  try {
    autoLog('Reading source folder contents…');
    const files = await listPdfsInDriveFolder(sourceFolderId);
    if (!files.length) {
      autoLog('No PDF files found in source folder.', 'err');
      setAutoStatus('No PDFs found in source folder.', 'err');
      return;
    }
    autoLog(`Found ${files.length} PDF file(s) in source folder.`, 'ok');
    setAutoStatus(`Source folder ready with ${files.length} PDF(s).`, 'ok');
  } catch (err) {
    autoLog('Source folder preview failed: ' + err.message, 'err');
    setAutoStatus('Source preview failed: ' + err.message, 'err');
  }
}

function renderAutomateRows() {
  const body = document.getElementById('auto-table-body');
  if (!body) return;
  body.innerHTML = '';
  if (!automateResults.length) {
    body.innerHTML = '<tr><td colspan="7" class="meta-empty">Run automation to populate this table.</td></tr>';
    return;
  }
  automateResults.forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.sourceFile}</td>
      <td>${row.cleanedFile}</td>
      <td><span class="match-badge ${row.excelMatch.toLowerCase().replace(' ', '-')}">${row.excelMatch}</span></td>
      <td>${row.title || 'Untitled'}</td>
      <td>${row.totalPages}</td>
      <td>${row.removed}</td>
      <td>${row.kept}</td>
    `;
    body.appendChild(tr);
  });
}


function downloadAutomateCsv() {
  if (!automateCsvText) return;
  const blob = new Blob([automateCsvText], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'folio_automated_titles.csv';
  a.click();
  URL.revokeObjectURL(a.href);
}

function downloadAllAutomatePdfs() {
  if (!automateResults.length) return;
  automateResults.forEach((r, idx) => {
    setTimeout(() => downloadBytes(r.cleanPdfBytes, r.cleanedFile), idx * 140);
  });
}

function cancelAutomatePipeline() {
  if (!automateRunning) return;
  automateCancelRequested = true;
  autoLog('Cancel requested. Finishing current request safely…');
}

async function extractTitleFromPdfBytes(cleanBytes, provider, model, apiKey) {
  const doc = await pdfjsLib.getDocument({ data: cleanBytes }).promise;
  const page = await doc.getPage(1);
  
  // Try text layer first
  let lines = await getPageTextLines(page);
  let fullText = lines.join(' ');
  
  // If text layer is too thin, try OCR
  if (fullText.trim().length < 50) {
    const b64 = await renderPdfPageToBase64(doc, 1, 2.0, 0.9);
    const ocrResult = await Tesseract.recognize('data:image/jpeg;base64,' + b64, 'eng');
    fullText = ocrResult.data.text;
    lines = fullText.split('\n');
    autoLog('Local OCR performed on page 1.', 'ok');
  }

  if (provider === 'local') {
    return extractEntitiesLocal(fullText, lines);
  }

  // AI Vision Provider fallback
  const b64 = await renderPdfPageToBase64(doc, 1, 1.6, 0.8);
  const prompt = "Extract the best concise title of this question paper. Return ONLY JSON with key 'title'. If unclear, return title as empty string.";

  async function openAiTitle(endpoint, authHeader, tokenField) {
    const body = {
      model,
      messages: [{ role: 'user', content: [
        { type: 'image_url', image_url: { url: 'data:image/jpeg;base64,' + b64, detail: 'high' } },
        { type: 'text', text: prompt }
      ] }]
    };
    body[tokenField] = 120;
    const res = await fetch(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        ...authHeader
      },
      body: JSON.stringify(body)
    });
    if (!res.ok) {
      const e = await res.json().catch(() => ({}));
      throw new Error(e.error?.message || `Title API error ${res.status}`);
    }
    const d = await res.json();
    return parseFirstJsonObject(d.choices?.[0]?.message?.content || '').title || '';
  }

  if (provider === 'anthropic') {
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model,
        max_tokens: 120,
        messages: [{ role: 'user', content: [
          { type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: b64 } },
          { type: 'text', text: prompt }
        ] }]
      })
    });
    if (!res.ok) {
      const e = await res.json().catch(() => ({}));
      throw new Error(e.error?.message || `Title API error ${res.status}`);
    }
    const d = await res.json();
    return parseFirstJsonObject(d.content?.[0]?.text || '').title || '';
  }

  if (provider === 'openai') {
    return openAiTitle('https://api.openai.com/v1/chat/completions', { 'Authorization': 'Bearer ' + apiKey }, 'max_completion_tokens');
  }

  if (provider === 'openrouter') {
    return openAiTitle('https://openrouter.ai/api/v1/chat/completions', {
      'Authorization': 'Bearer ' + apiKey,
      'HTTP-Referer': window.location.origin || 'http://localhost',
      'X-Title': 'Folio PDF Cleaner'
    }, 'max_tokens');
  }

  if (provider === 'groq') {
    return openAiTitle('https://api.groq.com/openai/v1/chat/completions', { 'Authorization': 'Bearer ' + apiKey }, 'max_tokens');
  }

  if (provider === 'gemini') {
    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        contents: [{ parts: [
          { inline_data: { mime_type: 'image/jpeg', data: b64 } },
          { text: prompt }
        ] }],
        generationConfig: { maxOutputTokens: 120 }
      })
    });
    if (!res.ok) {
      const e = await res.json().catch(() => ({}));
      throw new Error(e.error?.message || `Title API error ${res.status}`);
    }
    const d = await res.json();
    return parseFirstJsonObject(d.candidates?.[0]?.content?.parts?.[0]?.text || '').title || '';
  }

  throw new Error('Unknown provider for title extraction.');
}

async function processPdfBytes(fileName, bytes, provider, model, apiKey, fileIndex, fileCount) {
  currentPageResults = [];
  document.getElementById('page-chips').innerHTML = '';
  setProgress(0, `Loading ${fileName} (${fileIndex + 1}/${fileCount})…`, '');

  const renderBytes = new Uint8Array(bytes);
  const pdfDoc = await pdfjsLib.getDocument({ data: renderBytes }).promise;
  const total = pdfDoc.numPages;

  for (let i = 0; i < total; i++) addChip(i);
  setProgress(5, `Analyzing ${fileName} with ${model}…`, `0 / ${total}`);

  for (let i = 0; i < total; i++) {
    if (automateCancelRequested) throw new Error('Automation canceled by user.');
    updateChip(i, 'checking');
    setProgress(5 + Math.round((i / total) * 78), `File ${fileIndex + 1}/${fileCount} · page ${i + 1}/${total}…`, `${i} / ${total}`);
    let result;
    if (provider === 'local') {
      result = await classifyPageLocal(pdfDoc, i);
    } else {
      const b64 = await pageToBase64(pdfDoc, i);
      result = await classifyPageWithRetry(b64, provider, model, apiKey, 1);
    }
    currentPageResults.push(result);
    updateChip(i, result);
  }

  setProgress(88, `Rebuilding ${fileName}…`, `${total} / ${total}`);

  const srcPdf = await PDFLib.PDFDocument.load(bytes);
  const newPdf = await PDFLib.PDFDocument.create();
  const keepIdx = currentPageResults.map((r, i) => r === 'content' ? i : -1).filter(i => i >= 0);

  if (keepIdx.length) {
    const copied = await newPdf.copyPages(srcPdf, keepIdx);
    copied.forEach(p => newPdf.addPage(p));
  } else {
    newPdf.addPage([595, 842]);
  }

  const cleanPdfBytes = await newPdf.save();
  const kept = currentPageResults.filter(r => r === 'content').length;
  const removed = currentPageResults.filter(r => r === 'empty').length;
  setProgress(100, `Finished ${fileName}.`, `${total} / ${total}`);

  return {
    sourceFileName: fileName,
    outputName: `${sanitizeBaseName(fileName)}_cleaned.pdf`,
    totalPages: total,
    kept,
    removed,
    creditsUsed: total,
    cleanPdfBytes
  };
}

async function runAutomatePipeline() {
  const sourceFolderId = document.getElementById('auto-source-folder-id').value.trim();
  const destinationFolderId = getActiveDriveFolderId();
  
  const cleanProvider = document.getElementById('auto-clean-provider').value;
  const cleanModel = document.getElementById('auto-clean-model').value;
  const metaProvider = document.getElementById('auto-meta-provider').value;
  const metaModel = document.getElementById('auto-meta-model').value;
  
  const apiKey = document.getElementById('auto-api-key').value.trim() || document.getElementById('api-key').value.trim();
  const useFolderLogic = document.getElementById('auto-folder-logic').checked;
  const skipClean = document.getElementById('auto-skip-clean').checked;
  const skipMeta = document.getElementById('auto-skip-meta').checked;

  console.log('Automate Start:', { skipClean, skipMeta, cleanProvider, metaProvider });

  if (!sourceFolderId) return setAutoStatus('Provide source folder ID.', 'err');
  if (!destinationFolderId) return setAutoStatus('Provide destination folder ID.', 'err');
  
  const needsKey = (!skipClean && cleanProvider !== 'local') || (!skipMeta && metaProvider !== 'local');
  if (needsKey && !apiKey) return setAutoStatus('Provide API key for LLM phases.', 'err');

  automateRunning = true;
  automateCancelRequested = false;
  automateResults = [];
  automateCsvText = '';
  document.getElementById('auto-table-body').innerHTML = '<tr><td colspan="7" class="meta-empty">Pipeline running…</td></tr>';
  resetAutoSteps();
  document.getElementById('auto-log').innerHTML = '';
  autoLog('Automation started.');

  try {
    setAutoStepState(1, 'active');
    autoLog('Step 1/5: Reading source Drive folder…');
    const sourceFiles = await listPdfsInDriveFolder(sourceFolderId);
    if (!sourceFiles.length) throw new Error('No PDF files found.');
    autoLog(`Found ${sourceFiles.length} file(s).`, 'ok');
    setAutoStepState(1, 'done');

    setAutoStepState(2, 'active');
    const cleanFilesCache = new Map(); // Store [sourceName -> {bytes, removed, kept}]

    if (skipClean) {
      autoLog('Phase 1 (Cleaning) is SKIPPED.', 'ok');
      for (let i = 0; i < sourceFiles.length; i++) {
        const f = sourceFiles[i];
        // We don't download yet if we also skip meta, to make it faster
        cleanFilesCache.set(f.name, {
          id: f.id,
          bytes: null, 
          removed: 0,
          kept: 'Skipped',
          totalPages: '?'
        });
      }
      setAutoStepState(2, 'done');
    } else {
      autoLog(`Phase 1 (Cleaning) active using ${cleanProvider}…`);
      for (let i = 0; i < sourceFiles.length; i++) {
        if (automateCancelRequested) throw new Error('Canceled.');
        const f = sourceFiles[i];
        autoLog(`Processing ${i + 1}/${sourceFiles.length}: ${f.name}`);
        const bytes = await downloadDriveFileBytes(f.id);
        const cleaned = await processPdfBytes(f.name, bytes, cleanProvider, cleanModel, apiKey, i, sourceFiles.length);
        cleanFilesCache.set(f.name, {
          bytes: cleaned.cleanPdfBytes,
          removed: cleaned.removed,
          kept: cleaned.kept,
          totalPages: cleaned.totalPages
        });
        autoLog(`Cleaned ${f.name}.`, 'ok');
      }
      setAutoStepState(2, 'done');
    }

    setAutoStepState(3, 'active');
    const detectedPapers = [];

    if (skipMeta) {
      autoLog('Phase 2 (Metadata) is SKIPPED.', 'ok');
      for (const sourceName of cleanFilesCache.keys()) {
        detectedPapers.push({
          sourceFile: sourceName,
          pageIndex: 1,
          school: '', dept: '', subjectCode: '', subject: 'Extraction Skipped',
          month: '', year: '', excelMatch: 'No Match', excelRow: null, driveFileId: ''
        });
      }
      renderAutomateRowsFlattened(detectedPapers);
      setAutoStepState(3, 'done');
    } else {
      autoLog(`Phase 2 (Metadata) active using ${metaProvider}…`);
      for (const [sourceName, cache] of cleanFilesCache.entries()) {
        if (automateCancelRequested) throw new Error('Canceled.');
        
        // Download now if it was skipped in Step 2
        if (!cache.bytes) {
          autoLog(`Downloading ${sourceName}…`);
          cache.bytes = await downloadDriveFileBytes(cache.id);
        }

        const doc = await pdfjsLib.getDocument({ data: cache.bytes }).promise;
        autoLog(`Scanning every page of ${sourceName} (${doc.numPages})…`);

        for (let p = 1; p <= doc.numPages; p++) {
          const page = await doc.getPage(p);
          const meta = await extractQuestionPaperFieldsLocalFromPage(page, doc);
          
          if (meta.is_first_page || p === 1) {
            let finalMeta = meta;
            if (metaProvider !== 'local') {
              const b64 = await renderPdfPageToBase64(doc, p, 1.6, 0.8);
              finalMeta = await extractQuestionPaperFields(b64, metaProvider, metaModel, apiKey);
            }

            const paper = {
              sourceFile: sourceName,
              pageIndex: p,
              school: finalMeta.school || meta.school || '',
              dept: (finalMeta.departments || meta.departments || []).join(', '),
              subjectCode: finalMeta.subject_code || meta.subject_code || '',
              subject: finalMeta.subject || meta.subject || '',
              month: finalMeta.month || meta.month || '',
              year: finalMeta.year || meta.year || '',
              excelMatch: 'No Match',
              excelRow: null,
              driveFileId: ''
            };

            const matchData = getMappedExcelRow(sourceName, paper.subject || paper.subjectCode);
            if (matchData) {
              paper.excelMatch = 'Matched';
              paper.excelRow = matchData.row;
            }

            detectedPapers.push(paper);
            renderAutomateRowsFlattened(detectedPapers);
          }
        }
      }
      setAutoStepState(3, 'done');
    }
    automateResults = detectedPapers;

    setAutoStepState(4, 'active');
    autoLog('Step 4/5: Building folders and uploading…');
    
    // Group detected papers by source file for efficient uploading
    const papersByFile = new Map();
    automateResults.forEach(p => {
      if (!papersByFile.has(p.sourceFile)) papersByFile.set(p.sourceFile, []);
      papersByFile.get(p.sourceFile).push(p);
    });

    for (const [sourceName, papers] of papersByFile.entries()) {
      if (automateCancelRequested) throw new Error('Canceled.');
      
      const cache = cleanFilesCache.get(sourceName);
      if (!cache.bytes) {
        autoLog(`Downloading ${sourceName} for upload…`);
        cache.bytes = await downloadDriveFileBytes(cache.id);
      }
      
      let targetFolderId = destinationFolderId;
      
      // If folder logic is ON, we use the first paper's metadata for the folder
      // (Usually a bundle belongs to one dept/year)
      const p1 = papers[0];
      if (useFolderLogic) {
        const year = p1.year || 'Unknown-Year';
        const dept = p1.dept || 'Unknown-Dept';
        const code = p1.subjectCode || 'Unknown-Code';
        
        const yearId = await getOrCreateSubfolder(destinationFolderId, year);
        const deptId = await getOrCreateSubfolder(yearId, dept);
        targetFolderId = await getOrCreateSubfolder(deptId, code);
      }

      const outputName = `${sourceName.replace(/\.pdf$/i, '')}_cleaned.pdf`;
      autoLog(`Uploading ${outputName}…`);
      const up = await uploadBytesToDriveSmart(cache.bytes, outputName, targetFolderId);
      
      // Update all papers from this source with the same file ID
      papers.forEach(p => p.driveFileId = up.id);
      renderAutomateRowsFlattened(automateResults);
    }
    setAutoStepState(4, 'done');

    setAutoStepState(5, 'active');
    autoLog('Step 5/5: Generating Comprehensive Audit Report…');
    automateCsvText = buildAuditReportExtended();
    const csvBlob = new Blob([automateCsvText], { type: 'text/csv' });
    const csvBytes = new Uint8Array(await csvBlob.arrayBuffer());
    await uploadBytesToDriveSmart(csvBytes, `audit_report_${new Date().getTime()}.csv`, destinationFolderId, 'text/csv');
    setAutoStepState(5, 'done');

    setAutoStatus('Automation complete.', 'ok');
    autoLog('Pipeline complete. All detected papers logged.', 'ok');
  } catch (err) {
    autoLog(`Failed: ${err.message}`, 'err');
    setAutoStatus(`Failed: ${err.message}`, 'err');
  } finally { automateRunning = false; }
}

function renderAutomateRowsFlattened(papers) {
  const body = document.getElementById('auto-table-body');
  if (!body) return;
  body.innerHTML = '';
  if (!papers.length) {
    body.innerHTML = '<tr><td colspan="9" class="meta-empty">No papers detected yet.</td></tr>';
    return;
  }
  papers.forEach(p => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${p.sourceFile} (p.${p.pageIndex})</td>
      <td><span class="match-badge ${p.excelMatch.toLowerCase().replace(' ', '-')}">${p.excelMatch}</span></td>
      <td>${p.subjectCode || '—'}</td>
      <td>${p.subject || '—'}</td>
      <td>${p.school || '—'}</td>
      <td>${p.dept || '—'}</td>
      <td>${p.month || '—'}</td>
      <td>${p.year || '—'}</td>
      <td><a href="https://drive.google.com/file/d/${p.driveFileId}/view" target="_blank">View</a></td>
    `;
    body.appendChild(tr);
  });
}

function buildAuditReportExtended() {
  const header = 'source_file,start_page,school,department,subject_code,subject,month,year,excel_match,drive_file_id';
  const rows = automateResults.map(p => [
    p.sourceFile,
    p.pageIndex,
    p.school,
    p.dept,
    p.subjectCode,
    p.subject,
    p.month,
    p.year,
    p.excelMatch,
    p.driveFileId
  ].map(v => `"${String(v || '').replace(/"/g, '""')}"`).join(','));
  return [header, ...rows].join('\n');
}

async function uploadBytesToDrive(bytes, outputName, mimeType = 'application/pdf') {
  syncDriveFolderId();
  if (!driveFolderId) throw new Error('Set a Drive folder ID or create a folder first.');
  if (!googleAccessToken) await requestGoogleToken('consent');

  const metadata = { name: outputName, mimeType, parents: [driveFolderId] };
  const boundary = 'folio_' + Math.random().toString(16).slice(2);
  const pre = `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n`;
  const mid = `--${boundary}\r\nContent-Type: ${mimeType}\r\n\r\n`;
  const end = `\r\n--${boundary}--`;

  const body = new Blob([pre, mid, bytes, end], { type: 'multipart/related; boundary=' + boundary });

  let res = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink', {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + googleAccessToken,
      'Content-Type': `multipart/related; boundary=${boundary}`
    },
    body
  });

  if (res.status === 401) {
    await requestGoogleToken('');
    res = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink', {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + googleAccessToken,
        'Content-Type': `multipart/related; boundary=${boundary}`
      },
      body
    });
  }

  if (!res.ok) {
    const errBody = await res.json().catch(() => ({}));
    throw new Error(errBody.error?.message || `Drive upload error ${res.status}`);
  }

  return res.json();
}

async function uploadResultToDrive(result) {
  const statusEl = document.getElementById('drive-status');
  if (!statusEl) return;
  statusEl.className = 'drive-status';
  statusEl.textContent = `Uploading ${result.outputName}…`;

  try {
    const up = await uploadBytesToDrive(result.cleanPdfBytes, result.outputName);
    result.driveUploaded = true;
    result.driveUploadId = up.id;
    statusEl.className = 'drive-status ok';
    statusEl.textContent = `Uploaded to Drive folder (${up.name}).`;
    return up;
  } catch (err) {
    result.driveUploaded = false;
    statusEl.className = 'drive-status err';
    statusEl.textContent = 'Drive upload failed: ' + err.message;
    throw err;
  }
}

async function processOnePdf(file, provider, model, apiKey, fileIndex, fileCount) {
  currentPageResults = [];
  document.getElementById('page-chips').innerHTML = '';
  setProgress(0, `Loading ${file.name} (${fileIndex + 1}/${fileCount})…`, '');

  const renderBuffer = await file.arrayBuffer();
  const renderBytes = new Uint8Array(renderBuffer);
  const pdfDoc = await pdfjsLib.getDocument({ data: renderBytes }).promise;
  const total = pdfDoc.numPages;

  for (let i = 0; i < total; i++) addChip(i);
  setProgress(5, `Analyzing ${file.name} with ${model}…`, `0 / ${total}`);

  for (let i = 0; i < total; i++) {
    if (stopRequested) throw new Error('Processing canceled by user.');
    updateChip(i, 'checking');
    setProgress(5 + Math.round((i / total) * 78), `File ${fileIndex + 1}/${fileCount} · page ${i + 1}/${total}…`, `${i} / ${total}`);
    let result;
    if (provider === 'local') {
      result = await classifyPageLocal(pdfDoc, i);
    } else {
      const b64 = await pageToBase64(pdfDoc, i);
      result = await classifyPageWithRetry(b64, provider, model, apiKey, 1);
    }
    currentPageResults.push(result);
    updateChip(i, result);
  }

  setProgress(88, `Rebuilding ${file.name}…`, `${total} / ${total}`);

  const buildBuffer = await file.arrayBuffer();
  const srcPdf = await PDFLib.PDFDocument.load(buildBuffer);
  const newPdf = await PDFLib.PDFDocument.create();
  const keepIdx = currentPageResults.map((r, i) => r === 'content' ? i : -1).filter(i => i >= 0);

  if (keepIdx.length) {
    const copied = await newPdf.copyPages(srcPdf, keepIdx);
    copied.forEach(p => newPdf.addPage(p));
  } else {
    newPdf.addPage([595, 842]);
  }

  const cleanPdfBytes = await newPdf.save();
  const kept = currentPageResults.filter(r => r === 'content').length;
  const removed = currentPageResults.filter(r => r === 'empty').length;
  const creditsUsed = total;

  setProgress(100, `Finished ${file.name}.`, `${total} / ${total}`);

  return {
    sourceFileName: file.name,
    outputName: getOutputName(file, fileIndex),
    totalPages: total,
    kept,
    removed,
    creditsUsed,
    cleanPdfBytes,
    driveUploaded: false,
    driveUploadId: ''
  };
}

async function runAgent() {
  const apiKey = document.getElementById('api-key').value.trim();
  if (!pdfFiles.length) return showError('Please upload at least one PDF first.');

  syncDriveFolderId();
  const shouldAutoUpload = document.getElementById('auto-upload-toggle').checked;
  if (shouldAutoUpload && !driveFolderId) {
    return showError('Auto-upload is ON. Please set or create a Google Drive folder first.');
  }

  const provider = document.getElementById('provider-select').value;
  const model = document.getElementById('model-select').value;
  if (provider !== 'local' && !apiKey) return showError('Please enter your API key.');

  hideError();
  cleanResults = [];
  latestResult = null;
  totalCreditsUsed = 0;
  stopRequested = false;
  processing = true;
  document.getElementById('run-btn').disabled = true;
  document.getElementById('cancel-btn').style.display = 'block';
  document.getElementById('result-section').style.display = 'none';
  document.getElementById('progress-section').style.display = 'block';
  document.getElementById('page-chips').innerHTML = '';
  document.getElementById('result-list').innerHTML = '';
  setProgress(0, 'Loading document…', '');

  try {
    for (let i = 0; i < pdfFiles.length; i++) {
      if (stopRequested) throw new Error('Processing canceled by user.');
      const fileResult = await processOnePdf(pdfFiles[i], provider, model, apiKey, i, pdfFiles.length);
      cleanResults.push(fileResult);
      latestResult = fileResult;
      totalCreditsUsed += fileResult.creditsUsed;
      updateSummaryStats();
      refreshResultCards();

      if (shouldAutoUpload) {
        await uploadResultToDrive(fileResult);
      }
    }

    setProgress(100, `Complete. ${cleanResults.length} file(s) processed.`, `${cleanResults.length} / ${cleanResults.length}`);
    document.getElementById('result-section').style.display = 'block';
    document.getElementById('result-section').scrollIntoView({ behavior: 'smooth', block: 'nearest' });

  } catch (err) {
    showError('Error: ' + err.message);
    document.getElementById('progress-section').style.display = 'none';
  } finally {
    processing = false;
    document.getElementById('run-btn').disabled = false;
    document.getElementById('cancel-btn').style.display = 'none';
  }
}

function downloadBytes(bytes, filename) {
  const blob = new Blob([bytes], { type: 'application/pdf' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

function downloadResultByIndex(idx) {
  const r = cleanResults[idx];
  if (!r) return;
  latestResult = r;
  downloadBytes(r.cleanPdfBytes, r.outputName);
}

async function uploadResultByIndex(idx) {
  const r = cleanResults[idx];
  if (!r) return;
  latestResult = r;
  await uploadResultToDrive(r);
}

function downloadPDF() {
  if (!latestResult) return;
  downloadBytes(latestResult.cleanPdfBytes, latestResult.outputName);
}

function downloadAllPDFs() {
  if (!cleanResults.length) return;
  cleanResults.forEach((r, i) => {
    setTimeout(() => downloadBytes(r.cleanPdfBytes, r.outputName), i * 150);
  });
}

function downloadProcessingReport() {
  if (!cleanResults.length) return;
  const header = 'source_file,output_file,total_pages,kept,removed,credits_used,drive_uploaded,drive_file_id';
  const rows = cleanResults.map(r => [
    r.sourceFileName,
    r.outputName,
    r.totalPages,
    r.kept,
    r.removed,
    r.creditsUsed,
    r.driveUploaded ? 'yes' : 'no',
    r.driveUploadId || ''
  ].map(v => `"${String(v).replace(/"/g, '""')}"`).join(','));
  const csv = [header, ...rows].join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'folio_processing_report.csv';
  a.click();
  URL.revokeObjectURL(a.href);
}

async function uploadLatestToDrive() {
  if (!latestResult) return;
  await uploadResultToDrive(latestResult);
}

// EXCEL INGESTION
function handleExcelUpload(files) {
  if (!files || !files.length) return;
  const file = files[0];
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    excelData = XLSX.utils.sheet_to_json(worksheet);
    
    if (excelData.length > 0) {
      excelColumns = Object.keys(excelData[0]);
      populateMappingSelects();
      document.getElementById('excel-mapping-zone').style.display = 'block';
      autoLog(`Excel loaded: ${excelData.length} rows found.`, 'ok');
    }
  };
  reader.readAsArrayBuffer(file);
}

function populateMappingSelects() {
  ['map-title', 'map-code', 'map-year', 'map-dept'].forEach(id => {
    const sel = document.getElementById(id);
    sel.innerHTML = '<option value="">-- Skip --</option>';
    excelColumns.forEach(col => {
      const opt = document.createElement('option');
      opt.value = col;
      opt.textContent = col;
      
      // Auto-matching logic for common names
      const low = col.toLowerCase();
      if (id === 'map-title' && (low.includes('title') || low.includes('paper') || low.includes('subject'))) opt.selected = true;
      if (id === 'map-code' && (low.includes('code'))) opt.selected = true;
      if (id === 'map-year' && (low.includes('year'))) opt.selected = true;
      if (id === 'map-dept' && (low.includes('dept') || low.includes('department'))) opt.selected = true;
      
      sel.appendChild(opt);
    });
  });
}

function getMappedExcelRow(fileName, detectedTitle) {
  if (!excelData.length) return null;
  
  const titleCol = document.getElementById('map-title').value;
  const codeCol = document.getElementById('map-code').value;
  const yearCol = document.getElementById('map-year').value;
  const deptCol = document.getElementById('map-dept').value;

  let bestMatch = null;
  let bestScore = -1;

  const cleanFile = fileName.toLowerCase().replace(/\.pdf$/, '');
  const cleanDetected = (detectedTitle || '').toLowerCase();

  excelData.forEach(row => {
    let score = 0;
    const rowTitle = String(row[titleCol] || '').toLowerCase();
    const rowCode = String(row[codeCol] || '').toLowerCase();
    
    // Exact match on filename
    if (rowTitle && cleanFile.includes(rowTitle)) score += 10;
    if (rowCode && cleanFile.includes(rowCode)) score += 10;
    
    // Fuzzy match on detected title
    if (cleanDetected && rowTitle) {
      if (cleanDetected.includes(rowTitle) || rowTitle.includes(cleanDetected)) score += 5;
      // Simple word overlap
      const detectedWords = new Set(cleanDetected.split(/\W+/));
      const rowWords = rowTitle.split(/\W+/);
      let overlap = 0;
      rowWords.forEach(w => { if (w.length > 3 && detectedWords.has(w)) overlap++; });
      score += overlap;
    }

    if (score > bestScore) {
      bestScore = score;
      bestMatch = row;
    }
  });

  // Threshold
  if (bestScore > 3) return { row: bestMatch, score: bestScore };
  return null;
}

// FOLDER INTELLIGENCE
async function getOrCreateSubfolder(parentFolderId, folderName) {
  if (!folderName) return parentFolderId;
  
  if (!googleAccessToken) await requestGoogleToken('');
  
  const q = `'${parentFolderId}' in parents and name='${folderName.replace(/'/g, "\\'")}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
  const url = `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(q)}&fields=files(id,name)`;
  
  const res = await fetch(url, { headers: { 'Authorization': 'Bearer ' + googleAccessToken } });
  const data = await res.json();
  
  if (data.files && data.files.length > 0) {
    return data.files[0].id;
  }
  
  // Create if not found
  const createRes = await fetch('https://www.googleapis.com/drive/v3/files', {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + googleAccessToken,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      name: folderName,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [parentFolderId]
    })
  });
  const created = await createRes.json();
  return created.id;
}

async function uploadBytesToDriveSmart(bytes, outputName, folderId, mimeType = 'application/pdf') {
  if (!googleAccessToken) await requestGoogleToken('');

  const metadata = { name: outputName, mimeType, parents: [folderId] };
  const boundary = 'folio_' + Math.random().toString(16).slice(2);
  const pre = `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n`;
  const mid = `--${boundary}\r\nContent-Type: ${mimeType}\r\n\r\n`;
  const end = `\r\n--${boundary}--`;

  const body = new Blob([pre, mid, bytes, end], { type: 'multipart/related; boundary=' + boundary });

  const res = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name', {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + googleAccessToken,
      'Content-Type': `multipart/related; boundary=${boundary}`
    },
    body
  });

  return res.json();
}

function extractEntitiesLocal(text, lines) {
  const meta = extractPaperMetadataLocal(text, lines);
  return meta.subject || meta.school || lines[0] || '';
}

// Initialize components
document.addEventListener('DOMContentLoaded', () => {
  onProviderChange();
  onAutoProviderChange('clean');
  onAutoProviderChange('meta');
});
