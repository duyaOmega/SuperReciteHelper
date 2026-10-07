'use strict';

/*
 * SuperReciteHelper 网页版前端。
 * Python 核心逻辑（解析/抽题/记录）由 Pyodide 在浏览器中运行，
 * 本文件负责 UI 状态机、键盘交互与各弹窗。
 */

const PYODIDE_VERSION = '0.28.3';
const PYODIDE_BASES = [
  `https://cdn.jsdelivr.net/pyodide/v${PYODIDE_VERSION}/full/`,
  `https://fastly.jsdelivr.net/pyodide/v${PYODIDE_VERSION}/full/`,
  `https://unpkg.com/pyodide@${PYODIDE_VERSION}/`,
];
const THEME_KEY = 'srh_web_theme';

const $ = (id) => document.getElementById(id);

const el = {
  loading: $('loading'),
  loadingText: $('loading-text'),
  loadingSub: $('loading-sub'),
  loadingRetry: $('loading-retry'),
  bankTitle: $('bank-title'),
  statsBar: $('stats-bar'),
  themeBtn: $('theme-btn'),
  qNumber: $('q-number'),
  qType: $('q-type'),
  welcome: $('welcome'),
  qText: $('q-text'),
  options: $('options'),
  result: $('result'),
  history: $('history'),
  submitBtn: $('submit-btn'),
  nextBtn: $('next-btn'),
  statsBtn: $('stats-btn'),
  resetBtn: $('reset-btn'),
  modalRoot: $('modal-root'),
  toastRoot: $('toast-root'),
};

let pyodide = null;

const S = {
  ready: false,
  total: 0,
  current: null,
  selected: new Set(),
  submitted: false,
  revealed: false,
  revealedAnswer: '',
};

/* ---------- 主题 ---------- */

function currentTheme() {
  return document.documentElement.dataset.theme || 'light';
}

function applyTheme(name) {
  document.documentElement.dataset.theme = name;
  el.themeBtn.textContent = name === 'dark' ? '☀️' : '🌙';
  el.themeBtn.title = name === 'dark' ? '切换到浅色模式' : '切换到深色模式';
}

function initTheme() {
  const saved = localStorage.getItem(THEME_KEY);
  if (saved === 'light' || saved === 'dark') {
    applyTheme(saved);
  } else {
    applyTheme(matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
  }
}

el.themeBtn.addEventListener('click', () => {
  const next = currentTheme() === 'dark' ? 'light' : 'dark';
  applyTheme(next);
  localStorage.setItem(THEME_KEY, next);
});

matchMedia('(prefers-color-scheme: dark)').addEventListener('change', (e) => {
  if (!localStorage.getItem(THEME_KEY)) {
    applyTheme(e.matches ? 'dark' : 'light');
  }
});

/* ---------- 提示 ---------- */

function toast(msg, isError = false) {
  const node = document.createElement('div');
  node.className = 'toast' + (isError ? ' err' : '');
  node.textContent = msg;
  el.toastRoot.appendChild(node);
  setTimeout(() => node.remove(), 2800);
}

function errText(err) {
  const raw = String((err && err.message) || err || '未知错误');
  const lines = raw.split('\n').map((s) => s.trim()).filter(Boolean);
  return lines.length ? lines[lines.length - 1] : raw;
}

/* ---------- Pyodide 引导 ---------- */

function setLoadingText(text) {
  el.loadingText.textContent = text;
}

function showLoadingError(err) {
  el.loading.classList.add('error');
  el.loadingText.textContent = '加载失败：' + errText(err);
  el.loadingSub.textContent = '请检查网络连接后重试（需要访问 CDN 下载 Python 运行时）。';
  el.loadingRetry.hidden = false;
}

el.loadingRetry.addEventListener('click', () => window.location.reload());

function injectScript(src) {
  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = src;
    script.onload = resolve;
    script.onerror = () => reject(new Error(`无法加载 ${src}`));
    document.head.appendChild(script);
  });
}

async function loadRuntime() {
  let lastError = null;
  for (const base of PYODIDE_BASES) {
    try {
      const host = new URL(base).host;
      setLoadingText(`正在加载 Python 运行时…（${host}）`);
      await injectScript(base + 'pyodide.js');
      return await loadPyodide({ indexURL: base });
    } catch (err) {
      lastError = err;
      console.warn('Pyodide 镜像加载失败，尝试下一个：', base, err);
    }
  }
  throw lastError || new Error('Pyodide 加载失败');
}

async function fetchText(url) {
  const resp = await fetch(url, { cache: 'no-cache' });
  if (!resp.ok) {
    throw new Error(`下载失败：${url}（HTTP ${resp.status}）`);
  }
  return resp.text();
}

async function preparePythonFiles() {
  const FS = pyodide.FS;
  FS.mkdirTree('/home/pyodide/srh/core');
  FS.writeFile('/home/pyodide/srh/__init__.py', '');
  FS.writeFile('/home/pyodide/srh/core/__init__.py', '');
  for (const name of ['parser.py', 'question_bank.py', 'session.py']) {
    FS.writeFile(`/home/pyodide/srh/core/${name}`, await fetchText(`pycore/${name}`));
  }
  FS.writeFile('/home/pyodide/web_glue.py', await fetchText('py/web_glue.py'));
}

function pyCall(name, ...args) {
  const fn = pyodide.globals.get(name);
  if (!fn) {
    throw new Error(`内部函数缺失：${name}`);
  }
  try {
    return fn(...args);
  } finally {
    fn.destroy();
  }
}

function pyJson(name, ...args) {
  try {
    return JSON.parse(pyCall(name, ...args));
  } catch (err) {
    console.error(err);
    toast('内部错误：' + errText(err), true);
    return null;
  }
}

/* ---------- 顶部进度 ---------- */

function updateHeader() {
  const data = pyJson('api_overall');
  if (!data) return;
  el.statsBar.textContent =
    `已做 ${data.attempted}/${data.total} | 总答题 ${data.attempts} | 正确率 ${data.accuracy.toFixed(1)}%`;
}

function updateHistory(record) {
  if (record.attempts > 0) {
    const rate = (record.errors / record.attempts) * 100;
    el.history.textContent =
      `历史记录：答过 ${record.attempts} 次，错误 ${record.errors} 次，错误率 ${rate.toFixed(0)}%`;
  } else {
    el.history.textContent = '历史记录：首次作答';
  }
  el.history.hidden = false;
}

/* ---------- 刷题流程 ---------- */

const OBJECTIVE_TYPES = new Set(['single', 'multi', 'judge']);

function clearResult() {
  el.result.hidden = true;
  el.result.textContent = '';
  el.result.classList.remove('ok', 'err');
}

function setResult(text, kind) {
  el.result.hidden = false;
  el.result.textContent = text;
  el.result.classList.remove('ok', 'err');
  if (kind) el.result.classList.add(kind);
}

function clearOptions() {
  el.options.textContent = '';
}

function renderQuestion(q) {
  el.welcome.hidden = true;
  el.qText.hidden = false;
  el.qText.textContent = q.display_text;
  el.qNumber.textContent = `第 ${q.id} 题`;
  el.qType.hidden = false;
  el.qType.textContent = `【${q.type_name}题】`;
  clearOptions();
  clearResult();
  updateHistory(q.record);

  if (OBJECTIVE_TYPES.has(q.type)) {
    for (const key of Object.keys(q.options).sort()) {
      const btn = document.createElement('button');
      btn.className = 'option-btn';
      btn.textContent = `${key}. ${q.options[key]}`;
      btn.addEventListener('click', () => toggleOption(key));
      el.options.appendChild(btn);
    }
    el.submitBtn.textContent = '提交答案';
    el.submitBtn.disabled = false;
  } else {
    setResult('请先自己作答，点击「显示正确答案」后再进行自评。');
    el.submitBtn.textContent = '显示正确答案';
    el.submitBtn.disabled = false;
  }
}

function optionButtons() {
  return Array.from(el.options.querySelectorAll('.option-btn'));
}

function toggleOption(key) {
  if (S.submitted || !S.current) return;
  const type = S.current.type;
  if (type === 'single' || type === 'judge') {
    S.selected = new Set([key]);
  } else if (S.selected.has(key)) {
    S.selected.delete(key);
  } else {
    S.selected.add(key);
  }
  for (const btn of optionButtons()) {
    const k = btn.textContent.slice(0, 1);
    btn.classList.toggle('selected', S.selected.has(k));
    btn.classList.remove('correct', 'wrong', 'dimmed');
  }
}

function nextQuestion() {
  const q = pyJson('api_pick_next');
  if (!q) return;
  S.current = q;
  S.selected = new Set();
  S.submitted = false;
  S.revealed = false;
  S.revealedAnswer = '';
  renderQuestion(q);
}

function submitObjective() {
  if (!S.selected.size) {
    toast('请先选择一个答案！');
    return;
  }
  if (!S.current.has_answer) {
    toast('本题未识别出标准答案，暂无法自动判分。');
    return;
  }
  const res = pyJson('api_submit_objective', JSON.stringify(Array.from(S.selected)));
  if (!res) return;
  S.submitted = true;

  const correct = new Set(res.answer.split(''));
  for (const btn of optionButtons()) {
    const k = btn.textContent.slice(0, 1);
    btn.disabled = true;
    if (correct.has(k)) {
      btn.classList.add('correct');
      btn.classList.remove('selected', 'wrong', 'dimmed');
    } else if (S.selected.has(k)) {
      btn.classList.add('wrong');
      btn.classList.remove('selected', 'correct', 'dimmed');
    } else {
      btn.classList.add('dimmed');
      btn.classList.remove('selected', 'correct', 'wrong');
    }
  }

  if (res.correct) {
    setResult('✓ 回答正确！', 'ok');
  } else {
    setResult(`✗ 回答错误！正确答案是：${res.answer}`, 'err');
  }
  updateHistory(res.record);
  updateHeader();
  el.submitBtn.disabled = true;
}

function revealAnswer() {
  const res = pyJson('api_reveal_answer');
  if (!res) return;
  S.revealed = true;
  S.revealedAnswer = res.answer;
  setResult(`参考答案：${res.answer}\n\n请根据你的作答进行自评：`);
  clearOptions();

  const okBtn = document.createElement('button');
  okBtn.className = 'btn btn-ok';
  okBtn.textContent = '我答对了';
  okBtn.addEventListener('click', () => selfEval(true));

  const badBtn = document.createElement('button');
  badBtn.className = 'btn btn-ghost btn-danger';
  badBtn.textContent = '我答错了';
  badBtn.addEventListener('click', () => selfEval(false));

  el.options.appendChild(okBtn);
  el.options.appendChild(badBtn);
  el.submitBtn.disabled = true;
}

function selfEval(isCorrect) {
  if (S.submitted || !S.current) return;
  const res = pyJson('api_record_subjective', isCorrect);
  if (!res) return;
  S.submitted = true;
  setResult(
    `参考答案：${S.revealedAnswer}\n\n${isCorrect ? '✓ 已记录：你判定本题答对。' : '✗ 已记录：你判定本题答错。'}`,
    isCorrect ? 'ok' : 'err',
  );
  for (const btn of el.options.querySelectorAll('button')) {
    btn.disabled = true;
  }
  updateHistory(res.record);
  updateHeader();
}

function resetRecords() {
  if (!window.confirm('确定要重置所有错误记录吗？\n此操作不可撤销！')) return;
  const res = pyJson('api_reset_records');
  if (!res) return;
  updateHeader();
  if (S.current) {
    el.history.textContent = '历史记录：首次作答';
  }
  toast('所有记录已重置。');
}

/* ---------- 键盘 ---------- */

function normalizeKeyChar(ch) {
  const code = ch.codePointAt(0);
  if (code >= 0xff21 && code <= 0xff3a) return String.fromCharCode(code - 0xfee0);
  if (code >= 0xff41 && code <= 0xff5a) return String.fromCharCode(code - 0xfee0);
  return ch.toUpperCase();
}

function hasModal() {
  return el.modalRoot.querySelector('.modal-mask') !== null;
}

function closeTopModal() {
  const masks = el.modalRoot.querySelectorAll('.modal-mask');
  const top = masks[masks.length - 1];
  if (top && top._close) top._close();
}

document.addEventListener('keydown', (event) => {
  if (!S.ready) return;
  if (hasModal()) {
    if (event.key === 'Escape') closeTopModal();
    return;
  }
  if (!S.current && event.key !== 'Enter') return;

  if (event.key === 'Enter') {
    event.preventDefault();
    onEnter();
    return;
  }
  if (event.key.length !== 1) return;

  const ch = normalizeKeyChar(event.key);
  const type = S.current.type;

  if (OBJECTIVE_TYPES.has(type) && !S.submitted && /^[A-H]$/.test(ch) && S.current.options[ch]) {
    toggleOption(ch);
    return;
  }

  if ((type === 'blank' || type === 'short') && S.revealed && !S.submitted) {
    if (ch === 'T' || ch === 'Y' || ch === '对') {
      selfEval(true);
    } else if (ch === 'F' || ch === 'N' || ch === '错') {
      selfEval(false);
    }
  }
});

function onEnter() {
  if (!S.current) {
    nextQuestion();
    return;
  }
  if (S.submitted) {
    nextQuestion();
    return;
  }
  if (OBJECTIVE_TYPES.has(S.current.type)) {
    submitObjective();
    return;
  }
  if (!S.revealed) {
    revealAnswer();
  } else {
    toast('请输入 t（答对）或 f（答错）后回车自评。');
  }
}

/* ---------- 弹窗 ---------- */

function openModal(title) {
  const mask = document.createElement('div');
  mask.className = 'modal-mask';
  const panel = document.createElement('div');
  panel.className = 'modal-panel';

  const head = document.createElement('div');
  head.className = 'modal-head';
  const headText = document.createElement('span');
  headText.textContent = title;
  const closeBtn = document.createElement('button');
  closeBtn.className = 'icon-btn';
  closeBtn.textContent = '✕';
  closeBtn.title = '关闭';
  head.append(headText, closeBtn);

  panel.appendChild(head);
  mask.appendChild(panel);
  el.modalRoot.appendChild(mask);

  const close = () => {
    mask.remove();
  };
  closeBtn.addEventListener('click', close);
  mask.addEventListener('click', (event) => {
    if (event.target === mask) close();
  });
  mask._close = close;
  return { mask, panel, close };
}

function sortRows(rows, mode) {
  const copy = rows.slice();
  if (mode === '按错误次数') {
    copy.sort((a, b) => b.errors - a.errors || b.attempts - a.attempts || a.id - b.id);
  } else if (mode === '按错误率') {
    copy.sort((a, b) => b.error_rate - a.error_rate || b.attempts - a.attempts || a.id - b.id);
  } else if (mode === '按题号') {
    copy.sort((a, b) => a.id - b.id);
  } else {
    copy.sort((a, b) => b.attempts - a.attempts || b.errors - a.errors || a.id - b.id);
  }
  return copy;
}

function openStats() {
  const data = pyJson('api_stats');
  if (!data) return;

  const { panel } = openModal('考频统计');

  const summary = document.createElement('div');
  summary.className = 'modal-summary';
  const s = data.summary;
  summary.textContent =
    `总题数 ${s.total} | 已作答 ${s.attempted} | 总作答次数 ${s.attempts} | 总错误次数 ${s.errors} | 总错误率 ${s.error_rate.toFixed(1)}%`;
  panel.appendChild(summary);

  if (data.duplicate_groups > 0) {
    const dup = document.createElement('div');
    dup.className = 'modal-summary dup';
    dup.textContent =
      `重复题检测：${data.duplicate_groups} 组（共 ${data.duplicate_questions} 题），系统抽题时会自动规避短时间重复同组。`;
    panel.appendChild(dup);
  }

  const controls = document.createElement('div');
  controls.className = 'modal-controls';
  const sortLabel = document.createElement('label');
  sortLabel.textContent = '排序方式：';
  const sortSelect = document.createElement('select');
  const sortOptions = ['按作答次数', '按错误次数', '按错误率', '按题号'];
  for (const name of sortOptions) {
    const opt = document.createElement('option');
    opt.textContent = name;
    sortSelect.appendChild(opt);
  }
  sortLabel.appendChild(sortSelect);

  const onlyLabel = document.createElement('label');
  const onlyCheck = document.createElement('input');
  onlyCheck.type = 'checkbox';
  onlyCheck.checked = true;
  onlyLabel.append(onlyCheck, document.createTextNode('仅看已作答题目'));
  controls.append(sortLabel, onlyLabel);
  panel.appendChild(controls);

  const body = document.createElement('div');
  body.className = 'modal-body';
  const table = document.createElement('table');
  table.className = 'stats-table';
  const colNames = ['排名', '题号', '题型', '作答次数', '错误次数', '错误率', '题干预览'];
  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  for (const name of colNames) {
    const th = document.createElement('th');
    th.textContent = name;
    if (name !== '题干预览') th.classList.add('c');
    headRow.appendChild(th);
  }
  thead.appendChild(headRow);
  const tbody = document.createElement('tbody');
  table.append(thead, tbody);
  body.appendChild(table);
  panel.appendChild(body);

  const render = () => {
    const mode = sortSelect.value;
    const onlyAttempted = onlyCheck.checked;
    let rows = data.rows;
    if (onlyAttempted) rows = rows.filter((r) => r.attempts > 0);
    rows = sortRows(rows, mode);

    tbody.textContent = '';
    rows.forEach((row, index) => {
      const tr = document.createElement('tr');
      tr.className = 'clickable';
      const preview = row.text.length > 90 ? row.text.slice(0, 90) + '...' : row.text;
      const rateCls = row.error_rate >= 50 ? 'rate-bad' : row.error_rate > 0 ? 'rate-warn' : '';
      const cells = [
        ['c', String(index + 1)],
        ['c', String(row.id)],
        ['c', row.type],
        ['c', String(row.attempts)],
        ['c', row.errors > 0 ? `<span class="rate-bad">${row.errors}</span>` : '0'],
        ['c', `<span class="${rateCls}">${row.error_rate.toFixed(0)}%</span>`],
        ['preview', ''],
      ];
      for (const [cls, content] of cells) {
        const td = document.createElement('td');
        td.className = cls;
        if (cls === 'preview') {
          td.textContent = preview;
        } else {
          td.innerHTML = content;
        }
        tr.appendChild(td);
      }
      tr.addEventListener('click', () => openDetail(row.id));
      tbody.appendChild(tr);
    });

    if (!rows.length) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = colNames.length;
      td.className = 'c';
      td.textContent = onlyAttempted ? '暂无已作答题目。' : '暂无数据。';
      tr.appendChild(td);
      tbody.appendChild(tr);
    }
  };

  sortSelect.addEventListener('change', render);
  onlyCheck.addEventListener('change', render);
  render();
}

function openDetail(qid) {
  const data = pyJson('api_detail', String(qid));
  if (!data) return;

  const { panel } = openModal(`题目详情 - 第 ${data.id} 题`);

  const meta = document.createElement('div');
  meta.className = 'detail-meta';
  const r = data.record;
  const rate = r.attempts > 0 ? ((r.errors / r.attempts) * 100).toFixed(0) : '0';
  meta.textContent =
    `题型：${data.type} | 作答 ${r.attempts} 次  错误 ${r.errors} 次  错误率 ${rate}%`;
  panel.appendChild(meta);

  const body = document.createElement('div');
  body.className = 'detail-body';
  let text = `题干：\n${data.text}\n\n`;
  if (data.options.length) {
    text += '选项：\n' + data.options.map((o) => `${o.key}. ${o.text}`).join('\n') + '\n\n';
  }
  text += `答案：\n${data.answer}`;
  body.textContent = text;
  panel.appendChild(body);
}

/* ---------- 启动 ---------- */

async function boot() {
  initTheme();
  setLoadingText('正在加载题库数据…');
  const bank = await fetchText('data/bank.json').then((t) => JSON.parse(t));

  pyodide = await loadRuntime();

  setLoadingText('正在写入核心模块…');
  await preparePythonFiles();
  pyodide.runPython(
    "import sys\nif '/home/pyodide' not in sys.path:\n    sys.path.insert(0, '/home/pyodide')",
  );
  pyodide.runPython('from web_glue import *');

  setLoadingText('正在初始化题库…');
  const info = pyJson('api_init', JSON.stringify(bank));
  if (!info) throw new Error('题库初始化失败');

  S.ready = true;
  S.total = info.total;
  el.bankTitle.textContent = info.title;

  let welcomeText =
    `题库已加载 ${info.total} 道题目。\n\n` +
    '点击「下一题」开始刷题！\n\n' +
    '系统会根据你的错误率自动加权抽题，错得越多的题越容易被抽到哦～';
  if (info.duplicate_groups > 0) {
    welcomeText +=
      `\n\n检测到重复题组 ${info.duplicate_groups} 组（共 ${info.duplicate_questions} 题），` +
      '系统会自动尽量避免短时间重复抽到同组题。';
  }
  el.welcome.textContent = welcomeText;
  el.qNumber.textContent = '欢迎使用刷题工具！';

  el.nextBtn.disabled = false;
  el.statsBtn.disabled = false;
  el.resetBtn.disabled = false;
  updateHeader();

  el.loading.remove();
}

el.nextBtn.addEventListener('click', nextQuestion);
el.submitBtn.addEventListener('click', () => {
  if (!S.current || S.submitted) return;
  onEnter();
});
el.statsBtn.addEventListener('click', openStats);
el.resetBtn.addEventListener('click', resetRecords);

boot().catch((err) => {
  console.error(err);
  showLoadingError(err);
});
