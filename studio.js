const el = id => document.getElementById(id);
const labels = [
  ['lecture', '今日讲解', '01_今日讲解.docx'],
  ['exercises', '练习与答案', '02_练习与答案.docx'],
  ['vocab', '单词卡', '03_单词卡.xlsx'],
  ['listening', '听力材料', '04_听力材料.docx'],
  ['review', '滚动复习', '05_昨日复习.docx'],
];
const states = { preview: '首日预览', available: '已开放', scheduled: '待开放', pending: '资料生成中' };
const state = { course: null, statuses: {}, active: null, today: null, tab: labels[0][0], filter: 'all', done: new Set(), request: 0 };
const html = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
const idFor = n => 'Day' + String(n).padStart(2, '0');
const isReady = status => status === 'available' || status === 'preview';
const showDay = d => d.replaceAll('-', '.');
const zipFor = day => day.id + '_N3学习包.zip';
let speechSession = 0;

const publishedLessons = new Map();
const downloadUrls = new Map();
const mime = {"docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "zip": "application/zip"};
let planRequest;
let readyIds = new Set();

async function readData(path, fresh = false) {
  const url = new URL(path, location.href);
  if (fresh) url.searchParams.set('v', String(Math.floor(Date.now() / 30000)));
  const response = await fetch(url, { cache: fresh ? 'no-store' : 'default', credentials: 'omit' });
  if (!response.ok) throw new Error('HTTP ' + response.status);
  return response.json();
}

function coursePlan() {
  if (!planRequest) planRequest = readData('course.json');
  return planRequest;
}

function statusFor(day, now) {
  if (Date.parse(day.opens) > now) {
    return readyIds.has(day.id) && day.day === 1 ? 'preview' : 'scheduled';
  }
  return readyIds.has(day.id) ? 'available' : 'pending';
}

async function get(path) {
  const now = Date.now();
  if (path === '/api/course') return coursePlan();
  if (path === '/api/status') {
    const [course, manifest] = await Promise.all([coursePlan(), readData('available.json', true)]);
    readyIds = new Set(manifest.ready);
    return { serverTime: new Date(now).toISOString(), days: Object.fromEntries(course.plan.map(day => [day.id, statusFor(day, now)])) };
  }
  const id = path.slice('/api/lesson/'.length);
  const meta = (await coursePlan()).plan.find(day => day.id === id);
  if (!meta) throw new Error('Day unavailable');
  const status = statusFor(meta, now);
  if (!isReady(status)) return { meta, status };
  if (!publishedLessons.has(id)) {
    const bundle = await readData(id + '.json');
    publishedLessons.set(id, bundle);
  }
  return { meta, status, content: publishedLessons.get(id).content };
}

function urlFor(day, name) {
  const key = day.id + '/' + name;
  if (!downloadUrls.has(key)) {
    const data = publishedLessons.get(day.id)?.files[name];
    if (!data) return '#';
    const bytes = Uint8Array.from(atob(data), c => c.charCodeAt(0));
    downloadUrls.set(key, URL.createObjectURL(new Blob([bytes], { type: mime[name.split('.').at(-1)] })));
  }
  return downloadUrls.get(key);
}
function setMarkup(id, markup) {
  el(id).innerHTML = markup;
}
function safeDone() {
  try {
    const saved = JSON.parse(localStorage.getItem('n3-90-day-progress-v1') || '[]');
    return new Set(Array.isArray(saved) ? saved.filter(n => Number.isInteger(n) && n > 0 && n <= 90) : []);
  } catch {
    return new Set();
  }
}
function saveDone() {
  try {
    localStorage.setItem('n3-90-day-progress-v1', JSON.stringify([...state.done]));
    return true;
  } catch {
    return false;
  }
}

function courseDay(number) {
  return state.course?.plan[number - 1];
}
function directoryGroup(d) {
  if (d.kind === 'mock') return '模拟测试';
  if (d.kind === 'review') return '综合复习';
  return d.lessons.at(-1) <= 16 ? '中级上册' : '中级下册';
}
function lessonLine(d) {
  if (d.kind === 'learn') {
    return d.main.map(n => {
      const l = state.course.lessons[n - 1];
      return '第 ' + n + ' 课 ' + l.conversation + '・' + l.reading;
    }).join(' / ');
  }
  return d.lessons.length ? '第 ' + d.lessons[0] + '–' + d.lessons.at(-1) + ' 课' : d.goal;
}

function renderDirectory(scrollToActive = false) {
  const search = el('search').value.trim().toLocaleLowerCase();
  const tokens = search.split(/\s+/).filter(Boolean);
  let lastGroup = '';
  const output = [];
  for (const d of state.course.plan) {
    const status = state.statuses[d.id] ?? 'scheduled';
    if (state.filter === 'released' && !isReady(status)) continue;
    const keywords = [d.id.toLowerCase(), 'day ' + d.day, '第' + d.day + '天', d.date, d.title, d.phase, d.focus,
      ...d.lessons.flatMap(n => ['第' + n + '课', state.course.lessons[n - 1].conversation, state.course.lessons[n - 1].reading]),
    ].join(' ').toLowerCase();
    if (tokens.some(token => !keywords.includes(token))) continue;
    const group = directoryGroup(d);
    if (group !== lastGroup) {
      output.push('<p class="day-group">' + html(group) + '</p>');
      lastGroup = group;
    }
    const current = d.day === state.active?.day;
    const completed = state.done.has(d.day);
    const dotClass = completed ? 'done' : isReady(status) ? 'ready' : status === 'pending' ? 'pending' : '';
    output.push('<button type="button" class="day-link" data-day="' + d.day + '" aria-current="' + current + '" aria-label="' +
      html(d.id + ' ' + d.title + ' · ' + (completed ? '已完成' : states[status])) + '">' +
      '<span class="day-number">' + d.id.slice(3) + '</span><span><span class="day-name">' + html(d.title) +
      '</span><span class="day-label">' + html(d.date.slice(5).replace('-', '/')) + ' · ' + html(completed ? '已打卡' : states[status]) +
      '</span></span><span class="dot ' + dotClass + '" aria-hidden="true"></span></button>');
  }
  setMarkup('day-list', output.join('') || '<p class="empty">没有匹配的课程。</p>');
  if (scrollToActive) {
    const active = el('day-list').querySelector('[aria-current="true"]');
    if (active) el('day-list').scrollTop = Math.max(0, active.offsetTop - el('day-list').offsetTop - 110);
  }
}

function renderHero(day, status) {
  const ready = isReady(status);
  const meta = [day.minutes + ' 分钟'];
  if (day.kind === 'learn') meta.push(day.words + ' 个新词', day.grammar + ' 个句型');
  else meta.push(day.phase);
  setMarkup('breadcrumb', day.id);
  const action = ready
    ? '<a class="cta" href="' + urlFor(day, zipFor(day)) + '" download="' + zipFor(day) + '">下载学习包 <span aria-hidden="true">↓</span></a>'
    : '<span class="cta disabled">' + html(states[status] ?? '正在载入') + '</span>';
  setMarkup('hero-content', '<div class="hero-eyebrow">' + day.id.toUpperCase() + ' <span aria-hidden="true">/</span> ' + showDay(day.date) +
    (status ? ' <span class="badge">' + html(states[status]) + '</span>' : '') + '</div><h1 id="day-title">' + html(day.title) +
    '</h1><p class="hero-detail">' + html(lessonLine(day)) + '</p><div class="hero-bottom">' +
    meta.map(m => '<span class="pill">' + html(m) + '</span>').join('') + '<div class="hero-actions">' + action + '</div></div>');
}

function renderSide(day, status) {
  const complete = state.done.has(day.day);
  const percent = Math.round(100 * state.done.size / state.course.plan.length);
  setMarkup('progress', '<h2>学习进度</h2><div class="progress-num"><strong>' + state.done.size +
    '</strong><span>/ 90 天</span></div><div class="meter" role="progressbar" aria-label="已完成天数" aria-valuemin="0" aria-valuemax="90" aria-valuenow="' +
    state.done.size + '"><span style="width:' + percent + '%"></span></div><button type="button" id="mark" class="mark-button' + (complete ? ' checked' : '') +
    '" aria-pressed="' + complete + '"' + (!isReady(status) ? ' disabled' : '') + '>' + (complete ? '✓ 已完成 · 撤销' : '完成本日学习') +
    '</button><p id="progress-note">打卡记录保存在当前浏览器。</p>');
  const ready = isReady(status);
  const downloads = labels.map(([, name, file]) => {
    const label = file.startsWith('05') ? '昨日与滚动复习' : name;
    const inner = '<span>' + label + '</span><b>' + file.split('.').at(-1).toUpperCase() + '</b>';
    return ready ? '<a class="download-row" href="' + urlFor(day, file) + '" download="' + file + '">' + inner + '</a>' : '<span class="download-row unavailable">' + inner + '</span>';
  }).join('');
  setMarkup('downloads', '<h2>文件下载</h2><div class="download-list">' + downloads +
    '<a class="download-row plan" href="./plan.xlsx" download="N3_90天学习计划总表.xlsx"><span>90 天学习总表</span><b>XLSX</b></a></div>');
  el('mark').onclick = () => {
    if (state.done.has(day.day)) state.done.delete(day.day);
    else state.done.add(day.day);
    const saved = saveDone();
    renderDirectory();
    renderSide(day, status);
    if (!saved) setMarkup('progress-note', '浏览器暂时无法保存，刷新后记录可能丢失。');
    el('mark').focus({ preventScroll: true });
  };
}

const jpLine = value => '<p class="jp jp-line" lang="ja">' + html(value) + '</p>';
const cnLine = value => '<p class="zh-line">' + html(value) + '</p>';
const note = value => '<p class="note">' + html(value) + '</p>';

function lecture(data) {
  const tasks = data.textbook.map(t => {
    const l = state.course.lessons[t.lesson - 1];
    return '<div class="textbook"><h3>' + html(l.volume + '第' + l.number + '课 · ' + l.conversation + '／' + l.reading) +
      '</h3><p>' + html(t.task) + '</p></div>';
  }).join('');
  const explanations = data.grammar.map(g => '<article class="grammar"><h3 lang="ja">' + html(g.name) + '</h3>' +
    cnLine(g.pattern) + cnLine(g.meaning) + jpLine(g.jp) + cnLine(g.cn) + note('易错辨析  ' + g.contrast) + '</article>').join('');
  return '<section class="content-section"><h2>学习情境</h2><p class="intro">' + html(data.background) +
    '</p><div class="step-strip">' + data.steps.map(s => '<span>' + html(s.label) + ' · ' + s.minutes + ' 分</span>').join('') + '</div>' + tasks +
    '</section><section class="content-section"><h2>句型与易错辨析</h2>' + explanations + '</section>' +
    '<section class="content-section"><h2>原创配套短文 · 逐句精读</h2>' + data.translations.map(s =>
      '<article class="read-line">' + jpLine(s.jp) + cnLine(s.cn) + note(s.note) + '</article>').join('') +
    '</section><section class="content-section"><h2>学习自查</h2><ul>' + data.check.map(s => '<li>' + html(s) + '</li>').join('') + '</ul></section>';
}

function question(q, index, source = false) {
  const options = q.options?.length ? '<ol class="options">' + q.options.map((choice, i) =>
    '<li><b>' + String.fromCharCode(65 + i) + '.</b> ' + html(choice) + '</li>').join('') + '</ol>' : '';
  return '<div class="question"><p><span class="q-index">' + String(index + 1).padStart(2, '0') + '</span>' +
    (source && q.sourceDay ? '<span class="q-index">' + idFor(q.sourceDay) + '</span>' : '') + html(q.question) + '</p>' + options + '</div>';
}

function answers(questions, label = '查看答案与解析') {
  return '<details class="answer"><summary>' + label + '</summary><ol>' + questions.map(q => {
    const option = q.options?.length ? String.fromCharCode(65 + q.options.indexOf(q.answer)) + '. ' : '';
    return '<li><strong>' + html(option + q.answer) + '</strong><br>' + html(q.analysis) + '</li>';
  }).join('') + '</ol></details>';
}

function exercises(data) {
  if (!data.questions.length) return '<h2>练习与答案</h2><p class="intro">本场为听解。请在听力材料区作答。</p>' +
    '<button class="small-button" type="button" data-goto-tab="listening">开始听解</button>';
  const passages = (data.passages ?? []).map(p => '<div class="textbook"><h3>' + html(p.title) + '</h3>' +
    p.text.split('\n').map(jpLine).join('') + '</div>').join('');
  let previous = '';
  const questions = data.questions.map((q, i) => {
    const heading = previous !== q.type ? '<h3 class="read-line">' + html(q.type) + '</h3>' : '';
    previous = q.type;
    return heading + question(q, i);
  }).join('');
  return '<h2>练习与答案</h2><p class="intro">先独立作答，再展开末尾的解析。</p>' + passages + questions + answers(data.questions);
}

function vocabulary(data) {
  return '<h2>单词卡</h2><p class="intro">东京式声调 0 为平板，数字 n 表示第 n 拍后降调。下载 Excel 可逐词打卡。</p>' +
    '<div class="table-scroll" tabindex="0" aria-label="词汇表"><table><thead><tr>' +
    ['单词 / 假名', '声调', '词性与释义', '例句'].map(x => '<th scope="col">' + x + '</th>').join('') +
    '</tr></thead><tbody>' + data.vocab.map(v => '<tr><td><span class="jp" lang="ja">' + html(v.word) + '</span><br>' +
    '<span class="note" lang="ja">' + html(v.kana) + '</span><br><span class="note">' + html(v.romaji) + '</span></td><td>' + html(v.accent) +
    '</td><td><span class="note">' + html(v.pos) + '</span><br>' + html(v.cn) + '</td><td>' +
    '<span class="jp" lang="ja">' + html(v.example) + '</span><br><span class="note">' + html(v.exampleCn) + '</span></td></tr>').join('') +
    '</tbody></table></div>';
}

function listening(data) {
  return '<h2>听力材料</h2><p class="intro">先看题再播放。做完题后展开对话原文和解析。</p>' +
    data.listening.map((scene, i) => '<section class="content-section"><h3>' + html(scene.label) + '</h3>' +
      '<div class="player"><button type="button" data-audio="' + i + '" data-rate="1">▶ 正常播放</button><button type="button" data-audio="' + i +
      '" data-rate=".82">慢速播放</button><button type="button" data-stop>停止</button></div><p class="note" id="audio-status-' + i + '" role="status"></p>' +
      scene.questions.map((q, j) => question(q, j)).join('') + answers(scene.questions) +
      '<details class="answer"><summary>展开日文原文与中译</summary>' + scene.dialogue.map(l => jpLine(l.speaker + '  ' + l.text) + cnLine(l.cn)).join('') +
      '</details></section>').join('');
}

function review(data, day) {
  const intervals = day.review.map(r => '<span>' + idFor(r.day) + ' · D+' + r.interval + '</span>').join('');
  return '<h2>五分钟滚动复习</h2>' + (intervals ? '<div class="review-tags">' + intervals + '</div>' : '<p class="intro">课程首日 · 入门热身</p>') +
    data.reviews.map((q, i) => question(q, i, true)).join('') + answers(data.reviews);
}

function stopAudio() {
  speechSession += 1;
  window.speechSynthesis?.cancel();
}
function playAudio(index, rate) {
  const synth = window.speechSynthesis;
  const status = el('audio-status-' + index);
  const scene = state.result?.content?.listening[index];
  if (!scene || !status) return;
  if (!synth || typeof SpeechSynthesisUtterance === 'undefined') {
    status.textContent = '当前设备无法播放日语朗读。可展开原文跟读。';
    return;
  }
  stopAudio();
  const session = speechSession;
  document.querySelectorAll('[id^="audio-status-"]').forEach(p => { p.textContent = ''; });
  const voices = synth.getVoices().filter(voice => /^ja(?:[-_]|$)/i.test(voice.lang));
  if (!voices.length && synth.getVoices().length) {
    status.textContent = '当前设备未提供日语语音。可展开原文跟读。';
    return;
  }
  status.textContent = '正在播放日语朗读…';
  const speakers = [...new Set(scene.dialogue.map(l => l.speaker))];
  const utterances = scene.dialogue.map(line => {
    const utterance = new SpeechSynthesisUtterance(line.text);
    utterance.lang = 'ja-JP';
    utterance.rate = rate;
    if (voices.length) utterance.voice = voices[speakers.indexOf(line.speaker) % voices.length];
    utterance.onerror = () => {
      if (session === speechSession) status.textContent = '朗读未能播放。可展开原文跟读。';
    };
    return utterance;
  });
  utterances.at(-1).onend = () => {
    if (session === speechSession) status.textContent = '播放完毕';
  };
  for (const utterance of utterances) synth.speak(utterance);
}

function renderTabs() {
  setMarkup('tabs', labels.map(([id, label]) => '<button type="button" class="tab" role="tab" id="tab-' + id +
    '" aria-controls="panel" aria-selected="' + (state.tab === id) + '" tabindex="' + (state.tab === id ? '0' : '-1') +
    '" data-tab="' + id + '">' + label + '</button>').join(''));
  el('panel').setAttribute('aria-labelledby', 'tab-' + state.tab);
}

function renderPanel(day, status, data) {
  if (data) {
    const sections = { lecture, exercises, vocab: vocabulary, listening, review };
    setMarkup('panel', sections[state.tab](data, day));
    return;
  }
  const message = status === 'scheduled'
    ? '<p>本日资料于 ' + html(day.date) + ' 08:00（北京时间）开放。</p>'
    : status === 'pending' ? '<p>本日资料正在准备。资料发布后，网页会自动显示内容和文件下载。</p>'
    : status === 'error' ? '<p class="page-error">资料暂时无法读取。<button class="small-button" type="button" id="retry">重试</button></p>'
    : '<p>正在载入资料…</p>';
  setMarkup('panel', '<h2>' + html(states[status] ?? '学习资料') + '</h2><div class="notice">' + message +
    '<p>' + html(day.goal) + '</p>' + (day.focus ? '<p>句型重点 · ' + html(day.focus) + '</p>' : '') + '</div>');
  if (el('retry')) el('retry').onclick = () => selectDay(day.day, { history: false });
}

function changeTab(tab, focus = false) {
  if (!state.active || !labels.some(([id]) => id === tab)) return;
  stopAudio();
  state.tab = tab;
  renderTabs();
  renderPanel(state.active, state.result?.status, state.result?.content);
  if (focus) el('tab-' + tab).focus({ preventScroll: true });
}

function parseDay() {
  return Number(location.hash.match(/^#Day(\d{2})$/)?.[1]) || 0;
}

async function selectDay(number, { history = true, scroll = false } = {}) {
  const day = courseDay(number);
  if (!day) return;
  stopAudio();
  const request = ++state.request;
  state.active = day;
  state.result = null;
  state.tab = labels[0][0];
  if (history && location.hash !== '#' + day.id) window.history.pushState({}, '', '#' + day.id);
  const expected = isReady(state.statuses[day.id]) && !publishedLessons.has(day.id) ? null : state.statuses[day.id];
  renderDirectory(scroll);
  renderHero(day, expected);
  renderSide(day, expected);
  renderTabs();
  renderPanel(day, isReady(expected) ? null : expected);
  try {
    const result = await get('/api/lesson/' + day.id);
    if (request !== state.request) return;
    state.statuses[day.id] = result.status;
    state.result = result;
    renderHero(day, result.status);
    renderSide(day, result.status);
    renderPanel(day, result.status, result.content);
    renderDirectory(scroll);
  } catch {
    if (request !== state.request) return;
    renderPanel(day, 'error');
  }
}

function currentDay(data) {
  const now = Date.parse(data.serverTime);
  return state.course.plan.findLast(day => Date.parse(day.opens) <= now) ?? state.course.plan[0];
}

function drawer(open) {
  el('directory').classList.toggle('open', open);
  el('scrim').classList.toggle('show', open);
  el('open-directory').setAttribute('aria-expanded', String(open));
  if (open) el('search').focus();
}

function wire() {
  el('search').addEventListener('input', () => renderDirectory());
  for (const filter of document.querySelectorAll('[data-filter]')) filter.onclick = () => {
    state.filter = filter.dataset.filter;
    document.querySelectorAll('[data-filter]').forEach(button => button.setAttribute('aria-pressed', String(button === filter)));
    renderDirectory();
  };
  el('day-list').onclick = event => {
    const target = event.target.closest('[data-day]');
    if (!target) return;
    const isMobile = el('directory').classList.contains('open');
    drawer(false);
    selectDay(Number(target.dataset.day));
    if (isMobile) {
      el('main').scrollIntoView({ block: 'start' });
      el('open-directory').focus({ preventScroll: true });
    }
  };
  el('today-link').onclick = () => selectDay(state.today.day, { scroll: true });
  el('open-directory').onclick = () => drawer(true);
  el('close-directory').onclick = el('scrim').onclick = () => {
    drawer(false);
    el('open-directory').focus({ preventScroll: true });
  };
  document.addEventListener('keydown', event => {
    if (event.key === 'Escape' && el('directory').classList.contains('open')) el('close-directory').click();
  });
  el('tabs').onclick = event => {
    const target = event.target.closest('[data-tab]');
    if (target) changeTab(target.dataset.tab);
  };
  el('tabs').onkeydown = event => {
    const index = labels.findIndex(([id]) => id === state.tab);
    const delta = event.key === 'ArrowRight' ? 1 : event.key === 'ArrowLeft' ? -1 : 0;
    if (!delta && !['Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    const next = event.key === 'Home' ? 0 : event.key === 'End' ? labels.length - 1 : (index + delta + labels.length) % labels.length;
    changeTab(labels[next][0], true);
  };
  el('panel').onclick = event => {
    const button = event.target.closest('button');
    if (!button) return;
    if (button.hasAttribute('data-stop')) {
      stopAudio();
      document.querySelectorAll('[id^="audio-status-"]').forEach(p => { p.textContent = ''; });
    }
    if (button.dataset.audio !== undefined) playAudio(Number(button.dataset.audio), Number(button.dataset.rate));
    if (button.dataset.gotoTab) changeTab(button.dataset.gotoTab, true);
  };
  window.addEventListener('popstate', () => selectDay(parseDay() || state.today.day, { history: false }));
}

async function refresh() {
  try {
    const data = await get('/api/status');
    const previous = state.statuses[state.active?.id];
    state.statuses = data.days;
    const oldToday = state.today;
    state.today = currentDay(data);
    if (!parseDay() && oldToday.day !== state.today.day) {
      selectDay(state.today.day, { history: false, scroll: true });
    } else if (state.active && previous !== data.days[state.active.id]) selectDay(state.active.day, { history: false });
    else renderDirectory();
  } catch {
    // Keep the last readable course when a poll fails.
  }
}

async function init() {
  try {
    const [course, data] = await Promise.all([get('/api/course'), get('/api/status')]);
    state.course = course;
    state.statuses = data.days;
    state.today = currentDay(data);
    state.done = safeDone();
    setMarkup('course-range', showDay(course.start) + ' – ' + showDay(course.end));
    wire();
    await selectDay(courseDay(parseDay())?.day ?? state.today.day, { history: false, scroll: true });
    setInterval(refresh, 45000);
    window.addEventListener('focus', refresh);
  } catch {
    setMarkup('panel', '<p class="page-error">暂时无法加载学习室，请刷新页面。</p>');
  }
}
init();
