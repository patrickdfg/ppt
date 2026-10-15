'use strict';
const $ = id => document.getElementById(id);
const types = {skip:'제외',caption:'아래 자막',word:'말씀 전체 화면',verse:'인용 성구',passage:'서두 본문 성구',emphasis:'강조',emphasis2:'강조 · 다른 배경',notice:'광고·공지'};
const colorNames = {BLUE:'파란 글씨',GREEN:'녹색 글씨',RED:'빨간 글씨',BLACK:'검정 글씨',MANUAL:'직접 추가'};
let state = null, templateData = null, templateName = '기본 설교 자막 템플릿';
let defaultData = null, outputURL = null, busy = false, worker = null, serial = 0;
const pending = new Map();
function status(text, kind = '') {
  $('status').hidden = !text; $('status').textContent = text; $('status').className = 'status ' + kind;
}
function invalidate() {
  $('result').hidden = true;
  if (outputURL) { URL.revokeObjectURL(outputURL); outputURL = null; }
}
function setBusy(value) {
  busy = value;
  for (const control of document.querySelectorAll('input,textarea,select,button')) control.disabled = value;
  $('dropZone').setAttribute('aria-busy', String(value));
}
function getWorker() {
  if (worker) return worker;
  worker = new Worker('worker.js');
  worker.onmessage = ({data}) => {
    if (data.progress) { status(data.progress, 'busy'); return; }
    const task = pending.get(data.id);
    if (!task) return;
    pending.delete(data.id); clearTimeout(task.timer);
    data.error ? task.reject(new Error(data.error)) : task.resolve(data.result);
  };
  worker.onerror = () => {
    worker?.terminate(); worker = null;
    for (const task of pending.values()) { clearTimeout(task.timer); task.reject(new Error('도구를 불러오지 못했습니다. 인터넷 연결을 확인하고 다시 시도해 주세요.')); }
    pending.clear();
  };
  return worker;
}
function request(action, values = {}) {
  return new Promise((resolve, reject) => {
    const id = ++serial, w = getWorker();
    const timer = setTimeout(() => {
      w.terminate(); worker = null;
      for (const task of pending.values()) { clearTimeout(task.timer); task.reject(new Error('처리 시간이 길어 중단했습니다. 인터넷 연결을 확인하거나 원고를 나눠 다시 시도해 주세요.')); }
      pending.clear();
    }, 300000);
    pending.set(id, {resolve, reject, timer}); w.postMessage({id, action, ...values});
  });
}
function db() {
  return new Promise((resolve, reject) => {
    if (!window.indexedDB) return reject(new Error('이 브라우저는 템플릿 저장을 지원하지 않습니다.'));
    const open = indexedDB.open('seolgyo-jamak', 1);
    open.onupgradeneeded = () => open.result.createObjectStore('templates');
    open.onsuccess = () => resolve(open.result);
    open.onerror = () => reject(new Error('템플릿 저장소를 열 수 없습니다. 브라우저의 저장 공간 설정을 확인해 주세요.'));
  });
}
async function storedTemplate(operation, value) {
  const database = await db();
  return new Promise((resolve, reject) => {
    const transaction = database.transaction('templates', operation === 'get' ? 'readonly' : 'readwrite');
    const store = transaction.objectStore('templates');
    const query = operation === 'get' ? store.get('personal') : operation === 'delete' ? store.delete('personal') : store.put(value, 'personal');
    transaction.oncomplete = () => { database.close(); resolve(query.result); };
    transaction.onerror = transaction.onabort = () => { database.close(); reject(new Error('템플릿을 저장하지 못했습니다. 기기의 저장 공간을 확인해 주세요.')); };
  });
}
const templateLoaded = (async () => {
  try {
    const saved = await storedTemplate('get');
    if (saved) { templateData = saved.data; templateName = saved.name; $('templateName').textContent = templateName + ' · 이 기기에 저장됨'; }
  } catch (error) { status(error.message, 'error'); }
})();
async function getTemplate() {
  await templateLoaded;
  if (templateData) return templateData;
  if (!defaultData) {
    const response = await fetch('assets/template.pptx');
    if (!response.ok) throw new Error('기본 템플릿을 받지 못했습니다. 다시 시도해 주세요.');
    defaultData = await response.arrayBuffer();
  }
  return defaultData;
}
function node(tag, text, className) {
  const el = document.createElement(tag);
  if (text !== undefined) el.textContent = text;
  if (className) el.className = className;
  return el;
}
function updateCount() {
  const count = state.entries.filter(e => e.type !== 'skip' && e.text.trim()).length;
  $('counts').textContent = `자막 ${count}개 · 전체 문단 ${state.entries.length}개`;
}
function editEntry(article, entry, index) {
  const old = article.querySelector('.entry-controls');
  if (old) { old.remove(); return; }
  const controls = node('div', undefined, 'entry-controls'), text = node('textarea');
  text.value = entry.text; text.setAttribute('aria-label', `${index + 1}번 자막 내용`);
  text.addEventListener('input', () => {
    entry.text = text.value; article.querySelector('.entry-preview').textContent = entry.text.replace(/<\/?[rb]>/g, ''); updateCount(); invalidate();
  });
  controls.append(text, node('p', '줄바꿈을 직접 고칠 수 있습니다. <r>빨간 강조</r>는 PPT에 빨간색으로 들어갑니다.', 'small'));
  const options = node('div', undefined, 'entry-options');
  const refLabel = node('label', '성경 장절', 'ref-field'), ref = node('input');
  ref.value = entry.ref || ''; ref.placeholder = '예: 요 16:13';
  ref.oninput = () => { entry.ref = ref.value; invalidate(); }; refLabel.append(ref);
  const speakerLabel = node('label', '말씀 배경', 'speaker-field'), speaker = node('select');
  for (const [value, label] of [['','기본 배경'],['성령','성령님 배경']]) { const option = node('option', label); option.value = value; speaker.append(option); }
  speaker.value = entry.speaker || ''; speaker.onchange = () => { entry.speaker = speaker.value; invalidate(); }; speakerLabel.append(speaker);
  const revealLabel = node('label', undefined, 'check'), reveal = node('input');
  reveal.type = 'checkbox'; reveal.checked = !!entry.reveal; reveal.onchange = () => { entry.reveal = reveal.checked; invalidate(); };
  revealLabel.append(reveal, node('span', '한 줄씩 누적 공개')); options.append(refLabel, speakerLabel, revealLabel);
  const actions = node('div', undefined, 'entry-actions');
  for (const [label, delta] of [['↑ 위로',-1],['↓ 아래로',1]]) {
    const button = node('button', label); button.type = 'button'; button.disabled = index + delta < 0 || index + delta >= state.entries.length;
    button.onclick = () => {
      const destination = index + delta;
      if (destination < 0 || destination >= state.entries.length) return;
      [state.entries[index], state.entries[destination]] = [state.entries[destination], state.entries[index]]; invalidate(); render();
    }; actions.append(button);
  }
  const exclude = node('button', entry.type === 'skip' ? '아래 자막으로 넣기' : '이 문단 제외'); exclude.type = 'button';
  exclude.onclick = () => { entry.type = entry.type === 'skip' ? 'caption' : 'skip'; invalidate(); render(); }; actions.append(exclude);
  controls.append(options, actions); article.append(controls); text.focus();
}
function render() {
  $('entries').replaceChildren(); const all = $('showSkipped').checked, fragment = document.createDocumentFragment(); let visible = 0;
  state.entries.forEach((entry, index) => {
    if (entry.type === 'skip' && !all) return;
    visible++;
    const article = node('article', undefined, 'entry' + (entry.type === 'skip' ? ' skipped' : '')), head = node('div', undefined, 'entry-head'), select = node('select');
    select.setAttribute('aria-label', `${index + 1}번 문단 종류`);
    for (const [value, label] of Object.entries(types)) { const option = node('option', label); option.value = value; select.append(option); }
    select.value = entry.type; select.onchange = () => { entry.type = select.value; invalidate(); render(); };
    head.append(node('span', String(index + 1).padStart(2,'0'), 'entry-number'), select, node('span', colorNames[entry.color] || '색 확인 필요', 'entry-color'));
    const preview = node('button', entry.text.replace(/<\/?[rb]>/g, ''), 'entry-preview'); preview.type = 'button'; preview.title = '눌러서 자막 편집'; preview.onclick = () => editEntry(article, entry, index);
    article.append(head, preview); fragment.append(article);
  });
  $('entries').append(fragment); $('empty').hidden = visible !== 0; updateCount();
}
async function openManuscript(file) {
  if (!file || busy) return;
  if (!/\.(hwp|hwpx|txt)$/i.test(file.name)) return status('HWP, HWPX, TXT 파일을 선택해 주세요.', 'error');
  if (file.size > 30 * 1024 * 1024) return status('원고는 최대 30MB까지 가능합니다. 사진을 줄이거나 원고를 나눠 주세요.', 'error');
  setBusy(true); invalidate();
  try {
    const result = await request('analyze', {filename: file.name, data: await file.arrayBuffer()}); state = result;
    $('fileName').textContent = file.name; $('title').value = result.title; $('date').value = ''; $('editor').hidden = false; $('showSkipped').checked = false;
    $('warnings').replaceChildren(...result.warnings.map(w => node('p', w))); $('warnings').hidden = !result.warnings.length;
    render(); status('원고를 읽었습니다. 자막 종류와 줄바꿈을 확인한 뒤 PPT를 만들어 주세요.'); $('editor').scrollIntoView({behavior:'smooth', block:'start'});
  } catch (error) { status(error.message, 'error'); }
  finally { setBusy(false); $('manuscript').value = ''; }
}
$('manuscript').onchange = event => openManuscript(event.target.files[0]);
for (const name of ['dragenter', 'dragover']) $('dropZone').addEventListener(name, event => { event.preventDefault(); if (!busy) $('dropZone').classList.add('drag'); });
$('dropZone').addEventListener('dragleave', () => $('dropZone').classList.remove('drag'));
$('dropZone').addEventListener('drop', event => { event.preventDefault(); $('dropZone').classList.remove('drag'); openManuscript(event.dataTransfer.files[0]); });
$('showSkipped').onchange = () => { if (state) render(); };
$('title').oninput = () => { if (state) state.title = $('title').value; invalidate(); };
$('date').oninput = () => { if (state) state.date = $('date').value; invalidate(); };
$('addEntry').onclick = () => { state.entries.push({type:'caption',text:'새 자막을 입력해 주세요.',ref:'',speaker:'',color:'MANUAL',reveal:false}); invalidate(); render(); const last = $('entries').lastElementChild; last.querySelector('.entry-preview').click(); last.scrollIntoView({behavior:'smooth',block:'center'}); };
$('templateInput').onchange = async event => {
  const file = event.target.files[0]; if (!file || busy) return;
  if (!/\.pptx$/i.test(file.name) || file.size > 30 * 1024 * 1024) return status('30MB 이하의 PPTX 템플릿을 선택해 주세요.', 'error');
  setBusy(true);
  try {
    await templateLoaded; const data = await file.arrayBuffer(); await request('validate', {templateData:data}); await storedTemplate('put', {name:file.name, data});
    templateData = data; templateName = file.name; $('templateName').textContent = templateName + ' · 이 기기에 저장됨'; invalidate(); status('템플릿을 저장했습니다. 다음에 열어도 이 기기에서 계속 사용합니다.');
  } catch (error) { status(error.message, 'error'); }
  finally { setBusy(false); event.target.value = ''; }
};
$('resetTemplate').onclick = async () => {
  if (busy) return; setBusy(true);
  try { await templateLoaded; await storedTemplate('delete'); templateData = null; templateName = '기본 설교 자막 템플릿'; $('templateName').textContent = templateName; invalidate(); status('저장소의 기본 템플릿으로 되돌렸습니다.'); }
  catch (error) { status(error.message, 'error'); }
  finally { setBusy(false); }
};
$('build').onclick = async () => {
  if (!state || busy) return;
  setBusy(true); invalidate(); status('템플릿을 준비하고 있습니다.', 'busy');
  try {
    const result = await request('build', {state, templateData:await getTemplate()});
    outputURL = URL.createObjectURL(new Blob([result.data], {type:'application/vnd.openxmlformats-officedocument.presentationml.presentation'})); $('download').href = outputURL;
    $('download').download = (state.date || state.title || '설교').replace(/[<>:"/\\|?*\x00-\x1f]/g,'_').slice(0,100) + '_자막.pptx';
    $('resultDetail').textContent = `${result.count}장 · ${templateName}`; $('result').hidden = false; status('완료했습니다. 아래 버튼으로 PPT를 받아 주세요.'); $('result').scrollIntoView({behavior:'smooth',block:'center'});
  } catch (error) { status(error.message, 'error'); }
  finally { setBusy(false); }
};
window.addEventListener('beforeunload', event => { if (busy) { event.preventDefault(); event.returnValue = ''; } });
