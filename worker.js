/* All manuscript processing stays in this worker on the user's device. */
let ready;
function progress(text) { postMessage({ progress: text }); }
async function init() {
  progress('처음 사용할 도구를 준비하고 있습니다. 잠시만 기다려 주세요.');
  importScripts('https://cdn.jsdelivr.net/pyodide/v0.27.7/full/pyodide.js');
  const py = await loadPyodide({ indexURL: 'https://cdn.jsdelivr.net/pyodide/v0.27.7/full/' });
  py.FS.mkdirTree('/app/reference/scripts');
  py.FS.mkdirTree('/app/vendor');
  const files = ['engine.py', 'reference/scripts/build.py', 'vendor/olefile.py'];
  await Promise.all(files.map(async path => {
    const response = await fetch(path);
    if (!response.ok) throw new Error('도구 파일을 받지 못했습니다. 인터넷 연결을 확인해 주세요.');
    py.FS.writeFile('/app/' + path, new Uint8Array(await response.arrayBuffer()));
  }));
  py.runPython("import sys; sys.path.insert(0, '/app'); import engine, json");
  return py;
}
// Serialize requests; Python globals and the virtual filesystem belong to one worker.
let queue = Promise.resolve();
onmessage = event => {
  queue = queue.then(() => handle(event.data)).catch(() => {});
};
async function handle(message) {
  const { id, action, data, filename, state, templateData } = message;
  let py;
  try {
    if (!ready) ready = init().catch(error => { ready = null; throw error; });
    py = await ready;
    if (action === 'analyze') {
      progress('원고의 글자와 색을 읽고 있습니다.');
      py.FS.writeFile('/app/input', new Uint8Array(data));
      py.globals.set('input_name', filename);
      const result = py.runPython("json.dumps(engine.analyze(open('/app/input','rb').read(), input_name), ensure_ascii=False)");
      postMessage({ id, result: JSON.parse(result) });
    } else if (action === 'validate') {
      py.FS.writeFile('/app/template.pptx', new Uint8Array(templateData));
      py.runPython("engine.validate_template(open('/app/template.pptx','rb').read())");
      postMessage({ id, result: true });
    } else if (action === 'build') {
      progress('저장된 템플릿에 자막을 넣고 있습니다.');
      py.FS.writeFile('/app/template.pptx', new Uint8Array(templateData));
      py.globals.set('state_json', JSON.stringify(state));
      const count = py.runPython("output_data, output_count = engine.build_ppt(json.loads(state_json), open('/app/template.pptx','rb').read()); open('/app/output.pptx','wb').write(output_data); output_count");
      const output = py.FS.readFile('/app/output.pptx');
      postMessage({ id, result: { data: output.buffer, count } }, [output.buffer]);
    } else throw new Error('지원하지 않는 요청입니다.');
  } catch (error) {
    const detail = String(error.message || error);
    const useful = detail.match(/(?:ValueError|BadZipFile|KeyError|Error): ([^\n]+)/g);
    postMessage({ id, error: useful ? useful.at(-1).replace(/^\w+: /, '') : '처리하지 못했습니다. 인터넷 연결과 파일 형식을 확인한 뒤 다시 시도해 주세요.' });
    console.error(error);
  } finally {
    if (py) {
      for (const path of ['/app/input', '/app/output.pptx', '/app/template.pptx']) {
        try { py.FS.unlink(path); } catch {}
      }
      py.runPython("globals().pop('output_data', None); globals().pop('state_json', None); globals().pop('input_name', None); import gc; gc.collect()");
    }
  }
}
