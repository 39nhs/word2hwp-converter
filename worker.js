// 브라우저 안에서 파이썬(Pyodide)으로 word2hwp_core.py 를 실행하는 작업자.
// 파일은 서버로 보내지 않고 이 브라우저 안에서만 변환한다.
// 필요한 엔진 파일은 모두 이 사이트의 vendor/ 폴더에서 받는다 (외부 서버에 의존하지 않음).
const PYODIDE = new URL("vendor/pyodide/", self.location.href).href;
const HWPX_WHEEL = new URL("vendor/python_hwpx-6.12.0-py3-none-any.whl", self.location.href).href;
importScripts(PYODIDE + "pyodide.js");

let py = null;

function post(type, data) { self.postMessage(Object.assign({ type }, data)); }

async function init(coreUrl) {
  post("status", { text: "변환 엔진 내려받는 중… (처음 한 번만, 10~30초)", pct: 5 });
  py = await loadPyodide({ indexURL: PYODIDE });
  post("status", { text: "문서 처리 도구 준비 중…", pct: 35 });
  await py.loadPackage(["lxml", "Pillow"]);
  post("status", { text: "한글 파일 라이브러리 준비 중…", pct: 65 });
  await py.loadPackage(HWPX_WHEEL);
  post("status", { text: "변환기 불러오는 중…", pct: 90 });
  const code = await (await fetch(coreUrl, { cache: "no-cache" })).text();
  py.FS.mkdirTree("/app"); py.FS.mkdirTree("/work");
  py.FS.writeFile("/app/word2hwp_core.py", code);
  py.runPython("import sys; sys.path.insert(0, '/app'); import word2hwp_core");
  post("ready", {});
}

async function convert(id, name, bytes, fmt) {
  const dir = "/work/" + id;
  py.FS.mkdirTree(dir);
  const ext = (name.match(/\.[^.]+$/) || [".docx"])[0].toLowerCase();
  const src = dir + "/input" + ext;
  py.FS.writeFile(src, new Uint8Array(bytes));
  py.globals.set("_src", src); py.globals.set("_dir", dir); py.globals.set("_fmt", fmt);
  try {
    const res = py.runPython(`
import json, word2hwp_core as c
try:
    _out, _w = c.convert_file(_src, out_dir=_dir, fmt=_fmt, use_hancom=False, overwrite=True)
    _r = {"ok": True, "out": _out, "warnings": _w}
except ValueError as e:
    _r = {"ok": False, "error": str(e)}
json.dumps(_r, ensure_ascii=False)
`);
    const r = JSON.parse(res);
    if (r.ok) {
      const data = py.FS.readFile(r.out);
      post("done", { id, data, warnings: r.warnings }, );
    } else {
      post("failed", { id, error: r.error });
    }
  } catch (e) {
    const msg = String(e && e.message || e).trim().split("\n").filter(Boolean).pop();
    post("failed", { id, error: "예상 못한 오류: " + msg });
  } finally {
    try { py.runPython(`import shutil; shutil.rmtree(${JSON.stringify(dir)}, ignore_errors=True)`); } catch (_) {}
  }
}

self.onmessage = async (ev) => {
  const m = ev.data;
  try {
    if (m.type === "init") await init(m.coreUrl);
    else if (m.type === "convert") await convert(m.id, m.name, m.bytes, m.fmt);
  } catch (e) {
    post("fatal", { error: String(e && e.message || e) });
  }
};
