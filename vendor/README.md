# vendor/ — 웹 버전이 쓰는 외부 구성요소 (수정하지 않은 원본)

웹 버전(`index.html`, `worker.js`)이 외부 서버에 의존하지 않도록, 필요한 파일을 이 폴더에 함께 넣었습니다.
모든 파일은 아래 공식 배포처에서 받은 **수정하지 않은 원본**이며, 받은 뒤 공식 체크섬(sha256)으로 확인했습니다.

| 파일 | 구성요소 | 버전 | 라이선스 | 받은 곳 | 소스 코드 |
| --- | --- | --- | --- | --- | --- |
| `pyodide/pyodide.js`, `pyodide.asm.js`, `pyodide.asm.wasm`, `pyodide-lock.json` | Pyodide (브라우저용 CPython, Emscripten 포함) | 0.27.7 | MPL-2.0 | https://cdn.jsdelivr.net/pyodide/v0.27.7/full/ | https://github.com/pyodide/pyodide/tree/0.27.7 |
| `pyodide/python_stdlib.zip` | CPython 3.12 표준 라이브러리 | 3.12.7 | PSF | 위와 같음 | https://github.com/python/cpython/tree/v3.12.7 |
| `pyodide/lxml-5.2.1-…wasm32.whl` | lxml | 5.2.1 | BSD-3-Clause | 위와 같음 | https://github.com/lxml/lxml |
| `pyodide/pillow-10.2.0-…wasm32.whl` | Pillow | 10.2.0 | MIT-CMU (HPND) | 위와 같음 | https://github.com/python-pillow/Pillow |
| `python_hwpx-6.12.0-py3-none-any.whl` | python-hwpx | 6.12.0 | Apache-2.0 | https://pypi.org/project/python-hwpx/6.12.0/ | https://github.com/airmang/python-hwpx |

각 구성요소의 라이선스 원문은 [`licenses/`](licenses/) 에 있습니다
(`python-hwpx-NOTICE` 는 Apache-2.0 에 따라 함께 배포하는 고지 파일입니다).

이 폴더의 파일들은 이 프로젝트의 MIT 라이선스가 아니라 각자의 라이선스를 따릅니다.
