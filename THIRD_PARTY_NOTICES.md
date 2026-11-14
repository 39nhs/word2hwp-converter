# Third-party notices / 사용한 외부 구성요소

이 저장소에는 외부 라이브러리의 소스 코드가 들어 있지 않습니다.
아래 패키지들은 처음 실행할 때 `requirements.txt` 로 사용자의 PC 에 설치되며, 각자의 라이선스를 따릅니다.

This repository does not contain third-party source code. The packages below are
installed on the user's machine from `requirements.txt` and remain under their own licenses.

| 패키지 (Package) | 버전 (tested) | 라이선스 (License) | 프로젝트 (Project) |
| --- | --- | --- | --- |
| python-hwpx | 6.12.0 | Apache-2.0 | https://github.com/airmang/python-hwpx |
| python-docx | 1.2.0 | MIT | https://github.com/python-openxml/python-docx |
| lxml | 6.1.3 | BSD-3-Clause | https://lxml.de/ |
| tkinterdnd2 | 0.7.0 | MIT | https://github.com/Eliav2/tkinterdnd2 |
| Pillow | 13.0.0 | MIT-CMU (HPND) | https://python-pillow.org/ |
| pywin32 | 312 | PSF | https://github.com/mhammond/pywin32 |

웹 버전 (`index.html`, `worker.js`) 은 페이지를 열 때 아래를 CDN/PyPI 에서 불러옵니다 (저장소에 포함하지 않음):

| 구성요소 | 버전 | 라이선스 | 출처 |
| --- | --- | --- | --- |
| Pyodide (브라우저용 파이썬) | 0.27.7 | MPL-2.0 | https://pyodide.org (cdn.jsdelivr.net) |
| lxml, Pillow (Pyodide 패키지) | Pyodide 0.27.7 포함판 | BSD-3-Clause, MIT-CMU | Pyodide 배포본 |
| python-hwpx | 6.12.0 | Apache-2.0 | PyPI |

선택 사항 (Optional, not bundled):

- **LibreOffice** (MPL-2.0) — 설치되어 있으면 예전 `.doc` 파일을 `.docx` 로 바꾸는 데 사용합니다.
- **한컴오피스 (Hancom Office)** — 설치되어 있으면 한글 프로그램의 자동화 기능으로 변환합니다.
  한컴오피스와 한컴오피스 뷰어는 이 저장소에 포함되어 있지 않으며 재배포하지 않습니다.

## 한글 문서 형식 (HWP / HWPX)

본 제품은 한글과컴퓨터의 한/글 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.

HWPX 형식은 개방형 워드프로세서 마크업 언어(OWPML, KS X 6101) 표준을 따릅니다.

"한글", "한컴", "HWP" 는 (주)한글과컴퓨터의 상표입니다. 이 프로젝트는 (주)한글과컴퓨터와 관련이 없는 개인 프로젝트입니다.
"Hangul", "Hancom" and "HWP" are trademarks of Hancom Inc. This is an independent project, not affiliated with or endorsed by Hancom Inc.
"Microsoft Word" is a trademark of Microsoft Corporation.
