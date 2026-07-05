# 실행 지시서: 화면별 스크린샷 + 기획 설명 페이지 제작

> 새 세션용 자기완결 지시서. 이 파일을 읽었다면 아래 순서대로 실행하라.
> 작성: 2026-07-05 (이전 세션 인수인계)

## 1. 프로젝트 컨텍스트 (30초 요약)

- 뤼튼테크놀로지스 AX PM(Ontology) 채용 과제 제출물. 마감 임박.
- `site/index.html` = 제출용 문서 사이트(SPA, 로그인 포털 + PRD/로드맵/가격/화면설계 4개 북).
- `site/app.html` = 화면 설계(과제 4-2) 산출물인 **통합 인터랙티브 데모** (단일 자립 HTML).
  - URL 딥링크: `?screen=admin|reviewer|onboarding`, `?tab=dashboard|ontology|workflow|roles|integration|policy|audit|metering`, `?wf=card|invoice`, `?step=1~7`
- 관련 문서: `docs/prd.md`(v0.8), `docs/roadmap-and-pricing.md`, `docs/screen-design-spec.md`(v1, 대체됨)
- 주의: 사용자가 파일을 IDE로 열어두는 경우가 많아 Edit 도구가 "modified since read" 충돌을 낼 수 있음 → 재시도하거나 perl 인플레이스 편집으로 대체.

## 2. 과업

각 메뉴/화면을 스크린샷으로 캡처해 `site/screenshots/`에 저장하고, 문서 사이트의 **화면 설계 북**에 화면별 설명 페이지를 추가한다.

## 3. 스크린샷 캡처 (1단계)

로컬 서버 + 헤드리스 Chrome 사용 (file:// 는 쿼리스트링이 안 먹으므로 반드시 서버 경유):

```bash
cd "c:/workspace/디렉팅/뤼튼/site" && mkdir -p screenshots
(python -m http.server 8380 >/dev/null 2>&1 &) && sleep 1
CHROME="/c/Program Files/Google/Chrome/Application/chrome.exe"
# 캡처 패턴 (창 1600x1000, 가상시간 4초):
"$CHROME" --headless --disable-gpu --window-size=1600,1000 --virtual-time-budget=4000 \
  --screenshot="$(cygpath -w "$(pwd)")\\screenshots\\<파일명>.png" "http://localhost:8380/app.html?<파라미터>"
# 전부 끝나면: taskkill //F //IM python.exe
```

캡처 목록 (파일명 → URL 파라미터):

| 파일명 | 파라미터 | 화면 |
| --- | --- | --- |
| admin-01-dashboard.png | ?tab=dashboard | 대시보드 |
| admin-02-ontology.png | ?tab=ontology | 온톨로지 스튜디오 |
| admin-03-workflow.png | ?tab=workflow | 워크플로우 빌더 |
| admin-04-roles.png | ?tab=roles | 권한 관리 |
| admin-05-integration.png | ?tab=integration | 연동 관리 |
| admin-06-policy.png | ?tab=policy | 규정 센터 |
| admin-07-audit.png | ?tab=audit | 감사 로그 |
| admin-08-metering.png | ?tab=metering | 미터링 |
| reviewer-01-card.png | ?screen=reviewer&wf=card | 검토자 · 법인카드 |
| reviewer-02-invoice.png | ?screen=reviewer&wf=invoice | 검토자 · 세금계산서 3-Way |
| onboarding-step1.png ~ step7.png | ?screen=onboarding&step=N (N=1~7) | 온보딩 7단계 |

총 17장. 각 캡처 후 Read 도구로 이미지를 열어 깨진 화면이 없는지 확인할 것.

## 4. 문서 사이트에 설명 페이지 추가 (2단계)

`site/index.html`의 구조:
- 내비게이션: `<script type="application/json" id="nav-data">` 한 줄 JSON. `books` 배열의 `id:"ui"` 북("화면 설계")에 페이지 추가.
- 본문: `<script type="text/html" data-page="<id>">` 템플릿 블록을 추가 (기존 `data-page="ui-intro"` 블록 뒤에 나란히).

추가할 페이지 (nav-data의 ui 북 groups에 새 그룹으로):

```json
{"label":"화면별 상세","pages":[
 {"id":"ui-admin","title":"관리자 콘솔 (8탭)"},
 {"id":"ui-reviewer","title":"검토자 화면"},
 {"id":"ui-onboarding","title":"고객사 온보딩 (7단계)"}]}
```

17장을 페이지 17개로 쪼개지 말고 **3개 페이지로 묶는다** (관리자 8탭 = 1페이지에 8섹션).

각 섹션 구성(순서 고정):
1. `<h2>` 화면명
2. 스크린샷: `<img src="screenshots/파일명.png" alt="..." style="max-width:100%;border:1px solid var(--line);border-radius:10px;margin:6px 0 14px">`
3. **화면 목적** 1문장
4. **핵심 구성** 3~5개 불릿 (화면에 실제 보이는 요소 기준 — 스크린샷을 Read로 보고 쓸 것)
5. **PRD 근거** 1줄: 아래 매핑 사용

| 화면 | PRD 근거 |
| --- | --- |
| 대시보드 | §13 성공 지표(North Star=자동 처리 건수), §6.1 Audit & Metering |
| 온톨로지 스튜디오 | §3.4 강점 1·2(설정형 확장·차이 흡수), §6.2 |
| 워크플로우 빌더 | §8.7/§9.7 신뢰도 임계값 90/95, §3.3 설정형 워크플로우 |
| 권한 관리 | §3.4 강점 4(SoD·내부통제), §5.5 감사인 페르소나 |
| 연동 관리 | §16 SAP 리스크 대응, NFR-004 재처리 큐 |
| 규정 센터 | §6.1 Knowledge Layer, §16 규정 문서 품질 리스크 |
| 감사 로그 | NFR-001, §5.5 컴플라이언스 |
| 미터링 | §15.3 미터링 이벤트, 로드맵 §10 |
| 검토자 | §8.6/§9.6 검토자 UX, 일괄 처리 규칙 |
| 온보딩 | §3.4 온보딩 제품화 = 마진, §11.2 신규 개발 범위 |

문체: 기존 문서와 동일한 존댓말 없는 개조식/평서문. 과장 금지. 모든 수치는 "가상 데이터" 전제.

## 5. 검증 (3단계)

1. `node`로 nav-data JSON 파싱 확인 (깨지면 사이트 전체가 죽는다):
   `node -e "JSON.parse(require('fs').readFileSync('index.html','utf8').match(/id=\"nav-data\">(.*?)<\/script>/s)[1]);console.log('OK')"`
2. 헤드리스 Chrome으로 `index.html#ui-admin` 스크린샷 → 이미지가 표시되는지 확인 (로그인 포털이 덮고 있으면 sessionStorage 때문이니 `#ui-admin` 해시로도 포털이 뜨는 게 정상 — 포털의 4-2 링크 클릭 흐름은 수동 확인 대상으로 사용자에게 안내).
3. 완료 보고에 페이지 3개·스크린샷 17장 체크리스트 포함.

## 5.5 추가 과업: 로드맵 파트 사이트 동기화

`docs/roadmap-and-pricing.md`는 v0.4까지 개정됐지만(4-3 §3 "4인 스쿼드" 재작성 포함), `site/index.html`의 로드맵 페이지(`data-page="rm-1"`~`"rm-6"`)는 구버전이다. **가격 페이지(pr-7~pr-12)는 이미 v0.4로 동기화 완료** — 같은 방식으로 rm-1~rm-6 템플릿 블록을 md 내용 기준으로 교체하라. 교체 시 node 스크립트로 `<script type="text/html" data-page="rm-N">...</script>` 블록을 정규식 치환하고, 완료 후 nav-data JSON 파싱을 반드시 검증할 것. rm-3 제목도 md와 맞춰 갱신("3. 타임라인 & 개발 리소스 (4인 스쿼드, 동시진행)").

## 6. 하지 말 것

- app.html 수정 금지 (스크린샷 대상일 뿐).
- nav-data JSON을 여러 줄로 포매팅하지 말 것 (한 줄 유지).
- v1 목업 파일 복원 금지 (의도적으로 제거됨).
- 스크린샷 파일명·경로 임의 변경 금지 (위 표 그대로).
