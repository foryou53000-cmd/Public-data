const fs = require('fs');
const path = require('path');
const vm = require('vm');

const HTML_FILE = '경북체험학습추천 (2).html';
const OUT_FILE = '경북_교육과정_관광자원_전처리_근거자료.xlsx';

function readApp() {
  const html = fs.readFileSync(HTML_FILE, 'utf8');
  const match = html.match(/<script>([\s\S]*)<\/script>/);
  if (!match) throw new Error('script block not found');

  const context = {
    console,
    localStorage: { getItem() { return '{}'; }, setItem() {} },
    document: { getElementById() { return null; }, addEventListener() {} },
    window: {},
    setTimeout,
    fetch: async () => ({ ok: true, json: async () => ({}) }),
  };
  vm.createContext(context);
  vm.runInContext(match[1], context);
  vm.runInContext(`
    globalThis.__APP_EXPORT__ = {
      CUR, SOCIAL_CURRICULUM, CATS, SUB_MAP, TEXTBOOK_SOURCE, CURRICULUM_SOURCE, NIKH_SOURCE, HISTORY_PUBLIC_SOURCE, TOUR_CODE_GROUPS,
      GBK, SIGUNGU, TK, TOUR, APP,
      socialUnitProfile, curriculumRelevance, scorePlace, curriculumData, buildPlace,
      setState(g, s) { A.gr = g; A.sj = s; }
    };
  `, context);
  return context.__APP_EXPORT__;
}

function uniq(arr) {
  return Array.from(new Set((arr || []).filter(Boolean)));
}

function join(arr) {
  return uniq(arr).join(', ');
}

function regexText(v) {
  return v ? String(v) : '';
}

function profileRows(app) {
  const rows = [['학년', '교과', '선택 단원', '판정 유형', '최소 통과 점수', '강한 근거 키워드', '보조 키워드', '제외 패턴', '관광 API 코드 그룹', '코드 목록', '세부 주제 분기 규칙', '판정 방식']];
  Object.entries(app.CUR).forEach(([grade, subjects]) => {
    (subjects.사회 || []).forEach((unit) => {
      const p = app.socialUnitProfile(unit);
      const codes = p?.codeGroup ? (app.TOUR_CODE_GROUPS[p.codeGroup] || []) : [];
      const rules = (p?.topicByText || []).map((r) => `${r.re} -> ${r.topic}`).join(' / ');
      rows.push([
        `${grade}학년`, '사회', unit.unit || unit.name, p?.label || '핵심어 기반',
        p?.min || '', join(p?.strong), join(p?.keywords), regexText(p?.deny),
        p?.codeGroup || '', join(codes), rules,
        '강한 근거(+32), 보조 근거(+12), 관광분류 코드(+34), 세부주제 점수를 합산하고 최소 점수 미만은 제외',
      ]);
    });
  });
  return rows;
}

function unitRows(app) {
  const rows = [['학년', '교과', '선택 UI 단원명', '내부 세부주제 수', '내부 세부주제 목록', '통합 검색 키워드', '통합 장소유형 키워드', '통합 핵심 키워드', '선택 화면 정책', '자료 출처']];
  Object.entries(app.CUR).forEach(([grade, subjects]) => {
    (subjects.사회 || []).forEach((unit) => {
      rows.push([
        `${grade}학년`, '사회', unit.unit || unit.name, unit.subtopics?.length || 0,
        (unit.subtopics || []).map((t) => `${t.name}(${t.lesson})`).join(' / '),
        join(unit.searchTerms), join(unit.placeTypes), join(unit.terms),
        '선택 화면에서는 큰 단원 1개로 표시하고, 장소 추천 결과에서만 관련 세부 주제와 차시를 연결',
        unit.source || app.TEXTBOOK_SOURCE,
      ]);
    });
  });
  return rows;
}

function topicRows(app) {
  const rows = [['학년', '교과', '큰 단원', '내부 주제', '차시', '성취기준/학습 연결 진술', '수업 활동', '핵심 초점', '핵심 키워드', '검색 키워드', '장소유형 키워드', '자료 출처']];
  Object.entries(app.SOCIAL_CURRICULUM).forEach(([grade, units]) => {
    units.forEach((u) => {
      rows.push([
        `${grade}학년`, '사회', u.unit, u.name, u.lesson, u.standard, u.activity, u.focus,
        join(u.terms), join(u.searchTerms), join(u.placeTypes), u.source || app.TEXTBOOK_SOURCE,
      ]);
    });
  });
  return rows;
}

function categoryRows(app) {
  const rows = [['구분', '대분류/소분류', '아이콘', '하위 분류', '장소명/주소 매칭 키워드', '전처리 활용']];
  Object.entries(app.CATS).forEach(([cat, d]) => {
    rows.push(['대분류', cat, d.ic, join(d.subs), join(d.kw), '장소명과 주소를 기반으로 앱의 관광자원 대분류를 부여']);
  });
  Object.entries(app.SUB_MAP).forEach(([sub, kws]) => {
    rows.push(['소분류', sub, '', '', join(kws), '장소명과 주소를 기반으로 세부 장소유형을 부여']);
  });
  return rows;
}

function sourceRows(app) {
  return [
    ['데이터명', '제공/출처', '수집 방식', '주요 필드', '앱에서의 활용', '비고'],
    ['학교 기본정보', 'NEIS 학교기본정보 Open API', '사용자가 학교명을 검색할 때 실시간 조회', 'SCHUL_NM, ORG_RDNMA, SCHUL_KND_SC_NM', '학교 위치 기준 설정, 거리 계산 기준점 설정', '교육 공공데이터 활용 근거'],
    ['2022 개정 사회과 교육과정', '교육부 고시 제2022-33호 [별책 7] 사회과 교육과정', '공식 성취기준을 학년군·영역·성취기준 코드 단위로 구조화', '성취기준 코드, 성취기준 문장, 영역, 학습 내용', '5학년 2학기 역사 단원의 최상위 교육과정 기준으로 사용', '5, 6학년 2학기 교과서 공백을 공식 성취기준으로 보완'],
    ['국사편찬위원회 역사 원자료', '교육부 국사편찬위원회 한국사데이터베이스 공공데이터', '성취기준별로 연결 가능한 사료 DB를 시대·주제 태그로 매핑', '삼국유사 원문, 한국고대금석문, 고려사 원문, 한국사료총서, 일제감시대상인물카드, 역사지리정보', '역사 장소 추천의 공공데이터 검증 근거, 팝업 추천 이유와 직접 연계 근거 생성', '교육부 소속기관 데이터로 대회 교육공공데이터 활용 근거 강화'],
    ['초등 사회 교과서 PDF', '천재교과서 김정인 저 사회 교과서 PDF', '사용자 제공 PDF를 분석해 단원명, 주제, 차시, 핵심 개념 정리', '학년, 단원, 세부 주제, 차시, 학습활동, 핵심어', '교육과정 태그와 성취기준 연결 문장 생성', '3~4학년, 5-1, 6-1은 교과서 단원명 보조 기준으로 활용'],
    ['관광자원 정보', '한국관광공사 Tour API / 공공데이터포털', '경북 areaCode=35 기준 키워드 검색 및 위치 기반 검색', 'contentid, title, addr1, mapx, mapy, cat1, cat2, cat3, firstimage, sigungucode, overview', '관광자원 후보 생성, 이미지, 주소, 관광분류, 상세 설명 제공', '교육 데이터와 결합되는 이종 공공데이터'],
    ['지도/좌표 보정', 'OpenStreetMap Nominatim, Leaflet', '학교 주소를 좌표화하고 지도에 시각화', 'lat, lon, 주소', '학교-관광자원 거리 계산 및 지도 표시', '서비스 구현 보조 데이터'],
    ['앱 내부 선별 기준', HTML_FILE, '자바스크립트 규칙과 키워드 사전', '강한 근거, 보조 근거, 제외 패턴, 통과 기준 점수', '설명 가능한 추천 결과 산출', `생성일: ${new Date().toLocaleString('ko-KR')}`],
  ];
}

function processRows() {
  return [
    ['단계', '입력 데이터', '전처리 작업', '산출 데이터', '심사 어필 포인트'],
    ['1. 교육과정 원자료 수집', '2022 개정 사회과 교육과정, 천재교과서 김정인 저 초등 사회 교과서 PDF', '공식 성취기준은 코드·문장 단위로 구조화하고, 교과서는 큰 단원·세부 주제·차시 보조자료로 추출', '학년별 사회 단원 원자료 표', '추천의 출발점이 관광지가 아니라 공식 교육과정임을 보여줌'],
    ['2. 역사 원자료 수집', '국사편찬위원회 한국사데이터베이스 공공데이터', '삼국유사, 한국고대금석문, 고려사, 한국사료총서, 일제감시대상인물카드 등을 시대·사건·장소 키워드로 태깅', '역사 단원별 공공 사료 매핑표', '역사 장소 추천이 단순 관광 정보가 아니라 교육공공데이터 원자료와 연결됨'],
    ['3. 단원 구조화', '성취기준, 세부 주제/차시 단위 자료', '선택 화면은 큰 단원으로 병합하고, 세부 주제와 차시는 내부 분석용으로 보존', '선택 UI 단원표, 세부 주제표', '교사는 단순하게 단원을 고르고, AI는 내부에서 차시를 정밀 연결'],
    ['4. 교육과정 키워드 태깅', '단원명, 성취기준, 역사 원자료 태그, 학습활동', '핵심 키워드, 검색 키워드, 장소유형 키워드, 제외 키워드를 부여', '단원별 키워드/선별 기준표', '왜 어떤 장소가 추천/제외되는지 설명 가능'],
    ['5. 관광자원 원자료 수집', 'Tour API, 학교 위치, 경북 지역 코드', '경북 areaCode=35, 단원별 검색어, 위치 기반 후보를 수집', '관광자원 후보 원자료', '지역 관광자원과 교육 데이터를 결합'],
    ['6. 관광자원 정제', '장소명, 주소, 좌표, 관광분류, 이미지', '중복 contentid 제거, 좌표 없는 데이터 제외, 학교와의 거리 계산, 앱 대분류/소분류 부여', '정제된 장소 후보 목록', '단순 검색 결과가 아니라 수업 후보로 쓸 수 있는 데이터만 정제'],
    ['7. 교육과정 적합성 판정', '정제 장소 + 단원 키워드 사전 + 역사 원자료 태그', '강한 근거(+32), 보조 근거(+12), 관광분류 코드(+34), 세부 주제 점수를 합산하고 최소 점수 미만 제외', '교육과정 통과/제외 판정', '거리보다 교육과정 직접 관련성을 우선'],
    ['8. 추천 결과 생성', '통과 장소 + 세부 주제 판정 결과', '추천 이유, 관련 단원, 관련 주제·차시, 직접 연계 근거, 수업 활동, 핵심 키워드를 생성', '앱 팝업의 AI 교육과정 연결 결과', 'AI 추천이 근거 없는 문장이 아니라 전처리 데이터에 기반'],
    ['9. 검증/보완', '오매칭 사례와 사용자 피드백', '예: 영천문화원은 5학년 1단원에서 제외되도록 deny/min score 보정', '검증 예시표', '심사 질의응답에서 오매칭 개선 과정을 설명 가능'],
  ];
}

function scoringRows() {
  return [
    ['구성 요소', '점수/처리', '적용 조건', '의미'],
    ['강한 근거 키워드', '+32', '장소명/주소/상세 설명에 단원 핵심 장소어가 포함될 때', '직접 연계성이 높은 증거'],
    ['보조 키워드', '+12', '단원과 관련된 일반 개념어가 포함될 때', '연계 가능성을 보조하는 증거'],
    ['관광 API 분류 코드', '+34', '단원 성격과 맞는 Tour API cat3 또는 cat1 코드일 때', '장소명이 애매해도 공공데이터 분류로 보완'],
    ['세부 주제 점수', 'terms +3, searchTerms +2, placeTypes +2', '세부 주제별 키워드가 장소 정보에 포함될 때', '큰 단원 안에서 어느 주제/차시와 연결되는지 결정'],
    ['제외 패턴', '0점 처리 가능', '단원과 다른 성격의 장소가 강한 근거 없이 포함될 때', '예: 5학년 우리나라 국토 여행에서 문화원/시장/공방 제외'],
    ['최소 통과 점수', '단원별 22~34점', '단원 유형별 socialUnitProfile에 정의', '애매한 장소를 추천 결과에서 제거'],
    ['거리 보정', '+18/+12/+6', '최소 통과 후 8km/15km/30km 이내일 때', '교육과정 적합성을 먼저 통과한 장소만 거리 보정'],
    ['지역 보정', '+18', '학교가 있는 시군과 장소 시군이 같을 때', '지역화 수업 편의성 반영'],
  ];
}

function schemaRows() {
  return [
    ['원자료 필드', '출처', '전처리 방식', '앱 활용'],
    ['contentid', 'Tour API', '중복 제거 기준 ID', '동일 장소 중복 노출 방지'],
    ['title', 'Tour API', '공백 정리, 키워드 매칭 대상', '장소명 표시, 단원 키워드 판정'],
    ['addr1', 'Tour API', '경북 시군명 추출, 키워드 매칭 대상', '주소 표시, 지역/장소유형 판정'],
    ['mapx/mapy', 'Tour API', '숫자 변환, 좌표 없으면 제외', '지도 마커, 거리 계산'],
    ['cat1/cat2/cat3', 'Tour API', '관광분류 코드 그룹과 비교', '자연/역사/과학 등 공공데이터 분류 보정'],
    ['contenttypeid', 'Tour API', '관광지/문화시설 등 유형 확인', '후보 수집 범위 제한'],
    ['sigungucode', 'Tour API', '학교 시군 코드와 비교', '같은 시군일 때 지역화 보정'],
    ['firstimage/firstimage2', 'Tour API', '이미지 존재 여부 확인', '카드뉴스/장소 카드 이미지'],
    ['overview', 'Tour API detailCommon2', 'HTML 태그 제거, 420자 요약', '팝업 장소 설명과 추가 AI 연결 근거'],
    ['SCHUL_NM', 'NEIS API', '학교명 검색 결과 표시', '사용자 기준 학교 선택'],
    ['ORG_RDNMA', 'NEIS API', '주소 지오코딩', '학교 위치 기준점 생성'],
  ];
}

function aiRows() {
  return [
    ['AI/자동화 활용 지점', '입력', '처리', '출력', '통제 장치'],
    ['교과서 PDF 분석', '사회 교과서 PDF', '단원명, 주제, 차시, 핵심 개념을 구조화', '교육과정 데이터셋', '원자료 출처를 명시하고 큰 단원명은 교과서 표현 유지'],
    ['단원-장소 매칭', '단원 키워드 + 관광자원 필드', '키워드/분류코드/제외 규칙 기반 점수화', '추천/제외 판정', '최소 점수 미만 제외, 오매칭 사례 검증'],
    ['팝업 설명 생성', '선택 장소 + 선택 단원 + 매칭 근거', '관련 주제·차시, 직접 연계 근거, 수업 활동 문장화', 'AI 교육과정 연결 카드', '단원별 프로필에 없는 문장은 생성하지 않도록 제한'],
    ['역사 원자료 연결', '5학년 2학기 역사 성취기준 + 국사편찬위원회 사료 DB 태그', '삼국유사·금석문·고려사·감시대상인물카드 등과 장소 유형을 교차 매핑', '역사 단원 추천 이유와 교육공공데이터 근거', '공식 성취기준과 공공 사료명만 팝업에 표시'],
    ['카드뉴스 구성', '장소 이미지 + 교육과정 연결 결과', '이미지와 요약 문장을 카드 형태로 재구성', '장소별 카드뉴스', '출처 이미지 사용, 이미지 없으면 대체 표시'],
  ];
}

function getUnit(app, grade, unitName) {
  return (app.CUR[String(grade)].사회 || []).find((u) => u.unit === unitName || u.name === unitName);
}

function validationRows(app) {
  const rows = [['검증 목적', '학년', '선택 단원', '장소명', '주소', '예상 결과', '실제 점수', '관련 주제', '직접 연계 근거']];
  const cases = [
    {
      purpose: '오매칭 제거: 문화원은 국토 여행 단원과 직접 관련 부족',
      grade: '5',
      unit: '1. 우리나라 국토 여행',
      expected: '추천 제외',
      place: { title: '영천문화원', addr1: '경상북도 영천시', cat1: 'A02', cat3: 'A02060900', cat: '문화유산', sub: '', dist: 2, sigunguCode: '15' },
    },
    {
      purpose: '정상 매칭: 독도/울릉 자원은 국토 여행 단원과 연결',
      grade: '5',
      unit: '1. 우리나라 국토 여행',
      expected: '추천 유지',
      place: { title: '울릉도 독도박물관', addr1: '경상북도 울릉군', cat1: 'A02', cat3: 'A02060100', cat: '인물역사', sub: '박물관', dist: 120, sigunguCode: '17' },
    },
    {
      purpose: '정상 매칭: 시장은 경제활동·교류 단원과 연결',
      grade: '4',
      unit: '3. 경제활동과 지역 간 교류',
      expected: '추천 유지',
      place: { title: '영천공설시장', addr1: '경상북도 영천시', cat1: 'A02', cat3: 'A02030400', cat: '생활체험', sub: '전통시장', dist: 4, sigunguCode: '15' },
    },
    {
      purpose: '정상 매칭: 서원은 지역 국가유산 단원과 연결',
      grade: '4',
      unit: '2. 우리 지역의 국가유산',
      expected: '추천 유지',
      place: { title: '도산서원', addr1: '경상북도 안동시 도산면', cat1: 'A02', cat3: 'A02010700', cat: '문화유산', sub: '서원·향교', dist: 20, sigunguCode: '11' },
    },
    {
      purpose: '정상 매칭: 경주 고대 유적은 5학년 2학기 고대 생활 추론과 연결',
      grade: '5',
      unit: '4. 유적과 유물로 살펴본 옛 사람들의 생활',
      expected: '추천 유지',
      place: { title: '경주 대릉원 천마총', addr1: '경상북도 경주시 황남동', cat1: 'A02', cat3: 'A02010700', cat: '문화유산', sub: '고분', dist: 40, sigunguCode: '2' },
    },
    {
      purpose: '정상 매칭: 안동 서원은 조선 유교 문화와 생활 모습 단원에 연결',
      grade: '5',
      unit: '5. 달라지는 시대, 변화하는 생활 모습',
      expected: '추천 유지',
      place: { title: '안동 도산서원', addr1: '경상북도 안동시 도산면', cat1: 'A02', cat3: 'A02010700', cat: '문화유산', sub: '서원·향교', dist: 25, sigunguCode: '11' },
    },
    {
      purpose: '정상 매칭: 독립운동 기념관은 식민 통치와 저항 단원에 연결',
      grade: '5',
      unit: '6. 식민 통치와 저항, 전쟁이 바꾼 사회와 생활',
      expected: '추천 유지',
      place: { title: '경상북도독립운동기념관', addr1: '경상북도 안동시 임하면', cat1: 'A02', cat3: 'A02060200', cat: '인물역사', sub: '독립운동 유적', dist: 30, sigunguCode: '11' },
    },
  ];
  cases.forEach((c) => {
    const u = getUnit(app, c.grade, c.unit);
    app.setState(c.grade, '사회');
    const score = app.scorePlace(c.place, u, null);
    const d = app.curriculumData(c.place, u, '');
    rows.push([c.purpose, `${c.grade}학년`, c.unit, c.place.title, c.place.addr1, c.expected, score, d.topic, d.direct]);
  });
  return rows;
}

async function fetchTour(app, keyword) {
  const url = `${app.TOUR}/searchKeyword2?serviceKey=${app.TK}&keyword=${encodeURIComponent(keyword)}&areaCode=35&numOfRows=8&pageNo=1&MobileOS=ETC&MobileApp=${app.APP}&_type=json`;
  try {
    const res = await fetch(url);
    if (!res.ok) return [];
    const data = await res.json();
    const item = data?.response?.body?.items?.item;
    if (!item) return [];
    return Array.isArray(item) ? item : [item];
  } catch {
    return [];
  }
}

async function candidateRows(app) {
  const rows = [['학년', '교과', '선택 단원', '검색어', 'contentid', '장소명', '주소', 'contenttypeid', 'cat1', 'cat2', 'cat3', '시군코드', '좌표유무', '이미지유무', '앱 대분류', '앱 소분류', '경북중심 기준 거리(km)', '교육과정 점수', '최소 기준', '선별 결과', '관련 주제·차시', '근거 키워드', '직접 연계 근거']];
  const units = [];
  Object.entries(app.CUR).forEach(([grade, subjects]) => {
    (subjects.사회 || []).forEach((unit) => units.push({ grade, unit }));
  });

  for (const { grade, unit } of units) {
    const seeds = uniq([...(unit.searchTerms || []), ...(unit.placeTypes || []), ...(unit.terms || [])])
      .filter((v) => String(v).length >= 2)
      .slice(0, 5);
    const seen = new Set();
    for (const seed of seeds) {
      const items = await fetchTour(app, seed);
      for (const item of items) {
        if (!item?.contentid || seen.has(item.contentid)) continue;
        seen.add(item.contentid);
        app.setState(grade, '사회');
        const hasCoord = Boolean(item.mapx && item.mapy);
        let place = null;
        let score = 0;
        let min = '';
        let result = '좌표 없음 제외';
        let topic = '';
        let evidence = '';
        let direct = '';
        if (hasCoord) {
          place = app.buildPlace(item, app.GBK, unit, null);
          const rel = app.curriculumRelevance(place, unit, '');
          const data = app.curriculumData(place, unit, '');
          score = place.score;
          min = rel.min;
          result = score > 0 ? '통과' : '교육과정 근거 부족 제외';
          topic = `${data.topic} · ${data.lesson}`;
          evidence = data.keywords;
          direct = data.direct;
        }
        rows.push([
          `${grade}학년`, '사회', unit.unit || unit.name, seed, item.contentid || '', item.title || '', item.addr1 || '',
          item.contenttypeid || '', item.cat1 || '', item.cat2 || '', item.cat3 || '', item.sigungucode || '',
          hasCoord ? '있음' : '없음', item.firstimage || item.firstimage2 ? '있음' : '없음',
          place?.cat || '', place?.sub || '', place?.dist ?? '', score, min, result, topic, evidence, direct,
        ]);
      }
    }
  }
  return rows;
}

function colName(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - m) / 26);
  }
  return s;
}

function xmlEscape(value) {
  return String(value ?? '')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .slice(0, 32000)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function sheetXml(rows) {
  const maxCols = rows.reduce((m, r) => Math.max(m, r.length), 1);
  const lastRef = `${colName(maxCols)}${Math.max(rows.length, 1)}`;
  const cols = Array.from({ length: maxCols }, (_, i) => `<col min="${i + 1}" max="${i + 1}" width="${i === 0 ? 18 : 28}" customWidth="1"/>`).join('');
  const body = rows.map((row, rIdx) => {
    const r = rIdx + 1;
    const cells = row.map((v, cIdx) => {
      const ref = `${colName(cIdx + 1)}${r}`;
      if (typeof v === 'number' && Number.isFinite(v)) return `<c r="${ref}"><v>${v}</v></c>`;
      return `<c r="${ref}" t="inlineStr"><is><t>${xmlEscape(v)}</t></is></c>`;
    }).join('');
    return `<row r="${r}">${cells}</row>`;
  }).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols>${cols}</cols>
  <sheetData>${body}</sheetData>
  <autoFilter ref="A1:${lastRef}"/>
</worksheet>`;
}

function workbookXml(sheets) {
  const sheetTags = sheets.map((s, i) => `<sheet name="${xmlEscape(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheetTags}</sheets>
</workbook>`;
}

function workbookRels(sheets) {
  const rels = sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
}

function contentTypes(sheets) {
  const overrides = sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${overrides}
</Types>`;
}

function rootRels() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
}

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32(buf) {
  let c = 0xFFFFFFFF;
  for (const b of buf) c = crcTable[(c ^ b) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

function dosTimeDate(date = new Date()) {
  const year = Math.max(date.getFullYear(), 1980);
  const time = (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2);
  const day = ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
  return { time, day };
}

function zip(files) {
  const chunks = [];
  const central = [];
  let offset = 0;
  const { time, day } = dosTimeDate();
  files.forEach((file) => {
    const name = Buffer.from(file.name, 'utf8');
    const data = Buffer.from(file.data, 'utf8');
    const crc = crc32(data);
    const local = Buffer.alloc(30 + name.length);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(0, 8);
    local.writeUInt16LE(time, 10);
    local.writeUInt16LE(day, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(name.length, 26);
    local.writeUInt16LE(0, 28);
    name.copy(local, 30);
    chunks.push(local, data);

    const cent = Buffer.alloc(46 + name.length);
    cent.writeUInt32LE(0x02014b50, 0);
    cent.writeUInt16LE(20, 4);
    cent.writeUInt16LE(20, 6);
    cent.writeUInt16LE(0x0800, 8);
    cent.writeUInt16LE(0, 10);
    cent.writeUInt16LE(time, 12);
    cent.writeUInt16LE(day, 14);
    cent.writeUInt32LE(crc, 16);
    cent.writeUInt32LE(data.length, 20);
    cent.writeUInt32LE(data.length, 24);
    cent.writeUInt16LE(name.length, 28);
    cent.writeUInt16LE(0, 30);
    cent.writeUInt16LE(0, 32);
    cent.writeUInt16LE(0, 34);
    cent.writeUInt16LE(0, 36);
    cent.writeUInt32LE(0, 38);
    cent.writeUInt32LE(offset, 42);
    name.copy(cent, 46);
    central.push(cent);
    offset += local.length + data.length;
  });
  const centralSize = central.reduce((n, b) => n + b.length, 0);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(files.length, 8);
  end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(centralSize, 12);
  end.writeUInt32LE(offset, 16);
  end.writeUInt16LE(0, 20);
  return Buffer.concat([...chunks, ...central, end]);
}

function makeXlsx(sheets, outPath) {
  const files = [
    { name: '[Content_Types].xml', data: contentTypes(sheets) },
    { name: '_rels/.rels', data: rootRels() },
    { name: 'xl/workbook.xml', data: workbookXml(sheets) },
    { name: 'xl/_rels/workbook.xml.rels', data: workbookRels(sheets) },
    ...sheets.map((s, i) => ({ name: `xl/worksheets/sheet${i + 1}.xml`, data: sheetXml(s.rows) })),
  ];
  fs.writeFileSync(outPath, zip(files));
}

async function main() {
  const app = readApp();
  console.log('앱 데이터 추출 완료');
  const candidates = await candidateRows(app);
  console.log(`관광자원 후보 샘플 ${Math.max(candidates.length - 1, 0)}건 생성`);

  const sheets = [
    { name: '00_데이터출처', rows: sourceRows(app) },
    { name: '01_처리프로세스', rows: processRows() },
    { name: '02_단원선택UI', rows: unitRows(app) },
    { name: '03_교육과정세부주제', rows: topicRows(app) },
    { name: '04_단원별선별기준', rows: profileRows(app) },
    { name: '05_장소분류사전', rows: categoryRows(app) },
    { name: '06_API원자료스키마', rows: schemaRows() },
    { name: '07_추천점수규칙', rows: scoringRows() },
    { name: '08_장소후보전처리샘플', rows: candidates },
    { name: '09_검증예시', rows: validationRows(app) },
    { name: '10_AI활용기록', rows: aiRows() },
  ];
  makeXlsx(sheets, path.resolve(OUT_FILE));
  console.log(`완료: ${OUT_FILE}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
