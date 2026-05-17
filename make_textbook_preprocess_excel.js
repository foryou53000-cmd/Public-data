const fs = require('fs');
const path = require('path');

const TEXT_DIR = 'textbook_extracted';
const OUT_FILE = '사회교과서_단원차시_전처리.xlsx';

const BOOKS = [
  { grade: 3, semester: 1, pdf: '초_사회3-1(김정인)_교과서.pdf', txt: '초_사회3-1(김정인)_교과서.txt' },
  { grade: 3, semester: 2, pdf: '초_사회3-2(김정인)_교과서.pdf', txt: '초_사회3-2(김정인)_교과서.txt' },
  { grade: 4, semester: 1, pdf: '초_사회4-1(김정인)_교과서.pdf', txt: '초_사회4-1(김정인)_교과서.txt' },
  { grade: 4, semester: 2, pdf: '초_사회4-2(김정인)_교과서.pdf', txt: '초_사회4-2(김정인)_교과서.txt' },
  { grade: 5, semester: 1, pdf: '초_사회5-1(김정인)_교과서.pdf', txt: '초_사회5-1(김정인)_교과서.txt' },
  { grade: 6, semester: 1, pdf: '초_사회6-1(김정인)_교과서.pdf', txt: '초_사회6-1(김정인)_교과서.txt' },
];

const UNIT_TITLES = {
  '3-1': ['우리가 사는 곳', '일상에서 만나는 과거'],
  '3-2': ['사회 변화와 다양한 문화', '옛날과 오늘날의 생활 모습'],
  '4-1': ['지도로 만나는 우리 지역', '우리 지역의 국가유산', '경제활동과 지역 간 교류'],
  '4-2': ['민주주의와 자치', '지역문제를 해결하고 지역을 알리는 노력', '다양한 환경과 삶의 모습'],
  '5-1': ['우리나라 국토 여행', '우리나라 지리 탐구', '법과 인권의 보장'],
  '6-1': ['평화 통일을 위한 노력, 민주화와 산업화', '민주주의와 시민 참여', '지구, 대륙 그리고 국가들'],
};

const TOPIC_TITLES = {
  '3-1': {
    '1-1': '우리 생활 속 여러 장소에 대한 경험과 느낌',
    '1-2': '우리가 만드는 살기 좋은 곳',
    '2-1': '시간의 흐름과 우리',
    '2-2': '오래된 것이 알려 주는 과거',
    '2-3': '지역의 변화와 달라진 생활 모습',
  },
  '3-2': {
    '1-1': '사회 변화로 달라진 생활 모습',
    '1-2': '다양한 문화에 대한 이해와 존중',
    '2-1': '옛날과 오늘날의 풍습',
    '2-2': '교통의 변화로 달라진 생활 모습',
    '2-3': '통신수단의 변화로 달라진 생활 모습',
  },
  '4-1': {
    '1-1': '지도의 읽기와 활용',
    '1-2': '우리 지역의 위치와 특징',
    '2-1': '지역의 국가유산',
    '2-2': '지역의 박물관, 기념관, 유적지',
    '3-1': '경제활동과 합리적 선택',
    '3-2': '교류하며 발전하는 우리 지역',
  },
  '4-2': {
    '1-1': '학교 자치와 민주주의',
    '1-2': '주민 자치와 주민 참여',
    '2-1': '지역문제를 해결하려는 노력',
    '2-2': '지역을 알리려는 노력',
    '3-1': '지역의 다양한 환경과 변화',
    '3-2': '도시의 특징과 생활 모습',
  },
  '5-1': {
    '1-1': '우리나라 지형 여행',
    '1-2': '우리 땅 독도',
    '2-1': '우리나라 기후 탐구',
    '2-2': '우리나라 인구 분포 탐구',
    '3-1': '우리 생활 속 법과 인권',
    '3-2': '인권을 존중하는 우리',
  },
  '6-1': {
    '1-1': '평화 통일을 위한 노력',
    '1-2': '민주화와 산업화로 달라진 생활 문화',
    '2-1': '민주주의와 선거',
    '2-2': '국가기관이 하는 일',
    '2-3': '민주주의와 미디어',
    '3-1': '지구본과 지도로 보는 세계',
    '3-2': '세계의 대륙, 대양, 나라',
  },
};

const ORDINALS = ['하나', '둘', '셋', '넷', '다섯', '여섯', '일곱', '여덟', '아홉', '열', '열하나', '열둘'];

const KEYWORDS = [
  '지도', '방위', '방위표', '축척', '범례', '기호', '등고선', '디지털 영상 지도', '안내도', '노선도', '위치', '지리 정보',
  '지역', '우리 지역', '고장', '행정구역', '면적', '인구', '산', '산지', '하천', '강', '해안', '기온', '강수량', '기후',
  '자연환경', '인문환경', '환경', '생활 모습', '도시', '촌락', '교통', '통신',
  '국가유산', '문화유산', '무형유산', '박물관', '기념관', '유적지', '유물', '유적', '서원', '향교', '사찰', '고택', '종택',
  '경제활동', '생산', '소비', '희소성', '합리적 선택', '시장', '교류', '지역 간 교류', '상품', '기술', '정보',
  '민주주의', '자치', '주민 참여', '지역문제', '권리', '인권', '법', '헌법', '공공기관',
  '국토', '영토', '독도', '울릉도', '대륙', '대양', '세계지도', '위도', '경도', '세계 여러 나라',
  '역사', '선사 시대', '고조선', '삼국', '신라', '가야', '고려', '조선', '일제', '독립운동', '전쟁', '평화', '통일',
];

function clean(value = '') {
  return String(value).replace(/\u0000/g, '').replace(/\s+/g, ' ').trim();
}

function compact(value = '') {
  return clean(value).replace(/[^\p{Script=Hangul}A-Za-z0-9·~.,!?()\[\]\s-]/gu, '').trim();
}

function bookKey(book) {
  return `${book.grade}-${book.semester}`;
}

function knownUnitTitle(book, unitNo) {
  return UNIT_TITLES[bookKey(book)]?.[unitNo - 1] || '';
}

function knownTopicTitle(book, unitNo, topicNo) {
  return TOPIC_TITLES[bookKey(book)]?.[`${unitNo}-${topicNo}`] || '';
}

function fixKoreanSpacing(value = '') {
  return clean(value)
    .replace(/우\s+리/g, '우리')
    .replace(/지\s+도의/g, '지도의')
    .replace(/지\s+도/g, '지도')
    .replace(/다\s+양한/g, '다양한')
    .replace(/행\s+정구역/g, '행정구역')
    .replace(/지\s+리/g, '지리')
    .replace(/국\s+가유산/g, '국가유산')
    .replace(/지\s+역/g, '지역')
    .replace(/자\s+원/g, '자원')
    .replace(/주\s+변/g, '주변')
    .replace(/일\s+상/g, '일상')
    .replace(/시\s+간/g, '시간')
    .replace(/오\s+래된/g, '오래된')
    .replace(/옛\s+날/g, '옛날')
    .replace(/지\s+명/g, '지명')
    .replace(/교\s+통/g, '교통')
    .replace(/통\s+신/g, '통신')
    .replace(/저\s+출산/g, '저출산')
    .replace(/민\s+주주의/g, '민주주의')
    .replace(/학\s+교생활/g, '학교생활')
    .replace(/주\s+민/g, '주민')
    .replace(/생\s+활/g, '생활')
    .replace(/여\s+러/g, '여러')
    .replace(/환\s+경/g, '환경')
    .replace(/사\s+람/g, '사람')
    .replace(/도\s+시/g, '도시');
}

function normalizeOrdinals(text) {
  return clean(text)
    .replace(/하\s*나/g, '하나')
    .replace(/다\s*섯/g, '다섯')
    .replace(/여\s*섯/g, '여섯')
    .replace(/여\s*덟/g, '여덟')
    .replace(/아\s*홉/g, '아홉')
    .replace(/열\s*하\s*나/g, '열하나')
    .replace(/열\s*둘/g, '열둘');
}

function uniq(arr) {
  return Array.from(new Set((arr || []).filter(Boolean)));
}

function join(arr) {
  return uniq(arr).join(', ');
}

function textbookLabel(book) {
  return `${book.grade}학년 ${book.semester}학기`;
}

function readBook(book) {
  const txtPath = path.join(TEXT_DIR, book.txt);
  const text = fs.readFileSync(txtPath, 'utf8');
  const pages = text.split('\f').map((raw, idx) => ({
    pdfPage: idx + 1,
    raw,
    text: compact(raw),
  }));
  return { ...book, text, pages };
}

function isStandaloneKoreanTerm(text, term) {
  return new RegExp(`(^|[\\s,·])${term}(의|을|은|이|과|와|로|에서|,|·|\\s|$)`).test(text);
}

function keywordHit(text, keyword) {
  if (keyword === '산') return isStandaloneKoreanTerm(text, '산');
  if (keyword === '강') return isStandaloneKoreanTerm(text, '강') || /낙동강|한강|금강|영산강|섬진강/.test(text);
  if (keyword === '법') return isStandaloneKoreanTerm(text, '법') || /법과|법의|법을|법이/.test(text);
  return text.includes(keyword);
}

function deriveKeywords(...parts) {
  const text = compact(parts.join(' '));
  const hits = KEYWORDS.filter((k) => keywordHit(text, k));
  const extra = [];
  if (/지도|방위|축척|범례|등고선|디지털/.test(text)) extra.push('지도 읽기');
  if (/위치|면적|인구|지리 정보|행정구역/.test(text)) extra.push('지역 정보 조사');
  if (isStandaloneKoreanTerm(text, '산') || /산지|하천|해안|기온|강수량|자연환경|지형|낙동강|한강|금강|영산강|섬진강/.test(text)) extra.push('지형·환경 관찰');
  if (/국가유산|문화유산|무형유산|박물관|기념관|유적|유물/.test(text)) extra.push('국가유산 답사');
  if (/경제|생산|소비|시장|교류|상품|기술|정보/.test(text)) extra.push('경제 교류');
  if (/민주|자치|주민|지역문제|공공기관|권리|인권/.test(text)) extra.push('시민 참여');
  if (/독도|울릉|영토|국토|대륙|대양|세계지도|위도|경도/.test(text)) extra.push('공간 자료');
  return uniq([...hits, ...extra]).slice(0, 18);
}

function suggestPlaceTypes(keywords) {
  const text = keywords.join(' ');
  const types = [];
  const excludes = [];
  if (/지도|방위|축척|범례|등고선|위치|지리/.test(text)) types.push('전망대', '공원 안내도', '지질공원', '산·하천 조망지', '지도·기상 관련 시설');
  if (/(^|[, ]+)(산|산지)([, ]+|$)|하천|해안|기후|강수량|자연환경|지형/.test(text)) types.push('국립공원·도립공원', '대표 산', '하천 공원', '해수욕장·갯벌', '기상과학관');
  if (/국가유산|문화유산|박물관|기념관|유적/.test(text)) types.push('국가유산', '지역 박물관', '기념관', '유적지');
  if (/경제|시장|교류|생산|소비/.test(text)) types.push('대표 시장', '산업·물류 거점', '항만·철도 거점');
  if (/민주|자치|주민|공공기관|지역문제/.test(text)) types.push('공공기관', '주민 참여 시설', '안전체험관', '생활권 공원');
  if (/독도|울릉|국토|영토/.test(text)) types.push('독도 관련 박물관·교육관', '해양·영토 교육 시설');
  if (/경제|교류/.test(text)) excludes.push('소규모 공방', '마을 단위 체험장', '오일장', '특산물 판매장');
  if (/지도|위치|지리/.test(text)) excludes.push('지도·위치 관찰 근거가 약한 실내 체험 위주 장소');
  return { include: uniq(types).join(', '), exclude: uniq(excludes).join(', ') };
}

function extractUnitStarts(pages) {
  const units = [];
  for (const page of pages) {
    const text = page.text;
    const head = text.slice(0, 260);
    const m = head.match(/^(\d{1,2})\s+([가-힣A-Za-z0-9·,\s]+?)\s+1\s+[가-힣]/);
    if (!m) continue;
    const title = clean(m[2]).replace(/부분을.*$/, '').trim();
    if (title.length < 4 || title.length > 34) continue;
    if (!/[우리지역국가경제민주문제환경도시국토인권유적평화세계정치사회생활역사문화]/.test(title)) continue;
    if (units.some((u) => u.no === Number(m[1]) && u.title === title)) continue;
    units.push({ no: Number(m[1]), title, pdfPage: page.pdfPage, source: text.slice(0, 400) });
  }
  return units.sort((a, b) => a.pdfPage - b.pdfPage);
}

function findCurrentUnit(units, pdfPage) {
  let cur = null;
  for (const u of units) {
    if (u.pdfPage <= pdfPage) cur = u;
    else break;
  }
  return cur;
}

function extractObjectives(text) {
  const idx = text.indexOf('이 주제를 배우면 나는');
  if (idx < 0) return [];
  const after = text.slice(idx + '이 주제를 배우면 나는'.length, idx + 520);
  const sentences = after.match(/[가-힣A-Za-z0-9·,\s]+?(?:수 있어요|할 수 있어요|알 수 있어요|설명할 수 있어요|이해할 수 있어요)/g) || [];
  return uniq(sentences.map((s) => fixKoreanSpacing(s)).filter((s) => s.length >= 8)).slice(0, 4);
}

function extractTopicStarts(book, units) {
  const candidates = [];
  for (const page of book.pages) {
    const text = page.text;
    const idx = text.indexOf('이 주제를 배우면 나는');
    if (idx < 0 || page.pdfPage < 8) continue;
    const before = text.slice(Math.max(0, idx - 140), idx).replace(/부분을.*$/, '');
    const matches = [...before.matchAll(/(?:^|\s)(\d{1,2})\s+([가-힣A-Za-z0-9·,\s]{2,70})$/g)];
    let topicNo = '';
    let topicTitle = '';
    if (matches.length) {
      const m = matches[matches.length - 1];
      topicNo = Number(m[1]);
      topicTitle = clean(m[2]).replace(/^주제\s*/, '').trim();
    } else {
      const lines = page.raw.split(/\r?\n/).map(clean).filter(Boolean);
      const line = lines.find((v) => /^\d{1,2}\s+[가-힣]/.test(v));
      if (line) {
        const m = line.match(/^(\d{1,2})\s+(.{2,70})/);
        topicNo = m ? Number(m[1]) : '';
        topicTitle = m ? clean(m[2]) : '';
      }
    }
    if (!topicTitle || /이 주제를|부분을|단원/.test(topicTitle)) continue;
    candidates.push({
      page,
      idx,
      topicNo,
      rawTopicTitle: fixKoreanSpacing(topicTitle),
      objectives: extractObjectives(text),
    });
  }

  const rows = [];
  let unitNo = 0;
  let previousTopicNo = 0;
  for (const item of candidates.sort((a, b) => a.page.pdfPage - b.page.pdfPage)) {
    if (!unitNo || (Number.isFinite(item.topicNo) && item.topicNo <= previousTopicNo)) unitNo += 1;
    previousTopicNo = Number.isFinite(item.topicNo) ? item.topicNo : previousTopicNo + 1;

    const fallbackUnit = findCurrentUnit(units, item.page.pdfPage);
    const unitTitle = knownUnitTitle(book, unitNo) || fallbackUnit?.title || '';
    const topicTitle = knownTopicTitle(book, unitNo, item.topicNo) || item.rawTopicTitle;
    const objectives = item.objectives;
    const keyText = `${unitTitle} ${topicTitle} ${objectives.join(' ')}`;
    const keywords = deriveKeywords(keyText);
    const place = suggestPlaceTypes(keywords);
    rows.push({
      grade: book.grade,
      semester: book.semester,
      unitNo,
      unitTitle,
      topicNo: item.topicNo,
      topicTitle,
      pdfPage: item.page.pdfPage,
      objectives,
      core: objectives.length ? objectives.join(' / ') : `${topicTitle}의 핵심 개념과 활동을 교과서 원문에서 확인`,
      keywords,
      include: place.include,
      exclude: place.exclude,
      source: item.page.text.slice(Math.max(0, item.idx - 120), item.idx + 520),
      confidence: objectives.length ? '상' : '중',
    });
  }
  return rows;
}

function unitsFromTopics(topics) {
  const byUnit = new Map();
  for (const topic of topics) {
    const key = `${topic.unitNo}-${topic.unitTitle}`;
    if (!byUnit.has(key)) {
      byUnit.set(key, {
        no: topic.unitNo,
        title: topic.unitTitle,
        pdfPage: topic.pdfPage,
        source: `${topic.pdfPage}쪽 주제 "${topic.topicTitle}"에서 단원 "${topic.unitTitle}"로 매핑`,
      });
    }
  }
  return Array.from(byUnit.values()).sort((a, b) => a.pdfPage - b.pdfPage);
}

function extractActivityPhrases(text) {
  const phrases = [];
  const re = /([가-힣A-Za-z0-9·,\s]{2,34}?(?:알아보기|살펴보기|비교하기|조사하기|정리하기|만들기|읽기|찾기|표현하기|이야기하기|탐구하기|활용하기))/g;
  for (const m of text.matchAll(re)) {
    const phrase = clean(m[1]);
    if (phrase.length >= 4 && phrase.length <= 40 && !/배움 확인|준비물|놀이 방법/.test(phrase)) phrases.push(phrase);
  }
  return uniq(phrases).slice(0, 8);
}

function cleanLessonTitle(value = '') {
  return fixKoreanSpacing(value)
    .replace(/^애니메이션\s*/, '')
    .replace(/\s*해 보기\s*(?:의사소통 및 정보 활용 능력|정보 활용 능력|협업 능력)?\s*/g, ' ')
    .replace(/\s*(?:의사소통 및 정보 활용 능력|정보 활용 능력|협업 능력)\s*/g, ' ')
    .replace(/\s*리 들 사이\s*/g, ' ')
    .replace(/지역의역사를알수있는장소/g, '지역의 역사를 알 수 있는 장소')
    .replace(/우리지역/g, '우리 지역')
    .replace(/\s+/g, ' ')
    .trim();
}

function extractLessonHeadings(book, units, topics) {
  const rows = [];
  const seen = new Set();
  for (const page of book.pages) {
    if (page.pdfPage < 8) continue;
    const text = normalizeOrdinals(page.text);
    const topic = [...topics].reverse().find((t) => t.pdfPage <= page.pdfPage);
    const unit = findCurrentUnit(units, page.pdfPage) || (topic ? { no: topic.unitNo, title: topic.unitTitle } : null);
    const local = [];
    const splitHeadingRe = /(?:애니메이션\s+)?하\s+(?:애니메이션\s+)?([가-힣A-Za-z0-9·,\s]{2,70}?)\s+나\s+([가-힣A-Za-z0-9·,\s]{2,60}?(?:까요|볼까요|있을까요|나타낼까요|필요할까요|무엇일까요|해 볼까요|알아볼까요|살펴볼까요)\??)/g;
    for (const m of text.matchAll(splitHeadingRe)) {
      local.push({ marker: '하나', title: cleanLessonTitle(`${m[1]} ${m[2]}`) });
    }
    const headingRe = /(하나|둘|셋|넷|다섯|여섯|일곱|여덟|아홉|열|열하나|열둘)\s+([가-힣A-Za-z0-9·,\s]{3,80}?(?:까요|볼까요|있을까요|나타낼까요|필요할까요|무엇일까요|해 볼까요|알아볼까요|살펴볼까요)\??)/g;
    for (const m of text.matchAll(headingRe)) {
      local.push({ marker: m[1], title: cleanLessonTitle(m[2]) });
    }
    if (/주제 마무리/.test(text)) local.push({ marker: '마무리', title: '주제 마무리' });
    if (/단원 마무리|정리하기/.test(text) && /놀이 방법|배운 내용|되돌아보기/.test(text)) local.push({ marker: '정리', title: '단원 정리하기' });
    if (/더 나아가기/.test(text)) local.push({ marker: '확장', title: '더 나아가기' });

    const activities = extractActivityPhrases(text);
    for (const item of local) {
      const key = `${book.grade}-${book.semester}-${page.pdfPage}-${item.marker}-${item.title}`;
      if (seen.has(key)) continue;
      seen.add(key);
      const source = page.text.slice(Math.max(0, page.text.indexOf(item.marker) - 120), Math.min(page.text.length, page.text.indexOf(item.marker) + 520));
      const keywords = deriveKeywords(unit?.title, topic?.topicTitle, item.title, activities.join(' '), source);
      const place = suggestPlaceTypes(keywords);
      rows.push({
        grade: book.grade,
        semester: book.semester,
        unitNo: unit?.no || '',
        unitTitle: unit?.title || '',
        topicNo: topic?.topicNo || '',
        topicTitle: topic?.topicTitle || '',
        marker: item.marker,
        title: item.title,
        pdfPage: page.pdfPage,
        activities,
        core: `${item.title}${activities.length ? ` - ${activities.slice(0, 2).join(', ')}` : ''}`,
        keywords,
        include: place.include,
        exclude: place.exclude,
        source: source || page.text.slice(0, 600),
        confidence: item.marker === '마무리' || item.marker === '정리' ? '중' : '상',
      });
    }
  }
  const counters = new Map();
  return rows.map((r) => {
    const group = `${r.grade}-${r.semester}-${r.unitNo}-${r.topicNo || r.unitTitle}`;
    const n = (counters.get(group) || 0) + 1;
    counters.set(group, n);
    return { ...r, autoLessonNo: n };
  });
}

function guideRows() {
  return [
    ['항목', '내용'],
    ['목적', '사용자가 제공한 사회 교과서 PDF 6권을 pdftotext로 페이지 단위 추출하고 단원·주제·차시 후보·목표·핵심 키워드·장소 매칭 키워드를 전처리한 자료'],
    ['대상 PDF', BOOKS.map((b) => b.pdf).join(' / ')],
    ['추출 방식', 'C:\\Program Files\\Git\\mingw64\\bin\\pdftotext.exe -layout -enc UTF-8'],
    ['주의', '차시 후보는 교과서 원문 페이지의 하나/둘/셋 heading, 주제 마무리, 더 나아가기 등을 자동 추출한 값이다. 앱 반영 전 검수 필요 열과 원문 근거를 함께 확인한다.'],
    ['출력 일시', new Date().toLocaleString('ko-KR')],
  ];
}

function buildSheets() {
  const books = BOOKS.map(readBook);
  const validation = [['학년', '학기', 'PDF 파일', '추출 TXT', '총 문자수', 'PDF 페이지수', '내용 있는 페이지수', '추출 상태', '비고']];
  const unitRows = [['학년', '학기', '단원번호', '단원명_자동추출', '시작_PDF쪽', '원문근거']];
  const topicRows = [['학년', '학기', '단원번호', '단원명', '주제번호', '주제명', '시작_PDF쪽', '학습목표', '핵심내용', '추출키워드', '추천장소유형', '제외키워드/주의', '추출신뢰도', '원문근거']];
  const lessonRows = [['학년', '학기', '단원번호', '단원명', '주제번호', '주제명', '자동차시순번', '차시표지', '차시/활동 제목', 'PDF쪽', '활동후보', '핵심내용', '추출키워드', '추천장소유형', '제외키워드/주의', '추출신뢰도', '원문근거']];
  const matchRows = [['학년', '학기', '단원명', '주제/차시명', '구분', '키워드', '추천장소유형', '제외키워드/주의', '근거_PDF쪽']];
  const pageRows = [['학년', '학기', 'PDF쪽', '문자수', '감지표지', '원문앞부분']];

  for (const book of books) {
    const nonempty = book.pages.filter((p) => p.text.length > 50).length;
    validation.push([book.grade, book.semester, book.pdf, path.join(TEXT_DIR, book.txt), book.text.length, book.pages.length, nonempty, '추출완료', '페이지 구분 문자(폼피드) 기준으로 분리']);

    const rawUnits = extractUnitStarts(book.pages);
    const topics = extractTopicStarts(book, rawUnits);
    const units = unitsFromTopics(topics);
    units.forEach((u) => unitRows.push([book.grade, book.semester, u.no, u.title, u.pdfPage, u.source]));

    topics.forEach((t) => {
      topicRows.push([book.grade, book.semester, t.unitNo, t.unitTitle, t.topicNo, t.topicTitle, t.pdfPage, t.objectives.join(' / '), t.core, join(t.keywords), t.include, t.exclude, t.confidence, t.source]);
      matchRows.push([book.grade, book.semester, t.unitTitle, t.topicTitle, '주제', join(t.keywords), t.include, t.exclude, t.pdfPage]);
    });

    const lessons = extractLessonHeadings(book, units, topics);
    lessons.forEach((l) => {
      lessonRows.push([book.grade, book.semester, l.unitNo, l.unitTitle, l.topicNo, l.topicTitle, l.autoLessonNo, l.marker, l.title, l.pdfPage, l.activities.join(' / '), l.core, join(l.keywords), l.include, l.exclude, l.confidence, l.source]);
      matchRows.push([book.grade, book.semester, l.unitTitle, l.title, '차시후보', join(l.keywords), l.include, l.exclude, l.pdfPage]);
    });

    book.pages.forEach((p) => {
      const markers = [];
      if (p.text.includes('이 책의 차례')) markers.push('차례');
      if (p.text.includes('이 주제를 배우면 나는')) markers.push('주제목표');
      if (ORDINALS.some((o) => normalizeOrdinals(p.text).includes(o))) markers.push('차시표지후보');
      if (p.text.includes('주제 마무리')) markers.push('주제마무리');
      if (p.text.includes('더 나아가기')) markers.push('확장');
      pageRows.push([book.grade, book.semester, p.pdfPage, p.text.length, markers.join(', '), p.text.slice(0, 1200)]);
    });
  }

  return [
    { name: '00_안내', rows: guideRows() },
    { name: '01_추출검증', rows: validation },
    { name: '02_단원목록', rows: unitRows },
    { name: '03_주제목표', rows: topicRows },
    { name: '04_차시활동후보', rows: lessonRows },
    { name: '05_장소매칭키워드', rows: matchRows },
    { name: '06_페이지원문색인', rows: pageRows },
  ];
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
  const cols = Array.from({ length: maxCols }, (_, i) => `<col min="${i + 1}" max="${i + 1}" width="${i < 4 ? 16 : 34}" customWidth="1"/>`).join('');
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

function main() {
  const sheets = buildSheets();
  makeXlsx(sheets, path.resolve(OUT_FILE));
  console.log(`완료: ${OUT_FILE}`);
  sheets.forEach((s) => console.log(`${s.name}: ${s.rows.length - 1}건`));
}

if (require.main === module) {
  main();
}

module.exports = { buildSheets, BOOKS, readBook, extractUnitStarts, extractTopicStarts, extractLessonHeadings };
