const fs = require('fs');
const path = require('path');
const vm = require('vm');

const HTML_FILE = '경북체험학습추천 (2).html';
const DATA_FILES = [
  'data/public_data_sources.js',
  'data/gyeongbuk_public_places.js',
  'data/curated_representative_places.js',
];
const OUT_DIR = 'curation_export';

function clean(v = '') {
  return String(v ?? '').replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim();
}

function csvCell(v) {
  const s = clean(v);
  return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

function writeCsv(filename, rows) {
  const body = rows.map((row) => row.map(csvCell).join(',')).join('\r\n');
  fs.writeFileSync(path.join(OUT_DIR, filename), `\uFEFF${body}`, 'utf8');
}

function readApp() {
  const context = {
    console,
    setTimeout,
    clearTimeout,
    fetch: async () => { throw new Error('network disabled in curation export'); },
    localStorage: { getItem() { return '{}'; }, setItem() {} },
    document: {
      getElementById() { return null; },
      querySelectorAll() { return []; },
      addEventListener() {},
    },
  };
  context.window = context;
  context.globalThis = context;
  context.AbortController = global.AbortController;
  context.AbortSignal = global.AbortSignal;
  context.DOMParser = function DOMParser() {};
  vm.createContext(context);

  DATA_FILES.forEach((file) => {
    if (fs.existsSync(file)) vm.runInContext(fs.readFileSync(file, 'utf8'), context, { filename: file });
  });

  const html = fs.readFileSync(HTML_FILE, 'utf8');
  const scripts = [...html.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/gi)]
    .map((m) => m[1])
    .filter((s) => s.trim());
  if (!scripts.length) throw new Error('app script block not found');
  vm.runInContext(`${scripts.at(-1)}
    globalThis.__APP_EXPORT__ = {
      CUR, GBK, SIGUNGU, DAEGU_SIGUNGU,
      fetchAdditionalPublicCandidates,
      trimRecommendationCandidates,
      curriculumData,
      representativeScore,
      recommendationSortValue,
      areaName,
      setState(g, s) { A.gr = g; A.sj = s; }
    };
  `, context, { filename: HTML_FILE });
  return context.__APP_EXPORT__;
}

function uniq(arr) {
  return Array.from(new Set((arr || []).filter(Boolean)));
}

function allUnits(app) {
  const units = [];
  Object.entries(app.CUR).forEach(([grade, subjects]) => {
    Object.entries(subjects).forEach(([subject, list]) => {
      (list || []).forEach((unit) => units.push({ grade, subject, unit }));
    });
  });
  return units;
}

function unitTitle(unit) {
  return unit?.unit || unit?.name || '';
}

function unitSubtopics(unit) {
  return unit?.subtopics?.length ? unit.subtopics : [unit];
}

function cityFromPlace(app, place) {
  const addr = clean(place?.addr1 || '');
  const keys = [...Object.keys(app.SIGUNGU || {}), ...Object.keys(app.DAEGU_SIGUNGU || {})];
  return keys.find((name) => addr.includes(name)) || app.areaName?.(addr) || '';
}

function autoGrade(place, data, representative) {
  const status = clean(place?.publicData?.status || '');
  if (/검수필요|전년도|종료/.test(status)) return '검토필요';
  if (place?.sourceType === 'curatedRepresentative' || place?.publicData?.educationLevel === '직접연계') return '대표후보';
  if (representative >= 85) return '대표후보';
  if ((Number(place?.score) || 0) >= 260 || representative >= 55 || data?.direct) return '우선후보';
  return '보조후보';
}

function autoMemo(place, data, representative) {
  const memo = [];
  const status = clean(place?.publicData?.status || '');
  if (representative >= 85) memo.push('지역 대표성 높음');
  if (place?.publicData?.educationLevel === '직접연계') memo.push('공공데이터 직접연계');
  if (data?.direct) memo.push('성취기준 직접 근거 확인');
  if (/검수필요|전년도|종료/.test(status)) memo.push(`상태 확인 필요: ${status}`);
  if (!clean(place?.publicData?.overview || place?.sourceMeta || '')) memo.push('장소 소개 보강 필요');
  return memo.join(' / ') || '선별 검토';
}

function candidateRows(app, units) {
  const header = [
    '검토결과(대표/보조/삭제)',
    '검토메모',
    '자동등급',
    '자동메모',
    '단원내우선순위',
    '학년',
    '교과',
    '단원',
    '세부주제',
    '차시',
    '학습목표/성취기준',
    '활동',
    '장소명',
    '시군',
    '주소',
    '장소분류',
    '세부분류',
    '거리(대구경북기준km)',
    '추천점수',
    '대표성점수',
    '정렬점수',
    '연계키워드',
    '추천이유',
    '직접연계근거',
    '장소소개/원자료요약',
    '출처',
    '상태',
    '교육연계수준',
    '원자료ID',
    '출처URL',
  ];
  const rows = [header];
  units.forEach(({ grade, subject, unit }) => {
    app.setState(grade, subject);
    const candidates = app.fetchAdditionalPublicCandidates(unit, app.GBK, null)
      .filter((place) => Number(place?.score) > 0);
    const sorted = app.trimRecommendationCandidates(candidates, unit, null);
    sorted.forEach((place, index) => {
      const overview = clean(place?.publicData?.overview || place?.memorial?.overview || '');
      const data = app.curriculumData(place, unit, overview);
      const rep = app.representativeScore(place);
      const sortValue = app.recommendationSortValue(place);
      rows.push([
        '',
        '',
        autoGrade(place, data, rep),
        autoMemo(place, data, rep),
        index + 1,
        `${grade}학년`,
        subject,
        unitTitle(unit),
        data?.topic || '',
        data?.lesson || '',
        data?.standard || '',
        data?.activity || unit?.activity || '',
        place?.title || '',
        cityFromPlace(app, place),
        place?.addr1 || '',
        place?.cat || '',
        place?.sub || '',
        place?.dist ?? '',
        place?.score ?? '',
        rep,
        Number.isFinite(sortValue) ? Math.round(sortValue * 10) / 10 : '',
        data?.keywords || '',
        data?.reason || '',
        data?.direct || '',
        overview || clean(place?.sourceMeta || ''),
        place?.sourceName || '',
        place?.publicData?.status || '',
        place?.publicData?.educationLevel || '',
        place?.publicData?.rawId || place?.id || '',
        place?.sourceUrl || '',
      ]);
    });
  });
  return rows;
}

function unitRows(units) {
  const rows = [[
    '학년',
    '교과',
    '단원',
    '단원 핵심',
    '세부주제/차시',
    '성취기준/학습목표',
    '핵심키워드',
    '검색키워드',
    '장소유형키워드',
    '자료출처',
  ]];
  units.forEach(({ grade, subject, unit }) => {
    rows.push([
      `${grade}학년`,
      subject,
      unitTitle(unit),
      unit?.focus || '',
      unitSubtopics(unit).map((s) => [s.lesson, s.name].filter(Boolean).join(' · ')).join(' / '),
      unitSubtopics(unit).map((s) => s.standard).filter(Boolean).join(' / '),
      uniq(unit?.terms || []).join(', '),
      uniq(unit?.searchTerms || []).join(', '),
      uniq(unit?.placeTypes || []).join(', '),
      unit?.source || '',
    ]);
  });
  return rows;
}

function criteriaRows() {
  return [
    ['구분', '선별 기준', '운영 판단'],
    ['대표후보', '성취기준 직접 근거가 있고, 지역 대표성/공식성/방문성이 높은 장소', '앱의 기본 상단 추천 및 고품질 사전 자료 제작 대상'],
    ['우선후보', '단원 핵심어와 장소 성격이 뚜렷하게 맞는 장소', '대표후보가 부족한 지역이나 단원에서 보완 추천'],
    ['보조후보', '연계 가능성은 있으나 장소 소개나 교육활동 근거가 더 필요한 장소', '검토 후 유지/삭제 결정'],
    ['검토필요', '검수필요, 전년도/종료 축제, 소개 부족, 좌표/운영 확인 필요', '기본 추천에서는 보류하거나 별도 표시'],
    ['삭제', '오락·상업·숙박·음식·골프·여행사·인형/테디베어 등 교육과정 직접 근거가 약한 장소', '전역 제외 또는 특정 단원 제외'],
    ['다음 작업', '검토결과 칸에 대표/보조/삭제를 표시', '표시 결과를 앱 큐레이션 데이터로 반영'],
  ];
}

function summaryRows(candidateRowsData, units) {
  const data = candidateRowsData.slice(1);
  const byUnit = new Map();
  data.forEach((row) => {
    const key = `${row[5]}|${row[6]}|${row[7]}`;
    if (!byUnit.has(key)) byUnit.set(key, { total: 0, representative: 0, priority: 0, review: 0 });
    const item = byUnit.get(key);
    item.total += 1;
    if (row[2] === '대표후보') item.representative += 1;
    if (row[2] === '우선후보') item.priority += 1;
    if (row[2] === '검토필요') item.review += 1;
  });
  const rows = [['학년', '교과', '단원', '후보수', '대표후보수', '우선후보수', '검토필요수', '메모']];
  units.forEach(({ grade, subject, unit }) => {
    const key = `${grade}학년|${subject}|${unitTitle(unit)}`;
    const item = byUnit.get(key) || { total: 0, representative: 0, priority: 0, review: 0 };
    rows.push([
      `${grade}학년`,
      subject,
      unitTitle(unit),
      item.total,
      item.representative,
      item.priority,
      item.review,
      item.total ? '후보 검토 가능' : '후보 보강 필요',
    ]);
  });
  return rows;
}

function main() {
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const app = readApp();
  const units = allUnits(app);
  const candidates = candidateRows(app, units);
  writeCsv('01_단원별_장소선별후보.csv', candidates);
  writeCsv('02_단원요약.csv', unitRows(units));
  writeCsv('03_단원별_후보수요약.csv', summaryRows(candidates, units));
  writeCsv('04_선별기준.csv', criteriaRows());
  console.log(`단원 수: ${units.length}`);
  console.log(`단원-장소 후보 조합: ${candidates.length - 1}`);
  console.log(`완료: ${OUT_DIR}`);
}

main();
