const fs = require('fs');
const path = require('path');
const vm = require('vm');

const HTML_FILE = '경북체험학습추천 (2).html';
const OUT_FILE = '대구경북_장소데이터_전체점검용.xlsx';

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
      CUR, GBK, DAEGU, SERVICE_AREAS, TOUR_AREA_CODES, SIGUNGU, DAEGU_SIGUNGU, TK, TOUR, APP, EDU_CONTENT_TYPES, BLOCKED_CONTENT_TYPES,
      CONTENT_TYPE_NAMES, contentTypeName, isAllowedContentType, isNonLearningPlace,
      isUsableTourPlace, clf, clsub, buildPlace, scorePlace, curriculumData,
      curriculumRelevance, PLACE_DISCOVERY_GROUPS, placeDiscoverySeeds, allPlaceDiscoverySeeds,
      setState(g, s) { A.gr = g; A.sj = s; }
    };
  `, context);
  return context.__APP_EXPORT__;
}

function uniq(arr) {
  return Array.from(new Set((arr || []).filter(Boolean)));
}

function clean(v = '') {
  return String(v).replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim();
}

function join(arr) {
  return uniq(arr).join(', ');
}

function allUnits(app) {
  const units = [];
  Object.entries(app.CUR).forEach(([grade, subjects]) => {
    (subjects.사회 || []).forEach((unit) => units.push({ grade, subject: '사회', unit }));
  });
  return units;
}

function unitSeeds(unit) {
  return uniq([...(unit.searchTerms || []), ...(unit.terms || []).slice(0, 8)])
    .map((v) => String(v).trim())
    .filter((v) => v.length >= 2);
}

function apiUrl(app, endpoint, params) {
  const query = Object.entries(params)
    .filter(([, v]) => v !== undefined && v !== '')
    .map(([k, v]) => `${k}=${encodeURIComponent(v)}`)
    .join('&');
  return `${app.TOUR}/${endpoint}?serviceKey=${app.TK}&${query}&MobileOS=ETC&MobileApp=${app.APP}&_type=json`;
}

async function fetchItems(app, endpoint, params, sourceLabel) {
  const url = apiUrl(app, endpoint, params);
  try {
    const res = await fetch(url);
    if (!res.ok) return [];
    const data = await res.json();
    const item = data?.response?.body?.items?.item;
    const list = item ? (Array.isArray(item) ? item : [item]) : [];
    return list.map((it) => ({ ...it, __source: sourceLabel }));
  } catch {
    return [];
  }
}

async function runLimited(tasks, limit = 10) {
  const results = [];
  let index = 0;
  async function worker() {
    while (index < tasks.length) {
      const current = index++;
      results[current] = await tasks[current]();
      if ((current + 1) % 50 === 0) console.log(`수집 진행: ${current + 1}/${tasks.length}`);
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, tasks.length) }, worker));
  return results.flat();
}

function mergePlace(map, item) {
  if (!item?.contentid) return;
  const id = String(item.contentid);
  if (!map.has(id)) {
    map.set(id, { ...item, __sources: new Set(), __keywords: new Set() });
  }
  const row = map.get(id);
  if (item.__source) row.__sources.add(item.__source);
  if (item.__keyword) row.__keywords.add(item.__keyword);
}

function excludeReason(app, item) {
  if (!item?.mapx || !item?.mapy) return '좌표 없음';
  const type = String(item.contenttypeid || '');
  const typeName = app.contentTypeName(type);
  if ((app.BLOCKED_CONTENT_TYPES || []).includes(type)) return `학습장소 제외: ${typeName} 유형`;
  if (!app.isAllowedContentType(type)) return `허용하지 않는 관광 API 유형: ${typeName}`;
  if (app.isNonLearningPlace(item)) {
    if (type === '38') return '쇼핑 데이터 중 전통시장·공예·문화거리 성격이 약함';
    return '음식점·숙박·상점 성격 명칭으로 제외';
  }
  return '';
}

async function collectPlaces(app) {
  const units = allUnits(app);
  const allSeeds = uniq([
    ...units.flatMap(({ unit }) => unitSeeds(unit)),
    ...(app.allPlaceDiscoverySeeds ? app.allPlaceDiscoverySeeds() : []),
  ]);
  const placeMap = new Map();
  const serviceAreas = app.SERVICE_AREAS || [{ code: '35', name: '경북', center: app.GBK }];
  function sigunguMapFor(area) {
    if (String(area.code) === '4') return app.DAEGU_SIGUNGU || {};
    return app.SIGUNGU || {};
  }

  const areaTasks = [];
  serviceAreas.forEach((area) => {
    Object.entries(sigunguMapFor(area)).forEach(([sigungu, code]) => {
      app.EDU_CONTENT_TYPES.forEach((type) => {
        areaTasks.push(async () => fetchItems(app, 'areaBasedList2', {
          areaCode: area.code,
          sigunguCode: code,
          contentTypeId: type,
          arrange: 'A',
          numOfRows: 1000,
          pageNo: 1,
        }, `${area.name}전체:${sigungu}:${app.contentTypeName(type)}`));
      });
    });
  });

  console.log(`대구·경북 전체 허용 장소 수집 요청 ${areaTasks.length}건`);
  (await runLimited(areaTasks, 8)).forEach((item) => mergePlace(placeMap, item));

  const allowedKeywordTasks = [];
  allSeeds.forEach((seed) => {
    serviceAreas.forEach((area) => {
      app.EDU_CONTENT_TYPES.forEach((type) => {
        allowedKeywordTasks.push(async () => {
          const items = await fetchItems(app, 'searchKeyword2', {
            keyword: seed,
            areaCode: area.code,
            contentTypeId: type,
            arrange: 'A',
            numOfRows: 80,
            pageNo: 1,
          }, `보강검색:${area.name}:${seed}:${app.contentTypeName(type)}`);
          return items.map((it) => ({ ...it, __keyword: seed }));
        });
      });
    });
  });

  console.log(`허용 장소 보강 검색 요청 ${allowedKeywordTasks.length}건`);
  (await runLimited(allowedKeywordTasks, 8)).forEach((item) => mergePlace(placeMap, item));

  const blockedTypes = uniq([...(app.BLOCKED_CONTENT_TYPES || []), '39']);
  const blockedTasks = [];
  allSeeds.forEach((seed) => {
    serviceAreas.forEach((area) => {
      blockedTypes.forEach((type) => {
        blockedTasks.push(async () => {
          const items = await fetchItems(app, 'searchKeyword2', {
            keyword: seed,
            areaCode: area.code,
            contentTypeId: type,
            arrange: 'A',
            numOfRows: 50,
            pageNo: 1,
          }, `제외검사용:${area.name}:${seed}:${app.contentTypeName(type)}`);
          return items.map((it) => ({ ...it, __keyword: seed }));
        });
      });
    });
  });

  console.log(`음식점/숙박 오염 검사 요청 ${blockedTasks.length}건`);
  (await runLimited(blockedTasks, 8)).forEach((item) => mergePlace(placeMap, item));

  return { units, places: Array.from(placeMap.values()) };
}

function guideRows() {
  return [
    ['항목', '내용'],
    ['파일 목적', '앱이 대구·경북 관광자원 API에서 가져올 수 있는 장소 후보를 점검하기 위한 원자료 엑셀'],
    ['추천통과_단원별', '현재 앱 규칙으로 교육과정 점수를 통과한 단원-장소 조합'],
    ['장소마스터', '대구·경북 전체 허용 관광 API 유형과 오염 검사용 음식점/숙박 후보를 합친 장소 목록'],
    ['보강검색', '단원 키워드 외에 천문대, 천문과학관, 생태공원, 지질공원, 대표 경제거점처럼 장소군별 누락 방지 키워드로 추가 수집'],
    ['제외장소', '음식점, 숙박, 좌표 없음, 학습장소 성격 부족 등으로 앱 추천에서 제외되는 장소'],
    ['수정 방법', '장소마스터에서 남길 대표 장소를 표시해 주면 코드의 허용/제외 사전으로 반영 가능'],
    ['주의', '앱은 실시간 API를 사용하므로 학교 위치/검색 시점에 따라 후보는 조금 달라질 수 있음'],
  ];
}

function centerForPlace(app, raw) {
  const area = (app.SERVICE_AREAS || []).find((a) => String(a.code) === String(raw.areacode || raw.areaCode || ''));
  return area?.center || app.GBK;
}

function masterRows(app, places) {
  const rows = [['검토결과(유지/삭제)', '검토메모', '사용가능', '제외사유', 'contentid', '장소명', '주소', '관광API유형', 'contenttypeid', 'cat1', 'cat2', 'cat3', '시군코드', '좌표X', '좌표Y', '이미지', '앱대분류', '앱소분류', '수집경로', '검색어']];
  places
    .sort((a, b) => clean(a.title).localeCompare(clean(b.title), 'ko'))
    .forEach((p) => {
      const reason = excludeReason(app, p);
      const usable = !reason && app.isUsableTourPlace(p);
      rows.push([
        '',
        '',
        usable ? '사용 가능' : '제외',
        reason,
        p.contentid || '',
        clean(p.title),
        clean(p.addr1),
        app.contentTypeName(p.contenttypeid),
        p.contenttypeid || '',
        p.cat1 || '',
        p.cat2 || '',
        p.cat3 || '',
        p.sigungucode || '',
        p.mapx || '',
        p.mapy || '',
        p.firstimage || p.firstimage2 ? '있음' : '없음',
        usable ? app.clf({ title: p.title || '', addr1: p.addr1 || '' }) : '',
        usable ? app.clsub({ title: p.title || '', addr1: p.addr1 || '' }) : '',
        join(Array.from(p.__sources || [])),
        join(Array.from(p.__keywords || [])),
      ]);
    });
  return rows;
}

function recommendationRows(app, units, places) {
  const rows = [['검토결과(유지/삭제)', '검토메모', '학년', '교과', '선택 단원', 'contentid', '장소명', '주소', '관광API유형', '앱대분류', '앱소분류', '기준점거리(km)', '추천점수', '관련 주제', '관련 차시', '성취기준', '추천 이유', '직접 연계 근거', '출처 기준']];
  const candidatePlaces = places.filter((p) => {
    if (!p?.mapx || !p?.mapy) return false;
    const type = String(p.contenttypeid || '');
    return app.isAllowedContentType(type) && !(app.BLOCKED_CONTENT_TYPES || []).includes(type);
  });
  const recs = [];
  units.forEach(({ grade, subject, unit }) => {
    app.setState(grade, subject);
    candidatePlaces.forEach((raw) => {
      if (!app.isUsableTourPlace(raw, unit)) return;
      const place = app.buildPlace(raw, centerForPlace(app, raw), unit, null);
      if (!place.score) return;
      const data = app.curriculumData(place, unit, '');
      recs.push([
        '',
        '',
        `${grade}학년`,
        subject,
        unit.unit || unit.name,
        raw.contentid || '',
        clean(raw.title),
        clean(raw.addr1),
        app.contentTypeName(raw.contenttypeid),
        place.cat || '',
        place.sub || '',
        place.dist,
        place.score,
        data.topic,
        data.lesson,
        data.standard,
        data.reason,
        data.direct,
        data.source,
      ]);
    });
  });
  recs.sort((a, b) => String(a[2]).localeCompare(String(b[2]), 'ko') || String(a[4]).localeCompare(String(b[4]), 'ko') || Number(b[12]) - Number(a[12]));
  return rows.concat(recs);
}

function excludedRows(app, places) {
  const rows = [['제외사유', 'contentid', '장소명', '주소', '관광API유형', 'contenttypeid', '수집경로', '검색어']];
  places
    .map((p) => ({ p, reason: excludeReason(app, p) || (app.isUsableTourPlace(p) ? '' : '교육여행 장소로 부적합') }))
    .filter((x) => x.reason)
    .sort((a, b) => a.reason.localeCompare(b.reason, 'ko') || clean(a.p.title).localeCompare(clean(b.p.title), 'ko'))
    .forEach(({ p, reason }) => {
      rows.push([
        reason,
        p.contentid || '',
        clean(p.title),
        clean(p.addr1),
        app.contentTypeName(p.contenttypeid),
        p.contenttypeid || '',
        join(Array.from(p.__sources || [])),
        join(Array.from(p.__keywords || [])),
      ]);
    });
  return rows;
}

function unitRows(app, units) {
  const rows = [['학년', '교과', '선택 단원', '앱 검색어', '장소군 보강 검색어']];
  units.forEach(({ grade, subject, unit }) => rows.push([
    `${grade}학년`,
    subject,
    unit.unit || unit.name,
    join(unitSeeds(unit)),
    join(app.placeDiscoverySeeds ? app.placeDiscoverySeeds(unit) : []),
  ]));
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
  const cols = Array.from({ length: maxCols }, (_, i) => `<col min="${i + 1}" max="${i + 1}" width="${i < 3 ? 18 : 30}" customWidth="1"/>`).join('');
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
  const { units, places } = await collectPlaces(app);
  const recRows = recommendationRows(app, units, places);
  const master = masterRows(app, places);
  const excluded = excludedRows(app, places);
  const sheets = [
    { name: '00_안내', rows: guideRows() },
    { name: '01_추천통과_단원별', rows: recRows },
    { name: '02_장소마스터', rows: master },
    { name: '03_제외장소', rows: excluded },
    { name: '04_단원별검색어', rows: unitRows(app, units) },
  ];
  makeXlsx(sheets, path.resolve(OUT_FILE));
  console.log(`장소마스터 ${master.length - 1}건`);
  console.log(`추천통과 조합 ${recRows.length - 1}건`);
  console.log(`제외장소 ${excluded.length - 1}건`);
  console.log(`완료: ${OUT_FILE}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
