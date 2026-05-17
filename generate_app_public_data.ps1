param(
    [string]$XlsxPath = "경북_추가공공데이터_수집전처리_2026-05-12.xlsx",
    [string]$OutputDir = "data",
    [string]$CollectedAt = "2026-05-12",
    [string]$RawFolder = "추가공공데이터_수집_전처리_20260512_110300/raw"
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$XlsxFullPath = Join-Path $Root $XlsxPath
$OutFullPath = Join-Path $Root $OutputDir

New-Item -ItemType Directory -Force -Path $OutFullPath | Out-Null

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Write-Utf8File {
    param([string]$Path, [string]$Text)
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Text, $enc)
}

function Read-ZipEntryText {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$EntryName)
    $entry = $Zip.GetEntry($EntryName)
    if ($null -eq $entry) { throw "XLSX entry not found: $EntryName" }
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    try {
        return $reader.ReadToEnd()
    } finally {
        $reader.Close()
    }
}

function Get-ColIndex {
    param([string]$CellRef)
    $letters = ([regex]::Match($CellRef, "^[A-Z]+")).Value
    $index = 0
    foreach ($ch in $letters.ToCharArray()) {
        $index = ($index * 26) + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $index
}

function Get-CellText {
    param($Cell, [System.Xml.XmlNamespaceManager]$Ns)
    $textNode = $Cell.SelectSingleNode("x:is/x:t", $Ns)
    if ($textNode) { return [string]$textNode.InnerText }
    $valueNode = $Cell.SelectSingleNode("x:v", $Ns)
    if ($valueNode) { return [string]$valueNode.InnerText }
    return ""
}

function Read-XlsxSheet {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$SheetEntry)

    [xml]$doc = Read-ZipEntryText $Zip $SheetEntry
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

    $rows = $doc.SelectNodes("//x:sheetData/x:row", $ns)
    if ($rows.Count -eq 0) { return @() }

    $headerMap = @{}
    foreach ($cell in $rows[0].SelectNodes("x:c", $ns)) {
        $col = Get-ColIndex $cell.GetAttribute("r")
        $header = Get-CellText $cell $ns
        if (-not [string]::IsNullOrWhiteSpace($header)) {
            $headerMap[$col] = $header
        }
    }

    $records = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -lt $rows.Count; $i++) {
        $obj = [ordered]@{}
        foreach ($h in ($headerMap.GetEnumerator() | Sort-Object Name)) {
            $obj[$h.Value] = ""
        }
        foreach ($cell in $rows[$i].SelectNodes("x:c", $ns)) {
            $col = Get-ColIndex $cell.GetAttribute("r")
            if ($headerMap.ContainsKey($col)) {
                $obj[$headerMap[$col]] = Get-CellText $cell $ns
            }
        }
        $records.Add([pscustomobject]$obj) | Out-Null
    }
    return $records.ToArray()
}

function To-IntOrNull {
    param([string]$Value)
    $n = 0
    if ([int]::TryParse($Value, [ref]$n)) { return $n }
    return $null
}

function To-DoubleOrNull {
    param([string]$Value)
    $d = 0.0
    if ([double]::TryParse($Value, [ref]$d)) { return $d }
    return $null
}

function Split-Tags {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return @() }
    return @($Value -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Clean-OneLine {
    param([string]$Value)
    if ($null -eq $Value) { return "" }
    return (($Value -replace "\s+", " ").Trim())
}

function Source-Type {
    param([string]$SourceId)
    switch -Wildcard ($SourceId) {
        "S01*" { return "ecoTour" }
        "S02*" { return "cityPark" }
        "S03*" { return "nationalParkTrail" }
        "S04*" { return "forestTrail" }
        "S05*" { return "ruralVillage" }
        "S06*" { return "festival" }
        "S07*" { return "touristArea" }
        "S08*" { return "museumArt" }
        default { return "publicData" }
    }
}

function Visit-Note {
    param([string]$SourceType, [string]$Status)
    switch ($SourceType) {
        "nationalParkTrail" { return "탐방 전 통제 여부, 난이도, 기상 상황을 확인해야 합니다." }
        "forestTrail" { return "주소 필드가 없어 산명 기반 1차 후보입니다. 지도 표시 전 좌표 검수가 필요합니다." }
        "festival" {
            if ($Status -like "*전년도*") { return "전년도 또는 종료 자료일 수 있어 올해 개최 여부를 확인해야 합니다." }
            return "축제 일정과 운영 장소는 방문 전 공식 안내를 확인해야 합니다."
        }
        "ruralVillage" { return "체험 프로그램 운영 여부와 예약 가능일을 사전에 확인해야 합니다." }
        "cityPark" { return "생활권·공공시설 탐구용 후보입니다. 단원 직접 연계성은 현장 검토가 필요합니다." }
        "museumArt" { return "휴관일, 운영시간, 관람료를 방문 전 확인해야 합니다." }
        "ecoTour" { return "생태·자연 탐구 후보입니다. 위치 좌표는 기존 관광 API와 보강 검토가 필요합니다." }
        default { return "방문 전 운영 여부와 교육활동 가능 여부를 확인해야 합니다." }
    }
}

function Education-Level {
    param([string]$Status, [int]$Score)
    if ($Status -like "검수필요*") { return "검수필요" }
    if ($Status -like "보조자료*") { return "보조연계" }
    if ($Score -ge 80) { return "직접연계" }
    return "보조연계"
}

function Json-Array {
    param([object[]]$Rows)
    if ($Rows.Count -eq 0) { return "[]" }
    return ($Rows | ConvertTo-Json -Depth 10)
}

$zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $XlsxFullPath))
try {
    $sourceRows = Read-XlsxSheet $zip "xl/worksheets/sheet1.xml"
    $masterRows = Read-XlsxSheet $zip "xl/worksheets/sheet3.xml"
} finally {
    $zip.Dispose()
}

$sourceFileMap = @{
    "S01_ECO_TOUR" = "S01_ECO_TOUR.raw.json"
    "S02_CITY_PARK" = "S02_CITY_PARK.raw.csv"
    "S03_NATIONAL_PARK_TRAIL" = "S03_NATIONAL_PARK_TRAIL.raw.zip"
    "S04_FOREST_TRAIL" = "S04_FOREST_TRAIL.raw.csv"
    "S05_RURAL_VILLAGE" = "S05_RURAL_VILLAGE.raw.csv"
    "S06_FESTIVAL" = "S06_FESTIVAL.raw.csv"
    "S07_TOURIST_AREA" = "S07_TOURIST_AREA.raw.csv"
    "S08_MUSEUM_ART" = "S08_MUSEUM_ART.raw.csv"
}

$sources = @($sourceRows | ForEach-Object {
    $sourceId = $_.원천ID
    $rawLeaf = if ($sourceFileMap.ContainsKey($sourceId)) { $sourceFileMap[$sourceId] } else { Split-Path -Leaf $_.원자료파일 }
    [pscustomobject][ordered]@{
        id = $sourceId
        name = $_.데이터명
        provider = $_.제공기관
        officialUrl = $_.공식확인URL
        endpoint = $_.요청주소_키제외
        accessType = $_.접근방식
        approval = $_.승인상태
        status = $_.수집상태
        rawCount = To-IntOrNull $_.원자료건수
        gyeongbukCount = To-IntOrNull $_.경북필터건수
        rawFile = "$RawFolder/$rawLeaf"
        notes = $_.비고
        collectedAt = $CollectedAt
    }
})

$sourceCounters = @{}
$places = @($masterRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.장소명) } | ForEach-Object {
    $sourceId = $_.원천ID
    if (-not $sourceCounters.ContainsKey($sourceId)) { $sourceCounters[$sourceId] = 0 }
    $sourceCounters[$sourceId] = [int]$sourceCounters[$sourceId] + 1
    $seq = "{0:D4}" -f $sourceCounters[$sourceId]
    $sourceType = Source-Type $sourceId
    $score = To-IntOrNull $_.점수
    if ($null -eq $score) { $score = 0 }
    $status = if ([string]::IsNullOrWhiteSpace($_.추천상태)) { "추천후보" } else { $_.추천상태 }
    $overview = Clean-OneLine $_.설명요약
    $evidence = Clean-OneLine $_.직접연계근거
    [pscustomobject][ordered]@{
        id = "${sourceId}_${seq}"
        sourceId = $sourceId
        sourceType = $sourceType
        sourceName = $_.데이터명
        sourceUrl = $_.공식출처
        collectedAt = $CollectedAt
        title = $_.장소명
        addr1 = $_.주소
        sigungu = $_.시군
        lat = To-DoubleOrNull $_.위도
        lng = To-DoubleOrNull $_.경도
        cat = $_.추천카테고리
        sub = $_.세부유형
        status = $status
        educationLevel = Education-Level $status $score
        scoreBase = $score
        curriculumTags = Split-Tags $_.교육과정_연계후보
        evidence = $evidence
        overview = $overview
        visitNote = Visit-Note $sourceType $status
        rawId = $_.원자료ID
        referenceDate = $_.기준일자
        dedupeKey = $_.중복키
    }
})

$placesByCategory = $places | Group-Object cat | Sort-Object Name | ForEach-Object {
    [pscustomobject][ordered]@{ category = $_.Name; count = $_.Count }
}
$placesBySource = $places | Group-Object sourceId | Sort-Object Name | ForEach-Object {
    [pscustomobject][ordered]@{ sourceId = $_.Name; count = $_.Count }
}
$placesByStatus = $places | Group-Object status | Sort-Object Name | ForEach-Object {
    [pscustomobject][ordered]@{ status = $_.Name; count = $_.Count }
}

$sourceJs = @"
// Generated from $XlsxPath on $CollectedAt.
// Do not edit manually; rerun generate_app_public_data.ps1 after updating the preprocessing workbook.
(function(global){
  const PUBLIC_DATA_SOURCES = $(Json-Array $sources);
  const PUBLIC_DATA_SOURCE_SUMMARY = {
    collectedAt: "$CollectedAt",
    sourceCount: $($sources.Count),
    rawFolder: "$RawFolder"
  };
  global.PUBLIC_DATA_SOURCES = PUBLIC_DATA_SOURCES;
  global.PUBLIC_DATA_SOURCE_SUMMARY = PUBLIC_DATA_SOURCE_SUMMARY;
  if (typeof module !== 'undefined') {
    module.exports = { PUBLIC_DATA_SOURCES, PUBLIC_DATA_SOURCE_SUMMARY };
  }
})(typeof window !== 'undefined' ? window : globalThis);
"@

$placeJs = @"
// Generated from $XlsxPath on $CollectedAt.
// Do not edit manually; rerun generate_app_public_data.ps1 after updating the preprocessing workbook.
(function(global){
  const ADDITIONAL_PUBLIC_PLACES = $(Json-Array $places);
  const ADDITIONAL_PUBLIC_PLACE_SUMMARY = {
    collectedAt: "$CollectedAt",
    totalCount: $($places.Count),
    byCategory: $(Json-Array @($placesByCategory)),
    bySource: $(Json-Array @($placesBySource)),
    byStatus: $(Json-Array @($placesByStatus))
  };
  global.ADDITIONAL_PUBLIC_PLACES = ADDITIONAL_PUBLIC_PLACES;
  global.ADDITIONAL_PUBLIC_PLACE_SUMMARY = ADDITIONAL_PUBLIC_PLACE_SUMMARY;
  if (typeof module !== 'undefined') {
    module.exports = { ADDITIONAL_PUBLIC_PLACES, ADDITIONAL_PUBLIC_PLACE_SUMMARY };
  }
})(typeof window !== 'undefined' ? window : globalThis);
"@

Write-Utf8File (Join-Path $OutFullPath "public_data_sources.js") $sourceJs
Write-Utf8File (Join-Path $OutFullPath "gyeongbuk_public_places.js") $placeJs

$readme = @"
# 앱용 추가 공공데이터

생성일: $CollectedAt

## 파일

- public_data_sources.js: 원천별 공식 출처, 수집 상태, raw 파일 경로
- gyeongbuk_public_places.js: 앱에서 병합할 경북 후보 장소 데이터

## 생성 기준

- 원천 엑셀: $XlsxPath
- 장소 시트: 02_최종마스터_중복제거
- 출처 시트: 00_수집요약
- raw 폴더: $RawFolder

## 건수

- 앱용 장소 후보: $($places.Count)건
- 공공데이터 원천: $($sources.Count)개

## 카테고리별 건수

$($placesByCategory | ForEach-Object { "- $($_.category): $($_.count)건" } | Out-String)

## 상태별 건수

$($placesByStatus | ForEach-Object { "- $($_.status): $($_.count)건" } | Out-String)

## 주의

- `검수필요` 후보는 앱에서 기본 추천으로 과하게 노출하지 말고, 하단 또는 별도 토글로 표시한다.
- 축제 데이터는 날짜 상태를 확인해 `전년도 참고` 또는 `종료` 배지를 표시한다.
- 등산로 데이터는 코스/산명 자료이므로 일반 장소 카드와 다른 UI가 필요하다.
"@
Write-Utf8File (Join-Path $OutFullPath "README.md") $readme

Write-Output "DONE"
Write-Output "SOURCES=$($sources.Count)"
Write-Output "PLACES=$($places.Count)"
Write-Output "OUT=$OutFullPath"
