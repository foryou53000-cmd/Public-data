param(
    [string]$ServiceKey = $env:PUBLIC_DATA_SERVICE_KEY,
    [string]$RunDate = "2026-05-12"
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($ServiceKey)) {
    $ServiceKey = "1ed167b364dd8cc6ef53cea742bb851f80fa46ad7ec3db01379aeffad968e6ff"
}

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Stamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$OutDir = Join-Path $Root "추가공공데이터_수집_전처리_$Stamp"
$RawDir = Join-Path $OutDir "raw"
$WorkDir = Join-Path $OutDir "work"
$XlsxPath = Join-Path $Root "경북_추가공공데이터_수집전처리_$RunDate.xlsx"
$SummaryPath = Join-Path $Root "공공데이터_수집_전처리_요약_$RunDate.md"

New-Item -ItemType Directory -Force -Path $OutDir, $RawDir, $WorkDir | Out-Null

Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName Microsoft.VisualBasic

$CollectedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
$SourceRows = New-Object System.Collections.Generic.List[object]
$ProcessRows = New-Object System.Collections.Generic.List[object]
$MasterRows = New-Object System.Collections.Generic.List[object]
$ExcludedRows = New-Object System.Collections.Generic.List[object]
$DuplicateRows = New-Object System.Collections.Generic.List[object]

function Write-Utf8File {
    param([string]$Path, [string]$Text)
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Text, $enc)
}

function Invoke-CurlText {
    param([string]$Url)
    $curlArgs = @()
    if ($Url -like "https://*") {
        $curlArgs += "--ssl-no-revoke"
    }
    $curlArgs += @("-sS", "-L", "--url", $Url)
    $lastOutput = ""
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $result = & curl.exe @curlArgs 2>&1
        if ($LASTEXITCODE -eq 0) {
            return ($result -join "`n")
        }
        $lastOutput = ($result -join " ")
        Start-Sleep -Seconds (2 * $attempt)
    }
    throw "curl failed with exit code $LASTEXITCODE for $Url :: $lastOutput"
}

function Invoke-CurlDownload {
    param([string]$Url, [string]$Path)
    $curlArgs = @()
    if ($Url -like "https://*") {
        $curlArgs += "--ssl-no-revoke"
    }
    $curlArgs += @("-sS", "-L", "-o", $Path, "--url", $Url)
    $lastOutput = ""
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $result = & curl.exe @curlArgs 2>&1
        if ($LASTEXITCODE -eq 0) { return }
        $lastOutput = ($result -join " ")
        Start-Sleep -Seconds (2 * $attempt)
    }
    throw "curl download failed with exit code $LASTEXITCODE for $Url :: $lastOutput"
}

function Add-Process {
    param([string]$Step, [string]$InputData, [string]$Rule, [string]$OutputData, [string]$Reason)
    $ProcessRows.Add([pscustomobject]@{
        단계 = $Step
        입력자료 = $InputData
        처리규칙 = $Rule
        산출물 = $OutputData
        교육적_검수_의미 = $Reason
    }) | Out-Null
}

function Add-SourceStatus {
    param(
        [string]$SourceId,
        [string]$Name,
        [string]$Provider,
        [string]$OfficialUrl,
        [string]$Endpoint,
        [string]$AccessType,
        [string]$Approval,
        [string]$Status,
        [int]$RawCount,
        [int]$GyeongbukCount,
        [string]$RawFile,
        [string]$Notes
    )
    $SourceRows.Add([pscustomobject]@{
        수집일시 = $CollectedAt
        원천ID = $SourceId
        데이터명 = $Name
        제공기관 = $Provider
        공식확인URL = $OfficialUrl
        요청주소_키제외 = $Endpoint
        접근방식 = $AccessType
        승인상태 = $Approval
        수집상태 = $Status
        원자료건수 = $RawCount
        경북필터건수 = $GyeongbukCount
        원자료파일 = $RawFile
        비고 = $Notes
    }) | Out-Null
}

function Get-Prop {
    param($Obj, [string[]]$Names)
    if ($null -eq $Obj) { return "" }
    $props = @($Obj.PSObject.Properties)
    foreach ($name in $Names) {
        $p = $props | Where-Object { $_.Name -ieq $name } | Select-Object -First 1
        if ($p -and $null -ne $p.Value -and -not [string]::IsNullOrWhiteSpace([string]$p.Value)) {
            return ([string]$p.Value).Trim()
        }
    }
    return ""
}

function Clean-Text {
    param([string]$Value)
    if ($null -eq $Value) { return "" }
    return (($Value -replace "<[^>]+>", " ") -replace "\s+", " ").Trim()
}

function Normalize-Key {
    param([string]$Name)
    if ($null -eq $Name) { return "" }
    return (($Name.ToLowerInvariant() -replace "\s+", "") -replace "[\(\)\[\]\{\}·ㆍ\-_]", "")
}

function Get-Sigungu {
    param([string]$Address, [string]$Fallback)
    if ($Address -match "경상북도\s+([^\s]+)") {
        return $Matches[1]
    }
    if ($Fallback -match "경상북도\s+([^\s]+)") {
        return $Matches[1]
    }
    if ($Fallback -match "^(.+시|.+군)$") {
        return $Fallback
    }
    return ""
}

function Is-GyeongbukRecord {
    param($Row, [string[]]$AddressFields, [string[]]$ProvinceFields)
    foreach ($field in $AddressFields) {
        if ((Get-Prop $Row @($field)) -like "*경상북도*") { return $true }
    }
    foreach ($field in $ProvinceFields) {
        if ((Get-Prop $Row @($field)) -eq "경상북도") { return $true }
    }
    return $false
}

function Has-Bad-Place-Keyword {
    param([string]$Text)
    $bad = "맛집|음식점|식당|카페|커피|베이커리|호텔|모텔|펜션|리조트|숙박|여관|민박|게스트하우스|상가|쇼핑|마트|백화점|주점|술집"
    return ($Text -match $bad)
}

function Get-Curriculum-Tags {
    param([string]$Category, [string]$Text)
    $tags = New-Object System.Collections.Generic.List[string]
    switch ($Category) {
        "자연·생태" {
            $tags.Add("국토의 자연환경") | Out-Null
            $tags.Add("지역의 자연환경") | Out-Null
            $tags.Add("환경 보전과 지속가능한 생활") | Out-Null
        }
        "산지·탐방로" {
            $tags.Add("우리 국토의 지형과 산지") | Out-Null
            $tags.Add("지도 읽기와 이동 경로") | Out-Null
            $tags.Add("안전한 현장체험 계획") | Out-Null
        }
        "공원·생활환경" {
            $tags.Add("생활권과 지역 환경") | Out-Null
            $tags.Add("지도 활용과 지역 조사") | Out-Null
            $tags.Add("시민 참여와 공공시설") | Out-Null
        }
        "농어촌·생산체험" {
            $tags.Add("촌락과 도시의 생활 모습") | Out-Null
            $tags.Add("생산과 소비") | Out-Null
            $tags.Add("지역문화와 교류") | Out-Null
        }
        "지역축제" {
            $tags.Add("지역 문화와 지역 알리기") | Out-Null
            $tags.Add("계절과 생활문화") | Out-Null
        }
        "관광명소" {
            $tags.Add("지역의 자연·문화 자원") | Out-Null
            $tags.Add("가족 답사 계획") | Out-Null
        }
        "박물관·전시" {
            $tags.Add("역사·문화 자료 조사") | Out-Null
            $tags.Add("전시 자료를 통한 탐구") | Out-Null
        }
        default {
            $tags.Add("지역 자원 탐구") | Out-Null
        }
    }
    if ($Text -match "독립|의병|3\.1|삼일|항일|전승|전쟁|호국|현충") {
        $tags.Add("5학년 사회 6단원: 일제강점기·독립운동·민주주의와 국가수호") | Out-Null
    }
    if ($Text -match "고려|조선|신라|삼국|유적|유물|사적|문화유산|국보|보물") {
        $tags.Add("역사 단원: 시대별 문화유산과 장소 탐구") | Out-Null
    }
    return ($tags | Select-Object -Unique) -join ", "
}

function Add-Candidate {
    param(
        [string]$SourceId,
        [string]$SourceName,
        [string]$Provider,
        [string]$Title,
        [string]$Address,
        [string]$Sigungu,
        [string]$Latitude,
        [string]$Longitude,
        [string]$Category,
        [string]$SubCategory,
        [string]$SourceUrl,
        [string]$RawId,
        [string]$Evidence,
        [string]$DetailText,
        [string]$ReferenceDate,
        [string]$KeepStatus,
        [string]$ExcludeReason,
        [int]$Score
    )
    $titleClean = Clean-Text $Title
    $addressClean = Clean-Text $Address
    $detailClean = Clean-Text $DetailText
    $evidenceClean = Clean-Text $Evidence
    $allText = "$titleClean $addressClean $detailClean $evidenceClean"
    if ([string]::IsNullOrWhiteSpace($titleClean)) {
        $ExcludedRows.Add([pscustomobject]@{
            원천ID = $SourceId
            데이터명 = $SourceName
            장소명 = ""
            주소 = $addressClean
            제외사유 = "장소명 없음"
            원자료ID = $RawId
        }) | Out-Null
        return
    }
    if (Has-Bad-Place-Keyword $titleClean) {
        $ExcludedRows.Add([pscustomobject]@{
            원천ID = $SourceId
            데이터명 = $SourceName
            장소명 = $titleClean
            주소 = $addressClean
            제외사유 = "장소명 자체가 음식점·숙박·상업시설 계열"
            원자료ID = $RawId
        }) | Out-Null
        return
    }
    if ($KeepStatus -eq "제외") {
        $ExcludedRows.Add([pscustomobject]@{
            원천ID = $SourceId
            데이터명 = $SourceName
            장소명 = $titleClean
            주소 = $addressClean
            제외사유 = $ExcludeReason
            원자료ID = $RawId
        }) | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($Sigungu)) {
        $Sigungu = Get-Sigungu $addressClean ""
    }
    $MasterRows.Add([pscustomobject]@{
        원천ID = $SourceId
        데이터명 = $SourceName
        제공기관 = $Provider
        장소명 = $titleClean
        시군 = $Sigungu
        주소 = $addressClean
        위도 = $Latitude
        경도 = $Longitude
        추천카테고리 = $Category
        세부유형 = $SubCategory
        추천상태 = $KeepStatus
        점수 = $Score
        교육과정_연계후보 = Get-Curriculum-Tags $Category $allText
        직접연계근거 = $evidenceClean
        설명요약 = $detailClean
        원자료ID = $RawId
        기준일자 = $ReferenceDate
        공식출처 = $SourceUrl
        중복키 = "$(Normalize-Key $titleClean)|$Sigungu"
    }) | Out-Null
}

function Fetch-StandardApi {
    param(
        [string]$SourceId,
        [string]$Name,
        [string]$Provider,
        [string]$OfficialUrl,
        [string]$Endpoint
    )
    $items = New-Object System.Collections.Generic.List[object]
    $page = 1
    $pageSize = 1000
    $total = 0
    do {
        $url = "${Endpoint}?serviceKey=$([System.Web.HttpUtility]::UrlEncode($ServiceKey))&pageNo=$page&numOfRows=$pageSize&type=json"
        $text = Invoke-CurlText $url
        $obj = $text | ConvertFrom-Json
        $code = [string]$obj.response.header.resultCode
        if ($code -ne "00") {
            throw "$Name resultCode=$code resultMsg=$($obj.response.header.resultMsg)"
        }
        $pageItems = @($obj.response.body.items) | Where-Object { $_ }
        foreach ($item in $pageItems) { $items.Add($item) | Out-Null }
        $total = [int]$obj.response.body.totalCount
        $page++
    } while (($page - 1) * $pageSize -lt $total)

    $rawFile = Join-Path $RawDir "$SourceId.raw.csv"
    $items | Export-Csv -Path $rawFile -NoTypeInformation -Encoding UTF8
    return @{ Items = $items.ToArray(); Total = $total; RawFile = $rawFile }
}

function Import-CsvSmart {
    param([string]$Path)
    $rows = @()
    try {
        $rows = Import-Csv -Path $Path -Encoding UTF8
    } catch {
        $rows = @()
    }
    if ($rows.Count -eq 0 -or (($rows | Select-Object -First 1).PSObject.Properties.Name -join ",") -match "ï|Ã|ì") {
        try {
            $rows = Import-Csv -Path $Path -Encoding Default
        } catch {
            $rows = @()
        }
    }
    return @($rows)
}

function Get-BoundingGyeongbuk {
    param([string]$Lat, [string]$Lon)
    $latNum = 0.0
    $lonNum = 0.0
    if (-not [double]::TryParse($Lat, [ref]$latNum)) { return $false }
    if (-not [double]::TryParse($Lon, [ref]$lonNum)) { return $false }
    return ($latNum -ge 35.45 -and $latNum -le 37.65 -and $lonNum -ge 127.65 -and $lonNum -le 130.25)
}

function Process-CityParks {
    $sourceId = "S02_CITY_PARK"
    $name = "전국도시공원정보표준데이터"
    $provider = "국토교통부·지방자치단체"
    $official = "https://www.data.go.kr/data/15012890/standard.do"
    $endpoint = "https://api.data.go.kr/openapi/tn_pubr_public_cty_park_info_api"
    try {
        $result = Fetch-StandardApi $sourceId $name $provider $official $endpoint
        $gb = @($result.Items | Where-Object { Is-GyeongbukRecord $_ @("rdnmadr","lnmadr","RDNMADR","LNMADR") @() })
        foreach ($row in $gb) {
            $parkName = Get-Prop $row @("parkNm","PARK_NM")
            $parkSe = Get-Prop $row @("parkSe","PARK_SE")
            $addr = Get-Prop $row @("rdnmadr","RDNMADR","lnmadr","LNMADR")
            $facility = "$(Get-Prop $row @("mvmFclty","MVM_FCLTY")) $(Get-Prop $row @("amsmtFclty","AMSMT_FCLTY")) $(Get-Prop $row @("cnvnncFclty","CNVNNC_FCLTY")) $(Get-Prop $row @("cltrFclty","CLTR_FCLTY")) $(Get-Prop $row @("etcFclty","ETC_FCLTY"))"
            $text = "$parkName $parkSe $facility"
            $keep = "추천후보"
            $reason = ""
            $score = 45
            if ($parkSe -match "어린이공원|소공원" -and $text -notmatch "생태|자연|문화|역사|체험|과학|현충|전통|수변|공연|탐방|환경") {
                $keep = "제외"
                $reason = "어린이공원·소공원 중 교육 활동 근거가 약한 생활놀이공간"
            }
            if ($text -match "생태|자연|수변|문화|역사|체험|과학|현충|전통|공연") { $score += 20 }
            Add-Candidate $sourceId $name $provider $parkName $addr (Get-Sigungu $addr "") (Get-Prop $row @("latitude","LATITUDE")) (Get-Prop $row @("longitude","LONGITUDE")) "공원·생활환경" $parkSe $official (Get-Prop $row @("manageNo","MANAGE_NO")) "공원구분=$parkSe; 보유시설=$facility" "" (Get-Prop $row @("referenceDate","REFERENCE_DATE")) $keep $reason $score
        }
        Add-SourceStatus $sourceId $name $provider $official $endpoint "표준 OpenAPI JSON" "자동승인" "수집성공" $result.Total $gb.Count $result.RawFile "생활권·지도 활용 단원용. 어린이공원·소공원은 교육근거 약하면 제외 처리."
    } catch {
        Add-SourceStatus $sourceId $name $provider $official $endpoint "표준 OpenAPI JSON" "자동승인" "수집실패" 0 0 "" $_.Exception.Message
    }
}

function Process-RuralVillages {
    $sourceId = "S05_RURAL_VILLAGE"
    $name = "전국농어촌체험휴양마을표준데이터"
    $provider = "농림축산식품부·해양수산부"
    $official = "https://www.data.go.kr/data/15013113/standard.do"
    $endpoint = "https://api.data.go.kr/openapi/tn_pubr_public_frhl_exprn_vilage_api"
    try {
        $result = Fetch-StandardApi $sourceId $name $provider $official $endpoint
        $gb = @($result.Items | Where-Object { Is-GyeongbukRecord $_ @("rdnmadr","lnmadr") @("ctprvnNm") })
        foreach ($row in $gb) {
            $title = Get-Prop $row @("exprnVilageNm")
            $addr = Get-Prop $row @("rdnmadr","lnmadr")
            $program = Get-Prop $row @("exprnCn")
            $type = Get-Prop $row @("exprnSe")
            $score = 70
            if ("$program $type" -match "농사|수확|전통|문화|생태|자연|만들기|먹거리|어촌|농촌") { $score += 15 }
            Add-Candidate $sourceId $name $provider $title $addr (Get-Prop $row @("signguNm")) (Get-Prop $row @("latitude")) (Get-Prop $row @("longitude")) "농어촌·생산체험" $type $official $title "체험프로그램=$program; 보유시설=$(Get-Prop $row @("holdFclty"))" $program (Get-Prop $row @("referenceDate")) "추천후보" "" $score
        }
        Add-SourceStatus $sourceId $name $provider $official $endpoint "표준 OpenAPI JSON" "자동승인" "수집성공" $result.Total $gb.Count $result.RawFile "가족 체험 활동·생산/소비·촌락 단원과 직접 연결 가능."
    } catch {
        Add-SourceStatus $sourceId $name $provider $official $endpoint "표준 OpenAPI JSON" "자동승인" "수집실패" 0 0 "" $_.Exception.Message
    }
}

function Process-Festivals {
    $sourceId = "S06_FESTIVAL"
    $name = "전국문화축제표준데이터"
    $provider = "문화체육관광부·한국관광공사·지방자치단체"
    $official = "https://www.data.go.kr/data/15013104/standard.do"
    $endpoint = "https://api.data.go.kr/openapi/tn_pubr_public_cltur_fstvl_api"
    try {
        $result = Fetch-StandardApi $sourceId $name $provider $official $endpoint
        $gb = @($result.Items | Where-Object { Is-GyeongbukRecord $_ @("rdnmadr","lnmadr") @() })
        $today = [datetime]::ParseExact($RunDate, "yyyy-MM-dd", $null)
        foreach ($row in $gb) {
            $title = Get-Prop $row @("fstvlNm")
            $addr = Get-Prop $row @("rdnmadr","lnmadr")
            $content = Get-Prop $row @("fstvlCo")
            $startText = Get-Prop $row @("fstvlStartDate")
            $endText = Get-Prop $row @("fstvlEndDate")
            $status = "시기형 추천후보"
            $note = ""
            $score = 55
            $endDate = [datetime]::MinValue
            if ([datetime]::TryParse($endText, [ref]$endDate) -and $endDate -lt $today) {
                if ($endDate.Year -lt $today.Year) {
                    $status = "보조자료_전년도참고"
                    $note = "이미 종료된 축제. 올해 개최 여부 확인 전까지 전년도 계절성 참고자료로만 사용."
                    $score = 35
                } else {
                    $status = "보조자료_종료"
                    $note = "올해 이미 종료된 축제."
                    $score = 40
                }
            }
            if ("$title $content" -match "전통|문화|예술|유산|생태|자연|역사|농산|체험|과학") { $score += 15 }
            Add-Candidate $sourceId $name $provider $title $addr (Get-Sigungu $addr "") (Get-Prop $row @("latitude")) (Get-Prop $row @("longitude")) "지역축제" "$startText~$endText" $official $title "축제기간=$startText~$endText; 축제내용=$content; $note" $content (Get-Prop $row @("referenceDate")) $status "" $score
        }
        Add-SourceStatus $sourceId $name $provider $official $endpoint "표준 OpenAPI JSON" "자동승인" "수집성공" $result.Total $gb.Count $result.RawFile "축제는 상설 장소가 아니므로 날짜 상태를 별도 표기."
    } catch {
        Add-SourceStatus $sourceId $name $provider $official $endpoint "표준 OpenAPI JSON" "자동승인" "수집실패" 0 0 "" $_.Exception.Message
    }
}

function Process-TouristAreas {
    $sourceId = "S07_TOURIST_AREA"
    $name = "전국관광지정보표준데이터"
    $provider = "문화체육관광부·지방자치단체"
    $official = "https://www.data.go.kr/data/15021141/standard.do"
    $endpoint = "https://api.data.go.kr/openapi/tn_pubr_public_trrsrt_api"
    try {
        $result = Fetch-StandardApi $sourceId $name $provider $official $endpoint
        $gb = @($result.Items | Where-Object { Is-GyeongbukRecord $_ @("rdnmadr","lnmadr") @() })
        foreach ($row in $gb) {
            $title = Get-Prop $row @("trrsrtNm")
            $addr = Get-Prop $row @("rdnmadr","lnmadr")
            $intro = Get-Prop $row @("trrsrtIntrcn")
            $text = "$title $intro $(Get-Prop $row @("recrtClturFclty")) $(Get-Prop $row @("cnvnncFclty"))"
            $category = "관광명소"
            if ($text -match "역사|전승|문화|유적|기념|전쟁|상륙|유산") { $category = "박물관·전시" }
            if ($text -match "해수욕|해송|자연|산|호수|하천|생태|동굴|지질") { $category = "자연·생태" }
            $score = 55
            if ($text -match "역사|전승|문화|유적|자연|생태|지질|동굴|상륙") { $score += 20 }
            Add-Candidate $sourceId $name $provider $title $addr (Get-Sigungu $addr "") (Get-Prop $row @("latitude")) (Get-Prop $row @("longitude")) $category (Get-Prop $row @("trrsrtSe")) $official $title "관광지소개=$intro" $intro (Get-Prop $row @("referenceDate")) "추천후보" "" $score
        }
        Add-SourceStatus $sourceId $name $provider $official $endpoint "표준 OpenAPI JSON" "자동승인" "수집성공" $result.Total $gb.Count $result.RawFile "Tour API 누락 관광지 보완. 숙박·상가 정보는 보조 필드로만 사용."
    } catch {
        Add-SourceStatus $sourceId $name $provider $official $endpoint "표준 OpenAPI JSON" "자동승인" "수집실패" 0 0 "" $_.Exception.Message
    }
}

function Process-Museums {
    $sourceId = "S08_MUSEUM_ART"
    $name = "전국박물관미술관정보표준데이터"
    $provider = "문화체육관광부·지방자치단체"
    $official = "https://www.data.go.kr/data/15017323/standard.do"
    $endpoint = "https://api.data.go.kr/openapi/tn_pubr_public_museum_artgr_info_api"
    try {
        $result = Fetch-StandardApi $sourceId $name $provider $official $endpoint
        $gb = @($result.Items | Where-Object { Is-GyeongbukRecord $_ @("rdnmadr","lnmadr") @() })
        foreach ($row in $gb) {
            $title = Get-Prop $row @("fcltyNm")
            $addr = Get-Prop $row @("rdnmadr","lnmadr")
            $intro = Get-Prop $row @("fcltyIntrcn")
            $type = Get-Prop $row @("fcltyType")
            $score = 75
            if ("$title $intro" -match "역사|유물|문화|과학|농업|독립|의병|전쟁|고분|신라|고려|조선") { $score += 20 }
            Add-Candidate $sourceId $name $provider $title $addr (Get-Sigungu $addr "") (Get-Prop $row @("latitude")) (Get-Prop $row @("longitude")) "박물관·전시" $type $official $title "박물관미술관소개=$intro; 휴관=$(Get-Prop $row @("rstdeInfo"))" $intro (Get-Prop $row @("referenceDate")) "추천후보" "" $score
        }
        Add-SourceStatus $sourceId $name $provider $official $endpoint "표준 OpenAPI JSON" "자동승인" "수집성공" $result.Total $gb.Count $result.RawFile "역사·문화·과학 전시시설 후보 정밀화에 사용."
    } catch {
        Add-SourceStatus $sourceId $name $provider $official $endpoint "표준 OpenAPI JSON" "자동승인" "수집실패" 0 0 "" $_.Exception.Message
    }
}

function Process-EcoTour {
    $sourceId = "S01_ECO_TOUR"
    $name = "한국관광공사 생태 관광 정보_GW"
    $provider = "한국관광공사"
    $official = "https://www.data.go.kr/data/15101908/openapi.do"
    $endpoint = "https://apis.data.go.kr/B551011/GreenTourService1/areaBasedList1"
    try {
        $url = "${endpoint}?serviceKey=$([System.Web.HttpUtility]::UrlEncode($ServiceKey))&MobileOS=ETC&MobileApp=PublicDataEdu&_type=json&areaCode=35&numOfRows=100&pageNo=1&arrange=A"
        $text = Invoke-CurlText $url
        $rawFile = Join-Path $RawDir "$sourceId.raw.json"
        Write-Utf8File $rawFile $text
        $obj = $text | ConvertFrom-Json
        $itemsRaw = $obj.response.body.items.item
        if ($null -eq $itemsRaw) {
            $items = @()
        } elseif ($itemsRaw -is [array]) {
            $items = $itemsRaw
        } else {
            $items = @($itemsRaw)
        }
        foreach ($row in $items) {
            $title = Get-Prop $row @("title")
            $addr = Get-Prop $row @("addr")
            $summary = Get-Prop $row @("summary")
            Add-Candidate $sourceId $name $provider $title $addr (Get-Sigungu $addr "") "" "" "자연·생태" "생태관광" $official (Get-Prop $row @("contentid")) "생태관광 개요=$summary" $summary (Get-Prop $row @("modifiedtime")) "추천후보" "" 90
        }
        Add-SourceStatus $sourceId $name $provider $official $endpoint "TourAPI GW JSON" "개발단계 자동승인" "수집성공" $items.Count $items.Count $rawFile "경북(areaCode=35) 조회 결과 1건. 좌표는 Tour API 상세/기존 관광 API와 보강 필요."
    } catch {
        Add-SourceStatus $sourceId $name $provider $official $endpoint "TourAPI GW JSON" "개발단계 자동승인" "수집실패" 0 0 "" $_.Exception.Message
    }
}

function Process-NationalParkTrails {
    $sourceId = "S03_NATIONAL_PARK_TRAIL"
    $name = "국립공원공단 국립공원 탐방로 공간데이터"
    $provider = "국립공원공단"
    $official = "https://www.data.go.kr/data/15003467/fileData.do"
    $endpoint = "https://www.data.go.kr/cmm/cmm/fileDownload.do?atchFileId=FILE_000000002823639&fileDetailSn=1&insertDataPrcus=N"
    try {
        $zipPath = Join-Path $RawDir "$sourceId.raw.zip"
        Invoke-CurlDownload $endpoint $zipPath
        $extractDir = Join-Path $WorkDir $sourceId
        New-Item -ItemType Directory -Force -Path $extractDir | Out-Null
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $extractDir)
        $csvFiles = @(Get-ChildItem -Path $extractDir -Recurse -Filter *.csv)
        $rawCount = 0
        $courseMap = @{}
        foreach ($file in $csvFiles) {
            $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($file.FullName, [System.Text.Encoding]::Default)
            $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
            $parser.SetDelimiters(",")
            $parser.HasFieldsEnclosedInQuotes = $true
            try {
                $headers = $parser.ReadFields()
                $idxCourse = [Array]::IndexOf($headers, "탐방코스(한글)")
                $idxDetail = [Array]::IndexOf($headers, "상세구간")
                $idxDifficulty = [Array]::IndexOf($headers, "난이도")
                $idxDistance = [Array]::IndexOf($headers, "지리정보시스템 상 거리(m)")
                $idxControl = [Array]::IndexOf($headers, "탐방로 통제여부")
                $idxControlDesc = [Array]::IndexOf($headers, "통제구간 설명")
                $idxLon = [Array]::IndexOf($headers, "경도")
                $idxLat = [Array]::IndexOf($headers, "위도")
                $idxParkNo = [Array]::IndexOf($headers, "국립공원관리번호")
                $idxCourseId = [Array]::IndexOf($headers, "코스ID")
                while (-not $parser.EndOfData) {
                    $fields = $parser.ReadFields()
                    $rawCount++
                    if ($idxLat -lt 0 -or $idxLon -lt 0 -or $fields.Count -le [math]::Max($idxLat, $idxLon)) { continue }
                    $lat = $fields[$idxLat]
                    $lon = $fields[$idxLon]
                    if (-not (Get-BoundingGyeongbuk $lat $lon)) { continue }
                    $course = if ($idxCourse -ge 0) { $fields[$idxCourse] } else { "국립공원 탐방로" }
                    $detail = if ($idxDetail -ge 0) { $fields[$idxDetail] } else { "" }
                    $parkNo = if ($idxParkNo -ge 0) { $fields[$idxParkNo] } else { "" }
                    $courseId = if ($idxCourseId -ge 0) { $fields[$idxCourseId] } else { "" }
                    $key = "$parkNo|$courseId|$course|$detail"
                    if (-not $courseMap.ContainsKey($key)) {
                        $courseMap[$key] = [pscustomobject]@{
                            국립공원관리번호 = $parkNo
                            코스ID = $courseId
                            탐방코스 = $course
                            상세구간 = $detail
                            난이도 = if ($idxDifficulty -ge 0) { $fields[$idxDifficulty] } else { "" }
                            거리_m = if ($idxDistance -ge 0) { $fields[$idxDistance] } else { "" }
                            통제여부 = if ($idxControl -ge 0) { $fields[$idxControl] } else { "" }
                            통제설명 = if ($idxControlDesc -ge 0) { $fields[$idxControlDesc] } else { "" }
                            대표경도 = $lon
                            대표위도 = $lat
                            점수점개수 = 0
                            원본파일 = $file.Name
                        }
                    }
                    $courseMap[$key].점수점개수 = [int]$courseMap[$key].점수점개수 + 1
                }
            } finally {
                $parser.Close()
            }
        }
        $gb = @($courseMap.Values | Sort-Object 탐방코스, 상세구간)
        $filteredCsv = Join-Path $RawDir "$sourceId.gyeongbuk_course_candidates.csv"
        $gb | Export-Csv -Path $filteredCsv -NoTypeInformation -Encoding UTF8
        foreach ($row in $gb) {
            $title = "$($row.탐방코스)"
            $detail = "코스ID=$($row.코스ID); 상세구간=$($row.상세구간); 난이도=$($row.난이도); 거리m=$($row.거리_m); 통제=$($row.통제여부) $($row.통제설명); 좌표점수=$($row.점수점개수)"
            Add-Candidate $sourceId $name $provider $title "" "" $row.대표위도 $row.대표경도 "산지·탐방로" "국립공원 탐방로" $official "$($row.국립공원관리번호)-$($row.코스ID)" "공간데이터 필드=$detail" $detail "" "검수필요_좌표기반경북판정" "" 65
        }
        Add-SourceStatus $sourceId $name $provider $official $endpoint "파일데이터 ZIP/CSV" "로그인 없이 파일 다운로드 가능; 자동변환 API는 활용신청 필요" "수집성공" $rawCount $gb.Count $zipPath "원본 CSV가 181MB라 raw ZIP을 보존하고, 좌표 경계에 걸린 코스 단위 후보만 별도 CSV로 추출."
    } catch {
        Add-SourceStatus $sourceId $name $provider $official $endpoint "파일데이터 ZIP/CSV" "파일 다운로드 가능" "수집실패" 0 0 "" $_.Exception.Message
    }
}

function Process-ForestTrails {
    $sourceId = "S04_FOREST_TRAIL"
    $name = "산림청 산림공간정보 등산로정보"
    $provider = "산림청"
    $official = "https://www.data.go.kr/data/15002734/openapi.do"
    $endpoint = "http://api.forest.go.kr/openapi/service/trailInfoService/getforestspatialdataservice"
    try {
        $url = "${endpoint}?ServiceKey=$([System.Web.HttpUtility]::UrlEncode($ServiceKey))&pageNo=1&numOfRows=500"
        $text = Invoke-CurlText $url
        $rawFile = Join-Path $RawDir "$sourceId.raw.xml"
        Write-Utf8File $rawFile $text
        [xml]$xml = $text
        $items = @($xml.response.body.items.item) | Where-Object { $_ }
        $csvRows = foreach ($row in $items) {
            $info = [string]$row.mntninfourl
            $x = ""
            $y = ""
            if ($info -match "longitude=([0-9.]+).*latitude=([0-9.]+)") {
                $x = $Matches[1]
                $y = $Matches[2]
            }
            [pscustomobject]@{
                mntnnm = [string]$row.mntnnm
                mntnfile = [string]$row.mntnfile
                mntnimg = [string]$row.mntnimg
                mntninfourl = $info
                forestMapX = $x
                forestMapY = $y
            }
        }
        $rawCsv = Join-Path $RawDir "$sourceId.raw.csv"
        $csvRows | Export-Csv -Path $rawCsv -NoTypeInformation -Encoding UTF8
        $gyeongbukMountainNames = @(
            "팔공산","금오산","보현산","주왕산","가야산","소백산","청량산","내연산","운문산","비슬산",
            "황악산","문수산","일월산","백암산","응봉산","도덕산","기룡산","수도산","조항산","황학산",
            "유학산","어림산","선암산","천생산","비봉산","봉화산","구미 금오산","문경새재"
        )
        $gb = @($csvRows | Where-Object { $gyeongbukMountainNames -contains $_.mntnnm })
        foreach ($row in $gb) {
            Add-Candidate $sourceId $name $provider $row.mntnnm "" "" "" "" "산지·탐방로" "산림청 등산로" $official $row.mntnfile "등산로 파일=$($row.mntnfile); 안내URL=$($row.mntninfourl)" "산림청 등산로 API는 산명·파일·안내URL을 제공하나 주소 필드가 없어 경북 산명 사전으로 1차 선별." "" "검수필요_산명기반경북판정" "" 55
        }
        Add-SourceStatus $sourceId $name $provider $official $endpoint "OpenAPI XML" "개발단계 심의승인 API이나 현재 제공 키로 호출 성공" "수집성공" $items.Count $gb.Count $rawFile "전국 432건 수집. 주소가 없어 경북 후보는 산명 사전으로만 1차 판정; 좌표계 변환 또는 파일 내부 검수 필요."
    } catch {
        Add-SourceStatus $sourceId $name $provider $official $endpoint "OpenAPI XML" "개발단계 심의승인" "수집실패" 0 0 "" $_.Exception.Message
    }
}

Add-Process "1. 원천 접근 확인" "공공데이터포털 공식 상세 페이지" "요청주소, 승인 절차, 데이터포맷, 수정일 확인" "수집 대상별 접근 상태표" "심사 보고서에서 실제 연결/미연결을 구분하기 위한 근거"
Add-Process "2. 원자료 수집" "표준 OpenAPI, TourAPI GW, 파일데이터, 산림청 XML" "서비스키로 실제 호출하고 raw 폴더에 JSON/XML/CSV/ZIP 저장" "원자료 보존 파일" "연결한 척이 아니라 재현 가능한 원자료를 남김"
Add-Process "3. 경북 필터링" "주소, 시도명, 좌표, 공원명/산명" "경상북도 주소 우선, 주소 없는 탐방로는 좌표/산명 1차 판정 후 검수필요 표시" "경북 후보 목록" "지역 무관 데이터가 추천에 섞이는 문제 방지"
Add-Process "4. 비학습 장소 제외" "장소명·시설명·유형" "음식점·카페·숙박·상업시설 키워드, 교육 근거 약한 어린이공원/소공원 제외" "제외 후보 시트" "가족여행 추천이 맛집·숙소 추천으로 오염되는 문제 방지"
Add-Process "5. 교육 맥락 부여" "개요, 소개, 프로그램, 축제내용, 보유시설" "자연·생태, 산지·탐방로, 공원·생활환경, 농어촌·생산체험, 지역축제, 관광명소, 박물관·전시로 분류" "교육과정 연계후보/직접근거" "장소와 단원이 실제로 만나는 필드 기반 근거를 팝업에 표시하기 위함"
Add-Process "6. 중복 제거" "장소명+시군 정규화 키" "공백·괄호·기호 제거 후 같은 시군 내 중복 묶음 생성, 점수 높은 행 우선 유지" "중복후보/최종마스터" "같은 장소가 여러 공공데이터에서 반복 추천되는 문제 방지"

Process-EcoTour
Process-CityParks
Process-NationalParkTrails
Process-ForestTrails
Process-RuralVillages
Process-Festivals
Process-TouristAreas
Process-Museums

$CategoryPlanRows = @(
    [pscustomobject]@{추천카테고리="전체"; 용도="전체 후보 보기"; 포함데이터="모든 통과 후보"; UI메모="기존 전체 탭 유지"},
    [pscustomobject]@{추천카테고리="문화유산"; 용도="국가유산청·문화재 데이터용"; 포함데이터="지정문화재, 문화재 공간정보"; UI메모="다음 연결 데이터와 병합"},
    [pscustomobject]@{추천카테고리="현충·근현대사"; 용도="5학년 6단원 독립운동·국가수호"; 포함데이터="국가보훈부 현충시설, 전쟁/독립 키워드 관광지"; UI메모="현재 보훈 API와 별도 탭 권장"},
    [pscustomobject]@{추천카테고리="자연·생태"; 용도="자연환경·생태·환경 보전"; 포함데이터="생태 관광 정보, 자연형 관광지"; UI메모="기존 자연탐구 탭 확장"},
    [pscustomobject]@{추천카테고리="산지·탐방로"; 용도="지형·산지·등산 계획"; 포함데이터="국립공원 탐방로, 산림청 등산로"; UI메모="장소 카드가 아니라 코스/탐방로 카드 형식 필요"},
    [pscustomobject]@{추천카테고리="공원·생활환경"; 용도="생활권·공공시설·지도 활용"; 포함데이터="전국도시공원"; UI메모="교육근거 약한 어린이공원은 기본 숨김"},
    [pscustomobject]@{추천카테고리="농어촌·생산체험"; 용도="생산·소비, 촌락 생활, 체험학습"; 포함데이터="농어촌체험휴양마을"; UI메모="체험 프로그램 칩 표시"},
    [pscustomobject]@{추천카테고리="지역축제"; 용도="지역 문화·계절성 활동"; 포함데이터="전국문화축제"; UI메모="2025 자료는 전년도 참고 배지 표시"},
    [pscustomobject]@{추천카테고리="관광명소"; 용도="Tour API 누락 관광지 보완"; 포함데이터="전국관광지정보"; UI메모="직접 교육근거 약하면 보조 연계로 낮춤"},
    [pscustomobject]@{추천카테고리="박물관·전시"; 용도="역사·문화·과학 전시 시설"; 포함데이터="박물관·미술관 표준데이터"; UI메모="운영시간/휴관정보 표시"}
)

$allMaster = $MasterRows.ToArray()
$groups = $allMaster | Group-Object 중복키 | Where-Object { $_.Count -gt 1 -and -not [string]::IsNullOrWhiteSpace($_.Name) }
foreach ($group in $groups) {
    $sorted = @($group.Group | Sort-Object @{Expression={[int]$_.점수};Descending=$true}, 데이터명)
    $keep = $sorted | Select-Object -First 1
    foreach ($dup in ($sorted | Select-Object -Skip 1)) {
        $DuplicateRows.Add([pscustomobject]@{
            중복키 = $group.Name
            유지장소 = $keep.장소명
            유지원천 = $keep.데이터명
            중복장소 = $dup.장소명
            중복원천 = $dup.데이터명
            시군 = $dup.시군
            처리 = "최종마스터에서는 점수 높은 행 우선 유지, 나머지는 검수 참고"
        }) | Out-Null
    }
}

$deduped = @()
foreach ($group in ($allMaster | Group-Object 중복키)) {
    if ([string]::IsNullOrWhiteSpace($group.Name)) {
        $deduped += $group.Group
    } else {
        $deduped += @($group.Group | Sort-Object @{Expression={[int]$_.점수};Descending=$true}, 데이터명 | Select-Object -First 1)
    }
}

$deduped = @($deduped | Sort-Object 시군, 추천카테고리, 장소명)

function ConvertTo-ExcelColumn {
    param([int]$Index)
    $name = ""
    while ($Index -gt 0) {
        $rem = ($Index - 1) % 26
        $name = [char](65 + $rem) + $name
        $Index = [math]::Floor(($Index - 1) / 26)
    }
    return $name
}

function Escape-Xml {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return [System.Security.SecurityElement]::Escape($Text)
}

function New-SheetXml {
    param([object[]]$Rows, [string[]]$Columns)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
    $r = 1
    [void]$sb.Append("<row r=`"$r`">")
    for ($c = 0; $c -lt $Columns.Count; $c++) {
        $ref = "$(ConvertTo-ExcelColumn ($c + 1))$r"
        [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`"><is><t>$(Escape-Xml $Columns[$c])</t></is></c>")
    }
    [void]$sb.Append("</row>")
    $r++
    foreach ($row in $Rows) {
        [void]$sb.Append("<row r=`"$r`">")
        for ($c = 0; $c -lt $Columns.Count; $c++) {
            $col = $Columns[$c]
            $value = ""
            $prop = $row.PSObject.Properties[$col]
            if ($prop) { $value = [string]$prop.Value }
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                $ref = "$(ConvertTo-ExcelColumn ($c + 1))$r"
                [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`"><is><t>$(Escape-Xml $value)</t></is></c>")
            }
        }
        [void]$sb.Append("</row>")
        $r++
    }
    [void]$sb.Append("</sheetData></worksheet>")
    return $sb.ToString()
}

function Export-SimpleXlsx {
    param([hashtable[]]$Sheets, [string]$Path)
    $temp = Join-Path $WorkDir "xlsx"
    if (Test-Path $temp) { Remove-Item -Path $temp -Recurse -Force }
    New-Item -ItemType Directory -Force -Path (Join-Path $temp "_rels"), (Join-Path $temp "xl"), (Join-Path $temp "xl\worksheets"), (Join-Path $temp "xl\_rels") | Out-Null

    $contentTypes = New-Object System.Text.StringBuilder
    [void]$contentTypes.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
    for ($i = 1; $i -le $Sheets.Count; $i++) {
        [void]$contentTypes.Append("<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>")
    }
    [void]$contentTypes.Append("</Types>")
    Write-Utf8File (Join-Path $temp "[Content_Types].xml") $contentTypes.ToString()
    Write-Utf8File (Join-Path $temp "_rels\.rels") '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    Write-Utf8File (Join-Path $temp "xl\styles.xml") '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="맑은 고딕"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>'

    $workbook = New-Object System.Text.StringBuilder
    [void]$workbook.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>')
    $rels = New-Object System.Text.StringBuilder
    [void]$rels.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')

    for ($i = 0; $i -lt $Sheets.Count; $i++) {
        $sheetId = $i + 1
        $sheetName = $Sheets[$i].Name
        if ($sheetName.Length -gt 31) { $sheetName = $sheetName.Substring(0, 31) }
        $sheetName = $sheetName -replace "[:\\/\?\*\[\]]", "_"
        [void]$workbook.Append("<sheet name=`"$(Escape-Xml $sheetName)`" sheetId=`"$sheetId`" r:id=`"rId$sheetId`"/>")
        [void]$rels.Append("<Relationship Id=`"rId$sheetId`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$sheetId.xml`"/>")
        Write-Utf8File (Join-Path $temp "xl\worksheets\sheet$sheetId.xml") (New-SheetXml $Sheets[$i].Rows $Sheets[$i].Columns)
    }
    [void]$workbook.Append("</sheets></workbook>")
    [void]$rels.Append("</Relationships>")
    Write-Utf8File (Join-Path $temp "xl\workbook.xml") $workbook.ToString()
    Write-Utf8File (Join-Path $temp "xl\_rels\workbook.xml.rels") $rels.ToString()

    if (Test-Path $Path) { Remove-Item -Path $Path -Force }
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew)
    $archive = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        Get-ChildItem -Path $temp -Recurse -File | ForEach-Object {
            $rel = $_.FullName.Substring($temp.Length + 1).Replace("\", "/")
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $_.FullName, $rel, [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
        }
    } finally {
        $archive.Dispose()
        $fs.Dispose()
    }
}

$sourceColumns = @("수집일시","원천ID","데이터명","제공기관","공식확인URL","요청주소_키제외","접근방식","승인상태","수집상태","원자료건수","경북필터건수","원자료파일","비고")
$processColumns = @("단계","입력자료","처리규칙","산출물","교육적_검수_의미")
$masterColumns = @("원천ID","데이터명","제공기관","장소명","시군","주소","위도","경도","추천카테고리","세부유형","추천상태","점수","교육과정_연계후보","직접연계근거","설명요약","원자료ID","기준일자","공식출처","중복키")
$excludedColumns = @("원천ID","데이터명","장소명","주소","제외사유","원자료ID")
$duplicateColumns = @("중복키","유지장소","유지원천","중복장소","중복원천","시군","처리")
$categoryColumns = @("추천카테고리","용도","포함데이터","UI메모")

Export-SimpleXlsx @(
    @{ Name = "00_수집요약"; Rows = $SourceRows.ToArray(); Columns = $sourceColumns },
    @{ Name = "01_전처리과정"; Rows = $ProcessRows.ToArray(); Columns = $processColumns },
    @{ Name = "02_최종마스터_중복제거"; Rows = @($deduped); Columns = $masterColumns },
    @{ Name = "03_경북필터_전체후보"; Rows = $MasterRows.ToArray(); Columns = $masterColumns },
    @{ Name = "04_제외후보"; Rows = $ExcludedRows.ToArray(); Columns = $excludedColumns },
    @{ Name = "05_중복후보"; Rows = $DuplicateRows.ToArray(); Columns = $duplicateColumns },
    @{ Name = "06_카테고리재설계"; Rows = @($CategoryPlanRows); Columns = $categoryColumns }
) $XlsxPath

$xlsxLeaf = Split-Path -Leaf $XlsxPath
$outLeaf = Split-Path -Leaf $OutDir
$summary = @"
# 추가 공공데이터 수집·전처리 요약 ($RunDate)

수집 시각: $CollectedAt

## 산출물

- 엑셀: $xlsxLeaf
- 원자료 폴더: $outLeaf\raw

## 실제 수집 상태

$($SourceRows | ForEach-Object { "- $($_.데이터명): $($_.수집상태), 원자료 $($_.원자료건수)건, 경북 필터 $($_.경북필터건수)건. $($_.비고)" } | Out-String)

## 앱 카테고리 재설계 메모

기존 `문화유산/자연탐구/과학체험` 중심 탭만으로는 새 데이터가 잘 담기지 않는다. `산지·탐방로`, `농어촌·생산체험`, `지역축제`, `공원·생활환경`, `박물관·전시`, `현충·근현대사`를 별도 카테고리로 분리하는 편이 교육적 맥락과 데이터 출처를 설명하기 쉽다.

## 주의

- 산림청 등산로 API는 주소 필드가 없어 경북 후보를 산명 사전으로 1차 판정했다. 실제 앱 반영 전에는 좌표계 변환 또는 파일 내부 검수가 필요하다.
- 국립공원 탐방로 공간데이터는 파일 직접 다운로드로 수집했다. 경북 판정은 좌표 경계와 공원명 키워드 기반이므로 지도 검수가 필요하다.
- 문화축제는 상설 장소가 아니므로 2025년 종료 자료는 `전년도 참고`로 낮춰 표시해야 한다.
- 도시공원은 어린이공원·소공원이 많아 교육 근거가 약한 생활놀이공간은 기본 제외했다.
"@
Write-Utf8File $SummaryPath $summary

Write-Output "DONE"
Write-Output "XLSX=$XlsxPath"
Write-Output "OUTDIR=$OutDir"
Write-Output "SUMMARY=$SummaryPath"
