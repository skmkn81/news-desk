# ============================================================
#  뉴스 데스크 - 1단계 수집기
#  RSS에서 기사를 모아 data\news.json 으로 저장합니다.
#  실행 방법: 이 파일에서 마우스 오른쪽 클릭 > "PowerShell에서 실행"
#  ============================================================

$ErrorActionPreference = 'Stop'

# 윈도우의 옛 PowerShell에서 필요한 설정입니다.
# 리눅스(깃허브 서버)의 PowerShell 7에서는 필요 없고 오류가 날 수 있어 감싸 둡니다.
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

$Root = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }

$CfgPath    = Join-Path $Root '설정.json'
$DataDir    = Join-Path $Root 'data'
$RawRoot    = Join-Path $DataDir 'raw'
$NewsPath   = Join-Path $DataDir 'news.json'
$NewsJsPath = Join-Path $DataDir 'news.js'
$StatusPath = Join-Path $DataDir 'status.json'

$RunStart   = Get-Date
$RunStamp   = $RunStart.ToString('yyyy-MM-dd_HHmm')
$RunIso     = $RunStart.ToString('yyyy-MM-ddTHH:mm:sszzz')
$RawDir     = Join-Path $RawRoot $RunStamp

foreach ($d in @($DataDir, $RawRoot, $RawDir)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

# ---------- 도우미 함수 ----------

function Write-Utf8NoBom {
    param([string]$Path, [string]$Text)
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Text, $enc)
}

# XML 노드에서 글자만 꺼냅니다. CDATA로 감싸인 경우도 처리합니다.
function Get-NodeText {
    param($Node)
    if ($null -eq $Node) { return '' }
    if ($Node -is [string]) { return $Node.Trim() }
    if ($Node -is [System.Xml.XmlElement]) { return ([string]$Node.InnerText).Trim() }
    return ([string]$Node).Trim()
}

# HTML 태그를 걷어내고 &amp; 같은 문자를 되돌립니다.
function ConvertFrom-Html {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $t = [regex]::Replace($Text, '<[^>]+>', ' ')
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    $t = $t -replace '\s+', ' '
    return $t.Trim()
}

# RSS 날짜 문자열을 ISO 8601(UTC)로 바꿉니다.
function ConvertTo-IsoUtc {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return '' }
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    try {
        $dto = [System.DateTimeOffset]::Parse($Raw.Trim(), $inv)
        return $dto.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    } catch {
        return ''
    }
}

# 링크에서 추적용 꼬리표를 떼어 비교용 열쇠를 만듭니다.
function Get-LinkKey {
    param([string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return '' }
    $u = $Url.Trim()
    $u = [regex]::Replace($u, '([?&])(utm_[^&]*|fbclid=[^&]*|gclid=[^&]*|ref=[^&]*)', '$1')
    $u = $u -replace '[?&]+$', ''
    $u = $u -replace '#.*$', ''
    $u = $u -replace '^https?://', ''
    $u = $u -replace '^www\.', ''
    $u = $u -replace '/+$', ''
    return $u.ToLowerInvariant()
}

# 제목에서 공백·특수문자를 지워 비교용 열쇠를 만듭니다.
function Get-TitleKey {
    param([string]$Title)
    if ([string]::IsNullOrWhiteSpace($Title)) { return '' }
    $t = $Title.ToLowerInvariant()
    $t = [regex]::Replace($t, '[\s\p{P}\p{S}]', '')
    return $t
}

# 예전에 저장된 기사에는 없을 수도 있는 속성을 안전하게 설정합니다.
function Set-Prop {
    param($Object, [string]$Name, $Value)
    if ($Object.PSObject.Properties.Name -contains $Name) {
        $Object.$Name = $Value
    } else {
        $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force
    }
}

function Get-Feed {
    param([string]$Url)
    $lastErr = $null
    for ($i = 1; $i -le 2; $i++) {
        try {
            $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 25 `
                -MaximumRedirection 5 -Headers @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }
            $body = $resp.Content
            # PowerShell 7 은 응답을 바이트 배열로 줄 때가 있어 글자로 되돌립니다.
            if ($body -is [byte[]]) { $body = [System.Text.Encoding]::UTF8.GetString($body) }
            return [string]$body
        } catch {
            $lastErr = $_.Exception.Message
            Start-Sleep -Seconds 2
        }
    }
    throw $lastErr
}

# ---------- 설정 읽기 ----------

if (-not (Test-Path $CfgPath)) { throw "설정 파일을 찾을 수 없습니다: $CfgPath" }
$cfg = Get-Content -Path $CfgPath -Raw -Encoding UTF8 | ConvertFrom-Json

$keepDays     = if ($cfg.keepDays)     { [int]$cfg.keepDays }     else { 30 }
$maxPerSource = if ($cfg.maxPerSource) { [int]$cfg.maxPerSource } else { 100 }

Write-Host ''
Write-Host '=== 뉴스 수집 시작 ===' -ForegroundColor Cyan
Write-Host ("실행 시각 : {0}" -f $RunStart.ToString('yyyy-MM-dd HH:mm:ss'))
Write-Host ("원본 저장 : {0}" -f $RawDir)
Write-Host ''

# ---------- 기존 기사 불러오기 ----------

$articles  = New-Object System.Collections.ArrayList
$linkSeen  = @{}
$titleSeen = @{}

# 이전 기사 목록은 news.json 에서 읽고, 없으면 news.js 에서 읽습니다.
# (깃허브에는 용량을 아끼려고 news.js 만 올라갑니다)
$oldRaw = ''
if (Test-Path $NewsPath) {
    $oldRaw = Get-Content -Path $NewsPath -Raw -Encoding UTF8
} elseif (Test-Path $NewsJsPath) {
    $oldRaw = Get-Content -Path $NewsJsPath -Raw -Encoding UTF8
    $oldRaw = $oldRaw -replace '^\s*window\.NEWS_DATA\s*=\s*', ''
    $oldRaw = $oldRaw -replace ';\s*$', ''
}

if ($oldRaw) {
    try {
        $old = $oldRaw | ConvertFrom-Json
        foreach ($a in @($old.articles)) {
            # JSON을 읽을 때 날짜 글자가 날짜 자료형으로 바뀌는 경우가 있습니다.
            # 자료형이 섞이면 나중에 정렬이 실패하므로 전부 글자로 되돌립니다.
            if ($a.date -is [datetime]) {
                Set-Prop $a 'date' ($a.date.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))
            } else {
                Set-Prop $a 'date' ([string]$a.date)
            }
            Set-Prop $a 'collectedAt' ([string]$a.collectedAt)

            $lk = Get-LinkKey $a.link
            $tk = Get-TitleKey $a.title
            if ($lk -and $linkSeen.ContainsKey($lk))  { continue }
            if ($tk -and $titleSeen.ContainsKey($tk)) { continue }
            [void]$articles.Add($a)
            if ($lk) { $linkSeen[$lk]  = $articles.Count - 1 }
            if ($tk) { $titleSeen[$tk] = $articles.Count - 1 }
        }
        Write-Host ("기존 기사  : {0}건 불러옴" -f $articles.Count) -ForegroundColor DarkGray
    } catch {
        Write-Host ("기존 news.json 을 읽지 못해 새로 시작합니다: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}

# ---------- 이전 수집 상태 불러오기 (마지막 성공 시각 유지용) ----------

$prevSuccess = @{}
if (Test-Path $StatusPath) {
    try {
        $ps = Get-Content -Path $StatusPath -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($s in @($ps.sources)) {
            if ($s.lastSuccessAt) { $prevSuccess[$s.id] = $s.lastSuccessAt }
        }
    } catch { }
}

# ---------- 출처별 수집 ----------

$statusRows = New-Object System.Collections.ArrayList
$addedTotal = 0
$dupTotal   = 0

foreach ($src in @($cfg.sources)) {

    if (-not $src.enabled) {
        [void]$statusRows.Add([pscustomobject]@{
            id = $src.id; label = $src.label; tab = $src.tab; url = ''
            ok = $false; skipped = $true; count = 0; added = 0
            lastSuccessAt = $(if ($prevSuccess.ContainsKey($src.id)) { $prevSuccess[$src.id] } else { '' })
            error = '설정에서 꺼져 있음'
        })
        Write-Host ("건너뜀 | {0}" -f $src.label) -ForegroundColor DarkGray
        continue
    }

    # 출처 종류에 따라 실제 주소를 만듭니다.
    switch ($src.kind) {
        'google-search' {
            $url = 'https://news.google.com/rss/search?q=' + [uri]::EscapeDataString($src.query) + '&hl=ko&gl=KR&ceid=KR:ko'
        }
        'google-topic' {
            $url = 'https://news.google.com/rss/headlines/section/topic/' + $src.topic + '?hl=ko&gl=KR&ceid=KR:ko'
        }
        default {
            $url = $src.url
        }
    }

    $isGoogle = $src.kind -like 'google-*'
    $added = 0; $dup = 0; $count = 0; $ok = $false; $err = ''

    try {
        $content = Get-Feed -Url $url

        # 원본을 그대로 보관합니다.
        Write-Utf8NoBom -Path (Join-Path $RawDir ($src.id + '.xml')) -Text $content

        $xml   = [xml]$content
        $items = @($xml.rss.channel.item)
        if ($items.Count -eq 0) { $items = @($xml.SelectNodes('//item')) }
        $count = $items.Count

        foreach ($it in ($items | Select-Object -First $maxPerSource)) {

            $title = ConvertFrom-Html (Get-NodeText $it.title)
            $link  = Get-NodeText $it.link
            if (-not $title -or -not $link) { continue }

            # 구글 뉴스는 제목 뒤에 " - 언론사" 가 붙어 옵니다.
            $publisher = ''
            if ($it.source) { $publisher = Get-NodeText $it.source }
            if (-not $publisher -and $src.publisher) { $publisher = $src.publisher }
            if ($isGoogle -and $publisher -and $title.EndsWith(' - ' + $publisher)) {
                $title = $title.Substring(0, $title.Length - ($publisher.Length + 3)).Trim()
            }

            # 미리보기: 구글 뉴스는 본문 요약을 주지 않으므로 비워 둡니다.
            $preview = ''
            if (-not $isGoogle) {
                $preview = ConvertFrom-Html (Get-NodeText $it.description)
                # 미리보기가 제목을 그대로 반복하는 경우도 비워 둡니다.
                if ((Get-TitleKey $preview) -eq (Get-TitleKey $title)) { $preview = '' }
                if ($preview.Length -gt 300) { $preview = $preview.Substring(0, 300).TrimEnd() + '...' }
            }

            $date = ConvertTo-IsoUtc (Get-NodeText $it.pubDate)
            if (-not $date) { $date = $RunStart.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }

            $lk = Get-LinkKey $link
            $tk = Get-TitleKey $title

            # --- 중복 제거 ---
            $hitIndex = -1
            if ($lk -and $linkSeen.ContainsKey($lk))       { $hitIndex = $linkSeen[$lk] }
            elseif ($tk -and $titleSeen.ContainsKey($tk))  { $hitIndex = $titleSeen[$tk] }

            if ($hitIndex -ge 0) {
                $dup++
                $exist = $articles[$hitIndex]
                # 이미 있는 기사면 탭 정보만 합치고, 비어 있던 미리보기는 채워 줍니다.
                $tabs = @($exist.tabs) | Where-Object { $_ }
                if ($tabs -notcontains $src.tab) {
                    Set-Prop $exist 'tabs' (@($tabs) + $src.tab)
                }
                if (-not $exist.preview -and $preview) { Set-Prop $exist 'preview' $preview }
                continue
            }

            $row = [pscustomobject]@{
                title       = $title
                preview     = $preview
                source      = $publisher
                date        = $date
                link        = $link
                tab         = $src.tab
                tabs        = @($src.tab)
                feed        = $src.label
                feedId      = $src.id
                collectedAt = $RunIso
            }
            [void]$articles.Add($row)
            if ($lk) { $linkSeen[$lk]  = $articles.Count - 1 }
            if ($tk) { $titleSeen[$tk] = $articles.Count - 1 }
            $added++
        }

        $ok = $true
        $lastSuccess = $RunIso
        Write-Host ("성공   | {0,-30} | 받음 {1,3}건 | 새 기사 {2,3}건 | 중복 {3,3}건" -f $src.label, $count, $added, $dup) -ForegroundColor Green

    } catch {
        $err = $_.Exception.Message
        $lastSuccess = $(if ($prevSuccess.ContainsKey($src.id)) { $prevSuccess[$src.id] } else { '' })
        Write-Host ("실패   | {0,-30} | {1}" -f $src.label, $err) -ForegroundColor Red
    }

    $addedTotal += $added
    $dupTotal   += $dup

    [void]$statusRows.Add([pscustomobject]@{
        id = $src.id; label = $src.label; tab = $src.tab; url = $url
        ok = $ok; skipped = $false; count = $count; added = $added
        lastSuccessAt = $lastSuccess
        error = $err
    })
}

# ---------- 오래된 기사 정리 & 정렬 ----------

$cutoff = $RunStart.ToUniversalTime().AddDays(-$keepDays).ToString('yyyy-MM-ddTHH:mm:ssZ')
# 날짜는 반드시 글자로 비교·정렬합니다. (자료형이 섞이면 PowerShell 7에서 정렬이 실패합니다)
$final  = @(
    $articles |
        Where-Object { ([string]$_.date) -ge $cutoff } |
        Sort-Object -Property @{ Expression = { [string]$_.date } } -Descending
)

# ---------- 저장 ----------

$tabCounts = [ordered]@{}
foreach ($t in @($cfg.tabs)) {
    $tabCounts[$t.id] = @($final | Where-Object { @($_.tabs) -contains $t.id }).Count
}

$newsDoc = [pscustomobject]@{
    generatedAt = $RunIso
    keepDays    = $keepDays
    tabs        = @($cfg.tabs)
    tabCounts   = $tabCounts
    total       = $final.Count
    articles    = $final
}
$newsJson = $newsDoc | ConvertTo-Json -Depth 8
Write-Utf8NoBom -Path $NewsPath -Text $newsJson

# 브라우저에서 파일을 더블클릭해 열어도 읽을 수 있도록 .js 형태로도 내보냅니다.
# (file:// 로 열면 보안 정책 때문에 .json 을 직접 못 읽습니다)
Write-Utf8NoBom -Path $NewsJsPath -Text ("window.NEWS_DATA = " + $newsJson + ";" + [Environment]::NewLine)

$okCount = @($statusRows | Where-Object { $_.ok }).Count
$statusDoc = [pscustomobject]@{
    lastRunAt        = $RunIso
    lastSuccessAt    = $(if ($okCount -gt 0) { $RunIso } else { '' })
    rawFolder        = ('data/raw/' + $RunStamp)
    sourcesTotal     = @($cfg.sources).Count
    sourcesSucceeded = $okCount
    newArticles      = $addedTotal
    duplicatesSkipped= $dupTotal
    totalArticles    = $final.Count
    tabCounts        = $tabCounts
    sources          = @($statusRows)
}
Write-Utf8NoBom -Path $StatusPath -Text ($statusDoc | ConvertTo-Json -Depth 6)

# ---------- 결과 요약 ----------

Write-Host ''
Write-Host '=== 수집 완료 ===' -ForegroundColor Cyan
Write-Host ("출처 성공  : {0} / {1}" -f $okCount, @($cfg.sources).Count)
Write-Host ("새 기사    : {0}건" -f $addedTotal)
Write-Host ("중복 제외  : {0}건" -f $dupTotal)
Write-Host ("전체 보관  : {0}건 (최근 {1}일)" -f $final.Count, $keepDays)
Write-Host ''
Write-Host '탭별 기사 수' -ForegroundColor Cyan
foreach ($t in @($cfg.tabs)) {
    Write-Host ("  {0,-10} {1,5}건" -f $t.label, $tabCounts[$t.id])
}
Write-Host ''
Write-Host ("저장됨 : {0}" -f $NewsPath)
Write-Host ("저장됨 : {0}" -f $NewsJsPath)
Write-Host ("저장됨 : {0}" -f $StatusPath)
Write-Host ("원본   : {0}" -f $RawDir)
Write-Host ''
