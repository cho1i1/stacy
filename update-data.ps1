# Pulls Yahoo Finance data for the watch list and writes data.json for the phone page.
# Works in Windows PowerShell 5.1 and PowerShell 7, so it runs the same locally and on GitHub.
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$outFile = Join-Path $here 'data.json'
$stored = Get-Content (Join-Path $here 'stacy-data.json') -Raw -Encoding UTF8 | ConvertFrom-Json

$symbols = @(Get-Content (Join-Path $here 'watchlist.md') -Encoding UTF8 |
    ForEach-Object { if ($_ -match '^\s*-\s+([A-Za-z0-9.\-]+)\s+-\s') { $Matches[1].ToUpper() } } |
    Select-Object -Unique)
if ($symbols.Count -eq 0) { throw 'No tickers found in watchlist.md' }

$userAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36'
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$session.UserAgent = $userAgent
$base = 'https://query1.finance.yahoo.com'

# Visiting this host sets the cookie Yahoo requires. It normally answers with an error page, which is expected.
try { Invoke-WebRequest -Uri 'https://fc.yahoo.com' -WebSession $session -UseBasicParsing -TimeoutSec 30 | Out-Null } catch { }

$crumb = (Invoke-WebRequest -Uri "$base/v1/test/getcrumb" -WebSession $session -UseBasicParsing -TimeoutSec 30).Content
if ([string]::IsNullOrWhiteSpace($crumb) -or $crumb.Length -gt 50) { throw 'Yahoo did not return a valid crumb' }
$crumbEncoded = [uri]::EscapeDataString($crumb)

$trendingRaw = Invoke-RestMethod -Uri "$base/v1/finance/trending/US?count=25" -WebSession $session -TimeoutSec 30
$trendingSymbols = @($trendingRaw.finance.result[0].quotes | ForEach-Object { $_.symbol })

$allSymbols = @($symbols + $trendingSymbols | Select-Object -Unique)
$bySymbol = @{}
for ($i = 0; $i -lt $allSymbols.Count; $i += 100) {
    $last = [Math]::Min($i + 99, $allSymbols.Count - 1)
    $joined = [uri]::EscapeDataString(($allSymbols[$i..$last] -join ','))
    $quotes = Invoke-RestMethod -Uri "$base/v7/finance/quote?symbols=$joined&crumb=$crumbEncoded" -WebSession $session -TimeoutSec 30
    foreach ($q in $quotes.quoteResponse.result) { $bySymbol[$q.symbol] = $q }
}

$found = @($symbols | Where-Object { $bySymbol.ContainsKey($_) }).Count
if ($found -lt [Math]::Ceiling($symbols.Count / 2)) { throw "Yahoo returned only $found of $($symbols.Count) tickers" }

# Sectors rarely change, so reuse the ones already saved and only look up new tickers.
$knownSectors = @{}
if (Test-Path $outFile) {
    try {
        $previous = Get-Content $outFile -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($p in $previous.stocks) { if ($p.sector) { $knownSectors[$p.s] = $p.sector } }
    } catch { }
}
$sectorMap = @{
    'Technology' = 'Information Technology'; 'Financial Services' = 'Financials'
    'Consumer Cyclical' = 'Consumer Discretionary'; 'Consumer Defensive' = 'Consumer Staples'
    'Healthcare' = 'Health Care'; 'Basic Materials' = 'Materials'; 'Energy' = 'Energy'
    'Industrials' = 'Industrials'; 'Communication Services' = 'Communication Services'
    'Utilities' = 'Utilities'; 'Real Estate' = 'Real Estate'
}

function Get-IsoDate($unix) {
    if (-not $unix) { return $null }
    return [DateTimeOffset]::FromUnixTimeSeconds([long]$unix).UtcDateTime.ToString('yyyy-MM-dd')
}

$stocks = New-Object System.Collections.ArrayList
$latest = [long]0
$missing = New-Object System.Collections.ArrayList
foreach ($s in $symbols) {
    if (-not $bySymbol.ContainsKey($s)) { [void]$missing.Add($s); continue }
    $q = $bySymbol[$s]
    if ([long]$q.regularMarketTime -gt $latest) { $latest = [long]$q.regularMarketTime }
    $sector = $null
    if ($q.quoteType -eq 'EQUITY') {
        if ($knownSectors.ContainsKey($s)) {
            $sector = $knownSectors[$s]
        } else {
            try {
                $profile = Invoke-RestMethod -Uri "$base/v10/finance/quoteSummary/${s}?modules=assetProfile&crumb=$crumbEncoded" -WebSession $session -TimeoutSec 30
                $sector = $sectorMap[$profile.quoteSummary.result[0].assetProfile.sector]
            } catch { }
        }
    }
    $chg = $null; if ($null -ne $q.regularMarketChangePercent) { $chg = [Math]::Round([double]$q.regularMarketChangePercent, 2) }
    $pe = $null; if ($null -ne $q.trailingPE) { $pe = [Math]::Round([double]$q.trailingPE, 2) }
    [void]$stocks.Add([ordered]@{ s = $s; type = $q.quoteType; price = $q.regularMarketPrice; chg = $chg; pe = $pe; sector = $sector; earn = (Get-IsoDate $q.earningsTimestamp) })
}

$trending = New-Object System.Collections.ArrayList
foreach ($s in $trendingSymbols) {
    if ($trending.Count -ge 3) { break }
    if ($bySymbol.ContainsKey($s) -and $bySymbol[$s].quoteType -eq 'EQUITY') {
        [void]$trending.Add([ordered]@{ s = $s; chg = [Math]::Round([double]$bySymbol[$s].regularMarketChangePercent, 2) })
    }
}

try { $eastern = [TimeZoneInfo]::FindSystemTimeZoneById('America/New_York') }
catch { $eastern = [TimeZoneInfo]::FindSystemTimeZoneById('Eastern Standard Time') }
$culture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
$nowEastern = [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $eastern)
$priceEastern = [TimeZoneInfo]::ConvertTimeFromUtc([DateTimeOffset]::FromUnixTimeSeconds($latest).UtcDateTime, $eastern)
$isLive = ($stocks | Where-Object { $bySymbol[$_.s].marketState -eq 'REGULAR' } | Select-Object -First 1)
if ($isLive) { $priceAsOf = $priceEastern.ToString('MMM d, h:mm tt', $culture) + ' ET' }
else { $priceAsOf = $priceEastern.ToString('MMM d', $culture) + ' close' }

$output = [ordered]@{
    generatedAt = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
    today = $nowEastern.ToString('yyyy-MM-dd')
    priceAsOf = $priceAsOf
    sectorAsOf = $stored.sectorAsOf
    sectorPE = $stored.sectorPE
    flags = $stored.flags
    dates = $stored.dates
    missing = $missing
    trending = $trending
    stocks = $stocks
}
$json = $output | ConvertTo-Json -Depth 6 -Compress
[IO.File]::WriteAllText($outFile, $json, (New-Object Text.UTF8Encoding($false)))
Write-Output "Wrote $($stocks.Count) stocks. Prices as of $priceAsOf. Missing: $($missing -join ', ')"
