#Requires -Version 5.1
# ─────────────────────────────────────────────────────────────────────────────
# Windows Applicable CVE Checker  v3.0  —  TUI Edition
# Based on MSRC CVRF API (same approach as patch_review.ps1 by Fabian Bader)
# Matching via ProductTree ProductIDs — locale-independent
# No Admin | No Python | No external modules | PS 5.1+
# ─────────────────────────────────────────────────────────────────────────────
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue
[System.Globalization.CultureInfo]::CurrentCulture = [System.Globalization.CultureInfo]::CreateSpecificCulture('en-US')

# ═════════════════════════════════════════════════════════════════════════════
#  TUI ENGINE
# ═════════════════════════════════════════════════════════════════════════════

$TUI = @{
    W           = 78
    AccentColor = 'Cyan'
    BorderColor = 'DarkCyan'
    SelectColor = 'Black'
    SelectBg    = 'Cyan'
    NormalColor = 'White'
    DimColor    = 'DarkGray'
    OkColor     = 'Green'
    WarnColor   = 'Yellow'
    ErrColor    = 'Red'
    ExplColor   = 'Magenta'
}

function TUI-Clear { Clear-Host }

function TUI-Box {
    param([string]$Title = '', [int]$Width = $TUI.W)
    $inner = $Width - 2
    $top   = [char]0x2554 + ([string][char]0x2550 * $inner) + [char]0x2557
    $bot   = [char]0x255A + ([string][char]0x2550 * $inner) + [char]0x255D
    Write-Host $top -ForegroundColor $TUI.BorderColor
    if ($Title) {
        $pad  = [math]::Max(0, $inner - $Title.Length)
        $lpad = [math]::Floor($pad / 2)
        $rpad = $pad - $lpad
        Write-Host ([char]0x2551 + (' ' * $lpad) + $Title + (' ' * $rpad) + [char]0x2551) `
            -ForegroundColor $TUI.AccentColor
        Write-Host ([char]0x2560 + ([string][char]0x2550 * $inner) + [char]0x2563) `
            -ForegroundColor $TUI.BorderColor
    }
    return $bot
}

function TUI-Line {
    param([string]$Text = '', [string]$Color = 'White', [int]$Width = $TUI.W)
    $inner = $Width - 4
    if ($Text.Length -gt $inner) { $Text = $Text.Substring(0, $inner - 1) + '>' }
    $rpad = $inner - $Text.Length
    Write-Host ([char]0x2551 + ' ' + $Text + (' ' * $rpad) + ' ' + [char]0x2551) `
        -ForegroundColor $Color
}

function TUI-Sep {
    param([int]$Width = $TUI.W)
    Write-Host ([char]0x255F + ([string][char]0x2500 * ($Width - 2)) + [char]0x2562) `
        -ForegroundColor $TUI.BorderColor
}

function TUI-Close { param([string]$Bot) Write-Host $Bot -ForegroundColor $TUI.BorderColor }

function TUI-Header {
    TUI-Clear
    $bot = TUI-Box -Title ''
    TUI-Line '  Windows Applicable CVE Checker  v3.0' -Color $TUI.AccentColor
    TUI-Line '  MSRC CVRF API  |  ProductTree matching  |  No Admin  |  PS 5.1+' -Color $TUI.DimColor
    TUI-Line ''
    TUI-Close $bot
    Write-Host ''
}

function TUI-Menu {
    param(
        [string]$Title,
        [string[]]$Items,
        [int]$Default = 0,
        [string]$Hint = 'Arrow UP/DOWN to navigate    Enter to select    Esc to cancel'
    )
    $sel   = $Default
    $count = $Items.Count
    while ($true) {
        TUI-Header
        $bot = TUI-Box -Title " $Title "
        TUI-Line ''
        for ($i = 0; $i -lt $count; $i++) {
            $text = "  $(if ($i -eq $sel) { '>>' } else { '  ' })  $($Items[$i])"
            if ($i -eq $sel) {
                Write-Host -NoNewline ([char]0x2551 + ' ') -ForegroundColor $TUI.BorderColor
                Write-Host -NoNewline $text.PadRight($TUI.W - 4) `
                    -ForegroundColor $TUI.SelectColor -BackgroundColor $TUI.SelectBg
                Write-Host (' ' + [char]0x2551) -ForegroundColor $TUI.BorderColor
            } else {
                TUI-Line $text -Color $TUI.NormalColor
            }
        }
        TUI-Line ''
        TUI-Sep
        TUI-Line "  $Hint" -Color $TUI.DimColor
        TUI-Line ''
        TUI-Close $bot
        $key = [System.Console]::ReadKey($true)
        switch ($key.Key) {
            'UpArrow'   { $sel = ($sel - 1 + $count) % $count }
            'DownArrow' { $sel = ($sel + 1) % $count }
            'Enter'     { return $sel }
            'Escape'    { return -1 }
            'Q'         { return -1 }
        }
    }
}

function TUI-MultiSelect {
    param([string]$Title, [string[]]$Items, [bool[]]$Defaults)
    $sel    = $Defaults | ForEach-Object { $_ }
    $cursor = 0
    $count  = $Items.Count
    while ($true) {
        TUI-Header
        $bot = TUI-Box -Title " $Title "
        TUI-Line ''
        for ($i = 0; $i -lt $count; $i++) {
            $check = if ($sel[$i]) { '[X]' } else { '[ ]' }
            $text  = "  $(if ($i -eq $cursor) { '>>' } else { '  ' })  $check  $($Items[$i])"
            if ($i -eq $cursor) {
                Write-Host -NoNewline ([char]0x2551 + ' ') -ForegroundColor $TUI.BorderColor
                Write-Host -NoNewline $text.PadRight($TUI.W - 4) `
                    -ForegroundColor $TUI.SelectColor -BackgroundColor $TUI.SelectBg
                Write-Host (' ' + [char]0x2551) -ForegroundColor $TUI.BorderColor
            } else {
                $col = if ($sel[$i]) { $TUI.OkColor } else { $TUI.NormalColor }
                TUI-Line $text -Color $col
            }
        }
        TUI-Line ''
        TUI-Sep
        TUI-Line '  Arrow UP/DOWN    Space to toggle    Enter to confirm' -Color $TUI.DimColor
        TUI-Line ''
        TUI-Close $bot
        $key = [System.Console]::ReadKey($true)
        switch ($key.Key) {
            'UpArrow'   { $cursor = ($cursor - 1 + $count) % $count }
            'DownArrow' { $cursor = ($cursor + 1) % $count }
            'Spacebar'  { $sel[$cursor] = -not $sel[$cursor] }
            'Enter'     { return $sel }
            'Escape'    { return $Defaults }
        }
    }
}

function TUI-Input {
    param([string]$Title, [string]$Prompt, [string]$Default = '',
          [string]$Validate = '', [string]$ErrMsg = 'Invalid format')
    while ($true) {
        TUI-Header
        $bot = TUI-Box -Title " $Title "
        TUI-Line ''
        TUI-Line "  $Prompt" -Color $TUI.NormalColor
        TUI-Line ''
        if ($Default) { TUI-Line "  Default: $Default" -Color $TUI.DimColor }
        TUI-Line ''
        TUI-Sep
        TUI-Line '  Press Enter without input to use the default value' -Color $TUI.DimColor
        TUI-Close $bot
        Write-Host ''
        Write-Host '  >> ' -ForegroundColor $TUI.AccentColor -NoNewline
        $val = Read-Host
        if ($val -eq '' -and $Default -ne '') { $val = $Default }
        if ($val -eq '') {
            Write-Host '  [WARN] Value cannot be empty' -ForegroundColor $TUI.WarnColor
            Start-Sleep 1; continue
        }
        if ($Validate -and $val -notmatch $Validate) {
            Write-Host "  [WARN] $ErrMsg" -ForegroundColor $TUI.WarnColor
            Start-Sleep 1; continue
        }
        return $val
    }
}

function TUI-Notify {
    param([string]$Message, [string]$Level = 'Info')
    $color = switch ($Level) { 'Ok' { $TUI.OkColor } 'Warn' { $TUI.WarnColor }
                                'Err' { $TUI.ErrColor } default { $TUI.NormalColor } }
    Write-Host ''
    Write-Host "  [$Level] $Message" -ForegroundColor $color
    Write-Host '  Press any key to continue...' -ForegroundColor $TUI.DimColor
    [void][System.Console]::ReadKey($true)
}

function TUI-Progress {
    param([string]$Msg, [int]$Current, [int]$Total)
    $pct   = if ($Total -gt 0) { [int](($Current / $Total) * 100) } else { 0 }
    $bar   = [int](40 * $pct / 100)
    $empty = 40 - $bar
    Write-Host "`r  [" + ('#' * $bar) + ('.' * $empty) + "] $pct%  $Msg          " `
        -NoNewline -ForegroundColor $TUI.AccentColor
}

# ═════════════════════════════════════════════════════════════════════════════
#  SYSTEM INFORMATION  (locale-independent — build number + registry)
# ═════════════════════════════════════════════════════════════════════════════

function Get-SystemInfo {
    $os = $null
    try   { $os = Get-CimInstance Win32_OperatingSystem -EA Stop }
    catch { try { $os = Get-WmiObject Win32_OperatingSystem -EA Stop } catch {} }
    if (-not $os) {
        Write-Host '  [ERR] Failed to retrieve OS information.' -ForegroundColor Red; exit 1
    }

    $build    = [int]$os.BuildNumber
    $archNorm = if ($os.OSArchitecture -match '64') { 'x64' } else { 'x86' }

    # Derive WinMajor from build number — 100% locale-independent
    $winMajor = switch ($build) {
        { $_ -ge 22000 } { '11'  }
        { $_ -ge 10240 } { '10'  }
        { $_ -ge 9600  } { '8.1' }
        default          { '10'  }
    }

    # Registry keys are always English regardless of UI language
    $releaseId  = ''
    $englishName = ''
    try {
        $rk = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -EA Stop
        $releaseId   = if ($rk.DisplayVersion) { $rk.DisplayVersion } else { "$($rk.ReleaseId)" }
        $englishName = if ($rk.EnglishProductName) { $rk.EnglishProductName } else { "Windows $winMajor" }
    } catch { $englishName = "Windows $winMajor" }
    $englishName = $englishName -replace '^Microsoft\s+', ''

    # IsServer from DomainRole (numeric, locale-independent)
    $isServer = $false
    try {
        $cs = Get-CimInstance Win32_ComputerSystem -EA Stop
        $isServer = ([int]$cs.DomainRole -ge 2)
    } catch {
        try {
            $wcs = Get-WmiObject Win32_ComputerSystem -EA Stop
            $isServer = ([int]$wcs.DomainRole -ge 2)
        } catch {}
    }

    # Arch string as used in MSRC ProductTree names
    # Examples: "x64-based Systems", "32-bit Systems", "ARM64-based Systems"
    $msrcArch = switch ($archNorm) {
        'x64' { 'x64-based Systems'  }
        'x86' { '32-bit Systems'     }
        default { 'x64-based Systems' }
    }

    # Display caption (may be localized — only used for human display)
    $displayCaption = ($os.Caption -replace '^Microsoft\s+', '')

    return [PSCustomObject]@{
        Caption      = $displayCaption
        EnglishName  = $englishName          # always English
        WinMajor     = $winMajor             # "10" or "11"
        ReleaseId    = $releaseId            # "23H2", "22H2", etc.
        Build        = $build
        Arch         = $archNorm             # x64 / x86
        MsrcArch     = $msrcArch             # as in ProductTree
        IsServer     = $isServer
        FullString   = "$displayCaption $releaseId Build $build $archNorm"
    }
}

# ═════════════════════════════════════════════════════════════════════════════
#  INSTALLED KB LIST
# ═════════════════════════════════════════════════════════════════════════════

function Get-InstalledKBs {
    $kbs = @{}

    # Method 1: Get-HotFix
    try {
        Get-HotFix -EA Stop | ForEach-Object {
            $num = $_.HotFixID -replace '^KB', ''
            $kbs[$num] = $true      # store bare number to match MSRC Description.Value
        }
    } catch {}

    # Method 2: wmic qfe
    if ($kbs.Count -eq 0) {
        try {
            & wmic qfe list brief 2>$null | ForEach-Object {
                if ($_ -match 'KB(\d+)') { $kbs[$Matches[1]] = $true }
            }
        } catch {}
    }

    # Method 3: CBS registry
    if ($kbs.Count -eq 0) {
        try {
            $p = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages'
            if (Test-Path $p) {
                Get-ChildItem $p -EA SilentlyContinue | ForEach-Object {
                    if ($_.PSChildName -match 'KB(\d+)') { $kbs[$Matches[1]] = $true }
                }
            }
        } catch {}
    }

    return $kbs   # keys are bare KB numbers: "5031356" not "KB5031356"
}

# ═════════════════════════════════════════════════════════════════════════════
#  MSRC API HELPERS
# ═════════════════════════════════════════════════════════════════════════════

$MSRC_BASE = 'https://api.msrc.microsoft.com/cvrf/v2.0'
$MSRC_HEADERS = @{ Accept = 'application/json'; 'api-version' = '2016-08-01' }

$MonthMap = @{
    '01'='Jan';'02'='Feb';'03'='Mar';'04'='Apr';'05'='May';'06'='Jun'
    '07'='Jul';'08'='Aug';'09'='Sep';'10'='Oct';'11'='Nov';'12'='Dec'
}

function ConvertTo-MsrcId {
    # Converts YYYY-MM to YYYY-Mon  (e.g. 2025-09 -> 2025-Sep)
    param([string]$YearMonth)
    $parts = $YearMonth -split '-'
    return "$($parts[0])-$($MonthMap[$parts[1]])"
}

function Invoke-MsrcApi {
    param([string]$Url)
    try {
        return Invoke-RestMethod -Uri $Url -Method GET -Headers $MSRC_HEADERS `
            -UseBasicParsing -TimeoutSec 60 -EA Stop
    } catch {
        $code = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 0 }
        if ($code -eq 404) { return $null }
        Write-Host "`n  [WARN] API error ($code): $Url" -ForegroundColor Yellow
        return $null
    }
}

# Get list of available monthly release IDs in date range
function Get-ReleasesInRange {
    param([datetime]$Start, [datetime]$End)
    $index = Invoke-MsrcApi "$MSRC_BASE/updates"
    if (-not $index) { return @() }
    return @($index.value | Where-Object {
        try {
            $d = [datetime]::Parse($_.CurrentReleaseDate)
            $d -ge $Start -and $d -le $End
        } catch { $false }
    } | ForEach-Object { $_.ID })
}

# ═════════════════════════════════════════════════════════════════════════════
#  PRODUCT MATCHING  (ProductTree-based — correct MSRC approach)
# ═════════════════════════════════════════════════════════════════════════════

function Get-MatchingProductIDs {
    param($ProductTree, $SysInfo)

    # Build patterns from system info — these match MSRC ProductTree Value strings
    # which are always in English regardless of OS locale.
    #
    # Examples of MSRC FullProductName.Value:
    #   "Windows 11 Version 23H2 for x64-based Systems"
    #   "Windows 10 Version 22H2 for x64-based Systems"
    #   "Windows Server 2022"

    $winPattern  = "Windows $($SysInfo.WinMajor)"   # "Windows 11"
    $archPattern = $SysInfo.MsrcArch                # "x64-based Systems"

    # ReleaseId pattern — match "Version 23H2" or just the bare "23H2"
    $relPattern  = if ($SysInfo.ReleaseId) { $SysInfo.ReleaseId } else { '' }

    $matchedIDs = @{}

    foreach ($product in $ProductTree.FullProductName) {
        $name = $product.Value
        if (-not $name) { continue }

        # Server / Client separation
        $nameHasServer = ($name -match 'Windows Server')
        if ($SysInfo.IsServer  -and -not $nameHasServer) { continue }
        if (-not $SysInfo.IsServer -and $nameHasServer)  { continue }

        # Must match Windows major version
        if ($name -notmatch $winPattern) { continue }

        # Must match architecture
        if ($name -notmatch [regex]::Escape($archPattern)) { continue }

        # ReleaseId match (23H2, 22H2, 24H2 …)
        # If we have a ReleaseId, the product name must contain it
        if ($relPattern -and $name -notmatch [regex]::Escape($relPattern)) { continue }

        $matchedIDs[$product.ProductID] = $name
    }

    return $matchedIDs
}

# ═════════════════════════════════════════════════════════════════════════════
#  CVE PROCESSING FOR ONE CVRF DOCUMENT
# ═════════════════════════════════════════════════════════════════════════════

function Process-CvrfDocument {
    param($Cvrf, $SysInfo, $InstalledKBs, [string]$BaseScore = '8.0', [switch]$Quiet)

    if (-not $Cvrf.Vulnerability) { return @() }

    # Step 1: find ProductIDs that apply to this system
    $matchedProductIDs = Get-MatchingProductIDs -ProductTree $Cvrf.ProductTree -SysInfo $SysInfo

    if ($matchedProductIDs.Count -eq 0 -and -not $Quiet) {
        Write-Host "  [WARN] No matching products found in ProductTree for this system." `
            -ForegroundColor Yellow
        Write-Host "         WinMajor=$($SysInfo.WinMajor) ReleaseId=$($SysInfo.ReleaseId) Arch=$($SysInfo.MsrcArch)" `
            -ForegroundColor DarkGray
    }

    $VulnTypes = @(
        'Elevation of Privilege', 'Security Feature Bypass', 'Remote Code Execution',
        'Information Disclosure', 'Denial of Service', 'Spoofing', 'Edge - Chromium'
    )

    $allForSystem = @()

    foreach ($v in $Cvrf.Vulnerability) {
        if (-not $v.Title -or [string]::IsNullOrWhiteSpace($v.Title.Value)) { continue }

        # ── CVSS ──────────────────────────────────────────────────────────
        $cvssScore = 'n/a'
        if ($v.CVSSScoreSets -and $v.CVSSScoreSets.Count -gt 0) {
            # patch_review.ps1 approach: take first ScoreSet BaseScore
            $s = $v.CVSSScoreSets[0].BaseScore
            if ($s) { $cvssScore = $s }
        }

        # ── Severity / Criticality ─────────────────────────────────────
        $criticality = 'N/A'
        $possibleCrit = @('Critical','Important','Moderate','Low')
        foreach ($t in ($v.Threats | Where-Object { $_.Type -eq 3 })) {
            $desc = ($t.Description.Value -split ';' | Select-Object -Unique) -join ''
            if ($possibleCrit -contains $desc) { $criticality = $desc; break }
        }

        # ── Exploited / PubliclyDisclosed / ExploitationLikely ─────────
        $exploited   = $false
        $publicDisc  = $false
        $explLikely  = $false
        foreach ($t in ($v.Threats | Where-Object { $_.Type -eq 1 })) {
            $desc = $t.Description.Value
            if ($desc -match 'Exploited:Yes')            { $exploited  = $true }
            if ($desc -match 'Publicly Disclosed:Yes')   { $publicDisc = $true }
            if ($desc -match 'Exploitation More Likely') { $explLikely = $true }
        }
        # Also check split fields
        $allThreatDescs = ($v.Threats.Description.Value) -split ';' | Select-Object -Unique
        if ($allThreatDescs -contains 'Publicly Disclosed:Yes') { $publicDisc = $true }

        # ── Customer Action Required ───────────────────────────────────
        $custAction = $false
        $custNote = $v.Notes | Where-Object { $_.Title -eq 'Customer Action Required' } |
                    Select-Object -ExpandProperty Value -First 1
        if ($custNote -eq 'Yes') { $custAction = $true }

        # ── Vuln Type ─────────────────────────────────────────────────
        $vulnType = 'Other'
        foreach ($t in ($v.Threats | Where-Object { $_.Type -eq 0 })) {
            foreach ($vt in $VulnTypes) {
                if ($t.Description.Value -eq $vt -or
                    ($vt -eq 'Edge - Chromium' -and $t.ProductID[0] -eq '11655')) {
                    $vulnType = $vt; break
                }
            }
            if ($vulnType -ne 'Other') { break }
        }

        # ── Applicability check via ProductID ─────────────────────────
        # Check if ANY remediation for this CVE targets one of our matched ProductIDs
        $applicable = $false
        $kbNum      = ''
        $patchStatus = 'UNKNOWN'

        if ($matchedProductIDs.Count -gt 0) {
            foreach ($rem in ($v.Remediations | Where-Object { $_.Type -eq 1 })) {
                # Check if this remediation covers any of our matched ProductIDs
                $remProductIDs = @($rem.ProductID)
                $overlap = $remProductIDs | Where-Object { $matchedProductIDs.ContainsKey($_) }
                if ($overlap.Count -eq 0) { continue }

                $applicable = $true

                # KB number: Description.Value is bare number like "5031356"
                # URL is like "https://catalog.update.microsoft.com/...?q=KB5031356"
                $descVal = if ($rem.Description) { $rem.Description.Value } else { '' }

                # Try Description.Value first (bare number)
                if ($descVal -match '^\d{7,8}$') {
                    $kbNum = $descVal
                }
                # Fallback: extract from URL
                if (-not $kbNum -and $rem.URL -and $rem.URL -match 'KB(\d+)') {
                    $kbNum = $Matches[1]
                }
                # Fallback: extract number from description if it contains KB prefix
                if (-not $kbNum -and $descVal -match 'KB(\d+)') {
                    $kbNum = $Matches[1]
                }
                if ($kbNum) { break }
            }
        } else {
            # No product matching possible — include all Windows-related CVEs
            $applicable = $true
        }

        if (-not $applicable) { continue }

        # ── Patch status ───────────────────────────────────────────────
        if ($kbNum) {
            if ($InstalledKBs.ContainsKey($kbNum)) {
                $patchStatus = 'PATCHED'
            } else {
                $patchStatus = 'MISSING_PATCH'
            }
        }

        # ── HighestRated flag ──────────────────────────────────────────
        $highestRated = $false
        if ($cvssScore -ne 'n/a') {
            if ([float]$cvssScore -ge [float]$BaseScore) { $highestRated = $true }
            if ($criticality -eq 'Critical')             { $highestRated = $true }
        }

        $allForSystem += [PSCustomObject]@{
            CVE              = $v.CVE
            Title            = $v.Title.Value
            CvssScore        = $cvssScore
            Criticality      = $criticality
            VulnType         = $vulnType
            Exploited        = $exploited
            PubliclyDisclosed= $publicDisc
            ExploitationLikely = $explLikely
            CustomerActionRequired = $custAction
            HighestRated     = $highestRated
            KB               = if ($kbNum) { "KB$kbNum" } else { 'N/A' }
            KBNum            = $kbNum
            PatchStatus      = $patchStatus
            URL              = "https://msrc.microsoft.com/update-guide/vulnerability/$($v.CVE)"
        }
    }

    return $allForSystem
}

# ═════════════════════════════════════════════════════════════════════════════
#  FORMAT HELPERS
# ═════════════════════════════════════════════════════════════════════════════

function Format-Score {
    param($Score, [int]$Pad = 4)
    if ($Score -eq 'n/a') { return 'n/a '.PadLeft($Pad) }
    return ("{0:N1}" -f [float]$Score).PadLeft($Pad)
}

function Get-PatchColor {
    param([string]$Status)
    switch ($Status) {
        'PATCHED'       { return $TUI.OkColor   }
        'MISSING_PATCH' { return $TUI.ErrColor  }
        default         { return $TUI.DimColor   }
    }
}

# ═════════════════════════════════════════════════════════════════════════════
#  TUI CONFIGURATION SCREENS
# ═════════════════════════════════════════════════════════════════════════════

function Screen-Period {
    $idx = TUI-Menu -Title 'STEP 1 / 4  |  SELECT PERIOD' `
        -Items @(
            'Specific month  (Patch Tuesday of that month)'
            'Custom date range  From -> To'
            'Last N months'
            'All available records  (slow - many API calls)'
        ) -Default 0
    if ($idx -eq -1) { return $null }

    $now = Get-Date

    switch ($idx) {
        0 {
            $ms = TUI-Input -Title 'PERIOD - MONTH' `
                -Prompt 'Enter month  YYYY-MM  (e.g. 2025-09)' `
                -Default ($now.ToString('yyyy-MM')) `
                -Validate '^\d{4}-\d{2}$' -ErrMsg 'Required format: YYYY-MM'
            $msrcId = ConvertTo-MsrcId $ms
            return [PSCustomObject]@{
                Mode    = 'Single'
                Ids     = @($msrcId)
                Label   = "Patch Tuesday $ms  ($msrcId)"
            }
        }
        1 {
            $fs = TUI-Input -Title 'DATE RANGE - START' `
                -Prompt 'Start date  YYYY-MM-DD' `
                -Default ($now.AddMonths(-3).ToString('yyyy-MM-dd')) `
                -Validate '^\d{4}-\d{2}-\d{2}$' -ErrMsg 'Required: YYYY-MM-DD'
            $ts = TUI-Input -Title 'DATE RANGE - END' `
                -Prompt 'End date  YYYY-MM-DD' `
                -Default ($now.ToString('yyyy-MM-dd')) `
                -Validate '^\d{4}-\d{2}-\d{2}$' -ErrMsg 'Required: YYYY-MM-DD'
            $dtF = [datetime]::ParseExact($fs,'yyyy-MM-dd',$null)
            $dtT = [datetime]::ParseExact($ts,'yyyy-MM-dd',$null)
            if ($dtF -gt $dtT) { $tmp=$dtF; $dtF=$dtT; $dtT=$tmp }
            return [PSCustomObject]@{
                Mode    = 'Range'
                Start   = $dtF
                End     = $dtT.AddDays(1).AddSeconds(-1)
                Label   = "$fs -> $ts"
                Ids     = @()
            }
        }
        2 {
            $ns = TUI-Input -Title 'LAST N MONTHS' `
                -Prompt 'Number of months  (1-60)' `
                -Default '3' -Validate '^\d{1,2}$' -ErrMsg 'Enter a number'
            $n = [int]$ns
            return [PSCustomObject]@{
                Mode    = 'Range'
                Start   = $now.AddMonths(-$n)
                End     = $now
                Label   = "Last $n month(s)"
                Ids     = @()
            }
        }
        3 {
            return [PSCustomObject]@{
                Mode    = 'Range'
                Start   = [datetime]'2016-01-01'
                End     = $now
                Label   = 'All available records'
                Ids     = @()
            }
        }
    }
}

function Screen-Severity {
    $items = @('All severity levels','Critical','Important','Moderate','Low')
    $idx = TUI-Menu -Title 'STEP 2 / 4  |  SEVERITY FILTER' -Items $items -Default 0
    return if ($idx -le 0) { '' } else { $items[$idx] }
}

function Screen-StatusFilter {
    $statItems = @('MISSING_PATCH','UNKNOWN','PATCHED')
    $statDefs  = @($true, $true, $true)
    $statSel   = TUI-MultiSelect -Title 'STEP 3a/ 4  |  PATCH STATUS FILTER' `
        -Items $statItems -Defaults $statDefs

    $statuses = @()
    for ($i = 0; $i -lt $statItems.Count; $i++) {
        if ($statSel[$i]) { $statuses += $statItems[$i] }
    }

    $exIdx = TUI-Menu -Title 'STEP 3b/ 4  |  EXPLOITED IN THE WILD' `
        -Items @(
            'All CVEs'
            'Only CVEs exploited in the wild'
        ) -Default 0

    return [PSCustomObject]@{
        Statuses      = $statuses
        OnlyExploited = ($exIdx -eq 1)
    }
}

function Screen-Output {
    $idx = TUI-Menu -Title 'STEP 4 / 4  |  OUTPUT FORMAT' `
        -Items @('Table  - color-coded console output','JSON   - machine-readable') -Default 0
    return if ($idx -eq 1) { 'Json' } else { 'Table' }
}

# ═════════════════════════════════════════════════════════════════════════════
#  TUI TABLE VIEW  (interactive scrollable table)
# ═════════════════════════════════════════════════════════════════════════════

function Show-TuiTable {
    param($CVEs, $SysInfo, $Period)

    if ($CVEs.Count -eq 0) {
        TUI-Notify 'No CVEs to display in table view.' -Level Warn
        return 'None'
    }

    # Sort: MISSING_PATCH first, then UNKNOWN, then PATCHED; within each by CVSS desc
    $rows = @($CVEs | Sort-Object @{
        E = { switch ($_.PatchStatus) {
                'MISSING_PATCH' { 0 } 'UNKNOWN' { 1 } 'PATCHED' { 2 } default { 3 }
              }
        }
    }, @{ E = { if ($_.CvssScore -eq 'n/a') { 0.0 } else { [float]$_.CvssScore } }; D = $true })

    $total     = $rows.Count
    $scrollPos = 0
    $result    = 'None'

    # Summary stats — computed once, do not change during scrolling
    $cntMissing = ($CVEs | Where-Object PatchStatus -eq 'MISSING_PATCH').Count
    $cntPatched = ($CVEs | Where-Object PatchStatus -eq 'PATCHED').Count
    $cntUnknown = ($CVEs | Where-Object PatchStatus -eq 'UNKNOWN').Count
    $cntExpl    = ($CVEs | Where-Object Exploited).Count

    # VulnType short labels — fixed 8 chars to keep column aligned
    function Get-VTLabel { param([string]$vt)
        switch ($vt) {
            'Elevation of Privilege'  { 'EoP     ' }
            'Remote Code Execution'   { 'RCE     ' }
            'Security Feature Bypass' { 'SFB     ' }
            'Information Disclosure'  { 'InfoDisc' }
            'Denial of Service'       { 'DoS     ' }
            'Spoofing'                { 'Spoofing' }
            'Edge - Chromium'         { 'Chromium' }
            default                   { 'Other   ' }
        }
    }

    # ── Inline TUI helpers that accept a runtime Width ────────────────────────
    # These shadow the global helpers inside this function only,
    # so every draw call uses the current terminal width.

    function FT-Box {
        param([string]$Title = '', [int]$W)
        $inner = $W - 2
        $top   = [char]0x2554 + ([string][char]0x2550 * $inner) + [char]0x2557
        $bot   = [char]0x255A + ([string][char]0x2550 * $inner) + [char]0x255D
        Write-Host $top -ForegroundColor $TUI.BorderColor
        if ($Title) {
            $pad  = [math]::Max(0, $inner - $Title.Length)
            $lpad = [math]::Floor($pad / 2)
            $rpad = $pad - $lpad
            Write-Host ([char]0x2551 + (' ' * $lpad) + $Title + (' ' * $rpad) + [char]0x2551) `
                -ForegroundColor $TUI.AccentColor
            Write-Host ([char]0x2560 + ([string][char]0x2550 * $inner) + [char]0x2563) `
                -ForegroundColor $TUI.BorderColor
        }
        return $bot
    }

    function FT-Line {
        param([string]$Text = '', [string]$Color = 'White', [int]$W)
        $inner = $W - 4
        if ($Text.Length -gt $inner) { $Text = $Text.Substring(0, $inner - 1) + '>' }
        $rpad  = $inner - $Text.Length
        Write-Host ([char]0x2551 + ' ' + $Text + (' ' * $rpad) + ' ' + [char]0x2551) `
            -ForegroundColor $Color
    }

    function FT-Sep {
        param([int]$W)
        Write-Host ([char]0x255F + ([string][char]0x2500 * ($W - 2)) + [char]0x2562) `
            -ForegroundColor $TUI.BorderColor
    }

    # ─────────────────────────────────────────────────────────────────────────
    while ($true) {

        # ── Read terminal size on every frame (handles window resize) ─────────
        $W = [math]::Max(80, [System.Console]::WindowWidth  - 1)
        $H = [math]::Max(24, [System.Console]::WindowHeight)

        # Fixed overhead line count:
        #   top border(1) + title+sep(2) + sys(1) + period(1) + stats(1) +
        #   TUI-Sep(1) + hdr1(1) + hdr2(1) + TUI-Sep(1) +
        #   [data lines] +
        #   TUI-Sep(1) + pagination(1) + TUI-Sep(1) + hint(1) + empty(1) + bottom(1)
        # Fixed = 16 lines; remainder for data (2 lines per entry)
        $fixedLines = 16
        $pageSize   = [math]::Max(2, [math]::Floor(($H - $fixedLines) / 2))

        # Title column: inner usable = W-4, prefix "  " = 2,
        # line2 prefix "     [VTLabel(8)]  " = 5+1+8+1+2 = 17 → titleMax = W-4-2-17 = W-23
        $titleMax = [math]::Max(10, $W - 23)

        TUI-Clear
        $bot = FT-Box -Title ' INTERACTIVE TABLE VIEW ' -W $W

        # Info header
        $sysStr = $SysInfo.FullString
        $sysMax = $W - 16
        if ($sysStr.Length -gt $sysMax) { $sysStr = $sysStr.Substring(0, $sysMax - 3) + '...' }
        FT-Line "  System : $sysStr"  -Color $TUI.NormalColor -W $W
        FT-Line "  Period : $($Period.Label)" -Color $TUI.NormalColor -W $W
        FT-Line "  Total: $total   Missing: $cntMissing   Patched: $cntPatched   Unknown: $cntUnknown   Exploited: $cntExpl" `
            -Color $TUI.NormalColor -W $W
        FT-Sep -W $W

        # Column headers (line 1 + line 2 mirror the data rows)
        $colHdr1 = '{0,-3}  {1,-16}  {2,-4}  {3,-9}  {4,-13}  {5,-11}  {6}' -f `
            ' # ', 'CVE ID', 'CVSS', 'Criticality', 'PatchStatus', 'KB', 'Flg'
        $colHdr2 = '     [{0,-8}]  {1}' -f 'VulnType', 'Title'
        FT-Line "  $colHdr1" -Color $TUI.DimColor -W $W
        FT-Line "  $colHdr2" -Color $TUI.DimColor -W $W
        FT-Sep -W $W

        # ── Data rows ─────────────────────────────────────────────────────────
        $endRow = [math]::Min($scrollPos + $pageSize - 1, $total - 1)

        for ($i = $scrollPos; $i -le $endRow; $i++) {
            $c = $rows[$i]

            # Color scheme:
            #   Exploited      → Magenta  (highest priority)
            #   MISSING_PATCH  → Red
            #   PATCHED        → Green
            #   UNKNOWN        → Yellow   (default — all plain CVEs)
            $rowColor = switch ($c.PatchStatus) {
                'MISSING_PATCH' { $TUI.ErrColor  }
                'PATCHED'       { $TUI.OkColor   }
                default         { $TUI.WarnColor  }   # Yellow
            }
            if ($c.Exploited) { $rowColor = $TUI.ExplColor }

            # Flags: E=Exploited  L=ExploitationLikely  P=PubliclyDisclosed
            $flags  = ''
            $flags += if ($c.Exploited)          { 'E' } else { ' ' }
            $flags += if ($c.ExploitationLikely) { 'L' } else { ' ' }
            $flags += if ($c.PubliclyDisclosed)  { 'P' } else { ' ' }

            # Field truncation
            $cveId  = if ($c.CVE.Length         -gt 16) { $c.CVE.Substring(0,15)         + '>' } else { $c.CVE         }
            $crit   = if ($c.Criticality.Length -gt  9) { $c.Criticality.Substring(0, 8) + '>' } else { $c.Criticality }
            $kb     = if ($c.KB.Length          -gt 11) { $c.KB.Substring(0,10)          + '>' } else { $c.KB          }
            $status = if ($c.PatchStatus.Length -gt 13) { $c.PatchStatus.Substring(0,12) + '>' } else { $c.PatchStatus }
            $score  = Format-Score $c.CvssScore 4
            $vtLbl  = Get-VTLabel $c.VulnType
            $title  = if ($c.Title.Length -gt $titleMax) { $c.Title.Substring(0, $titleMax - 1) + '>' } else { $c.Title }

            $line1 = '{0,3}  {1,-16}  {2}  {3,-9}  {4,-13}  {5,-11}  {6}' -f `
                ($i + 1), $cveId, $score, $crit, $status, $kb, $flags
            $line2 = '     [{0}]  {1}' -f $vtLbl, $title

            FT-Line "  $line1" -Color $rowColor -W $W
            FT-Line "  $line2" -Color $rowColor -W $W
        }

        # Pad unused slots to hold layout stable
        $shown = $endRow - $scrollPos + 1
        for ($i = $shown; $i -lt $pageSize; $i++) {
            FT-Line '' -W $W
            FT-Line '' -W $W
        }

        FT-Sep -W $W

        # Pagination + legend
        $page  = [math]::Floor($scrollPos / $pageSize) + 1
        $pages = [math]::Max(1, [math]::Ceiling($total / $pageSize))
        FT-Line "  Entries $($scrollPos+1)-$($endRow+1) of $total   Page $page/$pages   E=Exploited  L=ExplLikely  P=Disclosed" `
            -Color $TUI.DimColor -W $W
        FT-Sep -W $W
        FT-Line '  UP/DOWN scroll    PgUp/PgDn page    Home/End    [E] Export JSON    [Q] Back' `
            -Color $TUI.AccentColor -W $W
        FT-Line '' -W $W
        Write-Host $bot -ForegroundColor $TUI.BorderColor

        # ── Input ─────────────────────────────────────────────────────────────
        $key        = [System.Console]::ReadKey($true)
        $shouldExit = $false

        switch ($key.Key) {
            'UpArrow'   { if ($scrollPos -gt 0) { $scrollPos-- } }
            'DownArrow' {
                $maxScroll = [math]::Max(0, $total - $pageSize)
                if ($scrollPos -lt $maxScroll) { $scrollPos++ }
            }
            'PageUp'    { $scrollPos = [math]::Max(0, $scrollPos - $pageSize) }
            'PageDown'  {
                $maxScroll = [math]::Max(0, $total - $pageSize)
                $scrollPos = [math]::Min($maxScroll, $scrollPos + $pageSize)
            }
            'Home'      { $scrollPos = 0 }
            'End'       { $scrollPos = [math]::Max(0, $total - $pageSize) }
            'E'         { $result = 'Export'; $shouldExit = $true }
            'Q'         { $shouldExit = $true }
            'Escape'    { $shouldExit = $true }
        }

        if ($shouldExit) { break }
    }

    return $result
}

# ═════════════════════════════════════════════════════════════════════════════
#  RESULTS OUTPUT
# ═════════════════════════════════════════════════════════════════════════════

function Show-Table {
    param($CVEs, $SysInfo, $Period, $BaseScoreThreshold)

    TUI-Clear
    $VulnTypes = @(
        'Elevation of Privilege','Security Feature Bypass','Remote Code Execution',
        'Information Disclosure','Denial of Service','Spoofing','Edge - Chromium','Other'
    )

    Write-Host ''
    Write-Host '[+] Windows CVE Checker v3.0  |  github.com/f-bader/msrc-api-ps (API source)' `
        -ForegroundColor $TUI.OkColor
    Write-Host "[+] System  : $($SysInfo.FullString)"  -ForegroundColor $TUI.OkColor
    Write-Host "[+] Period  : $($Period.Label)"         -ForegroundColor $TUI.OkColor
    Write-Host "[+] Found $($CVEs.Count) applicable vulnerabilities" -ForegroundColor $TUI.OkColor
    Write-Host ''

    # Category counts
    foreach ($vt in $VulnTypes) {
        $cnt = ($CVEs | Where-Object { $_.VulnType -eq $vt }).Count
        if ($cnt -gt 0) {
            Write-Host "  [-] $cnt $vt Vulnerabilities" -ForegroundColor $TUI.AccentColor
        }
    }
    Write-Host ''

    # Patch status summary
    $missing = ($CVEs | Where-Object PatchStatus -eq 'MISSING_PATCH').Count
    $patched = ($CVEs | Where-Object PatchStatus -eq 'PATCHED').Count
    $unknown = ($CVEs | Where-Object PatchStatus -eq 'UNKNOWN').Count
    Write-Host "[+] Patch status on this system:" -ForegroundColor $TUI.OkColor
    Write-Host "  [!] MISSING_PATCH : $missing" -ForegroundColor $TUI.ErrColor
    Write-Host "  [+] PATCHED       : $patched" -ForegroundColor $TUI.OkColor
    Write-Host "  [?] UNKNOWN       : $unknown" -ForegroundColor $TUI.DimColor
    Write-Host ''

    # Column widths
    $maxCVE   = [math]::Max(14, ($CVEs.CVE | Measure-Object -Property Length -Maximum).Maximum)
    $maxScore = if (($CVEs | Where-Object { $_.CvssScore -ne 'n/a' -and [float]$_.CvssScore -ge 10 }).Count -gt 0) { 4 } else { 3 }

    # Exploited in the wild
    $exploited = @($CVEs | Where-Object Exploited | Sort-Object @{E='CvssScore';D=$true})
    Write-Host "[+] Found $($exploited.Count) exploited in the wild" -ForegroundColor $TUI.OkColor
    foreach ($c in $exploited) {
        $sc = Format-Score $c.CvssScore $maxScore
        $pc = Get-PatchColor $c.PatchStatus
        Write-Host "  [-] $($c.CVE.PadRight($maxCVE)) $sc  $($c.Title)" `
            -ForegroundColor $TUI.ErrColor -NoNewline
        Write-Host "  [$($c.PatchStatus)]  $($c.KB)" -ForegroundColor $pc
    }
    Write-Host ''

    # Publicly disclosed
    $pubDisc = @($CVEs | Where-Object PubliclyDisclosed | Sort-Object @{E='CvssScore';D=$true})
    Write-Host "[+] Found $($pubDisc.Count) publicly disclosed vulnerabilities" `
        -ForegroundColor $TUI.OkColor
    foreach ($c in $pubDisc) {
        $sc = Format-Score $c.CvssScore $maxScore
        $pc = Get-PatchColor $c.PatchStatus
        Write-Host "  [-] $($c.CVE.PadRight($maxCVE)) $sc  $($c.Title)" `
            -ForegroundColor $TUI.ErrColor -NoNewline
        Write-Host "  [$($c.PatchStatus)]  $($c.KB)" -ForegroundColor $pc
    }
    Write-Host ''

    # Highest rated
    $highRated = @($CVEs | Where-Object HighestRated | Sort-Object @{E='CvssScore';D=$true})
    Write-Host "[+] Highest Rated Vulnerabilities CVE >= $BaseScoreThreshold" `
        -ForegroundColor $TUI.OkColor
    foreach ($c in $highRated) {
        $sc = Format-Score $c.CvssScore $maxScore
        $pc = Get-PatchColor $c.PatchStatus
        Write-Host "  [-] $($c.CVE.PadRight($maxCVE)) $sc  $($c.Title)" `
            -ForegroundColor $TUI.WarnColor -NoNewline
        if (-not $c.CustomerActionRequired) {
            Write-Host "  [FIXED]" -ForegroundColor $TUI.OkColor -NoNewline
        }
        Write-Host "  [$($c.PatchStatus)]  $($c.KB)" -ForegroundColor $pc
    }
    Write-Host ''

    # Exploitation likely
    $explLikely = @($CVEs | Where-Object ExploitationLikely | Sort-Object @{E='CvssScore';D=$true})
    Write-Host "[+] Found $($explLikely.Count) vulnerabilities more likely to be exploited" `
        -ForegroundColor $TUI.OkColor
    foreach ($c in $explLikely) {
        $sc = Format-Score $c.CvssScore $maxScore
        $pc = Get-PatchColor $c.PatchStatus
        Write-Host "  [-] $($c.CVE.PadRight($maxCVE)) $sc  $($c.Title)  $($c.URL)" `
            -ForegroundColor $TUI.WarnColor -NoNewline
        Write-Host "  [$($c.PatchStatus)]  $($c.KB)" -ForegroundColor $pc
    }
    Write-Host ''
}

function Show-Json {
    param($CVEs, $SysInfo, $Period)
    [PSCustomObject]@{
        meta = [PSCustomObject]@{
            system    = $SysInfo.FullString
            period    = $Period.Label
            generated = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
            total     = $CVEs.Count
        }
        cves = $CVEs | Select-Object CVE, Title, CvssScore, Criticality, VulnType,
                                     Exploited, PubliclyDisclosed, ExploitationLikely,
                                     CustomerActionRequired, KB, PatchStatus, URL
    } | ConvertTo-Json -Depth 5
}

# ═════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═════════════════════════════════════════════════════════════════════════════

function Main {
    $BaseScoreThreshold = 8.0

    # ── System detection ─────────────────────────────────────────────────────
    TUI-Clear
    $bot = TUI-Box -Title ' INITIALIZING '
    TUI-Line ''
    TUI-Line '  Detecting system parameters...' -Color $TUI.DimColor
    TUI-Line ''
    TUI-Close $bot
    Write-Host ''

    $sysInfo = Get-SystemInfo
    $kbs     = Get-InstalledKBs

    TUI-Header
    $bot = TUI-Box -Title ' SYSTEM DETECTED '
    TUI-Line ''
    TUI-Line "  $($sysInfo.FullString)"               -Color $TUI.OkColor
    TUI-Line "  English product name : $($sysInfo.EnglishName)"  -Color $TUI.NormalColor
    TUI-Line "  Windows major        : $($sysInfo.WinMajor)"      -Color $TUI.NormalColor
    TUI-Line "  Release ID           : $($sysInfo.ReleaseId)"     -Color $TUI.NormalColor
    TUI-Line "  MSRC arch string     : $($sysInfo.MsrcArch)"      -Color $TUI.NormalColor
    TUI-Line "  Is Server            : $($sysInfo.IsServer)"      -Color $TUI.NormalColor
    TUI-Line "  Installed KBs found  : $($kbs.Count)"            -Color $TUI.NormalColor
    TUI-Line ''
    TUI-Sep
    TUI-Line '  Press any key to configure analysis...' -Color $TUI.DimColor
    TUI-Line ''
    TUI-Close $bot
    [void][System.Console]::ReadKey($true)

    # ── Configuration ─────────────────────────────────────────────────────────
    $period    = Screen-Period
    if (-not $period) { return }

    $severity  = Screen-Severity
    $statusCfg = Screen-StatusFilter
    $outputFmt = Screen-Output

    # ── Resolve release IDs ───────────────────────────────────────────────────
    TUI-Clear
    $bot = TUI-Box -Title ' LOADING MSRC DATA '
    TUI-Line ''
    TUI-Line "  System : $($sysInfo.FullString)"  -Color $TUI.NormalColor
    TUI-Line "  Period : $($period.Label)"          -Color $TUI.NormalColor
    TUI-Line ''
    TUI-Sep
    TUI-Line '  Connecting to api.msrc.microsoft.com (HTTPS)...' -Color $TUI.DimColor
    TUI-Line ''
    TUI-Close $bot
    Write-Host ''

    $releaseIds = @()
    if ($period.Mode -eq 'Single') {
        $releaseIds = $period.Ids
    } else {
        Write-Host '  Fetching release index...' -ForegroundColor $TUI.DimColor
        $releaseIds = @(Get-ReleasesInRange -Start $period.Start -End $period.End)
        Write-Host "  Found $($releaseIds.Count) monthly release(s) in period." `
            -ForegroundColor $TUI.OkColor
    }

    if ($releaseIds.Count -eq 0) {
        TUI-Notify 'No MSRC releases found for the selected period.' -Level Warn
        return
    }

    # ── Download and process each release ────────────────────────────────────
    $allCVEs = @()
    $total   = $releaseIds.Count
    $idx     = 0
    Write-Host ''

    foreach ($rid in $releaseIds) {
        $idx++
        $sLine = "  >> [$idx/$total]  $($rid.PadRight(9))  |  Downloading..."
        Write-Host "`r$($sLine.PadRight(78))" -NoNewline -ForegroundColor $TUI.AccentColor

        $cvrf = Invoke-MsrcApi "$MSRC_BASE/cvrf/$rid"
        if (-not $cvrf) {
            $sLine = "  >> [$idx/$total]  $($rid.PadRight(9))  |  No data returned"
            Write-Host "`r$($sLine.PadRight(78))" -NoNewline -ForegroundColor $TUI.WarnColor
            Start-Sleep -Milliseconds 300
            continue
        }
        $results = @(Process-CvrfDocument -Cvrf $cvrf -SysInfo $sysInfo `
            -InstalledKBs $kbs -BaseScore $BaseScoreThreshold -Quiet)
        $allCVEs += $results
        $sColor = if ($results.Count -gt 0) { $TUI.OkColor } else { $TUI.DimColor }
        $sLine  = "  >> [$idx/$total]  $($rid.PadRight(9))  |  Found: $($results.Count)  |  Total: $($allCVEs.Count)"
        Write-Host "`r$($sLine.PadRight(78))" -NoNewline -ForegroundColor $sColor
        Start-Sleep -Milliseconds 200
    }
    Write-Host ''
    Write-Host "  Scan complete. Applicable CVEs: $($allCVEs.Count)" -ForegroundColor $TUI.OkColor

    # ── Apply filters ─────────────────────────────────────────────────────────
    if ($severity) {
        $allCVEs = @($allCVEs | Where-Object { $_.Criticality -eq $severity })
    }
    if ($statusCfg.OnlyExploited) {
        $allCVEs = @($allCVEs | Where-Object Exploited)
    }
    if ($statusCfg.Statuses.Count -gt 0 -and $statusCfg.Statuses.Count -lt 3) {
        $allCVEs = @($allCVEs | Where-Object { $statusCfg.Statuses -contains $_.PatchStatus })
    }

    # ── Render ────────────────────────────────────────────────────────────────
    if ($outputFmt -eq 'Json') {
        Show-Json -CVEs $allCVEs -SysInfo $sysInfo -Period $period
    } else {
        Show-Table -CVEs $allCVEs -SysInfo $sysInfo -Period $period `
            -BaseScoreThreshold $BaseScoreThreshold
    }

    # ── Pause then open interactive TUI table ─────────────────────────────────
    Write-Host ''
    Write-Host '  Press any key to open interactive table view...' `
        -ForegroundColor $TUI.DimColor
    [void][System.Console]::ReadKey($true)

    $exportAction = Show-TuiTable -CVEs $allCVEs -SysInfo $sysInfo -Period $period

    if ($exportAction -eq 'Export') {
        $fname = "cve_report_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
        try {
            Show-Json -CVEs $allCVEs -SysInfo $sysInfo -Period $period |
                Out-File -FilePath $fname -Encoding UTF8 -Force
            TUI-Notify "Saved: $(Resolve-Path $fname)" -Level Ok
        } catch { TUI-Notify "Write failed: $_" -Level Err }
    }

    # ── Post menu ─────────────────────────────────────────────────────────────
    $postIdx = TUI-Menu -Title 'WHAT NEXT?' `
        -Items @('Save report to JSON file','Run new analysis','Exit') `
        -Default 2 -Hint 'Arrow UP/DOWN    Enter to select'

    switch ($postIdx) {
        0 {
            $fname = "cve_report_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
            try {
                Show-Json -CVEs $allCVEs -SysInfo $sysInfo -Period $period |
                    Out-File -FilePath $fname -Encoding UTF8 -Force
                TUI-Notify "Saved: $(Resolve-Path $fname)" -Level Ok
            } catch { TUI-Notify "Write failed: $_" -Level Err }
            Main
        }
        1 { Main }
        default {
            TUI-Clear
            Write-Host ''
            Write-Host '  Done. Goodbye.' -ForegroundColor $TUI.DimColor
            Write-Host ''
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
Main
