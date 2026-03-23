# Windows Applicable CVE Checker v3.0

A fully interactive PowerShell TUI tool that queries the **Microsoft Security Response Center (MSRC) CVRF API** and reports which published CVEs apply to the current Windows system, along with their patch status based on installed KB updates.

---

## Key Features

- **No admin rights required** — runs entirely in user context
- **No Python, no external modules** — pure PowerShell 5.1+
- **Locale-independent system detection** — works correctly on any language Windows installation (Russian, Chinese, French, etc.)
- **ProductTree-based CVE matching** — uses the same MSRC API approach as [patch_review.ps1 by Fabian Bader](https://github.com/f-bader/msrc-api-ps)
- **Interactive full-screen TUI table** — scrollable, resizes to terminal window
- **JSON export** — save scan results directly from the table view
- **Color-coded output** — instant visual triage of patch status

---

## Requirements

| Requirement | Minimum |
|---|---|
| PowerShell | 5.1 (built into Windows 10/11) |
| OS | Windows 10 / Windows 11 / Windows Server 2016+ |
| Network | HTTPS access to `api.msrc.microsoft.com` |
| Privileges | Standard user (no admin) |

---

## Quick Start

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\vulncheck.ps1
```

No parameters needed. All configuration is done interactively through the TUI menus.

---

## How It Works

### 1. System Detection (locale-independent)

The script never relies on the localized Windows caption string (which may be in any language). Instead it uses:

| Property | Source |
|---|---|
| Windows major version | Build number (`>= 22000` → Win11, `>= 10240` → Win10) |
| Release ID (23H2, 22H2…) | `HKLM:\...\CurrentVersion\DisplayVersion` |
| English product name | `HKLM:\...\CurrentVersion\EnglishProductName` (always EN) |
| Server / Client | `Win32_ComputerSystem.DomainRole` (numeric, locale-independent) |
| Architecture | `Win32_OperatingSystem.OSArchitecture` |

### 2. Installed KB Detection

Three fallback methods are tried in order:

1. `Get-HotFix` (fastest, most reliable)
2. `wmic qfe list brief`
3. CBS registry: `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages`

KB numbers are stored as bare digits (e.g. `5031356`) to match MSRC `Remediations[].Description.Value` format exactly.

### 3. MSRC API Matching

For each CVRF document (monthly Patch Tuesday release):

1. **ProductTree scan** — finds all `FullProductName` entries matching the system (Windows version, release ID, architecture, client/server)
2. **ProductID extraction** — collects the matching ProductIDs (e.g. `"11926"` for "Windows 11 Version 23H2 for x64-based Systems")
3. **Remediation overlap** — for each CVE, checks whether any remediation entry targets one of the matched ProductIDs
4. **KB extraction** — reads the bare KB number from `Remediation.Description.Value`
5. **Patch status** — compares against installed KB list

### 4. Patch Status Values

| Status | Meaning |
|---|---|
| `MISSING_PATCH` | Fix KB identified, not installed on this system |
| `PATCHED` | Fix KB identified, installed on this system |
| `UNKNOWN` | No KB information in the MSRC data for this CVE |

---

## TUI Navigation — 4 Configuration Steps

All input is keyboard-driven. No command-line parameters.

### Navigation Keys

| Key | Action |
|---|---|
| `↑` / `↓` | Move cursor |
| `Enter` | Confirm selection |
| `Space` | Toggle item on/off (multi-select screens) |
| `Esc` / `Q` | Cancel / go back |

---

### Step 1 — Analysis Period

```
  STEP 1 / 4  |  SELECT PERIOD
  ┌──────────────────────────────────────────┐
  │  >>  Specific month  (Patch Tuesday)     │
  │      Custom date range  From -> To       │
  │      Last N months                       │
  │      All available records               │
  └──────────────────────────────────────────┘
```

| Option | Description |
|---|---|
| Specific month | Downloads the single Patch Tuesday release for `YYYY-MM` |
| Custom range | Arbitrary `YYYY-MM-DD` start and end dates |
| Last N months | Rolling window (default: 3) |
| All available | All MSRC releases since 2016 — **slow**, many API calls |

Text input fields accept Enter to use the pre-filled default value.

---

### Step 2 — Severity Filter

Filter by CVSS-based severity: `All`, `Critical`, `Important`, `Moderate`, or `Low`.

---

### Step 3a — Patch Status Filter

Multi-select (Space to toggle). Controls which patch statuses appear in the final output:

- `[X] MISSING_PATCH`
- `[X] UNKNOWN`
- `[X] PATCHED`

---

### Step 3b — Exploited Filter

Choose between all CVEs or only those marked **Exploited In The Wild** by Microsoft.

---

### Step 4 — Output Format

| Option | Description |
|---|---|
| Table | Color-coded console report in patch_review.ps1 style |
| JSON | Raw machine-readable output (also exportable from the table view) |

---

## Scan Progress Display

During the API download phase, a single line updates in place:

```
  >> [3/12]  2025-Sep  |  Found: 47  |  Total: 183
```

The line overwrites itself — only the release ID, found count, and total change. No scrolling noise.

If a release returns no data:
```
  >> [4/12]  2024-Dec  |  No data returned
```

---

## Console Report Output (Table mode)

After scanning, a patch_review-style report is printed:

```
[+] Windows CVE Checker v3.0
[+] System  : Windows 11 Home 23H2 Build 22631 x64
[+] Period  : Last 3 month(s)
[+] Found 183 applicable vulnerabilities

  [-] 44 Elevation of Privilege Vulnerabilities
  [-] 24 Remote Code Execution Vulnerabilities
  [-] 16 Information Disclosure Vulnerabilities
  ...

[+] Patch status on this system:
  [!] MISSING_PATCH : 31
  [+] PATCHED       : 142
  [?] UNKNOWN       : 10

[+] Found 2 exploited in the wild
  [-] CVE-2025-54110   8.8  Windows Kernel Elevation of Privilege  [MISSING_PATCH]  KB5058481
  [-] CVE-2025-54918   8.8  Windows NTLM Elevation of Privilege    [PATCHED]        KB5058480

[+] Highest Rated Vulnerabilities CVE >= 8.0
  [-] CVE-2025-55234  10.0  Azure Networking EoP  [FIXED]  [MISSING_PATCH]  KB5058490
  ...

[+] Found 9 vulnerabilities more likely to be exploited
  [-] CVE-2025-54110  8.8  Windows Kernel EoP  https://msrc.microsoft.com/...  [MISSING_PATCH]
```

**`[FIXED]`** — Microsoft marked the vulnerability as not requiring customer action (auto-patched by cloud provider or service update).

---

## Interactive Table View

After the console report, the full-screen interactive table opens automatically.

```
╔══════════════════════════════════════════════════════════════════════════╗
║                     INTERACTIVE TABLE VIEW                               ║
╠══════════════════════════════════════════════════════════════════════════╣
║  System : Windows 11 Home 23H2 Build 22631 x64                          ║
║  Period : Last 3 month(s)                                                ║
║  Total: 183   Missing: 31   Patched: 142   Unknown: 10   Exploited: 2   ║
╟──────────────────────────────────────────────────────────────────────────╢
║    #   CVE ID            CVSS  Criticality  PatchStatus    KB      Flg   ║
║        [VulnType]  Title                                                  ║
╟──────────────────────────────────────────────────────────────────────────╢
║    1   CVE-2025-54110   8.8  Critical     MISSING_PATCH  KB5058481  EL  ║
║        [EoP     ]  Windows Kernel Elevation of Privilege Vulnerability   ║
║    2   CVE-2025-54918   8.8  Important    PATCHED        KB5058480  E   ║
║        [EoP     ]  Windows NTLM Elevation of Privilege Vulnerability     ║
║  ...                                                                      ║
╟──────────────────────────────────────────────────────────────────────────╢
║  Entries 1-8 of 183   Page 1/23   E=Exploited  L=ExplLikely  P=Discl.  ║
╟──────────────────────────────────────────────────────────────────────────╢
║  UP/DOWN scroll    PgUp/PgDn page    Home/End    [E] Export JSON  [Q] Back║
╚══════════════════════════════════════════════════════════════════════════╝
```

### Table Columns

| Column | Description |
|---|---|
| `#` | Row index |
| `CVE ID` | CVE identifier |
| `CVSS` | Base score (CVSS v3 preferred, v2 fallback) |
| `Criticality` | Microsoft severity: Critical / Important / Moderate / Low |
| `PatchStatus` | `MISSING_PATCH` / `PATCHED` / `UNKNOWN` |
| `KB` | Associated KB update number |
| `Flg` | `E` Exploited · `L` Exploitation Likely · `P` Publicly Disclosed |
| `[VulnType]` | Abbreviated vulnerability category (line 2) |
| `Title` | Full CVE title, truncated to available terminal width (line 2) |

### VulnType Abbreviations

| Label | Full Name |
|---|---|
| `EoP` | Elevation of Privilege |
| `RCE` | Remote Code Execution |
| `SFB` | Security Feature Bypass |
| `InfoDisc` | Information Disclosure |
| `DoS` | Denial of Service |
| `Spoofing` | Spoofing |
| `Chromium` | Edge — Chromium |
| `Other` | All other types |

### Table Color Scheme

| Color | Meaning |
|---|---|
| 🔴 Red | `MISSING_PATCH` — update not installed |
| 🟢 Green | `PATCHED` — update installed |
| 🟡 Yellow | `UNKNOWN` — no KB data in MSRC |
| 🟣 Magenta | Exploited In The Wild (overrides other colors) |

### Table Navigation Keys

| Key | Action |
|---|---|
| `↑` / `↓` | Scroll one entry |
| `PgUp` / `PgDn` | Scroll one page |
| `Home` | Jump to first entry |
| `End` | Jump to last entry |
| `E` | Export current results to JSON file |
| `Q` / `Esc` | Exit table, return to post-scan menu |

The table **adapts to the terminal window size** automatically on every frame — resize the window and the next key press will redraw at the new dimensions. The Title column expands to fill available horizontal space.

---

## JSON Export

Export is available two ways:

1. Press **`E`** inside the interactive table view
2. Select **"Save report to JSON file"** from the post-scan menu

The file is saved in the current working directory as:
```
cve_report_YYYYMMDD_HHmmss.json
```

### JSON Structure

```json
{
  "meta": {
    "system":    "Windows 11 Home 23H2 Build 22631 x64",
    "period":    "Last 3 month(s)",
    "generated": "2025-11-14T18:42:07",
    "total":     183
  },
  "cves": [
    {
      "CVE":                    "CVE-2025-54110",
      "Title":                  "Windows Kernel Elevation of Privilege Vulnerability",
      "CvssScore":              "8.8",
      "Criticality":            "Critical",
      "VulnType":               "Elevation of Privilege",
      "Exploited":              true,
      "PubliclyDisclosed":      false,
      "ExploitationLikely":     true,
      "CustomerActionRequired": true,
      "KB":                     "KB5058481",
      "PatchStatus":            "MISSING_PATCH",
      "URL":                    "https://msrc.microsoft.com/update-guide/vulnerability/CVE-2025-54110"
    }
  ]
}
```

---

## Post-Scan Menu

After viewing the results and closing the interactive table, a final menu appears:

```
  WHAT NEXT?
  >>  Save report to JSON file
      Run new analysis
      Exit
```

---

## MSRC API Reference

The script uses the public MSRC CVRF v2.0 API — no API key required.

| Endpoint | Purpose |
|---|---|
| `GET /cvrf/v2.0/updates` | List of all available monthly releases |
| `GET /cvrf/v2.0/cvrf/{id}` | Full CVRF document for a release (e.g. `2025-Sep`) |

Rate limiting: 200ms delay between requests.

API reference: https://api.msrc.microsoft.com/cvrf/v2.0/swagger

---

## Troubleshooting

**Zero CVEs found — `"total": 0` in JSON**

The most common cause is system detection producing no ProductTree matches. The script displays the detected parameters on the "System Detected" screen before configuration begins:

```
  English product name : Windows 11 Home
  Windows major        : 11
  Release ID           : 23H2
  MSRC arch string     : x64-based Systems
  Is Server            : False
```

Verify these values match what you expect. If `Release ID` is blank, the registry key `DisplayVersion` may be missing — the script will then match any release ID, which may return fewer results.

**Network timeout / no data**

- Confirm HTTPS access to `api.msrc.microsoft.com`
- Corporate proxies or firewalls may block the endpoint
- Check with: `Invoke-RestMethod 'https://api.msrc.microsoft.com/cvrf/v2.0/updates'`

**KB list shows 0 installed**

All three KB detection methods failed. This happens when:
- Group Policy blocks `Get-HotFix` and `wmic`
- CBS registry path is inaccessible

When KB list is empty, all CVEs with a known fix KB will show `MISSING_PATCH`.

**ExecutionPolicy error**

The script sets `Bypass` for the current process only. If a policy at machine or domain level blocks even process-scoped bypass, run:
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\vulncheck.ps1
```

---

## Credits

MSRC API integration approach and CVE parsing logic based on  
[patch_review.ps1](https://github.com/f-bader/msrc-api-ps) by **Fabian Bader**  
Original Python implementation by **Kevin Breen, Immersive Labs**
