param(
    [switch]$dump,
    [string]$word,
    [int]$seed,
    [string]$envFile,
    [switch]$h,
    [switch]$help
)

# ---------- Help Menu ----------
if ($h -or $help) {
    Write-Host ""
    Write-Host "Usage:" -ForegroundColor Yellow
    Write-Host "  .\x.ps1 [options]" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Options:" -ForegroundColor Yellow
    Write-Host "  -word <string>        Build a word using characters found in environment variables."
    Write-Host "  -dump                 Save the full environment character map to TXT and JSON in the script directory."
    Write-Host "  -seed <int>           Use a fixed seed to make random selections reproducible."
    Write-Host "  -envFile <path>       Load environment variables from a TXT or JSON file instead of your live system environment."
    Write-Host "  -h, --help            Show this help menu and exit."
    Write-Host ""
    Write-Host "About -envFile:" -ForegroundColor Yellow
    Write-Host "  Use this option when you want to load environment variables from a file"
    Write-Host "  instead of using the ones currently set on your computer."
    Write-Host ""
    Write-Host "  The script supports two file formats:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "   1. TXT format  (each line is NAME=VALUE):"
    Write-Host "      Example file (env.txt):" -ForegroundColor DarkGray
    Write-Host "        USERNAME=fady"
    Write-Host "        COMPUTERNAME=FADY-PC"
    Write-Host "        PATH=C:\Tools;C:\Windows"
    Write-Host "        DriverData=SOMETHING"
    Write-Host ""
    Write-Host "      Run with:"
    Write-Host "        .\x.ps1 -envFile '.\env.txt' -word 'fady'" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "   2. JSON format  (object with key-value pairs):"
    Write-Host "      Example file (env.json):" -ForegroundColor DarkGray
    Write-Host "        {"
    Write-Host "          `"USERNAME`": `"fady`","
    Write-Host "          `"COMPUTERNAME`": `"FADY-PC`","
    Write-Host "          `"PATH`": `"C:\\Tools;C:\\Windows`","
    Write-Host "          `"DriverData`": `"SOMETHING`""
    Write-Host "        }"
    Write-Host ""
    Write-Host "      Run with:"
    Write-Host "        .\x.ps1 -envFile '.\env.json' -word 'fady'" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Tips:" -ForegroundColor Yellow
    Write-Host "   The script can randomize which environment variable it picks if multiple match the same character."
    Write-Host "   Use -seed <number> to make random results reproducible."
    Write-Host "   Use -dump to export the complete mapping to files."
    Write-Host ""
    Write-Host "Examples:" -ForegroundColor Yellow
    Write-Host "  .\x.ps1 -word 'fady'"
    Write-Host "  .\x.ps1 -word 'fady' -seed 42"
    Write-Host "  .\x.ps1 -word 'fady' -envFile '.\env.txt'"
    Write-Host "  .\x.ps1 -envFile '.\env.json' -dump"
    Write-Host ""
    Write-Host "=================================" -ForegroundColor Cyan
    exit
}

Write-Host "===== Building reverse character map from environment variables =====" -ForegroundColor Cyan

# ---------- Load Environment Variables ----------
function Load-EnvFromFile {
    param([Parameter(Mandatory=$true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "envFile not found: $Path"
    }

    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    $result = @{}

    if ($ext -eq ".json") {
        $raw = Get-Content -LiteralPath $Path -Raw
        $obj = $raw | ConvertFrom-Json
        if ($null -eq $obj) { throw "Invalid or empty JSON in $Path" }

        if ($obj -is [System.Collections.IDictionary]) {
            foreach ($k in $obj.Keys) { $result[$k] = [string]$obj[$k] }
        } else {
            # Accept array of {Name,Value}
            foreach ($item in $obj) {
                if ($item.PSObject.Properties.Name -contains 'Name' -and
                    $item.PSObject.Properties.Name -contains 'Value') {
                    $result[$item.Name] = [string]$item.Value
                }
            }
        }
    }
    elseif ($ext -eq ".txt") {
        $lines = Get-Content -LiteralPath $Path
        foreach ($line in $lines) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            $trim = $line.Trim()
            if ($trim.StartsWith("#") -or $trim.StartsWith(";")) { continue }
            $eqIndex = $trim.IndexOf("=")
            if ($eqIndex -lt 1) { continue }
            $name  = $trim.Substring(0, $eqIndex).Trim()
            $value = $trim.Substring($eqIndex + 1)
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                $result[$name] = [string]$value
            }
        }
    }
    else {
        throw "Unsupported envFile extension: $ext (use .json or .txt)"
    }

    return $result
}

$envSource = @{}
if ($PSBoundParameters.ContainsKey('envFile') -and $envFile) {
    try {
        Write-Host "Loading environment variables from file: $envFile" -ForegroundColor DarkCyan
        $envSource = Load-EnvFromFile -Path $envFile
    } catch {
        Write-Host "Error loading envFile: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
} else {
    Get-ChildItem Env: | ForEach-Object {
        $envSource[$_.Name] = [string]$_.Value
    }
}

# ---------- Build Character Map ----------
$charMap = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[object]]'

function Add-CharRef {
    param([string]$Char, [string]$Name, [int]$Index)
    if (-not $charMap.ContainsKey($Char)) {
        $charMap[$Char] = New-Object 'System.Collections.Generic.List[object]'
    }
    $charMap[$Char].Add([pscustomobject]@{
        Char  = $Char
        Name  = $Name
        Index = $Index
        Expr  = "[Environment]::GetEnvironmentVariable('$Name')[$Index]"
    })
}

# Only printable ASCII 32..126
foreach ($pair in $envSource.GetEnumerator()) {
    $name  = $pair.Key
    $value = $pair.Value
    if ([string]::IsNullOrEmpty($value)) { continue }

    for ($i = 0; $i -lt $value.Length; $i++) {
        $ch = $value[$i]
        $code = [int][char]$ch
        if ($code -ge 32 -and $code -le 126) {
            Add-CharRef -Char ([string]$ch) -Name $name -Index $i
        }
    }
}

Write-Host "Character map complete!" -ForegroundColor Green

# ---------- Helpers ----------
$Upper   = 65..90
$Lower   = 97..122
$Digits  = 48..57
$Sym1 = 32..47; $Sym2 = 58..64; $Sym3 = 91..96; $Sym4 = 123..126
$Symbols = @($Sym1 + $Sym2 + $Sym3 + $Sym4)
$AllPrintable = 32..126

function Label-ForChar([int]$c) {
    switch ($c) {
        32 { return "<space>" }
        92 { return '\' }
        default { return ([string][char]$c) }
    }
}

function Print-Group {
    param([string]$Title, [int[]]$Codes)
    Write-Host ""
    Write-Host "=== $Title ===" -ForegroundColor Yellow
    foreach ($c in $Codes) {
        $char  = [string][char]$c
        $label = Label-ForChar $c
        Write-Host "`n[$label]" -ForegroundColor Cyan
        if ($charMap.ContainsKey($char)) {
            foreach ($item in ($charMap[$char] | Sort-Object { $_.Expr.Length })) {
                Write-Host ("{0}  # '{1}'" -f $item.Expr, $item.Char)
            }
        } else {
            Write-Host "  (no matches)" -ForegroundColor DarkGray
        }
    }
}

# Random picker (supports optional seed for reproducibility)
function Pick-Choice {
    param([object[]]$Choices)
    if ($Choices.Count -le 1) { return $Choices[0] }
    if ($PSBoundParameters.ContainsKey('seed')) {
        return Get-Random -InputObject $Choices -SetSeed $seed
    } else {
        return Get-Random -InputObject $Choices
    }
}

# Build word using randomized picks
function Resolve-Word {
    param([Parameter(Mandatory=$true)][string]$Text)

    $parts   = New-Object System.Collections.Generic.List[string]
    $steps   = New-Object System.Collections.Generic.List[string]
    $missing = New-Object System.Collections.Generic.List[string]

    for ($k = 0; $k -lt $Text.Length; $k++) {
        $ch = [string]$Text[$k]
        if ($charMap.ContainsKey($ch) -and $charMap[$ch].Count -gt 0) {
            $choice = Pick-Choice -Choices $charMap[$ch]
            $parts.Add("([string]$($choice.Expr))")
            $steps.Add(("'{0}' -> {1}  # index {2} in ${3}" -f $ch, $choice.Expr, $choice.Index, $choice.Name))
        } else {
            $missing.Add(("'{0}' (pos {1})" -f $ch, $k))
        }
    }

    $exprLine = '$result = -join @(' + ($parts -join ', ') + '); Write-Host $result'
    return [pscustomobject]@{ Steps=$steps; Missing=$missing; ExprLine=$exprLine; Preview=$Text }
}

# ---------- Modes ----------
if ($word) {
    Write-Host ""
    Write-Host ("=== Constructing word: ""{0}"" ===" -f $word) -ForegroundColor Yellow

    $res = Resolve-Word -Text $word

    if ($res.Missing.Count -gt 0) {
        Write-Host "Missing characters:" -ForegroundColor Red
        $res.Missing | ForEach-Object { Write-Host "  $_" }
    }

    if ($res.Steps.Count -gt 0) {
        Write-Host "`nPer-character recipe (randomized when multiple options exist):" -ForegroundColor Cyan
        $res.Steps | ForEach-Object { Write-Host "  $_" }
    }

    Write-Host "`nOne-liner to reconstruct the word:" -ForegroundColor Cyan
    Write-Host $res.ExprLine

    Write-Host "`nLive preview:" -ForegroundColor Cyan
    Write-Host $res.Preview -ForegroundColor Green
}
elseif ($dump) {
    # Dump mode (no -word)
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
    if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
    $outTxt  = Join-Path $scriptDir "env_char_map.txt"
    $outJson = Join-Path $scriptDir "env_char_map.json"

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($c in $AllPrintable) {
        $char  = [string][char]$c
        $head  = "[" + (Label-ForChar $c) + "]"
        $lines.Add("")
        $lines.Add($head)
        if ($charMap.ContainsKey($char)) {
            foreach ($item in ($charMap[$char] | Sort-Object { $_.Expr.Length })) {
                $lines.Add(("{0}  # '{1}'" -f $item.Expr, $item.Char))
            }
        } else {
            $lines.Add("  (no matches)")
        }
    }
    $lines | Out-File -FilePath $outTxt -Encoding UTF8

    $ordered = [ordered]@{}
    foreach ($c in $AllPrintable) {
        $char = [string][char]$c
        if ($charMap.ContainsKey($char)) {
            $ordered[$char] = @(($charMap[$char] | Sort-Object { $_.Expr.Length }).Expr)
        } else {
            $ordered[$char] = @()
        }
    }
    $ordered | ConvertTo-Json -Depth 5 | Out-File -FilePath $outJson -Encoding UTF8

    Write-Host ""
    Write-Host "Dumped files:" -ForegroundColor Green
    Write-Host "  TXT : $outTxt"
    Write-Host "  JSON: $outJson"
}
else {
    # Default: print groups
    Print-Group -Title "Uppercase A-Z" -Codes $Upper
    Print-Group -Title "Lowercase a-z" -Codes $Lower
    Print-Group -Title "Digits 0-9"    -Codes $Digits
    Print-Group -Title "Symbols"        -Codes $Symbols
}

Write-Host ""
Write-Host "=================================" -ForegroundColor Cyan

