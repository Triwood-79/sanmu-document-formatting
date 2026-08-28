param(
    [string]$Python
)

$ErrorActionPreference = 'Stop'

function Resolve-Python {
    param([string]$Requested)
    if ($Requested) {
        if (Test-Path -LiteralPath $Requested) { return (Resolve-Path -LiteralPath $Requested).Path }
        throw "Specified Python executable was not found."
    }
    foreach ($name in @('python', 'py')) {
        $command = Get-Command $name -ErrorAction SilentlyContinue
        if ($command) { return $command.Source }
    }
    return $null
}

function Get-FontNames {
    $names = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($location in @(
        'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts',
        'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'
    )) {
        if (Test-Path -LiteralPath $location) {
            $item = Get-ItemProperty -LiteralPath $location
            foreach ($property in $item.PSObject.Properties) {
                if ($property.Name -notmatch '^PS') {
                    [void]$names.Add(($property.Name -replace '\s*\([^)]*\)\s*$', '').Trim())
                    if ($property.Value) {
                        [void]$names.Add([System.IO.Path]::GetFileNameWithoutExtension([string]$property.Value))
                    }
                }
            }
        }
    }
    return $names
}

$pythonPath = Resolve-Python -Requested $Python
$modules = [ordered]@{ python_docx = $false; lxml = $false }
if ($pythonPath) {
    $probe = & $pythonPath -c "import importlib.util,json; print(json.dumps({'python_docx': importlib.util.find_spec('docx') is not None, 'lxml': importlib.util.find_spec('lxml') is not None}))" 2>$null
    if ($LASTEXITCODE -eq 0 -and $probe) {
        $parsed = $probe | ConvertFrom-Json
        $modules.python_docx = [bool]$parsed.python_docx
        $modules.lxml = [bool]$parsed.lxml
    }
}

$fonts = Get-FontNames
$targets = @(
    @{ preferred = '方正小标宋_GBK'; fallback = '方正小标宋简体'; aliases = @() },
    @{ preferred = '方正黑体_GBK'; fallback = '黑体'; aliases = @('SimHei') },
    @{ preferred = '方正楷体_GBK'; fallback = '楷体_GB2312'; aliases = @('楷体GB2312', 'KaiTi_GB2312', 'KaiTi') },
    @{ preferred = '方正仿宋_GBK'; fallback = '仿宋_GB2312'; aliases = @('仿宋GB2312', 'FangSong_GB2312', 'FangSong') },
    @{ preferred = '宋体'; fallback = 'SimSun'; aliases = @('simsun') },
    @{ preferred = 'Times New Roman'; fallback = $null; aliases = @() }
)
$fontStatus = foreach ($target in $targets) {
    $fallbackCandidates = @($target.fallback) + @($target.aliases) | Where-Object { $_ }
    $fallbackAvailable = [bool]($fallbackCandidates | Where-Object { $fonts.Contains($_) } | Select-Object -First 1)
    [ordered]@{
        preferred = $target.preferred
        preferred_available = $fonts.Contains($target.preferred)
        fallback = $target.fallback
        fallback_available = $fallbackAvailable
    }
}

$renderer = $null
foreach ($name in @('soffice', 'libreoffice')) {
    $command = Get-Command $name -ErrorAction SilentlyContinue
    if ($command) { $renderer = $command.Source; break }
}
if (-not $renderer -and $env:OS -eq 'Windows_NT') {
    try {
        if ([type]::GetTypeFromProgID('Word.Application')) {
            $renderer = 'microsoft_word_com'
        }
    }
    catch {
        $renderer = $null
    }
}

[ordered]@{
    python = $pythonPath
    modules = $modules
    fonts = $fontStatus
    renderer = $renderer
    structural_ready = [bool]($pythonPath -and $modules.python_docx -and $modules.lxml)
    visual_ready = [bool]$renderer
    installs_performed = $false
} | ConvertTo-Json -Depth 5
