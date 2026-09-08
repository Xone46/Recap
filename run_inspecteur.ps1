param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$InputPath,

    [Parameter(Mandatory = $false, Position = 1)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$SignaturePath
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputDir = Join-Path $projectRoot "outputs"
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

function Get-InspectorOutputPath {
    param([string]$SourcePath)
    $leaf = Split-Path -Leaf $SourcePath
    if ($leaf.EndsWith("_cachet", [System.StringComparison]::OrdinalIgnoreCase)) {
        return Join-Path (Split-Path -Parent $SourcePath) ($leaf.Substring(0, $leaf.Length - 7) + "_inspecteur")
    }
    return Join-Path $outputDir (([System.IO.Path]::GetFileNameWithoutExtension($leaf)) + "_inspecteur")
}

function Invoke-Inspector {
    param([string]$SourcePath, [string]$TargetPath)
    $args = @(".\inspecteur_reports.py", $SourcePath, "-o", $TargetPath)
    if ($SignaturePath) { $args += @("--signature", $SignaturePath) }
    & rtk python @args | Write-Host
}

Set-Location $projectRoot

if (-not $InputPath) {
    $statePath = Join-Path $outputDir ".last_cachet_paths.txt"
    if (-not (Test-Path -LiteralPath $statePath)) {
        throw "Aucune sortie de cachet trouvee. Lancez d'abord run_cachet.bat."
    }
    $sources = Get-Content -LiteralPath $statePath | Where-Object { Test-Path -LiteralPath $_ -PathType Container }
    if (-not $sources) { throw "Les sorties de cachet enregistrees sont introuvables." }
    $targets = foreach ($source in $sources) {
        $target = Get-InspectorOutputPath $source
        Invoke-Inspector $source $target
        if ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $target -PathType Container)) { $target }
    }
    if (-not $targets) { throw "Aucun rapport n'a pu etre signe par l'Inspecteur." }
    Set-Content -LiteralPath (Join-Path $outputDir ".last_inspecteur_paths.txt") -Value $targets -Encoding utf8
    exit 0
}

$inputFullPath = (Resolve-Path -LiteralPath $InputPath).Path

if (-not $OutputPath) { $OutputPath = Get-InspectorOutputPath $inputFullPath }

Invoke-Inspector $inputFullPath $OutputPath
if ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $OutputPath -PathType Container)) {
    Set-Content -LiteralPath (Join-Path $outputDir ".last_inspecteur_paths.txt") -Value (Resolve-Path -LiteralPath $OutputPath).Path -Encoding utf8
}
