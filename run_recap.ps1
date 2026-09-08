param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$InputPath,

    [Parameter(Mandatory = $false, Position = 1)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputDir = Join-Path $projectRoot "outputs"
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

function Get-SafeName {
    param([string]$Name)
    $safe = [System.IO.Path]::GetFileNameWithoutExtension($Name)
    $safe = $safe -replace '[^\p{L}\p{Nd}_ -]+', '_'
    $safe = $safe.Trim()
    if (-not $safe) {
        $safe = "recap"
    }
    return $safe
}

function Invoke-Recap {
    param(
        [string]$SourcePath,
        [string]$TargetPath
    )

    Write-Host "Input : $SourcePath"
    Write-Host "Output: $TargetPath"
    & rtk python .\recap_docx.py $SourcePath -o $TargetPath
}

Set-Location $projectRoot

if (-not $InputPath) {
    $statePath = Join-Path $outputDir ".last_inspecteur_paths.txt"
    $sources = @()
    if (Test-Path -LiteralPath $statePath) {
        $sources = @(Get-Content -LiteralPath $statePath | Where-Object { Test-Path -LiteralPath $_ -PathType Container })
    }
    if (-not $sources) {
        $sources = @(Get-ChildItem -LiteralPath $outputDir -Directory -Filter "*_inspecteur" | Sort-Object LastWriteTime | Select-Object -ExpandProperty FullName)
    }
    if (-not $sources) { throw "Aucun dossier DOCX signe '*_inspecteur' trouve dans outputs." }
    foreach ($source in $sources) {
        $leaf = Split-Path -Leaf $source
        $target = Join-Path $outputDir (($leaf -replace "_inspecteur$", "") + "_recap.xls")
        Invoke-Recap $source $target
    }
    exit 0
}

$inputFullPath = (Resolve-Path -LiteralPath $InputPath).Path

if (Test-Path -LiteralPath $inputFullPath -PathType Container) {
    $hasDirectContent = Get-ChildItem -LiteralPath $inputFullPath -Force | Where-Object {
        $_.PSIsContainer -or $_.Extension.ToLowerInvariant() -eq ".docx"
    }

    if ($hasDirectContent) {
        if (-not $OutputPath) {
            $targetName = "{0}_recap.xls" -f (Get-SafeName (Split-Path -Leaf $inputFullPath))
            $folderTarget = Join-Path $outputDir $targetName
        } else {
            $folderTarget = $OutputPath
        }
        Invoke-Recap $inputFullPath $folderTarget
    }

    $archives = Get-ChildItem -LiteralPath $inputFullPath -File | Where-Object {
        $_.Extension.ToLowerInvariant() -in @(".rar", ".zip")
    } | Sort-Object Name

    if ($archives.Count -gt 0) {
        foreach ($archive in $archives) {
            $target = Join-Path $outputDir ("{0}_recap.xls" -f (Get-SafeName $archive.Name))
            if ($OutputPath -and $archives.Count -eq 1 -and -not $hasDirectContent) {
                $target = $OutputPath
            }
            Invoke-Recap $archive.FullName $target
        }
        exit 0
    }

    if ($hasDirectContent) {
        exit 0
    }
}

if (-not $OutputPath) {
    $targetName = "{0}_recap.xls" -f (Get-SafeName (Split-Path -Leaf $inputFullPath))
    $OutputPath = Join-Path $outputDir $targetName
}

Invoke-Recap $inputFullPath $OutputPath
