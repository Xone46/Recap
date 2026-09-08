param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$InputPath = ".\input",

    [Parameter(Mandatory = $false, Position = 1)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$CachetPath

)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$inputFullPath = (Resolve-Path -LiteralPath $InputPath).Path
$outputDir = Join-Path $projectRoot "outputs"
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
$completedOutputs = [System.Collections.Generic.List[string]]::new()

function Save-WorkflowState {
    if ($completedOutputs.Count -gt 0) {
        Set-Content -LiteralPath (Join-Path $outputDir ".last_cachet_paths.txt") -Value $completedOutputs -Encoding utf8
    }
}

function Get-SafeName {
    param([string]$Name)
    $safe = [System.IO.Path]::GetFileNameWithoutExtension($Name)
    $safe = $safe -replace '[^\p{L}\p{Nd}_ -]+', '_'
    $safe = $safe.Trim()
    if (-not $safe) {
        $safe = "reports"
    }
    return $safe
}

function Invoke-Stamp {
    param(
        [string]$SourcePath,
        [string]$TargetPath
    )

    Write-Host "Input : $SourcePath"
    Write-Host "Output: $TargetPath"

    $args = @(".\stamp_reports.py", $SourcePath, "-o", $TargetPath, "--sans-signature")
    if ($CachetPath) {
        $args += @("--cachet", $CachetPath)
    }
    & rtk python @args
    if ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $TargetPath -PathType Container)) {
        $completedOutputs.Add((Resolve-Path -LiteralPath $TargetPath).Path)
    }
}

Set-Location $projectRoot

if (Test-Path -LiteralPath $inputFullPath -PathType Container) {
    $archives = Get-ChildItem -LiteralPath $inputFullPath -File | Where-Object {
        $_.Extension.ToLowerInvariant() -in @(".rar", ".zip")
    } | Sort-Object Name

    if ($archives.Count -gt 0) {
        foreach ($archive in $archives) {
            $target = if ($OutputPath -and $archives.Count -eq 1) {
                $OutputPath
            } else {
                Join-Path $outputDir ("{0}_cachet" -f (Get-SafeName $archive.Name))
            }
            Invoke-Stamp $archive.FullName $target
        }
        Save-WorkflowState
        exit 0
    }
}

if (-not $OutputPath) {
    $targetName = "{0}_cachet" -f (Get-SafeName (Split-Path -Leaf $inputFullPath))
    $OutputPath = Join-Path $outputDir $targetName
}

Invoke-Stamp $inputFullPath $OutputPath
Save-WorkflowState
