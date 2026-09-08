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
        $safe = "pdf"
    }
    return $safe
}

function Invoke-ConvertDocx {
    param(
        [string]$SourceFile,
        [string]$TargetFile
    )

    Write-Host "DOCX : $SourceFile"
    Write-Host "PDF  : $TargetFile"
    & cscript //nologo .\convert_docx_to_pdf.vbs $SourceFile $TargetFile
}

function Expand-SourceRoot {
    param([string]$PathValue)

    $resolved = (Resolve-Path -LiteralPath $PathValue).Path
    $item = Get-Item -LiteralPath $resolved
    if ($item.PSIsContainer) {
        return [pscustomobject]@{
            Root = $resolved
            Temporary = $false
        }
    }

    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("recap_pdf_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null

    switch ($item.Extension.ToLowerInvariant()) {
        ".zip" {
            Expand-Archive -LiteralPath $resolved -DestinationPath $tempRoot -Force
        }
        ".rar" {
            $tar = Get-Command tar.exe -ErrorAction SilentlyContinue
            if (-not $tar) {
                throw "Archive .rar detectee, mais tar.exe est introuvable pour l'extraction."
            }
            & $tar.Path -xf $resolved -C $tempRoot
            if ($LASTEXITCODE -ne 0) {
                throw "Impossible d'extraire l'archive .rar."
            }
        }
        default {
            throw "Source non prise en charge: $resolved"
        }
    }

    return [pscustomobject]@{
        Root = $tempRoot
        Temporary = $true
    }
}

function Get-TargetRoot {
    param(
        [string]$SourcePath,
        [string]$ExplicitOutput
    )

    if ($ExplicitOutput) {
        $resolved = Resolve-Path -LiteralPath $ExplicitOutput -ErrorAction SilentlyContinue
        if ($resolved) {
            return $resolved.Path
        }
        return $ExplicitOutput
    }

    return Join-Path $outputDir ("{0}_pdf" -f (Get-SafeName (Split-Path -Leaf $SourcePath)))
}

function Get-PreferredSourcePath {
    param([string]$PathValue)

    $leaf = Split-Path -Leaf $PathValue
    $cachedCandidate = Join-Path $outputDir ("{0}_cachet" -f (Get-SafeName $leaf))
    if (Test-Path -LiteralPath $cachedCandidate -PathType Container) {
        return $cachedCandidate
    }
    return $PathValue
}

function Invoke-PdfRoot {
    param(
        [string]$SourcePath,
        [string]$TargetRoot
    )

    $sourceInfo = Expand-SourceRoot -PathValue $SourcePath
    $sourceRoot = $sourceInfo.Root
    $cleanupTemp = $sourceInfo.Temporary

    try {
        New-Item -ItemType Directory -Force -Path $TargetRoot | Out-Null
        $docs = Get-ChildItem -LiteralPath $sourceRoot -Recurse -File | Where-Object {
            $_.Extension.ToLowerInvariant() -eq ".docx" -and -not $_.Name.StartsWith("~$")
        } | Sort-Object FullName

        if (-not $docs) {
            throw "Aucun fichier DOCX trouve dans: $sourceRoot"
        }

        foreach ($docx in $docs) {
            $relative = $docx.FullName.Substring($sourceRoot.Length).TrimStart('\')
            $pdfRelative = [System.IO.Path]::ChangeExtension($relative, ".pdf")
            $targetFile = Join-Path $TargetRoot $pdfRelative
            $targetFolder = Split-Path -Parent $targetFile
            if ($targetFolder) {
                New-Item -ItemType Directory -Force -Path $targetFolder | Out-Null
            }
            if ((Test-Path -LiteralPath $targetFile -PathType Leaf) -and ((Get-Item -LiteralPath $targetFile).LastWriteTime -ge $docx.LastWriteTime)) {
                Write-Host "PDF deja cree : $targetFile"
                continue
            }
            Invoke-ConvertDocx -SourceFile $docx.FullName -TargetFile $targetFile
        }
    }
    finally {
        if ($cleanupTemp -and (Test-Path -LiteralPath $sourceRoot)) {
            Remove-Item -LiteralPath $sourceRoot -Recurse -Force
        }
    }
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
        $targetRoot = Join-Path $outputDir ($leaf -replace "_inspecteur$", "_pdf")
        Invoke-PdfRoot -SourcePath $source -TargetRoot $targetRoot
        Write-Host "Output root: $targetRoot"
    }
    exit 0
}

$inputFullPath = (Resolve-Path -LiteralPath $InputPath).Path

if (Test-Path -LiteralPath $inputFullPath -PathType Container) {
    $archives = Get-ChildItem -LiteralPath $inputFullPath -File | Where-Object {
        $_.Extension.ToLowerInvariant() -in @(".rar", ".zip")
    } | Sort-Object Name

    $hasDirectContent = Get-ChildItem -LiteralPath $inputFullPath -Force | Where-Object {
        $_.PSIsContainer -or $_.Extension.ToLowerInvariant() -eq ".docx"
    }

    if ($hasDirectContent) {
        $preferredSource = Get-PreferredSourcePath $inputFullPath
        $targetRoot = Get-TargetRoot -SourcePath $inputFullPath -ExplicitOutput $OutputPath
        Invoke-PdfRoot -SourcePath $preferredSource -TargetRoot $targetRoot
        Write-Host "Output root: $targetRoot"
    }

    if ($archives.Count -gt 0) {
        foreach ($archive in $archives) {
            $preferredSource = Get-PreferredSourcePath $archive.FullName
            $targetRoot = if ($OutputPath -and $archives.Count -eq 1) {
                $OutputPath
            } else {
                Join-Path $outputDir ("{0}_pdf" -f (Get-SafeName $archive.Name))
            }
            Invoke-PdfRoot -SourcePath $preferredSource -TargetRoot $targetRoot
            Write-Host "Output root: $targetRoot"
        }
        exit 0
    }

    if ($hasDirectContent) {
        exit 0
    }
}

$preferredSource = Get-PreferredSourcePath $inputFullPath
$targetRoot = Get-TargetRoot -SourcePath $InputPath -ExplicitOutput $OutputPath
Invoke-PdfRoot -SourcePath $preferredSource -TargetRoot $targetRoot
Write-Host "Output root: $targetRoot"
