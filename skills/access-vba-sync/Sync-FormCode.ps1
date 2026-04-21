[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [Alias("FormName")]
    [string]$ModuleName,

    [Parameter()]
    [string]$ProjectRoot = (Get-Location).Path,

    [Parameter()]
    [string]$DestinationRoot = "src",

    [Parameter()]
    [string]$FormsRoot,

    [Parameter()]
    [string]$ReportsRoot,

    [switch]$WhatIf
)

$ErrorActionPreference = "Stop"

function Resolve-AbsolutePath {
    param(
        [string]$BasePath,
        [string]$MaybeRelativePath
    )

    if ([string]::IsNullOrWhiteSpace($MaybeRelativePath)) {
        return $null
    }

    if ([System.IO.Path]::IsPathRooted($MaybeRelativePath)) {
        return $MaybeRelativePath
    }

    return (Join-Path $BasePath $MaybeRelativePath)
}

function Get-NewLine {
    param([string]$Text)

    if ($Text -match "`r`n") {
        return "`r`n"
    }

    return "`n"
}

function Normalize-NewLines {
    param(
        [string]$Text,
        [string]$NewLine
    )

    if ($null -eq $Text) {
        return ""
    }

    return ($Text -replace "`r`n|`n|`r", $NewLine)
}

function Split-CodeBehindSection {
    param([string]$Text)

    $match = [regex]::Match($Text, "(?ms)^(?<Prefix>.*?)(^(?<Marker>CodeBehind[^\r\n]*)\r?\n)(?<Body>.*)$")
    if (-not $match.Success) {
        throw "No se encontro una seccion CodeBehind valida."
    }

    return @{
        Prefix = $match.Groups["Prefix"].Value
        Marker = $match.Groups["Marker"].Value
        Body   = $match.Groups["Body"].Value
    }
}

function Split-VbaMetadataHeader {
    param(
        [string]$Text,
        [string]$NewLine
    )

    $normalized = Normalize-NewLines -Text $Text -NewLine $NewLine
    $lines = [regex]::Split($normalized, [regex]::Escape($NewLine))

    $startIdx = 0
    $inBeginBlock = $false
    $foundMeta = $false

    for ($i = 0; $i -lt $lines.Length; $i++) {
        $line = $lines[$i].TrimStart()

        if ($inBeginBlock) {
            if ($line -eq "END") {
                $inBeginBlock = $false
            }
            continue
        }

        if ($line -match "^VERSION\s+") {
            $foundMeta = $true
            continue
        }

        if ($line -eq "BEGIN") {
            $foundMeta = $true
            $inBeginBlock = $true
            continue
        }

        if ($line -match "^Attribute\s+VB_") {
            $foundMeta = $true
            continue
        }

        if ($foundMeta -and $line -eq "") {
            continue
        }

        $startIdx = $i
        break
    }

    if ($lines.Length -eq 0) {
        return @{
            Header = ""
            Body   = ""
        }
    }

    if ($startIdx -eq 0 -and $foundMeta) {
        $startIdx = $lines.Length
    }

    if ($startIdx -gt 0) {
        $header = [string]::Join($NewLine, $lines[0..($startIdx - 1)])
    } else {
        $header = ""
    }

    if ($startIdx -lt $lines.Length) {
        $body = [string]::Join($NewLine, $lines[$startIdx..($lines.Length - 1)])
    } else {
        $body = ""
    }

    return @{
        Header = $header
        Body   = $body
    }
}

function Join-CodeBehindBody {
    param(
        [string]$Header,
        [string]$Body,
        [string]$NewLine
    )

    $parts = @()

    if (-not [string]::IsNullOrWhiteSpace($Header)) {
        $normalizedHeader = Normalize-NewLines -Text $Header -NewLine $NewLine
        $parts += $normalizedHeader.TrimEnd("`r", "`n")
    }

    if (-not [string]::IsNullOrWhiteSpace($Body)) {
        $normalizedBody = Normalize-NewLines -Text $Body -NewLine $NewLine
        $parts += $normalizedBody.TrimEnd("`r", "`n")
    }

    if ($parts.Count -eq 0) {
        return ""
    }

    return [string]::Join($NewLine, $parts)
}

$destinationRootAbs = Resolve-AbsolutePath -BasePath $ProjectRoot -MaybeRelativePath $DestinationRoot

if ($FormsRoot) {
    $formsRootAbs = Resolve-AbsolutePath -BasePath $ProjectRoot -MaybeRelativePath $FormsRoot
} else {
    $formsRootAbs = Join-Path $destinationRootAbs "forms"
}

if ($ReportsRoot) {
    $reportsRootAbs = Resolve-AbsolutePath -BasePath $ProjectRoot -MaybeRelativePath $ReportsRoot
} else {
    $reportsRootAbs = Join-Path $destinationRootAbs "reports"
}

$formDocPath = Join-Path $formsRootAbs ($ModuleName + ".form.txt")
$formClsPath = Join-Path $formsRootAbs ($ModuleName + ".cls")
$reportDocPath = Join-Path $reportsRootAbs ($ModuleName + ".report.txt")
$reportClsPath = Join-Path $reportsRootAbs ($ModuleName + ".cls")

$docPath = $null
$clsPath = $null
$docKind = $null

if ((Test-Path -LiteralPath $formDocPath) -and (Test-Path -LiteralPath $formClsPath)) {
    $docPath = $formDocPath
    $clsPath = $formClsPath
    $docKind = "Form"
} elseif ((Test-Path -LiteralPath $reportDocPath) -and (Test-Path -LiteralPath $reportClsPath)) {
    $docPath = $reportDocPath
    $clsPath = $reportClsPath
    $docKind = "Report"
}

if (-not $docPath -or -not $clsPath) {
    throw "No se encontraron archivos sincronizables para '$ModuleName' en '$formsRootAbs' ni '$reportsRootAbs'."
}

Write-Host ("Sincronizando CodeBehind: {0} ({1})" -f $ModuleName, $docKind) -ForegroundColor Cyan

$docContent = Get-Content -LiteralPath $docPath -Raw -Encoding UTF8
$clsContent = Get-Content -LiteralPath $clsPath -Raw -Encoding UTF8
$newLine = Get-NewLine -Text $docContent

$docSection = Split-CodeBehindSection -Text $docContent
$docParsed = Split-VbaMetadataHeader -Text $docSection.Body -NewLine $newLine
$clsParsed = Split-VbaMetadataHeader -Text $clsContent -NewLine $newLine

if (-not [string]::IsNullOrWhiteSpace($docParsed.Header)) {
    $headerToKeep = $docParsed.Header
} else {
    $headerToKeep = $clsParsed.Header
}

$rebuiltBody = Join-CodeBehindBody -Header $headerToKeep -Body $clsParsed.Body -NewLine $newLine
$normalizedPrefix = Normalize-NewLines -Text $docSection.Prefix -NewLine $newLine
$markerLine = if ([string]::IsNullOrWhiteSpace($docSection.Marker)) { "CodeBehind" } else { $docSection.Marker }
$newContent = $normalizedPrefix + $markerLine + $newLine + $rebuiltBody

if ($WhatIf) {
    Write-Host "[WhatIf] Se reemplazaria solo el cuerpo real del codigo y se preservaria el header del CodeBehind del documento." -ForegroundColor Yellow
    Write-Host ("   Documento: {0}" -f $docPath) -ForegroundColor Gray
    Write-Host ("   Clase:     {0}" -f $clsPath) -ForegroundColor Gray
    exit 0
}

Set-Content -LiteralPath $docPath -Value $newContent -Encoding UTF8 -NoNewline
Write-Host ("Sincronizado: {0}" -f $docPath) -ForegroundColor Green
Write-Host ("   Se preservo el header del CodeBehind del documento y se sustituyo solo el cuerpo desde {0}" -f $clsPath) -ForegroundColor Gray
