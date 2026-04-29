# Sync-FormCode.ps1
# Synchronizes .cls code with CodeBehind section in .form.txt files
# Usage: .\Sync-FormCode.ps1 <formName> [-WhatIf]

param(
    [Parameter(Mandatory=$true)]
    [string]$FormName,
    [string]$FormsRoot = "src/forms",
    [switch]$WhatIf
)

$ErrorActionPreference = "Stop"

# Resolve paths
$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) { $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }
$projectRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }

$clsPath = Join-Path $projectRoot "$FormsRoot/$FormName.cls"
$formPath = Join-Path $projectRoot "$FormsRoot/$FormName.form.txt"

if (-not (Test-Path $clsPath)) {
    Write-Error "❌ No se encontró: $clsPath"
}
if (-not (Test-Path $formPath)) {
    Write-Error "❌ No se encontró: $formPath"
}

Write-Host "📋 Sincronizando: $FormName" -ForegroundColor Cyan

# Read .cls content
$clsContent = Get-Content $clsPath -Raw -Encoding UTF8

# Read form content and find CodeBehind
$formContent = Get-Content $formPath -Raw -Encoding UTF8
$cbIndex = $formContent.IndexOf("CodeBehind")

if ($cbIndex -eq -1) {
    Write-Error "❌ No se encontró CodeBehind en $FormName.form.txt"
}

# Extract the part before CodeBehind (UI definitions)
$uiPart = $formContent.Substring(0, $cbIndex)

# Build new CodeBehind (header + .cls content)
$newCodeBehind = "CodeBehind`r`n" + $clsContent

if ($WhatIf) {
    Write-Host "🔍 [WhatIf] Se substituiría CodeBehind completo" -ForegroundColor Yellow
    Write-Host "   Longitud actual: $($formContent.Length)" -ForegroundColor Gray
    Write-Host "   Nueva longitud: $($uiPart.Length + $newCodeBehind.Length)" -ForegroundColor Gray
} else {
    # Write new form content
    $newFormContent = $uiPart + $newCodeBehind
    Set-Content $formPath -Value $newFormContent -Encoding UTF8 -NoNewline
    
    Write-Host "✅ Sincronizado: $FormName.form.txt" -ForegroundColor Green
    Write-Host "   CodeBehind actualizado con contenido de $FormName.cls" -ForegroundColor Gray
}