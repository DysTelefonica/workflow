param(
    [Parameter(Mandatory=$true)]
    [string]$FromTag
)

$log = git log $FromTag..HEAD --pretty=format:"%s"

Write-Host ""
Write-Host "Changelog since $FromTag"
Write-Host ""

foreach ($line in $log) {
    if ($line -match "spec-") {
        Write-Host "- $line"
    }
}