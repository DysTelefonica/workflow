param(
    [Parameter(Mandatory=$true)]
    [string]$Version
)

Write-Host ""
Write-Host "================================="
Write-Host "Creating release $Version"
Write-Host "================================="
Write-Host ""

git checkout develop
git pull

git checkout main
git pull

git merge develop

git tag $Version

git push origin main
git push origin $Version

Write-Host ""
Write-Host "Release created:"
Write-Host $Version