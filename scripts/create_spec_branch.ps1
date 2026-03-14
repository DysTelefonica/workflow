param(
    [Parameter(Mandatory=$true)]
    [string]$Version
)

git checkout main
git pull

git merge hotfix-$Version

git tag $Version

git push origin main
git push origin $Version

git checkout develop
git merge main
git push origin develop

Write-Host ""
Write-Host "Hotfix released:"
Write-Host $Version