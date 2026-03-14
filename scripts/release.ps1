git checkout develop
git pull

git checkout main
git pull

git merge develop

$version = powershell ./scripts/next_release.ps1

Write-Host ""
Write-Host "Creating release $version"
Write-Host ""

git tag $version

git push origin main
git push origin $version

$lastTag = git tag --sort=-creatordate | Select-Object -Skip 1 -First 1

$changelog = powershell ./scripts/generate_changelog.ps1 $lastTag

gh release create $version --notes "$changelog"

Write-Host ""
Write-Host "Release published:"
Write-Host $version