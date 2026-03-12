param(
    [Parameter(Mandatory=$true)]
    [string]$Version
)

Write-Host "Starting release $Version"

# comprobar repo limpio
$changes = git status --porcelain
if ($changes) {
    Write-Host "Repo not clean. Commit or stash changes first."
    exit 1
}

# cambiar a main
git checkout main
git pull origin main

# merge develop
git merge develop --no-ff -m "Merge develop for release $Version"

# crear tag
git tag -a $Version -m "Release $Version"

# push
git push origin main
git push origin $Version

# generar changelog entre tags
$prevTag = git tag --sort=-creatordate | Select-Object -Skip 1 -First 1

$log = git log "$prevTag..$Version" --pretty=format:"- %s"

$releaseNotes = @"
## Release $Version

### Changes
$log
"@

# crear release en GitHub
gh release create $Version `
    --title "Release $Version" `
    --notes "$releaseNotes"

Write-Host "Release $Version created successfully"