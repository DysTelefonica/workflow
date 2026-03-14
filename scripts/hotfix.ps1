param(
    [Parameter(Mandatory=$true)]
    [string]$Issue
)

$branch = "hotfix-$Issue"

git checkout main
git pull

git checkout -b $branch

git push -u origin $branch