$tags = git tag

$year = (Get-Date).Year

$numbers = @()

foreach ($tag in $tags) {
    if ($tag -match "$year-(\d+)") {
        $numbers += [int]$matches[1]
    }
}

if ($numbers.Count -eq 0) {
    $next = 1
}
else {
    $next = ($numbers | Measure-Object -Maximum).Maximum + 1
}

$version = "$year-" + "{0:D3}" -f $next

Write-Host $version