$export = (Join-Path "$PSScriptRoot" "Export.ps1")
$builds = @("src", "test")

ForEach($build in $builds) {
    $file = (Join-Path "$PSScriptRoot" "VBEX$build.xlam")
    $dest = (Join-Path "$PSScriptRoot" "$build")
    Get-ChildItem "$dest" | Remove-Item
    & $export "$file" "$dest"
}
