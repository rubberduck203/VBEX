$scriptRoot = if ($PSVersionTable.PSVersion.Major -ge 3) {
    $PSScriptRoot
} else {
    Split-Path $MyInvocation.MyCommand.Path -Parent
}
$export = (Join-Path "$scriptRoot" "Export.ps1")
$builds = @("src", "test")

ForEach($build in $builds) {
    $file = (Join-Path "$scriptRoot" "VBEX$build.xlam")
    $dest = (Join-Path "$scriptRoot" "$build")
    Get-ChildItem "$dest" | Remove-Item
    & $export "$file" "$dest"
}
