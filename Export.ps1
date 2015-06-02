# Export.ps1
# 
# Exports all VBE files from a office file to a path
#
# Copywrite (C) 2015 Philip Wales
#
Param(
    [String]$sourceFile,
    [String]$outDest
)

function main {
   $fileCOM = (OpenOfficeFile $sourceFile)
   $prjCOM = ($fileCOM.VBProject)
   ExportModules $prjCOM $outDest
}
function OfficeComFromFileName([String] $fileName) {
}
# not the FOSS OpenOffice
function OpenOfficeFile([String] $sourceFile) {
}
function ExportModules ($prjCOM, [String] $outDest) {
}
#entry point
main
