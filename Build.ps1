# Build.ps1
#
# Collects VBEX files into an office add-in file
#
# Copywrite (C) 2015 Philip Wales
#
Param(
	[String]$buildPath,
	[System.Array]$sourceFiles,
	[System.Array]$references
)

function main {

	$fileExt = [System.IO.Path]::GetExtension($buildPath)
	$officeCOM = switch -wildcard ($fileExt.ToLower()) {
        ".xl*" {New-Object -ComObject Excel.Application; break}
        #".ac*" {New-Object -ComObject Acces.Application; break}
        default {throw "$fileName is not a supported office file."}
    } 
    dosEOLFolder $sourceFiles
    $srcAddin = (BuildExcelAddin $officeCOM $sourceFiles $references $buildPath)
    $officeCOM.Quit()
}
function BuildExcelAddin($officeCOM, 
                    [System.Array] $moduleFiles, 
                    [System.Array] $references,
                    [String] $outputPath) {

    $newFile = $officeCOM.Workbooks.Add()
    $prj = $newFile.VBProject
    $projectName = [System.IO.Path]::GetFileNameWithoutExtension($outputPath)
    BuildVBProject $prj $projectName $moduleFiles $references
    
    #save as addin
    $newFile.SaveAs($outputPath, 55)
    return $newFile
}
function BuildAccessAddin($officeCOM, [System.Array] $moduleFiles, 
        [System.Array] $references, [String] $outputPath) {

    $newDB = $officeCOM.DBEngine.CreateDatabase($outputPath)
    $prj = $officeCOM.VBE.VBProjects(1)
    $projectName = [System.IO.Path]::GetFileNameWithoutExtension($outputPath)
    BuildVBProject $prj $projectName $moduleFiles $references
    
    return $newDB
}
function BuildVBProject($prj, [String] $name, [System.Array] $moduleFiles,
        [System.Array] $references) {
    
    $prj.Name = $name
    $moduleFiles | ForEach-Object { $prj.VBComponents.Import( $_ ) }
    $references | ForEach-Object { $prj.References.AddFromFile( $_ ) }
    
}
function dosEOLFolder([System.Array] $textFiles) {
    $textFiles | ForEach-Object { dosEOL $_ }
}
function dosEOL([String] $textFile) {
    $tempOut = "$textFile-CRLF"
    Get-Content $textFile | Set-Content $tempOut
    Remove-Item $textFile
    Move-Item $tempOut $textFile
}
main # entry point
