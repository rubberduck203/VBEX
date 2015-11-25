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
    Write-Host "Will Build $buildPath"
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
    
    Write-Host "Saving Addin as $outputPath" -ForeGround Green
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
    Write-Host "=> " -ForeGround Blue -noNewLine
    Write-Host "Building VBProject $name`:"
    $moduleCount = $moduleFiles.length
    Write-Host "==> " -ForeGround Green -noNewLine
    Write-Host "Importing $moduleCount Modules:"
    ForEach($moduleFile in $modulefiles) {
        Write-Host "   -> " -ForeGround Yellow -noNewLine 
        Write-Host "$moduleFile"
        $prj.VBComponents.Import($moduleFile)
    }
    $refCount = $references.length
    Write-Host "==> " -ForeGround Green -noNewLine
    Write-Host "Linking $refCount References:"
    ForEach($reference in $references) {
        Write-Host "   -> " -ForeGround Yellow -noNewLine
        Write-Host "$reference"
        $prj.References.AddFromFile( $reference ) 
    }
}
function dosEOLFolder([System.Array] $textFiles) {
    $count = $textFiles.length
    Write-Host "Converting $count files to CRLF"
    $textFiles | ForEach-Object { dosEOL $_ }
}
function dosEOL([String] $textFile) {
    $tempOut = "$textFile-CRLF"
    Get-Content $textFile | Set-Content $tempOut
    Remove-Item $textFile
    Move-Item $tempOut $textFile
}
main # entry point
