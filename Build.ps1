# Build.ps1
#
# Collects VBEX files into an office add-in file
#
# Copywrite (C) 2015 Philip Wales
#
Param(
	[String]$buildPath,
	[System.Array]$refs,
    [String]$sourceDir
)

# locations of required libraries change according to OS arch.
# Our libs are 32 bit.
$programFiles = if ([Environment]::Is64BitOperatingSystem) { 
	"Program Files (x86)"
} else { 
	"Program Files"
}

function main {

	$fileExt = [System.IO.Path]::GetExtension($buildPath)
	$officeCOM = switch -wildcard ($fileExt.ToLower()) {
        ".xl*" {New-Object -ComObject Excel.Application; break}
        ".ac*" {throw "Access is not yet supported"; break} #{New-Object -ComObject Acces.Application; break}
        default {throw "$fileName is not a supported office file."}
    } 
    $srcs = (Get-ChildItem $sourceDir).FullName
    dosEOLFolder $srcs
    $srcAddin = (BuildAddin $officeCOM $srcs $refs $buildPath)
    $officeCOM.Quit()
}
function BuildAddin($officeCOM, 
                    [System.Array] $moduleFiles, 
                    [System.Array] $references,
                    [String] $outputPath) {

    $newFile = $officeCOM.Workbooks.Add()
    $prj = $newFile.VBProject
	
	$projectName = [System.IO.Path]::GetFileNameWithoutExtension($outputPath)
    $prj.Name = $projectName
	
	$moduleFiles | ForEach-Object { $prj.VBComponents.Import( $_ ) }
	$references | ForEach-Object { $prj.References.AddFromFile( $_ ) }
    
    #save as addin
    $newFile.SaveAs($outputPath, $addinFormat)
    return $newFile
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
