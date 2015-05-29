# Build.ps1
#
# Collects VBEX files into an office add-in file
#
# Copywrite (C) 2015 Philip Wales
#
Param(
    [ValidateSet("Excel")]
        [String]$officeApp = "Excel",
    [String]$sourceDir = "$PWD",
    [String]$buildDir = "$PWD"
)

# locations of required libraries change according to OS arch.
# Our libs are 32 bit.
if ([Environment]::Is64BitOperatingSystem) {
    $programFiles = "Program Files (x86)"
} else {
    $programFiles = "Program Files"
}

# constants are illegible in powershell
$VBA_EXTENSIBILITY_LIB = "C:\$programFiles\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
$VBA_EXTENSIBILITY_NAME = "VBIDE"
$VBA_SCRIPTING_LIB = "C:\Windows\system32\scrrun.dll"
$VBA_SCRIPTING_NAME = "Scripting"
$RUBBERDUCK_LIB = "C:\$programFiles\Rubberduck\Rubberduck\Rubberduck.tlb"
$RUBBERDUCK_NAME = "Rubberduck"

# enums are also equally illegible
$COMPTYPE_stdModule = 1
$COMPTYPE_classModule = 2
$COMPTYPE_msForm = 3
$COMPTYPE_activeXDesigner = 11
$COMPTYPE_document = 100

function main {

    $ext = "xlam" # switch on EXCEL/ACCESS
    $addinFormat = 55 # switch on EXCEL/ACCESS
    $officeCOM = GetOfficeCOM($officeApp)
    
    $buildPath = (Join-Path $buildDir "VBEX.$ext") 
    $testPath = (Join-Path $buildDir "VBEX-Testing.$ext") 

    $srcs = (Get-ChildItem (Join-Path $sourceDir "src")).FullName
    $tests = (Get-ChildItem (Join-Path $sourceDir "test")).FullName
    
    dosEOLFolder $srcs
    dosEOLFolder $tests
    
    $srcRefs = @($VBA_EXTENSIBILITY_LIB, $VBA_SCRIPTING_LIB)
    $testRefs = $srcRefs + @($RUBBERDUCK_LIB)
    
    $srcAddin = BuildAddin $officeCOM $srcs $srcRefs $buildPath "VBEX"
    $testAddin = BuildAddin $officeCOM $tests $testRefs $testPath "VBEXTesting"
    
    $officeCOM.Quit()
}
function BuildAddin($officeCOM, 
                    [System.Array] $moduleFiles, 
                    [System.Array] $references,
                    [String] $outputPath,
                    [String] $projectName) {

    $newFile = $officeCOM.Workbooks.Add()
    $prj = $newFile.VBProject
    $prj.Name = $projectName

    #add modules
    ForEach ($moduleFile in $moduleFiles) {
        $prj.VBComponents.Import($moduleFile)
    }

    #add references
    ForEach ($reference in $references) {
        $prj.References.AddFromFile($reference)
    }
    
    #save as addin
    $newFile.SaveAs($outputPath, $addinFormat)
    return $newFile
}
function GetOfficeCom([String] $officeAppName) {

    $officeCOM = switch ($officeApp.ToUpper()) {
        "EXCEL" {New-Object -ComObject Excel.Application; break}
        #"ACCESS" {New-Object -ComObject Acces.Application; break}
        default {throw "$officeApp is not a supported office application."}
    }
    
    return $officeCOM
}
function dosEOLFolder([System.Array] $textFiles) {
    ForEach ($textFile in $textFiles) {
        dosEOL $textFile
    }
}
function dosEOL([String] $textFile) {
    $tempOut = "$textFile-CRLF"
    Get-Content $textFile | Set-Content $tempOut
    Remove-Item $textFile
    Move-Item $tempOut $textFile
}
# Entry
main