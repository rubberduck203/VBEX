param([String]$officeApp = "Excel", [String]$sourceDir, [String]$buildDir)

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

$ext = "xlam" # switch on EXCEL/ACCESS
$addinFormat = 55
$officeCOM = switch ($officeApp.ToUpper()) {
    "EXCEL" {New-Object -ComObject Excel.Application; break}
    #"ACCESS" {New-Object -ComObject Acces.Application; break}
    default {throw "$officeApp is not a supported office application."}
}

function BuildAddin($moduleFiles, [String] $outputPath, [String] $projectName) {

    $newFile = $officeCOM.Workbooks.Add()
    $prj = $newFile.VBProject
    $prj.Name = $projectName

    #add modules
    ForEach ($moduleFile in $moduleFiles) {
        $prj.VBComponents.Import($moduleFile.FullName)
    }

    #add references
    $prj.References.AddFromFile($VBA_EXTENSIBILITY_LIB)
    $prj.References.AddFromFile($VBA_SCRIPTING_LIB)

    #save as addin
    $newFile.SaveAs($outputPath, $addinFormat)
}

$buildPath = (Join-Path $buildDir "VBEX.$ext") 
$testPath = (Join-Path $buildDir "VBEX-Testing.$ext") 

$srcModuleFiles = (Get-ChildItem (Join-Path $sourceDir "src"))
$testModuleFiles = (Get-ChildItem (Join-Path $sourceDir "test"))

$srcAddin = BuildAddin $srcModuleFiles $buildPath "VBEX"
$testAddin = BuildAddin $testModuleFiles $testPath "VBEXTesting"

$officeCOM.Quit()