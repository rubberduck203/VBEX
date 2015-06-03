# Build.ps1
#
# Collects VBEX files into an office add-in file
#
# Copywrite (C) 2015 Philip Wales
#
Param(
    [ValidateSet("Excel")]
        [String]$officeApp = "Excel"
)

# locations of required libraries change according to OS arch.
# Our libs are 32 bit.
$programFiles = if ([Environment]::Is64BitOperatingSystem) { 
	"Program Files (x86)"
} else { 
	"Program Files"
}

# constants are illegible in powershell
$VBA_EXTENSIBILITY_LIB = "C:\$programFiles\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
$VBA_SCRIPTING_LIB = "C:\Windows\system32\scrrun.dll"
$RUBBERDUCK_LIB = "C:\$programFiles\Rubberduck\Rubberduck\Rubberduck.tlb"
$buildScript = (Join-Path $PSScriptRoot "Build.ps1")

$srcPath = (Join-Path $PSScriptRoot "VBEX.xlam")
$srcFiles = (Get-ChildItem (Join-Path $PSScriptRoot "src")).FullName
$srcRefs = @("$VBA_EXTENSIBILITY_LIB", "$VBA_SCRIPTING_LIB")
& "$buildScript" "$srcPath" $srcFiles $srcRefs

$testPath = (Join-Path $PSScriptRoot "VBEXTesting.xlam")
$testFiles = (Get-ChildItem (Join-Path $PSScriptRoot "test")).FullName
$testRefs = @("$RUBBERDUCK_LIB")
& "$buildScript" "$testPath" $testFiles $testRefs