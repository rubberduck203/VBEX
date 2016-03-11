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


# Compatible with earlier powershell versions
$scriptRoot = if ($PSVersionTable.PSVersion.Major -ge 3) {
    $PSScriptRoot
} else {
    Split-Path $MyInvocation.MyCommand.Path -Parent
}
$buildScript = (Join-Path $scriptRoot "Build.ps1")

$ext = switch ($officeApp) {
    "Excel" {"xlam"}
    # "Access" {"accde"} Wow is the Acces Object Model shit.
    default {throw "$officeApp is not a supported office application."}
}

$VBA_EXTENSIBILITY_LIB = "C:\$programFiles\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
$VBA_SCRIPTING_LIB = "C:\Windows\system32\scrrun.dll"
$ACTIVEX_DATA_OBJECTS_LIB = "C:\$programFiles\Common Files\System\ado\msado15.dll"

$buildRefs = @{
    "src" = @("$VBA_EXTENSIBILITY_LIB", "$VBA_SCRIPTING_LIB", "$ACTIVEX_DATA_OBJECTS_LIB");
    "test" = @((Join-Path $scriptRoot "VBEXsrc.$ext"))
}

ForEach ($build In $buildRefs.Keys) {
    $path = (Join-Path $scriptRoot "VBEX$build.$ext")
    $files = (Get-ChildItem (Join-Path $scriptRoot $build)) | % { $_.FullName } # v3 and greater this would be just .FullName
    $refs = $buildRefs[$build]
    & "$buildScript" "$path" $files $refs
}

& (Join-Path "$scriptRoot" "Close-All-Of.ps1") "$officeApp"
