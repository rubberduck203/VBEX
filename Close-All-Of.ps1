#
<#
 .Synopsis
  Checks if the user wants to close all instances of a procss before doing so.

 .Description
  When using scripts that open processes in the background, it is often difficult to 
  ensure that all created instances of those programs are closed.  The simplest solution
  is to just close all processes of that name.  This can close instance that were not
  opened by a script, so this function asks the user if it is ok if we do so.

 .Parameter processName
  Name of the process to close.

 .Example
   # Close all instances of EXCEL.exe
   Close-All-Of "EXCEL"
   I Cannot close the instances of EXCEL I opened.
   May I close all instances of EXCEL? [y/N]
   > Yes
   WARNING! All instances of EXCEL were closed!
#>
Param (
    [String] $processName
)

Function Close-All-Of {
    Write-Host ""
    Write-Host "I cannot close the instances of $processName I opened."
    $closeAll = Read-Host "May I close all instances of $processName" + "? [y/N]"
    $msg = if ($closeAll -like "y*") {
        Stop-Process -Name "$processName"
        "WARNING! All instances of $processName were closed!"
    } else {
        "WARNING! There are unused instances of $processName in the background!"
    }
    Write-Host $msg -ForeGround Red
}

Close-All-Of
