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

$scriptRoot = if ($PSVersionTable.PSVersion.Major -ge 3) {
    $PSScriptRoot
} else {
    Split-Path $MyInvocation.MyCommand.Path -Parent
}
$display = (Join-Path $scriptRoot "Build-Display.ps1")

function Export {
        Write-Host "Will Export $sourceFile to $outDest`:"
	$officeCOM = (OfficeComFromFileName $sourceFile)
	$fileCOM = (OpenMSOfficeFile $officeCOM $sourceFile)
	$prjCOM = ($fileCOM.VBProject)
	ExportModules $prjCOM $outDest
	$officeCOM.Quit()
        Start-Sleep -seconds 1
        Write-Host "Removing $sourceFile" -ForeGround Red
        Remove-Item $sourceFile
}

function OfficeComFromFileName([String] $fileName) {
	$fileExt = [System.IO.Path]::GetExtension($fileName)
	$officeCOM = switch -wildcard ($fileExt.ToLower()) {
        ".xl*" {New-Object -ComObject Excel.Application; break}
        ".ac*" {throw "Access is not yet supported"; break} #{New-Object -ComObject Acces.Application; break}
        default {throw "$fileName is not a supported office file."}
    }
    return $officeCOM
}

function OpenMSOfficeFile($officeCOM, [String] $filePath) {
        Write-Host "Opening $filePath with Office Application."
	$fileCOM = ($officeCOM.Workbooks.Open($filePath))
	return $fileCOM
}

function ExportModules($prjCOM, [String] $outDest) {
    
	$vbComps = ($prjCOM.VBComponents)
        $count = $vbComps.count
        & "$display" 1 "Exporting $count modules to $outDest`:"
	ForEach ($component in $vbComps) {
		$compFileExt = (GetComponentExt($component))
		if ($compFileExt -ne "") {
			$compFileName = $component.name + $compFileExt
			$exportPath = (Join-Path $outDest $compFileName)
                        & "$display" 3 "$compFileName"
			$component.Export($exportPath)
                }
	}
}

function GetComponentExt($component) {
	$compExt = switch ($component.Type) {
		1 {".bas"}
		2 {".cls"}
		# form
		default {""}
	}
	return $compExt
}

Export # entry
