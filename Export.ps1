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
        Write-Host "Exporting $count modules:"
	ForEach ($component in $vbComps) {
		$compFileExt = (GetComponentExt($component))
		if ($compFileExt -ne "") {
			$compFileName = $component.name + $compFileExt
			$exportPath = (Join-Path $outDest $compFileName)
                        Write-Host "`t $compFileName => $exportPath"
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
main # entry point
