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
	$officeCOM = (OfficeComFromFileName $sourceFile)
	$fileCOM = (OpenMSOfficeFile $officeCOM $sourceFile)
	$prjCOM = ($fileCOM.VBProject)
	ExportModules $prjCOM $outDest
	$officeCOM.Quit()
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
	$fileCOM = ($officeCOM.Workbooks.Open($filePath))
	return $fileCOM
}
function ExportModules($prjCOM, [String] $outDest) {
    
	$vbComps = ($prjCOM.VBComponents)
	ForEach ($component in $vbComps) {
		$compFileExt = (GetComponentExt($component))
		if ($compFileExt -ne "") {
			$compFileName = $component.name + $compFileExt
			$exportPath = (Join-Path $outDest $compFileName)
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
