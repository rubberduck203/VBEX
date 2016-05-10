#

<#
#>
Param (
    [Int] $level,
    [String] $msg
)

function Build-Display {
    Switch ($level) {
        1 {
            Write-Host "=> " -ForeGround Blue -noNewLine
            break
        }
        2 {
            Write-Host "==> " -ForeGround Green -noNewLine
            break
        }
        default {
            $icon = "> ".PadLeft($level + 2, "-")
            Write-Host $icon  -ForeGround Yellow -noNewLine 
            break
        }
    }
    Write-Host $msg
}

Build-Display
