Param(
    [bool]$uninstall=$false
)

# Global definitions
$wtProfilesPath = "$env:LocalAppData\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"
$customConfigPath = "$PSScriptRoot\config.json"
$resourcePath = "$env:LOCALAPPDATA\WindowsTerminalContextIcons\"
$contextMenuIcoName = "terminal.ico"
$cmdIcoFileName = "cmd.ico"
$wslIcoFileName = "linux.ico"
$psIcoFileName = "powershell.ico"
$psCoreIcoFileName = "powershell-core.ico"
$azureCoreIcoFileName = "azure.ico"
$unknownIcoFileName = "unknown.ico"
$menuRegID = "WindowsTerminal"
$contextMenuLabel = "Open Windows Terminal here"
$contextMenuRegPath = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\$menuRegID"
$contextBGMenuRegPath = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\$menuRegID"
$subMenuRegRelativePath = "Directory\ContextMenus\$menuRegID"
$subMenuRegRoot = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\ContextMenus\$menuRegID"
$subMenuRegPath = "$subMenuRegRoot\shell\"

function Add-SubmenuReg ($regPath, $label, $iconPath, $command) {
    $cmdRegPath = "$regPath\command"
    [void](New-Item -Force -Path $regPath)
    [void](New-Item -Force -Path $cmdRegPath)
    [void](New-ItemProperty -Path $regPath -Name "MUIVerb" -PropertyType String -Value $label)
    [void](New-ItemProperty -Path $cmdRegPath -Name "(default)" -PropertyType String -Value $command)
    [void](New-ItemProperty -Path $regPath -Name "Icon" -PropertyType String -Value $iconPath)
}

# Clear register
if((Test-Path -Path $contextMenuRegPath)) {
    # If reg has existed
    Remove-Item -Recurse -Force -Path $contextMenuRegPath
    Write-Host "Clear reg $contextMenuRegPath"
}

if((Test-Path -Path $contextBGMenuRegPath)) {
    Remove-Item -Recurse -Force -Path $contextBGMenuRegPath
    Write-Host "Clear reg $contextBGMenuRegPath"
}

if((Test-Path -Path $subMenuRegRoot)) {
    Remove-Item -Recurse -Force -Path $subMenuRegRoot
    Write-Host "Clear reg $subMenuRegRoot"
}

if((Test-Path -Path $resourcePath)) {
    Remove-Item -Recurse -Force -Path $resourcePath
    Write-Host "Clear icon content folder $resourcePath"
}

if($uninstall) {
    Exit
}

# Setup icons
[void](New-Item -Path $resourcePath -ItemType Directory)
[void](Copy-Item -Path "$PSScriptRoot\icons\*.ico" -Destination $resourcePath)
Write-Output "Copy icons => $resourcePath"

# Load the custom config
if((Test-Path -Path $customConfigPath)) {
    $rawConfig = (Get-Content $customConfigPath -Encoding UTF8) -replace '^\s*\/\/.*' | Out-String
    $config = (ConvertFrom-Json -InputObject $rawConfig)
}

# Setup First layer context menu
[void](New-Item -Force -Path $contextMenuRegPath)
[void](New-ItemProperty -Path $contextMenuRegPath -Name ExtendedSubCommandsKey -PropertyType String -Value $subMenuRegRelativePath)
[void](New-ItemProperty -Path $contextMenuRegPath -Name Icon -PropertyType String -Value $resourcePath$contextMenuIcoName)
[void](New-ItemProperty -Path $contextMenuRegPath -Name MUIVerb -PropertyType String -Value $contextMenuLabel)
if($config.global.extended) {
    [void](New-ItemProperty -Path $contextMenuRegPath -Name Extended -PropertyType String)
}
Write-Host "Add top layer menu (shell) => $contextMenuRegPath"

[void](New-Item -Force -Path $contextBGMenuRegPath)
[void](New-ItemProperty -Path $contextBGMenuRegPath -Name ExtendedSubCommandsKey -PropertyType String -Value $subMenuRegRelativePath)
[void](New-ItemProperty -Path $contextBGMenuRegPath -Name Icon -PropertyType String -Value $resourcePath$contextMenuIcoName)
[void](New-ItemProperty -Path $contextBGMenuRegPath -Name MUIVerb -PropertyType String -Value $contextMenuLabel)
if($config.global.extended) {
    [void](New-ItemProperty -Path $contextBGMenuRegPath -Name Extended -PropertyType String)
}
Write-Host "Add top layer menu (background) => $contextMenuRegPath"

# Get Windows terminal profile
$rawContent = (Get-Content $wtProfilesPath -Encoding UTF8) -replace '^\s*\/\/.*' | Out-String
$json = (ConvertFrom-Json -InputObject $rawContent);

$profiles = $null;

if($json.profiles.list){
    Write-Host "Working with the new profiles style"
    $profiles = $json.profiles.list;
} else{
    Write-Host "Working with the old profiles style"
    $profiles = $json.profiles;
}

$profileSortOrder = 0

# Setup each profile item
$profiles | ForEach-Object {    
    $profileSortOrder += 1
    $profileSortOrderString = "{0:00}" -f $profileSortOrder 
    $profileName = $_.name
    $guid = $_.guid
    $configEntry = $config.profiles.$guid
        
    $leagaleName = $profileName -replace '[ \r\n\t]', '-'
    $subItemRegPath = "$subMenuRegPath$profileSortOrderString$leagaleName"
    $subItemAdminRegPath = "$subItemRegPath-Admin"

    if ($configEntry.hidden -eq $null) {
        $isHidden = $_.hidden
    } else {
        $isHidden = $configEntry.hidden
    }
    $commandLine = $_.commandline
    $source = $_.source
    $icoPath = ""

    # Final values
    $iconPath_f = ""
    $label_f = ""
    $labelAdmin_f = ""
    $command_f = ""
    $commandAdmin_f = ""

    if ($isHidden -eq $false) {

        # Decide label
        if ($configEntry.label) {
            $label_f = $configEntry.label
        }
        else {
            $label_f = $profileName
        }
        $labelAdmin_f = "$label_f (Admin)"
        
        $command_f = "`"$env:LOCALAPPDATA\Microsoft\WindowsApps\wt.exe`" -p `"$profileName`" -d `"%V\.`""
        $commandAdmin_f = "powershell -WindowStyle hidden -Command `"Start-Process powershell -WindowStyle hidden -Verb RunAs -ArgumentList `"`"`"`"-Command $env:LOCALAPPDATA\Microsoft\WindowsApps\wt.exe -p '$profileName' -d '%V\.'`"`"`"`""
        
        if($configEntry.icon){
            $useFullPath = [System.IO.Path]::IsPathRooted($configEntry.icon);
            $tmpIconPath = $configEntry.icon;            
            $icoPath = If (!$useFullPath) {"$resourcePath$tmpIconPath"} Else { "$tmpIconPath" }
        }
        elseif ($_.icon) {
            $icoPath = $_.icon
        }
        elseif(($commandLine -match "^cmd\.exe\s?.*")) {
            $icoPath = "$cmdIcoFileName"
        }
        elseif (($commandLine -match "^powershell\.exe\s?.*")) {
            $icoPath = "$psIcoFileName"
        }
        elseif ($source -eq "Windows.Terminal.Wsl") {
            $icoPath = "$wslIcoFileName"
        }
        elseif ($source -eq "Windows.Terminal.PowershellCore") {
            $icoPath = "$psCoreIcoFileName"
        }
        elseif ($source -eq "Windows.Terminal.Azure") {
            $icoPath = "$azureCoreIcoFileName"
        }else{
            # Unhandled Icon
            $icoPath = "$unknownIcoFileName"
            Write-Host "No icon found, using unknown.ico instead"
        }

        if($icoPath -ne "") {
            $iconPath_f = If ($configEntry.icon -or $_.icon) { "$icoPath" } Else { "$resourcePath$icoPath" }
        }

        Write-Host "Add new entry $profileName => $subItemRegPath"

        Add-SubmenuReg -regPath:$subItemRegPath -label:$label_f -iconPath:$iconPath_f -command:$command_f

        if ($configEntry.showRunAs) {
            Add-SubmenuReg -regPath:$subItemAdminRegPath -label:$labelAdmin_f -iconPath:$iconPath_f -command:$commandAdmin_f
        }
    }else{
        Write-Host "Skip entry $profileName => $subItemRegPath"
    }
}

# SIG # Begin signature block
# MIIFqQYJKoZIhvcNAQcCoIIFmjCCBZYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUSCHOMKDbS/rgmYX+eJiwkVjw
# XjygggM8MIIDODCCAiCgAwIBAgIQGAou2uZoRbVEqZJKHUO37TANBgkqhkiG9w0B
# AQsFADAiMSAwHgYDVQQDDBdtaWhpcnJhYmFkZUBvdXRsb29rLmNvbTAeFw0yMDEw
# MjgxNjM4MTFaFw0yMTEwMjgxNjU4MTFaMCIxIDAeBgNVBAMMF21paGlycmFiYWRl
# QG91dGxvb2suY29tMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAzw7I
# 8sjnYKOLTQX9s3VbSlzJxo4YTxy4y1D7xQ5K8BPW1TmKWOiS0TpbEtAzsc1eV2XL
# WGfRk49zDzcEcIyQz6anTXx72jMXpPvPOVqkxFSnMkJ78sZho1Ls21CLzWd0IHhw
# q+Vy6C9Kbx65hwt8vvXdNj1pZRWd7ccJtnFW/etyXPLniaNMvWFJqkOUCrgRMfIT
# KBzBXE9rzegtTlv8PueDSRPaF64ABy06pTFBSschmUgtJeAijq0AVIAxbBgmz0so
# EzbY1HvY7yvT6N0ZEgbqVq8oOqP1t7AZzcrIL35dVmLhS8GVBpx6aWE97gIo9APg
# tvTcrvJO/Ua8Frd7GQIDAQABo2owaDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwIgYDVR0RBBswGYIXbWloaXJyYWJhZGVAb3V0bG9vay5jb20w
# HQYDVR0OBBYEFLW4EjuvisCwJO8nWzmKLSkqfNiHMA0GCSqGSIb3DQEBCwUAA4IB
# AQC18hH/qbjIjRNWXUUTrASt0HtooaEYYyLvKUnfDWCmdR/ClupjW2JFvOiZsnR1
# PoE6yt/17JYqO/Uibhq2B0XZkBPyi9+4u4Sx8ebeDT30GD1yJoPvOK8KpTTY6ldM
# cyPC7zbu7TXDBQwn0nPc4yUTOb6UKrVzq3zJOdGsu9H5m0RSEdY7raDbWUcRibjl
# YPd2+jGmqNof+UxY05kn8VboFxhs8+UzHkYwZ2Y4kogz/R+syoQyMJQ2xRU7l+Sh
# IW/J0WQbl0djC3fSjyxFhzqzhf08rw5jWMIeWTWt4NofkcaXzwUn8MSypkHc+XE6
# txsz53KUHjsAGdqfcS/gwHYtMYIB1zCCAdMCAQEwNjAiMSAwHgYDVQQDDBdtaWhp
# cnJhYmFkZUBvdXRsb29rLmNvbQIQGAou2uZoRbVEqZJKHUO37TAJBgUrDgMCGgUA
# oHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYB
# BAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0B
# CQQxFgQUGv0qfnwVjQSK17RDQl4PDYV6ho8wDQYJKoZIhvcNAQEBBQAEggEApEi6
# As4GbwPD8iwOAg4QMauvf/ysIP09o5umX2un7jhiAG7jKXzo/p9crOoSguBEsD0+
# A4OyyZ3NHv/XGM7q92zZXR2WqgFpM9h449b+S/CEWsis01RzuDsLQ5nWP2RIAtsT
# If5qBRDOaCOdJz8PXNZGTPBFvkEadiz4OT21IywtFmMkXu69xsu4jBCxsb1Gj6td
# 3xAuYqP2YQdwD2cWIqMo3XElrvmbjfRXG4hN578V++2p0SvfYf1YZQJ9FR5y+ar1
# rvgpxEYWJpdhsa3CPnJ8shK6vhBW6od4lM1m5QArTa0LCMicHy3NOSjBEJ/dJvCO
# Wi2Ryo2Ids1gi18s8w==
# SIG # End signature block
