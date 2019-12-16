<#
.SYNOPSIS
SpecialFolderPathList.wsfの出力をファイルに保存する
#>

#Requires -Version 4.0

[CmdletBinding()]
param()

Set-StrictMode -Version Latest

$osVersion = [System.Environment]::OSVersion.Version.ToString(3)
$cpu = $env:PROCESSOR_ARCHITECTURE
$edition = (Get-Item 'HKLM:/SOFTWARE/Microsoft/Windows NT/CurrentVersion').GetValue('EditionID')
$now = Get-Date -Format yyyyMMdd-HHmmss

$tempFile = [System.IO.Path]::GetTempFileName()
$workDir = (Resolve-Path "$PSScriptRoot\..").Path
$setContentArg = @{
	LiteralPath = "$workDir\tools\$osVersion $cpu $edition $now.txt"
	Encoding = if ($PSVersionTable['PSVersion'] -ge '6.0') { 'utf8BOM' } else { 'UTF8' }
}

cmd.exe /c cscript.exe //u //nologo "$workDir\src\misc\wsh\SpecialFolderPathList.wsf" /bom /dbg /dir /c /d /f /t `> $tempFile
Get-Content -LiteralPath $tempFile | Set-Content @setContentArg
Remove-Item -LiteralPath $tempFile

Push-Location $workDir\tools

$txtFiles = Get-ChildItem "$osVersion $cpu $edition *.txt" | Sort-Object -Property Name -Descending
if (@($txtFiles).Length -ge 2) { fc.exe /n /20 $txtFiles[1].Name $txtFiles[0].Name }

Pop-Location
