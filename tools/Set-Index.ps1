<#
.SYNOPSIS
data.jsのSpecialFolder.getObject()内で参照しているindex番号を振りなおす
#>

#Requires -Version 4.0

[CmdletBinding()]
param()

Set-StrictMode -Version Latest

$modules = (Resolve-Path "$PSScriptRoot\..\src\modules").Path
$dataFile = "$modules\data.js"
$tmpFile = "$modules\data.js.tmp"

$index = 0

$encoding = if ($PSVersionTable['PSVersion'] -ge '6.0') { 'utf8BOM' } else { 'UTF8' }

Get-Content -LiteralPath $dataFile |
	ForEach-Object {
		if ($_ -cmatch '^\t\t\tcase \d+:$') { "`t`t`tcase $(($index++)):" } else { $_ }
	} |
	Out-File -LiteralPath $tmpFile -Encoding $encoding
[System.IO.File]::Replace($tmpFile, $dataFile, [nullstring]::Value)
