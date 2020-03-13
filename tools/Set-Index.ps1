#Requires -Version 4.0

<#
.SYNOPSIS
data.jsのSpecialFolder.getObject()内で参照しているindex番号を振りなおす
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest

$dataFile = (Resolve-Path "$PSScriptRoot\..\src\modules\data.js").Path

$tmpLine = New-Object System.Collections.Generic.List[string] 2048
$index = 0

Get-Content -LiteralPath $dataFile |
	ForEach-Object {
		$tmpLine.Add($(if ($_ -cmatch '^\t\tcase \d+:$') { "`t`tcase " + $index++ + ':' } else { $_ }))
	}

$encoding = if ($PSVersionTable['PSVersion'] -ge '6.0') { 'utf8BOM' } else { 'UTF8' }
Set-Content -LiteralPath $dataFile ($tmpLine -join "`r`n") -Encoding $encoding
