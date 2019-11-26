<#
.SYNOPSIS
data.jsに載っている特殊フォルダーとして使えないCLSIDがまだレジストリに存在しているか調べる
#>

#Requires -Version 4.0

[CmdletBinding()]
param()

Set-StrictMode -Version Latest

$dataFile = (Resolve-Path "$PSScriptRoot\..\src\modules\data.js").Path
$dataText = Get-Content -LiteralPath $dataFile -Raw
$unusableText = ($dataText -csplit 'category: "Unusable"', 2, 'SimpleMatch')[1]

([regex]'\{\w{8}-\w{4}-\w{4}-\w{4}-\w{12}\}').Matches($unusableText) |
	ForEach-Object {
		$clsid = $_.Value
		$path = "Microsoft.PowerShell.Core\Registry::HKEY_CLASSES_ROOT\CLSID\$clsid"
		$exists = Test-path $path
		
		Write-Output ([pscustomobject]@{
			Clsid = $clsid; Exists = $exists; Name = if ($exists) { (Get-Item $path).GetValue('') }
		})
	}
