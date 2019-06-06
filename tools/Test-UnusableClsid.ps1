<#
.SYNOPSIS
data.jsに載っている特殊フォルダーとして使えないCLSIDがまだレジストリに存在しているか調べる
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest

$dataFile = (Resolve-Path "$($MyInvocation.MyCommand.Path)\..\..\src\modules\data.js").Path
$dataText = @(Get-Content -LiteralPath $dataFile) -join "`n"
$unusableText = ($dataText -csplit 'category: "Unusable"', 2, 'SimpleMatch')[1]

([regex]'\{\w{8}-\w{4}-\w{4}-\w{4}-\w{12}\}').Matches($unusableText) |
	ForEach-Object {
		$clsid = $_.Value
		$path = "Microsoft.PowerShell.Core\Registry::HKEY_CLASSES_ROOT\CLSID\$clsid"
		$exists = Test-path $path
		
		Write-Output (
			New-Object psobject -Property @{
				'Clsid' = $clsid
				'Exists' = $exists
				'Name' = if ($exists) { (Get-Item $path).GetValue('') }
			} | Select-Object Clsid, Exists, Name
		)
	}
