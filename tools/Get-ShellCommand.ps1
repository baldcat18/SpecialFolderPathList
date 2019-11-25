<#
.SYNOPSIS
data.jsに載っていないshellコマンドの情報を返す
#>

#Requires -Version 4.0

[CmdletBinding()]
param([switch]$All)

Set-StrictMode -Version Latest

function const ([string]$name, $value) {
	New-Variable -Name $name -Value $value -Option Constant -Scope 1
}

if (!$All) {
	$dataFile = (Resolve-Path "$($MyInvocation.MyCommand.Path)\..\..\src\modules\data.js").Path
	$dataText = @(Get-Content -LiteralPath $dataFile) -join "`n"
}

$categories = @{ 1 = 'Virtual'; 2 = 'Fixed'; 3 = 'Common'; 4 = 'Per-user' }

Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions' |
	Where-Object { $All -or !$dataText.Contains("`"shell:$($_.GetValue('Name'))`"") } |
	ForEach-Object {
		[pscustomobject]@{
			Guid = $_.PSChildName
			Name = "shell:$($_.GetValue('Name'))"
			Category = $categories[$_.GetValue('Category')]
			PreCreate = $_.GetValue('PreCreate')
			ParsingName = $_.GetValue('ParsingName')
			ParentFolder = $_.GetValue('ParentFolder')
			RelativePath = $_.GetValue('RelativePath')
		}
	}
