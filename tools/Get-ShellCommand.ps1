<#
.SYNOPSIS
data.jsに載っていないshellコマンドの情報を返す
#>

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
	Select-Object -Property `
		@{ Name = 'Guid'; Expression = { $_.PSChildName } },
		@{ Name = 'Name'; Expression = { "shell:$($_.GetValue('Name'))" } },
		@{ Name = 'Category'; Expression = { $categories[$_.GetValue('Category')] } },
		@{ Name = 'PreCreate'; Expression = { $_.GetValue('PreCreate') } },
		@{ Name = 'ParsingName'; Expression = { $_.GetValue('ParsingName') } },
		@{ Name = 'ParentFolder'; Expression = { $_.GetValue('ParentFolder') } },
		@{ Name = 'RelativePath'; Expression = { $_.GetValue('RelativePath') } }
