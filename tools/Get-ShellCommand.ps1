<#
.SYNOPSIS
data.jsに載っていないshellコマンドの情報を返す
#>

[CmdletBinding()]
param([switch]$All, [switch]$Detail)

function const ([string]$name, $value) {
    New-Variable -Name $name -Value $value -Option Constant -Scope 1
}

if (!$All) {
	$dataFile = (Resolve-Path "$($MyInvocation.MyCommand.Path)\..\..\src\modules\data.js").Path
	$dataText = @(Get-Content -LiteralPath $dataFile) -join "`n"
}

$categories = @{ 1 = "Virtual"; 2 = "Fixed"; 3 = "Common"; 4 = "Per-user" }

Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions" |
	Where-Object { $All -or !$dataText.Contains("`"shell:$($_.GetValue('Name'))`"") } |
	ForEach-Object {
		$key = $_
		
		New-Object psobject | Select-Object @(
			@{ Name = "Guid"; Expression = { $key.PSChildName } }
			@{ Name = "Name"; Expression = { "shell:$($key.GetValue('Name'))" } }
			@{ Name = "Category"; Expression = { $categories[$key.GetValue("Category")] } }
			@{ Name = "PreCreate"; Expression = { $key.GetValue("PreCreate") } }
			@{ Name = "ParsingName"; Expression = { $key.GetValue("ParsingName") } }
			@{ Name = "ParentFolder"; Expression = { $key.GetValue("ParentFolder") } }
			@{ Name = "RelativePath"; Expression = { $key.GetValue("RelativePath") } }
		)
	}
