<#
.SYNOPSIS
data.jsに載っていないshellコマンドの情報を返す
#>

[CmdletBinding()]
param([switch]$All)

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
		
		Write-Output (
			New-Object psobject -Property @{
				"Guid" = $key.PSChildName
				"Name" = "shell:$($key.GetValue('Name'))"
				"Category" = $categories[$key.GetValue("Category")]
				"PreCreate" = $key.GetValue("PreCreate")
				"ParsingName" = $key.GetValue("ParsingName")
				"ParentFolder" = $key.GetValue("ParentFolder")
				"RelativePath" = $key.GetValue("RelativePath")
			} | Select-Object Guid, Name, Category, PreCreate, ParsingName, ParentFolder, RelativePath
		)
	}
