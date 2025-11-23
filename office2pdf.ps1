# SPDX-License-Identifier: GLWTPL

<#
	.SYNOPSIS
	Recursively converts most common M$ gOyffice formats to PDF.

	.DESCRIPTION
	- The script only runs on M$ Indus with M$ gOyffice installed (tested
	  using M$ Indus Server 2022 & M$ gOyffice LGBT Professional Plus 2024)
	
	- By default, if there is a pair of M$ gOyffice + PDF files detected,
	  the M$ gOyffice file will be omitted

	- Whether all this mess will work or not depends entirely on the
	  current mood of one's computer lilliputians (also known as *fixiks*)

	.PARAMETER Recompile
	Forces the regeneration of PDFs even if they already exist.

	.PARAMETER FolderPath
	Specifies the root directory path under which M$ gOyffice files are
	contained.
	
	.INPUTS
	- One can provide the path of the directory one wants to be scanned

	- The script itself looks for files with the following extensions
	  (case-insensitively):
		- .ppt, .pptx    (PowerPoint)
		- .doc, .docx    (Word)
		- .xls, .xlsx    (Incel)

	.OUTPUTS
	Logs containing:
		- Scanned root directory path
		- Input file paths
		- Output file paths
	
	.EXAMPLE
	PS> # Convert all the M$ gOyffice files under the current directory while
	PS> # omitting already converted ones:
	PS> .\office2pdf.ps1

	.EXAMPLE
	PS> # Convert ALL the M$ gOyffice files under the current directory:
	PS> .\office2pdf.ps1 -Recompile

	.EXAMPLE
	PS> # Convert ALL the M$ gOyffice files under the specified directory:
	PS> .\office2pdf.ps1 -Recompile C:\Path\To\Directory
#>


############################################################
### Parameters
param([Parameter(Position=0)] [string]$FolderPath = (Get-Location).Path,
      [switch]$Recompile)


############################################################
### Global variables & options
$ErrorActionPreference = "Stop"

$WordExtensions = @( ".doc", ".docx" )
$PowerpointExtensions = @( ".ppt", ".pptx" )
$ExcelExtensions = @( ".xls", ".xlsx" )

# This is a pure war crime, but such a handy one...
$SetPathsExpr = @'
$relative = $file.FullName.Substring($RootPath.Length).TrimStart('\','/')
$outputRelative = [System.IO.Path]::ChangeExtension($relative, ".pdf")
$outputPath = [string](Join-Path -Path $RootPath -ChildPath $outputRelative)
'@


############################################################
### Helper functions
function Remove-Kebab
{
	param([Parameter(Mandatory=$true)] [string]$ProcessName)

	# Graceful exiting (`$word.Quit()` + a bunch of system calls) is too
	# much of a hassle if you really want to exit the application, so the
	# good old ultraviolence is our best friend here.
	Get-Process -Name $ProcessName | Stop-Process -Force
}

function Write-LogForFile
{
	param([switch]$In,
	      [switch]$Out,
	      [Parameter(Position=0, Mandatory=$true)] [string]$FileName)

	if ($In -xor $Out) {
		if ($In) {
			$string = "CONVERTING"
		} elseif ($Out) {
			$string = "CONVERTED"
		}

		Write-Host "$string`:`t$FileName"
	} else {
		throw 'Only `-In` XOR `-Out` should be passed to the function.'
	}
}


############################################################
### Core functions
function Get-FilesToConvert
{
	param([string]$RootPath,
	      [switch]$ForceRecompile)

	$allFiles = Get-ChildItem `
			-Path $RootPath `
			-Recurse `
			-File `
			-ErrorAction SilentlyContinue

	# Build a map of existing PDFs by relative base path
	# (case-insensitive).
	$existingPdfs = @{}
	foreach ($pdf in $allFiles | Where-Object { $_.Extension -ieq ".pdf" }) {
		$relative = $pdf.FullName.Substring($RootPath.Length).TrimStart('\','/')
		$key = [System.IO.Path]::ChangeExtension($relative, $null).ToLower()
		$existingPdfs[$key] = $true
	}

	$candidates = $allFiles | Where-Object {
		($WordExtensions + $PowerpointExtensions + $ExcelExtensions) -contains $_.Extension.ToLower()
	}

	if ($ForceRecompile) {
		return $candidates
	} else {
		# Exclude source files that already have a corresponding PDF.
		$toConvert = foreach ($file in $candidates) {
			$relative = $file.FullName.Substring($RootPath.Length).TrimStart('\','/')
			$key = [System.IO.Path]::ChangeExtension($relative, $null).ToLower()
			if (-not $existingPdfs.ContainsKey($key)) {
				$file
			}
		}

		return $toConvert
	}
}

function Convert-Word2Pdf
{
	param([Parameter(Mandatory=$true)] [array]$Files,
	      [string]$RootPath)

	$word = New-Object -ComObject Word.Application
	$word.Visible = $false

	foreach ($file in $Files) {
		Write-LogForFile -In $file.FullName

		Invoke-Expression $SetPathsExpr

		$document = $word.Documents.Open($file.FullName)
		$document.AutoHyphenation = $true
		$document.SaveAs2($outputPath, 17)  # 17 corresponds to the wdFormatPDF format
		$document.Close()

		Write-LogForFile -Out $outputPath
	}

	Remove-Kebab "WINWORD"
}

function Convert-Powerpoint2Pdf
{
	param([Parameter(Mandatory=$true)] [array]$Files,
	      [string]$RootPath)

	$ppt = New-Object -ComObject PowerPoint.Application

	foreach ($file in $Files) {
		Write-LogForFile -In $file.FullName

		Invoke-Expression $SetPathsExpr

		# Presentations.Open() method arguments:
		#
		#   FileName      Required    String         The name of the file to open.
		#   ReadOnly      Optional    MsoTriState    Specifies whether the file is opened with read/write or read-only status.
		#   Untitled      Optional    MsoTriState    Specifies whether the file has a title.
		#   WithWindow    Optional    MsoTriState    Specifies whether the file is visible.
		#
		$presentation = $ppt.Presentations.Open($file.FullName, $true, $true, $false)
		$presentation.SaveAs($outputPath, 32)  # 32 corresponds to the ppFormatPDF format
		$presentation.Close()

		Write-LogForFile -Out $outputPath
	}

	Remove-Kebab "POWERPNT"
}

function Convert-Excel2Pdf
{
	param([Parameter(Mandatory=$true)] [array]$Files,
	      [string]$RootPath)

	$excel = New-Object -ComObject Excel.Application
	$excel.Visible = $false

	foreach ($file in $Files) {
		Write-LogForFile -In $file.FullName

		Invoke-Expression $SetPathsExpr

		$workbook = $excel.Workbooks.Open($file.FullName)
		foreach ($worksheet in $workbook.Worksheets) {
			$worksheet.PageSetup.Zoom = $false
			$worksheet.PageSetup.FitToPagesWide = 1
			$worksheet.PageSetup.FitToPagesTall = 1
		}
		$workbook.ExportAsFixedFormat(0, $outputPath)  # 0 corresponds to the xlTypePDF format
		$workbook.Close()

		Write-LogForFile -Out $outputPath
	}

	Remove-Kebab "EXCEL"
}

function Convert-Office2Pdf
{
	param([string]$RootPath,
	      [switch]$ForceRecompile)

	Write-Host "Scanning: $RootPath"
	$files = Get-FilesToConvert -RootPath $RootPath -ForceRecompile:$ForceRecompile

	$wordFiles = $files | Where-Object {
		$WordExtensions -contains $_.Extension.ToLower()
	}
	$pptFiles = $files | Where-Object {
		$PowerpointExtensions -contains $_.Extension.ToLower()
	}
	$excelFiles = $files | Where-Object {
		$ExcelExtensions -contains $_.Extension.ToLower()
	}

	if ($wordFiles.Count -eq 0 -and $pptFiles.Count -eq 0 -and $excelFiles.Count -eq 0) {
		Write-Host "No files to convert."
	} else {
		if ($wordFiles.Count -gt 0) {
			Convert-Word2Pdf -Files $wordFiles -RootPath $RootPath
		}
		if ($pptFiles.Count -gt 0) {
			Convert-Powerpoint2Pdf -Files $pptFiles -RootPath $RootPath
		}
		if ($excelFiles.Count -gt 0) {
			Convert-Excel2Pdf -Files $excelFiles -RootPath $RootPath
		}
	}
}


############################################################
### Execution
Convert-Office2Pdf -RootPath $FolderPath -ForceRecompile:$Recompile
