Function Update-MSExcelFindReplaceTextDocument {
	<#

		.SYNOPSIS
			When provided with an open Microsoft Excel document, executes a set of text only FindText/ReplaceWith operations throughout the whole document.

		.DESCRIPTION	
			When provided with an open Microsoft Excel document (Excel.Application's Workbooks.Open), executes a set of text only FindText/ReplaceWith operations throughout the whole document.  

			Microsoft Excel's Range.Replace operation performs a simple text match. There is no support wildcard or regular expressions [RegEx]. Formatting of the FindText is preserved.  
			To reduce the chance on unintended replacements, surround keywords with a marker.  Parenthesis are safe when used with Microsoft's Word, Excel, Outlook and PowerPoint documents.
			For example: (CompanyName)
			
		.OUTPUTS
			The output of this function is the modifications executed on the open document.
			The returned value from this function is the number of replacements made.  
			
		.PARAMETER Path [Microsoft.Office.Interop.Excel.Workbooks]
			An open Microsoft Excel document (Excel.Workbooks.Open).  If FindText is found this document will be modified.  
			
		.PARAMETER FindReplacePath String
			A hash table @{ [String] FindText, [String] ReplaceWith } pairs.  Each FindText ReplaceWith operation will be executed througout the whole document.  

		.EXAMPLE
			Description
			-----------
				This example opens a Mircosoft Excel application,
				opens a MS Excel document named '.\MySource.xlsx' in read only mode,
				creates a substitution hash table with 2 entries,
				and then calls this function.
				After the document is saved to another file name,
				and closed.  
			
				$excelApp = New-Object -ComObject Excel.Application
				$excelApp.Visible = $FALSE
				$excelApp.DisplayAlerts = $FALSE
				$document = $excelApp.Workbooks.Open( '.\MySource.xlsx', 0, $TRUE ) # Filename, UpdateLinks, ReadOnly
				
				$findReplaceTable = @{ 'INCORRECT' = 'correct'; '(Field)' = 'MyFieldValue' }
				
				Update-MSExcelFindReplaceTextDocument -Document $document -FindReplaceTable $findReplaceTable
				
				$document.Close( $TRUE, $outFilePathName ) # SaveChanges, Filename
				$excelApp.Quit()

		.NOTE
			Author: Terry E Dow
			Creation Date: 2018-03-02
			Last Modified: 2019-03-25

			Warning from Microsoft:
				Considerations for server-side Automation of Office https://support.microsoft.com/en-us/help/257757/considerations-for-server-side-automation-of-office
			
			Reference:
				https://codereview.stackexchange.com/questions/174455/powershell-script-to-find-and-replace-in-word-document-including-header-footer
				https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.find.execute
				https://learn-powershell.net/2014/12/31/beginning-with-powershell-and-word/
				https://wordmvp.com/FAQs/MacrosVBA/FindReplaceAllWithVBA.htm
				https://wordribbon.tips.net/T011489_Including_Headers_and_Footers_when_Selecting_All.html				
	#>
	[CmdletBinding(
		#SupportsShouldProcess = $TRUE # Enable support for -WhatIf by invoked destructive cmdlets.
	)]
	#[System.Diagnostics.DebuggerHidden()]
	Param(

		[Parameter(
		ValueFromPipeline=$TRUE,
		ValueFromPipelineByPropertyName=$TRUE,
		Position=0)]
		#[Microsoft.Office.Interop.Excel.Workbooks] # Cannot convert the "System.__ComObject" value of type "System.__ComObject#{000208da-0000-0000-c000-000000000046}" to type "Microsoft.Office.Interop.Excel.Workbooks".
		$Document = $NULL,
		
		[Parameter(
		ValueFromPipeline=$TRUE,
		ValueFromPipelineByPropertyName=$TRUE )]
		[Hashtable] $FindReplaceTable = $NULL

	)

	Begin {
	
		#Requires -version 3
		Set-StrictMode -Version Latest

		# Detect cmdlet common parameters.
		$cmdletBoundParameters = $PSCmdlet.MyInvocation.BoundParameters
		$Debug = If ( $cmdletBoundParameters.ContainsKey('Debug') ) { $cmdletBoundParameters['Debug'] } Else { $FALSE }
		# Replace default -Debug preference from 'Inquire' to 'Continue'.
		If ( $DebugPreference -Eq 'Inquire' ) { $DebugPreference = 'Continue' }
		$Verbose = If ( $cmdletBoundParameters.ContainsKey('Verbose') ) { $cmdletBoundParameters['Verbose'] } Else { $FALSE }
		$WhatIf = If ( $cmdletBoundParameters.ContainsKey('WhatIf') ) { $cmdletBoundParameters['WhatIf'] } Else { $FALSE }
		Remove-Variable -Name cmdletBoundParameters -WhatIf:$FALSE
			
		# https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.range.replace
		# What, Replacement, LookAt, SearchOrder, MatchCase, MatchByte, SearchFormat, ReplaceFormat
		# $replaceWhat, $replaceReplacement, $replaceLookAt, $replaceSearchOrder, $replaceMatchCase, $replaceMatchByte, $replaceSearchFormat, $replaceReplaceFormat
		#$replaceWhat = ''
		#$replaceReplacement = ''
		$replaceLookAt = [Microsoft.Office.Interop.Excel.XlLookAt]::xlPart
		Write-Debug "`$replaceLookAt:,$replaceLookAt"
		$replaceSearchOrder = [Microsoft.Office.Interop.Excel.XlSearchOrder]::xlByRows 
		Write-Debug "`$replaceSearchOrder:,$replaceSearchOrder"
		$replaceMatchCase = $FALSE
		$replaceMatchByte = $FALSE
		$replaceSearchFormat = 0 #?
		$replaceReplaceFormat = 0 #?
						
	}
	
	Process { 
		Write-Debug "(`$FindReplaceTable).Count:,$(($FindReplaceTable).Count)"
	
		# Initialize metrics.
		$isUpdated = $FALSE
				
		# Loop through each of the Document's StoryRanges
		$Document.Worksheets | 
			ForEach-Object {
				Write-Verbose 'Worksheets'
			
				# [Microsoft.Office.Interop.Excel.WorksheetClass] 
				$worksheet = $PSItem
				$range = $worksheet.UsedRange
				
				# For each substitution table entry, execute this worksheet range's Replace What/Replacement method.
				$FindReplaceTable.GetEnumerator() | 
					ForEach-Object {
						Write-Verbose "FindText:,$($PSItem.Key)"	
						$findTextEscaped = [System.Text.RegularExpressions.Regex]::Escape($PSItem.Key)
						Write-Debug "findTextEscaped:,$findTextEscaped"	

						# Update document's worksheet name with this FindText.
						Write-Verbose "`$worksheet.Name:,$($worksheet.Name)"
						If ( $worksheet.Name -Match $findTextEscaped ) { 
							$worksheet.Name = $worksheet.Name -Replace $findTextEscaped, $PSItem.Value 
							$isUpdated = $TRUE
							Write-Debug "`$worksheet.Name updated to:,$($worksheet.Name)"
						}						
						
						# Update document's worksheet used range with this FindText.  
						If ( $range.Replace( $PSItem.Key, $PSItem.Value, $replaceLookAt, $replaceSearchOrder, $replaceMatchCase ) ) { # , $replaceMatchByte, $replaceSearchFormat, $replaceReplaceFormat )
							$isUpdated = $TRUE
							Write-Debug "`$worksheet.UsedRange updated"
						}
					}			
			}

	}
	
	End {
		Write-Debug "`$isUpdated:,$isUpdated"
		Write-Output $isUpdated
	}
}
