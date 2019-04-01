Function Update-MSPowerPointFindReplaceTextDocument {
	<#

		.SYNOPSIS
			When provided with an open Microsoft PowerPoint document, executes a set of text only FindText/ReplaceWith operations throughout the whole document.

		.DESCRIPTION	
			When provided with an open Microsoft PowerPoint document (PowerPoint.Application's Workbooks.Open), executes a set of text only FindText/ReplaceWith operations throughout the whole document.  

			Microsoft PowerPoint's Range.Replace operation performs a simple text match. There is no support wildcard or regular expressions [RegEx]. Formatting of the FindText is preserved.  
			To reduce the chance on unintended replacements, surround keywords with a marker.  Parenthesis are safe when used with Microsoft's Word, PowerPoint, Outlook and PowerPoint documents.
			For example: (CompanyName)
			
		.OUTPUTS
			The output of this function is the modifications executed on the open document.
			The returned value from this function is the number of replacements made.  
			
		.PARAMETER Path [Microsoft.Office.Interop.PowerPoint.Presentations]
			An open Microsoft PowerPoint document (PowerPoint.Presentations.Open).  If FindText is found this document will be modified.  
			
		.PARAMETER FindReplacePath String
			A hash table @{ [String] FindText, [String] ReplaceWith } pairs.  Each FindText ReplaceWith operation will be executed througout the whole document.  

		.EXAMPLE
			Description
			-----------
				This example opens a Mircosoft PowerPoint application,
				opens a MS PowerPoint document named '.\MySource.pptx' in read only mode,
				creates a substitution hash table with 2 entries,
				and then calls this function.
				After the document is saved to another file name,
				and closed.  
			
				$PowerPointApp = New-Object -ComObject PowerPoint.Application
				$document = $PowerPointApp.Presentations.Open( '.\MySource.pptx', [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse ) # FileName, ReadOnly, Untitled, WithWindow
				
				$findReplaceTable = @{ 'INCORRECT' = 'correct'; '(Field)' = 'MyFieldValue' }
				
				Update-MSPowerPointFindReplaceTextDocument -Document $document -FindReplaceTable $findReplaceTable
				
				$document.SaveAs( $outFilePathName ) # Filename, FileFormat, EmbedTrueTypeFonts
				$document.Close() 
				$PowerPointApp.Quit()

		.NOTE
			Author: Terry E Dow
			Creation Date: 2018-03-30
			Last Modified: 2019-03-30

			Warning from Microsoft:
				Considerations for server-side Automation of Office https://support.microsoft.com/en-us/help/257757/considerations-for-server-side-automation-of-office
			
			Reference:
				
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
		#[Microsoft.Office.Interop.PowerPoint.Workbooks] # Cannot convert the "System.__ComObject" value of type "System.__ComObject#{000208da-0000-0000-c000-000000000046}" to type "Microsoft.Office.Interop.PowerPoint.Workbooks".
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
		
	}
	
	Process { 
		Write-Debug "(`$FindReplaceTable).Count:,$(($FindReplaceTable).Count)"
	
		# Initialize metrics.
		$isUpdated = $FALSE
				
		# Loop through each of the Document's StoryRanges
		$Document.Slides | 
			ForEach-Object {
				Write-Verbose 'Slide'
				$PSItem.Shapes |
					ForEach-Object {
						$shape = $PSItem			
						Write-Verbose ' Shape'
						
						$FindReplaceTable.GetEnumerator() | 
							ForEach-Object {
								Write-Verbose "  FindText:,$($PSItem.Key)"	
								$findTextEscaped = [System.Text.RegularExpressions.Regex]::Escape($PSItem.Key)
								Write-Debug "findTextEscaped:,$findTextEscaped"	
		
								If ( $shape.TextFrame.TextRange.Text -Match $findTextEscaped ) { 
									$shape.TextFrame.TextRange.Text = $shape.TextFrame.TextRange.Text -Replace $findTextEscaped, $PSItem.Value 
									$isUpdated = $TRUE
								}
							}
					}
			}
				
	}
	
	End {
		Write-Debug "`$isUpdated:,$isUpdated"
		Write-Output $isUpdated
	}
}
