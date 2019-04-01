Function Update-MSWordFindReplaceTextDocument {
	<#

		.SYNOPSIS
			When provided with an open Microsoft Word document, executes a set of text only FindText/ReplaceWith operations throughout the whole document.

		.DESCRIPTION	
			When provided with an open Microsoft Word document (Word.Application's Documents.Open), executes a set of text only FindText/ReplaceWith operations throughout the whole document.  

			MS Word supports opening the following file types:
				*.doc
				*.docm
				*.docx
				*.dot
				*.dotm
				*.dotx
				*.htm
				*.html
				*.htm
				*.html
				*.mht
				*.mhtml
				*.odt
				*.pdf
				*.rtf
				*.txt
				*.wps
				*.xml
				*.xml
				*.xps

			Microsoft's Word Range.Find operation performs a simple text match.  There is no support wildcard or regular expressions [RegEx]. Formatting of the FindText is preserved.  
			To reduce the chance on unintended replacements, surround keywords with a marker.  Parenthesis are safe when used with Microsoft's Word, Excel, Outlook and PowerPoint documents.
			For example: (CompanyName)
			
		.OUTPUTS
			The output of this function is the modifications executed on the open document.
			The returned value from this function is the number of replacements made.  
			
		.PARAMETER Path [Microsoft.Office.Interop.Word.DocumentClass]
			An open Microsoft Word document (Word.Application's Documents.Open).  If FindText is found this document will be modified.  
			
		.PARAMETER FindReplacePath String
			A hash table @{ [String] FindText, [String] ReplaceWith } pairs.  Each FindText ReplaceWith operation will be executed througout the whole document.  

		.EXAMPLE
			Description
			-----------
				This example opens a Mircosoft Word application,
				opens a MS Word document named '.\MySource.docx' in read only mode,
				creates a substitution hash table with 2 entries,
				and then calls this function.
				After the document is saved to another file name,
				and closed.  
			
				$wordApp = New-Object -ComObject Word.Application
				$document = $wordApp.Documents.Open( '.\MySource.docx', $FALSE, $TRUE ) # FileName, ConfirmConversions, ReadOnly
				$findReplace = @{ 'INCORRECT' = 'correct'; '(Field)' = 'MyFieldValue' }
				
				Update-MSWordFindReplaceTextDocument -Document $document -FindReplaceTable $findReplace
				
				$document.SaveAs( .\MyResults.docx' )
				$document.Close( $TRUE ) # SaveChanges

		.NOTE
			Author: Terry E Dow
			Creation Date: 2018-03-02
			Last Modified: 2019-03-16
			
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
		[Microsoft.Office.Interop.Word.DocumentClass] $Document = $NULL,
		
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
		
		# Define StoryRange's Find Execute method parameters.
		# https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.find.execute
		# FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace, MatchKashida, MatchDiacritics, MatchAlefHamza, MatchControl
		#$FindText = ''
		$findMatchCase = $FALSE
		$findMatchWholeWord = $FALSE
		$findMatchWildcards = $FALSE
		$findMatchSoundsLike = $FALSE
		$findMatchAllWordForms = $FALSE
		$findForward = $TRUE
		$findWrap = [Microsoft.Office.Interop.Word.WdFindWrap]::wdFindContinue
		$findFormat = $FALSE
		#$ReplaceWith = ''
		$findReplace = [Microsoft.Office.Interop.Word.WdReplace]::wdReplaceAll
			
	}
	
	Process { 
		Write-Debug "(`$FindReplaceTable).Count:,$(($FindReplaceTable).Count)"
	
		# Initialize metrics.
		$replacementCount = 0 
		$replacementCountTotal = 0 
				
		# Loop through each of the Document's StoryRanges
		$Document.StoryRanges | # -Verify:$FALSE
			ForEach-Object {
				Write-Verbose 'StoryRange'
			
				# [System.__ComObject#{0002095e-0000-0000-c000-000000000046}]
				$storyRange = $PSItem
				
				# For each substitution table entry, execute this StoryRange's Find FindText/ReplaceWith method.
				$FindReplaceTable.GetEnumerator() | 
					ForEach-Object {
						Write-Verbose "FindText:,$($PSItem.Key)"	
						
						# Accumulate the count of this StoryRange's executed replacements.
						$replacementCount = $storyRange.Find.Execute( $PSItem.Key, $findMatchCase, $findMatchWholeWord, $findMatchWildcards, $findMatchSoundsLike, $findMatchAllWordForms, $findForward, $findWrap, $findFormat, $PSItem.Value, $findReplace ) 
						Write-Verbose "`$replacementCount:,$replacementCount"
						$replacementCountTotal += $replacementCount
					}			
			}

	}
	
	End {
		#If ( $replacementCountTotal ) {
		#	# Update Document's Table of Contents.
		#	$tablesOfContents = $Document.TablesOfContents
		#	$tablesOfContents.Update()
		#}
		
		Write-Output $replacementCountTotal 
	}
}
