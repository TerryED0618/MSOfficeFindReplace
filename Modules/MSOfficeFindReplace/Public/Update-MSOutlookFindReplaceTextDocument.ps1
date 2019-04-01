Function Update-MSOutlookFindReplaceTextDocument {
	<#

		.SYNOPSIS
			When provided with an open Microsoft Outlook document, executes a set of text only FindText/ReplaceWith operations throughout the whole document.

		.DESCRIPTION	
			When provided with an open Microsoft Outlook document (Outlook.Application's CreateItemFromTemplate), executes a set of text only FindText/ReplaceWith operations throughout the whole document.  
			MS Outlook supports opening the following file types:
				*.eml
				*.msg

			The following MailItem properties are updated:
				To 
				Cc 
				Bcc
				Subject 
				Body 

			The replace operation performs a simple text match.  There is no support wildcard or regular expressions [RegEx].
			To reduce the chance on unintended replacements, surround keywords with a marker.  Parenthesis are safe when used with Microsoft's Word, Excel, Outlook and PowerPoint documents.
			For example: (CompanyName)
						
		.OUTPUTS
			The output of this function is the modifications executed on the open document.
			The returned [BOOL] from this function is TRUE if any of the FindText is updated, and FALSE if none of the the FindText was not found.
			
		.PARAMETER Path [Microsoft.Office.Interop.Outlook.MailItemClass]
			An open Microsoft Outlook document (Outlook.Application's Documents.Open).  If FindText is found this document will be modified.  
			
		.PARAMETER FindReplacePath String
			A hash table @{ [String] FindText, [String] ReplaceWith } pairs.  Each FindText ReplaceWith operation will be executed througout the whole document.  

		.EXAMPLE
			Description
			-----------
				This example opens a Mircosoft Outlook application,
				opens a MS Outlook document named '.\MySource.msg',
				creates a substitution hash table with 2 entries,
				and then calls this function.
				After the document is saved to another file name,
				and closed.  
			
				$OutlookApp = New-Object -ComObject Outlook.Application
				$document = $OutlookApp.CreateItemFromTemplate( ( '.\MySource.msg' ) 
				$findReplace = @{ '(INCORRECT)' = 'correct'; '(To1)' = 'Support@contoso.com' }
				
				Update-MSOutlookFindReplaceTextDocument -Document $document -FindReplaceTable $findReplace
				
				$document.SaveAs( .\MyResults.msgs' )
				$document.Close( [Microsoft.Office.Interop.Outlook.OlInspectorClose]::olSave )

		.NOTE
			Author: Terry E Dow
			Creation Date: 2018-03-17
			Last Modified: 2019-03-17

			Warning from Microsoft:
				Considerations for server-side Automation of Office https://support.microsoft.com/en-us/help/257757/considerations-for-server-side-automation-of-office
				
			Reference:
				Microsoft Outlook Constants https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa219371(v=office.11)
				[MS-OXMSG]: Outlook Item (.msg) File Format https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxmsg/b046868c-9fbf-41ae-9ffb-8de2bd4eec82
				System.Net.Mail Namespace https://docs.microsoft.com/en-us/dotnet/api/system.net.mail?redirectedfrom=MSDN&view=netframework-4.7.2
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
		#[Microsoft.Office.Interop.Outlook.MailItem] $Document, 
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
		
		# Sublist of [Microsoft.Office.Interop.Outlook.OlUserPropertyType]::olText properties available to update.
		# $document.ItemProperties | Where-Object { $PSItem.Type -Eq 1 } | FT Name
		#	BCC *
		#	BillingInformation
		#	Body *
		#	Categories
		#	CC *
		#	Companies
		#	ConversationTopic
		#	FlagRequest
		#	MessageClass
		#	Mileage
		#	OutlookVersion
		#	ReceivedByName
		#	ReceivedOnBehalfOfName
		#	ReminderSoundFile
		#	ReplyRecipientNames
		#	SenderEmailAddress
		#	SenderEmailType
		#	SenderName
		#	SentOnBehalfOfName
		#	Subject *
		#	To *
		#	VotingResponse
			
	}
	
	Process { 
		Write-Debug "(`$FindReplaceTable).Count:,$(($FindReplaceTable).Count)"
	
		# Initialize metrics.
		$isUpdated = $FALSE
		
		# For each substitution table entry...
		$FindReplaceTable.GetEnumerator() | 
			ForEach-Object {
				Write-Verbose "FindText:,$($PSItem.Key)"	
				$findTextEscaped = [System.Text.RegularExpressions.Regex]::Escape($PSItem.Key)
				Write-Debug "findTextEscaped:,$findTextEscaped"	
				
				# Update each of these Document's parts.				
				Write-Debug "`$Document.To:,$($Document.To)"
				If ( $Document.To -Match $findTextEscaped ) { 
					$Document.To = $Document.To -Replace $findTextEscaped, $PSItem.Value 
					$isUpdated = $TRUE
					Write-Debug "`$Document.To updated:,$($Document.To)"
				}
				
				Write-Debug "`$Document.Cc:,$($Document.Cc)"	
				If ( $Document.Cc -Match $findTextEscaped ) { 
					$Document.Cc = $Document.Cc -Replace $findTextEscaped, $PSItem.Value
					$isUpdated = $TRUE
					Write-Debug "`$Document.Cc updated:,$($Document.Cc)"	
				}
				
				Write-Debug "`$Document.Bcc:,$($Document.Bcc)"	
				If ( $Document.Bcc -Match $findTextEscaped ) { 
					$Document.Bcc = $Document.Bcc -Replace $findTextEscaped, $PSItem.Value
					$isUpdated = $TRUE
					Write-Debug "`$Document.Bcc updated:,$($Document.Bcc)"	
				}
				
				Write-Debug "`$Document.Subject:,$($Document.Subject)"	
				If ( $Document.Subject -Match $findTextEscaped ) { 
					$Document.Subject = $Document.Subject -Replace $findTextEscaped, $PSItem.Value
					$isUpdated = $TRUE
					Write-Debug "`$Document.Subject updated:,$($Document.Subject)"	
				}
				
				Write-Debug "`$Document.Body:,$($Document.Body)"	
				If ( $Document.Body -Match $findTextEscaped ) { 
					$Document.Body = $Document.Body -Replace $findTextEscaped, $PSItem.Value
					$isUpdated = $TRUE
					Write-Debug "`$Document.Body updated:,$($Document.Body)"
				}
				
				# HtmlBody is olOutlookInternal not olText
				#Write-Debug "`$Document.HtmlBody:,$($Document.HtmlBody)"	
				#If ( $Document.HtmlBody -Match $findTextEscaped ) { 
				#	$Document.HtmlBody = $Document.HtmlBody -Replace $findTextEscaped, $PSItem.Value
				#	$isUpdated = $TRUE
				#	#Write-Debug "`$Document.HtmlBody updated:,$($Document.HtmlBody)"	
				#}
			}
			
	}
	
	End {
		Write-Debug "`$isUpdated:,$isUpdated"
		Write-Output $isUpdated
	}
}
