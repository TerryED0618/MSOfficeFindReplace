Function New-MSOutlookFindReplaceTextFile {
	<#

		.SYNOPSIS
			When provided with an open Microsoft Outlook document, executes a set of text only FindText/ReplaceWith operations throughout the whole document.

		.DESCRIPTION	
			When provided with an open Microsoft Outlook document file name (wildcards are permitted), executes a set of text only FindText/ReplaceWith operations throughout the whole document.  
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
			One output file is generated per source document file, by default in a subfolder called '.\Reports\'.  Use -OutFolderPath to specify an alternate location.  The output file names are in the format of: 
				<source file base name>[-<execution source>]-<date/time/timezone stamp>[-<file name tag>].<Extension>
				
			If parameter -Debug or -Verbose is specified, then a second file, a PowerShell transcript (.LOG), is created in the same location.
			
		.PARAMETER Path String[]
			Specifies a path to Microsoft Outlook compatible document file pathname. Wildcards are permitted. The default location is the current directory.
			
		.PARAMETER FindReplacePath String
			Specifies a path to one Comma Separated Value (CSV) FindReplace file. The CSV must have at least two column headings (case insensitive), all other columns are ignored: 
			Find,Replace

		
		.PARAMETER Attributes FileAttributes
			Gets files and folders with the specified attributes. This parameter supports all attributes and lets you specify complex combinations of attributes.

			For example, to get non-system files (not directories) that are encrypted or compressed, type:
				Get-ChildItem -Attributes !Directory+!System+Encrypted, !Directory+!System+Compressed

			To find files and folders with commonly used attributes, you can use the Attributes parameter, or the Directory, File, Hidden, ReadOnly, and System switch parameters.

			The Attributes parameter supports the following attributes: Archive, Compressed, Device, Directory, Encrypted, Hidden, Normal, NotContentIndexed, Offline, ReadOnly, ReparsePoint, SparseFile, System, and Temporary. For a description of these attributes, see the FileAttributes enumeration at http://go.microsoft.com/fwlink/?LinkId=201508.

			Use the following operators to combine attributes.
				!    NOT
			   +    AND
			   ,      OR
			No spaces are permitted between an operator and its attribute. However, spaces are permitted before commas.

			You can use the following abbreviations for commonly used attributes:
				D    Directory
				H    Hidden
				R    Read-only
				S     System

		.PARAMETER Directory SwitchParameter
			Gets directories (folders).  

			To get only directories, use the Directory parameter and omit the File parameter. To exclude directories, use the File parameter and omit the Directory parameter, or use the Attributes parameter. 

			To get directories, use the Directory parameter, its "ad" alias, or the Directory attribute of the Attributes parameter.

		.PARAMETER File SwitchParameter
			Gets files. 

			To get only files, use the File parameter and omit the Directory parameter. To exclude files, use the Directory parameter and omit the File parameter, or use the Attributes parameter.

			To get files, use the File parameter, its "af" alias, or the File value of the Attributes parameter.

		.PARAMETER Hidden SwitchParameter
			Gets only hidden files and directories (folders).  By default, Get-ChildItem gets only non-hidden items, but you can use the Force parameter to include hidden items in the results.

		To get only hidden items, use the Hidden parameter, its "h" or "ah" aliases, or the Hidden value of the Attributes parameter. To exclude hidden items, omit the Hidden parameter or use the Attributes parameter.

		.PARAMETER ReadOnly SwitchParameter
			Gets only read-only files and directories (folders).  

		To get only read-only items, use the ReadOnly parameter, its "ar" alias, or the ReadOnly value of the Attributes parameter. To exclude read-only items, use the Attributes parameter.

		.PARAMETER System SwitchParameter
			Gets only system files and directories (folders).

			To get only system files and folders, use the System parameter, its "as" alias, or the System value of the Attributes parameter. To exclude system files and folders, use the Attributes parameter.

		.PARAMETER Force SwitchParameter
			Gets hidden files and folders. By default, hidden files and folder are excluded. You can also get hidden files and folders by using the Hidden parameter or the Hidden value of the Attributes parameter.

		.PARAMETER UseTransaction SwitchParameter
			Includes the command in the active transaction. This parameter is valid only when a transaction is in progress. For more information, see about_Transactions.

		.PARAMETER Depth UInt32
			{{Fill Depth Description}}

		.PARAMETER Exclude String[]
			Specifies, as a string array, an item or items that this cmdlet excludes in the operation. The value of this parameter qualifies the Path parameter. Enter a path element or pattern, such as *.txt. Wildcards are permitted.

		.PARAMETER Filter String
			Specifies a filter in the provider's format or language. The value of this parameter qualifies the Path parameter. The syntax of the filter, including the use of wildcards, depends on the provider. Filters are more efficient than other parameters, because the provider applies them when retrieving the objects, rather than having Windows PowerShell filter the objects after they are retrieved.

		.PARAMETER Include String[]
			Specifies, as a string array, an item or items that this cmdlet includes in the operation. The value of this parameter qualifies the Path parameter. Enter a path element or pattern, such as *.txt. Wildcards are permitted.
			
			The default is MS Outlook supported file types:
				*.eml
				*.msg

			The Include parameter is effective only when the command includes the Recurse parameter or the path leads to the contents of a directory, such as C:\Windows\*, where the wildcard character specifies the contents of the C:\Windows directory.

		.PARAMETER LiteralPath String[]
			Specifies, as a string arrya, a path to one or more locations. Unlike the Path parameter, the value of the LiteralPath parameter is used exactly as it is typed. No characters are interpreted as wildcards. If the path includes escape characters, enclose it in single quotation marks. Single quotation marks tell Windows PowerShell not to interpret any characters as escape sequences.

		.PARAMETER Name SwitchParameter
			Indicates that this cmdlet gets only the names of the items in the locations. If you pipe the output of this command to another command, only the item names are sent.

		.PARAMETER Path String[]
			Specifies a path to one or more locations. Wildcards are permitted. The default location is the current directory (.).

		.PARAMETER Recurse SwitchParameter
			Indicates that this cmdlet gets the items in the specified locations and in all child items of the locations.

			In Windows PowerShell 2.0 and earlier versions of Windows PowerShell, the Recurse parameter works only when the value of the Path parameter is a container that has child items, such as C:\Windows or C:\Windows\ , and not when it is an item does not have child items, such as C:\Windows\ .exe.

			
		.PARAMETER ExecutionSource
			Specifiy the script's execution environment source.  Must be either; 'msExchOrganizationName', 'ForestName', 'DomainName', 'ComputerName', or an arbitrary string including '' or $NULL.
			If msExchOrganizationName is requested, but there is no Microsoft Exchange organization, ForestName will be used.
			If ForestName is requested, but there is no forest, DomainName will be used.  The forest name is of the executing computer's domain membership.  
			If the DomainName is requested, but the computer is not a domain member, ComputerName is used.  The domain name is of the executing computer's domain membership.  
			An arbitrary string can be used in the case where the Microsoft Exchange organization name, forest name or domain name is too generic (e.g. 'EMAIL', 'CORP' or 'ROOT').
			Defaults is msExchOrganizationName.
		
		.PARAMETER OutFileNameTag
			Optional comment string added to the end of the output file name.
		
		.PARAMETER OutFolderPath
			Specify which folder path to write the outfile.  Supports UNC and relative reference to the current script folder.  Except for UNC paths, this function will attempt to create and compress the output folder if it doesn't exist.  The default is .\Reports subfolder.  
		
		.PARAMETER AlertOnly
			When enabled, only unhealthy items are reported and the optional mail subject will contain 'alert' instead of 'report', and if there are no unhealthy items there is no output.  
		
		.PARAMETER MailFrom
			Optionally specify the address from which the mail is sent. Enter a name (optional) and e-mail address, such as 'Name <LocalPart@domain.com>'. 
		
		.PARAMETER MailTo
			Optioanlly specify the addresses to which the mail is sent. Enter names (optional) and the e-mail address, such as 'Name1 <LocalPart1@domain.com>','Name2 <LocalPart2@domain.com>'. 
		
		.PARAMETER MailServer
			Optionally specify the name of the SMTP server that sends the mail message.
		
		.PARAMETER CompressAttachmentLargerThan
			Optionally specify that when a file attachment size is over this limit that it should be compressed when e-mailed.  The default is 5MB.  There is no guarantee the compressed attachment will be below the sender or recipeint's message size limit.  
			
		.EXAMPLE
			Description
			-----------
			If find/replace file '.\MyFindReplace.csv's finds matches in Microsoft Outlook document file '.\MySource.msg' then a new document '.\Reports\MySource-Mine-20190302T235959+12.docx file will be creatd.
			
			New-MSOutlookFindReplaceTextFile -Path .\MySource.msg -FindReplacePath .\MyFindReplace.csv -ExecutionSource Mine
			
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
		SupportsShouldProcess = $TRUE # Enable support for -WhatIf by invoked destructive cmdlets.
	)]
	#[System.Diagnostics.DebuggerHidden()]
	Param(

		[Parameter(
		ValueFromPipeline=$TRUE,
		ValueFromPipelineByPropertyName=$TRUE )]
		[ValidateScript({ If (Test-Path -Path $PSItem -PathType 'Leaf') {$TRUE} Else { Throw 'File not found.' } })] 
		[String] $FindReplacePath = $NULL,

		
		[Parameter(
		ValueFromPipeline=$TRUE,
		ValueFromPipelineByPropertyName=$TRUE )]
		[String] $Attributes = $NULL,

		[Parameter(
		ValueFromPipeline=$TRUE,
		ValueFromPipelineByPropertyName=$TRUE )]
		[Switch] $Directory = $NULL,

		[Parameter(
		ValueFromPipeline=$TRUE,
		ValueFromPipelineByPropertyName=$TRUE )]
		[Switch] $File = $NULL,

		[Parameter(
		ValueFromPipeline=$TRUE,
		ValueFromPipelineByPropertyName=$TRUE )]
		[Switch] $Hidden = $NULL,

		[Parameter(
		ValueFromPipeline=$TRUE,
		ValueFromPipelineByPropertyName=$TRUE )]
		[Switch] $ReadOnly = $NULL,

		[Parameter(
		ValueFromPipeline=$TRUE,
		ValueFromPipelineByPropertyName=$TRUE )]
		[Switch] $System = $NULL,

		[Parameter(
		ValueFromPipeline=$TRUE,
		ValueFromPipelineByPropertyName=$TRUE )]
		[Switch] $Force = $NULL,

		[Parameter(
		ValueFromPipeline=$TRUE,
		ValueFromPipelineByPropertyName=$TRUE )]
		[Switch] $UseTransaction = $NULL,

		[Parameter(
		ValueFromPipeline=$TRUE,
		ValueFromPipelineByPropertyName=$TRUE )]
		[UInt32] $Depth = $NULL,

		[Parameter(
		ValueFromPipeline=$TRUE,
		ValueFromPipelineByPropertyName=$TRUE )]
		[String[]] $Exclude = $NULL,

		[Parameter(
		ValueFromPipeline=$TRUE,
		Position=1)]
		[String] $Filter = $NULL,

		[Parameter(
		ValueFromPipeline=$TRUE,
		ValueFromPipelineByPropertyName=$TRUE )]
		[String[]] $Include = ( '*.eml', '*.msg' ),

		[Parameter(
		ValueFromPipeline=$TRUE,
		ValueFromPipelineByPropertyName=$TRUE )]
		[String[]] $LiteralPath = $NULL,

		[Parameter(
		ValueFromPipeline=$TRUE,
		ValueFromPipelineByPropertyName=$TRUE )]
		[Switch] $Name = $NULL,

		[Parameter(
		ValueFromPipeline=$TRUE,
		Position=0)]
		[String[]] $Path = $NULL,

		[Parameter(
		ValueFromPipeline=$TRUE,
		ValueFromPipelineByPropertyName=$TRUE )]
		[Switch] $Recurse = $NULL,

	#region Script Header

		[Parameter( HelpMessage='Specify the script''s execution environment source.  Must be either ''ComputerName'', ''DomainName'', ''msExchOrganizationName'' or an arbitrary string. Defaults is msExchOrganizationName.' ) ]
			[String] $ExecutionSource = $NULL,

		[Parameter( HelpMessage='Optional string added to the end of the output file name.' ) ]
			[String] $OutFileNameTag = $NULL,

		[Parameter( HelpMessage='Specify where to write the output file.' ) ]
			[String] $OutFolderPath = '.\Reports',

		[Parameter( HelpMessage='When enabled, only unhealthy items are reported.' ) ]
			[Switch] $AlertOnly = $FALSE,

		[Parameter( HelpMessage='Optionally specify the address from which the mail is sent.' ) ]
			[String] $MailFrom = $NULL,

		[Parameter( HelpMessage='Optioanlly specify the addresses to which the mail is sent.' ) ]
			[String[]] $MailTo = $NULL,

		[Parameter( HelpMessage='Optionally specify the name of the SMTP server that sends the mail message.' ) ]
			[String] $MailServer = $NULL,

		[Parameter( HelpMessage='If the mail message attachment is over this size compress (zip) it.' ) ]
			[Int] $CompressAttachmentLargerThan = 5MB
	)

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

	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
	# Collect script execution metrics.
	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

	$scriptStartTime = Get-Date
	Write-Verbose "`$scriptStartTime:,$($scriptStartTime.ToString('s'))"
	
	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
	# Build output and log file path name.
	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

	$outFilePathBase = New-OutFilePathBase -OutFolderPath $OutFolderPath -ExecutionSource $ExecutionSource -OutFileNameTag $OutFileNameTag

	$pathInfo = Get-Item -Path $Path
	$baseName = ($pathInfo).BaseName
	$extension = ($pathInfo).Extension
	
	#$outFilePathName = ( $( ( "$($outFilePathBase.FolderPath)$baseName",  $ExecutionSource, $outFilePathBase.DateTimeStamp, $OutFileNameTag ) | Where-Object { $PSItem } ) -Join '-').Trim( '-' ) +  $extension
	#Write-Debug "`$outFilePathName: $outFilePathName"
	
	$logFilePathName = "$($outFilePathBase.Value).log"
	Write-Debug "`$logFilePathName: $logFilePathName"
	
	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
	# Optionally start or restart PowerShell transcript.
	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

	If ( $Debug -Or $Verbose ) {
		Try {
			Start-Transcript -Path $logFilePathName -WhatIf:$FALSE
		} Catch {
			Stop-Transcript
			Start-Transcript -Path $logFilePathName -WhatIf:$FALSE
		}
	}

	#endregion Script Header

	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
	# Main process
	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

	# Create Get-ChildItem hash table to splat parameters.  
	$getChildItemParameters = @{}
	If ( $Attributes ) { $getChildItemParameters.Attributes = $Attributes }
	If ( $Directory ) { $getChildItemParameters.Directory = $Directory }
	If ( $File ) { $getChildItemParameters.File = $File }
	If ( $Hidden ) { $getChildItemParameters.Hidden = $Hidden }
	If ( $ReadOnly ) { $getChildItemParameters.ReadOnly = $ReadOnly }
	If ( $System ) { $getChildItemParameters.System = $System }
	If ( $Force ) { $getChildItemParameters.Force = $Force }
	If ( $UseTransaction ) { $getChildItemParameters.UseTransaction = $UseTransaction }
	If ( $Depth ) { $getChildItemParameters.Depth = $Depth }
	If ( $Exclude ) { $getChildItemParameters.Exclude = $Exclude }
	If ( $Filter ) { $getChildItemParameters.Filter = $Filter }
	If ( $Include ) { $getChildItemParameters.Include = $Include }
	If ( $LiteralPath ) { $getChildItemParameters.LiteralPath = $LiteralPath }
	If ( $Name ) { $getChildItemParameters.Name = $Name }
	If ( $Path ) { $getChildItemParameters.Path = $Path }
	If ( $Recurse ) { $getChildItemParameters.Recurse = $Recurse }
	If ( $Debug ) {
		ForEach ( $key In $getChildItemParameters.Keys ) {
			Write-Debug "`$getChildItemParameters[$key]`:,$($getChildItemParameters[$key])"
		}
	}

	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

	# Import FindText/ReplaceWith into a substitution hash table: [String] FindText, [String] ReplaceWith
	$findReplaceTable = @{}
	$findReplaceRecordCounter = 0
	Import-CSV -Path $FindReplacePath |
		ForEach-Object{
			$findReplaceRecordCounter++
			# NOT VERIFIED if this test is required by Outlook text properties.
			If ( $PSItem.Find.Length -LE 256 -And $PSItem.Replace.Length -LE 255 ) {
				Write-Debug "`$PSItem.Find.Length`:,$($PSItem.Find.Length)"
				Write-Debug "`$PSItem.Replace.Length`:,$($PSItem.Replace.Length)"
				$findReplaceTable.Add( $PSItem.Find, $PSItem.Replace ) 
				
			} Else {
				If ( $PSItem.Find.Length -LE 256 ) {
					Out-Host -InputObject "Length of FindText value '$PSItem.Find' on record $findReplaceRecordCounter is over 256 characters long and is not supported.  Record skipped."
				}
				If ( $PSItem.Replace.Length -LE 255 ) {
					Out-Host -InputObject "Length of ReplaceWith value '$PSItem.Replace' on record $findReplaceRecordCounter is over 255 characters long and is not supported.  Record skipped."
				}
			}
		}
	If ( $Debug ) {
		$findReplaceTable.GetEnumerator() |
			ForEach-Object {
				Write-Debug "`$findReplaceTable[$($PSItem.Key)]`:,$($PSItem.Value)"
			}
	}
	
	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

	# Microsoft Outlook only allows a single instance to be running.  Do not close an application started before execution.  
	# If the Microsoft Outlook application is not running, then start it.  Upon end, close the application.  
	# If Microsoft Outlook application is running, connect to it.  Upon end, do not close the application.  
	$wasOutlookRunning = [Bool] ( Get-Process | Where-Object { $PSItem.Name -EQ 'OUTLOOK' } )
	
	# Start Microsoft Outlook application.
	# [Microsoft.Office.Interop.Outlook.ApplicationClass]
	$OutlookApp = New-Object -ComObject Outlook.Application
	#$OutlookMapiNameSpace = $OutlookApp.GetNameSpace( 'MAPI' )
	
	# For each document file...
	Get-ChildItem @getChildItemParameters |
		ForEach-Object {
	
			# Open Microsoft Outlook document.  
			# [Microsoft.Office.Interop.Outlook.MailItemClass] [Microsoft.Office.Interop.Outlook.MailItem]
			$document = $OutlookApp.CreateItemFromTemplate( $PSItem.FullName ) # TemplatePath, InFolder 
			Write-Debug "`$document.GetType():,$($document.GetType())"
			
			# Construct an OutFile name.  
			$outFilePathName = ( $( ( "$($outFilePathBase.FolderPath)$($PSItem.BaseName)",  $ExecutionSource, $outFilePathBase.DateTimeStamp, $OutFileNameTag ) | Where-Object { $PSItem } ) -Join '-').Trim( '-' ) +  $PSItem.Extension
			Write-Debug "`$outFilePathName`:,$outFilePathName"
			
			# Update document, checking if any replacements were executed.  
			If ( Update-MSOutlookFindReplaceTextDocument -Document $document -FindReplaceTable $findReplaceTable ) {
			
				# Replacements executed, save document.
				$document.SaveAs( [Ref] $outFilePathName )
				$document.Close( [Microsoft.Office.Interop.Outlook.OlInspectorClose]::olSave ) # SaveMode
				
				# Write metrics.
				Out-Host -InputObject "New document saved to '$outFilePathName'."	
			} Else {
			
				# No replacements executed, don't save document.
				$document.Close( [Microsoft.Office.Interop.Outlook.OlInspectorClose]::olDiscard ) # SaveMode
				
				# Write metrics.
				Out-Host -InputObject "No FindText found in '$($PSItem.FullName)'"	
			}
		}
	
	
	# If the Microsoft Outlook application was already running when the this script was started, then do not close the application.  
	If ( -Not $wasOutlookRunning ) {
		# Close Microsoft Outlook application. 
		$OutlookApp.Quit()
		
		# Free up memory.  
		#$NULL = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$OutlookApp)
		$NULL = [System.Runtime.InteropServices.Marshal]::ReleaseComObject( $OutlookApp )
		[gc]::Collect()
		[gc]::WaitForPendingFinalizers()
	}
	Remove-Variable OutlookApp 

	#region Script Footer

	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
	# Optionally mail report.
	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

	If ( (Test-Path -PathType Leaf -Path $outFilePathName) -And $MailFrom -And $MailTo -And $MailServer ) {

		# Determine subject line report/alert mode.
		If ( $AlertOnly ) {
			$reportType = 'Alert'
		} Else {
			$reportType = 'Report'
		}

		$messageSubject = "New Microsoft Outlook replace text $reportType for $($outFilePathBase.ExecutionSourceName) on $((Get-Date).ToString('s'))"

		# If the out file is larger then a specified limit (message size limit), then create a compressed (zipped) copy.
		Write-Debug "$outFilePathName.Length:,$((Get-ChildItem -LiteralPath $outFilePathName).Length)"
		If ( $CompressAttachmentLargerThan -LT (Get-ChildItem -LiteralPath $outFilePathName).Length ) {

			$outZipFilePathName = "$outFilePathName.zip"
			Write-Debug "`$outZipFilePathName:,$outZipFilePathName"

			# Create a temporary empty zip file.
			Set-Content -Path $outZipFilePathName -Value ( "PK" + [Char]5 + [Char]6 + ("$([Char]0)" * 18) ) -Force -WhatIf:$FALSE

			# Wait for the zip file to appear in the parent folder.
			While ( -Not (Test-Path -PathType Leaf -Path $outZipFilePathName) ) {
				Write-Debug "Waiting for:,$outZipFilePathName"
				Start-Sleep -Milliseconds 20
			}

			# Wait for the zip file to be written by detecting that the file size is not zero.
			While ( -Not (Get-ChildItem -LiteralPath $outZipFilePathName).Length ) {
				Write-Debug "Waiting for ($outZipFilePathName\$($outFilePathBase.FileName).csv).Length:,$((Get-ChildItem -LiteralPath $outZipFilePathName).Length)"
				Start-Sleep -Milliseconds 20
			}

			# Bind to the zip file as a folder.
			$outZipFile = (New-Object -ComObject Shell.Application).NameSpace( $outZipFilePathName )

			# Copy out file into Zip file.
			$outZipFile.CopyHere( $outFilePathName )

			# Wait for the compressed file to be appear in the zip file.
			While ( -Not $outZipFile.ParseName("$($outFilePathBase.FileName).csv") ) {
				Write-Debug "Waiting for:,$outZipFilePathName\$($outFilePathBase.FileName).csv"
				Start-Sleep -Milliseconds 20
			}

			# Wait for the compressed file to be written into the zip file by detecting that the file size is not zero.
			While ( -Not ($outZipFile.ParseName("$($outFilePathBase.FileName).csv")).Size ) {
				Write-Debug "Waiting for ($outZipFilePathName\$($outFilePathBase.FileName).csv).Size:,$($($outZipFile.ParseName($($outFilePathBase.FileName).csv)).Size)"
				Start-Sleep -Milliseconds 20
			}

			# Send the report.
			Send-MailMessage `
				-From $MailFrom `
				-To $MailTo `
				-SmtpServer $MailServer `
				-Subject $messageSubject `
				-Body 'See attached zipped Excel (CSV) spreadsheet.' `
				-Attachments $outZipFilePathName

			# Remove the temporary zip file.
			Remove-Item -LiteralPath $outZipFilePathName

		} Else {

			# Send the report.
			Send-MailMessage `
				-From $MailFrom `
				-To $MailTo `
				-SmtpServer $MailServer `
				-Subject $messageSubject `
				-Body 'See attached Excel (CSV) spreadsheet.' `
				-Attachments $outFilePathName
		}
	}

	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
	# Optionally write script execution metrics and stop the Powershell transcript.
	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

	$scriptEndTime = Get-Date
	Write-Verbose "`$scriptEndTime:,$($scriptEndTime.ToString('s'))"
	$scriptElapsedTime =  $scriptEndTime - $scriptStartTime
	Write-Verbose "`$scriptElapsedTime:,$scriptElapsedTime"
	If ( $Debug -Or $Verbose ) {
		Stop-Transcript
	}
	#endregion Script Footer
}
