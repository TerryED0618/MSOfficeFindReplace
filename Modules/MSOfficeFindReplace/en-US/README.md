# MSOfficeFindReplace

### DESCRIPTION	
		When provided read-only access to a Microsoft Office compatible document, perform a set of text only FindText/ReplaceWith executions throughout the whole document, and then save to a new file.  

###	Functions Exported:
		New-MSOutlookFindReplaceTextFile
		New-MSWordFindReplaceTextFile
		Update-MSOutlookFindReplaceTextDocument
		Update-MSWordFindReplaceTextDocument
---
##	Function New-MSExcelFindReplaceTextFile 

###	SYNOPSIS
		When provided with an open Microsoft Excel document, executes a set of text only FindText/ReplaceWith operations throughout the whole document.

###	DESCRIPTION	
		When provided with an open Microsoft Excel document, executes a set of text only FindText/ReplaceWith operations throughout the whole document.  
		MS Excel supports opening the following file types:
			*.xlsx
			*.xlsm
			*.xlsb
			*.xltx
			*.xltm
			*.xls
			*.xlt
			*.xls
			*.xml
			*.xml
			*.xlam
			*.xla
			*.xlw
			*.xlr
			
		Microsoft Excel's Range.Replace operation performs a simple text match.  There is no support wildcard or regular expressions [RegEx].  Formatting of the FindText is preserved.  
		To reduce the chance on unintended replacements, surround keywords with a marker.  Parenthesis are safe when used with Microsoft's Word, Excel, Outlook and PowerPoint documents.
		For example: (CompanyName)
				
###	OUTPUTS
		One output file is generated per source document file, by default in a subfolder called '.\Reports\'.  Use -OutFolderPath to specify an alternate location.  The output file names are in the format of: 
			<source file base name>[-<execution source>]-<date/time/timezone stamp>[-<file name tag>].<Extension>
			
		If parameter -Debug or -Verbose is specified, then a second file, a PowerShell transcript (.LOG), is created in the same location.
			
###	PARAMETER Path String[]
		Specifies a path to Microsoft Excel compatible document file pathname. Wildcards are permitted. The default location is the current directory.
			
###	PARAMETER FindReplacePath String
		Specifies a path to one Comma Separated Value (CSV) FindReplace file. The CSV must have at least two column headings (case insensitive), all other columns are ignored: 
		Find,Replace

		
###	PARAMETER Attributes FileAttributes
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

###	PARAMETER Directory SwitchParameter
		Gets directories (folders).  

		To get only directories, use the Directory parameter and omit the File parameter. To exclude directories, use the File parameter and omit the Directory parameter, or use the Attributes parameter. 

		To get directories, use the Directory parameter, its "ad" alias, or the Directory attribute of the Attributes parameter.

###	PARAMETER File SwitchParameter
		Gets files. 

		To get only files, use the File parameter and omit the Directory parameter. To exclude files, use the Directory parameter and omit the File parameter, or use the Attributes parameter.

		To get files, use the File parameter, its "af" alias, or the File value of the Attributes parameter.

###	PARAMETER Hidden SwitchParameter
		Gets only hidden files and directories (folders).  By default, Get-ChildItem gets only non-hidden items, but you can use the Force parameter to include hidden items in the results.

		To get only hidden items, use the Hidden parameter, its "h" or "ah" aliases, or the Hidden value of the Attributes parameter. To exclude hidden items, omit the Hidden parameter or use the Attributes parameter.

###	PARAMETER ReadOnly SwitchParameter
		Gets only read-only files and directories (folders).  

		To get only read-only items, use the ReadOnly parameter, its "ar" alias, or the ReadOnly value of the Attributes parameter. To exclude read-only items, use the Attributes parameter.

###	PARAMETER System SwitchParameter
		Gets only system files and directories (folders).

		To get only system files and folders, use the System parameter, its "as" alias, or the System value of the Attributes parameter. To exclude system files and folders, use the Attributes parameter.

###	PARAMETER Force SwitchParameter
		Gets hidden files and folders. By default, hidden files and folder are excluded. You can also get hidden files and folders by using the Hidden parameter or the Hidden value of the Attributes parameter.

###	PARAMETER UseTransaction SwitchParameter
		Includes the command in the active transaction. This parameter is valid only when a transaction is in progress. For more information, see about_Transactions.

###	PARAMETER Depth UInt32
		{{Fill Depth Description}}

###	PARAMETER Exclude String[]
		Specifies, as a string array, an item or items that this cmdlet excludes in the operation. The value of this parameter qualifies the Path parameter. Enter a path element or pattern, such as *.txt. Wildcards are permitted.

###	PARAMETER Filter String
		Specifies a filter in the provider's format or language. The value of this parameter qualifies the Path parameter. The syntax of the filter, including the use of wildcards, depends on the provider. Filters are more efficient than other parameters, because the provider applies them when retrieving the objects, rather than having Windows PowerShell filter the objects after they are retrieved.

###	PARAMETER Include String[]
		Specifies, as a string array, an item or items that this cmdlet includes in the operation. The value of this parameter qualifies the Path parameter. Enter a path element or pattern, such as *.txt. Wildcards are permitted.
		
		The default is MS Excel supported file types:
			*.xlsx
			*.xlsm
			*.xlsb
			*.xltx
			*.xltm
			*.xls
			*.xlt
			*.xls
			*.xml
			*.xml
			*.xlam
			*.xla
			*.xlw
			*.xlr

		The Include parameter is effective only when the command includes the Recurse parameter or the path leads to the contents of a directory, such as C:\Windows\*, where the wildcard character specifies the contents of the C:\Windows directory.

###	PARAMETER LiteralPath String[]
		Specifies, as a string arrya, a path to one or more locations. Unlike the Path parameter, the value of the LiteralPath parameter is used exactly as it is typed. No characters are interpreted as wildcards. If the path includes escape characters, enclose it in single quotation marks. Single quotation marks tell Windows PowerShell not to interpret any characters as escape sequences.

###	PARAMETER Name SwitchParameter
		Indicates that this cmdlet gets only the names of the items in the locations. If you pipe the output of this command to another command, only the item names are sent.

###	PARAMETER Path String[]
		Specifies a path to one or more Microsoft Excel compatible document. Wildcards are permitted. The default location is the current directory (.).

###	PARAMETER Recurse SwitchParameter
		Indicates that this cmdlet gets the items in the specified locations and in all child items of the locations.

		In Windows PowerShell 2.0 and earlier versions of Windows PowerShell, the Recurse parameter works only when the value of the Path parameter is a container that has child items, such as C:\Windows or C:\Windows\ , and not when it is an item does not have child items, such as C:\Windows\ .exe.

			
###	PARAMETER ExecutionSource
		Specifiy the script's execution environment source.  Must be either; 'msExchOrganizationName', 'ForestName', 'DomainName', 'ComputerName', or an arbitrary string including '' or $NULL.
		If msExchOrganizationName is requested, but there is no Microsoft Exchange organization, ForestName will be used.
		If ForestName is requested, but there is no forest, DomainName will be used.  The forest name is of the executing computer's domain membership.  
		If the DomainName is requested, but the computer is not a domain member, ComputerName is used.  The domain name is of the executing computer's domain membership.  
		An arbitrary string can be used in the case where the Microsoft Exchange organization name, forest name or domain name is too generic (e.g. 'EMAIL', 'CORP' or 'ROOT').
		Defaults is msExchOrganizationName.
		
###	PARAMETER OutFileNameTag
		Optional comment string added to the end of the output file name.
		
###	PARAMETER OutFolderPath
		Specify which folder path to write the outfile.  Supports UNC and relative reference to the current script folder.  Except for UNC paths, this function will attempt to create and compress the output folder if it doesn't exist.  The default is .\Reports subfolder.  
		
###	PARAMETER AlertOnly
		When enabled, only unhealthy items are reported and the optional mail subject will contain 'alert' instead of 'report', and if there are no unhealthy items there is no output.  
		
###	PARAMETER MailFrom
		Optionally specify the address from which the mail is sent. Enter a name (optional) and e-mail address, such as 'Name <LocalPart@domain.com>'. 
		
###	PARAMETER MailTo
		Optioanlly specify the addresses to which the mail is sent. Enter names (optional) and the e-mail address, such as 'Name1 <LocalPart1@domain.com>','Name2 <LocalPart2@domain.com>'. 
		
###	PARAMETER MailServer
		Optionally specify the name of the SMTP server that sends the mail message.
		
###	PARAMETER CompressAttachmentLargerThan
		Optionally specify that when a file attachment size is over this limit that it should be compressed when e-mailed.  The default is 5MB.  There is no guarantee the compressed attachment will be below the sender or recipeint's message size limit.  
			
###	EXAMPLE
		Description
		-----------
		If find/replace file '.\MyFindReplace.csv's finds matches in Microsoft Excel document file '.\MySource.docx' then a new document '.\Reports\MySource-Mine-20190302T235959+12.docx file will be creatd.
		
		New-MSExcelFindReplaceTextFile -Path .\MySource.docx -FindReplacePath .\MyFindReplace.csv -ExecutionSource Mine
			
###	NOTE
		Author: Terry E Dow
		Creation Date: 2018-03-19
				
		Warning from Microsoft:
			Considerations for server-side Automation of Office https://support.microsoft.com/en-us/help/257757/considerations-for-server-side-automation-of-office

		Reference:

---

##	Function New-MSOutlookFindReplaceTextFile
	
###	SYNOPSIS
		When provided with an open Microsoft Outlook document, executes a set of text only FindText/ReplaceWith operations throughout the whole document.

###	DESCRIPTION	
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
			
###	OUTPUTS
		One output file is generated per source document file, by default in a subfolder called '.\Reports\'.  Use -OutFolderPath to specify an alternate location.  The output file names are in the format of: 
			<source file base name>[-<execution source>]-<date/time/timezone stamp>[-<file name tag>].<Extension>
			
		If parameter -Debug or -Verbose is specified, then a second file, a PowerShell transcript (.LOG), is created in the same location.
			
###	PARAMETER Path String[]
		Specifies a path to Microsoft Outlook compatible document file pathname. Wildcards are permitted. The default location is the current directory.
			
###	PARAMETER FindReplacePath String
		Specifies a path to one Comma Separated Value (CSV) FindReplace file. The CSV must have at least two column headings (case insensitive), all other columns are ignored: 
		Find,Replace

		
###	PARAMETER Attributes FileAttributes
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

###	PARAMETER Directory SwitchParameter
		Gets directories (folders).  

		To get only directories, use the Directory parameter and omit the File parameter. To exclude directories, use the File parameter and omit the Directory parameter, or use the Attributes parameter. 

		To get directories, use the Directory parameter, its "ad" alias, or the Directory attribute of the Attributes parameter.

###	PARAMETER File SwitchParameter
		Gets files. 

		To get only files, use the File parameter and omit the Directory parameter. To exclude files, use the Directory parameter and omit the File parameter, or use the Attributes parameter.

		To get files, use the File parameter, its "af" alias, or the File value of the Attributes parameter.

###	PARAMETER Hidden SwitchParameter
		Gets only hidden files and directories (folders).  By default, Get-ChildItem gets only non-hidden items, but you can use the Force parameter to include hidden items in the results.

		To get only hidden items, use the Hidden parameter, its "h" or "ah" aliases, or the Hidden value of the Attributes parameter. To exclude hidden items, omit the Hidden parameter or use the Attributes parameter.

###	PARAMETER ReadOnly SwitchParameter
		Gets only read-only files and directories (folders).  

		To get only read-only items, use the ReadOnly parameter, its "ar" alias, or the ReadOnly value of the Attributes parameter. To exclude read-only items, use the Attributes parameter.

###	PARAMETER System SwitchParameter
		Gets only system files and directories (folders).

		To get only system files and folders, use the System parameter, its "as" alias, or the System value of the Attributes parameter. To exclude system files and folders, use the Attributes parameter.

###	PARAMETER Force SwitchParameter
		Gets hidden files and folders. By default, hidden files and folder are excluded. You can also get hidden files and folders by using the Hidden parameter or the Hidden value of the Attributes parameter.

###	PARAMETER UseTransaction SwitchParameter
		Includes the command in the active transaction. This parameter is valid only when a transaction is in progress. For more information, see about_Transactions.

###	PARAMETER Depth UInt32
		{{Fill Depth Description}}

###	PARAMETER Exclude String[]
		Specifies, as a string array, an item or items that this cmdlet excludes in the operation. The value of this parameter qualifies the Path parameter. Enter a path element or pattern, such as *.txt. Wildcards are permitted.

###	PARAMETER Filter String
		Specifies a filter in the provider's format or language. The value of this parameter qualifies the Path parameter. The syntax of the filter, including the use of wildcards, depends on the provider. Filters are more efficient than other parameters, because the provider applies them when retrieving the objects, rather than having Windows PowerShell filter the objects after they are retrieved.

###	PARAMETER Include String[]
		Specifies, as a string array, an item or items that this cmdlet includes in the operation. The value of this parameter qualifies the Path parameter. Enter a path element or pattern, such as *.txt. Wildcards are permitted.
		
		The default is MS Outlook supported file types:
			*.eml
			*.msg

		The Include parameter is effective only when the command includes the Recurse parameter or the path leads to the contents of a directory, such as C:\Windows\*, where the wildcard character specifies the contents of the C:\Windows directory.

###	PARAMETER LiteralPath String[]
		Specifies, as a string arrya, a path to one or more locations. Unlike the Path parameter, the value of the LiteralPath parameter is used exactly as it is typed. No characters are interpreted as wildcards. If the path includes escape characters, enclose it in single quotation marks. Single quotation marks tell Windows PowerShell not to interpret any characters as escape sequences.

###	PARAMETER Name SwitchParameter
		Indicates that this cmdlet gets only the names of the items in the locations. If you pipe the output of this command to another command, only the item names are sent.

###	PARAMETER Path String[]
		Specifies a path to one or more locations. Wildcards are permitted. The default location is the current directory (.).

###	PARAMETER Recurse SwitchParameter
		Indicates that this cmdlet gets the items in the specified locations and in all child items of the locations.

		In Windows PowerShell 2.0 and earlier versions of Windows PowerShell, the Recurse parameter works only when the value of the Path parameter is a container that has child items, such as C:\Windows or C:\Windows\ , and not when it is an item does not have child items, such as C:\Windows\ .exe.

			
###	PARAMETER ExecutionSource
		Specifiy the script's execution environment source.  Must be either; 'msExchOrganizationName', 'ForestName', 'DomainName', 'ComputerName', or an arbitrary string including '' or $NULL.
		If msExchOrganizationName is requested, but there is no Microsoft Exchange organization, ForestName will be used.
		If ForestName is requested, but there is no forest, DomainName will be used.  The forest name is of the executing computer's domain membership.  
		If the DomainName is requested, but the computer is not a domain member, ComputerName is used.  The domain name is of the executing computer's domain membership.  
		An arbitrary string can be used in the case where the Microsoft Exchange organization name, forest name or domain name is too generic (e.g. 'EMAIL', 'CORP' or 'ROOT').
		Defaults is msExchOrganizationName.
		
###	PARAMETER OutFileNameTag
		Optional comment string added to the end of the output file name.
		
###	PARAMETER OutFolderPath
		Specify which folder path to write the outfile.  Supports UNC and relative reference to the current script folder.  Except for UNC paths, this function will attempt to create and compress the output folder if it doesn't exist.  The default is .\Reports subfolder.  
		
###	PARAMETER AlertOnly
		When enabled, only unhealthy items are reported and the optional mail subject will contain 'alert' instead of 'report', and if there are no unhealthy items there is no output.  
		
###	PARAMETER MailFrom
		Optionally specify the address from which the mail is sent. Enter a name (optional) and e-mail address, such as 'Name <LocalPart@domain.com>'. 
		
###	PARAMETER MailTo
		Optioanlly specify the addresses to which the mail is sent. Enter names (optional) and the e-mail address, such as 'Name1 <LocalPart1@domain.com>','Name2 <LocalPart2@domain.com>'. 
		
###	PARAMETER MailServer
		Optionally specify the name of the SMTP server that sends the mail message.
		
###	PARAMETER CompressAttachmentLargerThan
		Optionally specify that when a file attachment size is over this limit that it should be compressed when e-mailed.  The default is 5MB.  There is no guarantee the compressed attachment will be below the sender or recipeint's message size limit.  
			
###	EXAMPLE
		Description
		-----------
		If find/replace file '.\MyFindReplace.csv's finds matches in Microsoft Outlook document file '.\MySource.msg' then a new document '.\Reports\MySource-Mine-20190302T235959+12.docx file will be creatd.
		
		New-MSOutlookFindReplaceTextFile -Path .\MySource.msg -FindReplacePath .\MyFindReplace.csv -ExecutionSource Mine
			
###	NOTE
		Author: Terry E Dow
		Creation Date: 2018-03-17
		
		Warning from Microsoft:
			Considerations for server-side Automation of Office https://support.microsoft.com/en-us/help/257757/considerations-for-server-side-automation-of-office

		Reference:
			Microsoft Outlook Constants https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa219371(v=office.11)
			[MS-OXMSG]: Outlook Item (.msg) File Format https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxmsg/b046868c-9fbf-41ae-9ffb-8de2bd4eec82
			System.Net.Mail Namespace https://docs.microsoft.com/en-us/dotnet/api/system.net.mail?redirectedfrom=MSDN&view=netframework-4.7.2

---
##	Function New-MSPowerPointFindReplaceTextFile 

###	SYNOPSIS
		When provided with an open Microsoft PowerPoint document, executes a set of text only FindText/ReplaceWith operations throughout the whole document.

###	DESCRIPTION	
		When provided with an open Microsoft PowerPoint document, executes a set of text only FindText/ReplaceWith operations throughout the whole document.  
		MS PowerPoint supports opening the following file types:
			*.odp
			*.pot
			*.potm
			*.potx
			*.ppa
			*.ppam
			*.pps
			*.ppsm
			*.ppsx
			*.ppt
			*.pptm
			*.pptx
			*.pptx

		Microsoft PowerPoint's Range.Replace operation performs a simple text match.  There is no support wildcard or regular expressions [RegEx].  Formatting of the FindText is preserved.  
		To reduce the chance on unintended replacements, surround keywords with a marker.  Parenthesis are safe when used with Microsoft's Word, PowerPoint, Outlook and PowerPoint documents.
		For example: (CompanyName)
				
###	OUTPUTS
		One output file is generated per source document file, by default in a subfolder called '.\Reports\'.  Use -OutFolderPath to specify an alternate location.  The output file names are in the format of: 
			<source file base name>[-<execution source>]-<date/time/timezone stamp>[-<file name tag>].<Extension>
			
		If parameter -Debug or -Verbose is specified, then a second file, a PowerShell transcript (.LOG), is created in the same location.
			
###	PARAMETER Path String[]
		Specifies a path to Microsoft PowerPoint compatible document file pathname. Wildcards are permitted. The default location is the current directory.
			
###	PARAMETER FindReplacePath String
		Specifies a path to one Comma Separated Value (CSV) FindReplace file. The CSV must have at least two column headings (case insensitive), all other columns are ignored: 
		Find,Replace

		
###	PARAMETER Attributes FileAttributes
		Gets files and folders with the specified attributes. This parameter supports all attributes and lets you specify complex combinations of attributes.

		For example, to get non-system files (not directories) that are encrypted or compressed, type:
			Get-ChildItem -Attributes !Directory+!System+Encrypted, !Directory+!System+Compressed

		To find files and folders with commonly used attributes, you can use the Attributes parameter, or the Directory, File, Hidden, ReadOnly, and System switch parameters.

		The Attributes parameter supports the following attributes: Archive, Compressed, Device, Directory, Encrypted, Hidden, Normal, NotContentIndexed, Offline, ReadOnly, ReparsePoint, SparseFile, System, and Temporary. For a description of these attributes, see the FileAttributes enumeration at http://go.microsoft.com/fwlink/?LinkId=201508.

		Use the following operators to combine attributes.
			!    NOT
			+    AND
			,    OR
		No spaces are permitted between an operator and its attribute. However, spaces are permitted before commas.

		You can use the following abbreviations for commonly used attributes:
			D    Directory
			H    Hidden
			R    Read-only
			S     System

###	PARAMETER Directory SwitchParameter
		Gets directories (folders).  

		To get only directories, use the Directory parameter and omit the File parameter. To exclude directories, use the File parameter and omit the Directory parameter, or use the Attributes parameter. 

		To get directories, use the Directory parameter, its "ad" alias, or the Directory attribute of the Attributes parameter.

###	PARAMETER File SwitchParameter
		Gets files. 

		To get only files, use the File parameter and omit the Directory parameter. To exclude files, use the Directory parameter and omit the File parameter, or use the Attributes parameter.

		To get files, use the File parameter, its "af" alias, or the File value of the Attributes parameter.

###	PARAMETER Hidden SwitchParameter
		Gets only hidden files and directories (folders).  By default, Get-ChildItem gets only non-hidden items, but you can use the Force parameter to include hidden items in the results.

		To get only hidden items, use the Hidden parameter, its "h" or "ah" aliases, or the Hidden value of the Attributes parameter. To exclude hidden items, omit the Hidden parameter or use the Attributes parameter.

###	PARAMETER ReadOnly SwitchParameter
		Gets only read-only files and directories (folders).  

		To get only read-only items, use the ReadOnly parameter, its "ar" alias, or the ReadOnly value of the Attributes parameter. To exclude read-only items, use the Attributes parameter.

###	PARAMETER System SwitchParameter
		Gets only system files and directories (folders).

		To get only system files and folders, use the System parameter, its "as" alias, or the System value of the Attributes parameter. To exclude system files and folders, use the Attributes parameter.

###	PARAMETER Force SwitchParameter
		Gets hidden files and folders. By default, hidden files and folder are excluded. You can also get hidden files and folders by using the Hidden parameter or the Hidden value of the Attributes parameter.

###	PARAMETER UseTransaction SwitchParameter
		Includes the command in the active transaction. This parameter is valid only when a transaction is in progress. For more information, see about_Transactions.

###	PARAMETER Depth UInt32
		{{Fill Depth Description}}

###	PARAMETER Exclude String[]
		Specifies, as a string array, an item or items that this cmdlet excludes in the operation. The value of this parameter qualifies the Path parameter. Enter a path element or pattern, such as *.txt. Wildcards are permitted.

###	PARAMETER Filter String
		Specifies a filter in the provider's format or language. The value of this parameter qualifies the Path parameter. The syntax of the filter, including the use of wildcards, depends on the provider. Filters are more efficient than other parameters, because the provider applies them when retrieving the objects, rather than having Windows PowerShell filter the objects after they are retrieved.

###	PARAMETER Include String[]
		Specifies, as a string array, an item or items that this cmdlet includes in the operation. The value of this parameter qualifies the Path parameter. Enter a path element or pattern, such as *.txt. Wildcards are permitted.
		
		The default is MS PowerPoint supported file types:
			*.odp
			*.pot
			*.potm
			*.potx
			*.ppa
			*.ppam
			*.pps
			*.ppsm
			*.ppsx
			*.ppt
			*.pptm
			*.pptx
			*.pptx

		The Include parameter is effective only when the command includes the Recurse parameter or the path leads to the contents of a directory, such as C:\Windows\*, where the wildcard character specifies the contents of the C:\Windows directory.

###	PARAMETER LiteralPath String[]
		Specifies, as a string arrya, a path to one or more locations. Unlike the Path parameter, the value of the LiteralPath parameter is used exactly as it is typed. No characters are interpreted as wildcards. If the path includes escape characters, enclose it in single quotation marks. Single quotation marks tell Windows PowerShell not to interpret any characters as escape sequences.

###	PARAMETER Name SwitchParameter
		Indicates that this cmdlet gets only the names of the items in the locations. If you pipe the output of this command to another command, only the item names are sent.

###	PARAMETER Path String[]
		Specifies a path to one or more Microsoft PowerPoint compatible document. Wildcards are permitted. The default location is the current directory (.).

###	PARAMETER Recurse SwitchParameter
		Indicates that this cmdlet gets the items in the specified locations and in all child items of the locations.

		In Windows PowerShell 2.0 and earlier versions of Windows PowerShell, the Recurse parameter works only when the value of the Path parameter is a container that has child items, such as C:\Windows or C:\Windows\ , and not when it is an item does not have child items, such as C:\Windows\ .exe.

			
###	PARAMETER ExecutionSource
		Specifiy the script's execution environment source.  Must be either; 'msExchOrganizationName', 'ForestName', 'DomainName', 'ComputerName', or an arbitrary string including '' or $NULL.
		If msExchOrganizationName is requested, but there is no Microsoft Exchange organization, ForestName will be used.
		If ForestName is requested, but there is no forest, DomainName will be used.  The forest name is of the executing computer's domain membership.  
		If the DomainName is requested, but the computer is not a domain member, ComputerName is used.  The domain name is of the executing computer's domain membership.  
		An arbitrary string can be used in the case where the Microsoft Exchange organization name, forest name or domain name is too generic (e.g. 'EMAIL', 'CORP' or 'ROOT').
		Defaults is msExchOrganizationName.
		
###	PARAMETER OutFileNameTag
		Optional comment string added to the end of the output file name.
		
###	PARAMETER OutFolderPath
		Specify which folder path to write the outfile.  Supports UNC and relative reference to the current script folder.  Except for UNC paths, this function will attempt to create and compress the output folder if it doesn't exist.  The default is .\Reports subfolder.  
		
###	PARAMETER AlertOnly
		When enabled, only unhealthy items are reported and the optional mail subject will contain 'alert' instead of 'report', and if there are no unhealthy items there is no output.  
		
###	PARAMETER MailFrom
		Optionally specify the address from which the mail is sent. Enter a name (optional) and e-mail address, such as 'Name <LocalPart@domain.com>'. 
		
###	PARAMETER MailTo
		Optioanlly specify the addresses to which the mail is sent. Enter names (optional) and the e-mail address, such as 'Name1 <LocalPart1@domain.com>','Name2 <LocalPart2@domain.com>'. 
		
###	PARAMETER MailServer
		Optionally specify the name of the SMTP server that sends the mail message.
		
###	PARAMETER CompressAttachmentLargerThan
		Optionally specify that when a file attachment size is over this limit that it should be compressed when e-mailed.  The default is 5MB.  There is no guarantee the compressed attachment will be below the sender or recipeint's message size limit.  
			
###	EXAMPLE
		Description
		-----------
		If find/replace file '.\MyFindReplace.csv's finds matches in Microsoft PowerPoint document file '.\MySource.docx' then a new document '.\Reports\MySource-Mine-20190302T235959+12.docx file will be creatd.
		
		New-MSPowerPointFindReplaceTextFile -Path .\MySource.docx -FindReplacePath .\MyFindReplace.csv -ExecutionSource Mine
			
###	NOTE
		Author: Terry E Dow
		Creation Date: 2018-03-30
				
		Warning from Microsoft:
			Considerations for server-side Automation of Office https://support.microsoft.com/en-us/help/257757/considerations-for-server-side-automation-of-office

		Reference:
--
##	Function New-MSWordFindReplaceTextFile

###	SYNOPSIS
		When provided with an open Microsoft Word document, executes a set of text only FindText/ReplaceWith operations throughout the whole document.

###	DESCRIPTION	
		When provided with an open Microsoft Word document file name (wildcards are permitted), executes a set of text only FindText/ReplaceWith operations throughout the whole document.  
			
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
			
		Microsoft's Word Range.Find operation performs a simple text match.  There is no support wildcard or regular expressions [RegEx].  Formatting of the FindText is preserved.  
		To reduce the chance on unintended replacements, surround keywords with a marker.  Parenthesis are safe when used with Microsoft's Word, Excel, Outlook and PowerPoint documents.
		For example: (CompanyName)
			
###	OUTPUTS
		One output file is generated per source document file, by default in a subfolder called '.\Reports\'.  Use -OutFolderPath to specify an alternate location.  The output file names are in the format of: 
			<source file base name>[-<execution source>]-<date/time/timezone stamp>[-<file name tag>].<Extension>
			
		If parameter -Debug or -Verbose is specified, then a second file, a PowerShell transcript (.LOG), is created in the same location.
			
###	PARAMETER Path String[]
		Specifies a path to Microsoft Word compatible document file pathname. Wildcards are permitted. The default location is the current directory.
			
###	PARAMETER FindReplacePath String
		Specifies a path to one Comma Separated Value (CSV) FindReplace file. The CSV must have at least two column headings (case insensitive), all other columns are ignored: 
		Find,Replace

		
###	PARAMETER Attributes FileAttributes
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

###	PARAMETER Directory SwitchParameter
		Gets directories (folders).  

		To get only directories, use the Directory parameter and omit the File parameter. To exclude directories, use the File parameter and omit the Directory parameter, or use the Attributes parameter. 

		To get directories, use the Directory parameter, its "ad" alias, or the Directory attribute of the Attributes parameter.

###	PARAMETER File SwitchParameter
		Gets files. 

		To get only files, use the File parameter and omit the Directory parameter. To exclude files, use the Directory parameter and omit the File parameter, or use the Attributes parameter.

		To get files, use the File parameter, its "af" alias, or the File value of the Attributes parameter.

###	PARAMETER Hidden SwitchParameter
		Gets only hidden files and directories (folders).  By default, Get-ChildItem gets only non-hidden items, but you can use the Force parameter to include hidden items in the results.

		To get only hidden items, use the Hidden parameter, its "h" or "ah" aliases, or the Hidden value of the Attributes parameter. To exclude hidden items, omit the Hidden parameter or use the Attributes parameter.

###	PARAMETER ReadOnly SwitchParameter
		Gets only read-only files and directories (folders).  

		To get only read-only items, use the ReadOnly parameter, its "ar" alias, or the ReadOnly value of the Attributes parameter. To exclude read-only items, use the Attributes parameter.

###	PARAMETER System SwitchParameter
		Gets only system files and directories (folders).

		To get only system files and folders, use the System parameter, its "as" alias, or the System value of the Attributes parameter. To exclude system files and folders, use the Attributes parameter.

###	PARAMETER Force SwitchParameter
		Gets hidden files and folders. By default, hidden files and folder are excluded. You can also get hidden files and folders by using the Hidden parameter or the Hidden value of the Attributes parameter.

###	PARAMETER UseTransaction SwitchParameter
		Includes the command in the active transaction. This parameter is valid only when a transaction is in progress. For more information, see about_Transactions.

###	PARAMETER Depth UInt32
		{{Fill Depth Description}}

###	PARAMETER Exclude String[]
		Specifies, as a string array, an item or items that this cmdlet excludes in the operation. The value of this parameter qualifies the Path parameter. Enter a path element or pattern, such as *.txt. Wildcards are permitted.

###	PARAMETER Filter String
		Specifies a filter in the provider's format or language. The value of this parameter qualifies the Path parameter. The syntax of the filter, including the use of wildcards, depends on the provider. Filters are more efficient than other parameters, because the provider applies them when retrieving the objects, rather than having Windows PowerShell filter the objects after they are retrieved.

###	PARAMETER Include String[]
		Specifies, as a string array, an item or items that this cmdlet includes in the operation. The value of this parameter qualifies the Path parameter. Enter a path element or pattern, such as *.txt. Wildcards are permitted.
		
		The default is MS Word supported file types:
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

		The Include parameter is effective only when the command includes the Recurse parameter or the path leads to the contents of a directory, such as C:\Windows\*, where the wildcard character specifies the contents of the C:\Windows directory.

###	PARAMETER LiteralPath String[]
		Specifies, as a string arrya, a path to one or more locations. Unlike the Path parameter, the value of the LiteralPath parameter is used exactly as it is typed. No characters are interpreted as wildcards. If the path includes escape characters, enclose it in single quotation marks. Single quotation marks tell Windows PowerShell not to interpret any characters as escape sequences.

###	PARAMETER Name SwitchParameter
		Indicates that this cmdlet gets only the names of the items in the locations. If you pipe the output of this command to another command, only the item names are sent.

###	PARAMETER Path String[]
		Specifies a path to one or more Microsoft Word compatible document. Wildcards are permitted. The default location is the current directory (.).

###	PARAMETER Recurse SwitchParameter
		Indicates that this cmdlet gets the items in the specified locations and in all child items of the locations.

		In Windows PowerShell 2.0 and earlier versions of Windows PowerShell, the Recurse parameter works only when the value of the Path parameter is a container that has child items, such as C:\Windows or C:\Windows\ , and not when it is an item does not have child items, such as C:\Windows\ .exe.

			
###	PARAMETER ExecutionSource
		Specifiy the script's execution environment source.  Must be either; 'msExchOrganizationName', 'ForestName', 'DomainName', 'ComputerName', or an arbitrary string including '' or $NULL.
		If msExchOrganizationName is requested, but there is no Microsoft Exchange organization, ForestName will be used.
		If ForestName is requested, but there is no forest, DomainName will be used.  The forest name is of the executing computer's domain membership.  
		If the DomainName is requested, but the computer is not a domain member, ComputerName is used.  The domain name is of the executing computer's domain membership.  
		An arbitrary string can be used in the case where the Microsoft Exchange organization name, forest name or domain name is too generic (e.g. 'EMAIL', 'CORP' or 'ROOT').
		Defaults is msExchOrganizationName.
		
###	PARAMETER OutFileNameTag
		Optional comment string added to the end of the output file name.
		
###	PARAMETER OutFolderPath
		Specify which folder path to write the outfile.  Supports UNC and relative reference to the current script folder.  Except for UNC paths, this function will attempt to create and compress the output folder if it doesn't exist.  The default is .\Reports subfolder.  
		
###	PARAMETER AlertOnly
		When enabled, only unhealthy items are reported and the optional mail subject will contain 'alert' instead of 'report', and if there are no unhealthy items there is no output.  
		
###	PARAMETER MailFrom
		Optionally specify the address from which the mail is sent. Enter a name (optional) and e-mail address, such as 'Name <LocalPart@domain.com>'. 
		
###	PARAMETER MailTo
		Optioanlly specify the addresses to which the mail is sent. Enter names (optional) and the e-mail address, such as 'Name1 <LocalPart1@domain.com>','Name2 <LocalPart2@domain.com>'. 
		
###	PARAMETER MailServer
		Optionally specify the name of the SMTP server that sends the mail message.
		
###	PARAMETER CompressAttachmentLargerThan
		Optionally specify that when a file attachment size is over this limit that it should be compressed when e-mailed.  The default is 5MB.  There is no guarantee the compressed attachment will be below the sender or recipeint's message size limit.  
			
###	EXAMPLE
		Description
		-----------
		If find/replace file '.\MyFindReplace.csv's finds matches in Microsoft Word document file '.\MySource.docx' then a new document '.\Reports\MySource-Mine-20190302T235959+12.docx file will be creatd.
		
		New-MSWordFindReplaceTextFile -Path .\MySource.docx -FindReplacePath .\MyFindReplace.csv -ExecutionSource Mine
			
###	NOTE
		Author: Terry E Dow
		Creation Date: 2018-03-02

		Warning from Microsoft:
			Considerations for server-side Automation of Office https://support.microsoft.com/en-us/help/257757/considerations-for-server-side-automation-of-office

		Reference:
			https://codereview.stackexchange.com/questions/174455/powershell-script-to-find-and-replace-in-word-document-including-header-footer
			https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.find.execute
			https://learn-powershell.net/2014/12/31/beginning-with-powershell-and-word/
			https://wordmvp.com/FAQs/MacrosVBA/FindReplaceAllWithVBA.htm
			https://wordribbon.tips.net/T011489_Including_Headers_and_Footers_when_Selecting_All.html				

--
##	Function Update-MSExcelFindReplaceTextDocument

###	SYNOPSIS
		When provided with an open Microsoft Excel document, executes a set of text only FindText/ReplaceWith operations throughout the whole document.

###	DESCRIPTION	
		When provided with an open Microsoft Excel document (Excel.Application's Workbooks.Open), executes a set of text only FindText/ReplaceWith operations throughout the whole document.  

		Microsoft Excel's Range.Replace operation performs a simple text match. There is no support wildcard or regular expressions [RegEx]. Formatting of the FindText is preserved.  
		To reduce the chance on unintended replacements, surround keywords with a marker.  Parenthesis are safe when used with Microsoft's Word, Excel, Outlook and PowerPoint documents.
		For example: (CompanyName)
			
###	OUTPUTS
		The output of this function is the modifications executed on the open document.
		The returned value from this function is the number of replacements made.  
			
###	PARAMETER Path [Microsoft.Office.Interop.Excel.Workbooks]
		An open Microsoft Excel document (Excel.Workbooks.Open).  If FindText is found this document will be modified.  
			
###	PARAMETER FindReplacePath String
		A hash table @{ [String] FindText, [String] ReplaceWith } pairs.  Each FindText ReplaceWith operation will be executed througout the whole document.  

###	EXAMPLE
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

###	NOTE
		Author: Terry E Dow
		Creation Date: 2018-03-02
		
		Warning from Microsoft:
			Considerations for server-side Automation of Office https://support.microsoft.com/en-us/help/257757/considerations-for-server-side-automation-of-office
		
		Reference:
			https://codereview.stackexchange.com/questions/174455/powershell-script-to-find-and-replace-in-word-document-including-header-footer
			https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.find.execute
			https://learn-powershell.net/2014/12/31/beginning-with-powershell-and-word/
			https://wordmvp.com/FAQs/MacrosVBA/FindReplaceAllWithVBA.htm
			https://wordribbon.tips.net/T011489_Including_Headers_and_Footers_when_Selecting_All.html				

---
##	Function Update-MSOutlookFindReplaceTextDocument

###	SYNOPSIS
		When provided with an open Microsoft Outlook document, executes a set of text only FindText/ReplaceWith operations throughout the whole document.

###	DESCRIPTION	
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
						
###	OUTPUTS
		The output of this function is the modifications executed on the open document.
		The returned [BOOL] from this function is TRUE if any of the FindText is updated, and FALSE if none of the the FindText was not found.
			
###	PARAMETER Path [Microsoft.Office.Interop.Outlook.MailItemClass]
		An open Microsoft Outlook document (Outlook.Application's Documents.Open).  If FindText is found this document will be modified.  
			
###	PARAMETER FindReplacePath String
		A hash table @{ [String] FindText, [String] ReplaceWith } pairs.  Each FindText ReplaceWith operation will be executed througout the whole document.  

###	EXAMPLE
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

###	NOTE
		Author: Terry E Dow
		Creation Date: 2018-03-17

		Warning from Microsoft:
			Considerations for server-side Automation of Office https://support.microsoft.com/en-us/help/257757/considerations-for-server-side-automation-of-office
			
		Reference:
			Microsoft Outlook Constants https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa219371(v=office.11)
			[MS-OXMSG]: Outlook Item (.msg) File Format https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxmsg/b046868c-9fbf-41ae-9ffb-8de2bd4eec82
			System.Net.Mail Namespace https://docs.microsoft.com/en-us/dotnet/api/system.net.mail?redirectedfrom=MSDN&view=netframework-4.7.2

---
##	Function Update-MSPowerPointFindReplaceTextDocument
	
###	SYNOPSIS
		When provided with an open Microsoft PowerPoint document, executes a set of text only FindText/ReplaceWith operations throughout the whole document.

###	DESCRIPTION	
		When provided with an open Microsoft PowerPoint document (PowerPoint.Application's Workbooks.Open), executes a set of text only FindText/ReplaceWith operations throughout the whole document.  

		Microsoft PowerPoint's Range.Replace operation performs a simple text match. There is no support wildcard or regular expressions [RegEx]. Formatting of the FindText is preserved.  
		To reduce the chance on unintended replacements, surround keywords with a marker.  Parenthesis are safe when used with Microsoft's Word, PowerPoint, Outlook and PowerPoint documents.
		For example: (CompanyName)
			
###	OUTPUTS
		The output of this function is the modifications executed on the open document.
		The returned value from this function is the number of replacements made.  
			
###	PARAMETER Path [Microsoft.Office.Interop.PowerPoint.Workbooks]
		An open Microsoft PowerPoint document (PowerPoint.Workbooks.Open).  If FindText is found this document will be modified.  
			
###	PARAMETER FindReplacePath String
		A hash table @{ [String] FindText, [String] ReplaceWith } pairs.  Each FindText ReplaceWith operation will be executed througout the whole document.  

###	EXAMPLE
		Description
		-----------
			This example opens a Mircosoft PowerPoint application,
			opens a MS PowerPoint document named '.\MySource.xlsx' in read only mode,
			creates a substitution hash table with 2 entries,
			and then calls this function.
			After the document is saved to another file name,
			and closed.  
		
			$PowerPointApp = New-Object -ComObject PowerPoint.Application
			$PowerPointApp.Visible = $FALSE
			$PowerPointApp.DisplayAlerts = $FALSE
			$document = $PowerPointApp.Workbooks.Open( '.\MySource.xlsx', 0, $TRUE ) # Filename, UpdateLinks, ReadOnly
			
			$findReplaceTable = @{ 'INCORRECT' = 'correct'; '(Field)' = 'MyFieldValue' }
			
			Update-MSPowerPointFindReplaceTextDocument -Document $document -FindReplaceTable $findReplaceTable
			
			$document.Close( $TRUE, $outFilePathName ) # SaveChanges, Filename
			$PowerPointApp.Quit()

###	NOTE
		Author: Terry E Dow
		Creation Date: 2018-03-30

		Warning from Microsoft:
			Considerations for server-side Automation of Office https://support.microsoft.com/en-us/help/257757/considerations-for-server-side-automation-of-office
		
		Reference:

--
##	Function Update-MSWordFindReplaceTextDocument 

###	SYNOPSIS
		When provided with an open Microsoft Word document, executes a set of text only FindText/ReplaceWith operations throughout the whole document.

###	DESCRIPTION	
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
			
###	OUTPUTS
		The output of this function is the modifications executed on the open document.
		The returned value from this function is the number of replacements made.  
			
###	PARAMETER Path [Microsoft.Office.Interop.Word.DocumentClass]
		An open Microsoft Word document (Word.Application's Documents.Open).  If FindText is found this document will be modified.  
			
###	PARAMETER FindReplacePath String
		A hash table @{ [String] FindText, [String] ReplaceWith } pairs.  Each FindText ReplaceWith operation will be executed througout the whole document.  

###	EXAMPLE
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

###	NOTE
		Author: Terry E Dow
		Creation Date: 2018-03-02
		
		Warning from Microsoft:
			Considerations for server-side Automation of Office https://support.microsoft.com/en-us/help/257757/considerations-for-server-side-automation-of-office

		Reference:
			https://codereview.stackexchange.com/questions/174455/powershell-script-to-find-and-replace-in-word-document-including-header-footer
			https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.find.execute
			https://learn-powershell.net/2014/12/31/beginning-with-powershell-and-word/
			https://wordmvp.com/FAQs/MacrosVBA/FindReplaceAllWithVBA.htm
			https://wordribbon.tips.net/T011489_Including_Headers_and_Footers_when_Selecting_All.html				
