# MSOfficeFindReplace

### DESCRIPTION	
	A set of PowerShell functions that performs bulk simple text replacements on Microsoft Office documents:
		Excel
		Outlook
		PowerPoint
		Word
	Microsoft Office®, Microsoft Excel®, Microsoft Outlook®, Microsoft PowerPoint®, and Microsoft Word® are a registered trademarks of Microsoft Corporation.  

	Simple text substitution only
		No wildcard support
		No regular expression [RegEx] support
	Source text formatting is retained 
	One substitution file can be used for all operations with all documents and document types
	Respective Microsoft Office document applications installation is required on executing workstation

### Installing MSOfficeFindReplace on Windows PowerShell
	Install MSOfficeFindReplace module by copying the MSOfficeFindReplace folder of the this package’s .\Modules into the workstation’s $Env:PSModulePath
	For a single user use: 
		Copy-Item -Recurse -Path '.\Modules' -Destination  "$($Env:USERPROFILE)\Documents\WindowsPowerShell\Modules"
	For all users use:
		Copy-Item -Recurse -Path '.\Modules' -Destination  "$($Env:ProgramFiles)\WindowsPowerShell\Modules"

### Installing MSOfficeFindReplace on PowerShell Core
	Install MSOfficeFindReplace module by copying the MSOfficeFindReplace folder of the this package’s .\Modules into the workstation’s $Env:PSModulePath
	For a single user use: 
		Copy-Item -Recurse -Path '.\Modules' -Destination  "$($Env:USERPROFILE)\Documents\PowerShell\Modules"
	For all users use:
		Copy-Item -Recurse -Path '.\Modules' -Destination  "$($Env:ProgramFiles)\PowerShell\Modules"

###	Functions Exported:
		New-MSOutlookFindReplaceTextFile
		New-MSWordFindReplaceTextFile
		Update-MSOutlookFindReplaceTextDocument
		Update-MSWordFindReplaceTextDocument
		
		New-MSExcelFindReplaceTextFile 
		New-MSOutlookFindReplaceTextFile
		New-MSPowerPointFindReplaceTextFile
		New-MSWordFindReplaceTextFile
		Update-MSExcelFindReplaceTextDocument
		Update-MSOutlookFindReplaceTextDocument
		Update-MSPowerPointFindReplaceTextDocument
		Update-MSWordFindReplaceTextDocument