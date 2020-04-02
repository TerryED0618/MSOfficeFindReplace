﻿#
# Module manifest for module 'MSOfficeFindReplace' 
#
# Generated by: Terry E Dow
#
# Generated on: 2019-03-02
# Last Updated on: 2020-04-01
#

@{
	# Script module or binary module file associated with this manifest.
	RootModule = 'MSOfficeFindReplace.psm1'

	# Version number of this module.
	ModuleVersion = '3.3.1'

	# ID used to uniquely identify this module
	GUID = 'f54d3289-d09a-4361-b38e-542c06349359'

	# Author of this module
	Author = 'Terry E Dow'

	# Company or vendor of this module
	#CompanyName = ''

	# Copyright statement for this module
	Copyright = @'
MIT License

Copyright (c) 2019 Terry E Dow [Terry E Dow](https://github.com/TerryED0618)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'@

	# Description of the functionality provided by this module
	Description = 'PowerShell module generate and test complex passwords.'

	# Minimum version of the Windows PowerShell engine required by this module
	PowerShellVersion = '3.0'

	# Name of the Windows PowerShell host required by this module
	# PowerShellHostName = ''

	# Minimum version of the Windows PowerShell host required by this module
	# PowerShellHostVersion = ''

	# Minimum version of Microsoft .NET Framework required by this module
	# DotNetFrameworkVersion = ''

	# Minimum version of the common language runtime (CLR) required by this module
	# CLRVersion = ''

	# Processor architecture (None, X86, Amd64) required by this module
	# ProcessorArchitecture = ''

	# Modules that must be imported into the global environment prior to importing this module
	# RequiredModules = @()

	# Assemblies that must be loaded prior to importing this module
	# RequiredAssemblies = @()

	# Script files (.ps1) that are run in the caller's environment prior to importing this module.
	# ScriptsToProcess = @()

	# Type files (.ps1xml) to be loaded when importing this module
	# TypesToProcess = @()

	# Format files (.ps1xml) to be loaded when importing this module
	# FormatsToProcess = @()

	# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
	# NestedModules = @()

	# Functions to export from this module
	FunctionsToExport = @( 
		'New-MSExcelFindReplaceTextFile', 
		'New-MSOutlookFindReplaceTextFile', 
		'New-MSPowerPointFindReplaceTextFile',
		'New-MSWordFindReplaceTextFile', 
		'Update-MSExcelFindReplaceTextDocument',
		'Update-MSOutlookFindReplaceTextDocument', 
		'Update-MSPowerPointFindReplaceTextDocument',
		'Update-MSWordFindReplaceMailMergeDocument',
		'Update-MSWordFindReplaceTextDocument' 
	)

	# Cmdlets to export from this module
	# CmdletsToExport = '*'

	# Variables to export from this module
	# VariablesToExport = '*'

	# Aliases to export from this module
	# AliasesToExport = '*'

	# DSC resources to export from this module
	# DscResourcesToExport = @()

	# List of all modules packaged with this module
	# ModuleList = @()

	# List of all files packaged with this module
	# FileList = @()

	# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
	PrivateData = @{
	
		PSData = @{
	
			# Tags applied to this module. These help with module discovery in online galleries.
			Tags = @( 'Office', 'MSWord', 'Word', 'Find', 'Replace' )
	
			# A URL to the license for this module.
			LicenseUri = 'https://github.com/TerryED0618/MSOfficeFindReplace/blob/master/LICENSE'
	
			# A URL to the main website for this project.
			ProjectUri = 'https://github.com/TerryED0618/MSOfficeFindReplace'
	
			# A URL to an icon representing this module.
			# IconUri = ''
	
			# ReleaseNotes of this module
			ReleaseNotes = @'
2019-03-16 3.0.0 Initial release supporting New-MSWordFindReplaceTextFile and Update-MSWordFindReplaceTextDocument in support of *.DOC, *.DOCM, *.DOCX, *.RTF, etc.
2019-03-17 3.1.0 Added New-MSOutlookFindReplaceTextFile and Update-MSOutlookFindReplaceTextDocument in support for *.MSG and *.EML files.  
2019-03-31 3.2.0 Added New-MSExcelFindReplaceTextFile, Update-MSExcelFindReplaceTextDocument in support of Microsoft Excel files.
2019-03-31 3.2.0 Added New-MSPowerPointFindReplaceTextFile, Update-MSPowerPointFindReplaceTextDocument' in support of Microsoft PowerPoint files.
2019-09-02 3.2.1 Updated Added to include exported functions
2019-09-02 3.2.1 Updated documentation
2019-09-29 3.3.0 Updated New-MSOutlookFindReplaceTextFile to support optionally saving message to Outlook Draft folder instead of to a file using the SaveToDrafts switch.
2020-04-01 3.3.1 Update private module New-OutFilePathBase.ps1.
'@
	
		} # End of PSData hashtable
	
	} # End of PrivateData hashtable

	# HelpInfo URI of this module
	# HelpInfoURI = ''

	# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
	# DefaultCommandPrefix = ''

}