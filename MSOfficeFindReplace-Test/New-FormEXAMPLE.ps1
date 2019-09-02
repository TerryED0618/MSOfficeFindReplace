Import-Module MSOfficeFindReplace

Import-CSV -Path 'New-FormEXAMPLE.csv' |
	ForEach-Object {
		Write-Host $PSItem.Email
		
		$tempFileFullName = (New-TemporaryFile).FullName
	
		[PSCustomObject] @{
			'Find' = '(FirstName)';
			'Replace' = $PSItem.FirstName
		} | Export-CSV -Path $tempFileFullName
		
		[PSCustomObject] @{
			'Find' = '(LastName)';
			'Replace' = $PSItem.LastName
		} | Export-CSV -Path $tempFileFullName -Append

		[PSCustomObject] @{
			'Find' = '(Email)';
			'Replace' = $PSItem.Email
		} | Export-CSV -Path $tempFileFullName -Append
		
		# Calculated value not from form.
		[PSCustomObject] @{
			'Find' = '(PIN)';
			'Replace' = Get-Random -Minimum 1000 -Maximum 10000
		} | Export-CSV -Path $tempFileFullName -Append
		
		New-MSWordFindReplaceTextFile -Path '.\New-FormEXAMPLE.docx' -FindReplacePath $tempFileFullName -OutFileNameTag $PSItem.Email

		New-MSOutlookFindReplaceTextFile -Path '.\New-FormEXAMPLE.msg' -FindReplacePath $tempFileFullName -OutFileNameTag $PSItem.Email

		Remove-Item $tempFileFullName	
	}