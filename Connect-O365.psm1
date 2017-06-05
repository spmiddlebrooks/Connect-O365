<#
.SYNOPSIS
.DESCRIPTION
.PARAMETER
.EXAMPLE
.NOTES
	Version: 1.3.2
	Updated: 6/5/2017
	Author : Scott Middlebrooks
.INPUTS
.OUTPUTS
.LINK
#>
function New-SecureStringFile {
	<#
	.SYNOPSIS
		Store passwords in files as secure strings
	.DESCRIPTION
		Create files to store secure string passwords for reuse in scripts or later Connect-* module sessions.
	.PARAMETER FilePath
		Full path of file where secure string will be stored.  Must be full path or an error will occur.
	.EXAMPLE
		New-SecureStringFile -FilePath C:\Secure\MyPassword.txt
	.EXAMPLE
		New-SecureStringFile -FilePath .\Password.txt
	.NOTES
		Version: 1.0
		Updated: 7/6/2016
		Author : Scott Middlebrooks
	.LINK
	#>

	[cmdletbinding()]
	param(
		[Parameter(Mandatory=$True,Position=0)]
			[ValidateNotNullOrEmpty()]
			[ValidateScript({
				if ( Test-Path (Split-Path $_) ) {$True}
				else {Throw 'Invalid path'}
			})]
			[string] $FilePath
	)

	( (Get-Host).UI.PromptForCredential('Office 365 Credentials','Please enter your Office 365 Admin Password','No Username Required','') ).Password | ConvertFrom-SecureString | Out-File $FilePath
}
function Get-CredentialObject {
	[cmdletbinding()]
	param(
		[Parameter(Mandatory=$False,Position=0)]
			[string] $Username,		
		[Parameter(Mandatory=$False,Position=1)]
			[string] $Password
	)

	$UsernameRegexMatch = $Username -match "^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$"

	if ( ($UsernameRegexMatch -AND $Password) -AND (Test-Path $Password) ) {
		$sPassword = Get-Content $Password | ConvertTo-SecureString
		$CredObj = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username,$sPassword
	}
	elseif ($UsernameRegexMatch -AND $Password) {
		$sPassword = ConvertTo-SecureString -String $Password -AsPlainText -force
		$CredObj = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username,$sPassword
	}
	else {
		if ($UsernameRegexMatch) {
			#PromptForCredential(Title,Message,Username,Domain)
			$CredObj = (Get-Host).UI.PromptForCredential('Office 365 Credentials','Please enter your Office 365 Admin Credentials',$Username,'')
		}
		elseif ($Username -AND ! $UsernameRegexMatch) {
			$wshell = New-Object -ComObject Wscript.Shell
			$null = $wshell.Popup("Username not a valid userPrincipalName.`nUsername should be of the form username@domain.com.`nPlease re-enter your credentials.",0,"Username Format Invalid",0x30)
			$CredObj = (Get-Host).UI.PromptForCredential('Office 365 Credentials','Please re-enter your Office 365 Admin Credentials','','')
		}
		else {
			$CredObj = (Get-Host).UI.PromptForCredential('Office 365 Credentials','Please enter your Office 365 Admin Credentials','','')
		}
	}

	Return $CredObj
}
function Connect-O365Admin {
	[cmdletbinding(DefaultParameterSetName='Username')]
	param (	
		[Parameter(Mandatory=$False,Position=0,ParameterSetName='CredentialObject')]
			$CredentialObject,
		[Parameter(Mandatory=$False,Position=0,ParameterSetName='Username')]
			[string] $Username='',
		[Parameter(Mandatory=$false,Position=1)]
			[string] $Password=''
	)

	if (-Not ($CredentialObject -AND $CredentialObject.GetType().Name -eq 'PSCredential') ) {
		$CredentialObject = Get-CredentialObject -Username $Username -Password $Password
	}

	Test-Prerequisites 'O365Admin'

	try {
		connect-msolservice -credential $CredentialObject
	}
	catch {
		throw "Could not connect to O365: $($_.Exception.Message)"
	}	
}

function Connect-O365Exchange {
	[cmdletbinding(DefaultParameterSetName='Username')]
	param (	
		[Parameter(Mandatory=$False,Position=0,ParameterSetName='CredentialObject')]
			$CredentialObject,
		[Parameter(Mandatory=$False,Position=0,ParameterSetName='Username')]
			[string] $Username='',
		[Parameter(Mandatory=$false,Position=1)]
			[string] $Password=''
	)

	if (-Not ($CredentialObject -AND $CredentialObject.GetType().Name -eq 'PSCredential') ) {
		$CredentialObject = Get-CredentialObject -Username $Username -Password $Password
	}

	try {
		$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $CredentialObject -Authentication Basic -AllowRedirection
		import-module (import-pssession $session -DisableNameChecking -prefix 'OL' -AllowClobber) -DisableNameChecking -prefix 'OL' -Global
		return $session
	}
	catch {
		throw "Could not connect to Exchange Online: $($_.Exception.Message)"
	}	
}

function Connect-O365Skype {
	[cmdletbinding(DefaultParameterSetName='Username')]
	param (	
		[Parameter(Mandatory=$False,Position=0,ParameterSetName='CredentialObject')]
			$CredentialObject,
		[Parameter(Mandatory=$False,Position=0,ParameterSetName='Username')]
			[string] $Username='',
		[Parameter(Mandatory=$false,Position=1)]
			[string] $Password='',
		[Parameter(Mandatory=$false,Position=3)]
			[string] $TenantName
	)

	if (-Not ($CredentialObject -AND $CredentialObject.GetType().Name -eq 'PSCredential') ) {
		$CredentialObject = Get-CredentialObject -Username $Username -Password $Password
	}

	Test-Prerequisites -ServiceName 'O365Skype'

	if ($TenantName) {
		$OverrideAdminDomain = $TenantName + '.onmicrosoft.com'
	}

	try {
		if ($OverrideAdminDomain) {
			$session = new-csonlinesession -credential $CredentialObject -OverrideAdminDomain $OverrideAdminDomain
		}
		else {
			$session = new-csonlinesession -credential $CredentialObject
		}
		import-module (import-pssession $session -DisableNameChecking -prefix 'OL' -AllowClobber) -DisableNameChecking -prefix 'OL' -Global
		return $session
	}
	catch {
		throw "Could not connect to Skype Online: $($_.Exception.Message)"
	}	
}

function Test-Prerequisites {
	[cmdletbinding()]
	param (
		[Parameter(Mandatory=$true,Position=0)]
			[ValidateSet('O365Admin','O365Skype')]
			[string] $ServiceName
	)

	$InstalledPrograms = Get-ItemProperty HKLM:\software\microsoft\windows\currentversion\uninstall\*

	switch ($ServiceName) {
		'O365Admin' {  
			$Program = $InstalledPrograms | ? {$_.DisplayName -match "Microsoft Online.*Sign-In Assistant"}

			if ($Program.VersionMajor -lt 7 -OR !$Program) {
				throw "Microsoft Online Service Sign-In Assistant is not instaled or outdated - major version should be 7 or greater."
			}
		}
		'O365Skype' { 
			$Program = $InstalledPrograms | ? {$_.DisplayName -match "Skype for Business Online.*PowerShell"}

			if ($Program.VersionMajor -lt 6 -OR !$Program) {
				throw "Microsoft Skype for Business Online PowerSell Module not installed or outdated - major version should be 6 or greater."
			}
		}
	}
}

function Connect-O365All {
	[cmdletbinding(DefaultParameterSetName='Username')]
	param (	
		[Parameter(Mandatory=$False,Position=0,ParameterSetName='CredentialObject')]
			$CredentialObject,
		[Parameter(Mandatory=$False,Position=0,ParameterSetName='Username')]
			[string] $Username='',
		[Parameter(Mandatory=$false,Position=1)]
			[string] $Password='',
		[Parameter(Mandatory=$false,Position=3)]
			[string] $TenantName
	)

	if (-Not ($CredentialObject -AND $CredentialObject.GetType().Name -eq 'PSCredential') ) {
		$CredentialObject = Get-CredentialObject -Username $Username -Password $Password
	}

	Connect-O365Admin -CredentialObject $CredentialObject
	Connect-O365Exchange -CredentialObject $CredentialObject

	if ($TenantName) {
		Connect-O365Skype -CredentialObject $CredentialObject -TenantName $TenantName
	}
	else {
		Connect-O365Skype -CredentialObject $CredentialObject	
	}
}

New-Alias -Name coa -Value Connect-O365Admin
New-Alias -Name coe -Value Connect-O365Exchange
New-Alias -Name cos -Value Connect-O365Skype

Export-ModuleMember -Function New-SecureStringFile
Export-ModuleMember -Function Get-CredentialObject
Export-ModuleMember -Function Connect-O365All
Export-ModuleMember -Function Connect-O365Admin -Alias coa
Export-ModuleMember -Function Connect-O365Exchange -Alias coe
Export-ModuleMember -Function Connect-O365Skype -Alias cos
