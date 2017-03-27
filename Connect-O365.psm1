<#
.SYNOPSIS
.DESCRIPTION
.PARAMETER
.EXAMPLE
.NOTES
	Version: 1.2.11
	Updated: 10/17/2016
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

	if ( ($Username -AND $Password) -AND (Test-Path $Password) ) {
		$sPassword = Get-Content $Password | ConvertTo-SecureString
		$CredObj = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username,$sPassword
	}
	elseif ($Username -AND $Password) {
		$sPassword = ConvertTo-SecureString -String $Password -AsPlainText -force
		$CredObj = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username,$sPassword
	}
	else {
		if ($Username) {
			#PromptForCredential(Title,Message,Username,Domain)
			$CredObj = (Get-Host).UI.PromptForCredential('Office 365 Credentials','Please enter your Office 365 Admin Credentials',$Username,'')
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

function Connect-ExchangeOnline {
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

function Connect-SkypeOnline {
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

	Test-Prerequisites 'SkypeOnline'

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
			[ValidateSet('O365Admin','SkypeOnline')]
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
		'SkypeOnline' { 
			$Program = $InstalledPrograms | ? {$_.DisplayName -match "Skype for Business Online.*PowerShell"}

			if ($Program.VersionMajor -lt 6 -OR !$Program) {
				throw "Microsoft Skype for Business Online PowerSell Module not installed or outdated - major version should be 6 or greater."
			}
		}
	}
}

function Connect-AllO365 {
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
	Connect-ExchangeOnline -CredentialObject $CredentialObject

	if ($TenantName) {
		Connect-SkypeOnline -CredentialObject $CredentialObject -TenantName $TenantName
	}
	else {
		Connect-SkypeOnline -CredentialObject $CredentialObject	
	}
}

New-Alias -Name coa -Value Connect-O365Admin
New-Alias -Name ceo -Value Connect-ExchangeOnline
New-Alias -Name cso -Value Connect-SkypeOnline

Export-ModuleMember -Function New-SecureStringFile
Export-ModuleMember -Function Get-CredentialObject
Export-ModuleMember -Function Connect-AllO365
Export-ModuleMember -Function Connect-O365Admin -Alias coa
Export-ModuleMember -Function Connect-ExchangeOnline -Alias ceo
Export-ModuleMember -Function Connect-SkypeOnline -Alias cso