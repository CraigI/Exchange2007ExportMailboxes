$ErrorActionPreference = "silentlycontinue"
$global:MainMenuSelect = $NULL
$global:OUMenuSelect = $NULL
$global:MyUserName = [Environment]::UserName
$global:ExportDate = (Get-Date).AddDays(-90)
#Puts $global:ExportDate into the date format in which the Exchange export-mailbox is looking for
$global:ExportDate = (get-date $global:ExportDate -format d)

function MainMenu
{
	Clear-Host
	Write-Host "===== MAILBOX EXPORTER - MAIN MENU =====" -foregroundcolor "yellow"
	Write-Host " "
	Write-Host "Please select one of the following options." -foregroundcolor "yellow"
	Write-Host "0. Get me out of here!" -foregroundcolor "yellow"
	Write-Host "1. Export a single users mailbox" -foregroundcolor "yellow"
	Write-Host "2. Export an entire OU" -foregroundcolor "yellow"
	Write-Host " "
	$global:MainMenuSelect = Read-Host "Your option [1,2,..]"
	if ($global:MainMenuSelect -eq "0"){exit}
	if ($global:MainMenuSelect -eq "1"){SingleUserExport}
	if ($global:MainMenuSelect -eq "2"){OUExportMenu}
	else
	{
		Write-Host "Did you bother reading? Please select a correct option." -foregroundcolor "magenta"
		$TryAgain = Read-Host "Try again? [y/n]"
		if ($TryAgain -eq "n"){exit}n
		else{MainMenu}		
	}
}
function OUExportMenu
{
	Clear-Host
	Write-Host "===== MAILBOX EXPORTER - OU MENU =====" -foregroundcolor "yellow"
	Write-Host " "
	Write-Host "Please select one of the following options." -foregroundcolor "yellow"
	Write-Host "0. Return to Main Menu" -foregroundcolor "yellow"
	Write-Host "1. Single Stage OU Export" -foregroundcolor "yellow"
	Write-Host "2. Mutli Stage OU Export" -foregroundcolor "yellow"
	Write-Host " "
	$global:OUMenuSelect = Read-Host "Your option [1,2,..]"
	if ($global:OUMenuSelect -eq "0"){MainMenu}
	if ($global:OUMenuSelect -eq "1"){OUSingleStageExport}
	if ($global:OUMenuSelect -eq "2"){OUMultiStageExport}
	else
	{
		Write-Host "Did you bother reading? Please select a correct option." -foregroundcolor "magenta"
		$TryAgain = Read-Host "Try again? [y/n]"
		if ($TryAgain -eq "n"){exit}
		else{OUExportMenu}
	}
}
function SingleUserExport
{
	$username = Read-Host "Enter Username to Export"
	$PathToPST = Read-Host "Enter Path to Save PST (can be UNC)"
	Write-Host " "
	Write-Host "The user $global:MyUserName will be granted full access to the $username mailbox" -foregroundcolor "yellow"
	Write-Host "  and placed in the $PathToPST directory." -foregroundcolor "yellow"
	$IsCorrect = Read-Host "Is this correct? [y/n]"
	If($IsCorrect -ne "y"){SingleUserExport}
	else
	{
		$testusername = Get-User $username
		if($testusername.DisplayName -ne $NULL)
		{
			if((Test-Path $PathToPST) -eq $true)
			{
				AddAdminRights($username)
				export-mailbox -Identity $username -PSTFolderPath $PathToPST -Confirm:$false -ExcludeFolders "\RSS Feeds" -BadItemLimit 9999
				RemoveAdminRights($username)
			}
			Else
			{
				Write-Host "You apparently can't type, the PST path doesn't exist." -foregroundcolor "magenta"
				$TryAgain = Read-Host "Try again? [y/n]"
				if($TryAgain -eq "n")
				{MainMenu}
				else
				{
					Clear-Host
					$username = $NULL
					$PathToPST = $NULL
					SingleUserExport
				}
			}
		}
		else
		{
			Write-Host "Username Fat Fingered, user does not exist" -foregroundcolor "magenta"
			$TryAgain = Read-Host "Try again? [y/n]"
			if($TryAgain -eq "n")
			{MainMenu}
			else
			{
				Clear-Host
				$username = $NULL
				$PathToPST = $NULL
				SingleUserExport
			}
		}
	}
}
function OUSingleStageExport
{
	Write-Host "=======================================" -foregroundcolor "yellow"
	Write-Host "You will be exporting an entire OU in a" -foregroundcolor "yellow"
	Write-Host "single stage with this option." -foregroundcolor "yellow"
	Write-Host "=======================================" -foregroundcolor "yellow"
	Write-Host "Would you like to continue?" -foregroundcolor "yellow"
	$StartOver = Read-Host "[y/n]"
	if ($StartOver -eq "n"){OUExportMenu}
	else
	{		
		$OUName = Read-Host "Enter Full OU name"
		$PathToPST = Read-Host "Enter Path to Save PST (can be UNC)"
		Write-Host " "
		Write-Host "The user $global:MyUserName will be granted full access to all mailboxes in" -foregroundcolor "yellow"
		Write-Host "  the following OU: $OUName" -foregroundcolor "yellow"
		Write-Host "  and placed in the $PathToPST folder." -foregroundcolor "yellow"
		$IsCorrect = Read-Host "Is this correct? [y/n]"
		If($IsCorrect -eq "n"){OUSingleStageExport}
		else
		{
			$testOU = Get-Mailbox -OrganizationalUnit $OUName
			if($testOU -ne $NULL -AND ((Test-Path $PathToPST) -eq $true))
			{
				AddAdminRights($OUName)
				get-mailbox -OrganizationalUnit $OUName | export-mailbox -PSTFolderPath $PathToPST -Confirm:$false -ExcludeFolders "\RSS Feeds" -BadItemLimit 9999 -MaxThreads 6
				RemoveAdminRights($OUName)
			}
			else
			{
				Write-Host "You apparently can't type, the OU or folder doesn't exist." -foregroundcolor "magenta"
				$TryAgain = Read-Host "Try again? [y/n]"
				if($TryAgain -eq "n")
				{OUExportMenu}
				else
				{
					Clear-Host
					$OUName = $NULL
					$PathToPST = $NULL
					$testOU = $NULL
					OUSingleStageExport
				}
			}
		}
	}
}
function OUMultiStageExport
{
	Write-Host "=======================================" -foregroundcolor "yellow"
	Write-Host "You will be exporting an entire OU in " -foregroundcolor "yellow"
	Write-Host "three stages. The stages are as follows: Would you" -foregroundcolor "yellow"
	Write-Host "Stage 1 = Email Only, last 90 days (excludes special folders)" -foregroundcolor "yellow"
	Write-Host "Stage 2 = (special folders) \Calendar, \Contacts, \Journal, \Notes, \Tasks" -foregroundcolor "yellow"
	Write-Host "Stage 3 = Emails older than 90 days (excludes special folders)" -foregroundcolor "yellow"
	Write-Host " "
	Write-Host "Stage 1 & 2 will be merged to make 1 PST per user." -foregroundcolor "yellow"
	Write-Host "Stage 3 will be in a seperate PST." -foregroundcolor "yellow"
	Write-Host "90 Days ago it was $global:ExportDate" -foregroundcolor "yellow"
	Write-Host "=======================================" -foregroundcolor "yellow"
	Write-Host "Would you like to continue?" -foregroundcolor "yellow"
	$StartOver = Read-Host "[y/n]"
	if ($StartOver -eq "n"){OUExportMenu}
	else
	{		
		$OUName = Read-Host "Enter Full OU name"
		$PathToPST = Read-Host "Enter Path to Save PST (can be UNC)"
		Write-Host " "
		Write-Host "The user $global:MyUserName will be granted full access to all mailboxes in" -foregroundcolor "yellow"
		Write-Host "  the following OU: $OUName" -foregroundcolor "yellow"
		Write-Host "  and two new folders will be created under $PathToPST folder." -foregroundcolor "yellow"
		Write-Host "  Stage 1 & 2 location will be $PathToPST\Stage_1_2\" -foregroundcolor "yellow"
		Write-Host "  Stage 3 location will be $PathToPST\Stage_3\" -foregroundcolor "yellow"
		Write-Host " "
		Write-Host "  !!!Be sure the Stage subfolders to not exist!!!" -foregroundcolor "red"
		$IsCorrect = Read-Host "Is this correct? [y/n]"
		If($IsCorrect -eq "n"){OUSingleStageExport}
		else
		{
			$testOU = Get-Mailbox -OrganizationalUnit $OUName
			if($testOU -ne $NULL -AND ((Test-Path $PathToPST) -eq $true))
			{
				$Stage_1_2 = "$PathToPST\Stage_1_2"
				$Stage_3 = "$PathToPST\Stage_3"
				New-Item  $Stage_1_2 -type directory
				New-Item  $Stage_3 -type directory
				AddAdminRights($OUName)
				#Stage 1
				get-mailbox -OrganizationalUnit $OUName | export-mailbox -StartDate $global:ExportDate `
					-PSTFolderPath $Stage_1_2 -Confirm:$false -ExcludeFolders "\Junk E-Mail", "\RSS Feeds", `
					"\Calendar", "\Contacts", "\Journal", "\Notes", "\Tasks" -BadItemLimit 9999 -MaxThreads 6
				#Stage 2
				get-mailbox -OrganizationalUnit $OUName | export-mailbox -PSTFolderPath $Stage_1_2 -Confirm:$false `
					-IncludeFolders "\Calendar", "\Contacts", "\Journal", "\Notes", "\Tasks" -BadItemLimit 9999 -MaxThreads 6
				#Stage 3
				get-mailbox -OrganizationalUnit $OUName | export-mailbox -EndDate $global:ExportDate `
					-PSTFolderPath $Stage_3 -Confirm:$false -ExcludeFolders "\Junk E-Mail", "\RSS Feeds", `
					"\Calendar", "\Contacts", "\Journal", "\Notes", "\Tasks" -BadItemLimit 9999 -MaxThreads 6
				RemoveAdminRights($OUName)
			}
			else
			{
				Write-Host "You apparently can't type, the OU or folder doesn't exist." -foregroundcolor "magenta"
				$TryAgain = Read-Host "Try again? [y/n]"
				if($TryAgain -eq "n")
				{OUExportMenu}
				else
				{
					Clear-Host
					$OUName = $NULL
					$PathToPST = $NULL
					$testOU = $NULL
					OUSingleStageExport
				}
			}
		}
	}
}
function AddAdminRights($var1)
{
	If ($global:MainMenuSelect -eq "1")
	{
		add-mailboxpermission -Identity $var1 -user $global:MyUserName -Accessright Fullaccess -InheritanceType all
	}
	If ($global:MainMenuSelect -eq "2")
	{
		get-mailbox -OrganizationalUnit $var1 | add-mailboxpermission -user $global:MyUserName -Accessright Fullaccess -InheritanceType all
	}
}
function RemoveAdminRights($var1)
{
	If ($global:MainMenuSelect -eq "1")
	{
		remove-mailboxpermission -Identity $var1 -user $global:MyUserName -Accessright Fullaccess -InheritanceType all	
	}
	If ($global:MainMenuSelect -eq "2")
	{
		get-mailbox -OrganizationalUnit $var1 | remove-mailboxpermission -user $global:MyUserName -Accessright Fullaccess -InheritanceType all	
	}
}

MainMenu