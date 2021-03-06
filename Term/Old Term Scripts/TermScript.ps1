#include all scripts
. ".\RemoveGoogGroups.ps1"
. "C:\IT-Scripts\Helpers\sendMail.ps1"
write-Host -Foreground Gray "Loading AD Modules..."

#Commented out so you could just type in the username in the script
#if($args.length -ne 1)
#{
#	write-host "usage:  TermScript.ps1 <username>"
#	exit
#}

Add-PSSnapin Quest.ActiveRoles.ADManagement | out-null

<# Storing Creds here for service accounts... for now#>

<#Join.Me Admin Creds#>
$joinmeUser = "joinmeadmin@tradingtechnologies.com"
$joinmePass = "VZQk8m7xah"

Write-Host -ForegroundColor Green "___________                                        .__        __   
\__    ___/__________  _____     ______ ___________|__|______/  |_ 
  |    |_/ __ \_  __ \/     \   /  ___// ___\_  __ \  \____ \   __\
  |    |\  ___/|  | \/  Y Y  \  \___ \\  \___|  | \/  |  |_> >  |  
  |____| \___  >__|  |__|_|  / /____  >\___  >__|  |__|   __/|__|  
             \/            \/       \/     \/         |__|         "

Write-Host -Foreground Gray "------------------------------------------------------------------------------------------------------------------------"
Write-Host -Foreground Cyan "Termed Employee Information"
Write-Host -Foreground Gray "------------------------------------------------------------------------------------------------------------------------"
Write-Host -Foreground Black " "

$name = Read-Host "Enter the INTAD of the user you are running this script on in all lowercase"
connect-qadservice | out-null
$user = Get-QADUser $name


#This code is commented out as it does not work and needs to be rewritten

#first check to see if the employee is a tradeco employee
#connect-qadservice -service tradeco.int.tt.local | out-null 

#$user = Get-QADUser $name
#Write-Host "USERNAME $user"

<#
if($user -eq $null)
{
        $tradeco = "yes"
        write-host "TRADECO USER -manually go into AD and check to see if they have an INTAD account."
        Write-Host "If so, remove them from groups, disable them, and move them to the Disabled Users OU."
        try{Disable-QADUser $name | out-null}
        catch{}
        try{Remove-QADMemberOf $name -RemoveAll | out-null}
        catch{}
        try{Move-QADObject $name -NewParentContainer 'OU=Disabled Users,DC=int,DC=tt,DC=local' | out-null}
        catch{}


        disconnect-qadservice | out-null
        connect-qadservice | out-null
        $user = Get-QADUser $name
}
#>

 write-host "TRADECO USERS: manually go into AD and check to see if they have an INTAD account."
 Write-Host "If so, remove them from groups, disable them, and move them to the Disabled Users OU."


$user | out-default
Write-Host -Foreground Gray "------------------------------------------------------------------------------------------------------------------------"
Write-Host ""
Write-Host ""

$DispName = ""
$Location = ""
$suffix = ""

$DispName = $user.Name
$split = $DispName.split(" ")
$suffix = $split[2] 
try{$suffix = $suffix.substring(1,$suffix.length-2)}
catch{}
try{$location = $suffix.split("-")}
catch{}

$location = $location[1]
$Gmail = $user.Email

Write-Host -Foreground Yellow "Are you sure you want to start the term for"$DispName"?"
Read-Host 'If yes, press enter. If you do not wish to continue, please press Cntrl-C:' | Out-Null
Write-Host ""

Remove-QADMemberOf $name -RemoveAll | out-null

Disable-QADUser $name | out-null

Disconnect-QADService
Connect-QADService	
try{Disable-QADUser $name | out-null}
catch{}
try{Remove-QADMemberOf $name -RemoveAll | out-null}
catch{}
try{Move-QADObject $name -NewParentContainer 'OU=Disabled Users,DC=int,DC=tt,DC=local' | out-null}
catch{}



$username = $user.Name

sendMail -recipient "Joseph.Grankowski@TradingTechnologies.com" -subject "Disable MSDN Account For $DispName" -body "Please check MSDN to see if $DispName has an account, and if so disable it.  This person is no longer employed at Trading Technologies."

sendMail -recipient "IT-EnterpriseDevelopment@tradingtechnologies.com" -subject "Remove CMS Access For $DispName" -body "Please check AWS to see if $DispName has an account, and if so disable it.  This person is no longer employed at Trading Technologies."	 

sendMail -recipient "it-enterpriselabservices@tradingtechnologies.com" -subject "Remove AWS Access for $DispName" -body "Please check AWS for $Dispname and remove any accounts and keys if applicable. This person is no longer employed at Trading Technologies."

$ukArray = 'UK', 'FRA', 'GEN', 'FF'
$apArray = 'SIN', 'TO', 'HK', 'AU', 'IN'


if ($ukArray -contains $location)
{
	$ie = New-Object -ComObject InternetExplorer.Application
	$ie.Navigate("https://192.168.17.6/admin")
	$ie.Visible = $true
	Write-Host ""
	Write-Host ""
	Write-Host ""
	Write-Host ""
	Write-Host -Foreground Gray "Use the IE window that was launched to disconnect the user's VPN session."
	Read-Host 'Press Enter to continue...' | Out-Null
}
elseif ($apArray -contains $location)
{
	$ie = New-Object -ComObject InternetExplorer.Application
	$ie.Navigate("https://192.168.17.130/admin")
	$ie.Visible = $true
	Write-Host ""
	Write-Host ""
	Write-Host ""
	Write-Host ""
	Write-Host -Foreground Gray "Use the IE window that was launched to disconnect the user's VPN session."
	Read-Host 'Press Enter to continue...' | Out-Null
}
else
{
	$ie = New-Object -ComObject InternetExplorer.Application
	$ie.Navigate("https://192.168.250.50/admin")
	$ie.Visible = $true
	Write-Host ""
	Write-Host ""
	Write-Host ""
	Write-Host ""
	Write-Host -Foreground Gray "Use the IE window that was launched to remove end the user's VPN session (If Applicable)."
	Read-Host 'Press Enter to continue...' | Out-Null
}

#Move users to term group and change password
$GoogOrg = "Terms"
$GoogPword = "Default123!"

c:\gam\gam.exe update user $gmail org $GoogOrg password $GoogPword
	


#Set OOO message
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""	
	$answer=Read-Host "Do you wish to use a standard OOO reply (1),set a custom OOO reply (2), or do not set any OOO reply (3)?"
	if ($answer -eq 1)
	{
	    $oooreply=�This person is no longer an employee of Trading Technologies. For all TT related inquiries please contact us at +1(312)476-1000.�
	}
	elseif ($answer -eq 2)
	{
	    $oooreply=Read-Host "Please type the custom OOO message and then press enter"
	}
	
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""	
Write-Host -Foreground Gray "-----------------"
Write-Host -Foreground Cyan "OOO reply set to:" $oooreply
Write-Host -Foreground Gray "-----------------"	
	
	
#Set status of a users AutoReply and make sure it is set
if ($answer -eq 3)
	{
	    write-host "No OOO reply will be set for this user"
	}
else
	{
        c:\gam\gam.exe user $Gmail vacation on subject "This email is no longer valid" message $oooreply enddate 2154-1-18
	}
	

<# Google groups #>
removeGoogleGroups($Gmail)

<# Voicemail #>
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Navigate("172.17.24.35")
$ie.Visible = $true
Write-Host ""
Write-Host ""	
Write-Host -Foreground Gray "Use the IE window that was launched to check and remove the users voicemail."
Read-Host 'Press Enter to continue...' | Out-Null

[System.Diagnostics.Process]::Start("C:\Program Files (x86)\Avaya\Site Administration\bin\ASA.exe") | out-null
<# Extension #>
Write-Host ""
Write-Host ""	
Write-Host -Foreground Gray "Use your Avaya Site Administration Console to delete the user's extension."
Read-Host 'Press Enter to continue...' | Out-Null

<# Kill Wifi #>
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Navigate("http://meraki.cisco.com/")
$ie.Visible = $true

Write-Host ""
Write-Host ""
Write-Host -Foreground Gray "Use the IE window that was launched to kill any active wifi sessions for this user in Meraki."
Read-Host 'Press Enter to continue...' | Out-Null

<# SoHo Firewall #>
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Navigate("https://172.17.23.10")
$ie.Visible = $true
Write-Host ""
Write-Host ""
Write-Host -Foreground Gray "Use the IE window that was launched to remove the employee's user object from the Soho Firewall (If Applicable)."
Read-Host 'Press Enter to continue...' | Out-Null

<# JIRA #>
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Navigate("https://tradingtech.atlassian.net")
$ie.Visible = $true
Write-Host ""
Write-Host ""
Write-Host -Foreground Gray "Use the IE window that was launched to disable the user in JIRA (if applicable). (username: it-support, pword: on SS)."
Read-Host 'Press Enter to continue...' | Out-Null

<# Google Apps #>
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Navigate("https://admin.google.com/tradingtechnologies.com/AdminHome")
$ie.Visible = $true	

Write-Host ""
Write-Host ""	
Write-Host -Foreground Gray "Use the IE window that was launched to do the Following:"
Write-Host ""
Write-Host -Foreground Gray "   -Turn off 2-step Auth for the user"
Write-Host ""
Write-Host -Foreground Gray "   -Remove all app-specific passwords"
Write-Host ""
Write-Host -Foreground Gray "   -Reset Sign-in cookies"
Write-Host ""
Write-Host -Foreground Gray "   -Block, then delete, any devices assigned"
Write-Host ""
Write-Host -Foreground Gray "   -Transfer ownership of their documents."
Read-Host 'Press Enter to continue...' | Out-Null

<# Arkadin #>
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Navigate("https://docs.google.com/spreadsheets/d/1nhhIsR3ORu5OMLZvo-fxKPR4MPX6PUHVvlIlYELUUk0/edit?ts=568ec058#gid=123429426")
$ie.Visible = $true	
Write-Host ""
Write-Host ""	
Write-Host -Foreground Gray "Check the Arakin user spreadsheet so see if the employee has an account, update accordingly."
Read-Host 'Press Enter to continue...' | Out-Null

<# Wordpress #>
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Navigate("https://www.tradingtechnologies.com/wp-login.php?redirect_to=https%3A%2F%2Fwww.tradingtechnologies.com%2Fwp-admin%2F&reauth=1")
$ie.Visible = $true	
Write-Host ""
Write-Host ""	
Write-Host -Foreground Grey "Use the IT-Support Google login to access www.tradingtechnologies.com/wp-admin to remove the user if applicable.  Reach out to their manager to see who to transfer document ownership."
Read-Host 'Press Enter to continue...' | Out-Null

<# Join.Me #>
$ie = new-object -com "InternetExplorer.Application"
$ie.navigate("https://accounts.logme.in/login.aspx?clusterid=02&returnurl=https%3a%2f%2fwww.appguru.com%2ffederated%2floginsso&headerframe=https%3a%2f%2fwww.appguru.com%2fwelcome%2fcommon%2fproduct%2fpages%2fcls%2fheaderframe%2f&productframe=https%3a%2f%2fwww.appguru.com%2fwelcome%2fcommon%2fproduct%2fpages%2fcls%2fdefaultframe%2f&lang=en-US&skin=appguru&regtype=O&trackingproducttype=1&socialloginenabled=0")
$ie.visible = $true
$ie.height = 1000
while($ie.busy){Start-Sleep 1}
$ie.Document.getElementById("email").value=$joinmeUser
$ie.Document.getElementById("password").value=$joinMePass
$ie.Document.getElementById("btnSubmit").Click()
while($ie.busy){Start-Sleep 1}
$ie.navigate("https://www.appguru.com/policymanager")
while($ie.busy){Start-Sleep 1}
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""	
Write-Host -Foreground Gray "Use the IE window that was launched to do the Following:"
Write-Host ""
Write-Host -Foreground Gray "-Delete the User's Join.Me account."
Read-Host 'Press Enter to continue...' | Out-Null

<# Legal #>
Write-Host ""
Write-Host ""
$Legal=Read-Host "If this user was a member of the legal team, type YES and press enter (if not, just press enter to skip this)."

Write-Host ""

if($Legal -eq 'YES')
  {
  Write-Host "Please disable the users standard user account AND admin account in Anaqua (admin accounts follow the following naming convention: first.last+ADMIN@tradingtechnologies.com)."
  Write-Host ""
  $ie = New-Object -ComObject InternetExplorer.Application
  $ie.Navigate("https://tt.anaqua.com/anaqua")
  $ie.Visible = $true
  Read-Host 'Press Enter to continue...' | Out-Null
  }
  


<# Emergency #>
Write-Host ""
Write-Host ""
Write-Host -Foreground Red "IF THIS IS AN EMERGENCY TERM, REMOTE TO CCURE TO REMOVE THE USERS KEYCARD ACCESS FROM THE ADMINISTRATION CLIENT.  "
Write-Host "  (Login is IT SUPPORT , Password is enable)"
Write-Host ""
Write-Host "   Make a note of the users Card # from CCURE for the next step."
Write-Host ""
$keycard=Read-Host "Please enter the card number for the user (found on ccure) and press enter.  (If this is not an emergency term, just press enter.)"
$number = $keycard.length


if($number -ge 4)
{	
	function sendMail
	{     
		 Write-Host ""
		 Write-Host "Sending Email to Facilities to notify the building for removal of this users building access."
		 Write-Host ""
		
		 #SMTP server name
		 $smtpServer = "mail.int.tt.local"

		 #Creating a Mail object
		 $msg = new-object Net.Mail.MailMessage

		 #Creating SMTP server object
		 $smtp = new-object Net.Mail.SmtpClient($smtpServer)

		 #Email structure 
		 $msg.From = "noreply@tradingtechnologies.com"
		 $msg.ReplyTo = "it-support@tradingtechnologies.com"
		 $msg.To.Add("facilities@TradingTechnologies.com")
		 $msg.subject = "Contact building to remove door access for "+ $DispName
		 $msg.body = "Please contact the building to remove door access for "+ $DispName +".  The badge number for this person is "+ $keycard +"."

		 #Sending email 
		 $smtp.Send($msg)
	}	 
sendmail
}
else
{
Write-Host ""
Write-Host "No email will be sent to Facilities since this is not an Emergency Term."
Write-Host ""
}



disconnect-qadservice

#Disable TN User account (if applicable)

Write-Host "Checking the Test Network to see if the user has an account created, and disabling that account if applicable.."
Write-Host ""

Disconnect-QADService
Connect-QADService -service rostndc01.tn.local | out-null
try{Disable-QADUser $name | out-null}
catch{}
Disconnect-QADService


Write-Host -Foreground Gray "--------------------------------------------------------------------------------"
Write-Host -Foreground Cyan "You have successfully run the term script for"$DispName
Write-Host -Foreground Cyan	""
Write-Host -Foreground Cyan "	-User Account Disabled"
Write-Host -Foreground Cyan	""
Write-Host -Foreground Cyan	"	-account moved to Disabled Users OU"	
Write-Host -Foreground Cyan	""
Write-Host -Foreground Cyan	"	-removed from all Distribution, Security, and Google Groups"
Write-Host -Foreground Cyan	""
Write-Host -Foreground Cyan "	-email address hidden from Address Book"
Write-Host -Foreground Cyan	""
Write-Host -Foreground Cyan	"	-Google Apps Devices Deleted, and App Specific Passwords Removed"
Write-Host -Foreground Cyan ""
Write-Host -Foreground Cyan	"	-Avaya Extension deleted"	
Write-Host -Foreground Cyan ""
Write-Host -Foreground Cyan	"	-voicemail box deleted"	
Write-Host -Foreground Cyan ""
Write-Host -Foreground Cyan	"	-disabled in Staff Management"	
Write-Host -Foreground Cyan ""
Write-Host -Foreground Cyan	"	-Blackshield Token revoked"
Write-Host -Foreground Cyan ""
Write-Host -Foreground Cyan	"	-Google password changed for IT"
Write-Host -Foreground Cyan ""
Write-Host -Foreground Cyan "   -Arkadin accounts updated accordingly"
Write-Host -Foreground Cyan ""
Write-Host -Foreground Cyan ""
Write-Host -Foreground Gray "--------------------------------------------------------------------------------"
Write-Host -Foreground Red ""
Write-Host -Foreground Green	"ITEMS THAT NEED TO BE COMPLETED"				
Write-Host -Foreground Red ""
Write-Host -Foreground Red "	-Collect equipment and reallocated in Service Now"	
Write-Host -Foreground Red ""
Write-Host -Foreground Red "    -Retire employee accessories plan"
Write-Host -Foreground Red ""
Write-Host -Foreground Red "    -Take Backupify Exports and save them to \\chijob01\g$\Terms"
Write-Host -Foreground Red ""

Get-QADUser $name | out-default | fl name, phone, email, DN

Read-Host 'Press Enter to exit!' | Out-Null
