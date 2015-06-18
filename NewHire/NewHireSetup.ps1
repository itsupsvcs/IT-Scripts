#
# Create user accounts in AD and Google
# This script will read a CSV file and create user accounts based on that information.
#
# Requirements:
#  [+] http://www.quest.com/powershell/activeroles-server.aspx (For setting
#      attributes on an AD DS account)


#=============================================================================
#Begin Telnet session to create users extension
#=============================================================================

[System.Diagnostics.Process]::Start("C:\Program Files (x86)\Avaya\Site Administration\bin\ASA.exe") | out-null

Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""	
Write-Host -Foreground Gray "Use your Avaya Site Administration Console to create the user's extension.  Once finished, close that program, and press any key in this window to continue."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp")

#=============================================================================
# Convert the New Hire Spreadsheet to a CSV
#=============================================================================
$xlCSV=6
$CSVfilename1 = "$home\NewHire\user.csv"



$Excel = New-Object -comobject Excel.Application  
$Excel.Visible = $False 
$Excel.displayalerts=$False
$Excelfilename = "$home\NewHire\newhire.xlsx"


$Workbook = $Excel.Workbooks.Open($Excelfilename)
$Worksheet = $Workbook.Worksheets.item(3)
$Worksheet.SaveAs($CSVfilename1,$xlCSV)
$Excel.Quit()  

If(ps excel){kill -name excel}
(Get-Content $home\NewHire\user.csv) | % {$_ -replace '"', ""} | out-file -FilePath $home\NewHire\user.csv -Force -Encoding ascii

#=============================================================================
# Opening user detail list"
#=============================================================================
$UserInformation = Import-Csv user.csv
$cfgTab = [char]9
$cfgCompany = "Trading Technologies";
$cfgMailDomain = "@tradingtechnologies.com"; #E-Mail Domain

$cfgDomain = "int.tt.local"; #Domain
$cfgTCODomain = "tradeco.int.tt.local";
$Domain = "intad\";
$cfgDC = "chiintdc01.int.tt.local";
$cfgTCODC = "Trcdc03.tradeco.int.tt.local";
$TCODomain = "TRADECO\";
$TradeCoOU = "OU=Traders,OU=Tradeco,OU=USERS,OU=CHI,OU=US,DC=tradeco,DC=int,DC=tt,DC=local";
#=============================================================================
# Do some logic about our environment.
#=============================================================================
Foreach ($User in $UserInformation)
{
	$FirstName = $User.FirstName
	$LastName = $User.LastName
	$PersonalMail = $User.PersonalEmail
	$UserName = $User.Username
	$samAccountName = $FirstName + " " + $LastName
	$PlainPassword = "default12"
	$Manager = $User.Manager	
	$DomainUserName = $domain + $username
	$Distros = $User.Distros
	$Hdrive = "\\chifs01\home$\" + $Username
	$strDepartment = $User.Department
	$org = "Two-Factor Disabled Special cases"
	
	If ($strDepartment -eq "Tradeco")
	{$cfgDomain = $cfgTCODomain}

	
	#Exchange Specific
	$strMailAddress = $samAccountName -replace " ", ".";
	$strMailAddress = $strMailAddress + $cfgMailDomain;
	$strMailAlias = $samAccountName -replace " ", ".";
    $userPrincipalName = $strMailAddress
}

Add-PSSnapin Quest.ActiveRoles.ADManagement
$Extension = Read-Host "Enter users Extension"
If ($Extension -eq "")
	{
	$Extension = "0000"
	}

	
If ($strDepartment -eq "Tradeco")
	{
	$cfgDC = $cfgTCODC;
	$Domain = $TCODomain;
	}
else 
	{
	$cfgDC = "chiintdc01.int.tt.local";
	$Domain = "intad\";
	}
	
	
$nl = [Environment]::NewLine

#=============================================================================
# A series of hash tables for office information.
#=============================================================================
$cfgChicago = @{
  "Address" = "222 S. Riverside Plaza"+ $nl +"Suite 1100";
  "City" = "Chicago";
  "State" = "Illinois";
  "PostalCode" = "60606";
  "Country" = "US";
  "AreaCode" = "1312";
  "OU" = "OU=USERS,OU=CHI,OU=US,DC=int,DC=tt,DC=local";
  "DispNamLoc" = "(TT)";
  "ExtRange" = 1011..1199 + 2500..2549;
  "ExtRange2" = 6001..6099 + 6350..6599;
  "ExtRange3" = 1600..1699 + 4618..4697;
  "StaffGroup" = "staff-chicago@tradingtechnologies.com";
  }
 

$cfgNewYork = @{
  "Address" = "One Liberty Plaza"+ $nl +"27th Floor";
  "City" = "New York";
  "State" = "New York";
  "PostalCode" = "10006";
  "Country" = "US";
  "AreaCode" = "1212313";
  "OU" = "OU=USERS,OU=NYC,OU=US,DC=int,DC=tt,DC=local";
  "DispNamLoc" = "(TT-NY)";
  "ExtRange" = 5300..5399;
  "StaffGroup" = "staff-newyork@tradingtechnologies.com";
  }
  
$cfgHouston = @{
  "Address" = "9 Greenway Plaza"+ $nl +"Suite 3045";
  "City" = "Houston"
  "State" = "Texas"
  "PostalCode" = "77046"
  "Country" = "US";
  "AreaCode" = "1713568";
  "OU" = "OU=USERS,OU=HOU,OU=US,DC=int,DC=tt,DC=local";
  "DispNamLoc" = "(TT-TX)";
  "ExtRange" = 2019..2049;
  "StaffGroup" = "staff-houston@tradingtechnologies.com";
  }
  
$cfgLondon = @{
  "Address" = "85 King William Street"+ $nl +"7th Floor";
  "City" = "London"
  "State" = " "
  "PostalCode" = "EC4N 7BL"
  "Country" = "GB";
  "AreaCode" = "44207621";
  "OU" = "OU=USERS,OU=LON,OU=Europe,DC=int,DC=tt,DC=local";
  "DispNamLoc" = "(TT-UK)";
  "ExtRange" = 8100..8199; 
  "StaffGroup" = "staff-london@tradingtechnologies.com";
  }
 
$cfgGeneva = @{
  "Address" = "4th Floor"+ $nl +"14 Rue du Rhone";
  "City" = "Geneva"
  "State" = " "
  "PostalCode" = "1204"
  "Country" = "CH";
  "AreaCode" = "4122319 ";
  "OU" = "OU=USERS,OU=GEN,OU=Europe,DC=int,DC=tt,DC=local";
  "DispNamLoc" = "(TT-GEN)";
  "ExtRange" = 3426..3429;  
  "StaffGroup" = "staff-geneva@tradingtechnologies.com";
  }

$cfgFrankfurt = @{
  "Address" = "Goethestrasse 27"+ $nl +"6th Floor";
  "City" = "Frankfurt"
  "State" = " "
  "PostalCode" = "60313"
  "Country" = "DE";
  "AreaCode" = "49692972";
  "OU" = "OU=USERS,OU=FRA,OU=Europe,DC=int,DC=tt,DC=local";
  "DispNamLoc" = "(TT-FF)";
  "ExtRange" = 8200..8249;
  "StaffGroup" = "staff-frankfurt@tradingtechnologies.com";
  }
  
$cfgTokyo = @{
  "Address" = "Nihonbashi Muromachi Plaza Building 5F"+ $nl +"3-4-7 Nihonbashi Muromachi, Chuo-ku";
  "City" = "Tokyo"
  "State" = " "
  "PostalCode" = "103"+"-"+"0022"
  "Country" = "JP";
  "AreaCode" = "8134577";
  "OU" = "OU=USERS,OU=TO,OU=Asia-Pacific,DC=int,DC=tt,DC=local";
  "DispNamLoc" = "(TT-TO)";
  "ExtRange" = 8300..8349;
  "StaffGroup" = "staff-tokyo@tradingtechnologies.com";
  }

$cfgSingapore = @{
  "Address" = "3 Church Street #18-01/06"+ $nl +"Samsung Hub";
  "City" = "Singapore"
  "State" = " "
  "PostalCode" = "049483"
  "Country" = "SG";
  "AreaCode" = "656395";
  "OU" = "OU=USERS,OU=SIN,OU=Asia-Pacific,DC=int,DC=tt,DC=local";
  "DispNamLoc" = "(TT-SIN)";
  "ExtRange" = 7000..7099;
  "StaffGroup" = "staff-singapore@tradingtechnologies.com";
  }
  
$cfgSydney = @{
  "Address" = "50 Margaret Street Level 11"+ $nl +"Suite 1";
  "City" = "Sydney"
  "State" = " "
  "PostalCode" = "NSW 2000"
  "Country" = "AU";
  "AreaCode" = "6128022";
  "OU" = "OU=USERS,OU=SYD,OU=Asia-Pacific,DC=int,DC=tt,DC=local";
  "DispNamLoc" = "(TT-AU)";
  "ExtRange" = 1700..1799;
  "StaffGroup" = "staff-sydney@tradingtechnologies.com";
  }

$cfgHongKong = @{
  "Address" = "Room 2176, The Center"+ $nl +"99 Queens Road";
  "City" = "Hong Kong"
  "State" = " "
  "PostalCode" = " "
  "Country" = "HK";
  "AreaCode" = "8523478";
  "OU" = "OU=Sales,OU=USERS,OU=HK,OU=Asia-Pacific,DC=int,DC=tt,DC=local";
  "DispNamLoc" = "(TT-HK)";
  "ExtRange" = 7000..7099;
  "StaffGroup" = "staff-hongkong@tradingtechnologies.com";
  }

$cfgSaoPaulo = @{
  "Address" = "Av. Brigadeiro Faria Lima, 3729"+ $nl +"5th Floor, Suite 549/551";
  "City" = "Sao Paolo"
  "State" = " "
  "PostalCode" = "04538"+"-"+"905"
  "Country" = "BR";
  "AreaCode" = "+551134437212";
  "OU" = "OU=USERS,OU=SAO,OU=South America,DC=int,DC=tt,DC=local";
  "DispNamLoc" = "(TT-SAO)";
  "ExtRange" = 1000..1199;
  "StaffGroup" = "staff-saopaulo@tradingtechnologies.com";
  }
  
 
#=============================================================================
# Creates an array of the above hash tables.
#=============================================================================
$cfgOffices = @{
  "Chicago" = $cfgChicago;
  "New York" = $cfgNewYork;
  "Houston" = $cfgHouston;
  "London" = $cfgLondon;
  "Geneva" = $cfgGeneva;
  "Frankfurt" = $cfgFrankfurt;
  "Tokyo" = $cfgTokyo;
  "Hong Kong" = $cfgHongKong;
  "Singapore" = $cfgSingapore;
  "Sydney" = $cfgSydney;
  "Sao Paulo" = $cfgSaoPaulo;
  
  };


	# Attributes.
	$strOffice = $User.PhysicalDeliveryLocation;
	$strTitle = $user.JobTitle; 
	$strAddress = $cfgOffices.Get_Item( $strOffice ).Get_Item("Address");
	$strCity = $cfgOffices.Get_Item( $strOffice ).Get_Item("City");
	$strState = $cfgOffices.Get_Item( $strOffice ).Get_Item("State");
	$strPostalCode = $cfgOffices.Get_Item( $strOffice ).Get_Item("PostalCode");
	$strCountry = $cfgOffices.Get_Item( $strOffice ).Get_Item("Country");
	$strAreaCode = $cfgOffices.Get_Item( $strOffice ).Get_Item("AreaCode");
	$DisplayNameLocation = $cfgOffices.Get_Item( $strOffice ).Get_Item("DispNamLoc")
	$DisplayName = $samAccountName + " " +  $DisplayNameLocation
	$strActualOU = "OU=" + $strDepartment + "," + $strOU
	$StaffGroup = $cfgOffices.Get_Item( $strOffice ).Get_Item("StaffGroup")
	
	$strExtRange = $cfgOffices.Get_Item( $strOffice ).Get_Item("ExtRange")
	$strExtRange2 = $cfgOffices.Get_Item( $strOffice ).Get_Item("ExtRange2")
	$strExtRange3 = $cfgOffices.Get_Item( $strOffice ).Get_Item("ExtRange3")
	write-host""
	write-host""
	write-host""
	write-host""
	#Phone number and extension
	#DID ranges for each location
	$strTel = 0 ; 
	
	If ($strOffice -eq "Chicago")
		{
			If ($strExtRange -contains $Extension)
			{ 
				$strTel = "+" + "$strAreaCode" + "476" + "$Extension"
			}
			elseif ($strExtRange2 -contains $Extension)
			{
				$strTel = "+" + "$strAreaCode" + "698" + "$Extension"
			}
			elseif ($strExtRange3 -contains $Extension)
			{ 
				$strTel = "+" + "$strAreaCode" + "268" + "$Extension"
			}
			else
			{
			$strTel = "+" + "$strAreaCode" + "4761000"
			Write-Host -Foreground Red "-------------------------------------------------------------------------------"
			Write-Host -Foreground Yellow "Warning: The extension chosen is not a valid extension for the given location! "
			Write-Host -Foreground Red "-------------------------------------------------------------------------------"
			Write-Host ""
			}
		}
	elseif ($strOffice -eq "Sao Paulo")
		{
			$strTel = "+" + "$strAreaCode"
		}
	elseif ($strExtRange -contains $Extension)
		{
			$strTel = "+" + "$strAreaCode" + "$Extension"
		}
	else 
		{
			$strTel = "+" + "$strAreaCode" + "$Extension"
			Write-Host -Foreground Red "--------------------------------------------------------------------------------"
			Write-Host -Foreground Yellow "Warning: The extension chosen is not a valid extension for the given location! "
			Write-Host -Foreground Red "--------------------------------------------------------------------------------"
			Write-Host ""
		}
#=======================================================================		
#Specify the correct OU		
#=======================================================================
    IF ($strOffice -eq "Hong Kong")
	{$userOU = ""}
	elseif ($strDepartment -eq "Accounting")
	{$userOU = "OU=Accounting,"}
	elseif ($strDepartment -eq "Administration")
	{$userOU = "OU=Administration,"}
	elseif ($strDepartment -eq "Development")
	{$userOU = "OU=Development,"}
	elseif ($strDepartment -eq "Executives")
	{$userOU = "OU=Executives,"}
	elseif ($strDepartment -eq "Human Resources")
	{$userOU = "OU=Human Resources,"}
	elseif ($strDepartment -eq "Intellectual Property")
	{$userOU = "OU=Intellectual Property,"}
	elseif ($strDepartment -eq "Intern-Temp")
	{$userOU = "OU=Contractors Interns and Temps,"}
	elseif ($strDepartment -eq "IT-Enterprise Apps")
	{$userOU = "OU=Enterprise Applications,OU=IT,"}
	elseif ($strDepartment -eq "IT-Support Services")
	{$userOU = "OU=Support Services,OU=IT,"}
	elseif ($strDepartment -eq "IT-ELS")
	{$userOU = "OU=Trading Systems,OU=IT,"}
	elseif ($strDepartment -eq "IT-Ent Dev")
	{$userOU = "OU=Internal Development,OU=IT,"}
	elseif ($strDepartment -eq "Debesys")
	{$userOU = "OU=PMM,"}
	elseif ($strDepartment -eq "PMM-Documentation")
	{$userOU = "OU=Documentation,OU=PMM,"}
	elseif ($strDepartment -eq "PMM-Marketing")
	{$userOU = "OU=Marketing,OU=PMM,"}
	elseif ($strDepartment -eq "Recruiting")
	{$userOU = "OU=Recruiting,"}
	elseif ($strDepartment -eq "Sales")
	{$userOU = "OU=Sales,"}
	elseif ($strDepartment -eq "Support")
	{$userOU = "OU=Support,"}
	else
	{$userOU = ""}
	
$strOU = $userOU + $cfgOffices.Get_Item( $strOffice ).Get_Item("OU");

	If($strDepartment -eq "Tradeco")
	{$strOU = $TradeCoOU}

#==============================================================================================================
#Preview New Hire's Information
#==============================================================================================================		
Write-Host -Foreground Gray "--------------------------------------------------------------------------------"
Write-Host -Foreground Cyan "New Hire Information"
Write-Host -Foreground Gray "--------------------------------------------------------------------------------"
	Write-Host " Employee Name:"$DisplayName;
	Write-Host " Username:"$cfgTab$UserName;
    Write-Host " Manager:"$cfgTab$Manager;
	Write-Host " Office:"$cfgTab$strOffice;
	Write-Host " Job Title:"$cfgTab$strTitle;
	Write-Host " OU:"$cfgTab$cfgTab$strOU;
	Write-Host " E-Mail:"$cfgTab$strMailAddress;
	Write-Host " Phone:"$cfgTab$strTel;
	Write-Host " Mentor:"$cfgTab$Distros;
	
Write-Host -Foreground Gray "-------------------------------------------------------------------------------------"
Write-Host ""
	Write-Host -Foreground Cyan "Please review the above information for" $SamAccountName "before we add them to our system"
Write-Host ""
	Write-Host -Foreground Gray "-------------------------------------------------------------------------------------"
	start-sleep -s 2
Write-Host ""
Write-Host ""
Write-Host -Foreground Gray "If all of the above information is correct, press any key to continue. If you do not wish to continue, please press Cntrl-C:"
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp")

	# $Password = ConvertTo-SecureString $PlainPassword -AsPlainText -Force

	
	# Set Quest Active Directory stuff to use a DC in the local site.
	Connect-QADService -service $cfgDC | out-null

	# Create Active Directory user Account
	new-QADUser -ParentContainer $strOU -Name $SamAccountName -FirstName $FirstName -LastName $LastName -UserPrincipalName $userPrincipalName -DisplayName $DisplayName -SamAccountName $UserName -Email $strMailAddress -UserPassword $PlainPassword | out-null
	disconnect-qadservice
	Connect-QADService -service $cfgDC | out-null
	
	
	# Set attributes on AD DS account.
	Get-QADUser $Username | out-null
	start-sleep -s 2
	Set-QADUser $Username -office $strOffice -department $strDepartment -StreetAddress $StrAddress -city $strCity -StateOrProvince $strState -PostalCode $strPostalCode -description $strTitle -Company $cfgCompany -title $strTitle -PhoneNumber $strTel | out-null
	Set-QADUser $Username -objectattributes @{ipPhone=$Extension} | out-null
	Set-QADUser $Username -objectattributes @{c=$strCountry} | out-null
    Write-host ""

	If($Distros -eq "") {
        Write-host "The manager did not provide a user to mirror groups off of, you will have to add them to the appropriate security groups manually."
    }
	Else {
	    $K = Get-QADUser $Distros | select memberof 
	        foreach($user in $K.memberof) {
		        try{
                    # DO NOT COPY AWS PRIVS TO NEW USERS
                    if ($user -like "*AWS*")  {
                        Write-Host $user 
                        Write-Host -ForegroundColor Red "This will not be mirrored. Check in with managers for AWS access."
                        }
                    Else{
                        Add-QADGroupMember -Identity $user -Member $Username | out-null
                        }
                    }
		        catch{}	
        
	        }
        }
	Get-QADUser $UserName -includedproperties ipphone | out-default | fl displayname, title, department, manager, ipphone, email 
	
	# Disconnect QADService.
	disconnect-qadservice
	
import-module ActiveDirectory

$NIS = Get-ADObject "CN=int,CN=ypservers,CN=ypServ30,CN=RpcServices,CN=System,DC=int,DC=tt,DC=local" -Properties:*
$maxUid = $NIS.msSFU30MaxUidNumber + 1
Set-ADObject $NIS -Replace @{msSFU30MaxUidNumber = "$($maxUid)"}

 Set-ADUser -Identity "$($Username)" -Replace @{mssfu30nisdomain = "int"} #Enable NIS
   Set-ADUser -Identity "$($Username)" -Replace @{gidnumber="10000"} #Set Group ID
   $maxUid++ #Raise the User ID number
   Set-ADUser -Identity "$($Username)" -Replace @{uidnumber=$maxUid} #Set User ID
   Set-ADUser -Identity "$($Username)" -Replace @{loginshell="/bin/bash"} #Set user login shell
   Set-ADUser -Identity "$($Username)" -Replace @{msSFU30Name="$($Username)"}
   Set-ADUser -Identity "$($Username)" -Replace @{unixHomeDirectory="/home/$($Username)"}
   Write-Host -Backgroundcolor Green -Foregroundcolor Black $usr.SamAccountName changed #Write Changed Username to console	

#This is where we will Launch the GAM tool to add the user to Google

Write-Host -Foreground Green "The user's Google profile will now be created."
Write-Host ""
start-sleep -s 3
Write-Host -Foreground Green "Googlfying "$DisplayName"..."
Write-Host ""
.\Gam.ps1 
Write-Host ""
Write-Host -Foreground Green "Adding "$DisplayName" to the Staff-"$StrOffice" Google Group..."
c:\gam\gam.exe update group $StaffGroup add member $strMailAddress
Write-Host "" 
Write-Host -Foreground Green $DisplayName" can now use The Google."
Write-Host ""
c:\gam\gam.exe user $strMailAddress delegate to it-support@tradingtechnologies.com
Write-Host -Foregroung Green " "$Displayname"'s mailbox is now a delegate of IT-Support."



#$ie = New-Object -ComObject InternetExplorer.Application
#$ie.Navigate("http://intranet/staffmanagement/") | out-null
#try{$ie.Visible = $true}
#catch{}
#Write-Host ""
#Write-Host ""
#Write-Host ""
#Write-Host ""	
#Write-Host -Foreground Gray "Use the IE window that was launched to add the user to Staff Management.  Once finished, close that window, and press any key in this window to continue."
#$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp")

$ie = New-Object -ComObject InternetExplorer.Application
$ie.Navigate("https://172.17.24.35/") 
try{$ie.Visible = $true}
catch{}
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""	
Write-Host -Foreground Gray "Use the IE window that was launched to create the users voicemail.  Their Extension is " $Extension ".  Once finished, close that window, and press any key in this window to continue."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp")

$NewEmail = Read-Host "If you would like to send the new Hire email, type YES and press enter (if not, just press enter to skip this)."

Write-Host ""

if($NewEmail -eq 'YES')
	{     
		 Write-Host ""
		 Write-Host "Sending Welcome Email to the new hire's personal email address."
		 Write-Host ""
		
		 

			$smtp = "mail.int.tt.local"

			$to = $FirstName +" "+$LastName+ " <"+$PersonalMail+">"

			$from = "IT-Support <it-support@tradingtechnologies.com>"

			$subject = "Welcome to Trading Technologies!" 

			$body = "<img src='http://www.marketswiki.com/wiki/images/e/ee/TT_horizontal_2lines_4c_logo225.jpg'> <br>"


			$body += "Dear $to,<br>"

			$body += "Welcome to Trading Technologies! The first step in your new hire process at TT is logging into your TT Gmail account for the first time.  To do this, go to <a href=http://mail.google.com>mail.google.com</a>. <br>"

			$body += "<br>
					Your email address is: <b>$strMailAddress</b> <br> 
					Initial Password: <b>default12</b> (you will be prompted to create a new password) <br>"

			$body += "<br>
						We are sure you have many questions about your first day here so we have put together a <a href=https://sites.google.com/a/tradingtechnologies.com/new-hire-information>New Hire Information Site</a> in hopes of arming you with some information prior to your arrival.  You will be able to access this site once you login to your TT Gmail account. <br>"

			$body += "<br>
					   <b><font color=blue>Thank you,</b></font> <br>
					   IT Service Desk <br>
					   X1911 <br>
					   (312) 268-1607 <br>"
					   
			#### Now send the email using \> Send-MailMessage 

			send-MailMessage -SmtpServer $smtp -To $to -From $from -Subject $subject -Body $body -BodyAsHtml -Priority high
}
else
{
Write-Host ""
Write-Host "No email will be sent to the new hire."
Write-Host ""
}

