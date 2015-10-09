import-module ActiveDirectory
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

#[System.Diagnostics.Process]::Start("C:\Program Files (x86)\Avaya\Site Administration\bin\ASA.exe") | out-null

Write-Host -ForegroundColor Green " _______                    ___ ___ .__                
 \      \   ______  _  __  /   |   \|__|______   ____  
 /   |   \_/ __ \ \/ \/ / /    ~    \  \_  __ \_/ __ \ 
/    |    \  ___/\     /  \    Y    /  ||  | \/\  ___/ 
\____|__  /\___  >\/\_/    \___|_  /|__||__|    \___  >
        \/     \/                \/                 \/ 
                                                       "                                                                                                          
Write-Host ""
Write-Host ""
Write-Host ""	
Write-Host -Foreground Gray "Use your Avaya Site Administration Console to create the user's extension.  Once finished, close that program, and press any key in this window to continue."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp")

#=============================================================================
# Reads the user info from sheet3
#=============================================================================

$Excel = New-Object -comobject Excel.Application  
$Excel.Visible = $False 
$Excel.displayalerts=$False
$Excelfilename = $env:USERPROFILE + "\newhire\newhire.xlsx"

$Workbook = $Excel.Workbooks.Open($Excelfilename)
$Worksheet = $Workbook.Worksheets.item(3)
$userSheet = $Worksheet.Cells.Item(2, 1).text
$userSheet = $userSheet.split(',')
$Excel.Quit()  

If(ps excel){kill -name excel}

#=============================================================================
# Opening user detail list"
#=============================================================================

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

	$FirstName = $userSheet[0]
	$LastName = $userSheet[1]
    $UserName = $userSheet[2]
	$PersonalMail = $userSheet[3]
    $strOffice = $userSheet[4];
	$strTitle = $userSheet[5]; 
    $strDepartment = $userSheet[6]
    $Manager = $userSheet[7]
    $Distros = $userSheet[8]
	$pagerDutyUser = $userSheet[10]
	
	$samAccountName = $FirstName + " " + $LastName
	$PlainPassword = "default12"
	$DomainUserName = $domain + $username
	$Hdrive = "\\chifs01\home$\" + $Username
	$org = "Two-Factor Disabled Special cases"
	
	If ($strDepartment -eq "Tradeco")
	{$cfgDomain = $cfgTCODomain}

	
	#Exchange Specific
	$strMailAddress = $samAccountName -replace " ", ".";
	$strMailAddress = $strMailAddress + $cfgMailDomain;
	$strMailAlias = $samAccountName -replace " ", ".";
    $userPrincipalName = $strMailAddress

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
  "AreaCode" = "551134437212";
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
	elseif ($strDepartment -eq "Engineering")
	{$userOU = "OU=Development,"
    Write-Host "YES ENGINEERING"}
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



    
	
	# Disconnect QADService.
	disconnect-qadservice




 
#This is where the script will check if the user needs PagerDuty access and will grant it if needed

if ($pagerDutyUser -eq "TRUE")
{
$PDbody1 = @"
	{ "name"
"@
$PDbody2 = @"
	"$samAccountName", "email"
"@
$PDbody3 = @"
	"$strMailAddress" }
"@
	$PDbodyFull = "$PDbody1 : $PDbody2 : $PDbody3"
	
	$pagerDutyAdd = Invoke-RestMethod -Headers @{"Authorization"="Token token=v9E7rzAsDKxc1huTsjAz"} -Uri https://trading-technologies.pagerduty.com/api/v1/users -Method POST -ContentType "application/json" -Body $PDbodyFull
	
	if ($pagerDutyAdd) {Write-Host -ForegroundColor Green "PagerDuty account created for $samAccountName."}
}

