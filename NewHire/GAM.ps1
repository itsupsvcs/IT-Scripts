$UserInformation = Import-Csv user.csv
Foreach ($User in $UserInformation)
{
	$FirstName = $User.FirstName
	$LastName = $User.LastName
	$UserName = $User.Username
	$Password = "default12"
	$Manager = $User.Manager	
	$DomainUserName = $domain + $username
	$Distros = $User.Distros
	$Hdrive = "\\chifs01\home$\" + $Username
	$strDepartment = $User.Department
	$org = "Two-Factor Disabled Special cases"
    $UserEmail = $FirstName + "." + $LastName + "@tradingtechnologies.com"
	$strOffice = $User.PhysicalDeliveryLocation
}

If ($strOffice -eq "Chicago")
	{
	$Loc = "(TT)"
	}
elseif ($strOffice -eq "New York")
	{
	$Loc = "(TT-NY)"
	}
elseif ($strOffice -eq "Houston")
	{
	$Loc = "(TT-TX)"
	}
elseif ($strOffice -eq "London")
	{
	$Loc = "(TT-UK)"
	}
elseif ($strOffice -eq "Geneva")
	{
	$Loc = "(TT-GEN)"
	}
elseif ($strOffice -eq "Frankfurt")
	{
	$Loc = "(TT-FF)"
	}
elseif ($strOffice -eq "Tokyo")
	{
	$Loc = "(TT-TO)"
	}
elseif ($strOffice -eq "Singapore")
	{
	$Loc = "(TT-SIN)"
	}
elseif ($strOffice -eq "Sydney")
	{
	$Loc = "(TT-AU)"
	}
elseif ($strOffice -eq "Hong Kong")
	{
	$Loc = "(TT-HK)"
	}
elseif ($strOffice -eq "Sao Paulo")
	{
	$Loc = "(TT-SAO)"
	}
	

$LastNameLoc = $LastName +" "+ $Loc	

	
c:\gam\gam.exe create user $UserEmail firstname $FirstName lastname $LastNameLoc password $Password changepassword on org $org 
