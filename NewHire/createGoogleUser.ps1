#wrap GAM into fuction
#rename this to something like createGoogleUser 
#pass all arguments needed to create google user shown on the last line

Function createGoogleUser ($FirstName, $LastName, $strOffice, $UserEmail)
{
	$Password = "default12"
	$org = "Two-Factor Disabled Special cases"

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

}
