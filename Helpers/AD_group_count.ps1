import-module ActiveDirectory
<#This is an example script to find out how many users are in a specific group or set of groups#>

$userArray = @()
$names

#This looks through all groups that start with prod-aws 
$Groups = Get-ADGroup -Filter {Name -like "PROD-AWS*"} | Select-Object Name

#iterate through each group and grab all of the names from the AD objects
ForEach ($Group in $Groups)
{
Write-Host " "
Write-Host "$($group.name)" -ForegroundColor Red
Write-Host "-------"

$names += Get-ADGroupMember $Group.Name -Recursive | Select-Object Name

}

#split makes all the names enter an array, selet unique will make sure we don't have duplicates
$names = $names -split "`n"
$names = $names |Select -Unique
$names.count

#start all over again for debesys-ext

$Groups = Get-ADGroup -Filter {Name -like "debesys-ext*"} | Select-Object Name

ForEach ($Group in $Groups)
{
Write-Host " "
Write-Host "$($group.name)" -ForegroundColor Red
Write-Host "-------"

$names += Get-ADGroupMember $Group.Name -Recursive | Select-Object Name

}

$names = $names -split "`n"
$names = $names |Select -Unique
$names.count

$names

#this sloppy code below was used to get the groups member totals

<#
$totalCount = 0;
$groupCount = 0;
$counter = 0;

Write-Host "`n DEBESYS-EXT GROUPS `n "
$Groups = Get-ADGroup -Filter {Name -like "debesys-ext*"} | Select-Object Name


ForEach ($Group in $Groups)
{
Write-Host " "
Write-Host "$($group.name)" -ForegroundColor Red
Write-Host "-------"
(Get-ADGroupMember -identity $($group.name)).count
$counter = (Get-ADGroupMember -identity $($group.name)).count
$groupCount += $counter
$totalCount += $counter

}

Write-Host "COUNT: `n"$groupCount

$groupCount = 0;

Write-Host "`n PROD-AWS GROUPS `n "
$Groups = Get-ADGroup -Filter {Name -like "prod-aws*"} | Select-Object Name

ForEach ($Group in $Groups)
{
Write-Host " "
Write-Host "$($group.name)" -ForegroundColor Red
Write-Host "-------"
(Get-ADGroupMember -identity $($group.name)).count
$counter = (Get-ADGroupMember -identity $($group.name)).count
$groupCount += $counter
$totalCount += $counter

}

Write-Host "COUNT: `n"$groupCount

Write-Host "`n TOTAL COUNT `n" $totalCount



#>