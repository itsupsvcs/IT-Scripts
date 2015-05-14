Write-Host ""
Write-Host ""
$Legal=Read-Host -Foreground Red "If this user was a member of the legal team, type YES and press enter (if not, just press enter to skip this).  This will launch an IE window.  Please disable their standard user account AND admin account in Anaqua (admin accounts follow the following naming convention: first.last+ADMIN@tradingtechnologies.com). If this user is not a member of legal this step can be skipped."
Write-Host ""
Write-Host ""

if($Legal -eq 'YES')
  {
  write-host "YES"
  }