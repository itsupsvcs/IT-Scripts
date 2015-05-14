
$termUserEmail = "Lauren.Lavin@tradingtechnologies.com"
$joinmeUser = "joinmeadmin@tradingtechnologies.com"
$joinmePass = "VZQk8m7xah"

$ie = new-object -com "InternetExplorer.Application"


$ie.visible = $true
$ie.height = 1000
$ie.navigate("https://accounts.logme.in/login.aspx?clusterid=02&returnurl=https%3a%2f%2fwww.appguru.com%2ffederated%2floginsso&headerframe=https%3a%2f%2fwww.appguru.com%2fwelcome%2fcommon%2fproduct%2fpages%2fcls%2fheaderframe%2f&productframe=https%3a%2f%2fwww.appguru.com%2fwelcome%2fcommon%2fproduct%2fpages%2fcls%2fdefaultframe%2f&lang=en-US&skin=appguru&regtype=O&trackingproducttype=1&socialloginenabled=0")

while($ie.busy){Start-Sleep 1}

$ie.Document.getElementById("email").value=$joinmeUser
$ie.Document.getElementById("password").value=$joinMePass
$ie.Document.getElementById("btnSubmit").Click()
while($ie.busy){Start-Sleep 1}

$ie.navigate("https://www.appguru.com/policymanager")
while($ie.busy){Start-Sleep 1}























