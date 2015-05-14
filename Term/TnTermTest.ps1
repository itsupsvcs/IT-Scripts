Add-PSSnapin Quest.Activeroles.admanagement
Connect-QADservice -service rostndc01.tn.local
write-host "Disabling the GAtest user account in the TN network"
try{Disable-QADUser gatest | out-null}
catch{}
Disconnect-QADService
