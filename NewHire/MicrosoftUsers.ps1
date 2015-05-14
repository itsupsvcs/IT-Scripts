$xlCSV=6
$CSVfilename1 = "$home\MicrosoftUser.csv"


$Excel = New-Object -comobject Excel.Application  
$Excel.Visible = $False 
$Excel.displayalerts=$False
$Excelfilename = "$home\micusers.xlsx"


$Workbook = $Excel.Workbooks.Open($Excelfilename)
$Worksheet = $Workbook.Worksheets.item(3)
$Worksheet.SaveAs($CSVfilename1,$xlCSV)
$Excel.Quit()  

If(ps excel){kill -name excel}
(Get-Content $home\MicrosoftUser.csv) | % {$_ -replace '"', ""} | out-file -FilePath $home\MicrosoftUser.csv -Force -Encoding ascii