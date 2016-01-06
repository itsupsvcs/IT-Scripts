function sendMail($recipient, $subject, $body)
{ 
     #SMTP server name
     $smtpServer = "mail.int.tt.local"

     #Creating a Mail object
     $msg = new-object Net.Mail.MailMessage

     #Creating SMTP server object
     $smtp = new-object Net.Mail.SmtpClient($smtpServer)

     #Email structure 
     $msg.From = "noreply@tradingtechnologies.com"
     $msg.To.add($recipient)
     $msg.subject = $subject
     $msg.body = $body

     #Sending email 
     $smtp.Send($msg)

     Write-Host "Recipient value = $recipient"
     Write-Host "Subject value = $subject"
     Write-Host "Body value = $body"
}	

#If you want to test send email replace the recipient below and uncomment the line
#sendMail -recipient "Lauren.Lavin@tradingtechnologies.com" -subject "Disable MSDN Account For $DispName" -body "Please check MSDN to see if $DispName  has an account, and if so disable it.  This person is no longer employed at Trading Technologies."