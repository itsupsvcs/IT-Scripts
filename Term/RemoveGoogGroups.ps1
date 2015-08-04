$Gmail = $args


Function removeGoogleGroups($Gmail)
{
    Write-Host "Removing $Gmail from all Google Groups" -ForegroundColor Green

    #GAM call to get User information, create substring by finding lines after 'Groups:'
    $purge_usr = $Gmail
    $purge_chunk= c:\gam\gam.exe info user $purge_usr

    #If you want to see this information do a Write-Host $purge_chunk
    $purge_chunk= $purge_chunk | Select-String "Groups:" -context 0,100

    #Remove 'Groups: ' from string
    $purge_chunk=$purge_chunk.ToString().Substring(9)

    #Find length of string to determine if this is groupless
    $length = $purge_chunk.Length

    if($length.Equals(0))
    {
        Write-Host "This user is not a member of any groups. Skipping process" -ForegroundColor Green
        return 
    }

    #Meat and Potatoes -- Separate emails from group names by replacing "<" ">" with 
    #return/new line statements

    #Example output from GAM
    #Groups:
    #2sv <2sv@acme.org>
    #users <users@acme.org>
   
    $array = $purge_chunk.replace(">","`r`n").replace("<","`r`n").ToString().Split("`r`n")

    #Splitting creates an array of string objects, parse through each email for removal

    foreach ($line in $array)
    {
        if ($line.contains("@")) 
        {  
            Write-Host "$line" -ForegroundColor Yellow
            c:\gam\gam.exe update group $line remove member $purge_usr
        }

    }
}


#If you want to run this for a single user just uncomment below
#removeGoogleGroups($Gmail)
