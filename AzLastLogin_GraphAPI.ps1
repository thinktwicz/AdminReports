

#############
#
# Modules and File Things
#
#############

Write-Output "Program starting File and Function things"

$list = Import-Csv ''#UserPrincipalName must be used for look in what ever list or all user call you do

$TokenRefreshTime = Get-Date

Import-module MSAL.PS
$clientid = ""
$tenantid = ""
$MSALtoken = Get-MsalToken  -ClientId $clientID -TenantId $tenantID -ClientSecret (ConvertTo-SecureString '' -AsPlainText -Force)

#$MSALtoken = Get-MsalToken -Interactive -ClientId $clientID -TenantId $tenantID



# helper function for a new token 
function RefreshToken()
{
    return $MSALtoken = Get-MsalToken -ForceRefresh -ClientId $clientID -TenantId $tenantID -ClientSecret (ConvertTo-SecureString '' -AsPlainText -Force)
}

#make a filepath
$targetFile = 'C:\helloLogon.csv'
#remove any version 
rm $targetFile -ErrorAction SilentlyContinue
Add-Content $targetFile "DisplayName,UPN,lastSignInRequestId,LastDateInteractive,lastNonInteractiveSignInRequestId,LastDateNonInteractive" 


Write-Output "Looping through the users"

$count = 0
# Loop through each user in the list 
foreach($item in $list)
{


    ###########
    #
    # Graph Call 
    #
    ###########

    #header token
    $headers  = @{Authorization = "Bearer $($MSALtoken.accesstoken)" }
    
    #user we are checking 
    $UPN = $item.userPrincipalName.tostring()
    #$UPN = ''
	Write-Output "User: " $UPN $count

    #api endpoint
    $apiUrl = "https://graph.microsoft.com/beta/users?`$filter=startswith(userprincipalname,'$UPN')&`$select=displayName,signInActivity"

    # body of request
    $RequestBody = @{

                Uri = $apiUrl
                headers  = $headers
                method = 'GET'
                Contenttype = "application/json" 
       
    }
    #//

    #return response
    $Response = Invoke-RestMethod @RequestBody

    

    foreach($item in $Response)
    {
        if($Response.value.Count -eq 0)       
        {
            Write-Output "UserNull: "$UPN
			#TODO:! Error report
           
        }
        else
        {
            Write-Output "User: "$UPN $Response.value

             
            # take the values from the response
            $displayname = $item.value.displayName
            $LogonEventID = $item.value.id
            
            # you get back displayname and ID by default sometimes there is activity returned
            if($item.value.signInActivity.count -ne 0)
            {
            
              $lastSignDate =  $item.value.signInActivity.lastSignInDateTime
              $lastNonSignDate =  $item.value.signInActivity.lastNonInteractiveSignInDateTime
              $lastNonSignID = $item.value.signInActivity.lastNonInteractiveSignInRequestId

            }
            # null checking and setting
            if($displayname -eq $null)
            {
                $displayname = "Null"
            }
            else
            {
            }

            if($LogonEventID -eq $null)
            {
                $LogonEventID = "Null"

            }
   
                 

            $line = "$displayname,$UPN,$LogonEventID,$lastSignDate,$lastNonSignID,$lastNonSignDate"
			      Write-Output $line
            Add-Content $targetFile $line
			      $displayname,$UPN,$LogonEventID,$lastSignDate,$lastNonSignID,$lastNonSignDate = $null
        
        }


        # every 30 mins refresh the token

        $RefreshDiff = New-TimeSpan -Start $TokenRefreshTime -End (Get-Date)
        if ($RefreshDiff.Minutes -ge 30)
        {
            $MSALtoken = RefreshToken
            $TokenRefreshTime = Get-Date
        
        }

        #reset vars used in file writes and lookups
        $displayname,$UPN,$LogonEventID,$lastSignDate,$lastNonSignID,$lastNonSignDate = $null

		# WE MUST SLEEP - shh dont wake the throttle baby - 5 secs unless this one can get a throttle pass \(^-^)/
		Start-Sleep -Seconds 5

    }

	$count++

}






