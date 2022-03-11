

#############
#
# Modules and File Things
#
#############

Write-Output "Program starting File and Function things"

$list = Import-Csv ''#UserPrincipalName must be used for look in what ever list or all user call you do

$TokenRefreshTime = Get-Date

# app permisisons or admin account  -https://docs.microsoft.com/en-us/graph/api/signin-list?view=graph-rest-1.0&tabs=http
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

   try
    {
        $Response = Invoke-RestMethod @RequestBody -ErrorAction Stop
    }
    catch [System.Net.WebException]
    {
        Write-Verbose "Exception being Handled"
        $statusCode = [int]$_.Exception.Response.StatusCode
        Write-Verbose $statusCode
        Write-Verbose $_.Exception.Message


        if($statusCode -eq 401)
        {
            # Token might have expired! Renew token and try again
            #$authResult = $authContext.AcquireToken($MSGraphURI, $clientId, $redirectUri, "Auto")
            #$token = $authResult.AccessToken
            #$headers = Get-Headers($token)
            #$oneSuccessfulFetch = $False
            Write-Verbose "Exception being Handled - Token being refreshed"

            $MSALtoken = RefreshToken

    

        }
        elseif($statusCode -eq 429 -or $statusCode -eq 504 -or $statusCode -eq 503)
        {
            Write-Verbose "Exception being Handled - Throttled sleep for a bit"

            # throttled request or a temporary issue, wait for a few seconds and retry
            Start-Sleep -Seconds 120
           
            while($true)
            {
                try
                {
                    $Response = Invoke-RestMethod @RequestBody -ErrorAction Stop

                    #we made it to this line in the code no more 429 throttle break from looping
                    break;
                }
                catch [System.Net.WebException]
                {
                    if($statusCode -eq 429)
                    {
                    Write-Verbose "Exception loop starting to while true sleep"
                    $statusCode = [int]$_.Exception.Response.StatusCode
                    Write-Verbose $statusCode
                    Write-Verbose $_.Exception.Message
                    Start-Sleep -Seconds 180 #3mins
                    }
                }
            }
            
        }
        elseif($statusCode -eq 403 -or  $statusCode -eq 401)
        {
            Write-Verbose "Exception being Handled - This Blew up sorry bad request"
            
            Write-Output "Please check the permissions of the user"
            Start-Sleep -Seconds 5
            $Response = Invoke-RestMethod @RequestBody #-ErrorAction Continue

            #break;
        }
		elseif($statusCode -eq 400)
		{
			Write-Verbose "Exception being Handled - This Blew up sorry bad request"
            
            Write-Output "400 log and skip to next user"
            Write-Output "Bad request for USER: $UPN $count"
			$Response = $null

		}
    }

    

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
		#Start-Sleep -Seconds 5

    }

	$count++

}






