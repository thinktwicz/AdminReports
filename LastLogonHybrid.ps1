<#
	Author: Drew Crouch - still underwork never did keyvalt things
	Purpose: Use a list to run against the graph api for last logon information in azure and onprem ad to get attributes and report them

#>

#start time for script
$ScriptStartTime = Get-Date
#############
#
# Modules Import MSAL.PS
#
#############

Write-Output "import module"
Import-Module MSAL.PS

#############
#
# Error Preferences 
#
#############
# Set Preferences
#$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"

#############
#
# Modules and File Things
#
#############

##
# Azure App IDs 
$clientid = ""
$tenantid = ""


Write-Output "Program starting File and Function things"

$Cred = Get-AutomationPSCredential -Name ''
##
#
# - making the all users list
#
##
#Write-Verbose "Generating All Users List"

Write-Output "connecting to the graph and azure ad"

# Populate with the App Registration details and Tenant ID
$clientid = ""
$tenantid = ""
$secret = ''
 
$body =  @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $clientid
    Client_Secret = $secret
}
 
$connection = Invoke-RestMethod -Uri https://login.microsoftonline.com/$tenantid/oauth2/v2.0/token -Method POST -Body $body
 
$token = $connection.access_token
 
Connect-MgGraph -AccessToken $token

Connect-AzureAD -Credential $Cred

Write-Output "graph call all mg-graph"


#Get-MgUser -All | select Id,DisplayName,Mail,UserPrincipalName | export-csv "C:\{}.csv"


$list = import-csv "C:\{}.csv"
$targetFile = "C:\{}.csv"

$File = @()

Write-Output "Making Azure AD MG-Graph All user List"
Write-Output "Skipping"
<#
foreach($item in $list)
{
    
    $userinfo = Get-AzureADUser -ObjectId $item.Id | select DirSyncEnabled,UserType
   
    $FileData = New-Object psobject
    # Filling in our data and adding into the file array 
    $FileData | add-member -MemberType NoteProperty -Name DisplayName -Value $item.DisplayName
    $FileData | add-member -MemberType NoteProperty -Name Mail -Value $item.Mail
    $FileData | add-member -MemberType NoteProperty -Name Id -Value $item.Id
    $FileData | add-member -MemberType NoteProperty -Name UPN -Value $item.UserPrincipalName
    $FileData | add-member -MemberType NoteProperty -Name DirSyncEnabled -Value $userinfo.DirSyncEnabled
    $FileData | add-member -MemberType NoteProperty -Name UserType -Value $userinfo.UserType
    $File += $FileData    
}

#Export the file 
$File | export-csv $targetFile  -NoTypeInformation
#>


Write-Output "Loading up list"
$list = Import-Csv 'C:\{}.csv'

##
# Start our timmer for when we refresh our token
$TokenRefreshTime = Get-Date

Write-Output "Done with the all users list "

# Application Connecting passing Secure String Azure App Secret 
$MSALtoken = Get-MsalToken  -ClientId $clientID -TenantId $tenantID -ClientSecret (ConvertTo-SecureString '$secret' -AsPlainText -Force)
# Delegate connection
#$MSALtoken = Get-MsalToken -Interactive -ClientId $clientID -TenantId $tenantID

# helper function for a new token 
function RefreshToken()
{
    return $MSALtoken = Get-MsalToken -ForceRefresh -ClientId $clientID -TenantId $tenantID -ClientSecret (ConvertTo-SecureString '$secret' -AsPlainText -Force)
}

Write-Output "we have tokens and will setup files next"
##
#
# - Where do we write out errors and the main file
#
##
$targetFile = 'C:\{}csv'
$ErrortargetFile = 'C:\{}.csv'
$ErrorADLookupFile = 'C:\{}.csv'

##
#
# - File array objects that hold the rows
#
##
$File = @()
$ErrorFile = @()
$ErrorADFile = @()

#############
#
# Main
#
#############

# Count will count each pass this should match the count number of the list we run with
$count = 0
Write-Output "Starting Loop on All users to LastLogin by ID"
##
#
# - Loop through the list of user UPNs to check on
#
##
Write-Verbose "looping"
foreach($item in $list)
{

	# Displayname of the user we are checking
    $displayname = $item.DisplayName

    ###########
    #
    # Graph Call 
    #
    ###########

    #header token
    $headers  = @{Authorization = "Bearer $($MSALtoken.accesstoken)" }
    
    #user we are checking 
    $UPN = $item.UPN.tostring()
    $ID = $item.Id
    #api endpoint
    $apiUrl = "https://graph.microsoft.com/beta/users?`$filter=(id eq '$ID')&`$select=displayName,signInActivity"
    Write-Verbose $apiUrl
    # body of request
    $RequestBody = @{

                Uri = $apiUrl
                headers  = $headers
                method = 'GET'
                Contenttype = "application/json" 
       
    }
    #//

   
	# First Try of Invoke
    try
    {
        $Response = Invoke-RestMethod @RequestBody -ErrorAction Stop
		Write-Verbose $Response
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
            Write-Verbose "Exception being Handled - Token being refreshed"
            $MSALtoken = RefreshToken

        }
        elseif($statusCode -eq 429 -or $statusCode -eq 504 -or $statusCode -eq 503)
        {
            Write-Verbose "Exception being Handled - Throttled sleep for a bit"

            # throttled request or a temporary issue, wait for a few seconds and retry
            Start-Sleep -Seconds 120

			# We loop forever trying the same one till we get in sleep longer if we fail on our try
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

        }
		elseif($statusCode -eq 400)
		{
			Write-Verbose "Exception being Handled - This Blew up sorry bad request"
            Write-Output "400 log and skip to next user"
            Write-Output "Bad request for USER: $UPN $count"
			$Response = $null

		}
        
    }#end catch on first Invoke
    
	##
	#
	# - Run through the Response 
	#
	##
	if($Response -eq $null)
	{
		# Bad request response or its null 
		Write-Verbose "UserNull: $UPN $count bad request skip and move to next one"
        $ErrorFileData = New-Object psobject
	    # Filling in our data and adding into the file array 
	    $ErrorFileData | add-member -MemberType NoteProperty -Name DisplayName -Value $displayname
	    $ErrorFileData | add-member -MemberType NoteProperty -Name UPN -Value $UPN
	    $ErrorFileData | add-member -MemberType NoteProperty -Name Error -Value $statusCode
	    $ErrorFile += $ErrorFileData 
	}
	else{
		
		# Process each Item in Response
		foreach($responseitem in $Response)
		{
			if($Response.value.Count -eq 0)       
			{
				Write-Verbose "UserNull: $UPN $count"
				$ErrorFileData = New-Object psobject
				# Filling in our data and adding into the file array 
				$ErrorFileData | add-member -MemberType NoteProperty -Name DisplayName -Value $displayname
				$ErrorFileData | add-member -MemberType NoteProperty -Name UPN -Value $UPN
				$ErrorFileData | add-member -MemberType NoteProperty -Name Error -Value $statusCode
				$ErrorFile += $ErrorFileData 
	
			}
			else
			{

                # AD-User cant lookup with full UPN take UPN and try with Alias of it ie SamAccountName
                #get the AD onprem info
                $index = $UPN.ToString().IndexOf('@')
                if($index -ne -1)
                {
                    $SamLookup = $upn.ToString().SubString(0,$index)

                }

                ##
                #
                # - Try getting OnPrem User attributes that might not all be synced up
                #
                ##
                try
                {
                	$onpreminfo = Get-ADUser $Samlookup -properties * -ErrorAction Stop | select EmployeeID,EmployeeNumber,employeeType,LastLogonDate,whenCreated
                }
                catch
                {
                  #probally a cloud only user so special arent they
                  $ErrorFileData = New-Object psobject
                  # Filling in our data and adding into the file array 
                  $ErrorFileData | add-member -MemberType NoteProperty -Name DisplayName -Value $displayname
                  $ErrorFileData | add-member -MemberType NoteProperty -Name UPN -Value $UPN
                  $ErrorADFile += $ErrorFileData 
                }

				Write-Verbose "User: $UPN $count"
				Write-Output "User: "$UPN $count
				# take the values from the response
				$displayname = $responseitem.value.displayName
				$LogonEventID = $responseitem.value.id

				Write-Verbose -Message "$displayname"
				Write-Verbose -Message "$LogonEventID"
				
				# you get back displayname and ID by default sometimes there is activity returned
				if($responseitem.value.signInActivity.count -ne 0)
				{
					$lastSignDate =  $responseitem.value.signInActivity.lastSignInDateTime
					$lastNonSignDate =  $responseitem.value.signInActivity.lastNonInteractiveSignInDateTime
					$lastNonSignID = $responseitem.value.signInActivity.lastNonInteractiveSignInRequestId

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

                #simple date flip
                if($lastSignDate -ne $null)
				{
                	[System.DateTime]$SimpleDate = Get-Date($lastSignDate.tostring()) 	
				}
				if($lastNonSignDate -ne $null)
				{
                	[System.DateTime]$SimpleNonDate = Get-Date($lastNonSignDate.tostring()) 	
				}



				#item
				#-responseitem

				
				$FileData = New-Object psobject
				# Filling in our data and adding into the file array 
				$FileData | add-member -MemberType NoteProperty -Name DisplayName -Value $displayname
				$FileData | add-member -MemberType NoteProperty -Name UPN -Value $UPN
				$FileData | add-member -MemberType NoteProperty -Name lastSignInRequestId -Value $LogonEventID
				$FileData | add-member -MemberType NoteProperty -Name LastDateInteractive -Value $lastSignDate
				$FileData | add-member -MemberType NoteProperty -Name lastNonInteractiveSignInRequestId -Value $lastNonSignID
				$FileData | add-member -MemberType NoteProperty -Name LastDateNonInteractive -Value $lastNonSignDate
				$FileData | add-member -MemberType NoteProperty -Name OnPremADLastLogon -Value $onpreminfo.LastLogonDate
				$FileData | add-member -MemberType NoteProperty -Name OnPremCreated -Value $onpreminfo.whenCreated
				
                $FileData | add-member -MemberType NoteProperty -Name SimpleDateLastInteractive -Value $SimpleDate.ToShortDateString()
				$FileData | add-member -MemberType NoteProperty -Name SimpleDateNonInteractive -Value $SimpleNonDate.ToShortDateString()
				$FileData | add-member -MemberType NoteProperty -Name EmployeeID -Value $onpreminfo.EmployeeID
				$FileData | add-member -MemberType NoteProperty -Name EmployeeNumber -Value $onpreminfo.EmployeeNumber
				$FileData | add-member -MemberType NoteProperty -Name EmployeeType -Value $onpreminfo.employeeType
				
				$FileData | add-member -MemberType NoteProperty -Name Mail -Value $item.Mail
				$FileData | add-member -MemberType NoteProperty -Name ID -Value $item.id
				$FileData | add-member -MemberType NoteProperty -Name DirSyncEnabled -Value $item.DirSyncEnabled
				$FileData | add-member -MemberType NoteProperty -Name UserType -Value $item.UserType

				$File += $FileData     
				Write-Output $FileData
				$displayname,$UPN,$LogonEventID,$lastSignDate,$lastNonSignID,$lastNonSignDate,$SamLookup,$index,$onpreminfo = $null
			    $SimpleDate = Get-Date("3999-4-4")
			    $SimpleNonDate = Get-Date("3999-4-4")
			}

			# every 30 mins refresh the token
			$RefreshDiff = New-TimeSpan -Start $TokenRefreshTime -End (Get-Date)
			if ($RefreshDiff.Minutes -ge 30)
			{
				$MSALtoken = RefreshToken
				$TokenRefreshTime = Get-Date
			
			}
            
            #Every 1000 put a pin it for 5 mins refresh token is there a token refresh limit?

			#reset vars used in file writes
			$displayname,$UPN,$LogonEventID,$lastSignDate,$lastNonSignID,$lastNonSignDate = $null
		
		}#loop end for items in response 

		$count++

	}#else check for null bad values from return
}#loop end for list of users

#Export the file - "C:\{}.csv" 
$File | export-csv $targetFile  -NoTypeInformation

#Export the error file "C:\{}.csv"
$ErrorFile | export-csv  $ErrortargetFile -NoTypeInformation

#Export the AD error files 
$ErrorADFile | export-csv $ErrorADLookupFile -NoTypeInformation







