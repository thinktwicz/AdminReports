# you will need to fill in line 13,21,22,46
# you will need an Account with GA rights
# you will need to have setup an Azure App Registration with the right Permissions - This is using AuthMethods calls

# at this time the report list the users default auth method, if they have Push, if they have MS auth App, if they have just any Auth App, Stale count for both Auth Apps and Office location


#bring in the mods
Import-module MSAL.PS

Write-Host "Creating File" -BackgroundColor:Magenta -ForegroundColor:Green
#make a filepath
$targetFile = 'FilepathAndFileName' #example c:\temp\report.csv
#remove any version 
rm $targetFile -ErrorAction SilentlyContinue
Add-Content $targetFile "UPN, DisplayName, DefaultMethodType, HasAppPushNotifications, MS AuthApp, Other AuthApp, MultipleMSAuthApps, MultipleAuthApps, OfficeLocaiton"

Write-Host "Created" -BackgroundColor:Green -ForegroundColor:Blue


$clientid = "Client App ID"
$tenantid = "Tenant ID"

Write-Host "Getting GraphAPI Onbehalf Token" -BackgroundColor:Magenta -ForegroundColor:Green

$MSALtoken = Get-MsalToken -Interactive -ClientId $clientID -TenantId $tenantID

Write-Host "Token Granted" -BackgroundColor:Green -ForegroundColor:Blue


function RefreshToken()
{
    return $MSALtoken = Get-MsalToken -ForceRefresh -ClientId $clientID -TenantId $tenantID
}

Write-Host "Connecting to MSOL Service" -BackgroundColor:Magenta -ForegroundColor:Green

# used to get defualt method
Connect-MsolService

Write-Host "Connected" -BackgroundColor:Green -ForegroundColor:Blue

Write-Host "Importing User List and Starting" -BackgroundColor:Magenta -ForegroundColor:Green

# user list we check against 
$list = Import-Csv 'UserListtoReadFrom.csv' #Need displayName, userPrincipalName for the users that are targets for the report
$count = 0
$refreshOpCount = 0
# going to revoke
#Connect-AzureAD

foreach($user in $list)
{
    Write-Host "Count" $count "checking user" $user.displayName -BackgroundColor:DarkMagenta -ForegroundColor:Yellow

    #get the info from MSOL on methods and location
    $methodsettings = Get-MsolUser -UserPrincipalName $user.userPrincipalName | Select-Object -ExpandProperty StrongAuthenticationMethods 
    $office = Get-MsolUser -UserPrincipalName $user.userPrincipalName | Select-Object Office

    ###########
    #
    # find which method is set to default
    #
    ###########
        foreach($method in $methodsettings)
        {
    
            if($method.IsDefault -eq 'True')
            {
                $userMethod = $method.MethodType
            }

            if($method.MethodType -eq 'PhoneAppNotification')
            {
                $userHavePush = 'True'
            }
    
        }##// Method Default check
    ###########
    #
    # 
    #
    ###########

    ###########
    #
    # Office checking 
    #
    ###########
        $userupn = $user.userPrincipalName.ToString()
        $userdisplay = $user.displayName
        if($office.office -eq $null)
        {
           $office = 'Null'
        }
        else
        {
            $office = $office.office.tostring()
        }
    ###########
    #
    # 
    #
    ###########

    ###########
    #
    # Graph Call 
    #
    ###########

    #header token
    $headers  = @{Authorization = "Bearer $($MSALtoken.accesstoken)" }
    
    #user we are checking 
    $UPN = $user.userPrincipalName 

    #api endpoint
    $apiUrl = "https://graph.microsoft.com/beta/users/$($UPN)/authentication/methods"

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

    ###########
    #
    # Auth App Type check
    #
    ###########
    $StaleMSAuthCnt = 0
    $StaleAuthCnt = 0
        foreach($item in $Response.value)
        {

            if($item.'@odata.type' -eq "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod")
            {
                #Write-Host "MS Authenticator App"
                $msAuthApp = 'True'
                $StaleMSAuthCnt++
            }
    
            if($item.'@odata.type' -eq "#microsoft.graph.softwareOathAuthenticationMethod")
            {
                #Write-Host "Authenticator App"
                $otherAuthApp = "True"
                $StaleAuthCnt++

            }


        }##// Auth Method Type
    ###########
    #
    # 
    #
    ###########

    ###########
    #
    # making file
    #
    ###########
    $line = "$userupn,$userdisplay,$userMethod,$userHavePush,$msAuthApp,$otherAuthApp,$StaleMSAuthCnt,$StaleAuthCnt,$office"

    Write-Host "Details:" -BackgroundColor:Cyan -ForegroundColor:DarkGreen
    Write-Host $line -BackgroundColor:DarkCyan -ForegroundColor:Yellow
       
    Add-Content $targetFile $line
  
    ###########
    #
    # resetting file info
    #
    ###########
    $userupn = 'Null'
    $userMethod = 'Null'
    $userdisplay = 'Null'
    $userHavePush = 'Null'
    $officelocation = 'Null'
    $msAuthApp = 'Null'
    $otherAuthApp = 'Null'

    # display counter increase 
    $count++

    # refresh the token
    $refreshOpCount++
    if($refreshOpCount -eq 1000)
    {
        $MSALtoken = RefreshToken
        $refreshOpCount = 0
    }

}##// User in list

Write-Host "Finished" -BackgroundColor:Green -ForegroundColor:Blue
