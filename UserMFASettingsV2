<#

22,30,31,57


 we are getting the methods we are seaching for per user and for each of those listed to the user report on them so we can action on them
 https://docs.microsoft.com/en-us/graph/authenticationmethods-get-started
- Delegated Permisisons needed
UserAuthenticationMethod.Read.All, UserAuthenticationMethod.ReadWrite.All








#>
#bring in the mods
Import-module MSAL.PS

Write-Host "Creating File" -BackgroundColor:Magenta -ForegroundColor:Green
#make a filepath
$targetFile = 'ReportFilepathandName'
#remove any version 
rm $targetFile -ErrorAction SilentlyContinue
Add-Content $targetFile "UPN, DisplayName, DefaultMethodType, HasAppPushNotifications, MS AuthApp, Other AuthApp, MultipleMSAuthApps, MultipleAuthApps, OfficeLocaiton, MFAMethodID, MethodsDisplayName, MethodDeviceTag, MethodCreatedDate, MethodEmail, MicrosoftMethodType"

Write-Host "Created" -BackgroundColor:Green -ForegroundColor:Blue


$clientid = ""
$tenantid = ""

Write-Host "Getting GraphAPI Onbehalf Token" -BackgroundColor:Magenta -ForegroundColor:Green

$MSALtoken = Get-MsalToken -Interactive -ClientId $clientID -TenantId $tenantID

Write-Host "Token Granted" -BackgroundColor:Green -ForegroundColor:Blue

# Time things
$ScriptRunTime = Get-Date
$TokenRefreshTime = Get-Date

# Force a Refresh on our Token helper function 
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
$list = Import-Csv 'ListofUserstoWorkwith'
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
    $UPN = $user.userPrincipalName.tostring()

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
    
    $MFAMethodID = ""
    $MethodsDisplayName = ""
    $MethodDeviceTag = ""
    $MethodCreatedDate = ""
    $MicrosoftMethodType = ""
    $MethodEmail = ""

    $line = "$userupn,$userdisplay,$userMethod,$userHavePush,$msAuthApp,$otherAuthApp,$StaleMSAuthCnt,$StaleAuthCnt,$office,$MFAMethodID,$MethodsDisplayName,$MethodDeviceTag,$MethodCreatedDate,$MethodEmail,$MicrosoftMethodType"

    Write-Host "Details:" -BackgroundColor:Cyan -ForegroundColor:DarkGreen
    Write-Host $line -BackgroundColor:DarkCyan -ForegroundColor:Yellow

    #add in the first part of the users information       
    Add-Content $targetFile $line

    $userMethod = ""
    $userHavePush = ""
    $msAuthApp = ""
    $otherAuthApp = ""
    $StaleMSAuthCnt = ""
    $StaleAuthCnt = ""
    $office = ""


    #creating the second and or multiple lines of user method infomation
    foreach($item in $Response.value)
    {
        if($item.'@odata.type' -eq "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod")
        {
            $MFAMethodID = $item.id
            $MethodsDisplayName = $item.displayname
            $MethodDeviceTag = $item.deviceTag
            $MethodCreatedDate = $item.createdDateTime
            $MicrosoftMethodType = "Microsoft Authenticator"
            $line = "$userupn,$userdisplay,$userMethod,$userHavePush,$msAuthApp,$otherAuthApp,$StaleMSAuthCnt,$StaleAuthCnt,$office,$MFAMethodID,$MethodsDisplayName,$MethodDeviceTag,$MethodCreatedDate,$MethodEmail,$MicrosoftMethodType"
            Add-Content $targetFile $line
    
        }
        if($item.'@odata.type' -eq "#microsoft.graph.softwareOathAuthenticationMethod")
        {
            $MFAMethodID = $item.id
            $MethodsDisplayName = $item.displayname
            $MethodDeviceTag = $item.deviceTag
            $MethodCreatedDate = $item.createdDateTime
            $MicrosoftMethodType = "Authenticator App"

            $line = "$userupn,$userdisplay,$userMethod,$userHavePush,$msAuthApp,$otherAuthApp,$StaleMSAuthCnt,$StaleAuthCnt,$office,$MFAMethodID,$MethodsDisplayName,$MethodDeviceTag,$MethodCreatedDate,$MethodEmail,$MicrosoftMethodType"
            Add-Content $targetFile $line
        } 
        if($item.'@odata.type' -eq "#microsoft.graph.emailAuthenticationMethod")
        {
            $MFAMethodID = $item.id
            $MethodEmail = $item.emailAddress
            $MicrosoftMethodType = "Email"
            $line = "$userupn,$userdisplay,$userMethod,$userHavePush,$msAuthApp,$otherAuthApp,$StaleMSAuthCnt,$StaleAuthCnt,$office,$MFAMethodID,$MethodsDisplayName,$MethodDeviceTag,$MethodCreatedDate,$MethodEmail,$MicrosoftMethodType"
            Add-Content $targetFile $line

        }

        $MFAMethodID = ""
        $MethodsDisplayName = ""
        $MethodDeviceTag = ""
        $MethodCreatedDate = ""
        $MethodEmail = ""
        $MicrosoftMethodType = ""
    }

  
    ###########
    #
    # resetting file info
    #
    ###########
    $userupn = 'Null'
    $userMethod = 'Null'
    $userdisplay = 'Null'
    $userHavePush = 'Null'
    $office = 'Null'
    $msAuthApp = 'Null'
    $otherAuthApp = 'Null'

    # display counter increase 
    $count++

    # refresh the token every 30mins
    $RefreshDiff = New-TimeSpan -Start $TokenRefreshTime -End (Get-Date)
    if ($RefreshDiff.Minutes -ge 30)
    {
        $MSALtoken = RefreshToken
        $TokenRefreshTime = Get-Date
        
    }
    # refresh the token every 1000 operations 
    $refreshOpCount++
    if($refreshOpCount -eq 1000)
    {
        $MSALtoken = RefreshToken
        $refreshOpCount = 0
    }

}##// User in list

Write-Host "Finished" -BackgroundColor:Green -ForegroundColor:Blue
$diff = New-TimeSpan -Start $ScriptRunTime -End (Get-Date)
Write-Host "Run Time:" "Days:"$diff.Days "Hours:"$diff.Hours "Minutes:"$diff.Minutes "Seconds:"$diff.Seconds
