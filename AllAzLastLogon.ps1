<#
	Author: Drew Crouch
	Purpose: Grabs a list of all accounts and runs through them getting last logon. If there is no entry (null) or the date is outside
				our date ranage then that account shows up on the report. Last 30 Days

				Stale acconts will then be marked and processes started for removal/investigation 


#>
Write-Output "Connect to MSOL and Azure AD Powershells"

# MSOL Creds 
$MSOLCreds = Get-AutomationPSCredential -Name ''

# MSOL
Connect-MsolService -credential $MSOLCreds

# AzureAD
Connect-AzureAD -credential $MSOLCreds

Write-Output "Connect to MSOL and Azure AD Powershells"
#C:\

Get-MsolUser -All | where{$_.UserType -eq "Member"} | select UserPrincipalName,objectid,displayname | Export-Csv 'C:\.csv' -notypeinformation

Write-Output "File has been made"

$list = Import-Csv 'C:\.csv'

#make a filepath
$targetFile = 'C:\AllAzLastLogon.csv'
#remove any version 
rm $targetFile -ErrorAction SilentlyContinue
Add-Content $targetFile "userPrincipalName,DisplayName,EmployeeID,Manager,LastLogonDate"

[datetime]$myStart = (Get-date -Format "MM-dd-yyyy")

$myPastDay = (Get-date).AddDays(-30)
[datetime]$30dayAgo = $myPastDay.ToString("MM-dd-yyyy")
$count = 0

Write-Output "Math has been sorted and we are going through the list"


foreach($user in $list)
{
    $UPN = $user.UserPrincipalName
    Write-Output "User we are checking" $UPN $count

    ## This has a timeout -Message: Too Many Requests need to have a sleep 
    # the OData 3 filter returns UPN as all lower so FYI to that one 
    #$catch = Get-AzureADAuditSignInLogs -Top 1 -Filter "UserPrincipalName eq '$UPN' "  | select CreatedDateTime,AppDisplayName
    #$filter = "`"" + 'UserPrincipalName eq ' + "`'" + $UPN + "`'" +"`""
    $lower = $UPN.ToLower()
    $catch = Get-AzureADAuditSignInLogs -Top 1 -Filter "UserPrincipalName eq '$lower'" | select UserPrincipalName,AppDisplayName,CreatedDateTime
    #$catch = Get-AzureADAuditSignInLogs -Top 1 -Filter $filter.ToString() | select UserPrincipalName,AppDisplayName,CreatedDateTime
    if($catch -eq $null)
    {
        [string]$userLogin = "Null"
    }
    else
    {
        $userLogon = (get-date $catch.CreatedDateTime)
        [datetime]$userLogin = $userLogon.ToString("MM-dd-yyyy")
    }

    if(($userLogin -le $myStart) -and ($userLogin -ge $30dayAgo))
    {
        Write-Output "you are not on the report"
    }
    else
    {
        Write-Output "you are on the report" $UPN $catch
        #hybrid call the other values however you want need 
        $line = "$UPN,$UPN,$UPN,$UPN,$userLogin"
        Add-Content $targetFile $line
    }
    $catch = $null
    Start-Sleep -Seconds 5
    $count++
}





