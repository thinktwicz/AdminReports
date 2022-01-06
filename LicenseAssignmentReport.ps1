

<#
  Run through all the MSOL users with a valid license and report if its Direct user assignment or InHerited by Group Base assignment
  
  Like most of my scripts you will need to read and take and actual part in understand how it works and make changes to meet your needs below are some edits that you will have
  take care of
  
  This makes a connection to MSOL will need admin creds for lines 81 - 88
  This makes at the very least two files the first is from MSOL lines 81 - 88 user and the second is once the report is generated out from memory into a cvs file line 146
  at the end if you want it can send out via email message using a relay server if you have one setup to use or can comment out this section of code your choice 
  
  
#>
# get new sku names for checking
# set so an email fires off 

<#
  Helper Functions that do the work on if its directly or inheritedly assigned 

#>

#Returns TRUE if the user has the license assigned directly
function UserHasLicenseAssignedDirectly
{
    Param([Microsoft.Online.Administration.User]$user, [string]$skuId)

    foreach($license in $user.Licenses)
    {
        #we look for the specific license SKU in all licenses assigned to the user
        if ($license.AccountSkuId -ieq $skuId)
        {
            #GroupsAssigningLicense contains a collection of IDs of objects assigning the license
            #This could be a group object or a user object (contrary to what the name suggests)
            #If the collection is empty, this means the license is assigned directly - this is the case for users who have never been licensed via groups in the past
            if ($license.GroupsAssigningLicense.Count -eq 0)
            {
                return $true
            }

            #If the collection contains the ID of the user object, this means the license is assigned directly
            #Note: the license may also be assigned through one or more groups in addition to being assigned directly
            foreach ($assignmentSource in $license.GroupsAssigningLicense)
            {
                if ($assignmentSource -ieq $user.ObjectId)
                {
                    return $true
                }
            }
            return $false
        }
    }
    return $false
}
#Returns TRUE if the user is inheriting the license from a group
function UserHasLicenseAssignedFromGroup
{
    Param([Microsoft.Online.Administration.User]$user, [string]$skuId)

    foreach($license in $user.Licenses)
    {
        #we look for the specific license SKU in all licenses assigned to the user
        if ($license.AccountSkuId -ieq $skuId)
        {
            #GroupsAssigningLicense contains a collection of IDs of objects assigning the license
            #This could be a group object or a user object (contrary to what the name suggests)
            foreach ($assignmentSource in $license.GroupsAssigningLicense)
            {
                #If the collection contains at least one ID not matching the user ID this means that the license is inherited from a group.
                #Note: the license may also be assigned directly in addition to being inherited
                if ($assignmentSource -ine $user.ObjectId)
                {
                    return $true
                }
            }
            return $false
        }
    }
    return $false
}

# Connecting to MSOL to get all the user to make a file to then read from 
Connect-MsolService -Credential 
#get a list of all user things that have a vaild license of some kind
Get-MsolUser -All | where {$_.isLicensed -eq "TRUE"} | select UserPrincipalName | export-csv c:\.csv

$MSOLUserList = Import-Csv c:\.csv

#get the new skus 
# MSFT Changes these Sku's daily almost at times depending on what you are doing to your directory and volume license changes 
$SkuArray = Get-MsolAccountSku

#-----------------------------------------------------------------------------#
<#
    go through the entire directory and make a listed based on direct assignment 
        or inherited assigment per user for all skus 
    
#>

#FileData - this is the raw data we pulled from powershell 
$FileData = New-Object psobject
#File - this is an arrary that we just keep adding our FileData to and once done we have a completed file
$File = @()
$myTrue = "True"
$myFalse = "False"

for($j = 0; $j -le $MSOLUserList.Count - 1; $j++)
{

    $User = (Get-MsolUser -UserPrincipalName $MSOLUserList[$j].UserPrincipalName.ToString())

    for($i  = 0; $i -le $SkuArray.Count - 1; $i++)
    {

       $FileData = New-Object psobject
        #Filling in our data and adding into the file array 
        $FileData | add-member -MemberType NoteProperty -Name Name -Value $MSOLUserList[$j].UserPrincipalName.ToString()
        $FileData | add-member -MemberType NoteProperty -Name SoftwareSkuID -Value $SkuArray[$i].AccountSkuId.ToString()

        $catch = UserHasLicenseAssignedDirectly $User $SkuArray[$i].AccountSkuId

        if($catch)
        {
             $FileData | add-member -MemberType NoteProperty -Name LicenseDirect -Value $myTrue
        }
        else
        {
             $FileData | add-member -MemberType NoteProperty -Name LicenseDirect -Value $myFalse
        }

        $catch = UserHasLicenseAssignedFromGroup $User $SkuArray[$i].AccountSkuId
        
        if($catch)
        {
            $FileData | add-member -MemberType NoteProperty -Name LicenseInherited -Value $myTrue
        }
        else
        {
           $FileData | add-member -MemberType NoteProperty -Name LicenseInherited -Value $myFalse      
        }
       
        $File += $FileData 
    }
}
#Export the file
$File | export-csv  c:\.csv -NoTypeInformation


$body = "Hello`r`tThis is a generated License report you requested. Attached will be the report for Direct or Inherited licenses."


# message recipients
$recipients = @("user@domain.com")

Send-Mailmessage -smtpServer "RelayServer" -from "user@domain.com" -to $recipients -subject "License Report" -body $body  -Attachments c:\.csv


<#
    we put files onto the system clean up after your self
    file clean up 
    path
    path

#>
#clean up
remove-item -path c:\.csv  
remove-item -path c:\.csv 


# 
