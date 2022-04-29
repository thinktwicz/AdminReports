<#
get a list of all users from MG Graph API Module add in missing things from AzureAD into a bigger list
10-12 need file paths - exporting out to, importing in from and the new out put file location
todo: error handling always..... errors always

#>
Connect-MgGraph -Scopes "User.Read.All"
Connect-AzureAD

Get-MgUser -All | select Id,DisplayName,Mail,UserPrincipalName | export-csv "c:\my.csv"
$list = import-csv "c:\my.csv"
$targetFile = "c:\myNew.csv"

$File = @()

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
