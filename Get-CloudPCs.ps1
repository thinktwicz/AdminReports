
#https://docs.microsoft.com/en-us/graph/api/virtualendpoint-list-cloudpcs?view=graph-rest-beta&tabs=http

$apiUrl = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs"

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
{}

#might need to do a do while loop page loop if list is big - https://github.com/thinktwicz/PowerShellExamples/blob/main/PageLooping.ps1
$Response | fl

$JSponse = $Response.value | ConvertTo-Json

$devices = $JSponse | ConvertFrom-Json

$count = 0
foreach($item in $devices)
{
$count

$item.displayName
Write-Output "UPN: $($item.userPrincipalName) "
Write-Output "aadDeviceID: $($item.aadDeviceId)"
Write-Output "Status: $($item.status)"
Write-Output "managedDeviceID: $($item.managedDeviceId)"


$count++
}
