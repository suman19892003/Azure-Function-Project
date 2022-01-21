using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

$securePassword = ConvertTo-SecureString $env:tenant_pwd -AsPlainText -Force
$credentials = New-Object PSCredential ($env:tenant_user, $securePassword)

#Import-Module .\Modules\AzureAD

Connect-AzureAD -Credential $credentials

$users = Get-AzureADUser -Filter "startswith(UserPrincipalName,'suman')" -All $true |
Select-Object @{N='IdName';E={if($_.UserPrincipalName){$_.UserPrincipalName}else{""}}}

$json = ConvertTo-Json @($users)  -Compress
$json ='{ "value":'+ $json + '}'

$Stream = [IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes($json))

Write-Output "Connecting to SP List"

Connect-PnPOnline -Url https://libertymutual.sharepoint.com/sites/SmartekClaimsform_dev -Credential $credentials
Write-Output "SharePoint Connection successful in Web Jobs"
$splib = "MyTestLibrary"  
$data=Add-PnPFile  -Folder $splib -Stream $Stream -FileName "AzureAd.json"

Write-Output "All Task completed."

$body = "This HTTP triggered function executed successfully."

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
