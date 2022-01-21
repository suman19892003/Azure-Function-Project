#This is Azure Function which perform the operation of getting all list items and then upload the data
# in the local Azure Function path along with Upload the JSON and CSV file in SP Document Library

using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
$parameter=$Request.Body.Method
Write-Host "Performing Operation of $parameter"

$script = {
    $securePassword = ConvertTo-SecureString $env:tenant_pwd -AsPlainText -Force
    $credentials = New-Object PSCredential ($env:tenant_user, $securePassword)

    Connect-PnPOnline -Url https://libertymutual.sharepoint.com/sites/SmartekClaimsform_dev/ -Credentials $credentials

    $listItems= (Get-PnPListItem -List "MyTestList" -Fields "Title","ID", "Description","PublishDate","Author","Created")
    $Target= @()
    foreach($listItem in $listItems){
        $TargetProperties = @{ID=$listItem["ID"];Title=$listItem["Title"];PublishDate=$listItem["PublishDate"].Date;Employee=$listItem["Author"].LookupValue;Created=$listItem["Created"].Date}  
        $TargetObject = New-Object PSObject –Property $TargetProperties 
        $Target +=  $TargetObject
    }
    $date = Get-Date -format "dd-MMM-yyyy"
    $filenamejson = "AllListItems "+$date+".json"
    $filenamecsv = "AllListItems "+$date+".csv"

    #Upload to local directory in Azure Function
    $Target | export-csv "C:\home\temp\$filenamecsv" -notypeinformation
    
    #Upload CSV File
    Add-PnPFile -Path "C:\home\temp\$filenamecsv" -Folder "MyTestLibrary"

    #Upload JSON File
    $json = ConvertTo-Json @($Target)  -Compress
    $json ='{ "value":'+ $json + '}'
    $Stream = [IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes($json)) 
    Add-PnPFile  -Folder MyTestLibrary -Stream $Stream -FileName $filenamejson
}

$webTitle = Start-ThreadJob -Script $script | Receive-Job -Wait

$body = "Performed Operation Successfully for $parameter "

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body = $body
    })