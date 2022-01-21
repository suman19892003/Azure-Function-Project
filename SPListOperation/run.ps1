#This is Azure Function which perform the operation of following
#getting list items based on ID
#getting All List Items
#Return value in JSON format for GET Method

#Parameter to pass in Body should be in JSON Object as
# {
#   "Operation":"Get", or "Add" or "Update"
#   "Method": "GetAll/GetById",
#   "ItemId":19
# }

using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Get Parameter Values Passed in Body from the request
$Operation=$Request.Body.Operation #Get/Add/Update/Delete
$ItemId=$Request.Body.ItemId
$Method=$Request.Body.Method #GetAll/GetById

#Variable Initialization
$string = $ItemId
$idVal = [int]$string
$myObj=@{} #Object Declaration
$AllItems=@() #Array Declaration

$securePassword = ConvertTo-SecureString $env:tenant_pwd -AsPlainText -Force
$credentials = New-Object PSCredential ($env:tenant_user, $securePassword)

Connect-PnPOnline -Url https://libertymutual.sharepoint.com/sites/SmartekClaimsform_dev/ -Credentials $credentials

if($Operation -eq "Get"){
    if($Method -eq "GetById"){
        #Get Items Based on ID
        $Query= "<View><ViewFields><FieldRef Name='Title'/><FieldRef Name='Description'/><FieldRef Name='ID'/></ViewFields><Query><Where><Eq><FieldRef Name='ID' /><Value Type='FSObjType'>$idVal</Value></Eq></Where><OrderBy><FieldRef Name='Modified' Ascending='FALSE' /></OrderBy></Query></View>"
    }
    else{
        #Get All Items
        $Query= "<View><ViewFields><FieldRef Name='Title'/><FieldRef Name='Description'/><FieldRef Name='ID'/></ViewFields><Query><OrderBy><FieldRef Name='Modified' Ascending='FALSE' /></OrderBy></Query></View>"
    }
    $listItems= (Get-PnPListItem -List "MyTestList" -Query $Query)
    if($listItems.Count -gt 0)
    {
        foreach($item in $listItems)
        {
            $myObj.Title= $item["Title"]
            $myObj.Description= $item["Description"]
            $ResultObj=(New-Object PSObject -Property $myObj)
            $AllItems +=$ResultObj
        }
    }
    $body=$AllItems
    
}
if($Operation -eq "Add"){
    Add-PnPListItem -List "MyTestList" -Values @{"Title" = "PowerShell title";"Description" ="Description Added from Powershell"}
    $body="Item Added Successfully"
}

if($Operation -eq "Update"){
    Set-PnPListItem -List "MyTestList" -Identity $idVal  -Values @{"Title" = "Updated PS"}
    $body="Item Updated Successfully"
}

#$body="No Operation2"

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})