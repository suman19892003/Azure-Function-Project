#This is Azure Function which perform the operation of following
#getting All SPO Users from SharePoint Site and their Property
#getting SPO User Property based on Email ID
#And Ulpoad the CSV file to local Azure Function Dir

#Parameter to pass in Body should be in JSON Object as
# {
#   "Method":"Get",
#   "Email": "Suman.kumar01@libertymutual.com",
#   "Operation":"GetById"
# }

using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

$UserAccount=$Request.Body.Email
$Operation=$Request.Body.Operation
$Method=$Request.Body.Method

$UserData=@()
$securePassword = ConvertTo-SecureString $env:tenant_pwd -AsPlainText -Force
$credentials = New-Object PSCredential ($env:tenant_user, $securePassword)

Connect-PnPOnline -Url https://libertymutual.sharepoint.com/sites/SmartekClaimsform_dev/ -Credentials $credentials

if($Operation -eq "GetAll"){
    #Get All users of the site collection
    $Users = Get-PnPUser

    #Loop through Users and get properties
    ForEach ($User in $Users)
    {
        $UserData += New-Object PSObject -Property @{
            "User Name" = $User.Title
            "Login ID" = $User.LoginName
            "E-mail" = $User.Email
            "User Type" = $User.PrincipalType
        }        
    }
}
#if($Operation -eq "GetById"){
else{
    $Users = Get-PnPUserProfileProperty -Account $UserAccount | Select * -Expand UserProfileProperties

    $UserData = New-Object PSObject -Property @{
        "First Name" = $Users.FirstName
        "Last Name" = $Users.LastName
        "Login ID" = $Users.AccountName
        "User Type" = $Users.Title
    }
}

$date = Get-Date -format "dd-MMM-yyyy"
$filenamecsv = "AllUsersNew "+$date+".csv"

#Upload to local directory in Azure Function
$UserData | Export-Csv "C:\home\temp\$filenamecsv" -NoTypeInformation

$body = "User Details Fetched successfully and Uploaded as CSV inside Azure local temp folder"

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
