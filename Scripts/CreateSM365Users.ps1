[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [System.IO.FileInfo]
    $UserCSVPath
)

# Set Execution Policy, download and import AzureAD module
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Install-Module AzureAD -Force -Scope CurrentUser -Confirm:$false
Import-Module AzureAD

# Connect to AzureAD
Connect-AzureAD

# Store userlist in variable
$userlist = Import-Csv -Path $UserCSVPath

# Create Password and Passwordprofile
Add-Type -AssemblyName 'System.Web'
$Password = [System.Web.Security.Membership]::GeneratePassword(12, 3)
$PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
$PasswordProfile.Password = $password
$PasswordProfile.ForceChangePasswordNextLogin = $false
$PasswordProfile.EnforceChangePasswordPolicy = $false

# Retrieve Tenant-Domain
$TenantDomain = (Get-AzureADDomain | Where-Object Name -like "*.onmicrosoft.com").Name

# Loop through userlist and create users
Foreach ($user in $userlist)
{
    # build UPN
    $upnLeftpart = $user.GivenName + $user.Surname.Substring(0,1)
    $UpnRightpart = "@" + $TenantDomain
    
    # Create User
    try
    {
        New-AzureADUser -UserPrincipalName ($upnLeftpart + $UpnRightpart) -DisplayName $user.DisplayName -City $user.City -Country $user.Country -GivenName $user.GivenName -Surname $user.Surname -Department $user.Department -JobTitle $user.JobTitle -PhysicalDeliveryOfficeName $user.PhysicalDeliveryOfficeName -PostalCode $user.PostalCode -State $user.State -StreetAddress $user.StreetAddress -TelephoneNumber $user.TelephoneNumber -UsageLocation $user.UsageLocation -AccountEnabled $true -MailNickName $upnLeftpart -PasswordProfile $PasswordProfile -ErrorAction Stop
        Start-Sleep -Seconds 2
        Write-Host -ForegroundColor Green "Successfully created user $($User.Displayname)."
    }

    catch
    {
        Write-Host -ForegroundColor Red "Error creating user $($User.DisplayName). The error is: $($_.ErrorDetails)"
        Throw
    }
}

# output password for all users to screen and file in Directory where usercsv is stored
$UserPasswordFile = "UserPassword.txt"
Write-Host -ForegroundColor Green "Password for all Users is $password. This is also stored in file $UserPasswordFile in directory $($UserCSVPath.Directory)"
Set-Content -Value $Password -Path (Join-path -Path $UserCSVPath.Directory -ChildPath $UserPasswordFile)

