<# 
365-UserMgmt-AccessSetup.ps1

MWM, Claude assisted cleanup and recommendations/optimizations - 2026-10
Configure an enterprise application in Entra ID to enable login to 365 via certificate, scoped admin consent to User creation, group/distribution list membership, and Sharepoint site membership.
#>

# Initial checks/steps
# Set the SecurityProtocol to TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Verify that the system is running Powershell 7.0+ and Modules are available
#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Applications, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Identity.Governance

# Disconnect any currently open Office 365 connections
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

# Main script

# Where to store your certs
$cert_folder = Read-Host "Please enter the complete folder path to store certificates in.  Ex. c:\users\username\365-Certificates"

# Create a password for the certificate
$password_token = Read-Host -AsSecureString "Please enter a password to use for the certificate - make note of it in your password manager!"

Write-Host "Tenant Name needed"
$TenantName = Read-Host "Please enter a single word/no spaces name to use on the certificate."

# Define certificate name and path
$certname = "$TenantName-UserMgmt"
$certFilePath = "$cert_folder\$certname.cer"

# Check if the certificate already exists
Write-Output "Checking for Certificate"
if (-not (Test-Path -Path $certFilePath)) {
    Write-Host "Certificate does not exist. Creating the certificate..."
    $ErrorActionPreference = "Stop"
    # Create local certificate
    $cert = New-SelfSignedCertificate -Subject "CN=$certName" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256

    # Export the certificate to .pfx and .cer
    Export-PfxCertificate -Cert $cert -FilePath "$cert_folder\$certName.pfx" -Password $password_token
    Export-Certificate -Cert $cert -FilePath $certFilePath
    Write-Host "Certificate created and exported successfully."
    $Thumbprint = $Cert.Thumbprint
    $certStartDate = $Cert.NotBefore
    $certEndDate = $Cert.NotAfter
    else {
    Write-Host "Certificate already exists. Skipping certificate creation."
    $cert = [Security.Cryptography.X509Certificates.X509Certificate2]::new($certFilePath)
    $Thumbprint = $cert.Thumbprint
    $certStartDate = $cert.NotBefore
    $certEndDate = $cert.NotAfter
    }
}
#Store Certificate Key in Variable
[convert]::ToBase64String((Get-Content "$cert_folder\$certName.cer" -AsByteStream)) | Out-File -FilePath "$cert_folder\$certName.key.txt"
$Key = Get-Content "$cert_folder\$($certName).key.txt"

# Authenticate to Microsoft Graph
Write-Host "Connecting to Microsoft Graph"
Connect-MgGraph -Scopes AppRoleAssignment.ReadWrite.All,Application.ReadWrite.All,RoleManagement.ReadWrite.Directory -ContextScope Process

# Verify that we're in the correct tenant

Write-host (Get-MgOrganization).DisplayName
Write-host (Get-MgContext).Account
Write-host "Please verify that this is the correct tenant?  Ctrl-C to break/exit"
pause

# Get the Tenant ID
$tenantId = (Get-MgOrganization).Id

#Define API Permissions for Registered App
$requiredGrants = New-Object -TypeName System.Collections.Generic.List[Microsoft.Graph.PowerShell.Models.MicrosoftGraphRequiredResourceAccess]
$requiredResourceAccess = New-Object -TypeName Microsoft.Graph.PowerShell.Models.MicrosoftGraphRequiredResourceAccess
$requiredResourceAccess.ResourceAppId = "00000003-0000-0000-c000-000000000000"
$requiredResourceAccess.ResourceAccess+=@{ Id = "741f803b-c850-494e-b5df-cde7c675a1ca"; Type = "Role" }
$requiredResourceAccess.ResourceAccess+=@{ Id = "dbaae8cf-10b5-4b86-a4a1-f871c94c6695"; Type = "Role" }
$requiredResourceAccess.ResourceAccess+=@{ Id = "0121dc95-1b9f-4aed-8bac-58c5ac466691"; Type = "Role" }
$requiredResourceAccess.ResourceAccess+=@{ Id = "a82116e5-55eb-4c41-a434-62fe8a61c773"; Type = "Role" }
# Exchange Online - Exchange.ManageAsApp (app-only Connect-ExchangeOnline)
$exoResourceAccess = New-Object -TypeName Microsoft.Graph.PowerShell.Models.MicrosoftGraphRequiredResourceAccess
$exoResourceAccess.ResourceAppId = "00000002-0000-0ff1-ce00-000000000000"
$exoResourceAccess.ResourceAccess+=@{ Id = "dc50a0fb-09a3-484d-be87-e023b12c6440"; Type = "Role" }

$requiredGrants.Add($requiredResourceAccess)
$requiredGrants.Add($exoResourceAccess)

# Define the app name
$appName = "$certname"

# Get all registered applications in the directory
$apps = Get-MgApplication -all

# Check if the app already exists
$app = $apps | Where-Object { $_.DisplayName -eq $appName }
Write-Host "Tenant: $TenantName"
Write-Host "Tenant ID: $tenantId"
Write-Host "App ID: $($app.AppId)"
Write-Host "Thumbprint: $thumbprint"
if ($null -eq $app) {
    # If the app does not exist, create it
    Write-Host "App '$appName' does not exist. Creating the app..."
    $newApp = New-MgApplication -DisplayName $appName -SignInAudience "AzureADMyOrg" -RequiredResourceAccess $requiredGrants
    Write-Host "App '$appName' created successfully with ID: $($newApp.Id)"
    $app = $newApp
    start-sleep -Seconds 3
} else {
    #If the app exists, ensure it has all API Permissions
    Write-Host "App '$appName' already exists with ID: $($app.Id), Adding API Permissions..."
    Update-MgApplication -ApplicationId $app.Id  -RequiredResourceAccess $requiredGrants
}

$sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue
if (-not $sp) {
    Write-Host "Service Principal not found. Retrying in 5 seconds..."
    $sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 5
}

if (-not $sp) {
    Write-Host "Service Principal not found. Creating it..."
    $sp = New-MgServicePrincipal -AppId $app.AppId
}
$ErrorActionPreference = "Stop"

# Give admin consent to each of the RequiredGrants

foreach ($grant in $requiredGrants) {
    # Service principal of the API being granted (Graph, Exchange Online, ...)
    $resourceSp = Get-MgServicePrincipal -Filter "appId eq '$($grant.ResourceAppId)'"
    if (-not $resourceSp) {
        Write-Warning "Resource $($grant.ResourceAppId) not found in this tenant - skipping consent for it."
        continue
    }

    # Roles already granted, so re-runs don't fail on duplicates
    $assigned = (Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All |
        Where-Object ResourceId -eq $resourceSp.Id).AppRoleId

    foreach ($access in $grant.ResourceAccess) {
        if ($access.Id -in $assigned) {
            Write-Host "Already granted: $($access.Id)"
            continue
        }
        $roleName = ($resourceSp.AppRoles | Where-Object Id -eq $access.Id).Value
        Write-Host "Granting $roleName on $($resourceSp.DisplayName)"
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id `
            -PrincipalId $sp.Id -ResourceId $resourceSp.Id -AppRoleId $access.Id | Out-Null
    }
}

# Upload a certificate to the app
$params = @{
	keyCredentials = @(
		@{
			endDateTime = [System.DateTime]::Parse($certEndDate)
			startDateTime = [System.DateTime]::Parse($certStartDate)
			type = "AsymmetricX509Cert"
			usage = "Verify"
			key = [System.Text.Encoding]::ASCII.GetBytes($Key)
            displayName = "From Server"
		}
	)
}

Write-Host "Uploading certificate to app registration"
Update-MgApplication -ApplicationId $app.Id -BodyParameter $params

# Define the file path for the CSV
$filePath = "$cert_folder\$certname.csv"

# Check if the file exists, if not, create it and export the data
if (-not (Test-Path $filePath)) {
    Write-Host "File '$filePath' does not exist. Creating the file and exporting data..."

    # Export data to CSV
    [PSCustomObject] [ordered] @{
            Customer=$tenantName
            ApplicationID=$app.AppId
            TenantID=$tenantId
            Thumbprint=$thumbprint
            Expires=$((Get-ChildItem Cert:\CurrentUser\My\$($Thumbprint)).NotAfter)
        } | ConvertTo-Csv -NoTypeInformation | ForEach-Object { $_ -replace '"', '' } | Set-Content -Path $FilePath
    Write-Host "Data exported to '$filePath'."
} 
else {
    # Export data to CSV
    [PSCustomObject] [ordered] @{
            Customer=$($tenantName)
            ApplicationID=$($app.AppId)
            TenantID=$($tenantId)
            Thumbprint=$($thumbprint)
            Expires=$($cert.NotAfter)
        } | ConvertTo-Csv -NoTypeInformation -NoHeader | ForEach-Object { $_ -replace '"', ''} | Add-Content -Path $FilePath 
    Write-Host "Data appended to existing '$filePath'."
}
Remove-Item "$cert_folder\$($certName).key.txt" -Force

Write-Host "Company: $tenantName"
Write-Host "Tenant ID: $tenantId"
Write-Host "Application ID: $($app.AppId)"
Write-Host "Thumbprint: $thumbprint"
Write-Host "Expires: $($cert.NotAfter)"

# Close the powershell session
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
