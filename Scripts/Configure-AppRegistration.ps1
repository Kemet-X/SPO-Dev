# Configure-AppRegistration.ps1
# Uploads a certificate to Azure AD app registration for SharePoint authentication

param(
    [Parameter(Mandatory = $true)]
    [string]$CertificatePath,
    
    [Parameter(Mandatory = $true)]
    [string]$ApplicationId,
    
    [Parameter(Mandatory = $false)]
    [string]$DisplayName = "SharePoint RER Certificate"
)

Write-Host "=== Configure Azure AD App Registration with Certificate ===" -ForegroundColor Cyan

# Check if certificate file exists
if (!(Test-Path $CertificatePath)) {
    Write-Error "Certificate file not found: $CertificatePath"
    exit 1
}

# Connect to Azure AD
Write-Host "Connecting to Azure AD..."
$context = Get-AzContext
if (!$context) {
    Connect-AzAccount
}

# Get the application
Write-Host "Getting application: $ApplicationId"
$app = Get-AzADApplication -ApplicationId $ApplicationId

if (!$app) {
    Write-Error "Application not found: $ApplicationId"
    exit 1
}

# Read certificate
Write-Host "Reading certificate..."
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$cert.Import($CertificatePath)

# Create key credential
Write-Host "Creating key credential..."
$startDate = [DateTime]::UtcNow
$endDate = $cert.NotAfter
$keyId = [guid]::NewGuid()

$keyCredential = @{
    DisplayName = $DisplayName
    EndDateTime = $endDate
    Key = [System.Convert]::ToBase64String($cert.RawData)
    KeyId = $keyId
    StartDateTime = $startDate
    Type = "AsymmetricX509Cert"
    Usage = "Sign"
}

# Add key credential to application
Write-Host "Adding certificate to application..."
Update-AzADApplication -ApplicationId $ApplicationId -KeyCredential @($keyCredential)

Write-Host "✅ Certificate successfully added to Azure AD app!" -ForegroundColor Green
Write-Host "Certificate Details:" 
Write-Host "  Subject: $($cert.Subject)"
Write-Host "  Thumbprint: $($cert.Thumbprint)"
Write-Host "  Valid From: $($cert.NotBefore)"
Write-Host "  Valid To: $($cert.NotAfter)"
Write-Host "  Key ID: $keyId"

# Grant permissions
Write-Host "`n=== Granting Required Permissions ===" -ForegroundColor Cyan
Write-Host "Important: Your app registration needs the following permissions:"
Write-Host "  - SharePoint: Sites.Manage.All (or specific sites)"
Write-Host "  - Microsoft Graph: Mail.Read (if needed)"
Write-Host "`nPlease configure these manually in Azure Portal if not already done:"
Write-Host "1. Go to https://portal.azure.com"
Write-Host "2. Navigate to Azure AD > App registrations > $ApplicationId"
Write-Host "3. Go to API permissions"
Write-Host "4. Add SharePoint and Graph permissions as needed"
Write-Host "5. Grant admin consent"

Write-Host "`n✅ Azure AD app registration configuration completed!" -ForegroundColor Green
