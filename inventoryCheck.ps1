# Function to check if running with elevated privileges
function IsAdmin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

# Check if running with elevated privileges
if (-not (IsAdmin)) {
    Write-Host "This script must be run with elevated privileges. Please run as an administrator."
    exit
}

# List of applications to search for
$appNames = "Your-Application-Name1", "Your-Application-Name2", "Your-Application-Name3"

# Check in both the 32-bit and 64-bit registry view
$keys = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

# Iterate through the application names
foreach ($appName in $appNames) {
    # Flag to determine if the application is found
    $appFound = $false
    
    # Iterate through the keys and look for the application
    foreach ($key in $keys) {
        try {
            Get-ChildItem $key | ForEach-Object {
                $app = Get-ItemProperty -Path $_.PSPath
                if ($app.DisplayName -eq $appName) {
                    Write-Host "Application found:" $app.DisplayName
                    $appFound = $true
                    break
                }
            }
        } catch {
            Write-Host "An error occurred while checking the registry: $_"
            exit
        }
        
        if ($appFound) {
            break
        }
    }

    if (-not $appFound) {
        Write-Host "Application not found:" $appName
    }
}

# Get the domain information from the computer
$domainInfo = Get-WmiObject Win32_ComputerSystem

# Check if the computer is part of a domain
if ($domainInfo.PartOfDomain) {
    Write-Host "This computer is joined to a domain."
    Write-Host "Domain Name:" $domainInfo.Domain
} else {
    Write-Host "This computer is part of a workgroup."
    Write-Host "Workgroup Name:" $domainInfo.Workgroup
}
# Define the metadata endpoint and API version
$metadataUrl = "http://169.254.169.254/metadata/instance?api-version=2021-01-01"
$headers = @{"Metadata"="true"}

# Function to perform a GET request to the metadata endpoint
function Get-Metadata {
    param (
        [string]$url
    )

    try {
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
        return $response
    } catch {
        Write-Host "An error occurred while retrieving metadata: $_"
        exit
    }
}

# Get the complete metadata
$metadata = Get-Metadata -url $metadataUrl

# Print the hostname
Write-Host "Hostname:" $metadata.computer.vmName

# Get the tags (if available)
$tagsUrl = "http://169.254.169.254/metadata/instance/compute/tags?api-version=2021-01-01"
$tags = Get-Metadata -url $tagsUrl

# Print the tags
Write-Host "Tags:" $tags
