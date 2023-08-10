param (
    [Parameter(Mandatory=$true)]
    [string]$ESLFilePath
)

function Test-Privilege {
    try {
        $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        if (-not $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)){
            Write-Host "This script must be run with elevated privileges. Please run as an administrator." -ForegroundColor Red
            exit
            }
    } catch {
        Write-Host "This script must be run with elevated privileges. Please run as an administrator." -ForegroundColor Red
        exit
    }
}

function Get-Domain {
    $domainInfo = Get-WmiObject Win32_ComputerSystem
    if ($domainInfo.PartOfDomain) {
        return $domainInfo.Domain
    } else {
        Write-Host "Machine is not joined to a domain.  Please join the system to a domain and try again" -ForegroundColor Red
        return $domainInfo.Workgroup
    }
}

function Get-Inventory ($appNames){
    $keys = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    
    foreach ($appName in $appNames) {
        $appFound = $false
        foreach ($key in $keys) {
            try {
                Get-ChildItem $key | ForEach-Object {
                    $app = Get-ItemProperty -Path $_.PSPath
                    if ($app.DisplayName -like $appName) {
                        Write-Host "Application found:" $app.DisplayName -ForegroundColor Green
                        $appFound = $true
                        break
                    }
                }
            } catch {
                Write-Host "An error occurred while checking the registry: $_" -ForegroundColor Red
                exit
            }
            
            if ($appFound) {
                break
            }
        }
    
        if (-not $appFound) {
            Write-Host "Application not found:" $appName -ForegroundColor Red
        }
    }
}

function Install-ExcelModule {
    $moduleName = "ImportExcel"
    Write-Host "Installing ImportExcel Module"
    if (Get-Module -ListAvailable -Name $moduleName) {
        Write-Host "Module already intalled.  Skipping..."
    } else {
        try {
            Install-Module -Name $moduleName -Scope CurrentUser -Force
        } catch {
            Write-Host "An error occurred while installing the module: $_" -ForegroundColor Red
            exit
        }
    }
}

function Open-ESL ([string]$eslPath) {
    Import-Module ImportExcel

    #Import Excel content and skip the first two rows
    try {
        $excelData = Import-Excel -Path $eslPath -StartRow 3
        $json = $excelData | ConvertTo-Json
        return $json
    } catch {
        Write-Host "An error occurred while opening the ESL file.  Check the path and try again: $_" -ForegroundColor Red
        exit
    }
}

function Get-ESLDomain ($json, $hostName) {
    $jsonData = $json | ConvertFrom-Json
    try {
        $selectedRow = $jsonData | Where-Object { $_.Hostname -eq $hostName } | Select-Object -First 1
        if ($null -eq $selectedRow) {
            $message = @"
The ESL File does not contain the Hostname of this machine.  
This means either the machine was not named properly, or there is a problem with the ESL file.
Please validate the hostname and ESL file. 
"@
            Write-Host $message -ForegroundColor Red
            exit
        }
    } catch {
        Write-Host "ESL File exception occurred.  Please check the ESL file and try again: $_" -ForegroundColor Red
        exit
    }
    return $selectedRow.Domain
}

function Test-Attributes ($eslFilePath) {
    $hostName = hostname
    Install-ExcelModule
    $jsonOutput = Open-ESL $eslFilePath
    $eslDomain = Get-ESLDomain $jsonOutput $hostName
    $machineDomain = Get-Domain
    if ($eslDomain -eq $machineDomain) {
        Write-Host "Hostname $hostName matches ESL File" -ForegroundColor Green
        Write-Host "Domain $machineDomain matches ESL File" -ForegroundColor Green
    } else {
        Write-Host "Domain does not match the ESL File" -ForegroundColor Red
        Write-Host "Machine has joined $machineDomain" -ForegroundColor Red
        Write-Host "ESL States the domain should be $eslDomain" -ForegroundColor Red
        exit
    }
    Get-Inventory "Trend Micro*", "Cortex*", "OpsRamp*"

}

Test-Privilege
Install-ExcelModule
Test-Attributes $ESLFilePath


