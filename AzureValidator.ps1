param (
    [Parameter(Mandatory=$true)]
    [string]$ESLFilePath
)

function Test-AzureCLI {
    if (Get-Command az -ErrorAction SilentlyContinue) {
        Write-Output "Azure CLI is installed."
    } else {
        Write-Output "Azure CLI is not installed.  Please install the Azure CLI and try again. See https://aka.ms/installazurecliwindows" -ForegroundColor Red
        exit
    }
}

function Login-Az {
    $isLoggedIn = az account show 2>$null
    if (-not $isLoggedIn) {
        az login
    }
}

function Set-Subscription ($subscriptionName) {
    try {
        az account set --subscription $subscriptionName
    } catch {
        Write-Host "An error occurred while communication with Azure: $_" -ForegroundColor Red
        exit
    }
}

function Get-Tags ($resourceGroup, $hostName) {
    try {
        $tags = az vm show --resource-group $resourceGroup, --name $hostName --query "tags"
    } catch {
        Write-Host "An error occurred while querying Azure: $_" -ForegroundColor Red
        exit
    }
    return $tags
    
}
function Get-Backups ($resourceGroup, $hostName){
    $backupExists = $false
    try {
        $vaults = az backup vault list --resource-group $resourceGroup --query "[].name" -o tsv
        foreach ($vaultName in $vaults) {
            $backupItems = az backup item list --vault-name $vaultName --resource-group $resourceGroup --backup-management-type AzureIaasVM --query "[?properties.friendlyName == '$hostName']" -o tsv
            if (-not [string]::IsNullOrEmpty($backupItems)) {
                $backupExists = $true
                break
            }
        }
    } catch {
        Write-Host "An error occurred while querying Azure: $_" -ForegroundColor Red
        exit
    }
    if ($backupExists) {
        Write-Host "A backup policy is attached to the VM." -ForegroundColor Green
    } else {
        Write-Host "No backup policy is attached to the VM." -ForegroundColor Red
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
    #Tag Variables.  If tagging policy changes, this section will need to be updated.
    $columnOne = "Hostname"
    $columnTwo = "Resource_Description"
    $columnThree = "Environment"
    $columnFour = "Number"
    $columnFive = "Application_Name"
    $columnSix = "Entity_Name"
    $columnSeven = "Entity"
    $columnEight = "Organization"
    $columnNine = "Application_Owner"
    $columnTen = "Application_Technical_Contact"
    $columnEleven = "Distribution_List"
    $columnTwelve = "Launch_Date"
    $columnThirteen = "Backup_Schedule"
    $columnFourteen = "Backup_Retention"
    $columnFifteen = "Provisioning"
    $columnSixteen = "Maintenance_Window"
    $columnSeventeen = "OS_Version"
    $columnEighteen = "SubscriptionName"
    $columnNineteen = "Resource_Group"
    $columnTwenty = "Resource_Group for Recovery Service Vault*"

    #Import Excel content and skip the first two rows
    try {
        $excelData = Import-Excel -Path $eslPath -StartRow 3 | Select-Object $columnOne, $columnTwo, $columnThree, $columnFour, $columnFive, $columnSix, $columnSeven, $columnEight, $columnNine, $columnTen, $columnEleven, $columnTwelve, $columnThirteen, $columnFourteen, $columnFifteen, $columnSixteen, $columnSeventeen, $columnEighteen, $columnNineteen, $columnTwenty
        $json = $excelData | ConvertTo-Json
        return $json
    } catch {
        Write-Host "An error occurred while opening the ESL file.  Check the path and try again: $_" -ForegroundColor Red
        exit
    }
}

function Get-VMData ($json) {
    $jsonData = $json | ConvertFrom-Json
    try {
        foreach ($row in $jsonData) {
            $hostName = $row.Hostname
            $subscriptionName = $row.SubscriptionName
            $resourceGroup = $row.Resource_Group
            #$backupResourceGroup = $row."Resource_Group for Recovery Service Vault`n(Azure Backup)"
            Write-Host "Now Processing Host: $hostName"
            Set-Subscription $subscriptionName
            Get-Backups $resourceGroup $hostName
            $tags = Get-Tags $resourceGroup $hostName | ConvertFrom-Json
            if ($tags -eq $null) {
                Write-Host "An error occurred while scanning the host.  This usually means the resource group or subscription name in the ESL does not match the Azure VM or the Azure VM doesn't exist." -ForegroundColor Red
                continue
            }
            foreach ($cell in $row.PSObject.Properties) {
                $cellName = $cell.Name
                #Write-Host $cellName
                $tagValue = $tags.$cellName
                #Write-Host $tagValue
                if ($cellName -ne "SubscriptionName" -and $cellName -ne "Resource_Group"-and $cellName -ne "Launch_Date" -and $cellName -ne "Resource_Group for Recovery Service Vault`n(Azure Backup)") {
                    #Normalize the string coming from the ESL
                    $cellValue = $cell.Value.Replace(" / ", ",")
                    $cellValue = $cellValue.Replace("backup for ", "")
                    #Write-Host $cellValue
                    if ($tags | Get-Member -Name $cellName) {    
                        if ($cellValue -eq $tagValue){
                            Write-Host "Tags Match- $cellName : $cellValue" -ForegroundColor Green
                        } else {
                            Write-Host "$cellName Tag Does Not Match-" -ForegroundColor Red
                            Write-Host "  ESL Value: $cellValue" -ForegroundColor Red
                            Write-Host "  Discovered Value: $tagValue" -ForegroundColor Red
                        }
                    } else {
                        Write-Host "$cellName tag does not exist" -ForegroundColor Red
                    }
                } elseif ($cellName -eq "Launch_Date") {
                    if ($tags | Get-Member -Name $cellName) {
                        Write-Host "Launch_Date Tag Exists" -ForegroundColor Green
                    } else {
                        Write-Host "Launch_Date Tag Does Not Exist" -ForegroundColor Red
                    }
                }
            }

        }
    } catch {
        Write-Host "ESL File exception occurred.  Please check the ESL file and try again: $_" -ForegroundColor Red
        exit
    }
}



Test-AzureCLI
Login-Az
$eslData = Open-ESL "$ESLFilePath"
Get-VMData $eslData
