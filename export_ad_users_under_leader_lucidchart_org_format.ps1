# PowerShell script to export Active Directory hierarchy from a top-level user email down through the reporting structure
# This script requires the ActiveDirectory PowerShell module to be installed and active
# Check if the Active Directory module is installed, if not, install it
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Output "ActiveDirectory module not found. Installing..."
    Install-WindowsFeature -Name RSAT-AD-PowerShell -IncludeAllSubFeature -IncludeManagementTools -ErrorAction Stop
}
# Import the Active Directory Module
Import-Module ActiveDirectory
# Check if the ImportExcel module is installed, if not, install it
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Output "ImportExcel module not found. Installing..."
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
# Import the Excel Module
Import-Module -Name ImportExcel -Scope Local
# Global variable to track EmployeeID increment
$Global:CurrentEmployeeID = 1
# Function to generate a unique EmployeeID
function Generate-SequentialEmployeeID {
    $newID = "{0:D7}" -f $Global:CurrentEmployeeID
    $Global:CurrentEmployeeID++
    return $newID
}
# Function to retrieve direct reports recursively
function Get-OrgStructure {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserDN, # The distinguished name of the top-level user
        [Parameter(Mandatory=$true)]
        [ref]$Results
    )
    
    # Get user object and add to the results list
    $User = Get-ADUser -Identity $UserDN -Properties DisplayName,Title,Department,Mail,TelephoneNumber,DirectReports,EmployeeID,Manager
    
    # Assign a sequential EmployeeID, starting with "0000001" for the top user
    $EmployeeID = Generate-SequentialEmployeeID
    
    $Manager = if ($User.Manager) {
        $ManagerObj = Get-ADUser -Identity $User.Manager -Properties DisplayName, EmployeeID
        [string]$ManagerName = $ManagerObj.DisplayName
        [string]$ManagerEmployeeID = if ($ManagerObj.EmployeeID) { $ManagerObj.EmployeeID } else { '' }
    } else {
        $ManagerName = ''
        $ManagerEmployeeID = ''
    }
    
    $Results.Value.Add([PSCustomObject]@{
        Name = $User.DisplayName
        EmployeeID = $EmployeeID
        Manager = $ManagerName
        Manager_EmployeeID = $ManagerEmployeeID
        Title = $User.Title
        Department = $User.Department
        Email = $User.Mail
        Phone = $User.TelephoneNumber
    })
    # Get direct reports and iterate over them
    if ($User.DirectReports) {
        foreach ($report in $User.DirectReports) {
            Get-OrgStructure -UserDN $report -Results ([ref]$Results.Value)
        }
    }
}
# Main script
# Get the top individual by their email address
$TopUserEmail = Read-Host -Prompt "Please enter the email address of the top individual"
$TopUser = Get-ADUser -Filter {Mail -eq $TopUserEmail} -Properties DistinguishedName
if ($TopUser) {
    $Results = New-Object System.Collections.Generic.List[PSCustomObject]
    Get-OrgStructure -UserDN $TopUser.DistinguishedName -Results ([ref]$Results)
    # Export results to Excel Workbook
    $OutputPath = "C:\Exports\Lucidchart_Format_AD_Org_Chart.xlsx"
    $Results | Export-Excel -Path $OutputPath -WorksheetName "OrgChart" -AutoSize
    Write-Output "Export complete: $OutputPath"
    # Re-import the Excel file and update Manager_EmployeeID based on EmployeeID of the manager
    $ImportedResults = Import-Excel -Path $OutputPath -WorksheetName "OrgChart"
    foreach ($entry in $ImportedResults) {
        if ($entry.Manager -ne '') {
            $managerEntry = $ImportedResults | Where-Object { $_.Name -eq $entry.Manager }
            if ($managerEntry) {
                $entry.Manager_EmployeeID = $managerEntry.EmployeeID
            }
        }
    }
    
    # Export the updated results back to the Excel file
    $ImportedResults | Export-Excel -Path $OutputPath -WorksheetName "OrgChart" -AutoSize -ClearSheet
    Write-Output "Manager_EmployeeID fields updated: $OutputPath"
} else {
    Write-Error "Could not find the user with email address: $TopUserEmail"
}
