# This script must be run on a machine with the Active Directory PowerShell module installed.
# Ensure you have the necessary permissions to query Active Directory.

# Import Active Directory module
Import-Module ActiveDirectory

# Define the properties to be retrieved
$Properties = @(
    "GivenName",     # First Name
    "Surname",       # Last Name
    "Mail",          # Email Address
    "Title",         # Job Title
    "Department",    # Department
    "MobilePhone",   # Mobile Number
    "Manager"        # Manager (Distinguished Name)
)

# Get the manager user object for the specified email address
$ManagerEmail = "LeadersName@CompanyEmail.com"
$Manager = Get-ADUser -Filter {Mail -eq $ManagerEmail} -Properties DistinguishedName

# Recursive function to get all direct reports
function Get-AllReports {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ManagerDN
    )
    
    # Get direct reports of the current manager
    $DirectReports = Get-ADUser -Filter {Manager -eq $ManagerDN} -Properties $Properties
    
    # Create an array to hold all reports
    $AllReports = @()

    foreach ($Report in $DirectReports) {
        # Add current report to the list
        $AllReports += $Report

        # Recursively get reports of this report
        $AllReports += Get-AllReports -ManagerDN $Report.DistinguishedName
    }
    
    return $AllReports
}

# Get all employees under the specified manager recursively
$AllEmployees = Get-AllReports -ManagerDN $Manager.DistinguishedName |
    Select-Object -Property @{Name='EmployeeName'; Expression={$_.GivenName + ' ' + $_.Surname}}, Mail, Title, Department, MobilePhone, 
                  @{Name='Manager'; Expression={(Get-ADUser -Identity $_.Manager -Properties DisplayName).DisplayName}} |
    Sort-Object -Property EmployeeName

# Export the data to a CSV file
$ExportPath = "C:\Exports\ActiveDirectoryEmployeesUnderManager.csv"
$AllEmployees | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

# Notify user of completion
Write-Output "Export completed successfully. The file can be found at $ExportPath."
