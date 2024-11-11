# This script must be run on a machine with the Active Directory PowerShell module installed.
# Ensure you have the necessary permissions to query Active Directory.

# Import Active Directory module
Import-Module ActiveDirectory

# Function to get a distinct list of departments from Active Directory
function Get-DistinctDepartments {
    # Get all users and select distinct departments
    $AllUsers = Get-ADUser -Filter * -Properties Department | Where-Object { $_.Department -ne $null }
    $DistinctDepartments = $AllUsers | Select-Object -Property Department -Unique | Sort-Object Department
    return $DistinctDepartments
}

# Get the distinct list of departments
$DistinctDepartments = Get-DistinctDepartments

# Export the distinct departments to a CSV file
$DepartmentExportPath = "C:\Exports\DistinctDepartments.csv"
$DistinctDepartments | Export-Csv -Path $DepartmentExportPath -NoTypeInformation -Encoding UTF8

# Notify user of completion
Write-Output "The distinct departments file can be found at $DepartmentExportPath."