# This script must be run on a machine with the Active Directory PowerShell module installed.
# Ensure you have the necessary permissions to query Active Directory.

# Import Active Directory module
Import-Module ActiveDirectory

# Function to get all users and filter those whose department length is not equal to eight
function Get-InvalidDepartmentLengthEntries {
    # Get all users with a non-null Department attribute
    $AllUsers = Get-ADUser -Filter * -Properties Department | Where-Object { $_.Department -ne $null }
    
    # Filter users where the department length is not equal to eight characters
    $InvalidDepartments = $AllUsers | Where-Object {
        $_.Department.Length -ne 8
    }

    return $InvalidDepartments
}

# Get the list of users with invalid department length entries
$UsersWithInvalidDepartmentLength = Get-InvalidDepartmentLengthEntries

# Output the invalid department length entries for review
foreach ($User in $UsersWithInvalidDepartmentLength) {
    Write-Output "User: $($User.SamAccountName), Department: '$($User.Department)', Length: $($User.Department.Length)"
}

# Export invalid department length entries to a CSV file for further review
$ExportPath = "C:\Exports\InvalidDepartmentLength.csv"
$UsersWithInvalidDepartmentLength | Select-Object SamAccountName, Department | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

# Notify user of completion
Write-Output "The list of users with invalid department length entries has been exported to $ExportPath."
