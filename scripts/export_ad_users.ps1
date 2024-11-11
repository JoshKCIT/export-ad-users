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

# Use Get-ADUser to filter and retrieve only 250 active users with specified properties
$Users = Get-ADUser -Filter {Enabled -eq $true} -Properties $Properties -ResultSetSize 250 |
    Select-Object -Property GivenName, Surname, Mail, Title, Department, MobilePhone, 
                  @{Name='Manager'; Expression={(Get-ADUser -Identity $_.Manager -Properties DisplayName).DisplayName}} |
    Sort-Object -Property Surname, GivenName

# Export the data to a CSV file
$ExportPath = "C:\Exports\ActiveDirectoryUsers.csv"
$Users | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

# Notify user of completion
Write-Output "Export completed successfully. The file can be found at $ExportPath."
