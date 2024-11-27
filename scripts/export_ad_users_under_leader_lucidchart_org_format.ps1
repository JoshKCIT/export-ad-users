# Optimized PowerShell script to export Active Directory hierarchy from a top-level user email down through the reporting structure
# This script requires the ActiveDirectory PowerShell module to be installed and active
# Import required modules
Import-Module ActiveDirectory -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop
# Global hashtable to cache user data for reuse and minimize AD queries
$UserCache = @{}
# List all available domain controllers and prompt the user to select one
$DomainControllers = Get-ADDomainController -Filter *
for ($i = 0; $i -lt $DomainControllers.Count; $i++) {
    Write-Output "${i}: $($DomainControllers[$i].Name)"
}
$SelectedIndex = Read-Host -Prompt "Please select a domain controller by number"
if ($SelectedIndex -notmatch '^[0-9]+$' -or $SelectedIndex -ge $DomainControllers.Count) {
    Write-Error "Invalid selection. Please run the script again and select a valid number."
    exit
}
$DomainController = $DomainControllers[$SelectedIndex].Name
# Function to retrieve and cache user information
function Get-User {
    param([string]$UserDN)
    if ($UserCache.ContainsKey($UserDN)) {
        return $UserCache[$UserDN]
    }
    # Fetch user if not already cached
    $User = Get-ADUser -Identity $UserDN -Properties DisplayName,Title,Department,Mail,TelephoneNumber,DirectReports,EmployeeID,Manager
    if ($User) {
        $UserCache[$UserDN] = $User
    }
    return $User
}
# Function to generate the next Employee ID dynamically
function Get-NextEmployeeID {
    param ([int]$Counter)
    return "{0:D7}" -f $Counter
}
# Function to retrieve direct reports iteratively
function Get-OrgStructure {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserDN, # The distinguished name of the top-level user
        [ref]$Results # Use ref to pass the array by reference
    )
    
    # Initialize a queue to process users iteratively
    $queue = [System.Collections.Generic.Queue[string]]::new()
    $queue.Enqueue($UserDN)
    $EmployeeIDCounter = 1
    $EmployeeIDMap = @{}
    while ($queue.Count -gt 0) {
        $currentUserDN = $queue.Dequeue()
        $User = Get-User -UserDN $currentUserDN
        if ($User -eq $null) {
            continue
        }
        # Assign a sequential EmployeeID, starting with "0000001" for the top user
        $EmployeeID = Get-NextEmployeeID -Counter $EmployeeIDCounter
        $EmployeeIDMap[$currentUserDN] = $EmployeeID
        $EmployeeIDCounter++
        
        $ManagerName = ''
        $ManagerEmployeeID = ''
        if ($User.Manager) {
            # Ensure the manager is in the cache
            if (-not $UserCache.ContainsKey($User.Manager)) {
                $UserCache[$User.Manager] = Get-ADUser -Identity $User.Manager -Properties DisplayName,EmployeeID
            }
            # Retrieve the manager information from the cache
            if ($UserCache.ContainsKey($User.Manager)) {
                $ManagerObj = $UserCache[$User.Manager]
                $ManagerName = $ManagerObj.DisplayName
                $ManagerEmployeeID = $EmployeeIDMap[$User.Manager]
            }
        }
        $Results.Value += [PSCustomObject]@{
            Name               = $User.DisplayName
            EmployeeID         = $EmployeeID
            Manager            = $ManagerName
            Manager_EmployeeID = $ManagerEmployeeID
            Title              = $User.Title
            Department         = $User.Department
            Email              = $User.Mail
            Phone              = $User.TelephoneNumber
        }
        # Batch get direct reports and add them to the queue
        if ($User.DirectReports.Count -gt 0) {
            foreach ($directReport in $User.DirectReports) {
                if (-not $UserCache.ContainsKey($directReport)) {
                    $queue.Enqueue($directReport)
                }
            }
        }
    }
}
# Main script
# Prompt the user to enter the top-level user's email address
$TopUserEmail = Read-Host -Prompt "Please enter the email address of the top individual"
# Ensure the output path exists
$OutputPath = "C:\Exports\Lucidchart_Format_AD_Org_Chart.csv"
if (-not (Test-Path -Path (Split-Path -Path $OutputPath -Parent))) {
    New-Item -Path (Split-Path -Path $OutputPath -Parent) -ItemType Directory -Force
}
# Get the top individual by their email address
$TopUser = Get-ADUser -Filter {Mail -eq $TopUserEmail} -Properties DistinguishedName,DisplayName,Title,Department,Mail,TelephoneNumber,DirectReports,EmployeeID,Manager
if ($TopUser) {
    $UserCache[$TopUser.DistinguishedName] = $TopUser
    $Results = @()
    Get-OrgStructure -UserDN $TopUser.DistinguishedName -Results ([ref]$Results)
    
    # Export results to CSV in one go for better performance
    $Results | Export-Csv -Path $OutputPath -NoTypeInformation -Force
    Write-Output "Export complete: $OutputPath"
} else {
    Write-Error "Could not find the user with email address: $TopUserEmail"
}
