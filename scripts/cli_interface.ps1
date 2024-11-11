# cli_interface.ps1
# Command-line interface to interact with provided Active Directory export scripts

# Define parameter sets
param (
    [string]$Action,
    [string]$Department = "",
    [string]$Leader = "",
    [switch]$WithCount
)

# Ensure Action parameter is provided
if (-not $Action) {
    Write-Host "Please select an action to execute by typing the corresponding number." -ForegroundColor Yellow
    Write-Host "Available Actions:" -ForegroundColor Cyan
    Write-Host "  1 - Export distinct departments"
    Write-Host "  2 - Export all AD users"
    Write-Host "  3 - Export AD users by department"
    Write-Host "  4 - Export AD users under a specific leader"
    Write-Host "  5 - Export AD users under a leader with count"
    Write-Host "  6 - Export departments that don't match the format"
    $Action = Read-Host "Enter the number of the action you want to perform"
}

# Map action to the corresponding script
switch ($Action) {
    "1" {
        Write-Host "Executing: Export distinct departments" -ForegroundColor Green
        ./export_ad_distinct_departments.ps1
    }
    "2" {
        Write-Host "Executing: Export all AD users" -ForegroundColor Green
        ./export_ad_users.ps1
    }
    "3" {
        if (-not $Department) {
            Write-Host "Please provide a department name using the -Department parameter." -ForegroundColor Yellow
            exit 1
        }
        Write-Host "Executing: Export AD users for department '$Department'" -ForegroundColor Green
        ./export_ad_users_by_department.ps1 -Department $Department
    }
    "4" {
        if (-not $Leader) {
            Write-Host "Please provide a leader name using the -Leader parameter." -ForegroundColor Yellow
            exit 1
        }
        Write-Host "Executing: Export AD users under leader '$Leader'" -ForegroundColor Green
        ./export_ad_users_under_leader.ps1 -Leader $Leader
    }
    "5" {
        if (-not $Leader) {
            Write-Host "Please provide a leader name using the -Leader parameter." -ForegroundColor Yellow
            exit 1
        }
        Write-Host "Executing: Export AD users under leader '$Leader' with count" -ForegroundColor Green
        ./export_ad_users_under_leader_with_count.ps1 -Leader $Leader
    }
    "6" {
        Write-Host "Executing: Export departments that don't match format" -ForegroundColor Green
        ./export_ad_departments_dont_match_format.ps1
    }
    default {
        Write-Host "Invalid Action. Please choose a valid action." -ForegroundColor Red
        exit 1
    }
}

Write-Host "Script execution completed." -ForegroundColor Cyan
