# Export AD Users

This PowerShell script exports details of up to 250 active users from Active Directory, including the following fields:

- First Name
- Last Name
- Email Address
- Job Title
- Department
- Mobile Number
- Manager

## Prerequisites

- Active Directory PowerShell Module
- Sufficient permissions to query the domain

## Usage

Run the script on a machine that is part of the Active Directory domain:

```powershell
.\export_ad_users.ps1
