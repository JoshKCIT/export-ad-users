# Active Directory Export PowerShell Scripts

This repository contains a collection of PowerShell scripts designed to interact with Microsoft Active Directory (AD) to extract user and department data in various formats. These scripts are particularly useful for system administrators or IT professionals who need to export and analyze AD data for auditing, reporting, or administrative tasks.

## Scripts Overview

1. **export_ad_users.ps1**  
   This script exports all user information from Active Directory. It is useful for getting a complete overview of all users currently registered in the AD.

2. **export_ad_users_by_department.ps1**  
   This script exports user information based on departments within the organization. It allows you to filter AD user data by department, making it convenient for department-level reporting or analysis.

3. **export_ad_users_under_leader.ps1**  
   This script extracts users reporting to a specific leader or manager within the organization. This is particularly helpful for understanding reporting hierarchies and gathering information on all direct reports for a specific manager.

4. **export_ad_users_under_leader_with_count.ps1**  
   Similar to the `export_ad_users_under_leader.ps1`, this script also provides the number of users reporting to a specific leader. The count feature makes it easy to generate summary-level information for leadership reports.

5. **export_ad_departments_dont_match_format.ps1**  
   This script identifies departments within AD that do not match a specified format. It is useful for cleaning up department records, ensuring consistency, and aligning department names with predefined naming conventions.

6. **export_ad_distinct_departments.ps1**  
   This script exports a list of distinct departments from Active Directory. It helps administrators get an overview of all departments within the AD, which can be useful for auditing and verifying organizational structures.

## Requirements

- PowerShell 5.1 or later.
- Appropriate permissions to query Active Directory.
- Active Directory PowerShell module (`ActiveDirectory`).

## Usage

1. Clone the repository to your local environment.
2. Open PowerShell with administrator privileges.
3. Execute the scripts based on your reporting needs.

   For example, to export all users by department, use:
   
   ```powershell
   .\export_ad_users_by_department.ps1
   ```

## Typical Use Cases

- **Auditing Active Directory Users**: Get complete or filtered lists of users based on different criteria such as departments or managerial hierarchy.
- **Data Cleanup and Consistency Checks**: Identify discrepancies in department naming formats and correct them.
- **Reporting**: Generate managerial or department-level reports for organizational insights.

## Contributions

Feel free to open issues or submit pull requests if you wish to contribute or improve these scripts. Contributions that enhance the functionality or add new export formats are always welcome.

## License

This project is licensed under the MIT License. See the `LICENSE` file for more details.

