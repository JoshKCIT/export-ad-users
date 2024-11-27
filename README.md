# Active Directory Export PowerShell Scripts

This repository contains a collection of PowerShell scripts designed to interact with Microsoft Active Directory (AD) to extract user and department data in various formats. These scripts are particularly useful for system administrators or IT professionals who need to export and analyze AD data for auditing, reporting, or administrative tasks.

## Scripts Overview

1. **export_ad_users.ps1**  
   This script exports all user information from Active Directory, including user names, email addresses, department, and other relevant properties. It uses the `Get-ADUser` cmdlet from the Active Directory PowerShell module to query user data. It is useful for getting a complete overview of all users currently registered in the AD.

2. **export_ad_users_by_department.ps1**  
   This script exports user information based on departments within the organization. It uses the `Get-ADUser` cmdlet with filtering based on the department attribute to allow you to filter AD user data by department, making it convenient for department-level reporting or analysis.

3. **export_ad_users_under_leader.ps1**  
   This script extracts users reporting to a specific leader or manager within the organization. It utilizes the `Get-ADUser` cmdlet to find users based on the `Manager` attribute. You can specify the manager's name or ID, and the script will return all direct reports. This is particularly helpful for understanding reporting hierarchies and gathering information on all direct reports for a specific manager.

4. **export_ad_users_under_leader_with_count.ps1**  
   Similar to the `export_ad_users_under_leader.ps1`, this script also uses the `Get-ADUser` cmdlet and provides the number of users reporting to a specific leader. The count feature makes it easy to generate summary-level information for leadership reports. This script is especially useful for quickly determining team sizes.

5. **export_ad_departments_dont_match_format.ps1**  
   This script identifies departments within AD that do not match a specified format. It uses the `Get-ADUser` cmdlet to retrieve department information and then checks for discrepancies. It is useful for cleaning up department records, ensuring consistency, and aligning department names with predefined naming conventions. The script outputs both the department names that do not match and the users associated with them.

6. **export_ad_distinct_departments.ps1**  
   This script exports a list of distinct departments from Active Directory. It uses the `Get-ADUser` cmdlet to extract department information and compiles a list of unique department names. It helps administrators get an overview of all departments within the AD, which can be useful for auditing and verifying organizational structures. The output includes department names and the number of users within each department.

7. **export_ad_users_under_leader_lucidchart_org_format.ps1**  
   This script extracts users under a leader and structures the data in a format that is compatible with Lucidchart's organizational chart tools. It utilizes the `Get-ADUser` cmdlet to get the reporting structure and provides data that can be easily imported into Lucidchart to create an organizational chart.

8. **cli_interface.ps1**  
   A command-line interface script designed to streamline the use of the other PowerShell scripts in this repository. This provides an interactive way to execute specific export tasks by selecting options, improving ease of use. The CLI guides users through available options and provides prompts for required parameters, reducing the need for memorizing specific script commands.

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

   To use the CLI interface script for a more interactive experience:
   
   ```powershell
   .\cli_interface.ps1
   ```

## Typical Use Cases

- **Auditing Active Directory Users**: Get complete or filtered lists of users based on different criteria such as departments or managerial hierarchy.
- **Data Cleanup and Consistency Checks**: Identify discrepancies in department naming formats and correct them to maintain data integrity.
- **Reporting**: Generate managerial or department-level reports for organizational insights.
- **Visualizing Organizational Structure**: Extract data for use in tools like Lucidchart to visualize reporting hierarchies.
- **Team Size Analysis**: Use scripts like `export_ad_users_under_leader_with_count.ps1` to determine team sizes for workforce planning.

## Contributions

Feel free to open issues or submit pull requests if you wish to contribute or improve these scripts. Contributions that enhance the functionality or add new export formats are always welcome.

## License

This project is licensed under the MIT License. See the `LICENSE` file for more details.

