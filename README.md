# RFL.Microsoft.Intune
Automatic create an Intune Documentation to simplify the life of admins and consultants.

# Intune Documentation
Automatic create an Intune Documentation to simplify the life of admins and consultants.

# Usage
ExportIntuneData.ps1 - On a device with internet connectivity, open a PowerShell 5.1 session (not PowerShell core), it will connect to the Intune and Azure Services and export the data to a word or html format. It uses the PScribo PowerShell module (https://github.com/iainbrighton/PScribo) and MSAL.PS PowerShell module (https://github.com/AzureAD/MSAL.PS). It can be run from any Windows Device (Workstation, Server). As it uses external PowerShell module, it is recommended not to run from a Domain Controller.

# Pre-Requisites
To use the script, an Azure AD Application is required with the correct permissions
1. create a Azure AD Application - https://learn.microsoft.com/en-us/graph/toolkit/get-started/add-aad-app-registration
2. error AADSTS7000218 - fix: Navigate to Azure AD -> App registration -> manifest and change ""allowPublicClient": null," or ""allowPublicClient": false," to ""allowPublicClient": true,". 
3. under API Permissions 
4. Click add a permission and select migrosoft graph, click Application permissions and add the following permissions
			- DeviceManagementApps.Read.All
			- DeviceManagementConfiguration.Read.All
 			- DeviceManagementManagedDevices.PrivilegedOperations.All
 			- DeviceManagementManagedDevices.Read.All
 			- DeviceManagementRBAC.Read.All
 			- DeviceManagementServiceConfig.Read.All
 			- Directory.Read.All
 			- Group.Read.All
 			- User.Read.All
 			- CloudPC.Read.All
 			- RoleManagement.Read.CloudPC
 			- Organization.Read.All
 			- UserAuthenticationMethod.Read.All
      - AuditLog.Read.All
5. Click add a permission and select Intune, click Application permissions and add the following permissions
 			- get_data_warehouse
			- get_device_compliance
6. under the API Permissions click Grant consent for the tenant 
7. under Certificates & secrets create a new client secret and copy the value to ClientSecret


# Examples
Example01: Exports the Data in a word format using the ClientId xxx, CLientSecret xxx with TenantID tenant.com and will save the file to c:\temp folder

**.\ExportIntuneData.ps1 -ClientId 'xxx' -TenantId 'tenant.com' -ClientSecret 'xxx' -OutputFolderPath "c:\temp"**

Example 02: Exports the Data in a html format using the ClientId xxx, CLientSecret xxx with TenantID tenant.com and will save the file to c:\temp folder

**.\ExportIntuneData.ps1 -BetaAPI -OutputFormat @('HTML') -ClientId 'xxx' -TenantId 'tenant.com' -ClientSecret 'xxx' -OutputFolderPath "c:\temp"**

Example 03: Exports the Data in a word and html format using the ClientId xxx, CLientSecret xxx with TenantID tenant.com and will save the file to c:\temp folder and Add the company details to the header

**.\ExportIntuneData.ps1 -BetaAPI -OutputFormat @('Word', 'HTML') -ClientId 'xxx' -TenantId 'tenant.com' -ClientSecret 'xxx' -OutputFolderPath "c:\temp" -CompanyName 'RFL Systems' -CompanyWeb 'www.rflsystems.co.uk' -CompanyEmail 'team@rflsystems.co.uk'**

# Documentation
Access our Wiki at https://github.com/dotraphael/RFL.Microsoft.Intune/wiki

# Issues and Support
Access our Issues at https://github.com/dotraphael/RFL.Microsoft.Intune/issues
