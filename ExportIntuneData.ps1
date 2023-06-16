<#
    .SYSNOPSIS
        Export Microsoft Intune data to a word or html format

    .DESCRIPTION
        Export Microsoft Intune data to a word or html format

    .PARAMETER OutputFormat
        Format to export. Possible options are Word and HTML

    .PARAMETER BetaAPI
        Usage of Beta Graph API - https://learn.microsoft.com/en-us/graph/use-the-api#version

    .PARAMETER ClientId
        Graph API ClientId

    .PARAMETER TenantId
        Graph API TenantId

    .PARAMETER ClientSecret
        Graph API ClientSecret

    .PARAMETER CompanyName
        Company Name to be added onto the report's header

    .PARAMETER CompanyWeb
        Company URL to be added onto the report's header

    .PARAMETER CompanyEmail
        Company E-mail to be added onto the report's header

    .PARAMETER SectionTenantAdmin
        Export the tenant admin section'

    .PARAMETER SectionEnrollment
        Export the enrollment section'

    .PARAMETER SectionDevices
        Export the devices section

    .PARAMETER SectionUsers
        Export the users section

    .PARAMETER SectionGroups
        Export the groups section

    .PARAMETER SectionApps
        Export the apps section

    .PARAMETER SectionEndpointSecurity
        Export the Endpoint Security section

    .PARAMETER SectionPolicies
        Export the Policies section

    .NOTES
        Name: ExportIntuneData.ps1
        Author: Raphael Perez
        DateCreated: 17 May 2023 (v0.1)
        LatestUpdate: 16 June 2023 (v0.3)
        Website: http://www.endpointmanagers.com
        WebSite: https://github.com/dotraphael/RFL.Microsoft.Intune
        Twitter: @dotraphael
        Update: 24 May 2023 (v0.2)
            - Added filter to disable some sections from being exported
            - Added more platformsList
            - Added logging to Graph Connection Info
            - Added Filters section (apps and devices)
            - Added Apps Category Section
        Update: 16 June 2023 (v0.3)
            - Added more platformsList
            - Added templateIDList Enumeration
            - Added Ebooks Section
            - Added Ebooks category Section
            - Added Disk Encryption section
            - Added Firewall section (#Todo: ConfigMgr firewall policies)

        Create Azure AD Application
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


    .LINK
        http://www.endpointmanagers.com
        http://www.rflsystems.co.uk
        https://github.com/dotraphael/RFL.Microsoft.Intune

    .EXAMPLE
        .\ExportIntuneData.ps1 -ClientId 'xxx' -TenantId 'tenant.com' -ClientSecret 'xxx' -OutputFolderPath "c:\temp"
        .\ExportIntuneData.ps1 -BetaAPI -OutputFormat @('HTML') -ClientId 'xxx' -TenantId 'tenant.com' -ClientSecret 'xxx' -OutputFolderPath "c:\temp"
        .\ExportIntuneData.ps1 -BetaAPI -OutputFormat @('Word', 'HTML') -ClientId 'xxx' -TenantId 'tenant.com' -ClientSecret 'xxx' -OutputFolderPath "c:\temp" -CompanyName 'RFL Systems' -CompanyWeb 'www.rflsystems.co.uk' -CompanyEmail 'team@rflsystems.co.uk'
        .\ExportIntuneData.ps1 -BetaAPI -OutputFormat @('Word', 'HTML') -ClientId 'xxx' -TenantId 'tenant.com' -ClientSecret 'xxx' -OutputFolderPath "c:\temp" -CompanyName 'RFL Systems' -CompanyWeb 'www.rflsystems.co.uk' -CompanyEmail 'team@rflsystems.co.uk' -SectionTenantAdmin $false -SectionEnrollment $false -SectionDevices $false -SectionUsers $false -SectionGroups $false -SectionEndpointSecurity $false -SectionPolicies $false
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = 'Please provide the format you wish to export the report to')]
    [ValidateNotNullOrEmpty()]
    [ValidateSet('Word', 'HTML')]
    [string[]] $OutputFormat = @('Word'),

    [Parameter(Mandatory = $false, HelpMessage = 'Use the Beta Graph API')]
    [switch] $BetaAPI,

    [Parameter(Mandatory = $true, HelpMessage = 'Please provide the Graph API ClientId')]
    [string] $ClientId = '',

    [Parameter(Mandatory = $true, HelpMessage = 'Please provide the Graph API TenantId')]
    [string] $TenantId = '',

    [Parameter(Mandatory = $true, HelpMessage = 'Please provide the Graph API ClientSecret')]
    [string] $ClientSecret = '',

    [Parameter(Mandatory = $true, HelpMessage = 'Please provide the path to where the report files will be saved to')]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({ if ($_ | Test-Path -PathType 'Container') { $true } else { throw "$_ is not a valid folder path" }  })]
    [String]$OutputFolderPath,

    [Parameter(Mandatory = $false, HelpMessage = 'Please provide the company name')]
    [string] $CompanyName = '',
    [Parameter(Mandatory = $false, HelpMessage = 'Please provide the company web')]
    [string] $CompanyWeb = '',
    [Parameter(Mandatory = $false, HelpMessage = 'Please provide the company email')]
    [string] $CompanyEmail = '',

    [Parameter(Mandatory = $false, HelpMessage = 'Export the tenant admin section')]
    [bool]$SectionTenantAdmin = $true,

    [Parameter(Mandatory = $false, HelpMessage = 'Export the enrollment section')]
    [bool]$SectionEnrollment = $true,

    [Parameter(Mandatory = $false, HelpMessage = 'Export the devices section')]
    [bool]$SectionDevices = $true,

    [Parameter(Mandatory = $false, HelpMessage = 'Export the users section')]
    [bool]$SectionUsers = $true,

    [Parameter(Mandatory = $false, HelpMessage = 'Export the groups section')]
    [bool]$SectionGroups = $true,

    [Parameter(Mandatory = $false, HelpMessage = 'Export the apps section')]
    [bool]$SectionApps = $true,

    [Parameter(Mandatory = $false, HelpMessage = 'Export the Endpoint Security section')]
    [bool]$SectionEndpointSecurity = $true,

    [Parameter(Mandatory = $false, HelpMessage = 'Export the Policies section')]
    [bool]$SectionPolicies = $true
)

$Error.Clear()
#region Functions
#region Test-RFLAdministrator
Function Test-RFLAdministrator {
<#
    .SYSNOPSIS
        Check if the current user is member of the Local Administrators Group

    .DESCRIPTION
        Check if the current user is member of the Local Administrators Group

    .NOTES
        Name: Test-RFLAdministrator
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Test-RFLAdministrator
#>
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    (New-Object Security.Principal.WindowsPrincipal $currentUser).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}
#endregion

#region Set-RFLLogPath
Function Set-RFLLogPath {
<#
    .SYSNOPSIS
        Configures the full path to the log file depending on whether or not the CCM folder exists.

    .DESCRIPTION
        Configures the full path to the log file depending on whether or not the CCM folder exists.

    .NOTES
        Name: Set-RFLLogPath
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Set-RFLLogPath
#>
    if ([string]::IsNullOrEmpty($script:LogFilePath)) {
        $script:LogFilePath = $env:Temp
    }

    $script:ScriptLogFilePath = "$($script:LogFilePath)\$($Script:LogFileFileName)"
}
#endregion

#region Write-RFLLog
Function Write-RFLLog {
<#
    .SYSNOPSIS
        Write the log file if the global variable is set

    .DESCRIPTION
        Write the log file if the global variable is set

    .PARAMETER Message
        Message to write to the log

    .PARAMETER LogLevel
        Log Level 1=Information, 2=Warning, 3=Error. Default = 1

    .NOTES
        Name: Write-RFLLog
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Write-RFLLog -Message 'This is an information message'

    .EXAMPLE
        Write-RFLLog -Message 'This is a warning message' -LogLevel 2

    .EXAMPLE
        Write-RFLLog -Message 'This is an error message' -LogLevel 3
#>
param (
    [Parameter(Mandatory = $true)]
    [string]$Message,

    [Parameter()]
    [ValidateSet(1, 2, 3)]
    [string]$LogLevel=1
)
    $TimeNow = Get-Date   
    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'
    if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)) {
        $ScriptName = ''
    } else {
        $ScriptName = $MyInvocation.ScriptName | Split-Path -Leaf
    }

    $LineFormat = $Message, $TimeGenerated, $TimeNow.ToString('MM-dd-yyyy'), "$($ScriptName):$($MyInvocation.ScriptLineNumber)", $LogLevel
    $Line = $Line -f $LineFormat

    $Line | Out-File -FilePath $script:ScriptLogFilePath -Append -NoClobber -Encoding default
    $HostMessage = '{0} {1}' -f $TimeNow.ToString('dd-MM-yyyy HH:mm'), $Message
    switch ($LogLevel) {
        2 { Write-Host $HostMessage -ForegroundColor Yellow }
        3 { Write-Host $HostMessage -ForegroundColor Red }
        default { Write-Host $HostMessage }
    }
}
#endregion

#region Clear-RFLLog
Function Clear-RFLLog {
<#
    .SYSNOPSIS
        Delete the log file if bigger than maximum size

    .DESCRIPTION
        Delete the log file if bigger than maximum size

    .NOTES
        Name: Clear-RFLLog
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Clear-RFLLog -maxSize 2mb
#>
param (
    [Parameter(Mandatory = $true)][string]$maxSize
)
    try  {
        if(Test-Path -Path $script:ScriptLogFilePath) {
            if ((Get-Item $script:ScriptLogFilePath).length -gt $maxSize) {
                Remove-Item -Path $script:ScriptLogFilePath
                Start-Sleep -Seconds 1
            }
        }
    }
    catch {
        Write-RFLLog -Message "Unable to delete log file." -LogLevel 3
    }    
}
#endregion

#region Get-ScriptDirectory
function Get-ScriptDirectory {
<#
    .SYSNOPSIS
        Get the directory of the script

    .DESCRIPTION
        Get the directory of the script

    .NOTES
        Name: ClearGet-ScriptDirectory
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Get-ScriptDirectory
#>
    Split-Path -Parent $PSCommandPath
}
#endregion

#endregion

#region ENUM List
$OperatingSystemList = @{
    linux = 'Linux';
    aospUserAssociated = 'Android (AOSP User)';
    androidDedicated = 'androidDedicated';
    androidFullyManaged = 'Android (Fully Managed)';
    configMgrDevice = 'Configuration Manager';
    unknown = 'unknown';
    androidDeviceAdmin = 'Android (Device Admin)';
    androidWorkProfile = 'Android (Work Profile)';
    aospUserless = 'Android (AOSP Userless)';
    windows = 'Windows';
    chromeOS = 'Chrome OS';
    android = 'Android';
    macOS = 'macOS';
    windowsMobile = 'Windows Mobile';
    androidCorporateWorkProfile = 'Android (Corporate Work Profile)';
    ios = 'iOS/iPadOS'
}

$skuPartNumberList = @{
    Microsoft_Intune_Suite = 'Microsoft Intune Suite';
}

$statusList = @{
    done = 'Complete';
    pending = 'Pending';
}

$AuthenticationMethodList = @{
    '#microsoft.graph.emailAuthenticationMethod' = 'E-mail';
    '#microsoft.graph.phoneAuthenticationMethod' = 'Phone';
    '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' = 'Windows Hello For Business';
    '#microsoft.graph.passwordAuthenticationMethod' = 'Password';
    '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' = 'Authenticator App';
}

$AppTypeList = @{
    '#microsoft.graph.iosStoreApp' = 'Apple Store';
    '#microsoft.graph.managedAndroidStoreApp' = 'Android Store';
    '#microsoft.graph.managedIOSStoreApp' = 'Apple Store';
    '#microsoft.graph.microsoftStoreForBusinessApp' = 'Microsoft Store';
    '#microsoft.graph.win32LobApp' = 'Win32 App';
}

$platformsList = @{
    'windows10' = 'Windows 10 and later';
    'windows10AndLater' = 'Windows 10 and later';
    'iOS' = 'iOS/iPadOS';
    'androidMobileApplicationManagement' = 'Android';
    'android' = 'Android device administrator';
    'macOS' = 'macOS';
    'androidForWork' = 'Android Enterprise';
    'iOSMobileApplicationManagement' = 'iOS/iPadOS';
    'd1174162-1dd2-4976-affc-6667049ab0ae' = 'Windows 10 and later';
    'a239407c-698d-4ef8-b314-e3ae409204b8' = 'macOS';
    '5340aa10-47a8-4e67-893f-690984e4d5da' = 'macOS';
}

$templateIDList = @{
    'd1174162-1dd2-4976-affc-6667049ab0ae' = 'BitLocker';
    'a239407c-698d-4ef8-b314-e3ae409204b8' = 'FileVault';
}

#endregion

#region Variables
$script:ScriptVersion = '0.3'
$script:LogFilePath = $env:Temp
$Script:LogFileFileName = 'ExportIntuneData.log'
$script:ScriptLogFilePath = "$($script:LogFilePath)\$($Script:LogFileFileName)"
$Script:Modules = @('MSAL.PS', 'PScribo')
$Script:CurrentFolder = (Get-Location).Path
$Global:ExecutionTime = Get-Date

if($BetaAPI -eq $true){
	$version = "beta"
} else {
	$version = "v1.0"
}
$Global:baseUri = '{0}/{1}' -f 'https://graph.microsoft.com', $version
$Global:BetabaseUri = '{0}/{1}' -f 'https://graph.microsoft.com', 'beta'

$params = @{
    ClientId = $ClientId
    TenantId = $TenantId
    ClientSecret = $ClientSecret | ConvertTo-SecureString -AsPlainText -Force			
}
$Global:ReportFile = '{0}' -f $TenantId
#endregion

#region Main
try {
    #region Start Script and checks
    Set-RFLLogPath
    Clear-RFLLog 25mb

    Write-RFLLog -Message "*** Starting ***"
    Write-RFLLog -Message "Script version $($script:ScriptVersion)"
    Write-RFLLog -Message "Running as $($env:username) $(if(Test-RFLAdministrator) {"[Administrator]"} Else {"[Not Administrator]"}) on $($env:computername)"

	Write-RFLLog -Message "Please refer to the RFL.Microsoft.Intune github website for more detailed information about this project." -LogLevel 2
	Write-RFLLog -Message "Documentation: https://github.com/dotraphael/RFL.Microsoft.Intune" -LogLevel 2
	Write-RFLLog -Message "Issues or bug reporting: https://github.com/dotraphael/RFL.Microsoft.Intune/issues" -LogLevel 2

    Write-RFLLog -Message "Report Name: $($Global:ReportFile)"
    Write-RFLLog -Message "Export Path: $($OutputFolderPath)"

    $PSCmdlet.MyInvocation.BoundParameters.Keys | ForEach-Object { 
        Write-RFLLog -Message "Parameter '$($_)' is '$($PSCmdlet.MyInvocation.BoundParameters.Item($_))'"
    }

    $PSVersionTable.Keys | ForEach-Object { 
        Write-RFLLog -Message "PSVersionTable '$($_)' is '$($PSVersionTable.Item($_) -join ', ')'"
    }    

    if ($PSVersionTable.item('PSVersion').Tostring() -notmatch '5.1') {
        throw "The requested operation requires PowerShell 5.1"
    }

    if ($PSVersionTable.item('PSEdition').Tostring() -eq 'Core') {
        throw "The requested operation requires PowerShell 5.1 (Desktop)"
    }

    Write-RFLLog -Message "Getting list of installed modules"
    $InstalledModules = Get-Module -ListAvailable -ErrorAction SilentlyContinue
    $InstalledModules | ForEach-Object { 
        Write-RFLLog -Message "    Module: '$($_.Name)', Type: '$($_.ModuleTYpe)', Verison: '$($_.Version)', Path: '$($_.ModuleBase)'"
    }

    Write-RFLLog -Message "Validating required PowerShell Modules"
    $Continue = $true
    foreach($item in $Script:Modules) {
        $ModuleInfo = $InstalledModules | Where-Object {$_.Name -eq $item}
        if ($null -eq $ModuleInfo) {
            Write-RFLLog -Message "    Module $($item) not installed. Use Install-Module $($item) -force to install the required powershell modules" -LogLevel 3
            $Continue = $false
        } else {
            Write-RFLLog -Message "    Module $($item) installed. Type: '$($ModuleInfo.ModuleTYpe)', Verison: '$($ModuleInfo.Version)', Path: '$($ModuleInfo.ModuleBase)'"
        } 
    }
    if (-not $Continue) {
        throw "The requested operation requires missing PowerShell Modules. Install the missing PowerShell modules and try again"
    }

    Write-RFLLog -Message "Current Folder '$($Script:CurrentFolder)'"

    Write-RFLLog -Message "All checks completed successful. Starting collecting data for report"
    #endregion

    #region Connect to Graph API
    Write-RFLLog -Message "Getting Graph API Token"
    $token = Get-MsalToken @params -ErrorAction Stop

    Write-RFLLog -Message "Connecting to MS Graph"
    $graphAccess = Connect-MgGraph -AccessToken $token.AccessToken -ErrorAction Stop
    $mgContext = Get-MgContext

    Write-RFLLog -Message "MS Graph Context Connection Info"

    $mgContext.psobject.properties | foreach-object {Write-RFLLog -Message ('    {0} = {1}' -f $_.Name, ($_.Value -join ', '))}
    #endregion

    #region main script
    #region Report
    $Global:WordReport = Document $Global:ReportFile {
        #region style
        DocumentOption -EnableSectionNumbering -PageSize A4 -DefaultFont 'Arial' -MarginLeftAndRight 71 -MarginTopAndBottom 71 -Orientation Portrait
        Style -Name 'Title' -Size 24 -Color '0076CE' -Align Center
        Style -Name 'Title 2' -Size 18 -Color '00447C' -Align Center
        Style -Name 'Title 3' -Size 12 -Color '00447C' -Align Left
        Style -Name 'Heading 1' -Size 16 -Color '00447C'
        Style -Name 'Heading 2' -Size 14 -Color '00447C'
        Style -Name 'Heading 3' -Size 12 -Color '00447C'
        Style -Name 'Heading 4' -Size 11 -Color '00447C'
        Style -Name 'Heading 5' -Size 11 -Color '00447C' -Bold
        Style -Name 'Heading 6' -Size 11 -Color '00447C' -Italic
        Style -Name 'Normal' -Size 10 -Color '565656' -Default
        Style -Name 'Caption' -Size 10 -Color '565656' -Italic -Align Left
        Style -Name 'Header' -Size 10 -Color '565656' -Align Center
        Style -Name 'Footer' -Size 10 -Color '565656' -Align Center
        Style -Name 'TOC' -Size 16 -Color '00447C'
        Style -Name 'TableDefaultHeading' -Size 10 -Color 'FAFAFA' -BackgroundColor '0076CE'
        Style -Name 'TableDefaultRow' -Size 10 -Color '565656'
        Style -Name 'Critical' -Size 10 -BackgroundColor 'F25022'
        Style -Name 'Warning' -Size 10 -BackgroundColor 'FFB900'
        Style -Name 'Info' -Size 10 -BackgroundColor '00447C'
        Style -Name 'OK' -Size 10 -BackgroundColor '7FBA00'

        Style -Name 'HeaderLeft' -Size 10 -Color '565656' -Align Left -BackgroundColor BDD6EE
        Style -Name 'HeaderRight' -Size 10 -Color '565656' -Align Right -BackgroundColor E7E6E6
        Style -Name 'FooterRight' -Size 10 -Color '565656' -Align Right -BackgroundColor BDD6EE
        Style -Name 'FooterLeft' -Size 10 -Color '565656' -Align Left -BackgroundColor E7E6E6
        Style -Name 'TitleLine01' -Size 18 -Color '565656' -Align Left -BackgroundColor BDD6EE
        Style -Name 'TitleLine02' -Size 10 -Color '565656' -Align Left -BackgroundColor BDD6EE
        Style -Name '1stPageRowStyle' -Size 10 -Color '565656' -Align Left -BackgroundColor E7E6E6


        # Configure Table Styles
        $TableDefaultProperties = @{
            Id = 'TableDefault'
            HeaderStyle = 'TableDefaultHeading'
            RowStyle = 'TableDefaultRow'
            BorderColor = '0076CE'
            Align = 'Left'
            CaptionStyle = 'Caption'
            CaptionLocation = 'Below'
            BorderWidth = 0.25
            PaddingTop = 1
            PaddingBottom = 1.5
            PaddingLeft = 2
            PaddingRight = 2
        }

        TableStyle @TableDefaultProperties -Default
        TableStyle -Name Borderless -HeaderStyle Normal -RowStyle Normal -BorderWidth 0
        TableStyle -Name 1stPageTitle -HeaderStyle Normal -RowStyle 1stPageRowStyle -BorderWidth 0

        # Microsoft Intune Cover Page Layout
        # Header & Footer
        Header -FirstPage {
            $Obj = [ordered] @{
                "CompanyName" = $COmpanyName
                "CompanyWeb" = $CompanyWeb
                "CompanyEmail" = $CompanyEmail
            }
            [pscustomobject]$Obj | Table -Style Borderless -list -ColumnWidths 50, 50 
        }

        Header -Default {
            $hashtableArray = @(
                [Ordered] @{ "Private and Confidential" = "Microsoft Intune"; '__Style' = 'HeaderLeft'; "Private and Confidential__Style" = 'HeaderRight';}
            )
            Table -Hashtable $hashtableArray -Style Borderless -ColumnWidths 30, 70 -list
        }

        Footer -Default {
            $hashtableArray = @(
                [Ordered] @{ " " = 'Page <!# PageNumber #!> of <!# TotalPages #!>'; '__Style' = 'FooterLeft'; " __Style" = 'FooterRight';}
            )
            Table -Hashtable $hashtableArray -Style Borderless -ColumnWidths 30, 70 -list
        }

        BlankLine -Count 11
        $LineCount = 32 + $LineCount

        # Microsoft Logo Image
        Try {
            Image -Text 'Microsoft Logo' -Align 'Center' -Percent 20 -Base64 "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"
            BlankLine -Count 2
        } Catch {
            Write-RFLLog -Message ".NET Core is required for cover page image support. Please install .NET Core from https://dotnet.microsoft.com/en-us/download" -LogLevel 3
        }

        # Add Report Name
        $Obj = [ordered] @{
            " " = ""
            " __Style" = "TitleLine02"
            "  " = "Intune Report"
            "  __Style" = "TitleLine01"
            "   " = "Report Generated on $($Global:ExecutionTime.ToString("dd/MM/yyyy")) and $($Global:ExecutionTime.ToString("HH:mm:ss"))"
            "   __Style" = "TitleLine02"
            "    " = ""
            "    __Style" = "TitleLine02"
        }
        [pscustomobject]$Obj | Table -Style 1stPageTitle -list -ColumnWidths 10, 90 
        PageBreak

        # Add Table of Contents
        TOC -Name 'Table of Contents'
        PageBreak

        #region Executive Summary
        $sectionName = 'Introduction'
        Write-RFLLog -Message "Starting Section '$($sectionName)'"
        Section -Style Heading1 $sectionName {
	        try {
		        Paragraph "This document describes the overall configuration of the Microsoft Intune for tenant $($TenantId)."
		        BlankLine
                PageBreak
	        }
	        catch {
		        Write-RFLLog -Message $_.Exception.Message -LogLevel 3
	        }
        }
        #endregion
        #endregion

        #region Tenant Admin
        $sectionName = 'Tenant admin'
        if ($SectionTenantAdmin -eq $false) {
	        Write-RFLLog -Message "Exporting Section '$($sectionName)' is being ignored as the parameter to export is set to false" -LogLevel 2
        } else {
	        Write-RFLLog -Message "Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Tenant Details
                $SectionName = "Tenant Details"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
		            Write-RFLLog -Message "        MSGraph: organization"
                    $org = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/organization"

                    Write-RFLLog -Message "        MSGraph: servicePrincipals/appId=0000000a-0000-0000-c000-000000000000/endpoints"
                    $endpoints = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/servicePrincipals/appId=0000000a-0000-0000-c000-000000000000/endpoints"

                    Write-RFLLog -Message "        MSGraph: organization('id')?`$select=mobiledevicemanagementauthority"
                    $mdmAuth = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/organization('$($org.value.id)')?`$select=mobiledevicemanagementauthority"

		            Write-RFLLog -Message "        MSGraph: deviceManagement/subscriptionState"
		            $accStatus = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/deviceManagement/subscriptionState"

		            Write-RFLLog -Message "        MSGraph: subscribedSkus"
		            $subs = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/subscribedSkus"

                    Write-RFLLog -Message "        MSGraph: deviceManagement/managedDeviceOverview"
		            $mdo = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/deviceManagement/managedDeviceOverview"
                    #endregion

                    #region Generating Data
                    $OutObj = New-Object PSObject -Property @{
			            'Tenant Name' = $org.Value.displayName
			            'Tenant Location' = '{0} {1}' -f ($endpoints.Value | Where-Object {$_.providerName -eq 'Region'}).uri, ($endpoints.Value | Where-Object {$_.providerName -eq 'ASUName'}).Uri.Replace('AMSUB','')
			            'MDM authority' = $mdmAuth.mobileDeviceManagementAuthority
			            'Account status' = $accStatus.value
			            'Total enrolled devices' = $mdo.enrolledDeviceCount
			            'Total licensed users' = (($subs.Value | Where-Object {$_.skuPartNumber -in ('EMSPREMIUM', 'Microsoft_Intune_Suite')}).consumedUnits | Measure-Object -Sum).Sum
			            'Total Intune licenses' = (($subs.Value | Where-Object {$_.skuPartNumber -in ('EMSPREMIUM', 'Microsoft_Intune_Suite')}).prepaidunits.enabled | Measure-Object -Sum).Sum
		            }
		            $TableParams = @{
                        Name = $SectionName
                        List = $true
                        ColumnWidths = 40, 60
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($null -ne $OutObj) {
		                $OutObj | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Roles
                $SectionName = "Roles and Assignments"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
		            Write-RFLLog -Message "        MSGraph: deviceManagement/roleDefinitions"
                    $roles = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/deviceManagement/roleDefinitions"

		            Write-RFLLog -Message "        MSGraph: deviceManagement/roleDefinitions/{id}/roleAssignments"
                    $OutObj = @()
                    foreach($role in $roles.value) {
                        $assignments = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/deviceManagement/roleDefinitions/$($role.id)/roleAssignments"
                        $OutObj += New-Object PSObject -Property @{
			                'Display Name' = $role.displayName
                            'Built In' = $role.isBuiltIn
                            'Assignments' = $assignments.Value.displayName -join ', '
		                }
                    }
		        
                    Write-RFLLog -Message "        MSGraph (Beta): roleManagement/cloudPC/roleDefinitions"
                    ##Todo: use v1.0 version when available
                    $roles = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/roleManagement/cloudPC/roleDefinitions"

		            Write-RFLLog -Message "        MSGraph (Beta): roleManagement/cloudPC/roleAssignments?`$filter=roleDefinitionId eq `'{1}`'"
                    foreach($role in $roles.value) {
                        ##Todo: use v1.0 version when available
                        $assignments = Invoke-MgGraphRequest -Method GET -Uri ('{0}/roleManagement/cloudPC/roleAssignments?$filter=roleDefinitionId eq ''{1}''' -f $Global:BetabaseUri, $role.id)

                        $OutObj += New-Object PSObject -Property @{
			                'Display Name' = $role.displayName
                            'Built In' = $role.isBuiltIn
                            'Assignments' = $assignments.Value.displayName -join ', '
		                }
                    }
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        Header = 'Display Name','Built In', 'Assignments'
				        Columns = 'Display Name', 'Built In', 'Assignments'
                        ColumnWidths = 40, 20, 40
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($OutObj.Count -gt 0) {
		                $OutObj | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Device diagnostics
                $SectionName = "Device diagnostics"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
		            Write-RFLLog -Message "        MSGraph: deviceManagement/settings"
                    $outObj = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/deviceManagement/settings"
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $true
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($OutObj.Count -gt 0) {
		                $OutObj | select @{Name="Device diagnostics";Expression = { $_.enableLogCollection }}, @{Name="Autopilot diagnostics";Expression = { $_.enableAutopilotDiagnostics }} | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }		    
                    #endregion
                }
                #endregion

                #region Terms and Conditions
                $SectionName = "Terms and Conditions"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
		            Write-RFLLog -Message "        MSGraph: deviceManagement/termsAndConditions"
                    $tcList = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/deviceManagement/termsAndConditions"

		            Write-RFLLog -Message "        MSGraph: deviceManagement/termsAndConditions/{id}/assignments"
                    $OutObj = @()
                    foreach($tc in $tcList.value) {
                        $assignments = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/deviceManagement/termsAndConditions/$($tc.id)/assignments"
                        $groupInfo = @()
                        foreach($assignment in $assignments.Value.target.groupID) {
                            $groupInfo += Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/groups/$($assignment)"
                        }

                        $OutObj += New-Object PSObject -Property @{
                            'Name' = $tc.displayName
                            'Create Date' = $tc.createdDateTime
                            'Modified Date' = $tc.lastModifiedDateTime
                            'Version' = $tc.version
                            'Assignments' = $groupInfo.displayName -join ', '
		                }
                    }
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        Header = 'Name', 'Create Date', 'Modified Date', 'Version', 'Assignments'
				        Columns = 'Name', 'Create Date', 'Modified Date', 'Version', 'Assignments'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($OutObj.Count -gt 0) {
		                $OutObj | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Intune add-ons
                $SectionName = "Intune add-ons"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
		            Write-RFLLog -Message "        MSGraph: subscribedSkus"
                    $skus = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/subscribedSkus"
                    $OutObj = @()
                    foreach($item in ($skus.value | where-object {$_.skuPartNumber -in @('Microsoft_Intune_Suite')}) ) {
                        $OutObj += New-Object PSObject -Property @{
			                'Add-on' = $skuPartNumberList."$($item.skuPartNumber)"
                            'Purchased' = $item.prepaidUnits.enabled
			                'Consumed' = $item.consumedUnits
		                }
                    }
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        Header = 'Add-on','Purchased', 'Consumed'
				        Columns = 'Add-on', 'Purchased', 'Consumed'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($OutObj.Count -gt 0) {
		                $OutObj | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
        }
        #endregion

        #region Enrollment
        $sectionName = 'Enrollment'
        if ($SectionEnrollment -eq $false) {
	        Write-RFLLog -Message "Exporting Section '$($sectionName)' is being ignored as the parameter to export is set to false" -LogLevel 2
        } else {
            PageBreak
	        Write-RFLLog -Message "Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Device limit restrictions
                $SectionName = "Device limit restrictions"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    Write-RFLLog -Message "        MSGraph (Beta): deviceManagement/deviceEnrollmentConfigurations?`$expand=assignments&`$filter=deviceEnrollmentConfigurationType%20eq%20%27Limit%27"
                    ##Todo: use v1.0 version when available
                    $enrollmentlimitList = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/deviceEnrollmentConfigurations?`$expand=assignments&`$filter=deviceEnrollmentConfigurationType%20eq%20%27Limit%27"

                    Write-RFLLog -Message "        MSGraph (Beta): deviceManagement/deviceEnrollmentConfigurations/{id}/assignments"
                    $OutObj = @()
                    foreach($item in $enrollmentlimitList.value) {
                        $groupInfo = @()
                        if ($item.priority -eq 0) {
                            $groupInfo += New-Object PSObject -Property @{ 'displayName' = 'All users and all devices'}
                        } else {
                            ##Todo: use v1.0 version when available
                            $assignments = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/deviceEnrollmentConfigurations/$($item.id)/assignments"
                            foreach($assignment in $assignments.value.target.groupid) {
                                $groupInfo += Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/groups/$($assignment)"
                            }
                        }

                        $OutObj += New-Object PSObject -Property @{
                            'Name' = $item.displayName
                            'Priority' = $item.priority
                            'Limit' = $item.limit
                            'Assignments' = $groupInfo.displayName -join ', '
		                }
                    }
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        Header = 'Priority', 'Name', 'Limit', 'Assignments'
				        Columns = 'Priority', 'Name', 'Limit', 'Assignments'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($OutObj.Count -gt 0) {
		                $OutObj | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Device enrollment managers
                $SectionName = "Device enrollment managers"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    Write-RFLLog -Message "        MSGraph: users?`$filter=deviceEnrollmentLimit%20eq%201000&`$select=id,displayName,userPrincipalName&`$skip=0"
                    $outobj = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/users?`$filter=deviceEnrollmentLimit%20eq%201000&`$select=id,displayName,userPrincipalName&`$skip=0"
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
				        Header = 'Name','UPN'
				        Columns = 'displayName', 'userPrincipalName'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($outobj.value.count -gt 0) {
		                $outobj.value | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region todo:
                <#
                windows enrollment
                apple enrollment
                android enrollment
                enrollment device platform restrictions
                corporate device identifiers
                #>
                #endregion
            }
        }
        #endregion

        #region Devices
        $sectionName = 'Devices'
        if ($SectionDevices -eq $false) {
	        Write-RFLLog -Message "Exporting Section '$($sectionName)' is being ignored as the parameter to export is set to false" -LogLevel 2
        } else {
            PageBreak
	        Write-RFLLog -Message "Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Data Used by the next few sections
                Write-RFLLog -Message "    MSGraph (Beta): deviceManagement/managedDevices"
                ##Todo: use v1.0 version when complianceState, ownerType, jointype, autopilotEnrolled, managementState are available 
                $mdmList = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/managedDevices"

                #$mdmList.value | Group-Object complianceState - does not work as expected. it returns name = "" so fixing with foreach
                $mdmListFixed = @()
                foreach($item in $mdmList.value) {
                    $mdmListFixed += New-Object PSObject -Property @{
                        id = $item.id
                        complianceState = $item.complianceState
                        operatingSystem = $item.operatingSystem
                        osVersion = $item.osVersion
                        ownerType = $item.ownerType
                        model = $item.model
                        managementAgent = $item.managementAgent
                        joinType = $item.joinType
                        autopilotEnrolled = $item.autopilotEnrolled
                        managementState = $item.managementState
                    }
                }

                Write-RFLLog -Message "    MSGraph (Beta): deviceManagement/manageddevices/{id)}?`$select=deviceactionresults,managementstate,lostModeState,deviceRegistrationState,ownertype"
                $das = @()
                foreach($item in $mdmList.Value) {
                    ##Todo: use v1.0 version when available
                    $dasList = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/manageddevices/{$($item.id)}?`$select=deviceactionresults,managementstate,lostModeState,deviceRegistrationState,ownertype"
                    foreach($dasItem in $dasList.deviceActionResults) {
                        $dasobj = New-Object PSObject -Property @{
                            'DeviceID' = $item.id
                            'Name' = $item.deviceName
                            'Action' = ''
                            'Status' = $statusList."$($dasItem.actionState)"
                            'DateTime' = $dasItem.startDateTime
                            'actionName' = $dasItem.actionName
                        }

                        $dasobj.Action = switch ($dasItem.actionName.tolower()) {
                            'locatedevice' { 'Locate device' }
                            'windowsdefenderscan' { $dasItem.scanType }
                            default { $dasItem.actionName }
                        }

                        $das += $dasobj
                    }
                }
                #endregion

                #region Enrollment Status
                $SectionName = "Enrollment Status"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    Write-RFLLog -Message "        MSGraph: deviceManagement/managedDeviceOverview"
		            $mdo = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/deviceManagement/managedDeviceOverview"
                    #endregion

                    #region Generating Data
                    $OutObj = @()
                    foreach($item in $mdo.deviceOperatingSystemSummary.GetEnumerator()){
                        $OutObj += New-Object PSObject -Property @{
			                'Operating System' = $OperatingSystemList."$($item.key.Replace('Count',''))"
			                'Count' = $item.Value
		                }
                    }
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Operating System', 'Count'
				        Columns = 'Operating System', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($outobj.count -gt 0) {
		                $outobj | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Compliance
                $SectionName = "Compliance"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Name', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($mdmList.value.count -gt 0) {
                        #$mdmList.value | Group-Object complianceState does not work as expected. it returns name = "" so fixing with foreach
		                #$mdmList.value | Group-Object complianceState | Table @TableParams
                        $mdmListFixed | Group-Object complianceState | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Ownership
                $SectionName = "Ownership"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Name', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($mdmList.value.count -gt 0) {
                        #$mdmList.value | Group-Object ownerType does not work as expected. it returns name = "" so fixing with foreach
		                #$mdmList.value | Group-Object ownerType | Table @TableParams
                        $mdmListFixed | Group-Object ownerType | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Join Type
                $SectionName = "Join Type"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Name', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($mdmList.value.count -gt 0) {
                        #$mdmList.value | Group-Object joinType does not work as expected. it returns name = "" so fixing with foreach
		                #$mdmList.value | Group-Object joinType | Table @TableParams
                        $mdmListFixed | Group-Object joinType | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Autopilot Enrolled
                $SectionName = "Autopilot Enrolled"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Name', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($mdmList.value.count -gt 0) {
                        #$mdmList.value | Group-Object autopilotEnrolled does not work as expected. it returns name = "" so fixing with foreach
		                #$mdmList.value | Group-Object autopilotEnrolled | Table @TableParams
                        $mdmListFixed | Group-Object autopilotEnrolled | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion        

                #region Managed By
                $SectionName = "Managed By"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Name', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($mdmList.value.count -gt 0) {
                        #$mdmList.value | Group-Object managementAgent does not work as expected. it returns name = "" so fixing with foreach
		                #$mdmList.value | Group-Object managementAgent | Table @TableParams
                        $mdmListFixed | Group-Object managementAgent | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
		            #endregion
                }
                #endregion

                #region Management State
                $SectionName = "Management State"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Name', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($mdmList.value.count -gt 0) {
                        #$mdmList.value | Group-Object managementState does not work as expected. it returns name = "" so fixing with foreach
		                #$mdmList.value | Group-Object managementState | Table @TableParams
                        $mdmListFixed | Group-Object managementState | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion     

                #region Operating System Name/Version
                $SectionName = "Operating System Name/Version"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Name', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($mdmList.value.count -gt 0) {
                        #$mdmList.value | Group-Object operatingSystem,osVersion does not work as expected. it returns name = "" so fixing with foreach
		                #$mdmList.value | Group-Object operatingSystem,osVersion | Table @TableParams
                        $mdmListFixed | Group-Object operatingSystem,osVersion | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Hardware Model
                $SectionName = "Hardware Model"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 80, 20
                        Header = 'Name', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($mdmList.value.count -gt 0) {
                        #$mdmList.value | Group-Object model does not work as expected. it returns name = "" so fixing with foreach
		                #$mdmList.value | Group-Object model | Table @TableParams
                        $mdmListFixed | Group-Object model | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Device actions status
                $SectionName = "Device actions status"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Name', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($das.count -gt 0) {
		                $das | Group-Object Action,Status | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Device clean-up rules
                $SectionName = "Device clean-up rules"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    Write-RFLLog -Message "        MSGraph (Beta): deviceManagement/managedDeviceCleanupSettings"
                    ##Todo: use v1.0 version when available
                    $OutObj = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/managedDeviceCleanupSettings"
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $true
                        ColumnWidths = 80, 20
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($OutObj.count -gt 0) {
		                $OutObj | select @{Name="Delete devices based on last check-in date";Expression = { $null -ne $_.deviceInactivityBeforeRetirementInDays }}, @{Name="Delete devices that haven't checked in for this many days";Expression = { $_.deviceInactivityBeforeRetirementInDays }} | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Device Categories
                $SectionName = "Device Categories"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    Write-RFLLog -Message "        MSGraph: deviceManagement/deviceCategories"
                    $OutObj = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/deviceManagement/deviceCategories"
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($OutObj.Value.Count -gt 0) {
    		            $OutObj.Value | select @{Name="Display Name";Expression = { $_.displayName }}, @{Name="Description";Expression = { $_.description }} | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Device Filters
                $SectionName = "Device Filters"
                ##Todo: add filters scope and rules. Subsection?!
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    Write-RFLLog -Message "        MSGraph (Beta): deviceManagement/assignmentFilters"
                    ##Todo: use v1.0 version when available
                    $OutObj = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/assignmentFilters"
                    $TableOut = @()
                    $TableOut += $outobj.value | Where-Object {$_.assignmentFilterManagementType -eq 'devices'}
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 45, 30, 25
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($TableOut.Count -gt 0) {
    		            $TableOut | select @{Name="Display Name";Expression = { $_.displayName }}, @{Name="Platform";Expression = { $platformsList."$($_.platform)"}}, @{Name="Last modified";Expression = { $_.lastModifiedDateTime }} | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region todo:
                <#
            
                #>
                #endregion
            }
        }
        #endregion

        #region Users
        $sectionName = 'Users'
        if ($SectionUsers -eq $false) {
	        Write-RFLLog -Message "Exporting Section '$($sectionName)' is being ignored as the parameter to export is set to false" -LogLevel 2
        } else {
            PageBreak
	        Write-RFLLog -Message "Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Data Used by the next few sections
                Write-RFLLog -Message "    MSGraph (Beta): users"
                ##Todo: use v1.0 version when userType, accountEnabled, onPremisesSyncEnabled, passwordPolicies, usageLocation are available 
                $userList = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/users"

                #$userList.value | Group-Object userType - does not work as expected. it returns name = "" so fixing with foreach
                $userListFixed = @()
                foreach($item in $userList.value) {
                    $userListFixed += New-Object PSObject -Property @{
                        id = $item.id
                        userType = $item.userType
                        accountEnabled = $item.accountEnabled
                        onPremisesSyncEnabled = $item.onPremisesSyncEnabled
                        passwordPolicies = $item.passwordPolicies
                        usageLocation = $item.usageLocation
                    }
                }
                #endregion

                #region User Type
                $SectionName = "User Type"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Name', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($userList.value.count -gt 0) {
                        #$userList.value | Group-Object userType does not work as expected. it returns name = "" so fixing with foreach
		                #$userList.value | Group-Object userType | Table @TableParams
                        $userListFixed | Group-Object userType | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region account status
                $SectionName = "Account Enabled Status"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Name', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($userList.value.count -gt 0) {
                        #$userList.value | Group-Object accountEnabled does not work as expected. it returns name = "" so fixing with foreach
		                #$userList.value | Group-Object accountEnabled | Table @TableParams
                        $userListFixed | Group-Object accountEnabled | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region onprem sync enabled
                $SectionName = "onprem sync enabled"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Name', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($userList.value.count -gt 0) {
                        #$userList.value | Group-Object onPremisesSyncEnabled does not work as expected. it returns name = "" so fixing with foreach
		                #$userList.value | Group-Object onPremisesSyncEnabled | Table @TableParams
                        $userListFixed | Group-Object onPremisesSyncEnabled | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region password Policies
                $SectionName = "password Policies"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Name', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($userList.value.count -gt 0) {
                        #$userList.value | Group-Object passwordPolicies does not work as expected. it returns name = "" so fixing with foreach
		                #$userList.value | Group-Object passwordPolicies | Table @TableParams
                        $userListFixed | Group-Object passwordPolicies | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region usage Location
                $SectionName = "usage Location"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Name', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($userList.value.count -gt 0) {
                        #$userList.value | Group-Object usageLocation does not work as expected. it returns name = "" so fixing with foreach
		                #$userList.value | Group-Object usageLocation | Table @TableParams
                        $userListFixed | Group-Object usageLocation | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region identities
                $SectionName = "identities"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    $userIdentities = @()
                    foreach($item in $userList.value) {
                        foreach($identityItem in $item.identities) {
                            $userIdentities += New-Object PSObject -Property @{ 'issuer' = $identityItem.issuer }
                        }
                    }
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Name', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($userIdentities.count -gt 0) {
		                $userIdentities | Group-Object issuer | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region registeredDevices
                $SectionName = "Registered Devices"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    Write-RFLLog -Message "        MSGraph: users/{id}/registeredDevices"
                    $regDevices = @()
                    foreach($item in $userList.value) {
                        $userDevices = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/users/{$($item.id)}/registeredDevices"
                        $regDevices += New-Object PSObject -Property @{ 
                            'user' = $item.userPrincipalName
                            'count' = $userDevices.value.count
                        }
                    }
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Number of Registered Devices', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($regDevices.count -gt 0) {
		                $regDevices | Group-Object count | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Authentication Method
                $SectionName = "Authentication Method"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    Write-RFLLog -Message "        MSGraph: users/{UPN}/authentication/methods"
                    $authUsers = @()
                    foreach($item in $userList.value) {
                        $authMethods = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/users/$($item.userPrincipalName)/authentication/methods"
                        foreach($authItem in $authMethods.value) {
                            $authUsers += New-Object PSObject -Property @{ 
                                'user' = $item.userPrincipalName
                                'Method' = $AuthenticationMethodList."$($authItem.'@odata.type')"
                            }
                        }
                    }
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Authentication Method', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($authUsers.count -gt 0) {
		                $authUsers | Select-Object 'user','method' -Unique | Group-Object Method | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }        
                #endregion

                #region MFA Enabled
                $SectionName = "MFA Enabled"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    $mfaUsers = @()
                    foreach($item in $userList.value) {
                        $mfaUsers += New-Object PSObject -Property @{
                            'user' = $item.userPrincipalName
                            'MFA Status' = (($authUsers | Where-Object {($_.user -eq $item.userPrincipalName) -and ($_.Method -ne 'Password')} | Measure-Object).Count -gt 0)
                        }
                    }
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'MFA Enabled', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($mfaUsers.count -gt 0) {
		                $mfaUsers | Group-Object 'MFA Status' | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region last signin
                $SectionName = "Last Sign In"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    Write-RFLLog -Message "        MSGraph: users?`$select=displayName,signInActivity"
                    $userLastLogin = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/users?`$select=displayName,signInActivity"
                    $outobj = @()
                    foreach($item in $userLastLogin.value) {
                        $Obj = New-Object PSObject -Property @{
                            'user' = $item.displayName
                            'datetime' = $item.signInActivity.lastSignInDateTime
                            'DaysDifference' = 0
                            'DaysRange' = ''                    
                        }

                        if ($null -eq $item.signInActivity.lastSignInDateTime) {
                            $obj.DaysDifference = -1
                        } else {
                            $obj.DaysDifference = (New-TimeSpan -End (get-date) -Start ($item.signInActivity.lastSignInDateTime)).Days
                        }

                        $obj.DaysRange = switch ($obj.DaysDifference) {
                            -1 { 'Never' }
                            0 { 'Today' }
                            1..7 { 'Last 7 Days'}
                            8..14 { 'Last 2 Weeks' }
                            15..30 { 'This month' }
                            31..90 { 'Last 3 months' }
                            default { 'Over 90 days' }
                        }
                        $outobj += $obj
                    }
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Last Sign in', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($outobj.count -gt 0) {
		                $outobj | Group-Object DaysRange | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region todo:
                <#
                password reset 
                user settings
                #>
                #endregion

            }
        }
        #endregion

        #region Groups
        $sectionName = 'Groups'
        if ($SectionGroups -eq $false) {
	        Write-RFLLog -Message "Exporting Section '$($sectionName)' is being ignored as the parameter to export is set to false" -LogLevel 2
        } else {
            PageBreak
	        Write-RFLLog -Message "Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Data Used by the next few sections
                Write-RFLLog -Message "    MSGraph: groups"
                $groupList = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/groups"

                Write-RFLLog -Message "    MSGraph: groups/{id}/members"
                $groupInfo = @()
                foreach($item in $groupList.Value) {
                    $members = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/groups/{$($item.id)}/members"
                    $owners = Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/groups/{$($item.id)}/owners"

                    $groupInfo += New-Object PSObject -Property @{
                        'id' = $item.id
                        'groupname' = $item.displayName
                        'members' = $members.Value.count
                        'owners' = $owners.Value.count
                    }
                }
                #endregion

                #region Group Membership Type
                $SectionName = "Group Membership Type"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Group Membership Type', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($groupList.value.count -gt 0) {
		                $groupList.value | select id, @{Name="GroupType";Expression = { if ($null -eq $_.membershipRule) { 'Assigned' } else {'Dynamic'} }} | Group-Object GroupType | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Group Type
                $SectionName = "Group Type"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Group Type', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($groupList.value.count -gt 0) {
		                $groupList.value | select id, @{Name="GroupType";Expression = { if ($false -eq $_.mailEnabled) { 'Security' } else {'Microsoft 365'} }} | Group-Object GroupType | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Group Source
                $SectionName = "Group Source"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Source', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($groupList.value.count -gt 0) {
		                $groupList.value | select id, @{Name="GroupSource";Expression = { if ($null -eq $_.onPremisesLastSyncDateTime) { 'Cloud' } else {'On Premises'} }} | Group-Object GroupSource | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Group Membership Count
                $SectionName = "Group Membership Count"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Membership Count', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($groupInfo.count -gt 0) {
		                $groupInfo | Group-Object Members | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Group Owners Count
                $SectionName = "Group Owners Count"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Ownership Count', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($groupInfo.count -gt 0) {
		                $groupInfo | Group-Object owners | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
        }
        #endregion

        #region Apps
        $sectionName = 'Apps'
        if ($SectionApps -eq $false) {
	        Write-RFLLog -Message "Exporting Section '$($sectionName)' is being ignored as the parameter to export is set to false" -LogLevel 2
        } else {
            PageBreak
	        Write-RFLLog -Message "Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Data Used by the next few sections
                ##Todo: use v1.0 version when isAssigned is available 
                Write-RFLLog -Message "    MSGraph (Beta): deviceAppManagement/mobileApps"
                $AppList = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceAppManagement/mobileApps"

                Write-RFLLog -Message "    MSGraph (Beta): deviceAppManagement/mobileApps/{id}/installSummary"
                $AppInstallList = @()
                foreach($item in $AppList.Value) {
                    ##Todo: use v1.0 version when available
                    $installInfo = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceAppManagement/mobileApps/$($item.id)/installSummary"
                    $AppInstallList += New-Object PSObject -Property @{
                        'platform' = $AppTypeList."$($item.'@odata.type')"
                        'id' = $item.id
                        'installedDeviceCount' = $installInfo.installedDeviceCount
                        'failedDeviceCount' = $installInfo.failedDeviceCount
                        'notApplicableDeviceCount' = $installInfo.notApplicableDeviceCount
                        'notInstalledDeviceCount' = $installInfo.notInstalledDeviceCount
                        'pendingInstallDeviceCount' = $installInfo.pendingInstallDeviceCount
                        'installedUserCount' = $installInfo.installedUserCount
                        'notApplicableUserCount' = $installInfo.notApplicableUserCount
                        'failedUserCount' = $installInfo.failedUserCount
                        'notInstalledUserCount' = $installInfo.notInstalledUserCount
                        'pendingInstallUserCount' = $installInfo.pendingInstallUserCount
                    }
                }
                #endregion

                #region Apps Per Type
                $SectionName = "Apps Type"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'App Type', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($AppList.value.count -gt 0) {
		                $AppList.value | select id, @{Name="AppType";Expression = { $AppTypeList."$($_.'@odata.type')" }} | Group-Object AppType | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Assigned Apps
                $SectionName = "Assigned Apps Status"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 40, 60
                        Header = 'Assigned Status', 'Count'
				        Columns = 'Name', 'Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($AppList.value.count -gt 0) {
		                $AppList.value | Group-Object isAssigned | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region App Installation Status
                $SectionName = "Apps Installation Status"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 25, 25, 25, 25
                        Header = 'Platform', 'Application Count', 'Failed Device Count', 'Failed User Count'
				        Columns = 'Platform', 'Application Count', 'Failed Device Count', 'Failed User Count'
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($AppInstallList.count -gt 0) {
		                $AppInstallList | Group-Object platform,failedDeviceCount,failedUserCount | select @{Name="Application Count";Expression={$_.Count}}, @{Name="platform";Expression={$_.Group[0].Platform}}, @{Name="Failed Device Count";Expression={$_.group[0].failedDeviceCount}}, @{Name="Failed User Count";Expression={$_.group[0].failedUserCount}} | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region App categories
                $SectionName = "App categories"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    ##Todo: use v1.0 version when available
                    Write-RFLLog -Message "    MSGraph (Beta): deviceAppManagement/mobileAppCategories"
                    $OutObj = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceAppManagement/mobileAppCategories"
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 100
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($OutObj.value.count -gt 0) {
		                $OutObj.value | select @{Name="Name";Expression={$_.displayname}} | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region App Filters
                $SectionName = "App Filters"
                ##Todo: add filters scope and rules. Subsection?!
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    Write-RFLLog -Message "        MSGraph (Beta): deviceManagement/assignmentFilters"
                    ##Todo: use v1.0 version when available
                    $OutObj = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/assignmentFilters"
                    $TableOut = @()
                    $TableOut += $outobj.value | Where-Object {$_.assignmentFilterManagementType -eq 'apps'}
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 45, 30, 25
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($TableOut.Count -gt 0) {
    		            $TableOut | select @{Name="Display Name";Expression = { $_.displayName }}, @{Name="Platform";Expression = { $platformsList."$($_.platform)"}}, @{Name="Last modified";Expression = { $_.lastModifiedDateTime }} | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Ebook
                $SectionName = "Ebook"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    Write-RFLLog -Message "        MSGraph (Beta): deviceAppManagement/managedEBooks"
                    ##Todo: use v1.0 version when available
                    $OutObj = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceAppManagement/managedEBooks"
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 45, 30, 25
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($OutObj.value.Count -gt 0) {
    		            $OutObj.value | select @{Name="Display Name";Expression = { $_.displayName }}, @{Name="publisher";Expression = { $_.publisher}}, @{Name="Information Url";Expression = { $_.informationUrl }} | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Ebook categories
                $SectionName = "Ebook categories"
                ##Todo: add filters scope and rules. Subsection?!
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Collect Data
                    Write-RFLLog -Message "        MSGraph (Beta): deviceAppManagement/managedEBookCategories"
                    ##Todo: use v1.0 version when available
                    $OutObj = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceAppManagement/managedEBookCategories"
                    #endregion

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                        ColumnWidths = 100
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($OutObj.Value.Count -gt 0) {
    		            $OutObj.Value | select @{Name="Display Name";Expression = { $_.displayName }} | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion
                #region todo:
                <# 
                app protection policy
                app configuration policy
                ios app provisioning profiles
                s mode supplemental policies
                policies for office apps
                App selective wipe

                apple vpp tokens
                ebook installation status
                #>
                #endregion

            }
        }
        #endregion

        #region Endpoint security
        $sectionName = 'Endpoint security'
        if ($SectionEndpointSecurity -eq $false) {
	        Write-RFLLog -Message "Exporting Section '$($sectionName)' is being ignored as the parameter to export is set to false" -LogLevel 2
        } else {
            PageBreak
	        Write-RFLLog -Message "Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Anti-Virus
                $SectionName = "Anti-Virus"
                Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Data Used by the next few sections
                    #region AntiVirus
                    Write-RFLLog -Message "        MSGraph (Beta): deviceManagement/configurationPolicies?`$select=id,name,description,platforms,lastModifiedDateTime,technologies,settingCount,roleScopeTagIds,isAssigned,templateReference&`$filter=templateReference/TemplateFamily eq 'endpointSecurityAntivirus'"
                    ##Todo: use v1.0 version when available
                    $avList = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/configurationPolicies?`$select=id,name,description,platforms,lastModifiedDateTime,technologies,settingCount,roleScopeTagIds,isAssigned,templateReference&`$filter=templateReference/TemplateFamily eq 'endpointSecurityAntivirus'"

                    Write-RFLLog -Message "        MSGraph (Beta): deviceManagement/configurationPolicies/{1}/assignments"
                    $avAssignmentList = @()
                    foreach($item in $avList.value) {
                        ##Todo: use v1.0 version when available
                        $assignments = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/configurationPolicies/$($item.id)/assignments"
                        $groupInfo = @()
                        foreach($assignment in $assignments.Value.target.groupID) {
                            $groupInfo += Invoke-MgGraphRequest -Method GET -Uri "$($Global:baseUri)/groups/$($assignment)"
                        }

                        $avAssignmentList += New-Object PSObject -Property @{
                            'platforms' = $item.platforms
                            'name' = $item.name
                            'templatereference' = $item.templatereference
                            'technologies' = $item.technologies
                            'lastModifiedDateTime' = $item.lastModifiedDateTime
                            'Assignments' = $groupInfo.displayName -join ', '
		                }
                    }
                    #endregion
                    #endregion

                    #region Malware Overview
                    $SectionName = "Malware Overview"
                    Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                    Section -Style Heading3 $SectionName {
                        #region Collect Data
                        Write-RFLLog -Message "        MSGraph (Beta): deviceManagement/deviceProtectionOverview"
                        ##Todo: use v1.0 version when available
                        $objOut = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/deviceProtectionOverview"
                        #endregion

                        #region Generating Data
		                $TableParams = @{
                            Name = $SectionName
                            List = $true
                            ColumnWidths = 60, 40
                        }
		                $TableParams['Caption'] = "- $($TableParams.Name)"
                        if ($null -ne $objOut -gt 0) {
		                    $objOut | select  @{Name="Pending Signature update";Expression={$_.pendingSignatureUpdateDeviceCount}}, @{Name="Pending full scan";Expression={$_.pendingFullScanDeviceCount}}, @{Name="Pending Quick scan";Expression={$_.pendingQuickScanDeviceCount}}, @{Name="Pending restart";Expression={$_.pendingRestartDeviceCount}}, @{Name="Pending manual steps";Expression={$_.pendingManualStepsDeviceCount}}, @{Name="Pending offline scan";Expression={$_.pendingOfflineScanDeviceCount}}, @{Name="Critical failures";Expression={$_.criticalFailuresDeviceCount}}, @{Name="Inactive agent";Expression={$_.inactiveThreatAgentDeviceCount}}, @{Name="Unknown status";Expression={$_.unknownStateThreatAgentDeviceCount}} | Table @TableParams 
                        } else {
                            Paragraph "No $($sectionName) found"
                        }
                        #endregion
                    }
                    #endregion

                    #region Anti-Virus Policies
                    $SectionName = "Anti-Virus Policies"
                    Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                    Section -Style Heading3 $SectionName {
                        #region Collect Data
                        #endregion

                        #region Generating Data
		                $TableParams = @{
                            Name = $SectionName
                            List = $false
                            Header = 'Platform', 'Name', 'Policy Type', 'Target', 'Last Modified', 'Assignments'
				            Columns = 'platforms', 'name', 'templatereference', 'technologies', 'lastModifiedDateTime', 'Assignments'
                        }
		                $TableParams['Caption'] = "- $($TableParams.Name)"
                        if ($avAssignmentList.count -gt 0) {
		                    $avAssignmentList | select @{Name="platforms";Expression={$platformsList."$($_.platforms)"}}, name, technologies,lastModifiedDateTime, @{Name="templatereference";Expression={$_.templateReference.templateDisplayName}}, Assignments | Table @TableParams
                        } else {
                            Paragraph "No $($sectionName) found"
                        }
                        #endregion

                        #region Anti-Virus Policies Definition
                        foreach($item in $avList.value) {
                            $SectionName = "Policy: $($item.Name)"
                            Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                            Section -Style Heading4 $SectionName {
                                #region Collect Data
                                Write-RFLLog -Message "            MSGraph (Beta): deviceManagement/configurationPolicies/{id}/settings?`$expand=settingDefinitions&top=1000"
                                ##Todo: use v1.0 version when available
                                $polDef = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/configurationPolicies/$($item.id)/settings?`$expand=settingDefinitions&top=1000"
                                ##Todo: use v1.0 version when available
                                $setDef = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/configurationPolicyTemplates/$($item.templateReference.templateId)/settingTemplates?`$expand=settingDefinitions&top=1000"
                    
                                $outobj = @()
                                foreach($polItem in $polDef.Value) {
                                    if ($null -eq $politem.settingInstance.groupSettingCollectionValue.children) {
                                        $setDefItem = $setdef.value.settingdefinitions | Where-Object {$_.id -eq $politem.settingInstance.settingDefinitionId}
                                        $policeDef = New-Object PSObject -Property @{
			                                'SettingName' = $setDefItem.displayName
			                                'SettingMDM' = '{0}{1}' -f $setDefItem.baseUri, $setDefItem.offsetUri
			                                'Value' = 
                                                if (($null -eq ($polItem.settingInstance.choiceSettingValue)) -and ($null -eq $polItem.settingInstance.simpleSettingCollectionValue)) { 
                                                    $polItem.settingInstance.simpleSettingValue.value 
                                                } elseif ($null -eq $polItem.settingInstance.simpleSettingCollectionValue) { 
                                                    $polItem.settingInstance.choiceSettingValue.value.replace("$($polItem.settingInstance.settingDefinitionId)_",'') 
                                                } else {
                                                    $polItem.settingInstance.simpleSettingCollectionValue.value -join ', '
                                                }
                                
		                                }
                                        $outobj += $policeDef
                                    } else {
                                        foreach($polChild in $politem.settingInstance.groupSettingCollectionValue.children) {
                                            $setDefItem = $setdef.value.settingdefinitions | Where-Object {$_.id -eq $polChild.settingDefinitionId}
                                            $policeDef = New-Object PSObject -Property @{
			                                    'SettingName' = $setDefItem.displayName
			                                    'SettingMDM' = '{0}{1}' -f $setDefItem.baseUri, $setDefItem.offsetUri
			                                    'Value' = 
                                                    if (($null -eq ($polItem.settingInstance.choiceSettingValue)) -and ($null -eq $polItem.settingInstance.simpleSettingCollectionValue)) { 
                                                        $polItem.settingInstance.simpleSettingValue.value 
                                                    } elseif ($null -eq $polItem.settingInstance.simpleSettingCollectionValue) { 
                                                        $polItem.settingInstance.choiceSettingValue.value.replace("$($polItem.settingInstance.settingDefinitionId)_",'') 
                                                    } else {
                                                        $polItem.settingInstance.simpleSettingCollectionValue.value -join ', '
                                                    }
                                    
		                                    }
                                            $outobj += $policeDef
                                        }
                                    }
                                }
                                #endregion

                                #region Generating Data
		                        $TableParams = @{
                                    Name = $SectionName
                                    List = $false
                                    ColumnWidths = 30, 50, 20
                                    Header = 'Setting Name', 'MDM Path', 'Value'
				                    Columns = 'SettingName', 'SettingMDM', 'Value'
                                }
		                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                if ($outobj.count -gt 0) {
		                            $outobj | Table @TableParams
                                } else {
                                    Paragraph "No $($sectionName) found"
                                }
                                #endregion
                            }
                        }
                        #endregion
                    }
                    #endregion

                    #region Disk Encryption 
                    $SectionName = "Disk Encryption"
                    Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                    Section -Style Heading3 $SectionName {
                        #region Collect Data
                        $diskencryptionList = @()

                        ##Todo: use v1.0 version when available
                        ##Todo: always return 0, so not using at the moment
                        #Write-RFLLog -Message "        MSGraph (Beta): deviceManagement/configurationPolicies?`$select=id,name,description,platforms,lastModifiedDateTime,technologies,settingCount,roleScopeTagIds,isAssigned,templateReference&`$filter=templateReference/TemplateFamily eq 'endpointSecurityDiskEncryption'"
                        #$diskencryptionListTemplate = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/configurationPolicies?`$select=id,name,description,platforms,lastModifiedDateTime,technologies,settingCount,roleScopeTagIds,isAssigned,templateReference&`$filter=templateReference/TemplateFamily eq 'endpointSecurityDiskEncryption'"
                        #foreach($item in $diskencryptionListTemplate.value) {
                            ##Todo: not in use. mine always return 0
                        #}

                        Write-RFLLog -Message "        MSGraph (Beta): deviceManagement/intents?`$filter=templateId eq 'd1174162-1dd2-4976-affc-6667049ab0ae' or templateId eq 'a239407c-698d-4ef8-b314-e3ae409204b8'"
                        ##Todo: use v1.0 version when available
                        $diskencryptionListIntents = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/intents?`$filter=templateId eq 'd1174162-1dd2-4976-affc-6667049ab0ae' or templateId eq 'a239407c-698d-4ef8-b314-e3ae409204b8'"

                        foreach($item in $diskencryptionListIntents.value) {
                            ##Todo: use v1.0 version when available
                            $assignments = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/intents/$($item.id)/assignments"
                            $groupInfo = @()
                            foreach($assignment in $assignments.Value.target.groupID) {
                                ##Todo: use v1.0 version when available
                                $groupInfo += Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/groups/$($assignment)"
                            }

                            $diskencryptionList += New-Object PSObject -Property @{
                                'name' = $item.displayName
                                'policytype' = $templateIDList."$($item.templateId)"
                                'platforms' = $platformsList."$($item.templateId)"
                                'templateid' = $item.templateId
                                'lastModifiedDateTime' = $item.lastModifiedDateTime
                                'Assignments' = $groupInfo.displayName -join ', '
                                'graphtype' = 'intent'
                                'id' = $item.id
		                    }
                        }
                        #endregion

                        #region Generating Data
		                $TableParams = @{
                            Name = $SectionName
                            List = $false
                            Header = 'Platforms', 'Name', 'Policy Type', 'Last Modified', 'Assignments'
				            Columns = 'platforms', 'name', 'policytype', 'lastModifiedDateTime', 'Assignments'
                        }
		                $TableParams['Caption'] = "- $($TableParams.Name)"
                        if ($diskencryptionList.count -gt 0) {
		                    $diskencryptionList | select platforms, name, policytype, lastModifiedDateTime, Assignments | Table @TableParams
                        } else {
                            Paragraph "No $($sectionName) found"
                        }
                        #endregion

                        #region Disk Encryption Policies Definition
                        foreach($item in $diskencryptionList) {
                            $SectionName = "Policy: $($item.Name)"
                            Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                            Section -Style Heading4 $SectionName {
                                #region Collect Data
                                $outobj = @()
                                if ($item.graphtype -eq 'intent') {
                                    $template = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/templates/$($item.templateid)"
                                    $categories = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/templates/$($item.templateid)/categories"                                    
                                    foreach($catitem in $categories.value) {
                                        $setDef = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/templates/$($item.templateid)/categories/$($catitem.id)/settingDefinitions?"
                                        $poldef = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/intents/$($item.id)/categories/$($catitem.id)/settings?`$expand=Microsoft.Graph.DeviceManagementComplexSettingInstance/Value"

                                        foreach($politem in $poldef.value) {
                                            if ($politem.value -isnot [Array]) {
                                                $setDefItem = $setdef.value | Where-Object {$_.id -eq $politem.definitionId}
                                                $outobj += New-Object PSObject -Property @{
                                                    'SettingLevel' = $catitem.displayName
			                                        'isTopLevel' = $setDefItem.isTopLevel
			                                        'displayName' = $setDefItem.displayName
			                                        'documentationUrl' = $setDefItem.documentationUrl
			                                        'Value' = $polItem.Value
		                                        }
                                            } else {
                                                foreach($polSubitem in $politem.value) {
                                                    if ($polSubitem.value -isnot [Array]) {
                                                        $setDefItem = $setdef.value | Where-Object {$_.id -eq $polSubitem.definitionId}
                                                        $outobj += New-Object PSObject -Property @{
                                                            'SettingLevel' = $catitem.displayName
			                                                'isTopLevel' = $setDefItem.isTopLevel
			                                                'displayName' = $setDefItem.displayName
			                                                'documentationUrl' = $setDefItem.documentationUrl
			                                                'Value' = $polSubitem.Value
		                                                }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion


                                #region Generating Data
		                        $TableParams = @{
                                    Name = $SectionName
                                    List = $false
                                    ColumnWidths = 30, 50, 20
                                    Header = 'SettingLevel', 'displayName', 'Value'
				                    Columns = 'SettingLevel', 'displayName', 'Value'
                                }
		                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                if ($outobj.count -gt 0) {
		                            $outobj | Table @TableParams
                                } else {
                                    Paragraph "No $($sectionName) found"
                                }
                                #endregion
                            }
                        }
                        #endregion
                    }
                    #endregion

                    #region Firewall
                    $SectionName = "Firewall"
                    Write-RFLLog -Message "    Starting SubSection '$($sectionName)'"
                    Section -Style Heading3 $SectionName {
                        #region Collect Data
                        $firewallPolicyList = @()

                        Write-RFLLog -Message "        MSGraph (Beta): deviceManagement/configurationPolicies?`$select=id,name,description,platforms,lastModifiedDateTime,technologies,settingCount,roleScopeTagIds,isAssigned,templateReference&`$filter=templateReference/TemplateFamily%20eq%20%27endpointSecurityFirewall%27"
                        ##Todo: use v1.0 version when available
                        $firewallPolicyListtFirewall = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/configurationPolicies?`$select=id,name,description,platforms,lastModifiedDateTime,technologies,settingCount,roleScopeTagIds,isAssigned,templateReference&`$filter=templateReference/TemplateFamily%20eq%20%27endpointSecurityFirewall%27"
                        foreach($item in $firewallPolicyListtFirewall.value) {
                            ##Todo: use v1.0 version when available
                            $assignments = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/configurationPolicies/$($item.id)/assignments"
                            $groupInfo = @()
                            foreach($assignment in $assignments.Value.target.groupID) {
                                ##Todo: use v1.0 version when available
                                $groupInfo += Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/groups/$($assignment)"
                            }

                            $firewallPolicyList += New-Object PSObject -Property @{
                                'name' = $item.name
                                'id' = $item.id
                                'technologies' = $item.technologies
                                'templateReference' = $item.templateReference
                                'lastModifiedDateTime' = $item.lastModifiedDateTime
                                'platforms' = $platformsList."$($item.platforms)"
                                'templatetype' = $item.templateReference.templateDisplayName
                                'templateid' = $item.templateReference.templateId
                                'Assignments' = $groupInfo.displayName -join ', '
                                'graphtype' = 'firewall'
		                    }
                        }


                        Write-RFLLog -Message "        MSGraph (Beta): deviceManagement/configurationPolicies?select=id,name,description,technologies,platforms,lastModifiedDateTime,settingCount,roleScopeTagIds,creationSource,isAssigned,templateReference&`$filter=(technologies eq 'configManager' and creationSource eq 'Firewall')"
                        ##Todo: use v1.0 version when available
                        $firewallPolicyListConfigMgr = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/configurationPolicies?select=id,name,description,technologies,platforms,lastModifiedDateTime,settingCount,roleScopeTagIds,creationSource,isAssigned,templateReference&`$filter=(technologies eq 'configManager' and creationSource eq 'Firewall')"
                        foreach($item in $firewallPolicyListConfigMgr.value) {
                            ##Todo: use v1.0 version when available
                            $assignments = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/configurationPolicies/$($item.id)/assignments"
                            $groupInfo = @()
                            foreach($assignment in $assignments.Value.target.groupID) {
                                ##Todo: use v1.0 version when available
                                $groupInfo += Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/groups/$($assignment)"
                            }

                            $firewallPolicyList += New-Object PSObject -Property @{
                                'name' = $item.name
                                'id' = $item.id
                                'technologies' = $item.technologies
                                'templateReference' = $item.templateReference
                                'lastModifiedDateTime' = $item.lastModifiedDateTime
                                'platforms' = $platformsList."$($item.platforms)"
                                'templatetype' = 'Microsoft Defender Firewall' #$item.templateReference.templateDisplayName is always null
                                'templateid' = $item.templateReference.templateId
                                'Assignments' = $groupInfo.displayName -join ', '
                                'graphtype' = 'configmgr'
		                    }
                        }

                        Write-RFLLog -Message "        MSGraph (Beta): /deviceManagement/intents?`$filter=templateId eq 'c53e5a9f-2eec-4175-98a1-2b3d38084b91' or templateId eq '4356d05c-a4ab-4a07-9ece-739f7c792910' or templateId eq '5340aa10-47a8-4e67-893f-690984e4d5da'"
                        ##Todo: use v1.0 version when available
                        $firewallPolicyListIntents = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/intents?`$filter=templateId eq 'c53e5a9f-2eec-4175-98a1-2b3d38084b91' or templateId eq '4356d05c-a4ab-4a07-9ece-739f7c792910' or templateId eq '5340aa10-47a8-4e67-893f-690984e4d5da'"
                        foreach($item in $firewallPolicyListIntents.value) {
                            ##Todo: use v1.0 version when available
                            $assignments = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/intents/$($item.id)/assignments"
                            $groupInfo = @()
                            foreach($assignment in $assignments.Value.target.groupID) {
                                ##Todo: use v1.0 version when available
                                $groupInfo += Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/groups/$($assignment)"
                            }

                            $firewallPolicyList += New-Object PSObject -Property @{
                                'name' = $item.displayName
                                'id' = $item.id
                                'technologies' = 'mdm'
                                'templateReference' = $null
                                'lastModifiedDateTime' = $item.lastModifiedDateTime
                                'platforms' = $platformsList."$($item.templateId)"
                                'templatetype' = 'macOS firewall'
                                'templateid' = $item.templateId
                                'Assignments' = $groupInfo.displayName -join ', '
                                'graphtype' = 'mac'
		                    }
                        }
                        #endregion

                        #region Generating Data
		                $TableParams = @{
                            Name = $SectionName
                            List = $false
                            Header = 'Platform', 'Name', 'Policy Type', 'Target', 'Last Modified', 'Assignments'
				            Columns = 'platforms', 'name', 'templatetype', 'technologies', 'lastModifiedDateTime', 'Assignments'
                        }
		                $TableParams['Caption'] = "- $($TableParams.Name)"
                        if ($firewallPolicyList.count -gt 0) {
		                    $firewallPolicyList | select platforms, name, templatetype, technologies, lastModifiedDateTime, Assignments | Table @TableParams
                        } else {
                            Paragraph "No $($sectionName) found"
                        }
                        #endregion

                        #region Firewall Policies Definition
                        foreach($item in $firewallPolicyList) {
                            $SectionName = "Policy: $($item.Name)"
                            Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                            Section -Style Heading4 $SectionName {
                                #region Collect Data
                                $outobj = @()
                                if ($item.graphtype -eq 'mac') {
                                    $template = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/templates/$($item.templateid)"
                                    $categories = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/templates/$($item.templateid)/categories"
                                    foreach($catitem in $categories.value) {
                                        $setDef = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/templates/$($item.templateid)/categories/$($catitem.id)/settingDefinitions?"
                                        $poldef = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/intents/$($item.id)/categories/$($catitem.id)/settings?`$expand=Microsoft.Graph.DeviceManagementComplexSettingInstance/Value"

                                        foreach($politem in $poldef.value) {
                                            if ($politem.value -isnot [Array]) {
                                                $setDefItem = $setdef.value | Where-Object {$_.id -eq $politem.definitionId}
                                                $outobj += New-Object PSObject -Property @{
                                                    'SettingLevel' = $catitem.displayName
			                                        'isTopLevel' = $setDefItem.isTopLevel
			                                        'displayName' = $setDefItem.displayName
			                                        'documentationUrl' = $setDefItem.documentationUrl
			                                        'Value' = $polItem.Value
		                                        }
                                            } else {
                                                foreach($polSubitem in $politem.value) {
                                                    if ($polSubitem.value -isnot [Array]) {
                                                        $setDefItem = $setdef.value | Where-Object {$_.id -eq $polSubitem.definitionId}
                                                        $outobj += New-Object PSObject -Property @{
                                                            'SettingLevel' = $catitem.displayName
			                                                'isTopLevel' = $setDefItem.isTopLevel
			                                                'displayName' = $setDefItem.displayName
			                                                'documentationUrl' = $setDefItem.documentationUrl
			                                                'Value' = $polSubitem.Value
		                                                }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                } elseif ($item.graphtype -eq 'configmgr') {
                                    ##Todo
                                } elseif ($item.graphtype -eq 'firewall') {
                                    Write-RFLLog -Message "            MSGraph (Beta): deviceManagement/configurationPolicies/{id}/settings?`$expand=settingDefinitions&top=1000"
                                    ##Todo: use v1.0 version when available
                                    $polDef = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/configurationPolicies/$($item.id)/settings?`$expand=settingDefinitions&top=1000"
                                    ##Todo: use v1.0 version when available
                                    $setDef = Invoke-MgGraphRequest -Method GET -Uri "$($Global:BetabaseUri)/deviceManagement/configurationPolicyTemplates/$($item.templateReference.templateId)/settingTemplates?`$expand=settingDefinitions&top=1000"
                    
                                    foreach($polItem in $polDef.Value) {
                                        if ($null -eq $politem.settingInstance.groupSettingCollectionValue.children) {
                                            $setDefItem = $setdef.value.settingdefinitions | Where-Object {$_.id -eq $politem.settingInstance.settingDefinitionId}
                                            $policeDef = New-Object PSObject -Property @{
			                                    'displayName' = $setDefItem.displayName
			                                    'SettingLevel' = '{0}{1}' -f $setDefItem.baseUri, $setDefItem.offsetUri
			                                    'Value' = 
                                                    if (($null -eq ($polItem.settingInstance.choiceSettingValue)) -and ($null -eq $polItem.settingInstance.simpleSettingCollectionValue)) { 
                                                        $polItem.settingInstance.simpleSettingValue.value 
                                                    } elseif ($null -eq $polItem.settingInstance.simpleSettingCollectionValue) { 
                                                        $polItem.settingInstance.choiceSettingValue.value.replace("$($polItem.settingInstance.settingDefinitionId)_",'') 
                                                    } else {
                                                        $polItem.settingInstance.simpleSettingCollectionValue.value -join ', '
                                                    }
                                
		                                    }
                                            $outobj += $policeDef
                                        } else {
                                            foreach($polChild in $politem.settingInstance.groupSettingCollectionValue.children) {
                                                $setDefItem = $setdef.value.settingdefinitions | Where-Object {$_.id -eq $polChild.settingDefinitionId}
                                                $policeDef = New-Object PSObject -Property @{
			                                        'displayName' = $setDefItem.displayName
			                                        'SettingLevel' = '{0}{1}' -f $setDefItem.baseUri, $setDefItem.offsetUri
			                                        'Value' = 
                                                        if (($null -eq ($polItem.settingInstance.choiceSettingValue)) -and ($null -eq $polItem.settingInstance.simpleSettingCollectionValue)) { 
                                                            $polItem.settingInstance.simpleSettingValue.value 
                                                        } elseif ($null -eq $polItem.settingInstance.simpleSettingCollectionValue) { 
                                                            $polItem.settingInstance.choiceSettingValue.value.replace("$($polItem.settingInstance.settingDefinitionId)_",'') 
                                                        } else {
                                                            $polItem.settingInstance.simpleSettingCollectionValue.value -join ', '
                                                        }
                                    
		                                        }
                                                $outobj += $policeDef
                                            }
                                        }
                                    }
                                }
                                #endregion


                                #region Generating Data
		                        $TableParams = @{
                                    Name = $SectionName
                                    List = $false
                                    ColumnWidths = 30, 50, 20
                                    Header = 'Name', 'Setting', 'Value'
				                    Columns = 'displayName', 'SettingLevel', 'Value'
                                }
		                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                if ($outobj.count -gt 0) {
		                            $outobj | Table @TableParams
                                } else {
                                    Paragraph "No $($sectionName) found"
                                }
                                #endregion
                            }
                        }
                        #endregion
                    }
                    #endregion
                }
                #endregion

                #region todo:
                <#
                firewall
                endpoint privilege management
                endpoint detection and response
                attack surface reduction
                account protection
                device compliance
                conditional access
                #>
                #endregion

            }
        }
        #endregion

        #region Policies
        $sectionName = 'Policies'
        if ($SectionPolicies -eq $false) {
	        Write-RFLLog -Message "Exporting Section '$($sectionName)' is being ignored as the parameter to export is set to false" -LogLevel 2
        } else {
            PageBreak
	        Write-RFLLog -Message "Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region todo:
                <#
                compliance policies
                conditionl access
                configuration policies
                scripts
                Group policy analytics
                update rings for W10
                feature updates for W10
                quality updates for W10
                update policies for ios/ipados
                update policies for macOS
                #>
                #endregion

            }
        }
        #endregion
    }
    #endregion

    #region Export File
    foreach($OutPutFormatItem in $OutputFormat) {
        Write-RFLLog -Message "Exporting report format $($OutPutFormatItem) to $($OutputFolderPath)"
	    $Document = $Global:WordReport | Export-Document -Path $OutputFolderPath -Format:$OutPutFormatItem -Options @{ TextWidth = 240 } -PassThru
    }
    #endregion

    Write-RFLLog -Message "All Reports ($($OutputFormat.Count)) have been exported to '$($OutputFolderPath)'" -LogLevel 2
    #endregion

    #region Export Errors
    if ($Error.Count -gt 0) {
        Write-RFLLog -Message "Error found when running script"
        foreach($item in $Error) {
            Write-RFLLog -Message $item
        }
    }
    #endregion
} catch {
    Write-RFLLog -Message "An error occurred $($_)" -LogLevel 3
    Exit 3000
} finally {
    Set-Location $Script:CurrentFolder
    Write-RFLLog -Message "*** Ending ***"
}
#endregion