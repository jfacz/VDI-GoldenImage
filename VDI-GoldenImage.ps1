<#PSScriptInfo
.VERSION 2603.0
.GUID d25f8ad6-47fc-4b0c-a239-74ef3be16d2c
.AUTHOR FaraJan
.COMPANYNAME Data Protection Delivery Center s.r.o.
.COPYRIGHT 2026 DPDC CZ. All rights reserved.
.LICENSEURI https://mit-license.org
.TAGS VDI Omnissa Horizon GoldenImage DynamicEnvironmentManager AppVolumes MicrosoftTeams FSLogix OneDrive GoogleDrive OSOT Automation Maintenance VMware
.PROJECTURI https://github.com/jfacz/VDI-GoldenImage
.RELEASENOTES
[v2603.0 - 20260331] Revised and modified code (more structured form and a number of tasks as separate functions)
                     Bug fixes and improvements
[v2510.0 - 20251003] New standalone task SDelete to zero free disk space and reduce virtual disk size. Defrag task by default hidden in script configuration
                     Bug fixes and improvements
[v2504.0 - 20250408] Added task to install/update Horizon Recording Agent
[v2503.0 - 20250331] Full Support for Omnissa installer packages and configuration (version 2412 and later). Support for new FSLogix installer.
                     New function Get-SwRegDetails() for getting details of installed software
                     Bug fixes and improvements
[v2412.0 - 20241216] Removed task to install Microsoft Teams Classic (only New Microsoft Teams)
                     Default startup type of 'Optimize Drives' (defragsvc) service set to 'Manual' (because of FSLogix 'Disk Compaction' feature)
                     Updated to reflect changes in VMware and Omnissa 
                     Bug fixes and improvements
[v2404.0 - 20240430] New option to import configuration from a standalone file 'VDI-GoldenImage.ps1.config' or specify file by 'ConfigFile' parameter
                     Added task to install/update New Microsoft Teams with automatic download of the latest version from the web or offline MSI
                     Added task to install/update Google Drive with automatic download of the latest version from the web
[v2312.0 - 20231220] Added optional param DisableSpoolerRestart to App Volumes
[v2311.1 - 20231130] Bug fixes and improvements
[v2311.0 - 20231107] Added task to install/update Microsoft FSLogix & OneDrive with automatic download of the latest version from the web
[v2310.0 - 20231027] Added task to install/update Horizon Agent, Dynamic Environment Manager and App Volumes
[v2309.1 - 20230926] Added task to install/update Microsoft Teams for VDI with automatic download of the latest version from the web
[v2309.0 - 20230906] Initial version with OS/Office update management, OSOT finalize action, and VM Tools install/update task
#>

<#
.SYNOPSIS
   A robust, programmable maintenance script designed for managing VDI Golden Images (Master PCs) running Omnissa (VMware) Horizon.
.DESCRIPTION
   Script with a basic set of maintenance tasks for a VDI Golden Image running on Omnissa (VMware) Horizon.
   Typical tasks include updating OS, Office and other software as well as finalizing and cleaning up the image before final snapshot and deployment to virtual desktops.
   Other tasks include installing/updating agents used in the VDI environment (Horizon, Dynamic Environment Manager, App Volumes, FSlogix, MS Teams for VDI, OneDrive and Google Drive).
.NOTES
   Version:        2603
   Author:         Fara Jan
   Creation Date:  2023-09-06
   Last Update:    2026-03-31
.PARAMETER Action
   Specifies the action/task of the script. If not specified a menu with a list of actions is displayed.
.PARAMETER ConfigFile
   Specifies a configuration file to load and override script variables (useful for easy updating of this script)
   If not specified the script tries to find and import a file named as script + '.config' e.g. 'VDI-GoldenImage.ps1.config'
.INPUTS
   System.String or Empty
.OUTPUTS
   System.String Console/Log
.EXAMPLE
  # Run script and display menu list of all script actions
     powershell C:\ProgramData\VDI\VDI-GoldenImage.ps1
.EXAMPLE
  # Enable and run OS/Office/SW updates in Golden Image
     PS> VDI-GoldenImage.ps1 -Action Update
.EXAMPLE
  # Disable OS/Office/SW updates, run system clean-up tasks and prepare the Golden Image for Horizon
     PS> VDI-GoldenImage.ps1 -Action Finalize
.EXAMPLE
  # Install/update VM Tools in Golden Image - comparing current version and version in InstallSrcDir
     PS> VDI-GoldenImage.ps1 -Action VmTools
.EXAMPLE
  # Install/update Microsoft Teams for VDI - install the latest Microsoft Teams from an online or offline MSIX package
     PS> VDI-GoldenImage.ps1 -Action MsTeams
.NOTES
   Recommended locations for scripts & source installer packages: C:\ProgramData\VDI & C:\ProgramData\VDI\Install
#>


#----------------[ Script param ]----------------
param (
    [Parameter(Mandatory=$false)]
        [ValidateSet("Update", "Finalize", "FinalizeFast", "SDelete", "Defrag", "VmTools", "Horizon", "DEM", "AppVolumes", "HorizonRec", "FSLogix", "MsTeams", "OneDrive", "GoogleDrive", "CfgInfo", "Exit")]
        [string] $Action,
    [Parameter(Mandatory=$false)]
            [string] $ConfigFile
)

#----------------[ Settings ]----------------
$VAR = @{
 # Script Name
  ScriptName = "VDI Golden Image Maintenance [v2603]"
 # Script Path ($PSScriptRoot for CurrentPath)
  ScriptPath = $PSScriptRoot
 # Script Menu Title
  ScriptMenuTitle = "Please select a script action"

 # --- VDI Image - Updates Settings ---
 # Windows/Office (365/2019) updates managed/controlled by this script - disabled by default & run only during Golden Image maintenance/update
  ManageWindowsUpdates = $true
  ManageOfficeUpdates = $true
 # Other updates Settings (Web browser and other SW updates)
  ManageMsEdgeUpdate = $true
  ManageGoogleUpdates = $true
  ManageAdobeUpdates = $true
  ManageOneDriveUpdates = $true

 # --- Finalize Settings (Windows OS Optimalization Tool) ---
  # OSOT info web: https://techzone.omnissa.com/resource/windows-os-optimization-tool-horizon-guide
    # Path to the OSOT executable file (automatic getting of latest OSOT version)
   OsotPath = "C:\Program Files\OSOT"
  # Finalize Argument Settings (https://docs.omnissa.com/bundle/Optimizing-Images-for-Horizon/page/RunWindowsOSOptimizationToolforHorizonfromCommandLine.html)
  # Details/Help of Finalize Argument: optimization-tool.exe -h
  OsotFinalizeArg = "-v -f 0 1 2 3 4 5 7 9 10 11" #All excepts: '(6) Clears Default user profile', '(8) Creates local group policies'
   #OsotFinalizeArg = "-v -f 3 4" #demo
  # Finalize Argument Settings for Fast cleanup
   OsotFinalizeArgFast = "-v -f 3 4 9 10 11" # Disk Cleanup, EventLog Cleanup, Clears KMS Settings, Flush DNS cache, Releases IP Address
  # Shutdown VM after Finalize
   OsotShutdownAfterFinalize = $true
   OsotShutdownAfterFinalizeFast = $true
  # SCCM: Clear Configuration during finalize (eliminate duplicates)
   SccmClearConfig = $false

 # --- VM Tools Settings ---
  # Enable Carbon Black Helper (CBHelper)
  # Set to $true if you are using VMware Carbon Black Cloud for Endpoint Detection and Response (EDR) or Antivirus.
   VmToolsCarbonBlack = $false
  # Enable NSX Guest Introspection Drivers (NetworkIntrospection, FileIntrospection)
  # Set to $true if using NSX Agentless Antivirus (offloaded scans) or NSX Identity Firewall (user-based network micro-segmentation).
   VmToolsNsxIntrospection = $false

 # --- Horizon Agent Install options ---
  # Install Options DOC: https://docs.omnissa.com/bundle/Desktops-and-Applications-in-HorizonV2512/page/MicrosoftWindowsInstallerCommandLineOptions.html
  # Horizon Agent Features, if HorizonAgentAddLocal = ALL then HorizonAgentRemove is also used
   HorizonAgentAddLocal = "ALL" #"Core,PCoIP,USB,NGVC,RTAV,ClientDriveRedirection,GEOREDIR,V4V,VmwVaudio,VmwVidd,TSMMR,BlastUDP,PerfTracker,HelpDesk,PrintRedir,PSG"
   HorizonAgentRemove = "SerialPortRedirection,ScannerRedirection,SmartCard,SdoSensor"
  # Delete Horizon Peformance Tracker Icon from all users desktop
   DelPerfTrackerDesktopIcon = $true
  # Increasing a default timeout limit for Post-Sync/ClonePrep Customization Scripts during the customization phase of Instant Clones (default 20000 ms = 20 s)
   HorizonAgentExecScriptTimeout = 90000

 # --- DEM (Dynamic Environment Manager) Config share Path (Empty if no DEM) ---
  DemConfigPath = "\\domain.int\VDI$\DEMConfig\general"

 # --- App Volumes Agent Configuration ---
  # AppVol Manager(s) (DNS hostname/IP) - multiple array items for HA configuration; Empty Array if No App Volumes
   AppVolManager = @("vdi-avm01.domail.local", "vdi-avm02.domail.local")
  # Enforce SSL Certificate Validation - Set to False if the Golden Image does not trust the App Volumes Manager's certificate yet.
   AppVolEnforceSSLVal = $false
  # Deactivate Restarting the Spooler Service When Using Integrated Printing: https://docs.omnissa.com/bundle/AppVolumesAdminGuideV2512/page/DeactivateRestartingtheSpoolerServiceWhenUsingIntegratedPrinting.html
   AppVolDisableSpoolerRestart = $true
  # The maximum wait for a response from the AppVol Manager, in seconds. If set to 0, the wait for response is forever (default 300 sec = 5 min)
   AppVolMaxDelayTimeOutS = 30

 # --- Horizon Recording Agent Settings ---
  # FQDN (or) IP of Recording server (or) Load balancer. Note: Starts with https:// and ends with port number 9443.
   HorizonRecServerAdressProp = ""
  # Thumbprint of the recording server (no spaces or no colons).
   HorizonRecTrustedThumbprint = ""
  # Machine Template Switch True / False
   HorizonRecMachineIsTemplate = "True"
  # Credentrial for register machine to Recording server (script will promt for password)
   HorizonRecUser = "Administrator"

 # --- Microsoft Teams Settings ---
  # Install MS Teams from offline MSIX (stored in InstallSrcDir)
   MsTeamsOfflineMSIX = $false
   MsTeamsDisableAutoUpdate = $true
   MsTeamsDisableAutoStart = $false

 # --- Google Drive Settings ---
  GoogleDriveDesktopShortcuts = $false
  GoogleDriveGSuiteShortcuts = $true

 # --- Installer Settings ---
  # Script SW install Source Directory for agents (VM Tools, Horizon/DEM/AppVol Agent, MS Teams, ...)
   InstallSrcDir = "Install"  # relative path to $_.ScriptPath
  # Ask user before installing a newer version of any agent software
   InstallerAskIfNewer = $true  

 # --- Other Settings ---
  # Enable SDelete Action (Zero Free Space). Recommended for Thin Provisioned disks on VMFS/NFS.
   ActionSDeleteEnabled = $true
  # Enable Defrag Action/task. Not recommended for modern All-Flash storage or SSDs.
   ActionDefragEnabled = $false
  # Show details of install params
   ShowInstallParams = $false

 # --- LOG settings ---
  LogDir = "Logs" # relative path to $_.ScriptPath
  LogFileName = "VDI-GI-Maintenance-{0:yyyyMMdd_HHmmss}.txt"
  LogArchiveFiles = 14
}

# --- Settings for SW Install/Update (SW Name, SW Source Install File Name, Installation Log File Name, Ask for install of newer version) ---
$SwSet = @{
    VmTools =     @{Name = "VMware Tools";                SrcFile = "VMware-tools-*64.exe";                 LogFile = "Log_VmTools_install_{0:yyyyMMdd_HHmmss}.txt"}
    Horizon =     @{Name = "Horizon Agent";               SrcFile = "*Horizon-Agent-x86_64-*.exe";          LogFile = "Log_HorizonAgent_install_{0:yyyyMMdd_HHmmss}.txt"}
    DEM =         @{Name = "Dynamic Environment Manager"; SrcFile = "*Dynamic*Environment*Manager*x64.msi"; LogFile = "Log_DEMAgent_install_{0:yyyyMMdd_HHmmss}.txt"}
    AppVolumes =  @{Name = "App Volumes Agent";           SrcFile = "*AppVolumes*Agent*.msi";               LogFile = "Log_AppVolAgent_install_{0:yyyyMMdd_HHmmss}.txt"}
    FSLogix =     @{Name = "Microsoft FSLogix Apps";      SrcFile = "FSLogix_*.zip";                        LogFile = "Log_MsFSLogix_install_{0:yyyyMMdd_HHmmss}.txt"}
    MsTeams =     @{Name = "Microsoft Teams for VDI";     SrcFile = "teamsbootstrapper.exe";                LogFile = "Log_MsTeams_install_{0:yyyyMMdd_HHmmss}.txt"}
    OneDrive =    @{Name = "Microsoft OneDrive";          SrcFile = "OneDriveSetup.exe";                    LogFile = ""}
    GoogleDrive = @{Name = "Google Drive";                SrcFile = "GoogleDriveSetup.exe";                 LogFile = ""}
    HorizonRec =  @{Name = "Horizon Recording Agent";     SrcFile = "HorizonRecordingAgent-*.exe";          LogFile = "Log_HorizonRecAgent_install_{0:yyyyMMdd_HHmmss}.txt"}
}

# --------
Clear-Host
# --------
#----------------[ Initilization Functions ]------------------
# Script Environment Initilization/Setting 
function InitEnvConfig{
    param(
        [Parameter(Mandatory=$true)] [hashtable] $VarSet
    )

    $regArr     = @()
    $svcArr     = @()
    $schtaskArr = @()

    # Add Registry helper fce
    function AddItemReg($Path, $Name, $Type, $ValOn, $ValOff = $null, [switch]$DeleteOff){
        return [pscustomobject]@{ Path=$Path; Name=$Name; PropertyType=$Type; ValueOn=$ValOn; ValueOff=$ValOff; DeleteOff = $DeleteOff.IsPresent }
    }
    # Add Service helper fce
    function AddItemSvc($Name, $StartOn, $ActionOn, $StartOff, $ActionOff) {
        return [pscustomobject]@{ Name=$Name; StartupTypeOn=$StartOn; SvcActionOn=$ActionOn; StartupTypeOff=$StartOff; SvcActionOff=$ActionOff }
    }

    # Windows Update Registry Settings
    if($VarSet.ManageWindowsUpdates){
        $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
        $regArr += AddItemReg $regPath "BranchReadinessLevel" "Dword" 16 32
        $regArr += AddItemReg $regPath "DeferQualityUpdatesPeriodInDays" "Dword" 0 30
        $regArr += AddItemReg $regPath "DisableWindowsUpdateAccess" "Dword" 0 1
        $regArr += AddItemReg $regPath "DoNotConnectToWindowsUpdateInternetLocations" "Dword" 0 1
        $regArr += AddItemReg $regPath "PauseFeatureUpdatesStartTime" "String" $("{0:yyyy-12-31}" -f (Get-Date)) "2021-10-15" 
        $regArr += AddItemReg $regPath "PauseQualityUpdatesStartTime" "String" "2015-12-31" "2021-10-31"
        $regArr += AddItemReg $regPath "SetDisableUXWUAccess" "Dword" 0 1
        $regPath += "\AU"
        $regArr += AddItemReg $regPath "NoAutoUpdate" "Dword" 0 1
        $regArr += AddItemReg $regPath "AllowMUUpdateService" "Dword" 1 -DeleteOff
        $regArr += AddItemReg $regPath "AUOptions" "Dword" 4 -DeleteOff
        $regArr += AddItemReg $regPath "ScheduledInstallDay" "Dword" 0 -DeleteOff
        $regArr += AddItemReg $regPath "ScheduledInstallEveryWeek" "Dword" 1 -DeleteOff
        $regArr += AddItemReg $regPath "ScheduledInstallTime" "Dword" 3 -DeleteOff
    }
    # Office Update Registry Settings
    if($VarSet.ManageOfficeUpdates){
        $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate"
        $regArr += AddItemReg $regPath "EnableAutomaticUpdates" "Dword" 1 0
        $regArr += AddItemReg $regPath "HideEnableDisableUpdates" "Dword" 1 0
    }
    # Teams Update Registry Settings
    if($VarSet.MsTeamsDisableAutoUpdate){
        $regPath = "HKLM:\SOFTWARE\Microsoft\Teams"
        $regArr += AddItemReg $regPath "disableAutoUpdate" "Dword" 0 1
    }

    # Windows Service Settings
    if($VarSet.ManageWindowsUpdates){
        $svcArr += AddItemSvc "wuauserv" "Manual" "Start" "Disabled" "Stop"
        $svcArr += AddItemSvc "UsoSvc" "Manual" "Start" "Disabled" "Stop"
        $svcArr += AddItemSvc "StorSvc" "Automatic" "Start" "Manual" "Stop"
        # AppVolumes services
        if($VarSet.AppVolManager.Count -gt 0){
            $svcArr += AddItemSvc "svservice" "Disabled" "Stop" "Automatic" "Start"
            $svcArr += AddItemSvc "svdriver" "Disabled" "Stop" "Automatic" "Start"
        }
    }
    # MsEdge Update services
    if($VarSet.ManageMsEdgeUpdate){
        $svcArr += AddItemSvc "edgeupdate" "Manual" "None" "Disabled" "Stop"
        $svcArr += AddItemSvc "edgeupdatem" "Manual" "None" "Disabled" "Stop"
    }
    # Google Update services
    if($VarSet.ManageGoogleUpdates){
        $svcArr += AddItemSvc "GoogleUpdaterService*" "Manual" "None" "Disabled" "Stop"
        $svcArr += AddItemSvc "GoogleUpdaterInternalService*" "Manual" "None" "Disabled" "Stop"
        $svcArr += AddItemSvc "gupdate" "Manual" "None" "Disabled" "Stop"
        $svcArr += AddItemSvc "gupdatem" "Manual" "None" "Disabled" "Stop"
    }
    # Adobe Update services
    if($VarSet.ManageAdobeUpdates){
        $svcArr += AddItemSvc "AdobeARMservice" "Manual" "Start" "Disabled" "Stop"
    }

    # Scheduled Tasks settings
    if($VarSet.ManageAdobeUpdates){ $schtaskArr += "Adobe Acrobat Update Task" }
    if($VarSet.ManageOneDriveUpdates){ $schtaskArr += "OneDrive Per-Machine Standalone Update Task" }
    if($VarSet.ManageMsEdgeUpdate){
        $schtaskArr += "MicrosoftEdgeUpdateTaskMachineCore*"
        $schtaskArr += "MicrosoftEdgeUpdateTaskMachineUA*"
    }
    if($VarSet.ManageGoogleUpdates){
        $schtaskArr += "GoogleUpdateTaskMachineCore*"
        $schtaskArr += "GoogleUpdateTaskMachineUA*"
    }

    # Return the results as a structured object
    return [pscustomobject]@{
        Registry = $regArr
        Services = $svcArr
        Schtasks = $schtaskArr
    }
}

#----------------[ Declarations ]----------------
$InstallSrcDir = Join-Path $VAR.ScriptPath $VAR.InstallSrcDir
$LogDir = Join-Path $VAR.ScriptPath $VAR.LogDir
$LogFile = (Join-Path $LogDir $VAR.LogFileName) -f (Get-Date)

#----------------[ Initilization ]----------------
# Set environment based on the defined configuration
$EnvCfg = InitEnvConfig -VarSet $VAR
# Install source, Log folder and OSOT
if(!(Test-Path $InstallSrcDir -PathType Container)){ New-Item -Path $InstallSrcDir -ItemType Directory | out-null }
if(!(Test-Path $LogDir -PathType Container)){ New-Item -Path $LogDir -ItemType Directory | out-null }
if(!(Test-Path $VAR.OsotPath -PathType Container)){ New-Item -Path $VAR.OsotPath -ItemType Directory | out-null }

#----------------[ Logging ]----------------
# Transcript Log - Start
if($Host.Name -match "ConsoleHost"){ Start-Transcript -Path $LogFile -append }

#----------------[ Functions ]------------------
# Script Messages Function
# Displays a formatted message in the console with color-coding and optionally appends it to the global HTML EmailBody variable.
# Supports message types (info, warn, error, success, etc.), optional blank lines before/after, and stripping of HTML tags for console output.
Function MsgFce{
    param (
        [Parameter(Position=0, Mandatory=$True)] [string] $Msg,
        [ValidateSet("info", "warn", "error", "success", "verbose", "note", "header", "return")] [string] $Output="info",
        [int] $LinesBefore=0,
        [int] $LinesAfter=0,
        [switch] $StripHtml,
        [switch] $NoAddToEmailBody
    )

    $Color = switch($Output){ 
        "warn" {"Yellow"}; "error" {"Red"}; "success" {"Green"}; "verbose" {"Cyan"}; "note" {"DarkGray"}; "header" {"DarkYellow"} 
    }
    # EmailBody helper
    function AddEmailBody([string]$content){
            if(($Script:VAR.EmailBody -is [array]) -and !$NoAddToEmailBody){ $Script:VAR.EmailBody += $content}
    }

    # Padding: Empty Lines Before
    if($LinesBefore){ 1..$LinesBefore | ForEach-Object{ Write-Host ""; AddEmailBody "<br />" } }
    # Msg
    $Msg = if($StripHtml){ $Msg -replace "<[^>]*?>" } else{ $Msg }
    if($Output -eq "return"){ return "`r`n$($Msg)" }

    # Write-Host Msg
    $WriteHostArgs = @{ Object = $Msg }
    if($Color){ $WriteHostArgs.ForegroundColor = $Color }
    if($Output -eq "header"){
        $border = "-" * ($Msg.Length + 4)
        $WriteHostArgs.Object = "$($border)`n  $($Msg.ToUpper())`n$($border)"
        Write-Host @WriteHostArgs
        AddEmailBody "<p><b>$($Msg)</b></p>"
    } else{
        Write-Host @WriteHostArgs
        AddEmailBody $Msg
    }
    # Padding: Empty Lines After
    if($LinesAfter){ 1..$LinesAfter | ForEach-Object { Write-Host ""; AddEmailBody "<br />" } }
}

# Simple choice menu FCE (Writes an output of array items to select)
# Example of use: $MenuItems = @("Yes", "No"); $Title = "Contine?"
function MenuSimple{
    Param(
        [Parameter(Position=0, Mandatory=$True)] [string[]] $MenuItems,
        [string] $Title,
        [boolean] $Cls
    )

    $header = $null
    if(![string]::IsNullOrWhiteSpace($Title)){
        $len = [math]::Max(($MenuItems | Measure-Object -Maximum -Property Length).Maximum, $Title.Length)
        $header = "{0}{1}{2}" -f $Title, [Environment]::NewLine, ("-" * $len)
    }
    # menu items and space align if more than 9 items
    $len = if($MenuItems.Count -gt 9){ 2 } else{ 1 }
    $items = ($MenuItems | ForEach-Object{ "[{0}]{1}{2}" -f ++$i, $(if($i -lt 10){" " * $len} else{" "}), $_ }) -join [Environment]::NewLine

    # display the menu and return the chosen option
    while($true){
        if($Cls){ Clear-Host } else{ Write-Host }
        if($header){ Write-Host $header -ForegroundColor Yellow }
        Write-Host $items
        Write-Host

        $index = (Read-Host -Prompt 'Please make your choice')
        $index = $index -as [int]

        if((1..$MenuItems.Count) -contains $index){
            return $MenuItems[$index-1]
        } else{
            Write-Warning "Invalid choice.. Please try again."
            Start-Sleep -Seconds 2
        }
    }
}

# Helper function get action name from script menu
function GetMenuAction{
    Param(
        [Parameter(Position=0, Mandatory=$True)] [string] $MenuSel
    )
    return $($MenuSel.Split(":")[0])
}

# Registry settings FCE for enable/disable updates (enable = ValueOn / disable = ValueOff)
# Example of param: [pscustomobject] @{Path = ""; Name = ""; PropertyType = "Dword/String/..."; ValueOn = ; ValueOff = }
Function SetRegistry{
    param (
        [Parameter(Mandatory=$true)] [array] $RegArr,
        [Parameter(Mandatory=$true)] [ValidateSet("On","Off")] [string] $UpdateAction
    )

    $regValueKey = "Value$($UpdateAction)"
    foreach($reg in $RegArr){
        try{
            $currentVal = $reg.$regValueKey
            # delete
            if($UpdateAction -eq "Off" -and $reg.DeleteOff){
                if(Get-ItemProperty -Path $reg.Path -Name $reg.Name -ErrorAction SilentlyContinue){
                    MsgFce "Registry '$($reg.Name)' => Removing from '$($reg.Path)'"
                    Remove-ItemProperty -Path $reg.Path -Name $reg.Name -Force -ErrorAction Stop
                } else{
                    MsgFce "Registry '$($reg.Name)' doesn't exist at '$($reg.Path)'"
                }
            } else{
                if($null -ne $currentVal){
                    if(!(Test-Path $reg.Path)){ New-Item -Path $reg.Path -Force -ErrorAction Stop | Out-Null }
                    MsgFce "Registry '$($reg.Name)' ($($reg.PropertyType)) => Setting to '$currentVal' at '$($reg.Path)'"
                    New-ItemProperty -Path $reg.Path -Name $reg.Name -PropertyType $reg.PropertyType -Value $currentVal -Force -ErrorAction Stop | Out-Null
                }
            }
        } catch{
            MsgFce "ERROR: Failed to process registry '$($reg.Name)' at '$($reg.Path)'. Reason: $($_.Exception.Message)" -Output Error
        }
    }
}

# Windows service settings FCE for enable/disable updates (enable = StartupTypeOn | SvcActionOn  / disable = StartupTypeOff | SvcActionOff)
# Example of param: [pscustomobject] @{Name = ""; StartupTypeOn = "Manual/Automatic/Disabled"; SvcActionOn = "Start/Stop"; StartupTypeOff = "Manual/Automatic/Disabled"; SvcActionOff = "Start/Stop" }
Function SetService{
    param (
        [Parameter(Mandatory=$true)] [array] $SvcArr,
        [Parameter(Mandatory=$true)] [ValidateSet("On","Off")] [string] $UpdateAction
    )

    $svcStartupTypeKey = "StartupType$($UpdateAction)"
    $svcActionKey = "SvcAction$($UpdateAction)"
    foreach($svc in $SvcArr){
        try{
            $service = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
            if($service){
                $targetStartup = $svc.$svcStartupTypeKey
                $targetAction  = $svc.$svcActionKey
                # Update Startup Type
                MsgFce "Service '$($svc.Name)' ($($service.DisplayName)) => Startup type set to '$targetStartup'"
                $service | Set-Service -StartupType $targetStartup -ErrorAction Stop
                # Execute Service Action
                switch ($targetAction){
                    "Start"{
                        MsgFce "Service '$($svc.Name)' => Starting"
                        $service | Start-Service -ErrorAction Stop
                    } "Stop" {
                        MsgFce "Service '$($svc.Name)' => Stopping"
                        $service | Stop-Service -Force -ErrorAction Stop
                    } "None" {
                        # No action required (skip)
                    } default {
                        MsgFce "WARN: Unknown service action '$targetAction' for '$($svc.Name)'" -Output warn
                    }
                }
            } else{
                MsgFce "WARN: Service '$($svc.Name)' doesn't exist => Skipping" -Output warn
            }
        } catch{
            MsgFce "ERROR: Failed to configure service '$($svc.Name)'. Reason: $($_.Exception.Message)" -Output Error
        }
    }
}

# Windows Task Scheduler - Enabling/Disabling schtask ($Schtask = Scheduled Task Name mask)
Function SetSchtaskState{
    param (
        [Parameter(Mandatory=$true)] [string] $Schtask,
        [Parameter(Mandatory=$true)] [ValidateSet("Enable","Disable")] [string] $State
    )

    try{
        # Find tasks and suppress error if none are found
        $tasks = Get-ScheduledTask -TaskName $Schtask -ErrorAction SilentlyContinue
        $taskCount = @($tasks).Count # Using @() ensures we always have a .Count property
        if($taskCount -eq 1){
            $currentTask = $tasks[0] # Take the first and only task
            $taskName = $currentTask.TaskName
            # Perform action
            if($State -eq "Enable"){
                MsgFce "Scheduled task '$taskName' => Setting state to 'Enabled'"
                $currentTask | Enable-ScheduledTask -ErrorAction Stop | Out-Null
            } else{
                MsgFce "Scheduled task '$taskName' => Setting state to 'Disabled'"
                $currentTask | Disable-ScheduledTask -ErrorAction Stop | Out-Null
            }
        } elseif($taskCount -eq 0){
            MsgFce "WARN: Scheduled task '$Schtask' not found (Skip)" -Output warn
        } else{
            # Safety check: Multiple tasks found
            MsgFce "WARN: Scheduled task '$Schtask' match $taskCount items. Use a more specific name to ensure a unique match." -Output warn
        }
    } catch{
        MsgFce "ERROR: Failed to update task '$Schtask'. Reason: $($_.Exception.Message)" -Output Error
    }
}

# Getting latest version of "Windows OS Optimalization Tool" and call the OSOT exe with specified arguments
Function OsotCmd{
    param (
        [Parameter(Mandatory=$true)] [string] $OsotArg,
        [Parameter(Mandatory=$false)] [boolean] $OsotShutdown = $false
    )

    # Getting of latest version OSOT Executable
    $osotFileFilter = "*HorizonOSOptimizationTool-x86_64-*"
    $osotExe = Get-ChildItem -Path $Script:VAR.OsotPath -Filter $osotFileFilter -ErrorAction SilentlyContinue | Sort-Object { [version]$_.VersionInfo.FileVersion } | Select-Object -Last 1
    if($osotExe){
        # Prepare arguments
        $finalArgs = $OsotArg
        if($OsotShutdown){ $finalArgs += " -shutdown" }
        MsgFce "OSOT version: $($osotExe.VersionInfo.FileVersion)"
        MsgFce "OSOT Command: $($osotExe.Name) $finalArgs"
        try{
            $process = Start-Process -FilePath $osotExe.FullName -ArgumentList $finalArgs -Wait -NoNewWindow -PassThru -ErrorAction Stop
            if($process.ExitCode -ne 0){
                MsgFce "WARN: OSOT finished with ExitCode $($process.ExitCode)" -Output warn
            } else{
                MsgFce "OSOT completed successfully."
            }
        } catch{
            MsgFce "ERROR: OSOT execution failed: $($_.Exception.Message)" -Output Error
        }
    } else{
        MsgFce "WARN: No OSOT executable found in path ($($Script:VAR.OsotPath)) with filter ($osotFileFilter)" -Output warn
    }
}

# Getting software normalized version
function Get-SwNormalizedVersion {
    param (
        [Parameter(Mandatory=$true)] [string] $RawString
    )

    # Extract all numeric candidates (e.g., "8.16.0", "2506")
    $candidates = [regex]::Matches($RawString, '(\d+(\.\d+){0,3})') | ForEach-Object { $_.Value }
    $parsedVersions = foreach($c in $candidates){
        # Split into segments and check for Int32 overflow (max 2147483647)
        $parts = $c.Split('.')
        $validParts = @()
        foreach($p in $parts){
            if([int64]$p -le 2147483647){ $validParts += $p } else{ break }
        }
        if($validParts.Count -eq 0){ continue }
        # Reconstruct and normalize (YYMM -> YYMM.0)
        $vString = $validParts -join '.'
        if($vString -match '^\d+$'){ $vString += ".0" }
        try { [version]$vString } catch { $null }
    }
    # Prioritization Logic:
    # We prefer "Technical versions" (3 or 4 segments, e.g., 8.16.0) over "Marketing versions" (1 or 2 segments, e.g., 2506.0).
    $technical = $parsedVersions | Where-Object { $_.ToString().Split('.').Count -ge 3 }
    if($technical){
        return ($technical | Sort-Object -Descending | Select-Object -First 1)
    } else{
        return ($parsedVersions | Sort-Object -Descending | Select-Object -First 1)
    }
}

# Getting Installed software details from uninstall registry
function Get-SwRegDetails{
    param (
        [Parameter(Position=0, Mandatory=$true)] [string] $DisplayName,
        [ValidateSet("both","64bit","32bit")] [string] $regPath = "both"
    )

    # Installed software registry keys by SW architecture
    $keys = @()
    if($regPath -in ("both","64bit")){ $keys += "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" }
    if($regPath -in ("both","32bit")){ $keys += "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" }    

    # Search for the software
    $sw = Get-ItemProperty $keys -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like $DisplayName }
    if(!$sw){ return }

    # Process results
    $results = foreach($app in $sw){
        $sysVer = Get-SwNormalizedVersion -RawString $app.DisplayVersion
        [pscustomobject] @{ DisplayName = $app.DisplayName; DisplayVersion = $app.DisplayVersion; SystemVersion = $sysVer; PSPath = $app.PSPath; UninstallString = $app.UninstallString }
    }
    return $results | Sort-Object SystemVersion | Select-Object -Last 1
}

# Extracts metadata and property information from an MSI installer file
Function Get-MsiInformation{
    param (
        [Parameter(Position=0, Mandatory=$true)] [ValidateNotNullOrEmpty()]
        [System.IO.FileInfo[]] $MSI
    )

    if($MSI.Count -ne 1){
        Write-Warning "ERROR: Only ONE MSI file for fce Get-MsiInformation!"
        return
    }
    try{
        $file = Get-ChildItem $MSI -ErrorAction Stop
    } catch{
        Write-Warning "Unable to get file $($MSI) $($_.Exception.Message)"
        return
    }

    $dataMSI = [ordered] @{
        FileName         = $file.Name
        FileFullName     = $file.FullName
        "FileLength(MB)" = [math]::Round($file.Length / 1MB, 2)
    }

    # Read property from MSI database
    $winInstaller = New-Object -ComObject WindowsInstaller.Installer
    # open MSI DB read only
    $msiDbOpenReadOnly = 0
    try{
        $msiDb = $winInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $winInstaller, @($file.FullName, $msiDbOpenReadOnly))
    } catch{
        Write-Debug $_.Exception.Message
    }
    if($msiDb){
        $properties = @("ProductName", "ProductVersion", "Manufacturer", "ProductCode", "UpgradeCode")
        foreach($property in $properties){
            $query = "SELECT Value FROM Property WHERE Property = '$($property)'"
            $view = $msiDb.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $msiDb, ($query))
            $view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null)
            $record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)
            try{
                $value = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 1)
            } catch{
                Write-Debug "Unable to get '$property' $($_.Exception.Message)"
                $value = ""
            }
            $dataMSI.$property = $value
        }

        # Other MSI Details
        # https://docs.microsoft.com/en-us/windows/win32/msi/summary-information-stream-property-set
        $msiInfo = $msiDb.GetType().InvokeMember("SummaryInformation", "GetProperty",$null , $msiDb, $null)
        $propertyIndex = @{"Title"=2; "Subject"=3; "Author"=4; "Comment"=6; "CreationDate"=12; "RevisionNumber"=9;"ApplicationName"=18}
        foreach($key in $propertyIndex.Keys){
            $dataMSI.$key = $msiInfo.Property($propertyIndex.$key)
        }
    }
    # Close MSI Run garbage collection and release ComObject
    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($winInstaller)
    [System.GC]::Collect()

    return [pscustomobject] $dataMSI
}

# Function of clearing SCCM Configuration (Resets the SCCM/MECM client identity before system cloning)
Function SccmClearConfig{
    try{
        # Stop SCCM Service
        $svcName = "ccmexec"
        $service = Get-Service -Name $svcName -ErrorAction SilentlyContinue
        if($service){
            if($service.Status -ne 'Stopped'){
                MsgFce "Stopping service '$($svcName)'..."
                Stop-Service -Name $svcName -Force -ErrorAction Stop
            }
        } else{
            MsgFce "WARN: Service '$($svcName)' not found (Skipping stop)" -Output warn
        }
        # Remove SMSCFG.ini (The primary identity file)
        $sccmCfg = Join-Path $Env:WinDir "SMSCFG.ini"
        if(Test-Path $sccmCfg){
            MsgFce "Removing SCCM config file: $($sccmCfg)"
            Remove-Item -Path $sccmCfg -Force
        }
        # Clear SCCM Certificates (Deleting the entire SMS store is safer for cloning than filtering by subject)
        $smsCertPath = "Cert:\LocalMachine\SMS"
        if(Test-Path $smsCertPath){
            MsgFce "Clearing SCCM certificates in $($smsCertPath)"
            Get-ChildItem -Path $smsCertPath | Remove-Item -ErrorAction SilentlyContinue
        }
        # Clear WMI Identity (Avoids duplicate IDs in SCCM console)
        $wmiPath = "root\ccm\invagt"
        MsgFce "Clearing SCCM identity in WMI ($($wmiPath))"
        Get-WmiObject -Namespace $wmiPath -Class InventoryActionStatus -ErrorAction SilentlyContinue | Remove-WmiObject
        # Clear SCCM Cache
        $cachePath = Join-Path $Env:WinDir "ccmcache"
        if(Test-Path $cachePath){
            MsgFce "Clearing SCCM cache: $($cachePath)"
            Remove-Item "$cachePath\*" -Recurse -Force -ErrorAction SilentlyContinue
        }
        MsgFce "SCCM client cleanup completed successfully."
    } catch {
        MsgFce "ERROR: Failed to clear SCCM configuration: $($_.Exception.Message)" -Output Error
    }
}

#----------------[ Invoke Functions for script Actions ]---------------
# VDI Universal function to find, compare, and install/update VDI software components
function Invoke-VdiSoftwareInstall{
    param(
        [hashtable] $SwSettings,
        [string] $ExtraArgs = "",
        [scriptblock] $PreInstallAction = $null,
        [scriptblock] $PostInstallAction = $null,
        [ValidateSet("auto","fslogix","teams","others")] [string] $InstallType = "auto"
    )

    process{
        # Find the latest installation package
        $swSrcFile = Get-ChildItem -Path $script:InstallSrcDir -Filter $SwSettings.SrcFile | Sort-Object LastWriteTime, Name | Select-Object -Last 1
        if(!$swSrcFile){
            MsgFce "ERROR: Installation package for '$($SwSettings.Name)' (mask: $($SwSettings.SrcFile)) NOT found in '$($script:InstallSrcDir)'." -Output error
            return
        }

        # Determine source version string (MSI vs EXE)
        $sourceVersionRaw = if($swSrcFile.Extension -eq ".msi"){ 
            (Get-MsiInformation $swSrcFile.FullName).ProductVersion 
        } else{ 
            $swSrcFile.VersionInfo.FileVersion 
        }
        $sourceVersion = Get-SwNormalizedVersion -RawString "$($swSrcFile.Name) $sourceVersionRaw"
        # Safety check: Exit if version couldn't be retrieved
        if(!$sourceVersion){
            MsgFce "ERROR: Could not retrieve a valid version from $($swSrcFile.Name)" -Output error
            return
        }

        # Check current installation
        $swInstalled = Get-SwRegDetails -DisplayName "*$($SwSettings.Name)*"
        # Version comparison and installation
        if(!$swInstalled -or ($swInstalled.SystemVersion -lt $sourceVersion)){
            if($InstallType -ne "teams"){
                MsgFce "INFO: A newer version of '$($SwSettings.Name)' was found in the installation source path ('$($script:InstallSrcDir)')"
                $currentVersion = if($swInstalled){ $swInstalled.SystemVersion } else{ "Not installed" }
                MsgFce "Current version: $($currentVersion)"
                MsgFce "Source version:  $($sourceVersion)"
            } else{
                MsgFce "INFO: Installer for '$($SwSettings.Name)' is ready (version: $($sourceVersion))"
            }
            $continue = if($script:VAR.InstallerAskIfNewer){ MenuSimple -MenuItems "Yes","No" -Title "Continue the installation?" } else{ "Yes" }
            if($continue -eq "Yes"){
                if($PreInstallAction){
                    MsgFce "INFO: Running pre-installation tasks..."
                    $dynamicArgs = & $PreInstallAction
                    if($dynamicArgs){ $ExtraArgs = "$ExtraArgs $dynamicArgs".Trim() }
                }

                MsgFce "INFO: Starting installation of '$($SwSettings.Name)'..."
                $logFile = (Join-Path $script:LogDir $SwSettings.LogFile) -f (Get-Date)
                # Build Start-Process execution/command based on file type
                $startArgs = @{FilePath = ""; ArgumentList = ""; Wait = $true; ErrorAction = "Stop"}
                if($swSrcFile.Extension -eq ".msi"){
                    $startArgs.FilePath = "msiexec.exe"
                    $startArgs.ArgumentList = "/i ""$($swSrcFile.FullName)"" /qb /l* ""$logFile"" $ExtraArgs"
                }else{
                    $startArgs.FilePath = $swSrcFile.FullName
                    if($InstallType -in ("fslogix")){
                        $startArgs.ArgumentList = "$($ExtraArgs) /log ""$($logFile)"""
                    } elseif($InstallType -in ("teams","others")){
                        $startArgs.ArgumentList = $ExtraArgs
                        #$startArgs.RedirectStandardOutput = $logFile
                        $startArgs.NoNewWindow = $true
                    } else{
                        $startArgs.ArgumentList = "/S /v ""/qb /l* ""$logFile"" $ExtraArgs"""
                    }
                }
                if($script:VAR.ShowInstallParams){ MsgFce "DEBUG: Executing: $($startArgs.FilePath) $($startArgs.ArgumentList)" -Output note }
                try{
                    Start-Process @startArgs
                    MsgFce "INFO: $($SwSettings.Name) installation finished."
                    if($PostInstallAction){
                        MsgFce "INFO: Running post-installation tasks..."
                        & $PostInstallAction $swSrcFile
                    }
                } catch{
                    MsgFce "ERROR: Installation failed: $($_.Exception.Message)" -Output error
                }
            } else{
                MsgFce "INFO: The installation has been cancelled"
            }
        } else{
            MsgFce "INFO: '$($SwSettings.Name)' is already up to date" -Output success
            MsgFce "Current version: $($swInstalled.SystemVersion)"
            MsgFce "Source version:  $($sourceVersion)"
        }
    }
}

# Performs post-installation configuration for Horizon Agent (Removes the Performance Tracker shortcut from the public desktop and other registry configuration)
function Invoke-HorizonPostConfig{
    param(
        [hashtable] $VarSet = $script:VAR
    )

    process {
        # Remove PerfTracker icon from public desktop
        if($VarSet.DelPerfTrackerDesktopIcon){
            $IconPath = "$($env:PUBLIC)\Desktop\*Horizon Performance Tracker.lnk"
            if(Test-Path -Path $IconPath){
                MsgFce "INFO: Removing Horizon Performance Tracker icon from public desktop"
                Remove-Item -Path $IconPath -Force -ErrorAction SilentlyContinue | Out-Null
            }
        }

        # Configure ExecScriptTimeout registry
        if([int] $VarSet.HorizonAgentExecScriptTimeout -gt 2000){
            MsgFce "INFO: Setting a new Timeout Limit for ClonePrep Customization Scripts (default limit: 2000 ms => new limit: $($VAR.HorizonAgentExecScriptTimeout) ms)."
            $RegValName = "ExecScriptTimeout"
            $RegPropertyType = "Dword"
            $RegBase = "HKLM:\SYSTEM\CurrentControlSet\Services"
            
            # List of possible service names for both Omnissa/Vmware
            $TargetServices = @("omn-instantclone-ga", "vmware-viewcomposer-ga")
            $foundAny = $false
            foreach($service in $TargetServices){
                $RegPath = Join-Path $RegBase $service
                if(Test-Path -Path $RegPath){
                    $foundAny = $true
                    MsgFce "Registry '$($RegValName)' ($($RegPropertyType)) at '$($RegPath)'"
                    try{
                        New-ItemProperty -Path $RegPath -Name $RegValName -Value $VarSet.HorizonAgentExecScriptTimeout -PropertyType $RegPropertyType -Force -ErrorAction Stop | Out-Null
                    } catch{
                        MsgFce "ERROR: Failed to set '$($RegValName)' at '$($RegPath)': $($_.Exception.Message)" -Output error
                    }
                }
            }
            if(!$foundAny){
                MsgFce "WARN: No Horizon Agent service registry found to configure timeout." -Output warn
            }
        }
    }
}

# Performs post-installation configuration for App Volumes Agent (Configures additional App Volumes Managers for HA and other registry configuration)
function Invoke-AppVolPostConfig{
    param(
        [hashtable] $VarSet = $script:VAR
    )

    process{
        MsgFce "INFO: App Volumes Agent Post-Configuration"
        $RegPath = "HKLM:\SYSTEM\CurrentControlSet\Services\svservice\Parameters"
        if(!(Test-Path -Path $RegPath)){
            MsgFce "WARN: App Volumes service registry path not found ($($RegPath))" -Output warn
            return
        }
        # Configure additional Managers (Manager2, Manager3, ...) if provided
        # The primary manager is already set during MSI installation
        if($VarSet.AppVolManager.Count -gt 1){
            for($i=1; $i -lt $VarSet.AppVolManager.Count; $i++){
                $RegValName = "Manager$($i + 1)"
                $RegValData = "$($VarSet.AppVolManager[$i]):443"
                MsgFce "Configuring additional App Volumes Manager: '$($RegValData)' in '$($RegValName)'"
                New-ItemProperty -Path $RegPath -Name $RegValName -Value $RegValData -PropertyType String -Force | Out-Null
            }
        }
        # Next AppVol Configuration
        # Configure EnforceSSLCertificateValidation registry
        $RegValName = "EnforceSSLCertificateValidation"
        $RegVal = if($VarSet.AppVolEnforceSSLVal){ 1 } else{ 0 }
        MsgFce "AppVolumes registry settings for SSL validation ('$($RegValName)' = '$($RegVal)')"
        New-ItemProperty -Path $RegPath -Name $RegValName -Value $RegVal -PropertyType Dword -Force | Out-Null

        # Deactivate Spooler Service Restart
        if($VarSet.AppVolDisableSpoolerRestart){
            $RegValName = "DisableSpoolerRestart"
            $RegVal = 1
            MsgFce "AppVolumes registry settings for deactivating restarting of the spooler service ('$($RegValName)' = '$($RegVal)')"
            New-ItemProperty -Path $RegPath -Name $RegValName -Value $RegVal -PropertyType Dword -Force | Out-Null
        }
        # Configure Manager Connection Timeout (override the default timeout of 300 seconds if different)
        if($VarSet.AppVolMaxDelayTimeOutS -ne 300){
            $RegValName = "MaxDelayTimeOutS"
            $RegVal = $VarSet.AppVolMaxDelayTimeOutS
            MsgFce "AppVolumes registry settings for timeout to connect App Volumes Manager ('$($RegValName)' = '$($RegVal)')"
            New-ItemProperty -Path $RegPath -Name $RegValName -Value $RegVal -PropertyType Dword -Force | Out-Null
        }
    }
}

# Sdelete: Optimizes virtual disk size by zeroing out free space using Sysinternals SDelete
function Invoke-SDelete{
    param(
            [string] $InstallSrcDir = $Script:InstallSrcDir,
            [string] $OsotPath = $Script:VAR.OsotPath
    )

    process{
        $SDeleteFileName = "sdelete64.exe"
        $SDeletePath = Join-Path -Path $OsotPath -ChildPath $SDeleteFileName
        # Check if SDelete utility exists in the OSOT directory
        if(!(Test-Path -Path $SDeletePath)){
            MsgFce "WARN: SDelete utility ('$($SDeleteFileName)') was not found in OSOT path ('$($OsotPath)')" -Output warn
            $sel = MenuSimple -MenuItems "Yes", "No" -Title "Download SDelete from the web and save it to the OSOT directory?"
            if($sel -eq "Yes"){
                try{
                    $DownloadUrl = "https://download.sysinternals.com/files/SDelete.zip"
                    MsgFce "...downloading the latest SDelete utility (link: $($DownloadUrl)) and extracting '$($SDeleteFileName)' to OSOT directory" -Output note
                    
                    $ZipDest = Join-Path -Path $InstallSrcDir -ChildPath "SDelete.zip"
                    Invoke-WebRequest -Uri $DownloadUrl -OutFile $ZipDest -ErrorAction Stop

                    # Extract sdelete64.exe directly to the OSOT directory
                    Add-Type -Assembly System.IO.Compression.FileSystem
                    $ZipFile = [IO.Compression.ZipFile]::OpenRead($ZipDest)
                    $ZipFile.Entries | Where-Object {$_.FullName -eq $SDeleteFileName} | ForEach-Object{ [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, $SDeletePath, $true) }
                    $ZipFile.Dispose()
                    
                    MsgFce "INFO: SDelete has been successfully extracted to '$($OsotPath)'"
                }catch{
                    MsgFce "ERROR: Failed to download or extract SDelete: $($_.Exception.Message)" -Output error
                    return
                }
            }
        }
        # Execution SDelete
        if(Test-Path -Path $SDeletePath){
            $Arguments = @("-z", $env:SystemDrive)
            MsgFce "Starting '$($SDeletePath)' with parameters: $($Arguments -join ' ')"
            MsgFce "Zeroing free space to shrink the virtual disk. This may take several minutes..."
            try{
                Start-Process -FilePath $SDeletePath -ArgumentList $Arguments -Wait -NoNewWindow -ErrorAction Stop
                MsgFce "INFO: SDelete process finished successfully."
            }catch{
                MsgFce "ERROR: SDelete execution failed: $($_.Exception.Message)" -Output error
            }
        }else{
            MsgFce "INFO: SDelete utility still was not found in system => action was skipped."
        }
    }
}

# Defragmentation: Performs disk defragmentation on the system drive
function Invoke-Defrag{
    process{
        # Define service configuration for defragsvc
        $SvcDefrag = @([pscustomobject] @{Name = "defragsvc"; StartupTypeOn = "Manual"; SvcActionOn = "Start"; StartupTypeOff = "Manual"; SvcActionOff = "Stop" })
        # Enable and start the defrag service using existing SetService function
        SetService -SvcArr $SvcDefrag -UpdateAction On
        MsgFce "Running defragmentation on the system drive ($env:SystemDrive)..."
        try{
            $Arguments = @($env:SystemDrive, "/U", "/V") # Arguments: /U (Progress), /V (Verbose)
            # Start defrag and wait for completion
            Start-Process -FilePath "defrag.exe" -ArgumentList $Arguments -Wait -NoNewWindow -ErrorAction Stop
            MsgFce "INFO: Defragmentation finished successfully."
        }catch{
            MsgFce "ERROR: Defragmentation process failed: $($_.Exception.Message)" -Output error
        }
        # Stop and set service back to original state
        SetService -SvcArr $SvcDefrag -UpdateAction Off
    }
}

#----------------[ Main Execution ]---------------
MsgFce $VAR.ScriptName -Output Header
$msg = "The beginning of the script [{0:yyyy-MM-dd HH:mm:ss}]" -f (Get-Date)
MsgFce $msg -Output note

# UAC required
if(!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")){
    MsgFce "ERROR: Administrator rights are required to run this script!" -Output error
    break
}

# PowerShell responds settings for faster Invoke-WebRequest downloads (without progress)
$ProgressPreference = "SilentlyContinue"

# --- Load/Import script configuration file for redefine script variables defined in $VAR (useful for easy updating of this script) --- 
$cfgInfo = "" # string with details of configuraction details (listable in the Script Menu) 
# Configuration file for import (defined by param $ConfigFile or default file is same as script + '.config' e.g. 'VDI-GoldenImage.ps1.config'
$CfgFile2Import = (Split-Path -Path $PSCommandPath -Leaf) -replace ".ps1", ".ps1.config"
if($ConfigFile -notlike ""){ $CfgFile2Import = Split-Path -Path $ConfigFile -Leaf }
$CfgFile2Import = Join-Path $VAR.ScriptPath  $CfgFile2Import
# Load configuration file and import variables
if(Test-Path $CfgFile2Import -PathType Leaf){
    $cfgInfo += MsgFce "INFO: Configuration file '$($CfgFile2Import)' used to redefine `$VAR values" -Output return -LinesAfter 1
    $lineNum = 0
    Get-Content $CfgFile2Import | ForEach-Object {
        $lineNum++
        $line = $_.Trim()
        $lineNumF = "{0:D3}" -f $lineNum
        if($line -eq "" -or $line.StartsWith("#")){ return } # Skip empty and comments
        $key, $val = $line.Split("=", 2).Trim()
        if($key -notin $VAR.Keys){
            $cfgInfo += MsgFce "Line {0:D3}: Unknown variable '$key'" -f $lineNum -Output return
            return
        }
        # Detect target type - enhanced to catch all array types
        $currentVal = $VAR.$key
        $targetType = if($null -eq $currentVal){ "String" } 
                      elseif($currentVal -is [array]){ "Array" } 
                      else{ $currentVal.GetType().Name }

        $val = $val.Trim("'").Trim('"') # Remove quotes
        # --- Dynamic Type Conversion ---
        try{
            $parsedVal = switch -Wildcard ($targetType){
                "Boolean" {
                    $cleanVal = $val.TrimStart("$")
                    if($cleanVal -iin @("true", "1")){ $true }
                    elseif($cleanVal -iin @("false", "0")){ $false }
                    else{ throw "Invalid boolean value: $val" } 
                }
                "Int*" { [int] $val }
                "Array" {
                    $cleanVal = $val -replace '^\@\(', '' -replace '\)$', ''
                    if([string]::IsNullOrWhiteSpace($cleanVal)){
                        @()
                    } else{
                        $cleanVal -split "," | ForEach-Object { $_.Trim().Trim("'").Trim('"') }
                    }                    
                }
                default { $val } # String and others
            }
            # FIX: Force array type if the target was an array
            if($targetType -eq "Array"){ $VAR.$key = @($parsedVal) } else { $VAR.$key = $parsedVal }
            $cfgInfo += MsgFce "Line $($lineNumF): Variable '$($key)' set to '$($VAR.$key)'" -Output return
        } catch{
            $cfgInfo += MsgFce "Line $($lineNumF): ERROR - Cannot set '$($key)' to '$($val)' (Expected $($targetType))" -Output return
        }
    }
    # Reload environment based on redefined configuration
    $EnvCfg = InitEnvConfig -VarSet $VAR
} else{
    $cfgInfo += MsgFce "INFO: Configuration file '$($CfgFile2Import)' not found => using script defaults in '`$VAR'" -Output return
}

# --- Script Main Menu ---
# Prepare labels/text for some menu items
$updStr = @()
if($VAR.ManageWindowsUpdates){ $updStr += "Windows" }
if($VAR.ManageOfficeUpdates){ $updStr += "Office" }
$menuStrUpdate = if($updStr){ $updStr -join " and " } else{ "System" }
$menuStrFinalize = if($menuStrUpdate){ "Disable $menuStrUpdate updates & " } else{ "" }
$menuStrDownload = " (Internet download available)"

# --- Build Menu using a tiny local helper ---
$MENU = @()
$MenuAdd = { param($Label, $Condition=$true) if($Condition){ $script:MENU += $Label } }
& $MenuAdd "Update: Enable & run updates for $menuStrUpdate"
& $MenuAdd "Finalize: $($menuStrFinalize)Run system cleanup tasks and prepare Golden Image for Enrollment"
& $MenuAdd "FinalizeFast: Run cleanup tasks without NGEN optimization, DISM cleanup and without CompactOS (faster)"
& $MenuAdd "SDelete: Zero free space (virtual disk optimization)" $VAR.ActionSDeleteEnabled
& $MenuAdd "Defrag: Disk defragmentation (system drive)" $VAR.ActionDefragEnabled
& $MenuAdd "VmTools: Install/Update VMware Tools"
& $MenuAdd "Horizon: Install/Update Horizon Agent"
& $MenuAdd "DEM: Install/Update Dynamic Environment Agent" $VAR.DemConfigPath
& $MenuAdd "AppVolumes: Install/Update AppVolumes Agent" $VAR.AppVolManager.Count
& $MenuAdd "HorizonRec: Install/Update Horizon Recording Agent" $VAR.HorizonRecServerAdressProp
& $MenuAdd "FSLogix: Install/Update Microsoft FSLogix$menuStrDownload"
& $MenuAdd "MsTeams: Install/Update Microsoft Teams for VDI [MSIX]$menuStrDownload"
& $MenuAdd "OneDrive: Install/Update Microsoft OneDrive$menuStrDownload"
& $MenuAdd "GoogleDrive: Install/Update Google Drive$menuStrDownload"
& $MenuAdd "CfgInfo: Show script configuration details"
& $MenuAdd "Exit"

# Display menu if no Action specified ---
if($Action -eq ""){
    $menuSel = MenuSimple -MenuItems $MENU -Title $VAR.ScriptMenuTitle
    $Action = GetMenuAction $menuSel
}

# --- Show configuration details and again show Script Menu ---
While($Action -like "CfgInfo"){
    MsgFce "Script Configuration Details" -Output verbose -LinesBefore 1
    $cfgInfo
    MsgFce "INFO: List of all variables:" -LinesBefore 1
    $VAR.GetEnumerator() | Sort-Object Name | Format-Table
    # Script Menu
    $menuSel = MenuSimple -MenuItems $MENU -Title $VAR.ScriptMenuTitle
    $Action = GetMenuAction $menuSel
}

# Param/Action info
MsgFce "Script 'Action': $($Action)" -LinesBefore 1 -LinesAfter 1 -Output note

# if switch by Action param
# --- Enables & run Windows/Office updates ---
if($Action -like "Update"){
    # Set Update Registry
    if($EnvCfg.Registry.Count){
        $str += if($VAR.ManageWindowsUpdates -and $VAR.ManageOfficeUpdates){ "Windows and Office"} elseif($VAR.ManageWindowsUpdates){ "Windows" } else{ "Office" }
        MsgFce "INFO: Enable $($str) Update registry & run" -Output verbose
        SetRegistry -RegArr $EnvCfg.Registry -UpdateAction On
    }
    # Set Update Services
    if($EnvCfg.Services.Count){
        MsgFce "INFO: Enable updating services" -Output verbose
        SetService -SvcArr $EnvCfg.Services -UpdateAction On
    }

    # Run GPupdate
    if($EnvCfg.Registry.Count -or $EnvCfg.Services.Count){
        & gpupdate
    }

    # Start Windows OS Update GUI
    if($VAR.ManageWindowsUpdates){
        MsgFce "INFO: Starting Windows Update process..."
        Start-Process "ms-settings:windowsupdate"
    }

    # Start Office Update GUI
    if($VAR.ManageOfficeUpdates){
        $msoClientExe = "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe"
        $msoClientArg = "/update user"
        if(Test-Path $msoClientExe -PathType Leaf){
            MsgFce "INFO: Starting Microsoft Office update process..."
            Start-Process $msoClientExe $msoClientArg
        } else{
            MsgFce "WARN: Microsoft Office update client not found ('$($msoClientExe)') => skipped" -Output warn
        }
    }
}

# --- OSOT Finalize ---
elseif($Action -like "Finalize"){
    # Disable Update Services & Registry
    if($EnvCfg.Services.Count){
        MsgFce "INFO: Disable updating services" -Output verbose
        SetService -SvcArr $EnvCfg.Services -UpdateAction Off
    }

    # Set Update Registry
    if($EnvCfg.Registry.Count){
        $str += if($VAR.ManageWindowsUpdates -and $VAR.ManageOfficeUpdates){ "Windows and Office"} elseif($VAR.ManageWindowsUpdates){ "Windows" } else{ "Office" }
        MsgFce "INFO: Disable $($str) update registry" -Output verbose
        SetRegistry -RegArr $EnvCfg.Registry -UpdateAction Off
    }

    # Run GPupdate
    if($EnvCfg.Registry.Count -or $EnvCfg.Services.Count){
        & gpupdate
    }

    # Other maintenance - clear old windows updates files, log files etc.
    $dirArr = @()
    $dirArr += [pscustomobject] @{Path = "C:\ProgramData\VMware\VDM\Logs\"; FileMask = "*.*"; OlderThanXDays = 0 }
    $dirArr += [pscustomobject] @{Path = "C:\ProgramData\Omnissa\VDM\Logs\"; FileMask = "*.*"; OlderThanXDays = 0 }
    $dirArr += [pscustomobject] @{Path = "C:\Program Files (x86)\CloudVolumes\Agent\Logs"; FileMask = "*.log"; OlderThanXDays = 0 }

    if($dirArr.Count){
        MsgFce "INFO: Cleaning up some other directories" -Output verbose
        $dirSoftDistDownloads = "C:\Windows\SoftwareDistribution\Download"
        MsgFce "INFO: Directory: '$($dirSoftDistDownloads)' => Complete deletion of this directory"
        Remove-Item "$($dirSoftDistDownloads)\*" -Recurse -Force -ErrorAction SilentlyContinue
        foreach($dir in $dirArr){
            # select items to delete
            $items = Get-ChildItem -Path $dir.Path -Recurse -force -ErrorAction SilentlyContinue | Where-Object {-not $_.PsIsContainer -and ($_.Name -ilike $dir.FileMask)}
            if($dir.OlderThanXDays){ $items = $items | Where-Object {($_.LastwriteTime -lt (Get-Date).AddDays(-$dir.OlderThanXDays))} }
            if($items.Count){ $items | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue }
            # info msg
            $str = if($dir.OlderThanXDays){ ", older than $($dir.OlderThanXDays) day(s)"} else{ "" }
            $str += " => Total number of deleted files: " + $items.Count
            MsgFce "INFO: Directory: '$($dir.Path)', File Mask: '$($dir.FileMask)'$($str)"
        }
    }

    # Set/Disable Scheduled tasks
    if($EnvCfg.Schtasks.Count){
        MsgFce "INFO: Disabling some scheduled tasks" -Output verbose
        foreach($schtask in $EnvCfg.Schtasks){
            SetSchtaskState -Schtask $schtask -State Disable
        }
    }

    # SCCM Configuration Cleanup
    if($VAR.SccmClearConfig){
        MsgFce "INFO: Clearing SCCM Configuration" -Output verbose
        SccmClearConfig
    }

    # OSOT Finalize Cmd
    MsgFce "INFO: 'Windows OS Optimalization Tool' Finalize" -Output verbose
    OsotCmd -OsotArg $VAR.OsotFinalizeArg -OsotShutdown $VAR.OsotShutdownAfterFinalize
}

# --- OSOT Finalize fast ---
elseif($Action -like "FinalizeFast"){
    # SCCM Configuration Cleanup
    if($VAR.SccmClearConfig){
        MsgFce "INFO: Clearing SCCM Configuration" -Output verbose
        SccmClearConfig
    }

    # OSOT Finalize Cmd
    MsgFce "INFO: 'Windows OS Optimalization Tool' Finalize FAST" -Output verbose
    OsotCmd -OsotArg $VAR.OsotFinalizeArgFast -OsotShutdown $VAR.OsotShutdownAfterFinalizeFast
}

# --- SDelete (Zero free space) ---
elseif($Action -like "SDelete"){
    MsgFce "INFO: Action: SDelete (Zero Free Space)" -Output verbose
    Invoke-SDelete
}

# --- Disk defragmentation ---
elseif($Action -like "Defrag"){
    MsgFce "INFO: Action: Disk Defragmentation" -Output verbose
    Invoke-Defrag
}

# --- VMware Tools ---
elseif($Action -like "VmTools"){
    MsgFce "INFO: Action: Install/Update $($SwSet.$Action.Name)" -Output verbose
    # Prepare the components removal list based on VAR settings
    $vmToolsRemove = "VmwTimeProvider,ServiceDiscovery,VSS" # best practices for VDI
    if(!$VAR.VmToolsCarbonBlack){ $vmToolsRemove += ",CBHelper" }
    if(!$VAR.VmToolsNsxIntrospection){ $vmToolsRemove += ",NetworkIntrospection,FileIntrospection" }
    # Define ExtraArgs
    $ExtraArgs = "REBOOT=R ADDLOCAL=ALL REMOVE=$($vmToolsRemove)"
    # Invoke SW install
    Invoke-VdiSoftwareInstall -SwSettings $SwSet.VmTools -ExtraArgs $ExtraArgs
    
}

# --- Horizon Agent ---
elseif($Action -like "Horizon"){
    MsgFce "INFO: Action: Install/Update $($SwSet.$Action.Name)" -Output verbose
    # Define ExtraArgs
    $ExtraArgs = "VDM_VC_MANAGED_AGENT=1 SUPPRESS_RUNONCE_CHECK=1 ADDLOCAL=$($VAR.HorizonAgentAddLocal)"
    if($VAR.HorizonAgentRemove -ne ""){ $ExtraArgs += " REMOVE=$($VAR.HorizonAgentRemove)" }
    $ExtraArgs += " REBOOT=ReallySuppress"
    # Horizon agent post-installation configuration
    $PostInstall = { Invoke-HorizonPostConfig }
    # Invoke SW install
    Invoke-VdiSoftwareInstall -SwSettings $SwSet.Horizon -ExtraArgs $ExtraArgs -PostInstallAction $PostInstall
}

# --- Dynamic Environment Manager (DEM) Agent ---
elseif($Action -like "DEM"){
    MsgFce "INFO: Action: Install/Update $($SwSet.$Action.Name)" -Output verbose
    # DEM Args if the DEM configuration share path is defined
    $ExtraArgs = if($VAR.DemConfigPath -ne ""){ "COMPENVCONFIGFILEPATH='$($VAR.DemConfigPath)'" } else{""}
    # Invoke SW install
    Invoke-VdiSoftwareInstall -SwSettings $SwSet.DEM -ExtraArgs $ExtraArgs
}

# --- AppVolumes Agent ---
elseif($Action -like "AppVolumes"){
    MsgFce "INFO: Action: Install/Update $($SwSet.$Action.Name)" -Output verbose
    # Check if at least one manager is defined
    if($VAR.AppVolManager.Count -gt 0 -and $VAR.AppVolManager[0] -ne ""){
        $EnforceSSL = if($VAR.AppVolEnforceSSLVal){ 1 } else{ 0 }
        # Instalation options with primary manager and CertSSL
        $ExtraArgs = "MANAGER_ADDR=$($VAR.AppVolManager[0]) MANAGER_PORT=443 EnforceSSLCertificateValidation=$($EnforceSSL) REBOOT=ReallySuppress"
        # Post-installation AppVolumes configuration (Registry settings)
        $PostInstall = { Invoke-AppVolPostConfig }
        # Invoke SW install
        Invoke-VdiSoftwareInstall -SwSettings $SwSet.AppVolumes -ExtraArgs $ExtraArgs -PostInstallAction $PostInstall
    }else{
        MsgFce "ERROR: No App Volumes Manager defined in configuration (VAR.AppVolManager)" -Output warn
    }
}

# --- Horizon Recording Agent ---
elseif($Action -like "HorizonRec"){
    MsgFce "INFO: Action: Install/Update $($SwSet.$Action.Name)" -Output verbose
    if($script:VAR.HorizonRecServerAdressProp -match "^https:\/\/[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}:\d+$"){
        $PreInstall = {
            $Passw = Read-Host "Please enter password for recording user '$($script:VAR.HorizonRecUser)'" -AsSecureString
            $Passw = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Passw))
            $Extra = 'USER_PASSWORD="{0}"' -f $Passw
            return $Extra  # string appended to $ExtraArgs
        }
        $ExtraArgs = 'SERVERADDRESSPROP="{0}" TRUSTEDTHUMBPRINT="{1}" MACHINEISTEMPLATE="{2}" USER_NAME="{3}"' -f $VAR.HorizonRecServerAdressProp, $VAR.HorizonRecTrustedThumbprint, $VAR.HorizonRecMachineIsTemplate, $VAR.HorizonRecUser
        # Invoke SW install
        Invoke-VdiSoftwareInstall -SwSettings $SwSet.HorizonRec -ExtraArgs $ExtraArgs -PreInstallAction $PreInstall
    } else{
        MsgFce "ERROR: Config param 'HorizonRecServerAdressProp' ($($script:VAR.HorizonRecServerAdressProp)) is invalid => cannot continue with installation" -Output error
        MsgFce "INFO: Expected format: https://servername.domain.local:port"
    }
} 

# --- Microsoft FSLogix Agent ---
elseif($Action -like "FSLogix"){
    MsgFce "INFO: Action: Install/Update $($SwSet.$Action.Name)" -Output verbose
    # Download Installer?
    $sel = MenuSimple -MenuItems "Yes","No" -Title "Download the latest version of the '$($SwSet.$Action.Name)' installer from the web?"
    if($sel -eq "Yes"){
        $swDownloadUrl = "https://aka.ms/fslogix_download"
        MsgFce "...downloading the latest '$($SwSet.$Action.Name)' installation package (link: $($swDownloadUrl))" -Output note        
        # download header for getting fileName
        $downHead = Invoke-WebRequest -UseBasicParsing -Method Head $swDownloadUrl
        $destFile = Join-Path $InstallSrcDir $downHead.BaseResponse.ResponseUri.Segments[-1]
        Invoke-WebRequest $swDownloadUrl -OutFile $destFile
    }
    # Find the latest ZIP package in Install source
    $zipFile = Get-ChildItem -Path $InstallSrcDir -Filter $SwSet.$Action.SrcFile | Sort-Object LastWriteTime | Select-Object -Last 1
    if($zipFile){
        # Unzip FSLogix Agent Installer
        $FSLogixAppsSetup = "FSLogixAppsSetup.exe"
        $zipPathInside = "*x64/Release/$($FSLogixAppsSetup)"
        MsgFce "INFO: Extracting '$($FSLogixAppsSetup)' from ZIP package..."
        Add-Type -Assembly System.IO.Compression.FileSystem
        $zip = [IO.Compression.ZipFile]::OpenRead($zipFile.FullName)
        $zip.Entries | Where-Object { $_.FullName -like $zipPathInside } | ForEach-Object {[System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, "$($InstallSrcDir)\$($_.Name)", $true)}
        $zip.Dispose()
        # SwSet Clone and modify for FSLogix exe installer
        $SwSetFsl = $SwSet.FSLogix.Clone()
        $SwSetFsl.SrcFile = $FSLogixAppsSetup
        # PostInstall Cleanup - Delete the extracted EXE after the installation is finished
        $PostInstall = { 
            param($file) 
            MsgFce "INFO: Cleaning up extracted installer file..."
            Remove-Item $file.FullName -Force -ErrorAction SilentlyContinue
        }
        # Invoke SW install
        Invoke-VdiSoftwareInstall -SwSettings $SwSetFsl -ExtraArgs "/install /passive /norestart" -PostInstallAction $PostInstall -InstallType fslogix
    } else{
        MsgFce "ERROR: No FSLogix ZIP package (mask: $($SwSet.FSLogix.SrcFile)) found in '$InstallSrcDir'" -Output error
    }
}

# --- Microsoft Teams for VDI (MSIX) ---
elseif($Action -like "MsTeams"){
    MsgFce "INFO: Action: Install/Update $($SwSet.$Action.Name)" -Output verbose
    # Download Installer?
    $sel = MenuSimple -MenuItems "Yes","No" -Title "Download the latest version of the '$($SwSet.$Action.Name)' installer/bootstrapper from the web?"
    if($sel -eq "Yes"){
        $swDownloadUrl = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409"
        MsgFce "...downloading the latest '$($SwSet.$Action.Name)' installation Bootstrapper (link: $($swDownloadUrl))" -Output note                
        $destFile = Join-Path $InstallSrcDir $SwSet.$Action.SrcFile
        Invoke-WebRequest $swDownloadUrl -OutFile $destFile
    }
    $ExtraArgs = "-p"
    # Offline MSIX logic
    if($VAR.MsTeamsOfflineMSIX){
        $filter = "MSTeams-x64.msix"
        $msixFile = Get-ChildItem -Path $InstallSrcDir -Filter $filter -ErrorAction SilentlyContinue | Select-Object -Property Name, FullName, LastWriteTime
        if($msixFile){
            MsgFce "INFO: Offline MSIX package for '$($SwSet.$Action.Name)' exists in install source path ($($InstallSrcDir)). FileName: $($msixFile.Name); LastWriteTime: $($msixFile.LastWriteTime)" -Output note
            $ExtraArgs += " -o ""$($msixFile.FullName)"""
        } else{
            MsgFce "WARN: Offline MSIX '$($filter)' not found, bootstrapper will use online latest version from web." -Output warn
        }
    }
    # Post-installation detection of AppX version
    $PostInstall = { 
        $appx = Get-AppxPackage | Where-Object { $_.Name -like "MSTeams" } | Select-Object Name, Version
        if($appx){ MsgFce "INFO: Installed MS Teams AppX version: $($appx.Version)" }
    }
    # Invoke SW install
    Invoke-VdiSoftwareInstall -SwSettings $SwSet.$Action -ExtraArgs $ExtraArgs -PostInstallAction $PostInstall -InstallType teams
}

# --- Microsoft OneDrive (All Users) ---
elseif($Action -like "OneDrive"){
    MsgFce "INFO: Install/Update $($SwSet.$Action.Name) (All Users)" -Output verbose
    # Download Installer?
    $sel = MenuSimple -MenuItems "Yes","No" -Title "Download the latest version of the '$($SwSet.$Action.Name)' installer from the web?"
    if($sel -eq "Yes"){
        $swDownloadUrl = "https://go.microsoft.com/fwlink/?linkid=844652"
        MsgFce "...downloading the latest '$($SwSet.$Action.Name)' installation package (link: $($swDownloadUrl))" -Output note
        $destFile = Join-Path $InstallSrcDir $SwSet.$Action.SrcFile
        Invoke-WebRequest $swDownloadUrl -OutFile $destFile
    }
    $ExtraArgs = "/allusers" #/silent
    # Post-installation with disabling of update schtask
    $PostInstall = { 
        MsgFce "INFO: Waiting for OneDrive installation to settle..."
        Start-Sleep -Seconds 5
        $schtaskName = "OneDrive Per-Machine Standalone Update Task"
        if(Get-ScheduledTask -TaskName $schtaskName -ErrorAction SilentlyContinue){
            MsgFce "INFO: Disabling scheduled task '$($schtaskName)'"
            schtasks /change /tn $schtaskName /disable | Out-Null
        }
    }
    # Invoke SW install
    Invoke-VdiSoftwareInstall -SwSettings $SwSet.$Action -ExtraArgs $ExtraArgs -PostInstallAction $PostInstall -InstallType others
}    

# --- Google Drive ---
elseif($Action -like "GoogleDrive"){
    MsgFce "INFO: Install/Update $($SwSet.$Action.Name)" -Output verbose
    # Download Installer?
    $sel = MenuSimple -MenuItems "Yes","No" -Title "Download the latest version of the '$($SwSet.$Action.Name)' installer from the web?"
    if($sel -eq "Yes"){
        $swDownloadUrl = "https://dl.google.com/drive-file-stream/GoogleDriveSetup.exe"
        MsgFce "...downloading the latest '$($SwSet.$Action.Name)' installation package (link: $($swDownloadUrl))" -Output verbose
        $destFile = Join-Path $InstallSrcDir $SwSet.$Action.SrcFile
        Invoke-WebRequest $swDownloadUrl -OutFile $destFile
    }
    $ExtraArgs = "--silent --skip_launch_new"
    if($VAR.GoogleDriveDesktopShortcuts){ 
        $ExtraArgs += " --desktop_shortcut" 
    }
    if(!$VAR.GoogleDriveGSuiteShortcuts){ 
        $ExtraArgs += " --gsuite_shortcuts=false" 
    }
    Invoke-VdiSoftwareInstall -SwSettings $SwSet.GoogleDrive -ExtraArgs $ExtraArgs -InstallType others
}

# --- finish ---
elseif($Action -like "Exit"){
    MsgFce "INFO: No action" -Output verbose
} else{
    MsgFce "ERROR => No action parameter was specified" -Output error
}

#----------------[ Logging maintenance ]----------------
$LogFiles = Get-ChildItem $LogDir | Where-Object {-not $_.PSIsContainer}
if($LogFiles.Count -gt $VAR.LogArchiveFiles){
    MsgFce "There is currently $($LogFiles.Count) files In log archive folder '$($LogDir)' - it's more than set archive limit $($VAR.LogArchiveFiles)" -LinesBefore 1
    MsgFce "The oldest log files will be deleted."
    $LogFiles | Sort-Object LastWriteTime | Select-Object -First ($LogFiles.Count-$VAR.LogArchiveFiles) | Remove-Item -ErrorAction SilentlyContinue
}

# PowerShell responds back to default
$ProgressPreference = "Continue"

# --- End of the script ---
$msg = "End of the script [{0:yyyy-MM-dd HH:mm:ss}]" -f (Get-Date)
MsgFce $msg -Output note -LinesBefore 1

# Transcript Log - Stop
if($Host.Name -match "ConsoleHost"){ Stop-Transcript | out-null }
