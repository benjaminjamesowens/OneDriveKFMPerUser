<#
.SYNOPSIS
    - Background
        The script was created to deal with users who have network folder redirection in place already - enabling KFM for users with network folder redirection in place would result in the home drive data 
        being moved to the local drive which can cause significant issues with networking, conflicts and delay to user login.
        We required a way to enable OneDrive KFM:
            - a per user basis
            - which didn't move files
            - which properly updated the Documents and Pictures (and/or Desktop if required) in File Explorer to point to the OneDrive local folder    

    - Summary
        The script's purpose is to enable setting OneDrive KFM (known folder move) on a per user basis, as opposed to a per computer basis which is the only GPO or registry supported way to enable it.
        It is intended to change the registry settings for the Documents and Pictures location to point to the local OneDrive location (it is possible to amend the script to include the Desktop aswell)
        It is intended to be run under the context of the logged-in user, as opposed to running as the SYSTEM account due to changes ot the HKCU profile.
    
    - What the script does
        The key actions this script are as follows:
            - Creates c:\programdata\OneDriveFolderRedirection\ if it doesn't already exist for outputting log files
            - Create log file in c:\programdata\OneDriveFolderRedirection\ with usernmae, device, timestamp and unique two character value
            - Check whether OneDrive has been configured on the users profile and that the OneDrive user variable and registry key are present for Business1
            - Check whether the Documents and Pictures registry keys already point to OneDrive local folder - if so, exit
            - Check whether the Documents and Pictures registry keys point to a UNC path - in order to continue with the planned change - exit if not
            - Capture the Documents, Pictures and Desktop registry values before any change is made - output to log file
            - Amend the Documents, Pictures (Desktop can be included if required) to point to the local OneDrive folder using SHSetKnownFolderPath
            - Capture the Documents, Pictures and Desktop registry values after any change is made - output to log file
            - Add registry item with timestamp to verify script has completed (for use with Intune app package detection)
            - Copy log file to a network folder location

.PARAMETER KnownFolder
    The known folder whose path to set.
.PARAMETER Path
    The target path to redirect the folder to.
.NOTES
    Author: Ben Owens
    Created: 16/06/2021
    Started from base script taken from https://gist.github.com/aaronparker/cf124f13bb58d95342707527900d307b which is Forked from: https://gist.github.com/semenko/49a28675e4aae5c8be49b83960877ac5
    In the function Redirect-Folder, I have commented out the Move-Files function call as I don't wish to move files from the source to target folder, this can be commented out if required
    
#>


Function Set-WindowStyle {
param(
    [Parameter()]
    [ValidateSet('FORCEMINIMIZE', 'HIDE', 'MAXIMIZE', 'MINIMIZE', 'RESTORE', 
                 'SHOW', 'SHOWDEFAULT', 'SHOWMAXIMIZED', 'SHOWMINIMIZED', 
                 'SHOWMINNOACTIVE', 'SHOWNA', 'SHOWNOACTIVATE', 'SHOWNORMAL')]
    $Style = 'SHOW',
    [Parameter()]
    $MainWindowHandle = (Get-Process -Id $pid).MainWindowHandle
)
    $WindowStates = @{
        FORCEMINIMIZE   = 11; HIDE            = 0
        MAXIMIZE        = 3;  MINIMIZE        = 6
        RESTORE         = 9;  SHOW            = 5
        SHOWDEFAULT     = 10; SHOWMAXIMIZED   = 3
        SHOWMINIMIZED   = 2;  SHOWMINNOACTIVE = 7
        SHOWNA          = 8;  SHOWNOACTIVATE  = 4
        SHOWNORMAL      = 1
    }
    Write-Verbose ("Set Window Style {1} on handle {0}" -f $MainWindowHandle, $($WindowStates[$style]))

    $Win32ShowWindowAsync = Add-Type –memberDefinition @” 
    [DllImport("user32.dll")] 
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
“@ -name “Win32ShowWindowAsync” -namespace Win32Functions –passThru

    $Win32ShowWindowAsync::ShowWindowAsync($MainWindowHandle, $WindowStates[$Style]) | Out-Null
}
#Minimize all PowerShell Windows
(Get-Process -Name powershell).MainWindowHandle | ForEach { Set-WindowStyle FORCEMINIMIZE $_ }

Function Set-KnownFolderPath {
    # Purpose of this function is to the set the relevant path/s for folders in Windows.
    # The known folder and the path needs to be passed to the function. 
    Param (
            [Parameter(Mandatory = $true)]
            [ValidateSet('AddNewPrograms', 'AdminTools', 'AppUpdates', 'CDBurning', 'ChangeRemovePrograms', 'CommonAdminTools', 'CommonOEMLinks', 'CommonPrograms', 'CommonStartMenu', 'CommonStartup', 'CommonTemplates', 'ComputerFolder', 'ConflictFolder', 'ConnectionsFolder', 'Contacts', 'ControlPanelFolder', 'Cookies', 'Desktop', 'Documents', 'Downloads', 'Favorites', 'Fonts', 'Games', 'GameTasks', 'History', 'InternetCache', 'InternetFolder', 'Links', 'LocalAppData', 'LocalAppDataLow', 'LocalizedResourcesDir', 'Music', 'NetHood', 'NetworkFolder', 'OriginalImages', 'PhotoAlbums', 'Pictures', 'Playlists', 'PrintersFolder', 'PrintHood', 'Profile', 'ProgramData', 'ProgramFiles', 'ProgramFilesX64', 'ProgramFilesX86', 'ProgramFilesCommon', 'ProgramFilesCommonX64', 'ProgramFilesCommonX86', 'Programs', 'Public', 'PublicDesktop', 'PublicDocuments', 'PublicDownloads', 'PublicGameTasks', 'PublicMusic', 'PublicPictures', 'PublicVideos', 'QuickLaunch', 'Recent', 'RecycleBinFolder', 'ResourceDir', 'RoamingAppData', 'SampleMusic', 'SamplePictures', 'SamplePlaylists', 'SampleVideos', 'SavedGames', 'SavedSearches', 'SEARCH_CSC', 'SEARCH_MAPI', 'SearchHome', 'SendTo', 'SidebarDefaultParts', 'SidebarParts', 'StartMenu', 'Startup', 'SyncManagerFolder', 'SyncResultsFolder', 'SyncSetupFolder', 'System', 'SystemX86', 'Templates', 'TreeProperties', 'UserProfiles', 'UsersFiles', 'Videos', 'Windows')]
            [string]$KnownFolder,

            [Parameter(Mandatory = $true)]
            [string]$Path
    )

    # Define known folder GUIDs
    # These are known registry values for setting known key folders in Windows
    $KnownFolders = @{
        'AddNewPrograms' = 'de61d971-5ebc-4f02-a3a9-6c82895e5c04';
        'AdminTools' = '724EF170-A42D-4FEF-9F26-B60E846FBA4F';
        'AppUpdates' = 'a305ce99-f527-492b-8b1a-7e76fa98d6e4';
        'CDBurning' = '9E52AB10-F80D-49DF-ACB8-4330F5687855';
        'ChangeRemovePrograms' = 'df7266ac-9274-4867-8d55-3bd661de872d';
        'CommonAdminTools' = 'D0384E7D-BAC3-4797-8F14-CBA229B392B5';
        'CommonOEMLinks' = 'C1BAE2D0-10DF-4334-BEDD-7AA20B227A9D';
        'CommonPrograms' = '0139D44E-6AFE-49F2-8690-3DAFCAE6FFB8';
        'CommonStartMenu' = 'A4115719-D62E-491D-AA7C-E74B8BE3B067';
        'CommonStartup' = '82A5EA35-D9CD-47C5-9629-E15D2F714E6E';
        'CommonTemplates' = 'B94237E7-57AC-4347-9151-B08C6C32D1F7';
        'ComputerFolder' = '0AC0837C-BBF8-452A-850D-79D08E667CA7';
        'ConflictFolder' = '4bfefb45-347d-4006-a5be-ac0cb0567192';
        'ConnectionsFolder' = '6F0CD92B-2E97-45D1-88FF-B0D186B8DEDD';
        'Contacts' = '56784854-C6CB-462b-8169-88E350ACB882';
        'ControlPanelFolder' = '82A74AEB-AEB4-465C-A014-D097EE346D63';
        'Cookies' = '2B0F765D-C0E9-4171-908E-08A611B84FF6';
        'Desktop' = @('B4BFCC3A-DB2C-424C-B029-7FE99A87C641');
        'Documents' = @('FDD39AD0-238F-46AF-ADB4-6C85480369C7','f42ee2d3-909f-4907-8871-4c22fc0bf756');
        'Downloads' = @('374DE290-123F-4565-9164-39C4925E467B','7d83ee9b-2244-4e70-b1f5-5393042af1e4');
        'Favorites' = '1777F761-68AD-4D8A-87BD-30B759FA33DD';
        'Fonts' = 'FD228CB7-AE11-4AE3-864C-16F3910AB8FE';
        'Games' = 'CAC52C1A-B53D-4edc-92D7-6B2E8AC19434';
        'GameTasks' = '054FAE61-4DD8-4787-80B6-090220C4B700';
        'History' = 'D9DC8A3B-B784-432E-A781-5A1130A75963';
        'InternetCache' = '352481E8-33BE-4251-BA85-6007CAEDCF9D';
        'InternetFolder' = '4D9F7874-4E0C-4904-967B-40B0D20C3E4B';
        'Links' = 'bfb9d5e0-c6a9-404c-b2b2-ae6db6af4968';
        'LocalAppData' = 'F1B32785-6FBA-4FCF-9D55-7B8E7F157091';
        'LocalAppDataLow' = 'A520A1A4-1780-4FF6-BD18-167343C5AF16';
        'LocalizedResourcesDir' = '2A00375E-224C-49DE-B8D1-440DF7EF3DDC';
        'Music' = @('4BD8D571-6D19-48D3-BE97-422220080E43','a0c69a99-21c8-4671-8703-7934162fcf1d');
        'NetHood' = 'C5ABBF53-E17F-4121-8900-86626FC2C973';
        'NetworkFolder' = 'D20BEEC4-5CA8-4905-AE3B-BF251EA09B53';
        'OriginalImages' = '2C36C0AA-5812-4b87-BFD0-4CD0DFB19B39';
        'PhotoAlbums' = '69D2CF90-FC33-4FB7-9A0C-EBB0F0FCB43C';
        'Pictures' = @('33E28130-4E1E-4676-835A-98395C3BC3BB','0ddd015d-b06c-45d5-8c4c-f59713854639');
        'Playlists' = 'DE92C1C7-837F-4F69-A3BB-86E631204A23';
        'PrintersFolder' = '76FC4E2D-D6AD-4519-A663-37BD56068185';
        'PrintHood' = '9274BD8D-CFD1-41C3-B35E-B13F55A758F4';
        'Profile' = '5E6C858F-0E22-4760-9AFE-EA3317B67173';
        'ProgramData' = '62AB5D82-FDC1-4DC3-A9DD-070D1D495D97';
        'ProgramFiles' = '905e63b6-c1bf-494e-b29c-65b732d3d21a';
        'ProgramFilesX64' = '6D809377-6AF0-444b-8957-A3773F02200E';
        'ProgramFilesX86' = '7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E';
        'ProgramFilesCommon' = 'F7F1ED05-9F6D-47A2-AAAE-29D317C6F066';
        'ProgramFilesCommonX64' = '6365D5A7-0F0D-45E5-87F6-0DA56B6A4F7D';
        'ProgramFilesCommonX86' = 'DE974D24-D9C6-4D3E-BF91-F4455120B917';
        'Programs' = 'A77F5D77-2E2B-44C3-A6A2-ABA601054A51';
        'Public' = 'DFDF76A2-C82A-4D63-906A-5644AC457385';
        'PublicDesktop' = 'C4AA340D-F20F-4863-AFEF-F87EF2E6BA25';
        'PublicDocuments' = 'ED4824AF-DCE4-45A8-81E2-FC7965083634';
        'PublicDownloads' = '3D644C9B-1FB8-4f30-9B45-F670235F79C0';
        'PublicGameTasks' = 'DEBF2536-E1A8-4c59-B6A2-414586476AEA';
        'PublicMusic' = '3214FAB5-9757-4298-BB61-92A9DEAA44FF';
        'PublicPictures' = 'B6EBFB86-6907-413C-9AF7-4FC2ABF07CC5';
        'PublicVideos' = '2400183A-6185-49FB-A2D8-4A392A602BA3';
        'QuickLaunch' = '52a4f021-7b75-48a9-9f6b-4b87a210bc8f';
        'Recent' = 'AE50C081-EBD2-438A-8655-8A092E34987A';
        'RecycleBinFolder' = 'B7534046-3ECB-4C18-BE4E-64CD4CB7D6AC';
        'ResourceDir' = '8AD10C31-2ADB-4296-A8F7-E4701232C972';
        'RoamingAppData' = '3EB685DB-65F9-4CF6-A03A-E3EF65729F3D';
        'SampleMusic' = 'B250C668-F57D-4EE1-A63C-290EE7D1AA1F';
        'SamplePictures' = 'C4900540-2379-4C75-844B-64E6FAF8716B';
        'SamplePlaylists' = '15CA69B3-30EE-49C1-ACE1-6B5EC372AFB5';
        'SampleVideos' = '859EAD94-2E85-48AD-A71A-0969CB56A6CD';
        'SavedGames' = '4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4';
        'SavedSearches' = '7d1d3a04-debb-4115-95cf-2f29da2920da';
        'SEARCH_CSC' = 'ee32e446-31ca-4aba-814f-a5ebd2fd6d5e';
        'SEARCH_MAPI' = '98ec0e18-2098-4d44-8644-66979315a281';
        'SearchHome' = '190337d1-b8ca-4121-a639-6d472d16972a';
        'SendTo' = '8983036C-27C0-404B-8F08-102D10DCFD74';
        'SidebarDefaultParts' = '7B396E54-9EC5-4300-BE0A-2482EBAE1A26';
        'SidebarParts' = 'A75D362E-50FC-4fb7-AC2C-A8BEAA314493';
        'StartMenu' = '625B53C3-AB48-4EC1-BA1F-A1EF4146FC19';
        'Startup' = 'B97D20BB-F46A-4C97-BA10-5E3608430854';
        'SyncManagerFolder' = '43668BF8-C14E-49B2-97C9-747784D784B7';
        'SyncResultsFolder' = '289a9a43-be44-4057-a41b-587a76d7e7f9';
        'SyncSetupFolder' = '0F214138-B1D3-4a90-BBA9-27CBC0C5389A';
        'System' = '1AC14E77-02E7-4E5D-B744-2EB1AE5198B7';
        'SystemX86' = 'D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27';
        'Templates' = 'A63293E8-664E-48DB-A079-DF759E0509F7';
        'TreeProperties' = '5b3749ad-b49f-49c1-83eb-15370fbd4882';
        'UserProfiles' = '0762D272-C50A-4BB0-A382-697DCD729B80';
        'UsersFiles' = 'f3ce0f7c-4901-4acc-8648-d5d44b04ef8f';
        'Videos' = @('18989B1D-99B5-455B-841C-AB7C74E4DDFC','35286a68-3c57-41a1-bbb1-0eae73d76c95');
        'Windows' = 'F38BF404-1D43-42F2-9305-67DE0B28FC23';
    }

    # Define SHSetKnownFolderPath if it hasn't been defined already
    $Type = ([System.Management.Automation.PSTypeName]'KnownFolders').Type
    If (-not $Type) {
        $Signature = @'
[DllImport("shell32.dll")]
public extern static int SHSetKnownFolderPath(ref Guid folderId, uint flags, IntPtr token, [MarshalAs(UnmanagedType.LPWStr)] string path);
'@
        $Type = Add-Type -MemberDefinition $Signature -Name 'KnownFolders' -Namespace 'SHSetKnownFolderPath' -PassThru
    }

	# Makes directory path, if it doesn't already exist
	If (!(Test-Path $Path -PathType Container)) {
		New-Item -Path $Path -Type Directory -Force
    }

    # Validate the path
    If (Test-Path $Path -PathType Container) {
        # Call SHSetKnownFolderPath
        #  return $Type::SHSetKnownFolderPath([ref]$KnownFolders[$KnownFolder], 0, 0, $Path)
        ForEach ($guid in $KnownFolders[$KnownFolder]) {
            $result = $Type::SHSetKnownFolderPath([ref]$guid, 0, 0, $Path)
            If ($result -ne 0) {
                $errormsg = "Error redirecting $($KnownFolder). Return code $($result) = $((New-Object System.ComponentModel.Win32Exception($result)).message)"
                Throw $errormsg
            }
        }
    } Else {
        Throw New-Object System.IO.DirectoryNotFoundException "Could not find part of the path $Path."
    }
	
	# Fix up permissions, if we're still here
	Attrib +r $Path
    
    Return $Path
}
Function Get-KnownFolderPath {
    
    <#
    .SYNOPSIS
    Gets a known folder's path using GetFolderPath.
    .PARAMETER KnownFolder
    The known folder whose path to get.
    .NOTES
    https://stackoverflow.com/questions/16658015/how-can-i-use-powershell-to-call-shgetknownfolderpath
    #>

    Param (
            [Parameter(Mandatory = $true)]
            [ValidateSet('AdminTools','ApplicationData','CDBurning','CommonAdminTools','CommonApplicationData','CommonDesktopDirectory','CommonDocuments','CommonMusic','CommonOemLinks','CommonPictures','CommonProgramFiles','CommonProgramFilesX86','CommonPrograms','CommonStartMenu','CommonStartup','CommonTemplates','CommonVideos','Cookies','Desktop','DesktopDirectory','Favorites','Fonts','History','InternetCache','LocalApplicationData','LocalizedResources','MyComputer','MyDocuments','MyMusic','MyPictures','MyVideos','NetworkShortcuts','Personal','PrinterShortcuts','ProgramFiles','ProgramFilesX86','Programs','Recent','Resources','SendTo','StartMenu','Startup','System','SystemX86','Templates','UserProfile','Windows')]
            [string]$KnownFolder
    )

    Return [Environment]::GetFolderPath($KnownFolder)
}
Function Move-Files {

    <#
    .SYNOPSIS
        Moves contents of a folder with output to a log.
        Uses Robocopy to ensure data integrity and all moves are logged for auditing.
        Means we don't need to re-write functionality in PowerShell.
    .PARAMETER Source
        The source folder.
    .PARAMETER Destination
        The destination log.
    .PARAMETER Log
        The log file to store progress/output
    #>

    Param (
            [Parameter(Mandatory = $true)]
            [string]$Source,

            [Parameter(Mandatory = $true)]
            [string]$Destination,

            [Parameter(Mandatory = $true)]
            [string]$Log
    )

    If (!(Test-Path (Split-Path $Log))) { New-Item -Path (Split-Path $Log) -ItemType Container }
    Robocopy.exe $Source $Destination /E /MOV /XJ /R:1 /W:1 /NP /LOG+:$Log
}
Function Redirect-Folder {

    <#
    .SYNOPSIS
        Function exists to reduce code required to redirect each folder.
    #>

    Param (
        $SyncFolder,
        $GetFolder,
        $SetFolder,
        $Target
    )

    # Get current Known folder path
    $Folder = Get-KnownFolderPath -KnownFolder $GetFolder

    # If paths don't match, redirect the folder
    If ($Folder -ne "$SyncFolder\$Target") {
        # Redirect the folder
        Set-KnownFolderPath -KnownFolder $SetFolder -Path "$SyncFolder\$Target"
        # Move files/folders into the redirected folder 

        # Move-Files -Source $Folder -Destination "$SyncFolder\$Target" -Log "$env:LocalAppData\RedirectLogs\Robocopy$Target.log"

        # Hide the source folder (rather than delete it)
        #Attrib +h $Folder
    }    
}
Function LogWrite {

        <#
    .SYNOPSIS
        Writes log entries to log files
    .PARAMETER Logstring
        To be details on information to add to log file
    #>

   Param ([string]$logstring)

   Add-content $Logfile -value $logstring
}

Function Toast-Notification {

    <#
    .SYNOPSIS
        Sends a notification to the end user
    .PARAMETER ToastAppID
        The AppID launcher associated to notification - run Get-StartApps for list of apps
    .PARAMETER ToastImage
        Image which will displayed
    .PARAMETER ToastTitle
        The main title of the script
    .PARAMETER ToastMessage
        The main text in the message
    #>

    Param (
            
            [Parameter(Mandatory = $true)]
            [string]$ToastAppID,

            [Parameter(Mandatory = $true)]
            [string]$ToastImage,

            [Parameter(Mandatory = $true)]
            [string]$ToastTitle,
            
            [Parameter(Mandatory = $true)]
            [string]$ToastAttribution,

            [Parameter(Mandatory = $true)]
            [string]$ToastText
    )
#Specify Launcher App ID
$LauncherID = $ToastAppID

#Load Assemblies
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
 
#Build XML Template
#for reference
#<binding template="ToastImageAndText03">
[xml]$ToastTemplate = @"
<toast duration="long">
    <visual>
    <binding template="ToastGeneric">
            <text id="1">$ToastTitle</text>
            <text id="2">$ToastText</text>
            <image placement="hero" id="1" src="$ToastImage"/>
            <text placement="attribution">$ToastAttribution</text>
        </binding>
    </visual>
    <actions>
    <action activationType="protocol" content="Open" arguments="$env:OneDriveCommercial" />
    </actions>
</toast>
"@
 
#Prepare XML
$ToastXml = [Windows.Data.Xml.Dom.XmlDocument]::New()
$ToastXml.LoadXml($ToastTemplate.OuterXml)
 
#Prepare and Create Toast
$ToastMessage = [Windows.UI.Notifications.ToastNotification]::New($ToastXML)
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($LauncherID).Show($ToastMessage)

}

#----------------------------------------------------------[Declarations]----------------------------------------------------------
#Log location where script will attempt to upload log file at the end
$NetworkLogLocation = "\\SERVERNAME\LOGS"

#Create unique log file name based on time, user and computer
$Timestamp = Get-Date -UFormat %Y%m%d-%H%M
# generate a random 6 digit number
$Random = -join(48..57+65..90+97..122 | ForEach-Object {[char]$_} | Get-Random -Count 2)
$hostname = $env:computername
$Username = $env:username
$UPN = whoami /upn
$LogFilename = "ODFR-$Username-$Hostname-$Timestamp-$Random.log"
$LogTranscriptFilename = "ODFB-$Username-$Hostname-$Timestamp-$Random.transcript"
$LogDir = "$($env:ProgramData)\OneDriveFolderRedirection\"
$LogFile = "$LogDir" + "$LogFilename"
$LogTranscriptFile = "$LogDir" + "$LogTranscriptFilename"

#-----------------------------------------------------------[Execution]------------------------------------------------------------

# Create a directory for tag or log files to be created
if (-not (Test-Path "$($env:ProgramData)\OneDriveFolderRedirection"))
{
    Mkdir "$($env:ProgramData)\OneDriveFolderRedirection"
}

# A transcript to outputs all commands and outputs from a script to a file
Start-Transcript -Path $LogTranscriptFile

#Declaring user/workstation/time in script
LogWrite -logstring $Timestamp
LogWrite -logstring $UPN
LogWrite -logstring $hostname

# Get OneDrive sync folder value, if it exists
Try {
    $OneDriveFolder = Get-ItemPropertyValue -Path 'HKCU:\Software\Microsoft\OneDrive\Accounts\Business1' -Name 'UserFolder' -ErrorAction Stop
}
Catch {
    LogWrite -logstring $($Error[0])
} 

#Exit script if OneDriveFolderPath is blank
If ($OneDriveFolder -eq $NULL) {
    LogWrite -logstring "Script will exit as OneDrive registry key is blank and doesn't appeared signed in and configured"
    Stop-Transcript
    Robocopy.exe $LogDir $NetworkLogLocation $LogFilename /XJ /NP
    Robocopy.exe $LogDir $NetworkLogLocation $LogTranscriptFilename /XJ /NP
    Return 999
    Exit
}

#Gather the details of the current registry keys for HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders
Try {
    $HKCUShellFolders = Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" -ErrorAction Stop
}
Catch {
    LogWrite -logstring $($Error[0])
}

#Gather the details of the current registry keys for HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders
Try {
    $HKCUUserShellFolders = Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -ErrorAction Stop
}
Catch {
    LogWrite -logstring $($Error[0])
}

LogWrite -logstring "-----------------------------------------------"
LogWrite -logstring "| Registry values before any changes in script |"
LogWrite -logstring "-----------------------------------------------"
LogWrite -logstring "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal - $($HKCUShellFolders.Personal)"
LogWrite -logstring "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Pictures - $($HKCUShellFolders."My Pictures")"
LogWrite -logstring "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Desktop - $($HKCUShellFolders."Desktop")"
LogWrite -logstring "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Personal - $($HKCUUserShellFolders.Personal)"
LogWrite -logstring "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\{F42EE2D3-909F-4907-8871-4C22FC0BF756} - $($HKCUUserShellFolders."{F42EE2D3-909F-4907-8871-4C22FC0BF756}")"
LogWrite -logstring "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\My Pictures - $($HKCUUserShellFolders."My Pictures")"
LogWrite -logstring "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\{0DDD015D-B06C-45D5-8C4C-F59713854639} - $($HKCUUserShellFolders."{0DDD015D-B06C-45D5-8C4C-F59713854639}")"
LogWrite -logstring "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Desktop - $($HKCUUserShellFolders.Desktop)"
LogWrite -logstring "-----------------------------------------------"

#Checks if the User Shell Folders or Shell Folder registry keys for Documents and Pictures point the local OneDrive location. Script will exit if one ore more are already pointing at OneDrive
If ($HKCUShellFolders.Personal -like "*$OneDriveFolder*" -or $HKCUShellFolders."My Pictures" -like "*$OneDriveFolder*" -or $HKCUUserShellFolders.Personal -like "*$OneDriveFolder*" -or $HKCUUserShellFolders."{F42EE2D3-909F-4907-8871-4C22FC0BF756}" -like "*$OneDriveFolder*" -or $HKCUUserShellFolders."My Pictures" -like "*$OneDriveFolder*" -or $HKCUUserShellFolders."{0DDD015D-B06C-45D5-8C4C-F59713854639}" -like "*$OneDriveFolder*") {
    LogWrite -logstring "Script will exit as one or more registry values for documents/pictures already point to OneDrive for Business local folder locations"
    Stop-Transcript
    Robocopy.exe $LogDir $NetworkLogLocation $LogFilename /XJ /NP
    Robocopy.exe $LogDir $NetworkLogLocation $LogTranscriptFilename /XJ /NP
    Return 999
    Exit
}

#Checks if the User Shell Folders or Shell Folder registry keys for Documents and Pictures point UNC path - if not, script will exit
# NOTE remove this section if you don't want the UNC path check
If ($HKCUShellFolders.Personal -notlike "*\\*" -or $HKCUShellFolders."My Pictures" -notlike "*\\*" -or $HKCUUserShellFolders.Personal -notlike "*\\*" -or $HKCUUserShellFolders."{F42EE2D3-909F-4907-8871-4C22FC0BF756}" -notlike "*\\*" -or $HKCUUserShellFolders."My Pictures" -notlike "*\\*" -or $HKCUUserShellFolders."{0DDD015D-B06C-45D5-8C4C-F59713854639}" -notlike "*\\*") {
    LogWrite -logstring "Script will exit as one or more registry values for documents/pictures do not point to network UNC path"
    Stop-Transcript
    Robocopy.exe $LogDir $NetworkLogLocation $LogFilename /XJ /NP
    Robocopy.exe $LogDir $NetworkLogLocation $LogTranscriptFilename /XJ /NP
    Return 999
    Exit
}

# Redirect select folders
If (Test-Path $OneDriveFolder) {
    
    #NOTE remove comment against desktop to include that also
    LogWrite -logstring "Attempting to update paths to OneDrive location"
    #Redirect-Folder -SyncFolder $OneDriveFolder -GetFolder 'Desktop' -SetFolder 'Desktop' -Target 'Desktop'
    Redirect-Folder -SyncFolder $OneDriveFolder -GetFolder 'MyDocuments' -SetFolder 'Documents' -Target 'Documents'
    Redirect-Folder -SyncFolder $OneDriveFolder -GetFolder 'MyPictures' -SetFolder 'Pictures' -Target 'Pictures'

    Try {
        #LogWrite -logstring "Creating .tag file in $($env:ProgramData)\OneDriveFolderRedirection"
        #Set-Content -Path "$($env:ProgramData)\OneDriveFolderRedirection\OneDriveFolderRedirection.tag" -Value "Installed" -ErrorAction Stop
        LogWrite -logstring "Creating registry value 'ODRedirectionCompleted' in 'HKCU\SOFTWARE\Microsoft\OneDrive' for package detection"
        reg.exe add "HKCU\SOFTWARE\Microsoft\OneDrive" /v "ODRedirectionCompleted" /t REG_SZ /d $Timestamp /f
    }
    Catch {
        LogWrite -logstring $($Error[0])
    }

    #Gather the details of the current registry keys for HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders
    Try {
        $HKCUShellFolders = $NULL
        $HKCUShellFolders = Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" -ErrorAction Stop
    }
    Catch {
        LogWrite -logstring $($Error[0])
    }

    #Gather the details of the current registry keys for HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders
    Try {
        $HKCUUserShellFolders = $NULL
        $HKCUUserShellFolders = Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -ErrorAction Stop
    }
    Catch {
        LogWrite -logstring $($Error[0])
}

LogWrite -logstring "-----------------------------------------------"
LogWrite -logstring "| Registry values after changes in script |"
LogWrite -logstring "-----------------------------------------------"
LogWrite -logstring "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal - $($HKCUShellFolders.Personal)"
LogWrite -logstring "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Pictures - $($HKCUShellFolders."My Pictures")"
LogWrite -logstring "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Desktop - $($HKCUShellFolders."Desktop")"
LogWrite -logstring "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Personal - $($HKCUUserShellFolders.Personal)"
LogWrite -logstring "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\{F42EE2D3-909F-4907-8871-4C22FC0BF756} - $($HKCUUserShellFolders."{F42EE2D3-909F-4907-8871-4C22FC0BF756}")"
LogWrite -logstring "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\My Pictures - $($HKCUUserShellFolders."My Pictures")"
LogWrite -logstring "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\{0DDD015D-B06C-45D5-8C4C-F59713854639} - $($HKCUUserShellFolders."{0DDD015D-B06C-45D5-8C4C-F59713854639}")"
LogWrite -logstring "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Desktop - $($HKCUUserShellFolders.Desktop)"
LogWrite -logstring "-----------------------------------------------"

}

Toast-Notification -ToastAppID "Microsoft.SkyDrive.Desktop" -ToastImage "$PSScriptRoot\OneDriveLogo.png" -ToastTitle "OneDrive Migration Complete" -ToastAttribution "IT Team" -ToastText "Your home drive data has been migrated to OneDrive.  Your 'Documents' and 'Pictures' will now point to your OneDrive location."

Stop-Transcript
Robocopy.exe $LogDir $NetworkLogLocation $LogFilename /XJ /NP
Robocopy.exe $LogDir $NetworkLogLocation $LogTranscriptFilename /XJ /NP
Return 0




