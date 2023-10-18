<#
        .SYNOPSIS
        Microsoft Endpoint Manager PowerShell script as win32 applicaton template.
        This script writes/removes a marker file to track the installation as an application for ease of tracking.
        Add details here:

        .PARAMETER Mode
        Specifies the method of operation.
        Install or Remove are the required parameters.

        .EXAMPLE
        %WINDIR%\Sysnative\WindowsPowerShell\v1.0\powershell.exe -executionpolicy bypass -file Script.ps1 -Mode Install

        .EXAMPLE
        %WINDIR%\Sysnative\WindowsPowerShell\v1.0\powershell.exe -executionpolicy bypass -file Script.ps1 -Mode Remove

        .DETECTION
        Rule Type: File
        Path: %PROGRAMDATA%\!SUPPORT\_Tags
        File or Folder: <ScriptName>.tag
        Detection Method: File or Folder exists

        .LINK
        We are running this script as an application vs the native method due to reasons found in the "before you begin" section so we can reapply if needed.
        https://docs.microsoft.com/en-us/mem/intune/apps/intune-management-extension

        .NOTES
        Version:        1.0
        Author:         Dustin Knight
        Creation Date:  10/20/2021
        Purpose/Change: Initial script
#>


Param(
[Parameter(Mandatory=$true)]
[ValidateSet("Install", "Remove")]
[String]$Mode
)

# Static Variables Section - Don't Edit

# Gets the name of the running powershell script & trims off the .ps1 extension
$scriptName = ([io.fileinfo]$MyInvocation.MyCommand.Definition).BaseName

# Dynamically sets the Marker directory to the progamdata typically C:ProgramData\!SUPPORT\_Tags
$csMarkerPath = Join-Path -Path $Env:PROGRAMDATA -ChildPath "!SUPPORT\_Tags"

# The Marker file named as the running script, & created in the Marker directory
$csMarkerFile = Join-Path -Path $csMarkerPath -ChildPath "$scriptName.tag"

# Dynamically sets the log directory to the system drive typically C:ProgramData\!SUPPORT\_LogFiles
$csLogPath = Join-Path -Path $Env:PROGRAMDATA -ChildPath "!SUPPORT\_LogFiles"

# The log file named as the running script, & created in the Log directory
$csLogFile = Join-Path -Path $csLogPath -ChildPath "$scriptName.log"


# Script Specific Variables Section - Create Here
$count = 0
$exitCode = 0

# Static Functions Section - Don't Edit

function InitLogging { 

    if (!(Test-Path -Path $csLogPath)) {
        New-Item -Path $csLogPath -ItemType Directory -Force | Out-Null
    }

    Start-Transcript -Path $csLogFile -Append -Force
}

# Reminder StopLogging needs to be called before exiting the script 

function StopLogging {

    Stop-Transcript
}

# Creates a marker file to show the script "application" has executed & is considered installed

function CreateMarker {

    if (!(Test-Path -Path $csMarkerPath)) {
        New-Item -Path $csMarkerPath -ItemType Directory -Force | Out-Null
    }

   New-Item $csMarkerFile -type file -Force | Out-Null

   Write-Output "Marker File $csMarkerFile Created"
   Write-Output ""
   
}

# Removes the marker file to show the script "application" is removed & eligble to re-install

function RemoveMarker {

    if (Test-Path -Path $csMarkerPath) {
        Remove-Item $csMarkerFile -Force

        Write-Output "Marker File $csMarkerFile Removed"
        Write-Output ""
    }
        
}

# Add New Script Functions Below, if desired


#Init - Code Execution Starts Here

# Start Logging
InitLogging

# Add Install Code Here
If ($Mode -eq "Install")
{

    #detection
    $product_name = "Microsoft Office 365", "Microsoft 365"
    foreach ($product in $product_name) {
    $x32 = gci "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" -ErrorAction SilentlyContinue | foreach { gp $_.PSPath } | ? { $_ -match $product } | select Displayname, UninstallString
    $x64 = gci "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" -ErrorAction SilentlyContinue | foreach { gp $_.PSPath } | ? { $_ -match $product } | select DisplayName, UninstallString
        if ($x32){ 
        $count++
        $x32.DisplayName
        Write-Output "Office x86 detected"
        $x32.UninstallString 
        }

        if ($x64){ 
        $count++
        $x64.DisplayName
        Write-Output  "Office x64 detected"
        $x64.UninstallString 
        }  
    }

    #removal 
    if ($count -ne '0') {
        Write-Output "Installed OEM Office 365 Detected, Removing"
        $remove = Start-Process -FilePath "setup.exe" -ArgumentList "/configure uninstall.xml" -WindowStyle Hidden -PassThru
        $remove.WaitForExit()
        Write-Output "ODT uninstall exit code: $($remove.ExitCode)"
        $exitCode = $remove.ExitCode
    } else {
        Write-Output ""
        Write-Output "No OEM Installed Office 365 Detected"
    }

    # Finish Success
    CreateMarker
    StopLogging
    exit $exitCode

}

# Add Uninstall Code Here
If ($Mode -eq "Remove")

{
    # Finish Success
    RemoveMarker
    StopLogging
    exit 0

} 