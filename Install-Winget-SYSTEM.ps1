<#
.SYNOPSIS
The script can be used to download, install, and configure Windows Package Manager (winget) for use in the SYSTEM context, you can also use script to add a private winget repository by passing the necessary paramaters.

.DESCRIPTION
The script does the following:

Accepts optional parameters PrivateWingetName and PrivateWingetURL
Sets package name, MSIXBundle file name, and download URL for Windows Package Manager
Sets the local path for logs and starts the transcript for logging the installation process
Creates the installation folder for the package
Downloads the MSIXBundle file from the URL and saves it to the installation folder
Installs the WinGet MSIXBundle without prompting for a license
Sets the location for the WinGet executable and VCRuntime, and resolves the path of the WinGet executable
Adds both the WinGet executable and VCRuntime folder paths to the PATH variable for the current user
If both PrivateWingetName and PrivateWingetURL parameters are provided, it adds the private winget source to the client
Required Powershell version:
5.1 or higher

.PARAMETER PrivateWingetName
The name of the private winget source (optional)
.PARAMETER PrivateWingetURL
The URL of the private winget source (optional)

.NOTES
Version: 1.0
Author: Florian Salzmann
Creation Date: 22/11/2022
Source: https://github.com/FlorianSLZ/scloud/tree/main/winget/Windows%20Package%20Manager
Purpose/Change: For distribution/installation of the Winget Package Manager with InTune, details on Florian's Blog https://scloud.work/en/how-to-winget-intune/

Version: 2.0
Author: Niall Jennings
Creation Date: 22/04/2022
Purpose/Change: 1 Used original script for installation of Winget in the SYSTEM Context,with the addition of Environment variables for remediation of VCRuntime issues decribed at https://github.com/microsoft/winget-cli/issues/3049
Purpose/Change: 2 Added optional paramaters PrivateWingetName and PrivateWingetURL to allow the addition of a private winget repo as part of the install. The paramaters can be passed to the script at run time or defined directly if prefered 
#>


# Accept two optional parameters: PrivateWingetName and PrivateWingetURL
param (
    [Parameter(mandatory = $false)]
    [string]$PrivateWingetName,
    [Parameter(mandatory = $false)]
    [string]$PrivateWingetURL
)

# Variables for package name, MSIXBundle file name, and the download URL
$PackageName = "WindowsPackageManager"
$MSIXBundle = "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
$URL_msixbundle = "https://aka.ms/getwinget"

# Set the local path for logs
$Path_local = "$Env:Programfiles\_MEM"
# Start the transcript for logging the installation process
Start-Transcript -Path "$Path_local\Log\$PackageName-install.log" -Force

# Create the installation folder for the package
$Folder_install = "$Path_local\Data\$PackageName"
New-Item -Path $Folder_install -ItemType Directory -Force -Confirm:$false

# Initialize the WebClient object to download the MSIXBundle file
$wc = New-Object net.webclient
# Download the MSIXBundle file from the URL and save it to the installation folder
$wc.Downloadfile($URL_msixbundle, "$Folder_install\$MSIXBundle")

# Install the WinGet MSIXBundle
try {
    # Install the MSIXBundle without prompting for a license
    Add-AppxProvisionedPackage -Online -PackagePath "$Folder_install\$MSIXBundle" -SkipLicense 
    # Log the successful installation of the package
    Write-Host "Installation of $PackageName finished"
    # Log the process of setting PATH variables for the system
    Write-Host "Setting PATH Variables for SYSTEM"
    # Set the location for the WinGet executable
    $WinGetInstallLocation = "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe"
    # Resolve the path of the WinGet executable
    $winget_folder = Resolve-Path "$WinGetInstallLocation"
    $winget_folder = $winget_folder.path
    # Set the location for the WinGet VCRuntime
    $winget_VCRuntime = "C:\Program Files\WindowsApps\Microsoft.VCLibs.140.00.UWPDesktop_14.0.30704.0_x64__8wekyb3d8bbwe"
    # Log the WinGet executable and VCRuntime folder paths
    Write-Host "Winget Folder is $winget_folder"
    Write-Host "Winget VCRuntime folder is $winget_VCRuntime"
    # Add both the WinGet executable and VCRuntime folder paths to the PATH variable for the current user
    $newPaths = @("$winget_folder", "$winget_VCRuntime")
    $currentPath = [Environment]::GetEnvironmentVariable("PATH", [System.EnvironmentVariableTarget]::User)
    $pathsArray = $currentPath -split ';'

    # Check if the new paths already exist in the current user's PATH variable
    $pathsToAdd = $newPaths | Where-Object { $_ -notin $pathsArray }

    if ($pathsToAdd) {
        $newPath = $currentPath + ";" + ($pathsToAdd -join ';')
        [Environment]::SetEnvironmentVariable("PATH", $newPath, [System.EnvironmentVariableTarget]::User)
    }

    # Check if both PrivateWingetName and PrivateWingetURL parameters are provided
    if ($PrivateWingetName -and $PrivateWingetURL) {
        # Add the necessary Path variables for the life of the script to allow winget to execute
        $newPaths = @("$winget_folder", "$winget_VCRuntime")
        $env:PATH += ";" + ($newPaths -join ';')
        # Add the private winget source to your client
        winget source add --name $PrivateWingetName $PrivateWingetURL -t Microsoft.Rest
    }
} catch {
    Write-Error "Failed to install $PackageName!"
}

# Install file cleanup
Start-Sleep 3 # to unblock installation file
Remove-Item -Path "$Folder_install" -Force -Recurse

Stop-Transcript