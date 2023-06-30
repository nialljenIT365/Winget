# Converts the first letter of a string to lowercase
function Convert-FirstLetterToLowercase {
    param (
        [Parameter(Mandatory = $true)]
        [string] $InputString
    )
    return $InputString.Substring(0, 1).ToLower()
}

# Fetches the list of vendor folders from the GitHub winget-pkgs repository
function Get-WingetVendorFolders {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FirstFolderLetter
    )
    $vendorFolders = Get-GitHubContent -OwnerName microsoft -RepositoryName winget-pkgs -Path "manifests/$FirstFolderLetter"
    return $vendorFolders.entries.path | ForEach-Object { $_.Split('/')[-1] }
}

# Returns all possible case combinations for a given string
function Get-CaseCombinations {
    param (
        [string]$InputString
    )
    $combinations = @('')
    for ($i = 0; $i -lt $InputString.Length; $i++) {
        $currentChar = $InputString[$i]
        $count = $combinations.Count
        $newCombinations = New-Object System.Collections.Generic.List[string]
        if ($currentChar -cmatch '[^A-Za-z]') {
            for ($j = 0; $j -lt $count; $j++) {
                $currentCombination = $combinations[$j]
                $newCombinations.Add($currentCombination + $currentChar)
            }
        }
        else {
            for ($j = 0; $j -lt $count; $j++) {
                $currentCombination = $combinations[$j]
                $newCombinations.Add($currentCombination + $currentChar.ToString().ToLower())
                $newCombinations.Add($currentCombination + $currentChar.ToString().ToUpper())
            }
        }
        $combinations = $newCombinations
    }
    return $combinations
}

# Finds and returns the matching pair in two lists
function Get-MatchingPair {
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$List1,

        [Parameter(Mandatory = $true)]
        [string[]]$List2
    )
    $matchingPair = Compare-Object $List1 $List2 -IncludeEqual -ExcludeDifferent |
    Where-Object { $_.SideIndicator -eq '==' } |
    ForEach-Object { $_.InputObject }
    return $matchingPair
}

# Fetches the list of application folders for a given vendor from the GitHub winget-pkgs repository
function Get-WingetApplicationFolders {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FirstFolderLetter,

        [Parameter(Mandatory = $true)]
        [string]$VendorFolder
    )
    $applicationFolders = Get-GitHubContent -OwnerName microsoft -RepositoryName winget-pkgs -Path "manifests/$FirstFolderLetter/$VendorFolder"
    return $applicationFolders.entries.path | ForEach-Object { $_.Split('/')[-1] }
}

# Fetches the list of version folders for a given application from the GitHub winget-pkgs repository
function Get-WingetVersionFolders {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderLetter,

        [Parameter(Mandatory = $true)]
        [string]$VendorFolder,

        [Parameter(Mandatory = $true)]
        [string]$ApplicationFolder
    )
    $versionFolders = Get-GitHubContent -OwnerName microsoft -RepositoryName winget-pkgs -Path "manifests/$FolderLetter/$VendorFolder/$ApplicationFolder"
    return $versionFolders.entries | Where-Object { $_.type -ne 'file' }
}

# Fetches the path to the most recent version of a given application from the GitHub winget-pkgs repository
function Get-MostRecentVersionPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderLetter,

        [Parameter(Mandatory = $true)]
        [string]$VendorFolder,

        [Parameter(Mandatory = $true)]
        [string]$ApplicationFolder
    )
    $versionFolderList = Get-WingetVersionFolders -FolderLetter $FolderLetter -VendorFolder $VendorFolder -ApplicationFolder $ApplicationFolder
    $results = foreach ($versionFolder in $versionFolderList) {
        $commitUrl = "https://api.github.com/repos/microsoft/winget-pkgs/commits?path=$($versionFolder.path)"
        $lastModified = (Invoke-WebRequest -Method HEAD -UseBasicParsing $commitUrl).Headers."Last-Modified"
        [PSCustomObject]@{
            "Path"         = $versionFolder.path
            "LastModified" = [datetime]::ParseExact($lastModified, "R", [Globalization.CultureInfo]::InvariantCulture)
        }
    }
    $mostRecentVersion = $results | Where-Object { $_.Path -notmatch 'preview' } | Sort-Object -Property LastModified -Descending | Select-Object -First 1
    return "https://api.github.com/repos/microsoft/winget-pkgs/contents/$($mostRecentVersion.path)"
}

# Main function to retrieve Winget application manifests
function Get-WinGetApplicationManifestPaths {
    param(
        [Parameter(ParameterSetName = 'Search')]
        [switch]$Search,

        [Parameter(ParameterSetName = 'Url')]
        [switch]$Url,

        [Parameter(ParameterSetName = 'Search')]
        [string]$Vendor,

        [Parameter(ParameterSetName = 'Search')]
        [string]$Application,

        [Parameter(ParameterSetName = 'Url')]
        [string]$ManifestUrl,

        [Parameter()]
        [string]$Version = $null

    )

    switch ($PsCmdlet.ParameterSetName) {
        'Search' {
            if ([string]::IsNullOrEmpty($Vendor) -or [string]::IsNullOrEmpty($Application)) {
                Write-Error "When using -Search, both -Vendor and -Application parameters must be provided."
                return
            }
            Write-Host "Script will start to search Github for the latest release of the $Vendor $Application manifest files, note this may take a few minutes and it will exclude any manifest files with preview in the version path , if you know the manifest URL for your prefered Vendor, Application and Version on the WInget Repo consider the -url switch, this will be more efficient"
            $firstFolderLetter = Convert-FirstLetterToLowercase -InputString $Vendor
            $vendorFolderList = Get-WingetVendorFolders -FirstFolderLetter $firstFolderLetter
            $vendorFolderCombinations = Get-CaseCombinations -InputString $Vendor
            $vendorFolder = Get-MatchingPair -List1 $vendorFolderList -List2 $vendorFolderCombinations
            if ([string]::IsNullOrEmpty($vendorFolder)) {
                Write-Error "The Vendor you are searching for does not appear to be in the winget-pks repository, check spelling and try again."
                Write-Output $vendorFolderList
                return
            }
            $applicationFolderList = Get-WingetApplicationFolders -FirstFolderLetter $firstFolderLetter -VendorFolder $vendorFolder
            $applicationFolderCombinations = Get-CaseCombinations -InputString $Application
            $applicationFolder = Get-MatchingPair -List1 $applicationFolderList -List2 $applicationFolderCombinations
            if ([string]::IsNullOrEmpty($applicationFolder)) {
                Write-Error "The Application you are searching for does not appear to be in the winget-pks repository, check spelling and try again."
                Write-Output $applicationFolderList
                return
            }
            $manifestUrl = Get-MostRecentVersionPath -FolderLetter $firstFolderLetter -VendorFolder $vendorFolder -ApplicationFolder $applicationFolder
            $response = Invoke-RestMethod -Uri $manifestUrl -Method Get
            $WinGetApplicationManifestPaths = foreach ($item in $response) {
                New-Object -TypeName PSObject -Property @{
                    FileName     = $item.name
                    DownloadPath = $item.download_url
                } 
            }
            return $WinGetApplicationManifestPaths
        }
        'Url' {
            if ([string]::IsNullOrEmpty($ManifestUrl)) {
                Write-Error "When using -Url, the -ManifestUrl parameter must be provided."
                return
            }
        
            if ($ManifestUrl -match '^manifests/') {
                $ManifestUrl = "https://api.github.com/repos/microsoft/winget-pkgs/contents/" + $ManifestUrl
            }
        
            $urlParts = $ManifestUrl -split '/manifests/'
            if ($urlParts.Count -le 1) {
                Write-Error "Invalid ManifestUrl. The URL must contain '/manifests/'."
                return
            } 

            if ($Version -eq 'latest') {
                Write-Host "Note: You have used the -Version switch in combination with the -url switch, if you have included a specific version in your manifest url path this will now be disregarded in favour of the latest non preview version available in manifests/<Letter>/<Vendor>/<Application>, if you need and have specified a specific application version you will need to remove the -Version switch from your command"
                $manifestParts = $urlParts[1] -split '/'

                $firstFolderLetter = $manifestParts[0]
                $vendorFolder = $manifestParts[1]
                $applicationFolder = $manifestParts[2]

                $manifestUrl = Get-MostRecentVersionPath -FolderLetter $firstFolderLetter -VendorFolder $vendorFolder -ApplicationFolder $applicationFolder
                $response = Invoke-RestMethod -Uri $manifestUrl -Method Get
                $WinGetApplicationManifestPaths = foreach ($item in $response) {
                    New-Object -TypeName PSObject -Property @{
                        FileName     = $item.name
                        DownloadPath = $item.download_url
                    } 
                }
                return $WinGetApplicationManifestPaths
            }
            else {
                $newUrl = "https://api.github.com/repos/microsoft/winget-pkgs/contents/manifests/" + $urlParts[1]
                $response = Invoke-RestMethod -Uri $newUrl -Method Get
        
                $WinGetApplicationManifestPaths = foreach ($item in $response) {
                    New-Object -TypeName PSObject -Property @{
                        FileName     = $item.name
                        DownloadPath = $item.download_url
                    }
                }
                return $WinGetApplicationManifestPaths
            }
        }

    }
}
function Get-MainObject {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [Array]$Files
    )

    if ($Files -isnot [Array]) {
        Write-Error "Input must be an array"
        return
    }

    $MainFile = $null

    #note this excludes any non en-US locales
    foreach ($File in $Files) {
        $fileName = $File.FileName

        # Check if the file name does not end with "locale.*.yaml" or "installer.yaml"
        if (($fileName -notlike "*locale.*.yaml") -and ($fileName -notlike "*installer.yaml")) {
            $MainFile = $File.DownloadPath
        }
    }

    if ($null -eq $MainFile) {
        Write-Warning "The main file was not found in the input array."
    }

    # Read the contents of the files and convert to PowerShell Objects
    $MainObject = ConvertFrom-Yaml -Yaml (Invoke-RestMethod -Uri $MainFile)

    return $MainObject
}

function Get-InstallerObject {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [Array]$Files,
        [string]$Locale = "en-US",
        [ValidateSet('x86', 'x64', 'arm', 'arm64')]
        [string]$Architecture = "x64",
        [ValidateSet('user', 'machine')]
        [string]$Scope = "machine",
        [switch]$All
    )

    if ($Files -isnot [Array]) {
        Write-Error "Input must be an array"
        return
    }

    $InstallerFile = $null

    foreach ($File in $Files) {
        $fileName = $File.FileName

        if ($fileName -like "*installer.yaml") {
            $InstallerFile = $File.DownloadPath
        }
    }

    if ($null -eq $InstallerFile) {
        Write-Warning "The installer file was not found in the input array."
        return
    }

    $InstallerObject = ConvertFrom-Yaml -Yaml (Invoke-RestMethod -Uri $InstallerFile)

    # If All switch is chosen, return all installers
    if ($All) {
        return $InstallerObject
    }

    # If there is only one InstallerUrl, ignore all switches
    if ($InstallerObject.Installers.Count -eq 1) {
        Write-Warning "Only a single installer entry found, disregarding all switches."
        return $InstallerObject
    }

    # Filter based on Architecture
    $FilteredInstallers = $InstallerObject.Installers | Where-Object { $_.Architecture -eq $Architecture }

    if ($FilteredInstallers.Count -eq 0) {
        Write-Error "No installers found for the specified architecture '$Architecture'."
        return
    }

    # Filter based on Locale if there are still multiple installers
    if ($FilteredInstallers.Count -gt 1) {
        $LocaleFilteredInstallers = $FilteredInstallers | Where-Object { $_.InstallerLocale -eq $Locale }
        if ($LocaleFilteredInstallers.Count -gt 0) {
            $FilteredInstallers = $LocaleFilteredInstallers
        }
    }

    # Filter based on Scope if there are still multiple installers
    if ($FilteredInstallers.Count -gt 1) {
        $ScopeFilteredInstallers = $FilteredInstallers | Where-Object { $_.Scope -eq $Scope }
        if ($ScopeFilteredInstallers.Count -gt 0) {
            $FilteredInstallers = $ScopeFilteredInstallers
        }
    }

    $InstallerObject.Installers = $FilteredInstallers

    # Issue a warning if there are still multiple installers and the All switch was not chosen
    if ($InstallerObject.Installers.InstallerUrl.Count -gt 1 -and !$All) {
        Write-Warning "More than one installer found."
    }

    return $InstallerObject
}

function Set-InstallerObject {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSCustomObject]$InstallerObject,
        [string]$Locale = "en-US",
        [ValidateSet('x86', 'x64', 'arm', 'arm64')]
        [string]$Architecture = "x64",
        [ValidateSet('user', 'machine')]
        [string]$Scope = "machine"
    )

    if ($InstallerObject.Installers.InstallerUrl.Count -ne 1) {
        Write-Error "The input object must have exactly one installer."
        return
    }

    $FilteredInstaller = $InstallerObject.Installers

    if ([string]::IsNullOrEmpty($FilteredInstaller.InstallerType)) {
        # Perform action if InstallerType is null or empty
        $FilteredInstaller.InstallerType = $InstallerObject.InstallerType
    }
    
    $InstallerObject["InstallerIdentifier"] = "$($Architecture).$($Locale).$($FilteredInstaller.InstallerType)"
    $InstallerObject["InstallerType"] = $FilteredInstaller.InstallerType
    $InstallerObject["InstallerURL"] = $FilteredInstaller.InstallerUrl
    $InstallerObject["InstallerSha256"] = $FilteredInstaller.InstallerSha256
    $InstallerObject["InstallerLocale"] = $Locale
    $InstallerObject["Architecture"] = $Architecture
    $InstallerObject["Scope"] = $Scope

    $InstallerObject.Remove("Installers")

    return $InstallerObject
}

function Get-LocaleObject {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [Array]$Files,
        [string]$Locale = "en-US"
    )

    if ($Files -isnot [Array]) {
        Write-Error "Input must be an array"
        return
    }

    $LocaleFile = $null

    foreach ($File in $Files) {
        $fileName = $File.FileName

        # Check if the file name ends with "locale.$Locale.yaml"
        if ($fileName -like "*locale.$Locale.yaml") {
            $LocaleFile = $File.DownloadPath
        }
    }

    if ($null -eq $LocaleFile) {
        Write-Warning "The locale file for '$Locale' was not found in the input array."
    }

    # Read the contents of the files and convert to Powershell Objects
    $LocaleObject = ConvertFrom-Yaml -Yaml (Invoke-RestMethod -Uri $LocaleFile)

    # Remove spaces from tags
    if ($LocaleObject.PSObject.Properties.Name -contains 'Tags') {
        $LocaleObject.Tags = $LocaleObject.Tags | ForEach-Object { $_ -replace ' ', '' }
    }

    return $LocaleObject
}

function Get-VersionObject {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [PSCustomObject]$MainObject,

        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [PSCustomObject]$LocaleObject
    )

    # Create a new PSCustomObject with the desired properties
    $TempObject = [PSCustomObject] @{
        PackageVersion = $MainObject.PackageVersion
        DefaultLocale  = $LocaleObject
    }

    # Convert the PSCustomObject to a YAML string
    $YamlString = $TempObject | ConvertTo-Yaml 

    # Convert the YAML string back to an object
    $VersionObject = ConvertFrom-Yaml -Yaml $YamlString

    return $VersionObject
}

function Set-PackageIdentifier {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [PSCustomObject]$YAMLObject,

        [Parameter(Mandatory = $true)]
        [string]$Identifier
    )

    # Validate the Identifier and throw an error if it's not valid
    try {
        if ($Identifier -notmatch '^[a-zA-Z0-9.]+$') {
            throw "The Identifier must only contain alphanumeric characters and/or periods."
        }
    }
    catch {
        Write-Error $_.Exception.Message
        return
    }

    # Modify the PackageIdentifier property of the MainYAMLObject
    $YAMLObject.PackageIdentifier = $Identifier

    return $YAMLObject
}

function Set-InstallerUrl {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        $YAMLObject,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [Uri]
        $InputURI
    )

    $YAMLObject.Installers.InstallerUrl = $InputURI.AbsoluteUri
    return $YAMLObject
}

function Set-Tags {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        $YAMLObject,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $Tags
    )

    $tagsArray = $Tags -split "," | ForEach-Object { $_.Trim().Replace(" ", "") }

    if ($YAMLObject.Tags) {
        $existingTags = $YAMLObject.Tags -split "," | ForEach-Object { $_.Trim().Replace(" ", "") }
        $YAMLObject.Tags = ($existingTags + $tagsArray) | Select-Object -Unique
    }

    if ($YAMLObject.DefaultLocale.Tags) {
        $existingDefaultLocaleTags = $YAMLObject.DefaultLocale.Tags -split "," | ForEach-Object { $_.Trim().Replace(" ", "") }
        $YAMLObject.DefaultLocale.Tags = ($existingDefaultLocaleTags + $tagsArray) | Select-Object -Unique
    }

    return $YAMLObject
}

function Remove-Tags {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        $YAMLObject,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $Tags
    )

    $tagsArray = $Tags | ForEach-Object { $_.Trim() }

    if ($YAMLObject.Tags) {
        $YAMLObject.Tags = $YAMLObject.Tags -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -notin $tagsArray }
    }

    if ($YAMLObject.DefaultLocale.Tags) {
        $YAMLObject.DefaultLocale.Tags = $YAMLObject.DefaultLocale.Tags -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -notin $tagsArray }
    }

    return $YAMLObject
}

function Set-InstallerScope {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [PSCustomObject]$YAMLObject,

        [Parameter(Mandatory = $true)]
        [ValidateSet('user', 'machine')]
        [string]$Scope
    )

    # Check if the Scope property exists in the YAML object
    if ($YAMLObject.PSObject.Properties.Name -contains 'Scope') {
        # If the Scope property exists, overwrite its value
        $YAMLObject.Scope = $Scope
    }
    else {
        # If the Scope property doesn't exist, add it to the YAML object
        $YAMLObject | Add-Member -Type NoteProperty -Name Scope -Value $Scope
    }

    return $YAMLObject
}

function Set-PackageIdentifier {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [PSCustomObject]$YAMLObject,

        [Parameter(Mandatory = $true)]
        [string]$Identifier
    )

    # Validate the Identifier and throw an error if it's not valid
    try {
        if ($Identifier -notmatch '^[a-zA-Z0-9.]+$') {
            throw "The Identifier must only contain alphanumeric characters and/or periods."
        }
    }
    catch {
        Write-Error $_.Exception.Message
        return
    }

    # Modify the PackageIdentifier property of the MainYAMLObject
    $YAMLObject.PackageIdentifier = $Identifier

    return $YAMLObject
}

function Set-InstallerUrl {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        $YAMLObject,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [Uri]
        $InputURI
    )

    $YAMLObject.Installers.InstallerUrl = $InputURI.AbsoluteUri
    return $YAMLObject
}

function Set-Tags {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        $YAMLObject,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $Tags
    )

    $tagsArray = $Tags -split "," | ForEach-Object { $_.Trim().Replace(" ", "") }

    if ($YAMLObject.Tags) {
        $existingTags = $YAMLObject.Tags -split "," | ForEach-Object { $_.Trim().Replace(" ", "") }
        $YAMLObject.Tags = ($existingTags + $tagsArray) | Select-Object -Unique
    }

    if ($YAMLObject.DefaultLocale.Tags) {
        $existingDefaultLocaleTags = $YAMLObject.DefaultLocale.Tags -split "," | ForEach-Object { $_.Trim().Replace(" ", "") }
        $YAMLObject.DefaultLocale.Tags = ($existingDefaultLocaleTags + $tagsArray) | Select-Object -Unique
    }

    return $YAMLObject
}

function Remove-Tags {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        $YAMLObject,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $Tags
    )

    $tagsArray = $Tags | ForEach-Object { $_.Trim() }

    if ($YAMLObject.Tags) {
        $YAMLObject.Tags = $YAMLObject.Tags -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -notin $tagsArray }
    }

    if ($YAMLObject.DefaultLocale.Tags) {
        $YAMLObject.DefaultLocale.Tags = $YAMLObject.DefaultLocale.Tags -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -notin $tagsArray }
    }

    return $YAMLObject
}

function Set-InstallerScope {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [PSCustomObject]$YAMLObject,

        [Parameter(Mandatory = $true)]
        [ValidateSet('user', 'machine')]
        [string]$Scope
    )

    # Check if the Scope property exists in the YAML object
    if ($YAMLObject.PSObject.Properties.Name -contains 'Scope') {
        # If the Scope property exists, overwrite its value
        $YAMLObject.Scope = $Scope
    }
    else {
        # If the Scope property doesn't exist, add it to the YAML object
        $YAMLObject | Add-Member -Type NoteProperty -Name Scope -Value $Scope
    }

    return $YAMLObject
}
