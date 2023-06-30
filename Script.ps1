$endpointBaseUrl = "https://appwgdemone692023.azurewebsites.net/api/"
$hostKey = "QEAGn-w2v5M63UypPU2errLwzwR_BdVMb-9I3RXpDx-aAzFuj-r8JQ=="
$headers = @{ "x-functions-key" = $hostKey }


######RetrieveManifests from GitHub###############################################
$ManifestFiles = Get-WinGetApplicationManifestPaths -Url -ManifestUrl  https://raw.githubusercontent.com/microsoft/winget-pkgs/master/manifests/m/Microsoft/Bicep -Version latest
##################################################################################

######Import YAML Manifest Contents to PS Objects###############################################
$MainObject = Get-MainObject -Files $ManifestFiles
$InstallerObject = Set-InstallerObject -InstallerObject (Get-InstallerObject -Files $ManifestFiles -Architecture x64 -Scope machine -Locale en-us) -Architecture x64 -Scope machine -Locale en-us
$LocaleEnUSObject = Get-LocaleObject -Files $ManifestFiles -Locale en-us
$VersionEnUSObject = Get-VersionObject -MainObject $MainObject -LocaleObject $LocaleEnUSObject
##################################################################################

######Modifications to Objects can go here######


##################################################################################

######Convert PS Objects to JSON format###############################################
$MainJSON = $MainObject | ConvertTo-Json -Depth 10
$InstallerJSON = $InstallerObject | ConvertTo-Json -Depth 10
$LocaleEnUSJSON = $LocaleEnUSObject | ConvertTo-Json -Depth 10
$VersionEnUSJSON = $VersionEnUSObject | ConvertTo-Json -Depth 10
##################################################################################

###############Send the HTTP request without an intermediary file#########################################
Invoke-RestMethod -Uri "$endpointBaseUrl/packages" -Method Post -Headers $headers -Body $MainJSON -ContentType "application/json"

Invoke-RestMethod -Uri "$endpointBaseUrl/packages/$($MainObject.PackageIdentifier)/versions" -Method Post -Headers $headers -Body $VersionEnUSJSON -ContentType "application/json"

Invoke-RestMethod -Uri "$endpointBaseUrl/packages/$($MainObject.PackageIdentifier)/versions/$($MainObject.PackageVersion)/locales" -Method Post -Headers $headers -Body $LocaleEnUSJSON -ContentType "application/json"

Invoke-RestMethod -Uri "$endpointBaseUrl/packages/$($MainObject.PackageIdentifier)/versions/$($MainObject.PackageVersion)/installers" -Method Post -Headers $headers -Body $InstallerJSON -ContentType "application/json"

##################################################################################


###############WinGetCommands#########################################
# winget source add --name appwgdemone692023 https://appwgdemone692023.azurewebsites.net/api/ -t Microsoft.Rest
# winget search "Chrome" -s appwgdemone692023
# winget install --id Google.Chrome --source appwgdemone692023
# winget uninstall --id Google.Chrome