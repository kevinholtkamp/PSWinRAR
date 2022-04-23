$ApiKey = Read-Host "Please Enter API Key: "

New-Item "./Publishing/PSWinRAR" -ItemType Directory -Force

Copy-Item ./PSWinRAR.psd1 ./Publishing/PSWinRAR/PSWinRAR.psd1
Copy-Item ./PSWinRAR.psm1 ./Publishing/PSWinRAR/PSWinRAR.psm1

Publish-Module -Path ./Publishing/PSWinRAR -NuGetApiKey $ApiKey -Repository "PSGallery" -ErrorAction "continue"

[version]$version = (Import-PowerShellDataFile ./PSWinRAR.psd1).ModuleVersion
[version]$NewVersion = "{0}.{1}.{2}" -f $Version.Major, $Version.Minor, ($Version.Build + 1)
Update-ModuleManifest -Path ./PSWinRAR.psd1 -ModuleVersion $NewVersion

Remove-Item "./Publishing" -Recurse -Force -Verbose