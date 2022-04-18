# Powershell WinRAR (PSWinRAR)
This is a simple Powershell module that wraps the WinRAR command line into Powershell cmdlets with easier to read parameter naming


## How to Use
PSWinRAR can be easily installed by using `Install-Module -Name PSWinRAR` inside of powershell, after which the functions are usable


## Functions
- `Compress-WinRAR` is used to compress the contents of the directory specified in `-DirectoryToCompress` into the .rar file specified in `-ArchivePath`
- `Expand-WinRAR` is used to decompress the contents of the .rar file specified in `-ArchivePath` into the directory specified in `-TargetDirectory`
- `Test-WinRAR` is a wrapper for `Check-WinRAR` which returns `$true` if `Check-WinRAR` returns `0`, and returns `$false` otherwise. If the Switch-Parameter `-GetReturnCode` is set, the returncode from the winrar executable is returned instead
- `Repair-WinRAR` is used to check an archive for errors and correct them which returns `$true` if the WinRAR executable returns `0`, and returns `$false` otherwise. If the Switch-Parameter `-GetReturnCode` is set, the returncode from the winrar executable is returned instead