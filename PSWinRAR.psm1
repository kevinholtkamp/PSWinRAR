function Compress-WinRAR(){
    <#
    .SYNOPSIS
        Compress the specified directory into the specified .rar archive
    .DESCRIPTION
        Compress the specified directory into the specified .rar archive
    .PARAMETER DirectoryToCompress
        Directory which will be compressed into the Archive
    .PARAMETER ArchivePath
        Path to .rar file into which the directory will be compressed
    .PARAMETER CompressionLevel
        Level of compression. Higher values use more compute power and will possibly result in lower archive size. Permitted Values 0-5
    .PARAMETER DictionarySize
        Size of dictionary to use while compressing. A bigger dictionary results in higher RAM usage and possibly lower archive size. Permitted values are 1-1024
    .PARAMETER DictionarySizeUnit
        Unit for parameter DictionarySize. Possible values are 'k', 'm' and 'g' for kilobytes, megabytes and gigabytes respectively
    .PARAMETER Threads
        Number of Threads to use for compression. Permitted values are 1 up to the number of threads of the system
    .PARAMETER Password
        Password to be set for the .rar file
    .PARAMETER ErrorLogFile
        Path to file where errors will be logged
    .PARAMETER FileMask
        Filter files to include
    .PARAMETER RecoveryPercentage
        Percentage of archive size that will be dedicated to recovery data
    .PARAMETER Delete
        Delete original directory after compression
    .PARAMETER SafeDelete
        Move original directory to recycle bin after compression
    .PARAMETER IgnoreEmptyDirectories
        Ignore empty directories in the source directory and not adding them to the archive
    .PARAMETER Recurse
        Recurse the source directory
    .PARAMETER SolidArchive
        Create a solid archive. Results in smaller size but slower read/modify speeds for the archive
    .PARAMETER TestArchive
        Test the archive after creation
    .PARAMETER StructureInArchive
        Set the directory structure in the Archive.
        # 'Flat' moves all files into the root of the archive.
        # 'Relative' moves all files into the archive with their relative paths to the root directory into the archive
        # 'Full' includes the full path of every file in the archive
        # 'DriveLetter' acts like 'Full' but also includes the drive letter
    .PARAMETER UseRAR4
        Use older RAR4 compression method instead of RAR5
    .PARAMETER PassThruParameters
        This string gets passed through to the winrar command line tools
    .INPUTS
        Pipeline inputs get used as DirectoryToCompress
    .OUTPUTS
        Returns the archive Path
    .EXAMPLE
        Compress-WinRAR ./Directory/ ./Archive.rar -Threads 8
    #>

    [CmdletBinding(SupportsShouldProcess)]
    param(
    [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
    [Alias("Directory")]
    [String] $DirectoryToCompress,

    [Parameter(Position = 1)]
    [Alias("Archive")]
    [ValidatePattern(".*\.rar")]
    [String] $ArchivePath,

    [Parameter(Position = 2)]
    [ValidateRange(0,5)]
    [Byte] $CompressionLevel = 3,

    [Parameter(Position = 3)]
    [ValidateRange(1,1024)]
    [Int16] $DictionarySize = 128,

    [Parameter(Position = 4)]
    [ValidateSet("k", "m", "g")]
    [String] $DictionarySizeUnit = "m",

    [Parameter(Position = 5)]
    [ValidateScript({$_ -gt 0 -and $_ -le $Env:NUMBER_OF_PROCESSORS})]
    [Int16] $Threads = 1,

    [Parameter(Position = 6)]
    [ValidateSet("Relative", "Full", "Flat", "DriveLetter")]
    [String] $ArchiveFileStructure = "Relative",

    [Parameter(Position = 7)]
    [ValidateRange(0, 1000)]
    [Int16] $RecoveryPercentage,

    [String] $Password,
    [String] $ErrorLogFile,
    [String] $FileMask,
    [String] $PassThruParameters,

    [Switch] $Delete,
    [Switch] $SafeDelete,
    [Switch] $IgnoreEmptyDirectories,
    [Switch] $NoRecurse,
    [Switch] $SolidArchive,
    [Switch] $TestArchive,
    [Switch] $UseRAR4
    )

    if(!$ArchivePath){
        $ArchivePath = $DirectoryToCompress.TrimEnd("\")
    }

    if(!$NoRecurse){
        $Switches = "$Switches -r"
        Write-Verbose "Setting parameter -r since switch Recursive is set"
    }
    if($Delete){
        $Switches = "$Switches -df"
        Write-Verbose "Setting parameter -df since switch Delete is set"
    }
    if($SafeDelete){
        $Switches = "$Switches -dr"
        Write-Verbose "Setting parameter -dr since switch SafeDelete is set"
    }
    if($IgnoreEmptyDirectories){
        $Switches = "$Switches -ed"
        Write-Verbose "Setting parameter -ed since switch IgnoreEmptyDirectories is set"
    }
    if($ErrorLogFile){
        $Switches = "$Switches -ilog$ErrorLogFile"
        Write-Verbose "Setting parameter -ilog$ErrorLogFile since parameter ErrorLogFile is set"
    }
    if($FileMask){
        $Switches = "$Switches -n$FileMask"
        Write-Verbose "Setting parameter -n$FileMask since parameter FileMask is set"
    }
    if($Password){
        $Switches = "$Switches -p$Password"
        Write-Verbose "Setting parameter -p$Password since parameter Password is set"
    }
    if($RecoveryPercentage){
        $Switches = "$Switches -rr$($RecoveryPercentage)p"
        Write-Verbose "Setting parameter -rr$($RecoveryPercentage)p since parameter RecoveryPercentage is set"
    }
    if($SolidArchive){
        $Switches = "$Switches -s"
        Write-Verbose "Setting parameter -s since switch SolidArchive is set"
    }
    if($TestArchive){
        $Switches = "$Switches -t"
        Write-Verbose "Setting parameter -t since switch TestArchive is set"
    }
    if($Confirm){
        $Switches = "$Switches -y"
        Write-Verbose "Setting parameter -y since switch Confirm is set"
    }
    if($UseRAR4){
        $Switches = "$Switches -ma4"
        Write-Verbose "Setting parameter -ma4 since switch UseRAR4 is set"
    }
    switch($ArchiveFileStructure){
        "Flat"{
            $Switches = "$Switches -ep"
            Write-Verbose "Setting parameter -ep since StructureInArchive is set to 'Flat'"
        }
        "Relative"{
            $Switches = "$Switches -ep1"
            Write-Verbose "Setting parameter -ep1 since StructureInArchive is set to 'Relative'"
        }
        "Full"{
            $Switches = "$Switches -ep2"
            Write-Verbose "Setting parameter -ep2 since StructureInArchive is set to 'Full'"
        }
        "DriveLetter"{
            $Switches = "$Switches -ep3"
            Write-Verbose "Setting parameter -ep3 since StructureInArchive is set to 'DriveLetter'"
        }
    }

    $Switches = "$Switches -m$CompressionLevel"
    Write-Verbose "Setting parameter -m$CompressionLevel with value from parameter CompressionLevel"
    $Switches = "$Switches -md$DictionarySize$DictionarySizeUnit"
    Write-Verbose "Setting parameter -md$DictionarySize$DictionarySizeUnit with value from parameter DictionarySize"
    $Switches = "$Switches -mt$Threads"
    Write-Verbose "Setting parameter -mt$Threads with value from parameter Threads"

    Write-Verbose "Calling WinRAR via command line: Start-Process -FilePath '$(Get-WinRARPath -ErrorAction "Stop")\Rar.exe' -ArgumentList 'a $PassThruParameters $Switches $ArchivePath $DirectoryToCompress' -Wait -PassThru"
    $ExitCode = (Start-Process -FilePath "$(Get-WinRARPath -ErrorAction "Stop")\Rar.exe" -ArgumentList "a $PassThruParameters $Switches $ArchivePath $DirectoryToCompress" -Wait -PassThru).ExitCode
    if($ExitCode -ne 0){
        throw "Winrar stopped with exit code $ExitCode"
    }

    return $ArchivePath
}

function Expand-WinRAR(){
    <#
    .SYNOPSIS
        Expand ("Decompress") a .rar archive into a directory
    .DESCRIPTION
        Expand ("Decompress") the specified .rar file into the specified directory
    .PARAMETER ArchivePath
        Path to .rar file which will be expanded ("Decompressed")
    .PARAMETER TargetDirectory
        Directory into which the contents of the file specified in ArchivePath will be moved. If it does not exist, it will be created
    .PARAMETER Password
        Password for the .rar file
    .INPUTS
        Pipeline inputs get used as ArchivePath
    .OUTPUTS
        Returns the directory path
    .EXAMPLE
        Expand-WinRAR ./Archive.rar ./Directory/
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
    [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "Path to .rar file which will be decompressed")]
    [Alias("Archive", "Path")]
    [ValidatePattern(".*\.rar")]
    [String] $ArchivePath,

    [Parameter(Position = 1, Mandatory = $true, HelpMessage = "Directory into which the archive will be decompressed")]
    [Alias("Target", "Directory")]
    [String] $TargetDirectory,

    [Parameter(Position = 2, HelpMessage = "Password to the .rar file")]
    [String] $Password
    )

    if(!(Test-Path $TargetDirectory)){
        New-Item $TargetDirectory -ItemType Directory -Force
    }

    if($Password){
        $Switches = "$Switches -p$Password"
        Write-Verbose "Setting parameter -p$Password since parameter Password is set"
    }
    if($Confirm){
        $Switches = "$Switches -y"
        Write-Verbose "Setting parameter -y since switch Confirm is set"
    }

    Write-Verbose "Calling WinRAR via command line: Start-Process -FilePath '$(Get-WinRARPath -ErrorAction "Stop")\UnRAR.exe' -ArgumentList 'x $Switches $ArchivePath $TargetDirectory' -Wait -PassThru"
    $ExitCode = (Start-Process -FilePath "$(Get-WinRARPath -ErrorAction "Stop")\UnRAR.exe" -ArgumentList "x $Switches $ArchivePath $TargetDirectory" -Wait -PassThru).ExitCode
    if($ExitCode -ne 0){
        throw "Winrar stopped with exit code $ExitCode"
    }

    return $TargetDirectory
}

function Test-WinRAR(){
    <#
    .SYNOPSIS
        Test a .rar archive for validity
    .DESCRIPTION
        Test a .rar archive for validity
    .PARAMETER ArchivePath
        Path to .rar file which will be tested
    .PARAMETER Password
        Password for the .rar file
    .PARAMETER GetReturnCode
        Pass through the return code from the winrar command line tool instead of $true/$false
    .INPUTS
        Pipeline inputs get used as ArchivePath
    .OUTPUTS
        Returns $true if the archive is valid, $false if the archive is invalid
        If the parameter -GetReturnCode is set, the returncode from the winrar command line tool will be passed through instead
    .EXAMPLE
        Test-WinRAR ./Archive.rar -GetReturnCode
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "Path to .rar file which will be tested")]
        [Alias("Archive", "Path")]
        [String] $ArchivePath,

        [Parameter(Position = 1, HelpMessage = "Password to .rar file")]
        [String] $Password,

        [Parameter(HelpMessage = "Return the returncode from the winrar executable instead of true/false")]
        [Switch] $GetReturnCode
    )

    if($Password){
        $Switches = "$Switches -p$Password"
        Write-Verbose "Setting parameter -p$Password since parameter Password is set"
    }

    Write-Verbose "Calling WinRAR via command line: Start-Process -FilePath $(Get-WinRARPath -ErrorAction "Stop")\Rar.exe -ArgumentList t $Switches $ArchivePath -PassThru -Wait"
    $ReturnCode = (Start-Process -FilePath "$(Get-WinRARPath -ErrorAction "Stop")\Rar.exe" -ArgumentList "t $Switches $ArchivePath" -PassThru -Wait).ExitCode
    if($GetReturnCode){
        return $ReturnCode
    }
    else{
        if($ReturnCode -eq 0){
            return $true
        }
        else{
            return $false
        }
    }
}

function Repair-WinRAR(){
    <#
    .SYNOPSIS
        Repair a .rar archive
    .DESCRIPTION
        Repair a .rar archive
    .PARAMETER ArchivePath
        Path to .rar file which will be repaired
    .PARAMETER Password
        Password for the .rar file
    .PARAMETER GetReturnCode
        Pass through the return code from the winrar command line tool instead of $true/$false
    .INPUTS
        Pipeline inputs get used as ArchivePath
    .OUTPUTS
        Returns $true if the archive was repaired successfully, $false if the archive was not repaired
        If the parameter -GetReturnCode is set, the returncode from the winrar command line tool will be passed through instead
    .EXAMPLE
        Test-WinRAR ./Archive.rar -GetReturnCode
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "Path to .rar file which will be tested")]
        [Alias("Archive", "Path")]
        [String] $ArchivePath,

        [Parameter(Position = 1, HelpMessage = "Password to .rar file")]
        [String] $Password,

        [Parameter(HelpMessage = "Return the returncode from the winrar executable instead of true/false")]
        [Switch] $GetReturnCode
    )

    if($Password){
        $Switches = "$Switches -p$Password"
        Write-Verbose "Setting parameter -p$Password since parameter Password is set"
    }

    Write-Verbose "Calling WinRAR via command line: Start-Process -FilePath $(Get-WinRARPath -ErrorAction "Stop")\Rar.exe -ArgumentList r $Switches $ArchivePath -PassThru -Wait"
    $ReturnCode = (Start-Process -FilePath "$(Get-WinRARPath -ErrorAction "Stop")\Rar.exe" -ArgumentList "r $Switches $ArchivePath" -PassThru -Wait).ExitCode
    if($GetReturnCode){
        return $ReturnCode
    }
    else{
        if($ReturnCode -eq 0){
            return $true
        }
        else{
            return $false
        }
    }
}

function Get-WinRARPath(){
    [CmdletBinding(SupportsShouldProcess)]
    param()

    if(Test-Path "C:\Program Files\WinRAR"){
        return "C:\Program Files\WinRAR"
    }
    if(Test-Path "C:\Program Files (x86)\WinRAR"){
        return "C:\Program Files (x86)\WinRAR"
    }
    throw "WinRAR installation not found"
}

#Check for WinRAR directory on module install and abort installation when WinRAR is not installed
if(!(Get-WinRARPath -ErrorAction SilentlyContinue)){
    throw "WinRAR installation not found, aborting module installation"
}