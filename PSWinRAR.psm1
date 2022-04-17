function Compress-WinRAR(){
    [CmdletBinding()]
    param(
    [Parameter(Position = 0, Mandatory = $true)]
    [String] $DirectoryToCompress,

    [Parameter(Position = 1, Mandatory = $true)]
    [String] $ArchivePath,

    [Switch] $Delete,
    [Switch] $SafeDelete,
    [Switch] $IgnoreEmptyDirectories,
    [String] $ErrorLogFile,
    [Byte] $CompressionLevel = 3,
    [Int16] $DictionarySize = 128,
    [Byte] $Threads = 4,
    [String] $FileMask,
    [String] $Password,
    [Switch] $Recursive,
    [Byte] $RecoveryPercentage,
    [Switch] $SolidArchive,
    [Switch] $TestData,
    [Switch] $Confirm,
    [Switch] $FullPath
    )

    if($Recursive){
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
    if($TestData){
        $Switches = "$Switches -t"
        Write-Verbose "Setting parameter -t since switch TestData is set"
    }
    if($Confirm){
        $Switches = "$Switches -y"
        Write-Verbose "Setting parameter -y since switch Confirm is set"
    }
    if(!$FullPath){
        $Switches = "$Switches -ep"
        Write-Verbose "Setting parameter -ep since switch FullPath is not set"
    }

    $Switches = "$Switches -m$CompressionLevel"
    Write-Verbose "Setting parameter -m$CompressionLevel with value from parameter CompressionLevel"
    $Switches = "$Switches -md$DictionarySize"
    Write-Verbose "Setting parameter -md$DictionarySize with value from parameter DictionarySize"
    $Switches = "$Switches -mt$Threads"
    Write-Verbose "Setting parameter -mt$Threads with value from parameter Threads"

    Write-Verbose "Calling WinRAR via command line: \
Start-Process -FilePath $(Get-WinRARPath -ErrorAction "Stop")\Rar.exe -ArgumentList a -ep $Switches $ArchivePath $DirectoryToCompress"
    Start-Process -FilePath "$(Get-WinRARPath -ErrorAction "Stop")\Rar.exe" -ArgumentList "a -ep $Switches $ArchivePath $DirectoryToCompress"
}

function Expand-WinRAR(){
    [CmdletBinding()]
    param(
    [Parameter(Position = 0, Mandatory = $true)]
    [String] $TargetDirectory,

    [Parameter(Position = 1, Mandatory = $true)]
    [String] $ArchivePath,

    [String] $Password,
    [Switch] $Confirm
    )

    if($Password){
        $Switches = "$Switches -p$Password"
        Write-Verbose "Setting parameter -p$Password since parameter Password is set"
    }
    if($Confirm){
        $Switches = "$Switches -y"
        Write-Verbose "Setting parameter -y since switch Confirm is set"
    }

    Write-Verbose "Calling WinRAR via command line: \
Start-Process -FilePath $(Get-WinRARPath -ErrorAction "Stop")\UnRAR.exe -ArgumentList x $Switches $ArchivePath $TargetDirectory"
    Start-Process -FilePath "$(Get-WinRARPath -ErrorAction "Stop")\UnRAR.exe" -ArgumentList "x $Switches $ArchivePath $TargetDirectory"
}

function Test-WinRAR(){
    [CmdletBinding()]
    param(
    [Parameter(Position = 0, Mandatory = $true)]
    [String] $ArchivePath,

    [String] $Password
    )

    if((Check-WinRAR -ArchivePath $ArchivePath -Password $Password) -eq 0){
        return $True
    }
    else{
        return $False
    }
}

function Check-WinRAR(){
    [CmdletBinding()]
    param(
    [Parameter(Position = 0, Mandatory = $true)]
    [String] $ArchivePath,

    [String] $Password
    )

    if($Password){
        $Switches = "$Switches -p$Password"
        Write-Verbose "Setting parameter -p$Password since parameter Password is set"
    }

    Write-Verbose "Calling WinRAR via command line: \
Start-Process -FilePath $(Get-WinRARPath -ErrorAction "Stop")\Rar.exe -ArgumentList t $Switches $ArchivePath -PassThru -Wait"
    return (Start-Process -FilePath "$(Get-WinRARPath -ErrorAction "Stop")\Rar.exe" -ArgumentList "t $Switches $ArchivePath" -PassThru -Wait).ExitCode
}

function Get-WinRARPath(){
    [CmdletBinding()]
    param()

    if(Test-Path "C:\Program Files\WinRAR\"){
        return "C:\Program Files\WinRAR\"
    }
    if(Test-Path "C:\Program Files (x86)\WinRAR\"){
        return "C:\Program Files (x86)\WinRAR\"
    }
    throw "WinRAR installation not found"
}

#Check for WinRAR directory on module install and abort installation when WinRAR is not installed
if(!(Get-WinRARPath -ErrorAction SilentlyContinue)){
    throw "WinRAR installation not found, aborting module installation"
}