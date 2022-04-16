function Compress-WinRAR(){
    param(
    [String] $CompressPath,
    [String] $DestinationPath,
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
    }
    if($Delete){
        $Switches = "$Switches -df"
    }
    if($SafeDelete){
        $Switches = "$Switches -dr"
    }
    if($IgnoreEmptyDirectories){
        $Switches = "$Switches -ed"
    }
    if($ErrorLogFile){
        $Switches = "$Switches -ilog$ErrorLogFile"
    }
    if($FileMask){
        $Switches = "$Switches -n$FileMask"
    }
    if($Password){
        $Switches = "$Switches -p$Password"
    }
    if($RecoveryPercentage){
        $Switches = "$Switches -rr$($RecoveryPercentage)p"
    }
    if($SolidArchive){
        $Switches = "$Switches -s"
    }
    if($TestData){
        $Switches = "$Switches -t"
    }
    if($Confirm){
        $Switches = "$Switches -y"
    }
    if(!$FullPath){
        $Switches = "$Switches -ep"
    }

    $Switches = "$Switches -m$CompressionLevel"
    $Switches = "$Switches -md$DictionarySize"
    $Switches = "$Switches -mt$Threads"

    Start-Process -FilePath "C:\Program Files\WinRAR\Rar.exe" -ArgumentList "a -ep $Switches $DestinationPath $CompressPath"
}

function Expand-WinRAR(){
    param(
    [String] $TargetPath,
    [String] $ArchivePath,
    [String] $Password,
    [Switch] $Confirm
    )

    if($Password){
        $Switches = "$Switches -p$Password"
    }
    if($Password){
        $Switches = "$Switches -y"
    }

    Start-Process -FilePath "C:\Program Files\WinRAR\UnRAR.exe" -ArgumentList "x $Switches $ArchivePath $TargetPath"
}

function Validate-WinRAR(){
    param(
        [String] $ArchivePath,
        [String] $Password
    )

    if($Password){
        $Switches = "$Switches -p$Password"
    }

    if((Start-Process -FilePath "C:\Program Files\WinRAR\Rar.exe" -ArgumentList "t $Switches $ArchivePath" -PassThru -Wait).ExitCode -eq 0){
        return $True
    }
    else{
        return $False
    }
}

function Test-WinRAR(){
    param(
        [String] $ArchivePath,
        [String] $Password
    )

    if($Password){
        $Switches = "$Switches -p$Password"
    }

    return (Start-Process -FilePath "C:\Program Files\WinRAR\Rar.exe" -ArgumentList "t $Switches $ArchivePath" -PassThru -Wait).ExitCode
}