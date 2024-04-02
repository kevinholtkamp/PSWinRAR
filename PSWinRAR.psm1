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
    .PARAMETER PassThru
        Return the archive path
    .PARAMETER ExitCode
        Return the exit code from the winrar cli
    .INPUTS
        Pipeline inputs get used as DirectoryToCompress
    .OUTPUTS
        Returns the return code of the winrar command line tool when -ExitCode is provided, returns the archive path if -PassThru is provided, otherwise returns true if no error occured, false if an error occured
    .EXAMPLE
        Compress-WinRAR ./Directory/ ./Archive.rar -Threads 8
    #>

    #DefaultParameterSetName in case neither PassThru nor ExitCode gets switched, it doesn't matter though
    [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = "PassThru")]
    param(
    [Parameter(Position = 0, Mandatory, ValueFromPipeline)]
    [Alias("Directory")]
    [String] $DirectoryToCompress,

    [Alias("Archive")]
    [ValidatePattern(".*\.rar")]
    [String] $ArchivePath,

    [Parameter(Position = 2)]
    [ValidateRange(0,5)]
    [Byte] $CompressionLevel = 3,

    [Parameter(Position = 3)]
    [ValidateRange(1,[Int32]::MaxValue)]
    [Int32] $DictionarySize = 128,

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

    [ValidateSet("Store", "Lowest", "Low", "Medium", "High", "Highest")]
    [String] $Preset,

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
    [Switch] $UseRAR4,

    [Parameter(ParameterSetName = "PassThru")]
    [Switch] $PassThru,
    [Parameter(ParameterSetName = "ExitCode")]
    [Alias("ReturnCode")]
    [Switch] $ExitCode
    )
    Begin{
        $PresetValues = @{
            Store = @{
                CompressionLevel = 0
                DictionarySize = 0
                DictionarySizeUnit = "k"
                Threads = 1
            }
            Lowest = @{
                CompressionLevel = 1
                DictionarySize = 128
                DictionarySizeUnit = "k"
                Threads = 1
            }
            Low = @{
                CompressionLevel = 2
                DictionarySize = 1
                DictionarySizeUnit = "m"
                Threads = if($Env:NUMBER_OF_PROCESSORS -gt 1){
                    2
                }
                else{
                    1
                }
            }
            Medium = @{
                CompressionLevel = 3
                DictionarySize = 128
                DictionarySizeUnit = "m"
                Threads = if($Env:NUMBER_OF_PROCESSORS -gt 1){
                    2
                }
                else{
                    1
                }
            }
            High = @{
                CompressionLevel = 4
                DictionarySize = if([Math]::Truncate((Get-CIMInstance Win32_OperatingSystem -Verbose:$false -Debug:$false).FreePhysicalMemory / 8000) -gt 1024){
                    1024
                }
                else{
                    [Math]::Truncate((Get-CIMInstance Win32_OperatingSystem -Verbose:$false -Debug:$false).FreePhysicalMemory / 8000)
                }
                DictionarySizeUnit = "m"
                Threads = if($Env:NUMBER_OF_PROCESSORS -gt 4){
                    4
                }
                else{
                    $Env:NUMBER_OF_PROCESSORS - 1
                }
            }
            Highest = @{
                CompressionLevel = 5
                DictionarySize = if([Math]::Truncate((Get-CIMInstance Win32_OperatingSystem -Verbose:$false -Debug:$false).FreePhysicalMemory / 4000) -gt 1024){
                    1024
                }
                else{
                    [Math]::Truncate((Get-CIMInstance Win32_OperatingSystem -Verbose:$false -Debug:$false).FreePhysicalMemory / 4000)
                }
                DictionarySizeUnit = "m"
                Threads = if($Env:NUMBER_OF_PROCESSORS -gt 1){
                    $Env:NUMBER_OF_PROCESSORS - 1
                }
                else{
                    1
                }
            }
        }
        if($Preset){
            Write-Verbose "Applying preset $Preset"
            "CompressionLevel", "DictionarySize", "DictionarySizeUnit", "Threads" | ForEach-Object {
                if(!$PSBoundParameters[$_]){
                    Write-Verbose "Setting $_ to $($PresetValues[$Preset][$_]) because of preset $Preset"
                    Set-Variable -Name $_ -Value $PresetValues[$Preset][$_] -Scope "Local"
                }
            }
        }

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
        else{
            $Switches = "$Switches -ma5"
            Write-Verbose "Setting parameter -ma5 since switch UseRAR4 is not set"
        }
        switch ($ArchiveFileStructure){
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

        $Paths = [System.Collections.ArrayList]@()
    }
    Process{
        $Paths.Add($DirectoryToCompress) | Out-Null
    }
    End{
        $FilePath = "$(Get-WinRARPath -ErrorAction "Stop")\Rar.exe"
        $ArgumentList = "a $PassThruParameters $Switches -IBCK `"$ArchivePath`" $($Paths | ForEach-Object {"'$_'"}) > null"
        $Expression = "(Start-Process -FilePath `"$FilePath`" -ArgumentList `"$($ArgumentList.Replace('"', '``"'))`" -Wait -PassThru -WindowStyle `"Hidden`").ExitCode"

        Write-Verbose "Calling WinRAR via command line: $Expression"
        $ExitCodeInternal = Invoke-Expression -Command $Expression

        if($ExitCodeInternal -ne 0 -and $PassThru){
            throw "Winrar stopped with exit code $ExitCodeInternal"
        }

        if($PassThru){
            return $ArchivePath
        }
        elseif($ExitCode){
            return $ExitCodeInternal
        }
        else{
            if($ExitCodeInternal -eq 0){
                return $true
            }
            else{
                return $false
            }
        }
    }
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

    Begin{
        if(!(Test-Path $TargetDirectory)){
            New-Item $TargetDirectory -ItemType Directory -Force | Out-Null
        }

        if($Password){
            $Switches = "$Switches -p$Password"
            Write-Verbose "Setting parameter -p$Password since parameter Password is set"
        }
        if($Confirm){
            $Switches = "$Switches -y"
            Write-Verbose "Setting parameter -y since switch Confirm is set"
        }

        $ExitCodes = [System.Collections.ArrayList]@()
    }
    Process{
        Write-Verbose "Calling WinRAR via command line: Start-Process -FilePath '$(Get-WinRARPath -ErrorAction "Stop")\UnRAR.exe' -ArgumentList 'x $Switches $ArchivePath $TargetDirectory' -Wait -PassThru"
        $ExitCode = (Start-Process -FilePath "$(Get-WinRARPath -ErrorAction "Stop")\UnRAR.exe" -ArgumentList "x $Switches $ArchivePath $TargetDirectory" -Wait -PassThru).ExitCode
        if($ExitCode -ne 0){
            Write-Error "Winrar stopped with exit code $ExitCode"
        }
        $ExitCodes.Add($ExitCode) | Out-Null
    }
    End{
        return $ExitCodes
    }
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
    .PARAMETER ExitCode
        Pass through the return code from the winrar command line tool instead of $true/$false
    .INPUTS
        Pipeline inputs get used as ArchivePath
    .OUTPUTS
        Returns $true if the archive is valid, $false if the archive is invalid
        If the parameter -ExitCode is set, the returncode from the winrar command line tool will be passed through instead
    .EXAMPLE
        Test-WinRAR ./Archive.rar -ExitCode
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "Path to .rar file which will be tested")]
        [Alias("Archive", "Path")]
        [String] $ArchivePath,

        [Parameter(Position = 1, HelpMessage = "Password to .rar file")]
        [String] $Password,

        [Parameter(HelpMessage = "Return the returncode from the winrar executable instead of true/false")]
        [Alias("ReturnCode")]
        [Switch] $ExitCode
    )
    Begin{
        if($Password){
            $Switches = "$Switches -p$Password"
            Write-Verbose "Setting parameter -p$Password since parameter Password is set"
        }
        $ExitCodes = [System.Collections.ArrayList]@()
    }
    Process{
        Write-Verbose "Before: $ExitCodes"
        Write-Verbose "Calling WinRAR via command line: Start-Process -FilePath $(Get-WinRARPath -ErrorAction "Stop")\Rar.exe -ArgumentList t $Switches $ArchivePath -PassThru -Wait"
        $ExitCodes.Add((Start-Process -FilePath "$(Get-WinRARPath -ErrorAction "Stop")\Rar.exe" -ArgumentList "t $Switches $ArchivePath" -PassThru -Wait).ExitCode) | Out-Null
        Write-Verbose "After: $ExitCodes"
    }
    End{
        if($ExitCode){
            Write-Verbose "End: $($ExitCodes.Length)"
            return $ExitCodes
        }
        else{
            if($ExitCodes | Where-Object {$_ -ne 0}){
                return $false
            }
            else{
                return $true
            }
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

        [Parameter(HelpMessage = "Return the exit code from the winrar executable instead of true/false")]
        [Alias("ReturnCode")]
        [Switch] $ExitCode
    )
    Begin{
        if($Password){
            $Switches = "$Switches -p$Password"
            Write-Verbose "Setting parameter -p$Password since parameter Password is set"
        }
        $ExitCodes = [System.Collections.ArrayList]@()
    }
    Process{
        Write-Verbose "Calling WinRAR via command line: Start-Process -FilePath $(Get-WinRARPath -ErrorAction "Stop")\Rar.exe -ArgumentList r $Switches $ArchivePath -PassThru -Wait"
        Push-Location (Split-Path $ArchivePath)
        $ExitCodes.Add((Start-Process -FilePath "$(Get-WinRARPath -ErrorAction "Stop")\Rar.exe" -ArgumentList "r $Switches $ArchivePath" -PassThru -Wait).ExitCode) | Out-Null
        Pop-Location
    }
    End{
        if($ExitCode){
            return $ExitCodes
        }
        else{
            if($ExitCodes | Where-Object {$_ -ne 0}){
                return $false
            }
            else{
                return $true
            }
        }
    }
}

function Get-WinRARPath(){
    [CmdletBinding()]
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