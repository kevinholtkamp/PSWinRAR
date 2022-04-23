BeforeAll{
    Import-Module "$(Get-Location)"
    if(!(Get-WinRARPath -ErrorAction "silentlycontinue")){
        Invoke-WebRequest -Uri 'https://www.win-rar.com/fileadmin/winrar-versions/winrar/winrar-x64-611.exe' -OutFile "$($env:TMP)\winrar.exe"
        Start-Process -FilePath "$($env:TMP)\winrar.exe" -ArgumentList "/S" | Out-Null
    }
}
Describe "Main functions"{
    BeforeAll{
        $TestFolder = "$($env:TEMP)\PSWinRAR_Pester_Tests"
        Remove-Item $TestFolder -Recurse -Force -ErrorAction "silentlycontinue"
        New-Item "$TestFolder\FolderToCompress\InnerFolder\" -ItemType directory -Force
        New-Item "$TestFolder\FolderToCompress\InnerFolder\InnerFile.txt" -ItemType file
        Set-Content "$TestFolder\FolderToCompress\InnerFolder\InnerFile.txt" "Content of test File"
    }
    AfterAll{
        Remove-Item $TestFolder -Recurse -Force
    }
    Context "Main functions working normally"{
        It "Compressing files"{
            Compress-WinRAR -DirectoryToCompress "$TestFolder\FolderToCompress" -ArchivePath "$TestFolder\Archive.rar"
            "$TestFolder\Archive.rar" | Should -Exist
        }
        It "Testing archive"{
            Test-WinRAR -ArchivePath "$TestFolder\Archive.rar" -GetReturnCode | Should -Be 0
            Test-WinRAR -ArchivePath "$TestFolder\Archive.rar" | Should -Be $true
        }
        It "Expanding archive"{
            Expand-WinRAR -ArchivePath "$TestFolder\Archive.rar" -TargetDirectory "$TestFolder\TargetDirectory\"
            "$TestFolder\TargetDirectory\FolderToCompress\InnerFolder\InnerFile.txt" | Should -Exist
            "$TestFolder\TargetDirectory\FolderToCompress\InnerFolder\InnerFile.txt" | Should -FileContentMatch "Content of test File"
        }
    }
    Context "Testing repair functionality"{
        BeforeEach{
            Compress-WinRAR -DirectoryToCompress "$TestFolder\FolderToCompress" -ArchivePath "$TestFolder\RecoveryArchive.rar" -RecoveryPercentage $Recovery
        }
        AfterEach{
            Remove-Item "$TestFolder\RecoveryArchive.rar" -ErrorAction "silentlycontinue"
            Remove-Item "$TestFolder\fixed.RecoveryArchive.rar" -ErrorAction "silentlycontinue"
            Remove-Item "$TestFolder\rebuilt.RecoveryArchive.rar" -ErrorAction "silentlycontinue"
        }
        It "Repair-WinRAR with recoverable errors: <Recovery>% recovery data and removing <Bytes> byte" -ForEach @(
            @{Recovery = 20; Bytes = 1}
            @{Recovery = 40; Bytes = 1}
            @{Recovery = 10; Bytes = 1}
            @{Recovery = 1; Bytes = 20}
            @{Recovery = 1; Bytes = 150}
            @{Recovery = 1; Bytes = 100}
            @{Recovery = 300; Bytes = 200}
        ){
            "$TestFolder\RecoveryArchive.rar" | Should -Exist
            Set-Content -Path "$TestFolder\RecoveryArchive.rar" -NoNewLine -Value (Get-Content "$TestFolder\RecoveryArchive.rar" -Raw).Remove(101,$Bytes)
            Repair-WinRAR -ArchivePath "$TestFolder\RecoveryArchive.rar" -GetReturnCode | Should -Be 3
            "$TestFolder\fixed.RecoveryArchive.rar" | Should -Exist
            "$TestFolder\rebuilt.RecoveryArchive.rar" | Should -Not -Exist
        }
        It "Repair-WinRAR with too little recovery data: <Recovery>% recovery data and removing <Bytes> byte" -ForEach @(
            @{Recovery = 1; Bytes = 200}
            @{Recovery = 100; Bytes = 300}
        ){
            Set-Content -Path "$TestFolder\RecoveryArchive.rar" -NoNewLine -Value (Get-Content "$TestFolder\RecoveryArchive.rar" -Raw).Remove(101,$Bytes)
            Repair-WinRAR -ArchivePath "$TestFolder\RecoveryArchive.rar" -GetReturnCode | Should -BeIn @(0,3)
            "$TestFolder\rebuilt.RecoveryArchive.rar" | Should -Exist
            "$TestFolder\fixed.RecoveryArchive.rar" | Should -Not -Exist
        }
        It "Repair-WinRAR without recovery data: <Recovery>% recovery data and removing <Bytes> byte" -ForEach @(
            @{Recovery = 0; Bytes = 20}
            @{Recovery = 0; Bytes = 1}
        ){
            Set-Content -Path "$TestFolder\RecoveryArchive.rar" -NoNewLine -Value (Get-Content "$TestFolder\RecoveryArchive.rar" -Raw).Remove(101,$Bytes)
            Repair-WinRAR -ArchivePath "$TestFolder\RecoveryArchive.rar" -GetReturnCode | Should -BeIn @(0,3)
            "$TestFolder\rebuilt.RecoveryArchive.rar" | Should -Exist
            "$TestFolder\fixed.RecoveryArchive.rar" | Should -Not -Exist
        }
    }
    Context "Compress-WinRAR with out-of-spec parameters"{
        BeforeAll{
            $Splat = @{
                DirectoryToCompress = "$TestFolder\FolderToCompress"
                ArchivePath = "$TestFolder\Archive.rar"
            }
        }
        AfterAll{
            Remove-Item "$TestFolder\Archive.rar" -ErrorAction "silentlycontinue"
        }
        It "Terminating wrong parameters"{
            {Compress-WinRAR @Splat -CompressionLevel 6} | Should -Throw
            {Compress-WinRAR @Splat -CompressionLevel -5} | Should -Throw
            {Compress-WinRAR @Splat -DictionarySize 0} | Should -Throw
            {Compress-WinRAR @Splat -DictionarySize 2000} | Should -Throw
            {Compress-WinRAR @Splat -DictionarySizeUnit "t"} | Should -Throw
            {Compress-WinRAR @Splat -Threads 0} | Should -Throw
            {Compress-WinRAR @Splat -Threads ($env:NUMBER_OF_PROCESSORS * 2)} | Should -Throw
            {Compress-WinRAR @Splat -ArchiveFileStructure "INVALID"} | Should -Throw
            {Compress-WinRAR @Splat -RecoveryPercentage -10} | Should -Throw
            {Compress-WinRAR @Splat -PassThruParameters "-a -ab"} | Should -Throw
        }
        It "Recoverable wrong parameters"{
            Compress-WinRAR @Splat -PassThruParameters "/INVALID" | Should -Exist
        }
    }
    Context "Main functions with missing files"{
        It "Compressing non-existing folder"{
            {Compress-WinRAR -DirectoryToCompress "$TestFolder\NonExistingFolder" -ArchivePath "$TestFolder\Archive.rar"} | Should -Throw "Winrar stopped with exit code 10"
            "$TestFolder\Archive.rar" | Should -Not -Exist
        }
        It "Expanding non-existing archive"{
            {Expand-WinRAR -ArchivePath "$TestFolder\Archive.rar" -TargetDirectory "$TestFolder\TargetDirectory"} | Should -Throw "Winrar stopped with exit code 10"
        }
        It "Testing non-existing archive"{
            Test-WinRAR -ArchivePath "$TestFolder\Archive.rar" -GetReturnCode | Should -Be 10
        }
        It "Repairing non-existing winrar"{
            Repair-WinRAR -ArchivePath "$TestFolder\Archive.rar" -GetReturnCode | Should -Be 10
        }
    }
}