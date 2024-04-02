$Rep = @{
    Name = "CustomRepository"
    SourceLocation = "\\10.0.0.2\PowershellRepository\"
    PublishLocation = "\\10.0.0.2\PowershellRepository\"
    InstallationPolicy = "Trusted"
}
Register-PSRepository @Rep -ErrorAction SilentlyContinue
if(!(Get-InstalledModule "Pester" -RequiredVersion 5.3.2)){
    Install-Module -Name "Pester" -Repository "CustomRepository" -RequiredVersion 5.3.2 -Force
}

$config = New-PesterConfiguration -HashTable @{
    Run = @{
        Path = "$(Get-Location)\Tests"
    }
#    CodeCoverage = @{
#        Enabled = $true
#        OutputPath = "./Tests/"
#    }
#    TestResult = @{
#        Enabled = $true
#        OutputPath = "./Tests/"
#    }
    Should = @{
        ErrorAction = 'Continue'
    }
    Output = @{
        Verbosity = "Detailed"
    }
}


Invoke-Pester -Configuration $config