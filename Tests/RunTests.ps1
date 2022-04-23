if(!(Get-InstalledModule "Pester")){
    Install-Module -Name "Pester" -RequiredVersion 5.2.0
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