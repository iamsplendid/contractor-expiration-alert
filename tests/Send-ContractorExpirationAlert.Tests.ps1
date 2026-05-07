BeforeAll {
    # Keep in sync with Send-ContractorExpirationAlert.ps1
    function Test-IsGuid {
        param([string]$Value)
        return $Value -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
    }

    # Keep in sync with Send-ContractorExpirationAlert.ps1
    function Test-IsEmail {
        param([string]$Value)
        return $Value -match '^[^@\s]+@[^@\s]+\.[^@\s]+$'
    }
}

Describe 'Test-IsGuid' {
    It 'returns true for a valid lowercase GUID' {
        Test-IsGuid 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' | Should -Be $true
    }
    It 'returns true for a valid uppercase GUID' {
        Test-IsGuid 'AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE' | Should -Be $true
    }
    It 'returns true for a mixed-case GUID' {
        Test-IsGuid 'Aaaaaaaa-BBBB-cccc-DDDD-eeeeeeeeeeee' | Should -Be $true
    }
    It 'returns false for an empty string' {
        Test-IsGuid '' | Should -Be $false
    }
    It 'returns false for a plain word' {
        Test-IsGuid 'not-a-guid' | Should -Be $false
    }
    It 'returns false for a partial GUID' {
        Test-IsGuid 'aaaaaaaa-bbbb-cccc' | Should -Be $false
    }
}

Describe 'Test-IsEmail' {
    It 'returns true for a valid email address' {
        Test-IsEmail 'user@domain.com' | Should -Be $true
    }
    It 'returns true for a UPN-style address' {
        Test-IsEmail 'alerts@contoso.onmicrosoft.com' | Should -Be $true
    }
    It 'returns false for a string without @' {
        Test-IsEmail 'nodomain' | Should -Be $false
    }
    It 'returns false for @ at the start with no local part' {
        Test-IsEmail '@nodomain' | Should -Be $false
    }
    It 'returns false for @ at the end with no domain' {
        Test-IsEmail 'user@' | Should -Be $false
    }
    It 'returns false for a domain with no dot' {
        Test-IsEmail 'user@nodot' | Should -Be $false
    }
    It 'returns false for an empty string' {
        Test-IsEmail '' | Should -Be $false
    }
}

Describe 'Save-AlertConfig / Read-AlertConfig' {
    BeforeAll {
        # Keep in sync with Send-ContractorExpirationAlert.ps1
        function Get-AlertConfigPath {
            return Join-Path $PSScriptRoot "config\$env:USERNAME.xml"
        }

        # Keep in sync with Send-ContractorExpirationAlert.ps1
        function Read-AlertConfig {
            param([string]$Path = (Get-AlertConfigPath))
            if (-not (Test-Path $Path)) { return $null }
            return Import-Clixml -Path $Path
        }

        # Keep in sync with Send-ContractorExpirationAlert.ps1
        function Save-AlertConfig {
            param([hashtable]$Config, [string]$Path = (Get-AlertConfigPath))
            $dir = Split-Path $Path
            if (-not (Test-Path $dir)) { New-Item $dir -ItemType Directory | Out-Null }
            $Config | Export-Clixml -Path $Path
        }
    }

    It 'round-trips all four fields including the encrypted secret' {
        $tempPath = Join-Path $TestDrive "config\$env:USERNAME.xml"
        $null = New-Item (Split-Path $tempPath) -ItemType Directory -Force

        $config = @{
            TenantId     = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
            ClientId     = 'ffffffff-0000-1111-2222-333333333333'
            ClientSecret = ConvertTo-SecureString 'test-secret' -AsPlainText -Force
            FromAddress  = 'alerts@test.com'
        }
        Save-AlertConfig -Config $config -Path $tempPath
        Test-Path $tempPath | Should -Be $true
        $loaded = Read-AlertConfig -Path $tempPath
        $loaded.ClientSecret | Should -BeOfType [System.Security.SecureString]

        $loaded.TenantId    | Should -Be 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
        $loaded.ClientId    | Should -Be 'ffffffff-0000-1111-2222-333333333333'
        $loaded.FromAddress | Should -Be 'alerts@test.com'

        $ptr   = [Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($loaded.ClientSecret)
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringUni($ptr)
        [Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($ptr)
        $plain | Should -Be 'test-secret'
    }

    It 'returns null when the config file does not exist' {
        Read-AlertConfig -Path (Join-Path $TestDrive 'nonexistent.xml') | Should -BeNullOrEmpty
    }
}
