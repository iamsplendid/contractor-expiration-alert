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
