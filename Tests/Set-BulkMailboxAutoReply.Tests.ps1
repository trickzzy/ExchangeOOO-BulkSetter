$ErrorActionPreference = 'Stop'

Describe "ExchangeOOO-BulkSetter" {
    It "script file exists" {
        Test-Path "$PSScriptRoot\..\Set-BulkMailboxAutoReply.ps1" | Should -BeTrue
    }

    It "has CmdletBinding" {
        $content = Get-Content "$PSScriptRoot\..\Set-BulkMailboxAutoReply.ps1" -Raw
        $content | Should -Match '\[CmdletBinding'
    }
}