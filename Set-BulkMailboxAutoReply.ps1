<#
.SYNOPSIS
Bulk-enable Exchange Online automatic replies (OOO) for multiple mailboxes from a CSV mapping file.

.DESCRIPTION
Reads a CSV with mailbox + internal/external message content, applies Set-MailboxAutoReplyConfiguration,
and writes a results log CSV.

PREREQS
- ExchangeOnlineManagement module
- Connected session to Exchange Online (Connect-ExchangeOnline)
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ })]
    [string]$CsvPath,

    [Parameter()]
    [ValidateSet('All','Known','None')]
    [string]$ExternalAudience = 'All',

    [Parameter()]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$LogDirectory = $env:TEMP
)

function Assert-ExchangeConnected {
    if (-not (Get-Command Set-MailboxAutoReplyConfiguration -ErrorAction SilentlyContinue)) {
        throw "Set-MailboxAutoReplyConfiguration not found. Connect to Exchange Online and/or install ExchangeOnlineManagement."
    }
}

Assert-ExchangeConnected

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logPath = Join-Path $LogDirectory "OOO_SetResults_$timestamp.csv"

$items   = Import-Csv -Path $CsvPath
$results = @()

foreach ($row in $items) {
    $mailbox  = $row.Mailbox
    $internal = $row.InternalMessage
    $external = if ([string]::IsNullOrWhiteSpace($row.ExternalMessage)) { $internal } else { $row.ExternalMessage }

    if ($PSCmdlet.ShouldProcess($mailbox, "Enable OOO and set internal/external messages")) {
        try {
            Set-MailboxAutoReplyConfiguration `
                -Identity $mailbox `
                -AutoReplyState Enabled `
                -ExternalAudience $ExternalAudience `
                -InternalMessage $internal `
                -ExternalMessage $external `
                -ErrorAction Stop

            $results += [pscustomobject]@{
                Mailbox = $mailbox
                Status  = "Success"
                When    = Get-Date
                Notes   = ""
            }
        }
        catch {
            $results += [pscustomobject]@{
                Mailbox = $mailbox
                Status  = "Failed"
                When    = Get-Date
                Notes   = $_.Exception.Message
            }
        }
    }
}

$results | Export-Csv -NoTypeInformation -Path $logPath
Write-Output "Done. Log saved to: $logPath"
$results | Sort-Object Status,Mailbox