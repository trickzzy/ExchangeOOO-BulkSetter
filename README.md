[![CI](https://github.com/trickzzy/ExchangeOOO-BulkSetter/actions/workflows/ci.yml/badge.svg)](https://github.com/trickzzy/ExchangeOOO-BulkSetter/actions/workflows/ci.yml)

# ExchangeOOO-BulkSetter

Bulk-enable Exchange Online automatic replies (OOO) for multiple mailboxes using a CSV file, with logging.

## Prerequisites
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Connect-ExchangeOnline