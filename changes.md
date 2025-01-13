# Changes

## Version 2.0 (2025-01-13)

- Made into a module.
- Removed dependency on Exchange Online and PnP PowerShell modules.
- Operations are now handeled with Microsoft Graph API calls.

## Version 1.1 (2024-04-30)

- Added the `-SplitTimeIntoChunksOf` parameter with the default value of 1.
  - This parameter indicates the number of chunks to split the duration of the search period.
  - The purpose is to prevent reaching the maximum search limit of 50,000 of the `Search-UnifiedAuditLog` command.

## Version 1.0

- Initial release.
