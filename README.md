# PS-Export-PowerBI-Audit-Log

![PowerShell Gallery Version](https://img.shields.io/powershellgallery/v/PS-Export-PowerBI-Audit-Log)
![Downloads](https://img.shields.io/powershellgallery/dt/PS-Export-PowerBI-Audit-Log)
![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20PowerShell%20Core-blue)

> A PowerShell module for exporting **Power BI audit logs** using the Microsoft Graph Security API.

---

## ✨ Features

* Submits asynchronous audit log search queries via Microsoft Graph (`beta` endpoint)
* Filters audit log searches to only return **PowerBIAudit** records
* Waits for search completion and paginates results
* Returns raw log records for export, inspection, or further processing

---

## 📦 Module Functions

| Function                     | Description                                              |
| ---------------------------- | -------------------------------------------------------- |
| `New-PBIAuditLogQuery`       | Creates a new Power BI audit log search query job        |
| `Get-PBIAuditLogQuery`       | Retrieves details about one or more audit log query jobs |
| `Get-PBIAuditLogQueryRecord` | Retrieves records/results of a completed audit log query |

---

## 🔧 Requirements

* PowerShell 5.1 or later (PowerShell 7+ recommended)
* [Microsoft.Graph](https://learn.microsoft.com/powershell/microsoftgraph/overview) PowerShell module

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

---

## 🔐 Permissions

You must authenticate with Microsoft Graph using delegated or application permissions that include:

* `AuditLog.Read.All`

```powershell
Connect-MgGraph -Scopes "AuditLog.Read.All"
```

---

## 🚀 Getting Started

### 1. Create a new audit log query

```powershell
$query = New-PBIAuditLogQuery -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date)
```

### 2. Wait and retrieve query records

```powershell
$records = Get-PBIAuditLogQueryRecord -Id $query.id -Wait 30
```

### 3. Export results to CSV

```powershell
$records | Export-Csv -Path "PBIAuditLogs.csv" -NoTypeInformation
```

---

## 📘 Function Details

### `New-PBIAuditLogQuery`

Creates a new asynchronous search query for Power BI audit logs.

#### Parameters

| Name         | Type                | Description                             |
| ------------ | ------------------- | --------------------------------------- |
| `StartDate`  | `DateTime`          | Start of the audit log time range (UTC) |
| `EndDate`    | `DateTime`          | End of the audit log time range (UTC)   |
| `SearchName` | `String` (optional) | Custom display name for the search      |

Returns: the created query job (object with `id`, `status`, etc.)

---

### `Get-PBIAuditLogQuery`

Retrieves the status of a search query or list of queries.

#### `Get-PBIAuditLogQuery` Parameters

| Name                      | Type                | Description                                         |
| ------------------------- | ------------------- | --------------------------------------------------- |
| `Id`                      | `String` (optional) | Specific query ID to retrieve                       |
| `IncludeNonPBIAuditQuery` | `Switch`            | Include queries that are not Power BI audit queries |

Returns: query object(s)

---

### `Get-PBIAuditLogQueryRecord`

Retrieves results from a **completed** audit log query.

#### `Get-PBIAuditLogQueryRecord` Parameters

| Name   | Type             | Description                                                                         |
| ------ | ---------------- | ----------------------------------------------------------------------------------- |
| `Id`   | `String`         | ID of the audit query job                                                           |
| `Wait` | `Int` (optional) | Wait interval (in seconds) to poll the query job until it completes (range: 30–300) |

Returns: list of audit log records

---

## 📌 Example

```powershell
Connect-MgGraph -Scopes "AuditLog.Read.All"

# Step 1: Submit the query
$query = New-PBIAuditLogQuery `
    -StartDate (Get-Date).AddHours(-6) `
    -EndDate (Get-Date)

# Step 2: Wait and get results
$records = Get-PBIAuditLogQueryRecord -Id $query.id -Wait 60

# Step 3: Save results
$records | Export-Csv -Path ".\PowerBIAuditRecords.csv" -NoTypeInformation
```

---

## 🥮 Notes

* The search is **asynchronous**. You must wait for it to complete before retrieving records.
* Pagination is automatically handled in `Get-PBIAuditLogQueryRecord` via `@odata.nextLink`.
* Records are retrieved using Microsoft Graph's `/security/auditLog/queries/{id}/records` endpoint.

---

## 🛠 Troubleshooting

* If no records are returned, ensure:

  * You're licensed for audit logging (E5 or equivalent)
  * Audit logs are enabled in Microsoft Purview
  * The time range is valid and data exists

---

## 📄 License

MIT License. See [LICENSE](LICENSE) for full details.

---

## 👤 Author

Developed by [June Castillote](https://github.com/junecastillote)
