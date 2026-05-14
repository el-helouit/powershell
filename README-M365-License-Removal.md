# 📄 M365 Licence Removal Script (Graph PowerShell)

## 📌 Overview

This script performs **bulk removal of Microsoft 365 licences** from users using Microsoft Graph PowerShell.

It:
- Reads users and licence SKUs from a CSV file
- Attempts to remove licences per user
- Handles **group-based licensing correctly**
- Outputs a **full execution log for audit and review**

---

## ✅ Key Behaviour

| Scenario | Outcome |
|--------|--------|
| Licence removed successfully | `Success` |
| Licence assigned via group | `Skipped - Group Assigned` |
| Invalid SKU / user issue | `Failed` |
| Dry Run enabled | `DryRun` |

---

## 📁 CSV Input File

### ✅ Required Format

```csv
UserId,SkuPartNumber
