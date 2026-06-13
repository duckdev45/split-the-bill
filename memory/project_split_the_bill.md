---
name: project-split-the-bill
description: Context for the 美西分帳 2026 app — trip participants, currencies, tech stack, and Google Sheets integration plan
metadata:
  type: project
---

鴨子/三三/虎妞 的美西旅行分帳系統（2026/06/19–06/29）。

**Why:** Trip to US West Coast; 3 people with different home currencies (鴨子=TWD, 三三/虎妞=CNY), all also use USD. Exchange rates entered at end of trip for final settlement.

**Tech:** Single HTML file (`index.html`) with vanilla JS + localStorage. Press Start 2P pixel-art style (8-bit aesthetic).

**Data structure key fields per expense:**
- `id`, `date`, `desc`, `cat`, `paidBy`, `cur`, `amt`, `among`, `splitType` (equal|custom), `custom` (obj), `notes`
- Google Sheets reserved: `gsRowId`, `gsSyncedAt`

**Google Sheets integration (not yet built):**
- `S.gs` = `{ spreadsheetId, sheetName, scriptUrl, lastSyncAt }` already saved to localStorage
- Plan: Apps Script Web App URL → POST expenses, GET sync

**Settlement algorithm:** Greedy debt-minimization. All amounts converted to USD via TWD_USD / CNY_USD rates. Transactions shown with home-currency equivalents.

**How to apply:** When continuing this project, check if GS integration is the next milestone and design the Apps Script endpoint to accept the existing expense JSON shape.
