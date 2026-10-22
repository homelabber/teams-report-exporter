# Teams Report Exporter

A portable, browser-based prototype for discovering and exporting Teams Voice users. It is intentionally dependency-free so it can be opened locally with `index.html` or packaged into a Tauri desktop application later.

## Run locally

Double-click `index.html`, or serve the folder with any static web server. No credentials are collected by this prototype; the Microsoft 365 sign-in flow is simulated.

## Included MVP flow

- Microsoft 365 connection screen (simulated)
- Report catalog showing only the supported Teams Voice users report
- User-status, location, and license filters
- Preview table
- CSV export
- Excel-compatible `.xls` export
- Local export history stored in browser storage
- Responsive layout suitable for a Tauri desktop shell

## Desktop build

The `Desktop` folder contains a Windows Forms host using Microsoft Edge WebView2. Build a self-contained Windows executable with:

```powershell
dotnet publish Desktop\TeamsReportExporterDesktop.csproj --configuration Release --runtime win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o Desktop\publish\win-x64
```

The portable executable is written to `Desktop\publish\win-x64\TeamsReportExporterDesktop.exe`. The target computer needs the Microsoft Edge WebView2 runtime, which is present on most current Windows installations.

## Customer-owned Entra ID setup

The desktop app does not use a shared multi-tenant registration. Each customer creates an app registration in their own Entra tenant and enters its two identifiers in **Settings → Configure tenant**.

1. Open **Microsoft Entra admin center → App registrations → New registration**.
2. Choose **Accounts in this organizational directory only (single tenant)**.
3. Under **Authentication**, add a **Mobile and desktop applications** platform and enable `https://login.microsoftonline.com/common/oauth2/nativeclient`.
4. Under **API permissions → Microsoft Graph → Delegated permissions**, add `User.Read`, `User.Read.All`, and `Reports.Read.All`.
5. Grant admin consent for the tenant.
6. Copy the **Application (client) ID** and **Directory (tenant) ID** into the application setup screen.

The app then signs the user in with MSAL.NET and calls the Microsoft Graph Teams user activity report for the last 30 days. The customer’s tenant administrator controls all access.

Microsoft Graph’s currently documented report endpoint provides Teams user activity, not a complete Teams Phone inventory. The adapter is deliberately isolated so a supported Teams Voice endpoint can be added when Microsoft exposes one.
