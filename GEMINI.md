# GEMINI.md

## Project Overview
**CRMLSOG** is a specialized real estate automation tool built with **VB.NET** and **WPF** on **.NET 8.0**. It serves as a central hub for real estate agents to interact with multiple platforms simultaneously, scrape property data, manage leads, and automate outreach.

### Key Technologies
- **UI Framework:** WPF (Windows Presentation Foundation)
- **Runtime:** .NET 8.0 (Windows-specific)
- **Web Integration:** Microsoft Edge WebView2
- **Database:** SQLite (System.Data.SQLite)
- **Office Interop:** Microsoft Outlook and Excel via COM/Interop
- **Utilities:** 
  - **HtmlAgilityPack:** For parsing HTML from real estate portals.
  - **ClosedXML:** For generating and manipulating Excel reports.
  - **Newtonsoft.Json:** For JSON data handling.

## Building and Running

### Prerequisites
- **Visual Studio 2022** with the ".NET desktop development" workload.
- **.NET 8.0 SDK**.
- **WebView2 Runtime** (usually included with modern Windows/Edge).
- **Microsoft Outlook & Excel** installed for interop features.

### Build Commands
- **Restore Dependencies:** `dotnet restore`
- **Build Project:** `dotnet build`
- **Run Application:** `dotnet run --project CRMLS.vbproj`

## Project Structure
- `MainWindow.xaml / .vb`: The primary application dashboard containing the UI layout and orchestration logic.
- `AppHelpers.vb`: Contains specialized helper classes:
  - `DatabaseHelper`: SQLite CRUD operations.
  - `OutlookHelper`: Email automation and account management.
  - `WebViewHelper`: Script injection and data scraping logic for WebView2.
  - `ExcelHelper`: Data export functionality.
- `Common.vb`: General utility functions and logging.
- `DB/`: 
  - `crmls.db`: Local SQLite database for property and lead storage.
  - `login.xml`: Application settings, portal credentials, and email configuration.
  - `Attachment/`: Default directory for email attachments.

## Development Conventions
- **Helper Pattern:** Business logic and external integrations (DB, Office, Web) should be placed in `AppHelpers.vb` or new helper classes rather than `MainWindow.xaml.vb`.
- **Async Web Interaction:** Use `Await WebView2.ExecuteScriptAsync` for web scraping to prevent UI freezes.
- **Settings Management:** Sensitive information and configurable parameters are stored in `DB\login.xml`. Ensure this file is handled securely.
- **Error Logging:** All major operations should be wrapped in try-catch blocks, with exceptions passed to `fxCommon.GenerateLog(ex)`.
