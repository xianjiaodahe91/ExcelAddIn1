# Excel Worksheet Navigator

> A side navigation bar for Excel — a VSTO + WPF add-in that puts every worksheet one click away.

**English** | [简体中文](README_zh.md)

## Overview

Long workbooks turn into a tab-scrolling exercise: tabs shrink, hidden sheets disappear, and finding the one you need means hunting along the bottom of the window. This add-in docks a compact pane on the left side of the Excel window that lists every worksheet, so switching, renaming and managing sheets no longer requires touching the tab bar.

The pane is built with WPF and hosted inside a standard Office task pane, so it follows Excel's left dock like any built-in pane.

## Features

- **Sheet list** — every worksheet of the active workbook, with the active one highlighted and hidden ones dimmed
- **Search** — filter the list as you type
- **Switch** — click a row to activate that worksheet
- **Visibility** — the eye icon on each row hides or unhides the sheet
- **Rename / delete / new** — from the pane header or the row context menu
- **Tab colors** — worksheet tab colors from Excel are shown as a colored stripe on each row
- **Settings page** — toggle whether hidden sheets are listed, and set the pane width (220–700 px)
- **Remembers state** — if the pane was open when you closed Excel, it opens again next time
- **Ribbon integration** — a *导航* button in the ribbon toggles the pane

## Requirements

| Component | Version |
|---|---|
| Excel | 2013 or later (32-bit and 64-bit) |
| .NET Framework | 4.7.2 or later |
| VSTO runtime | Microsoft Visual Studio 2010 Tools for Office Runtime |
| Build machine | Visual Studio 2022 with the *Office/SharePoint development* workload |

## Build

```text
git clone <repository>
```

1. Open `ExcelAddIn1.sln` in Visual Studio 2022.
2. Switch the solution configuration to **Release**.
3. Build → Rebuild Solution (`Ctrl+Alt+F7`).
4. Output lands in `ExcelAddIn1\bin\Release`:

```text
ExcelAddIn1.dll
ExcelAddIn1.dll.manifest
ExcelAddIn1.vsto
ExcelAddIn1.pdb
Microsoft.Office.Tools.Common.v4.0.Utilities.dll
```

### Signing

The key is **not in the repository**, and neither are the signing settings — the project builds with signing turned **off**, so a fresh clone compiles and debugs out of the box.

To produce an installable, signed package, drop a local `Signing.props` next to `ExcelAddIn1.csproj` (git-ignored):

```xml
<Project>
  <PropertyGroup>
    <SignManifests>true</SignManifests>
    <ManifestKeyFile>$(MSBuildProjectDirectory)\ExcelAddIn1_SigningKey.pfx</ManifestKeyFile>
    <ManifestCertificateThumbprint>YOUR_CERTIFICATE_THUMBPRINT</ManifestCertificateThumbprint>
  </PropertyGroup>
</Project>
```

Use **your own** certificate — either generate one with `New-SelfSignedCertificate -Type CodeSigningCert`, or pick an existing PFX through *Project Properties → Signing*, which fills in both values for you. Remember to hand out the matching `.cer` alongside the installer.

A build without `Signing.props` produces an unsigned manifest: fine for development, rejected by Office on machines that do not trust the publisher.

## Deploying to another machine

The add-in is published with ClickOnce and signed, so the target machine has to trust the publisher once:

```powershell
Import-Certificate -FilePath .\ExcelAddIn1.cer -CertStoreLocation Cert:\CurrentUser\Root
Import-Certificate -FilePath .\ExcelAddIn1.cer -CertStoreLocation Cert:\CurrentUser\TrustedPublisher
```

Then run `setup.exe` from the publish folder. Without the certificate, Excel refuses to load the add-in and reports that the publisher of the customization is not trusted.

`ExcelAddIn1.cer` is the **public** half of the signing certificate. It is not kept in the repository — it belongs to whoever signs the build, so export it from your key when you publish and ship it next to the installer (or attach it to a GitHub release).

## Configuration

User settings live in:

```text
%LOCALAPPDATA%\ExcelAddIn1\settings.ini
```

| Key | Meaning |
|---|---|
| `PaneWidth` | Pane width in pixels (clamped to 220–700) |
| `ShowHiddenSheets` | Whether hidden worksheets appear in the list |
| `LastPaneVisible` | Whether the pane was open when Excel last closed |

## Project layout

```text
ExcelAddIn1.sln
ExcelAddIn1/
  ThisAddIn.cs            Add-in entry point: Excel events, pane lifecycle, sheet operations
  PaneHost.cs             WinForms shell that hosts the WPF pane inside the task pane
  RibbonNavigation.cs     Ribbon button that toggles the pane
  UI/                     WPF user interface (contains no Excel types)
    NavPaneControl.cs     Root control: icon rail on the left, pages on the right
    SheetListPage.cs      Sheet list, search box, header actions
    SettingsPage.cs       Settings UI
    SheetInfo.cs          Plain view model for a worksheet
    PaneOptions.cs        Reads and writes settings.ini
    UiTheme.cs            Colors, fonts, icons and control styles
    UiHelpers.cs
    UiPrompt.cs           Rename / confirmation dialog
  Properties/
```

## Notes and limitations

- The task pane's **title bar** (the row showing *工作表导航* with the collapse and close buttons) is drawn by Office. VSTO exposes no API to change its font, size or styling.
- The pane shell is plain WinForms and is created with Excel, so it appears immediately; the WPF content is attached a moment later. This keeps the WPF startup cost out of Excel's launch path.
- Hiding the active worksheet is not allowed by Excel, so the add-in switches to another visible sheet first when needed.
