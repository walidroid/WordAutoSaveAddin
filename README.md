# Word Auto-Save Add-in
### Microsoft Word 2019 · VSTO · C# · .NET Framework 4.7.2 · Fully Offline

---

## What it does

Automatically saves your active Word document every **10 seconds**.

| Feature | Detail |
|---------|--------|
| Save interval | 10 seconds (configurable in `AutoSaveManager.cs`) |
| Trigger | Saves only when `doc.Saved == false` (no unnecessary disk writes) |
| New documents | Skipped until the document has been saved at least once (has a path) |
| UI | Custom **Auto-Save** ribbon tab with toggle button + live status label |
| Thread safety | `System.Windows.Forms.Timer` fires on the UI thread – no marshalling |
| Shutdown | Timer is stopped and disposed cleanly when Word exits |

---

## Prerequisites

### On the development machine

| Requirement | Version |
|-------------|---------|
| Windows | 10 or 11 (64-bit) |
| Visual Studio | 2019 (any edition – Community is free) |
| Visual Studio workload | **Office/SharePoint Development** |
| .NET Framework | 4.7.2 (included with VS workload) |
| Microsoft Word | 2019 (desktop) installed locally |

> **No internet required.** All packages are included with the VS workload.

---

## Project Structure

```
WordAutoSaveAddin/
├── WordAutoSaveAddin.sln               ← Open this in Visual Studio
│
├── Install-AddIn.ps1                   ← PowerShell installer
├── Install-AddIn.reg                   ← Manual registry installer
├── Uninstall-AddIn.ps1                 ← PowerShell uninstaller
│
└── WordAutoSaveAddin/
    ├── WordAutoSaveAddin.csproj        ← Project file
    ├── app.config                      ← .NET runtime config
    │
    ├── ThisAddIn.cs                    ← VSTO entry point (Startup/Shutdown)
    ├── ThisAddIn.Designer.cs           ← VSTO designer boilerplate
    │
    ├── AutoSaveManager.cs              ← Timer + save logic
    │
    ├── AutoSaveRibbon.cs               ← Ribbon callbacks
    ├── AutoSaveRibbon.Designer.cs      ← Ribbon designer placeholder
    ├── AutoSaveRibbon.xml              ← RibbonX markup (custom tab)
    │
    └── Properties/
        └── AssemblyInfo.cs
```

---

## Step 1 — Open in Visual Studio

1. Launch **Visual Studio 2019**.
2. **File → Open → Project/Solution** → select `WordAutoSaveAddin.sln`.
3. Visual Studio may prompt you to install missing components.  
   Accept and install **"Microsoft Office Developer Tools"** if asked.

---

## Step 2 — Verify References

In **Solution Explorer → WordAutoSaveAddin → References** confirm these are present
(they come from the VSTO workload – no NuGet packages needed):

```
Microsoft.Office.Interop.Word
Microsoft.Office.Tools.Common
Microsoft.Office.Tools.Common.v4.0.Utilities
Microsoft.Office.Tools.Word
Microsoft.Office.Tools.Word.v4.0.Utilities
System
System.Windows.Forms
```

If any are missing, right-click **References → Add Reference → COM** and browse to the
Office PIA (Primary Interop Assembly) folder:

```
C:\Program Files (x86)\Microsoft Visual Studio\2019\<Edition>\
  Visual Studio Tools for Office\PIA\Office15\
```

---

## Step 3 — Build

1. Set configuration to **Release** (dropdown in the toolbar).
2. **Build → Build Solution** (`Ctrl+Shift+B`).
3. Output goes to: `WordAutoSaveAddin\bin\Release\`

You should see:
```
WordAutoSaveAddin.dll
WordAutoSaveAddin.vsto
WordAutoSaveAddin.dll.manifest
```

> **Debug build** also works and produces the same files in `bin\Debug\`.

---

## Step 4 — Install (Offline, No Web Server)

The `vstolocal` suffix in the manifest path tells VSTO to load the add-in
directly from the local file system – **no ClickOnce or web server needed**.

### Option A — PowerShell (recommended)

Open **PowerShell** (no elevation required for HKCU) and run:

```powershell
cd "C:\path\to\WordAutoSaveAddin"
.\Install-AddIn.ps1 -BuildOutputPath "C:\path\to\WordAutoSaveAddin\WordAutoSaveAddin\bin\Release"
```

### Option B — Registry file

1. Open `Install-AddIn.reg` in Notepad.
2. Replace `C:\\YourPath\\` with your actual build output path  
   (use double backslashes: `C:\\Users\\You\\WordAutoSaveAddin\\bin\\Release\\`).
3. Save the file, then double-click it to merge into the registry.

### What the registry entry looks like

```
HKCU\SOFTWARE\Microsoft\Office\Word\Addins\WordAutoSaveAddin
  FriendlyName  = "Word Auto-Save Add-in"
  Description   = "Automatically saves the active Word document every 10 seconds."
  LoadBehavior  = 3        (load at startup, keep loaded)
  Manifest      = "C:\...\WordAutoSaveAddin.vsto|vstolocal"
```

---

## Step 5 — Trust the Add-in

The first time Word loads the add-in you may see a security prompt:

> *"Microsoft Office has identified a potential security concern."*

Click **Enable this add-in for this session** (or add the folder to Trusted
Locations via **File → Options → Trust Center → Trusted Locations**).

To permanently trust the folder:

1. **Word → File → Options → Trust Center → Trust Center Settings**.
2. **Trusted Locations → Add new location**.
3. Browse to `WordAutoSaveAddin\bin\Release\` and tick **Subfolders of this
   location are also trusted**.

---

## Step 6 — Verify it Works

1. Open or create a Word document **and save it once** (so it has a file path).
2. Look for the **Auto-Save** tab in the ribbon (between Home and Insert).
3. The tab shows:
   - A large **"Auto-Save ON"** toggle button (green/pressed by default).
   - A status label: `Auto-save: ON  |  Last saved 14:32:05`
4. Make a change – within 10 seconds the document is saved automatically.
5. Click the toggle to turn it **OFF**; the label changes to `Auto-save: OFF`.

---

## Changing the Save Interval

Open `AutoSaveManager.cs` and change the constant:

```csharp
private const int SaveIntervalMs = 10_000;   // 10 seconds  ← change this
```

Rebuild and reinstall.

---

## Uninstall

```powershell
.\Uninstall-AddIn.ps1
```

Or delete the registry key manually:

```
HKCU\SOFTWARE\Microsoft\Office\Word\Addins\WordAutoSaveAddin
```

---

## Troubleshooting

### Add-in tab does not appear

- Check **Word → File → Options → Add-ins**.
- If the add-in is listed under **Disabled Application Add-ins**, select  
  **COM Add-ins** in the Manage dropdown and click **Go…** to re-enable it.
- Verify the `Manifest` registry value points to the correct `.vsto` file path.

### "Could not load file or assembly" error in Word

- Rebuild the project in Visual Studio.
- Ensure `LoadBehavior` is `3` (DWORD) in the registry.
- Confirm .NET Framework 4.7.2 is installed (`winver` → Settings → Apps → .NET).

### Add-in loads but does not save

- The document must have been **saved at least once** (File → Save As) before
  auto-save kicks in. Unsaved new documents have no path and are skipped.
- Open the Visual Studio **Output** window while debugging to see
  `[AutoSaveManager]` log lines.

### Word crashes on startup after installing

- Run `Uninstall-AddIn.ps1` to remove the registry entry, restart Word, then
  re-examine the build output for errors before reinstalling.

---

## Architecture Summary

```
ThisAddIn (VSTO host)
  │  Startup  → creates AutoSaveManager, hooks Word events
  │  Shutdown → disposes AutoSaveManager
  │
  ├─ AutoSaveManager
  │    System.Windows.Forms.Timer (10 s, UI thread)
  │    OnTimerTick → checks doc open? has path? has changes? → doc.Save()
  │    Start() / Stop() / Toggle() / IsRunning / StatusText / LastSavedAt
  │
  └─ AutoSaveRibbon  (IRibbonExtensibility)
       AutoSaveRibbon.xml  → custom tab, toggleButton, labelControl
       Callbacks: OnToggleButton_Click, GetStatusLabel, GetToggleButtonPressed
       A 2-second UI timer invalidates the ribbon to refresh the status label
```
