# 📖 Technical Notes — VB6 on Windows 11

This document provides a detailed explanation of the internal mechanisms and compatibility adjustments used by the **VB6 on Windows 10/11** package.  
These notes are intended for developers, system administrators, and power users who want to understand *how* and *why* the installation works the way it does.

---

## 1. Installer Overview

The installation is performed via a self-extracting WinRAR archive (`Visual-Basic-6-for-Win11.exe`) that includes:

- All necessary VB6 runtime files (DLLs, OCXs, etc.) based on the official Visual Basic 6.0 Service Pack 6 (SP6) distribution.
- Batch scripts (`.bat` files) that automate the setup process, including file copying, registration, and configuration.
- Registry files (`.reg`) that apply essential settings to ensure compatibility with modern Windows versions.
- A desktop shortcut that launches VB6 in compatibility mode with elevated privileges via a scheduled task.

## 2. Scheduled Task for VB6 Launch

The installation script creates a **Windows Scheduled Task** that launches `VB6.exe` in:

- **Windows XP compatibility mode**, and  
- **with elevated (administrator) privileges**,  
- **without triggering a UAC prompt**.

This ensures that the VB6 IDE starts seamlessly by clicking its desktop shortcut, just like it did on Windows XP.  

```bat
rem Create or update the scheduled task
schtasks /Delete /TN "M%TASKNAME%-%USERNAME%" /F
schtasks /Create /TN "M%TASKNAME%-%USERNAME%" /xml "%TMPXML%"
```

📝 On Windows XP, VB6 runs directly without any scheduled task. On Windows 7, compatibility is handled differently, but Windows 10/11 require this method for a smooth launch experience.

---

## 3. MSCOMCTL.OCX Version Note

This package includes version 6.1.98.34 of MSCOMCTL.OCX, which is the official version shipped with Microsoft Visual Basic 6.0 Service Pack 6 (SP6).

Some systems may have newer versions (e.g. 6.1.98.46, 6.1.98.50) installed via:

- Microsoft Office (2007–2016),
- Windows Update security patches,
- Other third-party software.

These newer versions are not part of the official SP6 distribution and occasionally introduce compatibility issues with legacy VB6 projects.

---

## 4. DLL/OCX Registration and Backup

All required VB6 runtime components (.DLL and .OCX files) are copied into the appropriate system directories and registered silently using regsvr32.
Before overwriting, the script creates a backup of any existing files with the .old extension (e.g., MSCOMCTL.OCX.old) to preserve the previous state - for example:

``` plaintext
mscomctl.ocx.old
comdlg32.ocx.old
tabctl32.ocx.old
etc.
```

This approach:

- Ensures predictable behavior across different machines,
- Preserves user-installed versions in case they need to restore them later,
- Helps avoid subtle runtime mismatches that often happen with ActiveX controls.

If an existing OCX/DLL file is found during setup, it will be backed up before being replaced — for example:

This allows you to easily restore your previous version — which may be newer — if you wish to keep it.
👉 If you rely on a newer version of MSCOMCTL.OCX or any other ActiveX control, simply re-register your preferred file manually, e.g.:

```bat
regsvr32 /s "C:\Path\To\Your\Preferred\MSCOMCTL.OCX"
```

or restore the backed-up .old file after installation.
(Adjust paths as needed — the `.old` files will be in the same folder as the replaced files unless moved manually.)
This project aims for maximum stability and compatibility with legacy VB6 projects, so original SP6 versions are shipped by default — but your backups give you full control.

---

## 5. Compatibility Adjustments

Several key compatibility fixes are included to make VB6 fully functional on modern systems:

**Data Environment & Data Report**
Registry settings are applied to ensure proper loading and editing of Data Report layouts in the VB6 IDE, not just runtime execution.

**Double Agent (MS Agent replacement)**
Projects that use Microsoft Agent controls can work with Double Agent, which provides modern 32/64-bit replacements compatible with VB6.

Register using:

```bat
regsvr32 /s "C:\Program Files (x86)\Double Agent\DaControl.dll"
```

**Toolbar, StatusBar, ImageList Controls**
For editing these controls in the IDE (not just runtime), the following must be registered:

```bat
regsvr32 /s "%SystemRoot%\system32\msstkprp.dll"
```

---

## 6. 32-bit vs 64-bit Considerations

VB6 is a 32-bit development environment and all bundled runtime components are 32-bit. On 64-bit Windows:

- Files are intentionally copied to System32, not SysWOW64, because for 32-bit COM components this is the expected registration location under Windows-on-Windows redirection.
- The installer and scripts run in a 32-bit administrative context to ensure consistent behavior.

---

## 7. Optional Customizations

Advanced users may wish to modify some default behaviors:

- Disabling the Scheduled Task: You can bypass the task creation step if you prefer to launch VB6 manually with “Run as Administrator” and compatibility mode.
- Using your own runtime versions: Replace or re-register any OCX/DLL files after installation to fit your specific project needs.
- Silent / automated installs: The provided `.bat` scripts can be adapted for mass deployment or integration with custom provisioning systems.

---

## 8. Why Not Just Use twinBASIC?

While this project makes it possible to run and maintain classic VB6 applications on modern Windows, the **future** of VB6-style development is [twinBASIC](https://twinbasic.com/).

- ✅ 32-bit **and** 64-bit executables  
- ✅ Modern component support (including WebView2)  
- ✅ Actively developed and community-driven  
- ❌ Still missing some legacy-specific features (e.g., DataEnvironment, DataReport, MS Agent/Double Agent), which are important in certain business applications  

This project serves as a **compatibility bridge** — ensuring that old VB6 applications can still be compiled and maintained today, while twinBASIC continues to evolve into a full replacement.  

👉 **Note:** If your project depends on these legacy features, the original VB6 IDE remains the only option for editing and compiling. For all new development, however, twinBASIC is strongly recommended as the long-term solution.

---

## 9. Troubleshooting

Always test on a clean VM when in doubt (especially Windows 11 25H2), to isolate environmental issues.
If custom controls are involved (e.g. Crystal Reports, third-party OCX), ensure all required dependencies are installed and registered.
Use OLE/COM Object Viewer (oleview.exe) or Sysinternals ProcMon to debug missing registrations.

---

⚖️ © 2025 Miloš Perunović — For educational and legacy maintenance purposes.
All rights reserved.
