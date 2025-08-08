# Visual Basic 6.0 for Windows 10/11 — No VM Required

[![License: MIT](https://img.shields.io/badge/License-as_is-yellow.svg)](LICENSE) [![Visual Basic 6](https://img.shields.io/badge/Visual%20Basic-6-darkcyan.svg)](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/visual-basic-6.0-documentation)

> Enables editing and running legacy VB6 applications directly on modern Windows.

## 💾 About

This package provides a ready-to-use setup for running and compiling VB6 projects on Windows 10 and 11, without needing virtual machines or manual OCX registration. It's ideal for maintaining old software, analyzing legacy code, or building quick utilities in a familiar environment.

👉 For background information, setup explanations, and optional components, see the [blog post](https://milos-p-lab.github.io/VB6-on-Windows-11/blog).

> ✍️ **Author:** Miloš Perunović  
> 🗓️ **Date:** 2025-08-08

---

## 🚀 Features

- Fully working VB6 IDE on Windows 10 and 11 (including 24H2)
- No need to manually register ActiveX controls or system DLLs
- Registry and batch files automate all configuration steps
- Supports compiling, editing, and running old projects out-of-the-box
- Ideal for educational purposes, small utilities, and legacy support

---

## 📂 How It Works

The entire environment is packaged as a self-extracting archive containing:

- The original `VB6.exe` executable and IDE files
- All necessary runtime libraries and dependencies
- Preconfigured registry settings to ensure compatibility with modern Windows
- Batch scripts to automate the registration of ActiveX controls and system libraries
- Required OCX/DLL files (e.g. `mscomctl.ocx`, `msstkprp.dll`)
- Registration scripts (`.reg`, `.bat`) that ensure VB6 components work properly
- A desktop and Start menu shortcut for easy access to the VB6 IDE

All actions are performed locally on your machine — no system-level installations or persistent changes beyond what's required for VB6 to function.

---

## 🔐 Antivirus Warning

Some antivirus programs — including Windows Defender — may flag the self-extracting installation archive as suspicious or even as malware. This is a **false positive**, triggered by the nature of the WinRAR SFX archive, which includes scripts (.bat, .reg) that modify system settings to enable VB6 functionality.

> You can safely run the installer if you downloaded it from this repository. If in doubt, inspect or extract the contents manually before execution.

We recommend temporarily disabling real-time protection only during installation, if necessary, and restoring it immediately afterward.

---

## ⚠️ Limitations

- The solution has been successfully tested on Windows 7, Windows 10, and Windows 11 (24H2), but due to variations in system configurations and updates, full compatibility is not guaranteed.
- Microsoft officially ended support for VB6 development tools; this package is provided for educational, archival, and legacy support purposes.
- This setup is **not recommended** for creating new business applications — for long-term development, consider porting to modern technologies such as .NET.

---

### Note on MSCOMCTL.OCX Version

This setup includes version 6.1.98.34 of MSCOMCTL.OCX, which is the official version shipped with Microsoft Visual Basic 6.0 Service Pack 6 (SP6).

You may have a newer version installed (e.g., 6.1.98.46, 6.1.98.50), which could have been registered by:

- Microsoft Office (2007–2016)
- Windows Update security patches
- Other applications bundling ActiveX controls

These newer versions are not part of the official VB6 SP6 distribution and sometimes introduce subtle compatibility issues.

👉 Recommendation:
If you rely on a newer version of MSCOMCTL.OCX, feel free to re-register it after installing this package using:

```batch
regsvr32 /s "C:\Path\To\Your\Newer\MSCOMCTL.OCX"
```

This project aims for maximum stability and compatibility with legacy VB6 projects, and therefore ships the original SP6 version by default.

---

## 🛠️ Installation

1. Download the `Visual-Basic-6-for-Win10-11.exe` installer from the [Releases](https://github.com/milos-p-lab/VB6-on-Windows-11/releases/) page.
2. Run the installer (self-extracting archive) with administrative privileges.
3. Follow the prompts; all required files will be extracted and configured automatically.
4. After installation, launch VB6 from the desktop shortcut, which is created automatically by the installer for your convenience.

✅ Tested successfully on:

- **Windows 11 (23H2, 24H2)**
- **Windows 10 (22H2)**
- **Windows 7 SP1**
- **Windows Server 2003 R2**
- **Windows XP SP3**

---

## 💡 Why not just recompile in C#, Java, etc.?

You could — but certain legacy behaviors (e.g., precise startup timing, COM interop, message handling, or agent-like UI features) can’t be easily replicated in .NET or Java.

For example, utilities like `delaycmd.exe` allow precise control over when and how a dependent application is launched. This can be critical when dealing with shell extensions, startup scripts, or software requiring legacy timing patterns that modern `Process.Start()` (C#) or `Runtime.exec()` (Java) don't handle well.

---

## 🧳 Use Cases

- Recovering and analyzing old VB6 code
- Maintaining legacy desktop utilities still in production use
- Developing small tools or prototypes using familiar VB6 workflow
- Educational exploration of legacy programming paradigms
- Avoiding the need for full virtual machine environments

---

## 📃 License

This repository provides a preconfigured environment based on existing redistributable components. The author does not distribute proprietary software — the package assumes you have the legal right to use VB6 and related tools.

Use at your own risk.
