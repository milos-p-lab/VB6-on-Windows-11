# 🛠️ Bringing VB6 Back to Life on Windows 10/11 — Without Virtual Machines

Modern Windows systems (10, 11, and even 24H2) can still run and compile legacy VB6 projects — but only with the right combination of setup, registration, and compatibility tweaks. This project packages everything you need into a single self-extracting installer: **no VM required**, no registry headaches, no missing OCXs.

Whether you're fixing a 20-year-old tool, exploring classic codebases, or simply feel at home in the VB6 IDE, this project gives you a ready-to-use experience.

## 🧩 Why This Project Exists

Like many developers with a long history in Windows programming, I still occasionally need to maintain or run legacy Visual Basic 6.0 applications. Whether it's for preserving old utilities, analyzing historical code, or simply accessing data created in a previous era, the need for a working VB6 IDE still arises.

Unfortunately, setting up VB6 on modern systems — especially Windows 10 and 11 — has become increasingly problematic. While some opt for virtual machines or compatibility layers, I wanted a **simple, fast, and direct** solution that didn't require any VMs, no sketchy patches, and no endless DLL troubleshooting.

That's how this project began.

---

## ⚙️ Challenges with Running VB6 on Windows 10/11

Getting VB6 to run on Windows 11 (especially 24H2) wasn't straightforward. Here are some of the common issues I faced:

- The **original installer fails** or freezes on modern Windows.
- The IDE loads, but **fails to render** certain controls like `Toolbar`, `StatusBar`, or `ImageList`.
- The infamous `MSCOMCTL.OCX` is either missing or fails to register properly.
- **Double Agent** (used as a modern MS Agent replacement) must be manually registered.
- Running old `.exe` files compiled in VB6 often fails without **runtime dependencies**.
- Antivirus software occasionally flags **self-extracting archives** (SFX) as false positives.

---

## 🔄 The Process: Trial and Error

The process was a mix of research, scripting, testing on fresh machines, and documenting every detail. After manually assembling all required components and resolving DLL and OCX registrations, I bundled everything into a **WinRAR self-extracting executable**.

✅ Tested successfully on:

- **Windows 11 (23H2, 24H2)**
- **Windows 10 (22H2)**
- **Windows 7 SP1**
- **Windows XP SP3**

In all cases, VB6 IDE launched correctly, projects loaded and compiled, and ActiveX controls worked inside the designer — without needing a virtual machine.

---

## 🧪 Delay Launcher: Solving `Process.Start()` Issues

One interesting part of this project was the creation of `delaycmd.exe` — a small utility that **delays execution of another program** by a few seconds. While this may seem trivial, it's often essential for:

- Ensuring registry keys or COM components are fully registered before launch.
- Delaying startup in scripts, without relying on PowerShell or batch tricks.
- Working around race conditions during installations or first-run setups.

Unlike C#’s `Process.Start()` or Java’s `Runtime.getRuntime().exec()`, which may exit before child processes are fully initialized, this utility offers **precise control** in a native binary with zero dependencies.

---

## 📦 What's in the Package?

The archive includes:

- Visual Basic 6.0 IDE (fully functional)
- All essential ActiveX controls and design-time libraries
- Pre-registered components (`MSCOMCTL.OCX`, `MSFLXGRD.OCX`, `MSCHRT20.OCX`, etc.)
- IDE fixes for Windows 10/11 (e.g., `msstkprp.dll` for toolbar editing)
- Automated setup via `.bat` and `.reg` files
- No patching or hex-editing — just unpack and use

---

## 🧩 Optional: MS Agent Compatibility (Double Agent)

VB6 originally supported Microsoft Agent — a discontinued animation/speech system. If your old project used MS Agent, it will fail on modern Windows, as the original agents are deprecated and unsupported.

Instead, you can use **[Double Agent](https://doubleagent.sourceforge.net/)** — a drop-in replacement compatible with `AgentCtl.Agent` components.

To enable it:

1. Download and install Double Agent from the official site.
2. Your existing VB6 forms using MS Agent will work without modification.

This is not bundled in the main installer to reduce size and because only a small number of legacy projects depend on it.

---

## 🔍 Additional Examples

Some old utilities and small tools have been extracted and modernized as part of this effort. These are placed in a separate folder:

➡️ [vb6-projects](https://github.com/milos-p-lab/VB6-on-Windows-11/tree/main/vb6-projects)

This includes:

- delaycmd — a utility for delaying the launch of a program without leaving an open Command Prompt window. Useful in scripting and installer logic.
- Source code (.vbp, .frm, .bas) is included for learning or reuse.

None of these examples are included in the main setup (Visual-Basic-6-for-Win10-11.exe) to keep it clean and focused. You can browse and build them separately if needed.

## ⚠️ About Compatibility

Although tested on several systems, **I can't guarantee it will work on all hardware and configurations**. Use it at your own discretion — especially in production environments. This is not a Microsoft-supported solution, but a community-driven effort to **preserve legacy tools**.

---

## 🙌 Final Thoughts

Visual Basic 6.0 may be old, but its speed, simplicity, and RAD capabilities still make it relevant for certain use cases. With this setup, I hope to help others keep their tools alive, preserve old projects, or simply explore a fascinating part of programming history — without resorting to virtual machines or unsafe hacks.

Feel free to contribute, open issues, or share your experience with this setup.

---

> ✍️ **Author:** Miloš Perunović  
> 🗓️ **Date:** 2025-07-29
