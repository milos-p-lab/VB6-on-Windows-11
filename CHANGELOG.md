# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.0.3] – 2025-08-08

### Changed

- The setup now makes a backup of any existing OCX/DLL file before overwriting it (e.g. mscomctl.ocx.old). This allows you to restore your original version if needed.
  - ☝ This project aims for maximum stability and compatibility with legacy VB6 projects, so original SP6 versions are shipped by default — but your backups give you full control.

## [1.0.2] – 2025-07-29

### Changed

- Start VB6 IDE without UAC prompt.

## [1.0.1] – 2025-07-28

### Fixed

- Import registry settings on Windows 7.

## [1.0.0] – 2025-07-28

### Added

- Initial release.
- Self-extracting archive containing VB6 IDE and runtime files.
- Automated registration of ActiveX controls and system libraries.
- Preconfigured OCX/DLL files required for VB6 applications.
- Desktop shortcut for easy access to the VB6 IDE.
- Support for editing and maintaining legacy VB6 applications on Windows 10 and 11, including 24H2.
- Ideal for educational purposes, small utilities, and legacy software maintenance.
