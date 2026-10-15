# 📊 Folder Size Reporter to Excel

A robust, enterprise-ready PowerShell script designed to scan immediate subfolders, calculate their cumulative sizes recursively in Megabytes (MB), and export a clean, sorted report directly into **Microsoft Excel**.

---

## ✨ Key Features

* **🗂️ Recursive Deep Scan:** Drills down through all nested directories to calculate the true aggregate footprint of every primary subfolder.
* **🛡️ Windows Server Optimized:** Safely handles highly populated production **Windows File Servers** and **Windows Server** clusters without impacting underlying file operations.
* **🚫 Silent Error Resilience:** Gracefully skips system files, symlinks, or locked directories throwing permission or access errors instead of breaking execution.
* **📟 Real-Time Console Preview:** Prints a live top-5 ranking table of the largest directories directly to the console for quick visual analysis.
* **🎛️ Intelligent Workbook Control:** Automatically detects if your target `.xlsx` file exists. It will append/overwrite a designated tab or generate a fresh workbook dynamically.
* **🎨 Autopilot Formatting:** Automatically generates high-visibility light-gray headers, applies bold weights, auto-fits cell borders, and maps explicit two-decimal (`0.00`) numeric alignment.
* **🧹 Zero-Leak Process Management:** Explicitly breaks down Excel COM interfaces and forces standard system garbage collection routines to prevent rogue `EXCEL.EXE` background instances from hanging on your server infrastructure.

---

## 🖥️ Windows Server Security & Deployment Notes

This utility is structured specifically to run inside sensitive production infrastructure environments:
* **Read-Only Operations:** The script uses native `Get-ChildItem` streams to capture structural metadata without ever locking, touching, or altering data payload structures.
* **Excel Dependency Strategy:** Because the script uses the system COM application interface, Microsoft Excel must be installed on the machine running the code. 
* **Headless Architecture Tip:** If you are operating on a hardened server where Excel installations are restricted, simply execute the script from an administrative workstation and pass the target server's network UNC network pathway (e.g., `\\YourFileServer\SharedVolume`) directly into the script parameters.

---

## 🛠️ Prerequisites

* **Operating System:** Windows Server (2012 R2 up to 2025/2026+) or Windows 10/11.
* **Shell Engine:** Windows PowerShell 5.1+ or PowerShell 7+.
* **Automation Bridge:** Local installation of Microsoft Excel.

---

## 📌 Parameters & Arguments

| Parameter | Data Type | Required | Description |
| :--- | :---: | :---: | :--- |
| `-RootPath` | `[string]` | **Yes** | The local absolute path or network UNC address of the directory tree you want to audit. |
| `-OutputPath` | `[string]` | **Yes** | The complete path location where the destination Excel `.xlsx` workbook will be stored. |
| `-SheetName` | `[string]` | **Yes** | The exact string name applied to the targeted worksheet tab inside the file. |

---

## 🚀 Execution & Usage Guide

1. Save the source script to your system as `Get-FolderSizes.ps1`.
2. Open an administrative PowerShell console session.
3. Pass your runtime target paths into the parameter flags as shown below:

```powershell
.\Get-FolderSizes.ps1 `
    -RootPath "D:\Shares\CompanyData" `
    -OutputPath "C:\Audits\DiskSpaceStorageReport.xlsx" `
    -SheetName "Department Footprints"
```

---

## 📊 Structured Output Schema

The targeted Excel sheet dynamically purges old static content before injecting a freshly sorted grid matrix:

| Column Name | Data Presentation | Description |
| :--- | :--- | :--- |
| **Folder Name** | Text | The base system directory string folder name. |
| **Full Path** | Text Path String | The complete local or UNC structural path. |
| **Size (MB)** | Numeric Float (`0.00`) | The recursive calculation total, sorted cleanly from largest to smallest. |

---

## 📄 Licensing & Disclaimer
This script is provided "as-is" under standard open administrative distribution workflows. Always validate performance limits against sandbox structures before pointing scripts at arrays containing millions of deeply nested paths.
