
Folder Size Reporter to Excel

A PowerShell script that scans the immediate subfolders of a specified directory, calculates their sizes recursively in megabytes (MB), sorts them from largest to smallest, and exports the data into a beautifully formatted Microsoft Excel spreadsheet.

Features

• Recursive Size Calculation: Loops through all nested files inside subfolders to calculate total space used.
• Enterprise Environment Safe: Built specifically to safely run on production Windows File Servers and Windows Server environments without causing high memory leaks or service disruptions.
• Robust Error Handling: Gracefully skips system files or directories that throw access/permission errors instead of stopping.
• Live Console Preview: Prints a top-5 ranking of the largest folders directly to the terminal for immediate verification.
• Smart Workbook Injection: Automatically creates a new Excel file or opens an existing workbook to append/overwrite a designated tab.
• Professional Layout: Formats the generated grid automatically with light-gray headers, bold text, auto-fitted columns, and precise two-decimal tracking (0.00) for MBs.
• Clean Process Management: Force-releases COM objects and triggers garbage collection to prevent hidden EXCEL.EXE background processes from lagging your server.

Windows Server Deployment Notes

This script is safe to execute on production Windows Server instances because:
• It uses standard, non-intrusive Get-ChildItem operations that read directory metadata without modifying files.
• It uses -ErrorAction SilentlyContinue to prevent permission-restricted admin shares or system directories from crashing the workflow.
• Note on Excel: Since the script leverages the Microsoft Excel COM interface, Excel must be installed on the machine execution path. For headless servers where installing Excel is restricted, it is recommended to run this script from a management workstation targeting the server's network share path (e.g., \\ServerName\ShareName).

Prerequisites

• Windows Server / Windows OS with PowerShell 5.1+
• Microsoft Excel installed on the system where the script runs (required for COM automation).

Script Parameters


Parameter	Type	Required	Description
RootPath	String	Yes	The absolute path (or UNC share path) of the target directory to scan.
OutputPath	String	Yes	The folder path and file name where the .xlsx file will be created or updated.
SheetName	String	Yes	The exact name of the worksheet tab where the data will be written.


Usage Example

Save the script as Export-FolderSizes.ps1 and run it from a PowerShell terminal with named arguments:
powershell
.\Export-FolderSizes.ps1 `
    -RootPath "D:\SharedData" `
    -OutputPath "C:\Reports\ServerDiskSpaceSummary.xlsx" `
    -SheetName "Share Sizes"
Use code with caution.

Generated Output Format

The target Excel sheet will be cleared of existing content and populated with three cleanly auto-fit columns:
1. Folder Name (The base name of the subfolder)
2. Full Path (The complete system or network address)
3. Size (MB) (Numeric column sorted in descending order)