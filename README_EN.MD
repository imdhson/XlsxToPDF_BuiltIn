# Excel Batch PDF Converter | xlsx to PDF Batch Converter

## 📥 **Download**

[![Download Excel Macro](https://img.shields.io/badge/📥%20Download-Excel%20Macro%20File-brightgreen?style=for-the-badge&logo=microsoft-excel)](xlsx_To_PDF[imdhson].xlsm)
[![Latest Release](https://img.shields.io/badge/📦%20Latest-Release-blue?style=for-the-badge&logo=github)](https://github.com/imdhson/XlsxToPDF_BuiltIn/releases)

## Introduction

This is a VBA macro tool that converts Excel XLSX files to PDF in batch. It addresses privacy concerns and speed issues associated with online conversion tools by utilizing Excel’s official PDF conversion feature, enabling safe and fast batch processing in your local environment.

## Development Background

- **Resolves privacy concerns** with online PDF conversion tools
- **Fast batch conversion** in a local environment
- **Minimizes dependency issues** by using Excel’s official conversion feature
- **Enhances efficiency** for multiple file processing

## Installation & Setup

### 1. Requirements

- **Microsoft Excel** (version supporting macros)
- **Windows operating system**

### 2. Download the File

Click the **Download** button above to download the [`xlsx_To_PDF[imdhson].xlsm`](xlsx_To_PDF[imdhson].xlsm) file.

### 3. Enable Macros

Allow macros to run when opening the file in Excel.

### 4. Ready to Use

Open the Excel file and click the **[Start]** cell to begin the conversion process.

## Main Features

- **Batch Conversion:** Convert all XLSX files in a selected folder to PDF at once
- **Local Processing:** Process files safely without an internet connection
- **Official Excel Conversion:** High-quality conversion using the `ExportAsFixedFormat` function
- **Simple Interface:** Intuitive operation with just a cell click
- **Automatic Folder Creation:** Automatically creates a `./pdf` folder to organize converted files
- **Timestamp Filename:** Saves files as `originalfilename_lastmodifiedtime.pdf`

## Usage Instructions

### 1. Starting the Conversion

1. **Open** the `xlsx to PDF[imdhson].xlsm` file in Excel.
2. **Click** the **[Start]** cell.
3. **Select** the folder containing the XLSX files to convert.

### 2. Conversion Process

1. **Select Folder:** Choose the folder with your XLSX files.
2. **File Search:** Automatically finds all `.xlsx` files in the selected folder.
3. **PDF Folder Creation:** Automatically creates a `./pdf` subfolder in the selected folder.
4. **Batch Conversion:** Opens each file sequentially and converts it to PDF.
5. **File Saving:** Saves each PDF as `originalfilename_YYYYMMDD_HHMMSS.pdf` in the `./pdf` folder.
6. **Completion Notice:** Displays a notification when all files are converted.

### 3. Filename Rules

Converted PDF files are saved with the following rules:
- **Save Location:** `selected_folder/pdf/`
- **Filename Format:** `originalfilename_YYYYMMDD_HHMMSS.pdf`
- **Example:** `report.xlsx` → `report_20250712_220730.pdf`

### 4. Notes

- **Subfolders Not Supported:** Only files in the selected folder are converted.
- **PDF Folder Auto-creation:** The `pdf` folder is created if it does not exist.
- **Screen Updates Disabled:** Excel screen updates are disabled during conversion for better performance.

## Core VBA Code

### Main Conversion Function

```vba
Sub Batch_Export_XLSX_to_PDF()
    Dim folderPath As String
    Dim pdfFolderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim fileModifiedTime As Date
    Dim timeStamp As String
    Dim baseName As String
    Dim pdfFileName As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select a directory with .xlsx files"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            Exit Sub
        End If
    End With
    
    ' Set and create PDF folder path
    pdfFolderPath = folderPath & "pdf\"
    If Dir(pdfFolderPath, vbDirectory) = "" Then
        MkDir pdfFolderPath
    End If
    
    fileName = Dir(folderPath & "*.xlsx")
    Application.ScreenUpdating = False
    
    Do While fileName <> ""
        ' Get file modified time
        fileModifiedTime = FileDateTime(folderPath & fileName)
        timeStamp = Format(fileModifiedTime, "YYYYMMDD_HHMMSS")
        
        ' Base filename (without extension)
        baseName = Replace(fileName, ".xlsx", "")
        
        ' Generate PDF filename
        pdfFileName = baseName & "_" & timeStamp & ".pdf"
        
        ' Open file and export as PDF
        Set wb = Workbooks.Open(folderPath & fileName)
        wb.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfFolderPath & pdfFileName
        wb.Close False
        
        fileName = Dir
    Loop
    
    Application.ScreenUpdating = True
    MsgBox "Export complete. Path: ./pdf - dev:imdhson.com"
End Sub
```

### Cell Click Event

```vba
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Address = "$A$1" Then
        Call Batch_Export_XLSX_to_PDF
    ElseIf Target.Address = "$A$2" Then
        Call Batch_Export_XLSX_to_PDF
    End If
End Sub
```

## Technical Features

- **ExportAsFixedFormat:** Uses Excel’s official PDF conversion API
- **Batch Processing:** Efficient file traversal using the Dir function
- **User-Friendly:** Intuitive folder selection via FileDialog
- **Performance Optimization:** Disables ScreenUpdating for faster conversion
- **Automatic Folder Management:** Creates PDF folder automatically using MkDir
- **Timestamp Handling:** Accurate time information using FileDateTime and Format functions

## License

This project is distributed under the [MIT License](LICENSE.md).  
You are free to use it for both personal and commercial purposes.

## Contact

**Developer:** Donghwi Son

- **Email:** mail@imdhson.com
- **Instagram:** @imdhson  
- **GitHub:** [https://github.com/imdhson](https://github.com/imdhson)
- **Website:** http://imdhson.com

---