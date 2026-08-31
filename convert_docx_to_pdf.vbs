Option Explicit

Const wdExportFormatPDF = 17
Const wdAlertsNone = 0

Dim args
Set args = WScript.Arguments

If args.Count < 2 Then
    WScript.Echo "Usage: cscript //nologo convert_docx_to_pdf.vbs input.docx output.pdf"
    WScript.Quit 1
End If

Dim inputPath, outputPath
inputPath = args(0)
outputPath = args(1)

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

If Not fso.FileExists(inputPath) Then
    WScript.Echo "Input file not found: " & inputPath
    WScript.Quit 2
End If

Dim outputFolder
outputFolder = fso.GetParentFolderName(outputPath)
If Len(outputFolder) > 0 Then
    If Not fso.FolderExists(outputFolder) Then
        CreateFolders fso, outputFolder
    End If
End If

Dim wordApp, doc
On Error Resume Next
Set wordApp = CreateObject("Word.Application")
If Err.Number <> 0 Then
    WScript.Echo "Cannot start Microsoft Word: " & Err.Description
    WScript.Quit 3
End If
Err.Clear

wordApp.Visible = False
wordApp.DisplayAlerts = wdAlertsNone

Set doc = wordApp.Documents.Open(inputPath, False, True, False)
If Err.Number <> 0 Then
    WScript.Echo "Cannot open document: " & Err.Description
    wordApp.Quit
    WScript.Quit 4
End If
Err.Clear

doc.ExportAsFixedFormat outputPath, wdExportFormatPDF
If Err.Number <> 0 Then
    WScript.Echo "Cannot export PDF: " & Err.Description
    doc.Close False
    wordApp.Quit
    WScript.Quit 5
End If

doc.Close False
wordApp.Quit

Set doc = Nothing
Set wordApp = Nothing

WScript.Echo outputPath
WScript.Quit 0

Sub CreateFolders(fso, folderPath)
    Dim parentPath
    parentPath = fso.GetParentFolderName(folderPath)
    If Len(parentPath) > 0 Then
        If Not fso.FolderExists(parentPath) Then
            CreateFolders fso, parentPath
        End If
    End If
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub
