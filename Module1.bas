Attribute VB_Name = "Module1"
Option Explicit

' ==============================================================================
' MACRO 1: Extract, Clean, and Format Serial Numbers from PDF via Word OCR
' ==============================================================================
Sub ImportPDFDataViaWord()
    Dim FileToOpen As Variant
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim wdTable As Object
    Dim wdRow As Object
    Dim wdCell As Object
    Dim ws As Worksheet
    Dim destRow As Long
    Dim cellText As String
    Dim WshShell As Object
    Dim regPath As String
    Dim colIdx As Long
    Dim rowIdx As Long
    Dim maxCols As Long
    Dim targetCol As Long
    Dim maxTargetCol As Long
    Dim rowHasData As Boolean
    Dim tableHadData As Boolean
    Dim i As Long
    Dim r As Long
    Dim c As Long
    Dim dataRowCounter As Long
    Dim startCol As Long
    
    ' Set the starting column to K (Column 11)
    startCol = 11
    maxTargetCol = startCol
    
    ' 1. Prompt user to select the scanned PDF
    FileToOpen = Application.GetOpenFilename("PDF Files (*.pdf), *.pdf", , "Select Scanned Serial Number PDF")
    If FileToOpen = False Then Exit Sub
    
    ' 2. Prepare the Active Excel Sheet (Protecting Columns A through J)
    Set ws = ActiveSheet
    
    If Application.WorksheetFunction.CountA(ws.Range("K:XFD")) > 0 Then
        If MsgBox("This will clear existing data in the extraction area (Columns K+). Do you want to proceed?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If
    
    ' Only clear the data area, leaving the buttons and instructions intact
    ws.Range("K:XFD").Clear
    ws.Range("K:XFD").Interior.Pattern = xlNone
    destRow = 1
    
    ' 3. Launch Invisible Word Instance
    Application.StatusBar = "Opening Word in background... Please wait, PDF conversion takes a moment."
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error GoTo ErrorHandler
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    wdApp.DisplayAlerts = 0
    
    ' --- THE REGISTRY FIX ---
    Set WshShell = CreateObject("WScript.Shell")
    regPath = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & wdApp.Version & "\Word\Options\DisableConvertPdfWarning"
    On Error Resume Next
    WshShell.RegWrite regPath, 1, "REG_DWORD"
    On Error GoTo ErrorHandler
    ' ------------------------
    
    ' 4. Open PDF
    Set wdDoc = wdApp.Documents.Open(Filename:=CStr(FileToOpen), ConfirmConversions:=False, ReadOnly:=True)
    
    ' 5. Extract and Clean Data (Spaced Grid starting at Column K)
    Application.StatusBar = "Extracting data from Word tables..."
    
    If Not wdDoc Is Nothing Then
        If wdDoc.Tables.Count > 0 Then
            For Each wdTable In wdDoc.Tables
                
                On Error Resume Next
                maxCols = wdTable.Columns.Count
                On Error GoTo ErrorHandler
                
                If maxCols > 0 Then
                    tableHadData = False
                    
                    ' Loop through rows top-to-bottom
                    For rowIdx = 1 To wdTable.Rows.Count
                        rowHasData = False
                        
                        ' Loop through columns left-to-right
                        For colIdx = 1 To maxCols
                            Set wdCell = Nothing
                            On Error Resume Next
                            Set wdCell = wdTable.cell(rowIdx, colIdx)
                            On Error GoTo ErrorHandler
                            
                            If Not wdCell Is Nothing Then
                                cellText = wdCell.Range.Text
                                cellText = Replace(cellText, Chr(13), "")
                                cellText = Replace(cellText, Chr(7), "")
                                cellText = Trim(cellText)
                                
                                cellText = Replace(cellText, " ", "")
                                cellText = Replace(cellText, ".", "-")
                                cellText = Replace(cellText, ",", "-")
                                cellText = UCase(cellText)
                                
                                ' THE STRICT GATEKEEPER
                                If Left(cellText, 2) = "JQ" Then
                                    rowHasData = True
                                    tableHadData = True
                                    
                                    ' Math to skip every other column starting at K (1->K, 2->M, 3->O, 4->Q)
                                    targetCol = startCol + (colIdx * 2) - 2
                                    If targetCol > maxTargetCol Then maxTargetCol = targetCol
                                    
                                    ' AUTO-FIX 1: Missing Dash
                                    If Len(cellText) = 11 And InStr(cellText, "-") = 0 Then
                                        cellText = Left(cellText, 8) & "-" & Right(cellText, 3)
                                    End If
                                    
                                    ' AUTO-FIX 2: Extra trailing characters
                                    If Len(cellText) > 12 Then
                                        If Left(cellText, 12) Like "[A-Z][A-Z]#[A-Z]####-###" Then
                                            cellText = Left(cellText, 12)
                                        End If
                                    End If
                                    
                                    ' Write to Excel
                                    ws.Cells(destRow, targetCol).Value = cellText
                                    
                                    ' VALIDATE & FLAG
                                    If Not cellText Like "[A-Z][A-Z]#[A-Z]####-###" Then
                                        ws.Cells(destRow, targetCol).Interior.Color = vbYellow
                                    End If
                                End If
                            End If
                        Next colIdx
                        
                        ' Move down one row
                        If rowHasData Then destRow = destRow + 1
                    Next rowIdx
                    
                    ' Leave a single blank row between physical pages to show the page break
                    If tableHadData Then destRow = destRow + 1
                End If
            Next wdTable
        Else
            MsgBox "Word successfully converted the PDF, but could not detect any grid tables.", vbExclamation
            GoTo Cleanup
        End If
    End If
    
    ' 6. Clean up the formatting for the target area only
    If maxTargetCol >= startCol Then
        ws.Range(ws.Columns(startCol), ws.Columns(maxTargetCol)).Columns.AutoFit
        ws.Range(ws.Columns(startCol), ws.Columns(maxTargetCol)).HorizontalAlignment = xlCenter
        
        ' Shrink spacer columns (L, N, P...)
        For i = startCol + 1 To maxTargetCol Step 2
            ws.Columns(i).ColumnWidth = 2.14
        Next i
        
        ' 7. Apply Zebra Striping (Data columns only)
        dataRowCounter = 1
        For r = 1 To destRow
            If Left(ws.Cells(r, startCol).Value, 3) = "---" Then
                dataRowCounter = 1 ' Reset alternating colors at the start of each new scanned page
            ElseIf Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, startCol), ws.Cells(r, maxTargetCol))) > 0 Then
                If dataRowCounter Mod 2 = 0 Then
                    ' Step 2 ensures we only stripe the data columns and skip the spacers
                    For c = startCol To maxTargetCol Step 2
                        ' Only apply grey background if the cell isn't already flagged yellow for an error
                        If ws.Cells(r, c).Interior.ColorIndex = xlNone Then
                            ws.Cells(r, c).Interior.Color = RGB(235, 235, 235) ' Light Grey
                        End If
                    Next c
                End If
                dataRowCounter = dataRowCounter + 1
            End If
        Next r
    End If

Cleanup:
    ' 8. Close Word and release memory safely
    Application.StatusBar = "Cleaning up..."
    If Not wdDoc Is Nothing Then
        wdDoc.Close False
    End If
    If Not wdApp Is Nothing Then
        wdApp.Quit
    End If
    
    Set WshShell = Nothing
    Set wdCell = Nothing
    Set wdRow = Nothing
    Set wdTable = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Data extraction complete! Please review the data for errors. If all serial numbers are correct, click on ""Download Reports"".", vbInformation
    Exit Sub

ErrorHandler:
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

' ==============================================================================
' MACRO 2: Find and Download Latest Reports from Synced SharePoint Folder
' ==============================================================================
Sub DownloadLatestReports()
    Dim ws As Worksheet
    Dim sourceFolder As String
    Dim destFolder As String
    Dim downloadsFolder As String
    Dim defaultPath As String
    Dim searchRange As Range
    Dim cell As Range
    Dim serialNum As String
    Dim currentFile As String
    Dim newestFile As String
    Dim fso As Object
    Dim foundCount As Long
    Dim missingCount As Long
    
    Set ws = ActiveSheet
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 1. Auto-Detect the Synced SharePoint Folder
    ' This uses the active user's Windows profile to silently find the synced Proteor folder
    defaultPath = Environ("USERPROFILE") & "\proteor.com\QualityControlDataSync - GC_Outoing_QC"
    
    If fso.FolderExists(defaultPath) Then
        sourceFolder = defaultPath
    Else
        ' Failsafe: If the folder isn't where we expect, ask the user to find it
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Select the Synced GC_Outgoing_QC Folder"
            If .Show = -1 Then
                sourceFolder = .SelectedItems(1)
            Else
                MsgBox "Folder selection canceled. Exiting macro.", vbExclamation
                Exit Sub
            End If
        End With
    End If
    
    ' Ensure trailing slash on source path
    If Right(sourceFolder, 1) <> "\" Then sourceFolder = sourceFolder & "\"
    
    ' 2. Create the Destination Folder in "Downloads"
    downloadsFolder = Environ("USERPROFILE") & "\Downloads"
    destFolder = downloadsFolder & "\" & Format(Date, "yyyy-mm-dd") & " - QNX0401 Test Reports"
    
    If Not fso.FolderExists(destFolder) Then
        fso.CreateFolder destFolder
    End If
    
    ' 3. Define the Search Area (Only look in columns K and beyond, protecting UI elements)
    Set searchRange = Intersect(ws.UsedRange, ws.Range("K:XFD"))
    
    If searchRange Is Nothing Then
        MsgBox "No data found in the extraction area to process.", vbInformation
        Exit Sub
    End If
    
    foundCount = 0
    missingCount = 0
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Searching for and copying reports... Please wait."
    
    ' 4. Loop Through the Excel Grid
    For Each cell In searchRange
        serialNum = Trim(cell.Value)
        
        ' Only process valid serial numbers that have NOT been flagged with yellow OCR errors
        If serialNum Like "[A-Z][A-Z]#[A-Z]####-###" And cell.Interior.Color <> vbYellow Then
            
            newestFile = ""
            ' Look for any PDF starting with this exact serial number
            currentFile = Dir(sourceFolder & serialNum & "*.pdf")
            
            ' Loop through all matches to find the one with the latest date/time string
            Do While currentFile <> ""
                If currentFile > newestFile Then
                    newestFile = currentFile
                End If
                currentFile = Dir
            Loop
            
            ' Execute the Copy and Color Code the Results
            If newestFile <> "" Then
                On Error Resume Next
                fso.CopyFile sourceFolder & newestFile, destFolder & "\" & newestFile, True
                
                If Err.Number = 0 Then
                    cell.Interior.Color = RGB(198, 239, 206) ' Light Green: Success
                    foundCount = foundCount + 1
                Else
                    cell.Interior.Color = vbRed ' Red: Found it, but failed to copy
                    missingCount = missingCount + 1
                End If
                On Error GoTo 0
            Else
                cell.Interior.Color = vbRed ' Red: No matching file found in the synced folder
                missingCount = missingCount + 1
            End If
            
        End If
    Next cell
    
    ' 5. Clean Up and Display Results
    Set fso = Nothing
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Automatically pop open the destination folder so the operator can see the files
    Call Shell("explorer.exe """ & destFolder & """", vbNormalFocus)
    
    MsgBox "Download process complete!" & vbCrLf & vbCrLf & _
           "Files Successfully Copied: " & foundCount & vbCrLf & _
           "Missing / Not Found: " & missingCount, vbInformation, "Status"
           
End Sub

