Attribute VB_Name = "modConsolidator"
'==============================================================================
' Excel File Consolidator
' Author : Swaminathan Narayanan
' Repo   : https://github.com/Swami529796/excel-file-consolidator
'
' Combines every Excel / CSV file in a chosen folder into one "Master" sheet.
'  - Validates that each file has the same headers as the first file
'  - Adds "Source File" and "Source Sheet" columns for traceability
'  - Logs every file (loaded / skipped / error) to a "Log" sheet
'  - Uses arrays (not cell-by-cell copying) so it stays fast on large files
'
' Run: ConsolidateFiles
'==============================================================================
Option Explicit

'---- Settings ----------------------------------------------------------------
Private Const MASTER_SHEET As String = "Master"
Private Const LOG_SHEET As String = "Log"
Private Const ALL_SHEETS As Boolean = False     ' False = first sheet of each file only
Private Const SKIP_BLANK_ROWS As Boolean = True ' True  = ignore completely empty rows

'---- Run state ---------------------------------------------------------------
Private mHeaders As Variant     ' expected headers (taken from the first valid sheet)
Private mHeaderCount As Long
Private mNextRow As Long        ' next free row on Master
Private mLogRow As Long         ' next free row on Log
Private mRowsAdded As Long
Private mLoaded As Long
Private mSkipped As Long
Private mCalcMode As XlCalculation

'==============================================================================
' Entry point
'==============================================================================
Public Sub ConsolidateFiles()
    Dim folderPath As String
    folderPath = PickFolder()
    If Len(folderPath) = 0 Then Exit Sub

    Dim files As Collection
    Set files = ListFiles(folderPath)
    If files.Count = 0 Then
        MsgBox "No Excel or CSV files found in:" & vbLf & folderPath, vbExclamation
        Exit Sub
    End If

    Dim startTime As Double
    startTime = Timer

    Dim wsMaster As Worksheet, wsLog As Worksheet
    Set wsMaster = GetOrCreateSheet(MASTER_SHEET)
    Set wsLog = GetOrCreateSheet(LOG_SHEET)

    On Error GoTo CleanFail
    SpeedUp True

    ' reset output and counters
    wsMaster.Cells.Clear
    wsLog.Cells.Clear
    wsLog.Range("A1:E1").Value = Array("File", "Sheet", "Rows Added", "Status", "Message")
    mHeaderCount = 0
    mNextRow = 2
    mLogRow = 2
    mRowsAdded = 0
    mLoaded = 0
    mSkipped = 0

    Dim f As Variant
    For Each f In files
        Application.StatusBar = "Consolidating: " & f
        ProcessFile folderPath, CStr(f), wsMaster, wsLog
    Next f

    FormatOutput wsMaster, wsLog
    SpeedUp False
    Application.StatusBar = False

    MsgBox "Consolidation complete." & vbLf & vbLf & _
           "Files found:      " & files.Count & vbLf & _
           "Sheets loaded:    " & mLoaded & vbLf & _
           "Skipped / errors: " & mSkipped & vbLf & _
           "Rows added:       " & Format(mRowsAdded, "#,##0") & vbLf & _
           "Time taken:       " & Format(Timer - startTime, "0.0") & " sec" & vbLf & vbLf & _
           "See the '" & LOG_SHEET & "' sheet for details.", vbInformation
    Exit Sub

CleanFail:
    SpeedUp False
    Application.StatusBar = False
    MsgBox "Consolidation stopped:" & vbLf & Err.Description, vbCritical
End Sub

'==============================================================================
' Open one file and process its sheet(s)
'==============================================================================
Private Sub ProcessFile(ByVal folderPath As String, ByVal fileName As String, _
                        ByVal wsMaster As Worksheet, ByVal wsLog As Worksheet)
    Dim wb As Workbook
    Dim ws As Worksheet

    On Error GoTo OpenFail
    Set wb = Workbooks.Open(Filename:=folderPath & fileName, UpdateLinks:=0, ReadOnly:=True)

    On Error GoTo ProcessFail
    For Each ws In wb.Worksheets
        ProcessSheet ws, fileName, wsMaster, wsLog
        If Not ALL_SHEETS Then Exit For
    Next ws

    wb.Close SaveChanges:=False
    Exit Sub

OpenFail:
    WriteLog wsLog, fileName, "", 0, "Error", "Could not open file: " & Err.Description
    Exit Sub

ProcessFail:
    WriteLog wsLog, fileName, "", 0, "Error", Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
End Sub

'==============================================================================
' Validate headers and append one sheet's data to Master
'==============================================================================
Private Sub ProcessSheet(ByVal ws As Worksheet, ByVal fileName As String, _
                         ByVal wsMaster As Worksheet, ByVal wsLog As Worksheet)
    Dim lastRow As Long, lastCol As Long
    lastRow = LastUsedRow(ws)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        WriteLog wsLog, fileName, ws.Name, 0, "Skipped", "No data rows"
        Exit Sub
    End If

    ' read the whole block into memory in one go
    Dim data As Variant
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value

    Dim c As Long
    If mHeaderCount = 0 Then
        ' the first valid sheet defines the expected layout
        mHeaderCount = lastCol
        ReDim mHeaders(1 To lastCol)
        For c = 1 To lastCol
            mHeaders(c) = Trim$(CStr(data(1, c)))
        Next c
        WriteMasterHeader wsMaster
    Else
        Dim mismatch As String
        mismatch = CompareHeaders(data, lastCol)
        If Len(mismatch) > 0 Then
            WriteLog wsLog, fileName, ws.Name, 0, "Skipped", "Header mismatch - " & mismatch
            Exit Sub
        End If
    End If

    ' build the output block: Source File | Source Sheet | data...
    Dim outArr() As Variant
    ReDim outArr(1 To lastRow - 1, 1 To mHeaderCount + 2)

    Dim r As Long, n As Long
    For r = 2 To lastRow
        If Not (SKIP_BLANK_ROWS And RowIsBlank(data, r, mHeaderCount)) Then
            n = n + 1
            outArr(n, 1) = fileName
            outArr(n, 2) = ws.Name
            For c = 1 To mHeaderCount
                outArr(n, c + 2) = data(r, c)
            Next c
        End If
    Next r

    If n = 0 Then
        WriteLog wsLog, fileName, ws.Name, 0, "Skipped", "Only blank rows"
        Exit Sub
    End If

    If mNextRow + n - 1 > wsMaster.Rows.Count Then
        WriteLog wsLog, fileName, ws.Name, 0, "Error", "Master sheet row limit reached"
        Exit Sub
    End If

    ' write everything in a single operation
    wsMaster.Cells(mNextRow, 1).Resize(n, mHeaderCount + 2).Value = outArr
    mNextRow = mNextRow + n
    mRowsAdded = mRowsAdded + n
    WriteLog wsLog, fileName, ws.Name, n, "Loaded", ""
End Sub

'==============================================================================
' Helpers
'==============================================================================
Private Function CompareHeaders(ByVal data As Variant, ByVal lastCol As Long) As String
    If lastCol <> mHeaderCount Then
        CompareHeaders = "expected " & mHeaderCount & " columns, found " & lastCol
        Exit Function
    End If

    Dim c As Long, found As String
    For c = 1 To lastCol
        found = Trim$(CStr(data(1, c)))
        If StrComp(found, mHeaders(c), vbTextCompare) <> 0 Then
            CompareHeaders = "column " & c & ": expected '" & mHeaders(c) & "', found '" & found & "'"
            Exit Function
        End If
    Next c
End Function

Private Function RowIsBlank(ByVal data As Variant, ByVal r As Long, ByVal colCount As Long) As Boolean
    Dim c As Long, v As Variant
    For c = 1 To colCount
        v = data(r, c)
        If IsError(v) Then Exit Function
        If Not IsEmpty(v) Then
            If Len(Trim$(CStr(v))) > 0 Then Exit Function
        End If
    Next c
    RowIsBlank = True
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim found As Range
    Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                              SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not found Is Nothing Then LastUsedRow = found.Row
End Function

Private Sub WriteMasterHeader(ByVal wsMaster As Worksheet)
    Dim c As Long
    wsMaster.Cells(1, 1).Value = "Source File"
    wsMaster.Cells(1, 2).Value = "Source Sheet"
    For c = 1 To mHeaderCount
        wsMaster.Cells(1, c + 2).Value = mHeaders(c)
    Next c
End Sub

Private Sub WriteLog(ByVal wsLog As Worksheet, ByVal fileName As String, ByVal sheetName As String, _
                     ByVal rowsAdded As Long, ByVal status As String, ByVal message As String)
    wsLog.Cells(mLogRow, 1).Resize(1, 5).Value = Array(fileName, sheetName, rowsAdded, status, message)
    mLogRow = mLogRow + 1
    If status = "Loaded" Then
        mLoaded = mLoaded + 1
    Else
        mSkipped = mSkipped + 1
    End If
End Sub

Private Sub FormatOutput(ByVal wsMaster As Worksheet, ByVal wsLog As Worksheet)
    Dim r As Long

    With wsLog
        .Rows(1).Font.Bold = True
        For r = 2 To mLogRow - 1
            Select Case .Cells(r, 4).Value
                Case "Loaded":  .Cells(r, 4).Font.Color = RGB(0, 128, 0)
                Case "Skipped": .Cells(r, 4).Font.Color = RGB(204, 102, 0)
                Case "Error":   .Cells(r, 4).Font.Color = RGB(192, 0, 0)
            End Select
        Next r
        .UsedRange.Columns.AutoFit
    End With

    With wsMaster
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(221, 235, 247)
        .UsedRange.Columns.AutoFit
        .Activate
    End With

    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
End Sub

Private Function ListFiles(ByVal folderPath As String) As Collection
    Dim col As New Collection
    Dim pattern As Variant, f As String

    ' collect names first - calling Dir() while opening files is unreliable
    For Each pattern In Array("*.xls*", "*.csv")
        f = Dir(folderPath & pattern)
        Do While Len(f) > 0
            If Left$(f, 2) <> "~$" And _
               StrComp(folderPath & f, ThisWorkbook.FullName, vbTextCompare) <> 0 Then
                col.Add f
            End If
            f = Dir()
        Loop
    Next pattern

    Set ListFiles = col
End Function

Private Function PickFolder() As String
    With Application.FileDialog(4) ' msoFileDialogFolderPicker
        .Title = "Select the folder containing the files to consolidate"
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
            If Right$(PickFolder, 1) <> Application.PathSeparator Then
                PickFolder = PickFolder & Application.PathSeparator
            End If
        End If
    End With
End Function

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Private Sub SpeedUp(ByVal turnOn As Boolean)
    With Application
        If turnOn Then
            mCalcMode = .Calculation
            .Calculation = xlCalculationManual
        ElseIf mCalcMode <> 0 Then
            .Calculation = mCalcMode
        End If
        .ScreenUpdating = Not turnOn
        .EnableEvents = Not turnOn
        .DisplayAlerts = Not turnOn
    End With
End Sub
