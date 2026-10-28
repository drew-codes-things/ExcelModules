Attribute VB_Name = "HighlightDuplicates"
Option Explicit

Const HIGHLIGHT_COLOR     As Long = 65535
Const CROSS_SHEET_COLOR   As Long = 255

Sub HighlightDuplicatesMultiColumn()
    On Error GoTo ErrorHandler

    Dim ws          As Worksheet
    Dim rng         As Range
    Dim keys        As Object
    Dim rowKey      As String
    Dim hasValue    As Boolean
    Dim r           As Long
    Dim c           As Long
    Dim highlighted As Long

    Set ws = ActiveSheet
    Set rng = Selection

    If MsgBox("Create a backup sheet before running?" & vbCrLf & _
              "A copy of the active sheet will be saved as Backup_YYYYMMDD.", _
              vbYesNo + vbQuestion, "Backup?") = vbYes Then
        Call CreateBackupSheet(ws)
    End If

    Set keys = CreateObject("Scripting.Dictionary")
    highlighted = 0

    Dim firstRow As Long
    Dim lastRow  As Long
    Dim firstCol As Long
    Dim lastCol  As Long
    firstRow = rng.Row
    lastRow = rng.Row + rng.Rows.Count - 1
    firstCol = rng.Column
    lastCol = rng.Column + rng.Columns.Count - 1

    For r = firstRow To lastRow
        rowKey = RowKeyOf(ws, r, firstCol, lastCol, hasValue)
        If hasValue Then keys(rowKey) = keys(rowKey) + 1
    Next r

    For r = firstRow To lastRow
        rowKey = RowKeyOf(ws, r, firstCol, lastCol, hasValue)
        If hasValue Then
            If keys(rowKey) > 1 Then
                ws.Range(ws.Cells(r, firstCol), ws.Cells(r, lastCol)).Interior.Color = HIGHLIGHT_COLOR
                highlighted = highlighted + 1
            End If
        End If
    Next r

    If highlighted = 0 Then
        MsgBox "No duplicate rows found across the selected columns.", vbInformation
    Else
        MsgBox highlighted & " row(s) highlighted (combined key is duplicated).", vbInformation
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Something went wrong during multi-column duplicate highlighting." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Possible causes: the sheet is protected or a merged cell was encountered.", _
           vbCritical, "Highlight Error"
End Sub

Sub HighlightDuplicatesAcrossSheets()
    On Error GoTo ErrorHandler

    Dim wsA         As Worksheet
    Dim wsB         As Worksheet
    Dim rngA        As Range
    Dim cell        As Range
    Dim sheetBName  As String
    Dim lookups     As Object
    Dim highlighted As Long

    Set wsA = ActiveSheet
    Set rngA = Selection

    If MsgBox("Create a backup sheet before running?" & vbCrLf & _
              "A copy of the active sheet will be saved as Backup_YYYYMMDD.", _
              vbYesNo + vbQuestion, "Backup?") = vbYes Then
        Call CreateBackupSheet(wsA)
    End If

    sheetBName = InputBox( _
        "Enter the name of the second sheet to compare against:" & vbCrLf & _
        "(Cells in your current selection that also appear in the same" & vbCrLf & _
        " column on that sheet will be highlighted in red.)", _
        "Cross-Sheet Duplicate Check")

    If sheetBName = "" Then Exit Sub

    On Error Resume Next
    Set wsB = ActiveWorkbook.Sheets(sheetBName)
    On Error GoTo ErrorHandler
    If wsB Is Nothing Then
        MsgBox "Sheet '" & sheetBName & "' was not found in this workbook.", vbExclamation
        Exit Sub
    End If
    If wsB Is wsA Then
        MsgBox "Pick a different sheet - every cell would match itself.", vbExclamation
        Exit Sub
    End If

    Set lookups = CreateObject("Scripting.Dictionary")
    highlighted = 0
    For Each cell In rngA
        If Not IsEmpty(cell.Value) Then
            If Not lookups.exists(cell.Column) Then
                lookups.Add cell.Column, ColumnValues(wsB, cell.Column)
            End If
            If lookups(cell.Column).exists(CellKey(cell.Value)) Then
                cell.Interior.Color = CROSS_SHEET_COLOR
                highlighted = highlighted + 1
            End If
        End If
    Next cell

    If highlighted = 0 Then
        MsgBox "No matching values found between the two sheets.", vbInformation
    Else
        MsgBox highlighted & " cell(s) highlighted in red (also present on '" & sheetBName & "').", vbInformation
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Something went wrong during cross-sheet duplicate checking." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Possible causes: the sheet name was wrong, the sheet is protected, " & _
           "or a merged cell was encountered.", _
           vbCritical, "Cross-Sheet Error"
End Sub

Sub ClearDuplicateHighlights()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Selection.Interior.ColorIndex = xlNone
    MsgBox "Highlights cleared.", vbInformation
End Sub

Private Function CellKey(v As Variant) As String
    If IsError(v) Then
        CellKey = "#ERROR"
    Else
        CellKey = CStr(v)
    End If
End Function

Private Function RowKeyOf(ws As Worksheet, r As Long, firstCol As Long, lastCol As Long, _
                          ByRef hasValue As Boolean) As String
    Dim c As Long
    hasValue = False
    For c = firstCol To lastCol
        If Not IsEmpty(ws.Cells(r, c).Value) Then hasValue = True
        RowKeyOf = RowKeyOf & CellKey(ws.Cells(r, c).Value) & Chr(0)
    Next c
End Function

Private Function ColumnValues(ws As Worksheet, col As Long) As Object
    Dim values  As Object
    Dim lastRow As Long
    Dim r       As Long
    Set values = CreateObject("Scripting.Dictionary")
    values.CompareMode = vbTextCompare
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    For r = 1 To lastRow
        If Not IsEmpty(ws.Cells(r, col).Value) Then values(CellKey(ws.Cells(r, col).Value)) = True
    Next r
    Set ColumnValues = values
End Function

Private Sub CreateBackupSheet(ws As Worksheet)
    Dim backupName As String
    Dim testName   As String
    Dim counter    As Integer

    backupName = "Backup_" & Format(Now(), "YYYYMMDD")
    testName = backupName
    counter = 1
    Do While SheetExists(ws.Parent, testName)
        testName = backupName & "_" & counter
        counter = counter + 1
    Loop
    ws.Copy After:=ws.Parent.Sheets(ws.Parent.Sheets.Count)
    ActiveSheet.Name = testName
    ws.Activate
    MsgBox "Backup created: '" & testName & "'", vbInformation, "Backup"
End Sub

Private Function SheetExists(wb As Workbook, shName As String) As Boolean
    Dim sh As Object
    On Error Resume Next
    Set sh = wb.Sheets(shName)
    On Error GoTo 0
    SheetExists = Not sh Is Nothing
End Function
