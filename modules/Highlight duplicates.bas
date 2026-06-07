Attribute VB_Name = "HighlightDuplicates"
Option Explicit

Const HIGHLIGHT_COLOR     As Long = 65535   ' Yellow
Const CROSS_SHEET_COLOR   As Long = 16711680 ' Red

' =======================================================================
' Sub 1: Highlight rows where the COMBINED KEY across all selected
'        columns is duplicated. Works across multiple columns simultaneously
'        - not possible with built-in Conditional Formatting.
' Usage: Select the range (including all columns to check), run macro.
' =======================================================================
Sub HighlightDuplicatesMultiColumn()
    On Error GoTo ErrorHandler

    Dim ws          As Worksheet
    Dim rng         As Range
    Dim cell        As Range
    Dim keys        As Object
    Dim rowKey      As String
    Dim r           As Long
    Dim c           As Long
    Dim highlighted As Long

    Set ws  = ActiveSheet
    Set rng = Selection

    ' --- Backup prompt ---
    If MsgBox("Create a backup sheet before running?" & vbCrLf & _
              "A copy of the active sheet will be saved as Backup_YYYYMMDD.", _
              vbYesNo + vbQuestion, "Backup?") = vbYes Then
        Call CreateBackupSheet(ws)
    End If

    Set keys = CreateObject("Scripting.Dictionary")
    highlighted = 0

    ' First pass: build a key for every row in the selection
    Dim firstRow As Long
    Dim lastRow  As Long
    Dim firstCol As Long
    Dim lastCol  As Long
    firstRow = rng.Row
    lastRow  = rng.Row + rng.Rows.Count - 1
    firstCol = rng.Column
    lastCol  = rng.Column + rng.Columns.Count - 1

    For r = firstRow To lastRow
        rowKey = ""
        For c = firstCol To lastCol
            rowKey = rowKey & CStr(ws.Cells(r, c).Value) & Chr(0)
        Next c
        If keys.exists(rowKey) Then
            keys(rowKey) = keys(rowKey) + 1
        Else
            keys.Add rowKey, 1
        End If
    Next r

    ' Second pass: highlight rows whose key appeared more than once
    For r = firstRow To lastRow
        rowKey = ""
        For c = firstCol To lastCol
            rowKey = rowKey & CStr(ws.Cells(r, c).Value) & Chr(0)
        Next c
        If keys(rowKey) > 1 Then
            For c = firstCol To lastCol
                ws.Cells(r, c).Interior.Color = HIGHLIGHT_COLOR
            Next c
            highlighted = highlighted + 1
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


' =======================================================================
' Sub 2: Highlight cells in Sheet A column that also appear in the SAME
'        column on Sheet B. Built-in Conditional Formatting cannot compare
'        across sheets.
' Usage: Select the column range on Sheet A, run macro. You will be
'        prompted for the name of Sheet B.
' =======================================================================
Sub HighlightDuplicatesAcrossSheets()
    On Error GoTo ErrorHandler

    Dim wsA         As Worksheet
    Dim wsB         As Worksheet
    Dim rngA        As Range
    Dim cell        As Range
    Dim sheetBName  As String
    Dim colB        As Range
    Dim highlighted As Long

    Set wsA  = ActiveSheet
    Set rngA = Selection

    ' --- Backup prompt ---
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

    ' Build the comparison range on Sheet B (same column, used range rows)
    Set colB = wsB.Range( _
        wsB.Cells(1, rngA.Column), _
        wsB.Cells(wsB.Cells(wsB.Rows.Count, rngA.Column).End(xlUp).Row, rngA.Column))

    highlighted = 0
    For Each cell In rngA
        If Not IsEmpty(cell.Value) Then
            If WorksheetFunction.CountIf(colB, cell.Value) > 0 Then
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
    Dim cell As Range
    For Each cell In Selection
        cell.Interior.ColorIndex = xlNone
    Next cell
    MsgBox "Highlights cleared.", vbInformation
End Sub
