Attribute VB_Name = "AlphabeticalSort"
Option Explicit

Sub SortSelectedColumns()
    On Error GoTo ErrorHandler

    Dim ws          As Worksheet
    Dim rng         As Range
    Dim colToSort   As Range
    Dim sortDir     As XlSortOrder
    Dim lastRow     As Long
    Dim colIdx      As Integer
    Dim hasHeader   As Integer
    Dim response    As Integer
    Dim cell        As Range
    Dim hasText     As Boolean

    Set ws = ActiveSheet

    If Selection Is Nothing Then
        MsgBox "No selection found. Please select at least one column.", vbExclamation
        Exit Sub
    End If

    ' --- Backup prompt ---
    If MsgBox("Create a backup sheet before running?" & vbCrLf & _
              "A copy of the active sheet will be saved as Backup_YYYYMMDD.", _
              vbYesNo + vbQuestion, "Backup?") = vbYes Then
        Call CreateBackupSheet(ws)
    End If

    response = MsgBox("Sort Ascending? Click No for Descending, Cancel to abort.", _
                      vbYesNoCancel + vbQuestion, "Sort Direction")
    If response = vbCancel Then Exit Sub
    sortDir = IIf(response = vbYes, xlAscending, xlDescending)

    hasHeader = MsgBox("Does your selection include a header row in row 1?", _
                       vbYesNo + vbQuestion, "Header Row?")

    Set rng = Selection
    lastRow = ws.Cells(ws.Rows.Count, rng.Column).End(xlUp).Row

    For colIdx = 1 To rng.Columns.Count
        Set colToSort = ws.Range( _
            rng.Cells(1, colIdx), _
            ws.Cells(lastRow, rng.Cells(1, colIdx).Column))

        hasText = False
        For Each cell In colToSort
            If cell.Value Like "*[A-Za-z]*" Then
                hasText = True
                Exit For
            End If
        Next cell

        If hasText Then
            ws.Sort.SortFields.Clear
            ws.Sort.SortFields.Add Key:=colToSort, Order:=sortDir
            With ws.Sort
                .SetRange colToSort
                .Header = IIf(hasHeader = vbYes, xlYes, xlNo)
                .MatchCase = False
                .Orientation = xlTopToBottom
                .Apply
            End With
        End If
    Next colIdx

    MsgBox "Done. Alphabetical columns sorted " & _
           IIf(sortDir = xlAscending, "A → Z", "Z → A") & ".", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Something went wrong while sorting." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Possible causes: the sheet is protected, a cell is merged, " & _
           "or the selection contains an unexpected data type.", _
           vbCritical, "Sort Error"
End Sub

' -----------------------------------------------------------------------
' Shared helper: copy the active sheet to a new tab named Backup_YYYYMMDD
' -----------------------------------------------------------------------
Sub CreateBackupSheet(ws As Worksheet)
    Dim backupName As String
    backupName = "Backup_" & Format(Now(), "YYYYMMDD")
    ' Avoid duplicate names by appending a counter
    Dim counter As Integer
    counter = 1
    Dim testName As String
    testName = backupName
    Do While SheetExists(ws.Parent, testName)
        testName = backupName & "_" & counter
        counter = counter + 1
    Loop
    ws.Copy After:=ws.Parent.Sheets(ws.Parent.Sheets.Count)
    ws.Parent.Sheets(ws.Parent.Sheets.Count).Name = testName
    MsgBox "Backup created: '" & testName & "'", vbInformation, "Backup"
End Sub

Function SheetExists(wb As Workbook, shName As String) As Boolean
    Dim sh As Object
    On Error Resume Next
    Set sh = wb.Sheets(shName)
    On Error GoTo 0
    SheetExists = Not sh Is Nothing
End Function
