Attribute VB_Name = "AlphabeticalSort"
Option Explicit

Sub SortSelectedColumns()
    On Error GoTo ErrorHandler

    Dim ws          As Worksheet
    Dim rng         As Range
    Dim colToSort   As Range
    Dim sortDir     As XlSortOrder
    Dim lastRow     As Long
    Dim colIdx      As Long
    Dim hasHeader   As Integer
    Dim response    As Integer
    Dim cell        As Range
    Dim hasText     As Boolean

    If TypeName(Selection) <> "Range" Then
        MsgBox "No selection found. Please select at least one column.", vbExclamation
        Exit Sub
    End If

    Set ws = ActiveSheet
    Set rng = Selection

    If MsgBox("Create a backup sheet before running?" & vbCrLf & _
              "A copy of the active sheet will be saved as Backup_YYYYMMDD.", _
              vbYesNo + vbQuestion, "Backup?") = vbYes Then
        Call CreateBackupSheet(ws)
    End If

    response = MsgBox("Sort Ascending? Click No for Descending, Cancel to abort.", _
                      vbYesNoCancel + vbQuestion, "Sort Direction")
    If response = vbCancel Then Exit Sub
    sortDir = IIf(response = vbYes, xlAscending, xlDescending)

    hasHeader = MsgBox("Does the first row of your selection contain headers?", _
                       vbYesNo + vbQuestion, "Header Row?")

    For colIdx = 1 To rng.Columns.Count
        lastRow = ws.Cells(ws.Rows.Count, rng.Cells(1, colIdx).Column).End(xlUp).Row
        If lastRow >= rng.Row Then
            Set colToSort = ws.Range( _
                rng.Cells(1, colIdx), _
                ws.Cells(lastRow, rng.Cells(1, colIdx).Column))

            hasText = False
            For Each cell In colToSort
                If Not IsError(cell.Value) Then
                    If CStr(cell.Value) Like "*[A-Za-z]*" Then
                        hasText = True
                        Exit For
                    End If
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
        End If
    Next colIdx

    MsgBox "Done. Alphabetical columns sorted " & _
           IIf(sortDir = xlAscending, "A to Z", "Z to A") & ".", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Something went wrong while sorting." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Possible causes: the sheet is protected, a cell is merged, " & _
           "or the selection contains an unexpected data type.", _
           vbCritical, "Sort Error"
End Sub

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
