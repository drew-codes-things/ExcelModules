Attribute VB_Name = "StripWhitespace"
Option Explicit

Sub TrimCells()
    On Error GoTo ErrorHandler

    If MsgBox("Create a backup sheet before running?" & vbCrLf & _
              "A copy of the active sheet will be saved as Backup_YYYYMMDD.", _
              vbYesNo + vbQuestion, "Backup?") = vbYes Then
        Call CreateBackupSheet(ActiveSheet)
    End If

    Dim cell    As Range
    Dim trimmed As String
    Dim changed As Long
    changed = 0

    For Each cell In Selection
        If VarType(cell.Value) = vbString And Not cell.HasFormula Then
            trimmed = WorksheetFunction.Trim(CStr(cell.Value))
            If trimmed <> CStr(cell.Value) Then
                cell.Value = trimmed
                changed = changed + 1
            End If
        End If
    Next cell

    MsgBox changed & " cell(s) trimmed.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Something went wrong while trimming whitespace." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Possible causes: the sheet is protected or a merged cell was encountered.", _
           vbCritical, "Trim Error"
End Sub

Sub RemoveAllSpaces()
    On Error GoTo ErrorHandler

    If MsgBox("Create a backup sheet before running?" & vbCrLf & _
              "A copy of the active sheet will be saved as Backup_YYYYMMDD.", _
              vbYesNo + vbQuestion, "Backup?") = vbYes Then
        Call CreateBackupSheet(ActiveSheet)
    End If

    Dim cell     As Range
    Dim original As String
    Dim stripped As String
    Dim changed  As Long
    changed = 0

    For Each cell In Selection
        If VarType(cell.Value) = vbString And Not cell.HasFormula Then
            original = CStr(cell.Value)
            stripped = Replace(original, " ", "")
            If stripped <> original Then
                cell.Value = stripped
                changed = changed + 1
            End If
        End If
    Next cell

    MsgBox changed & " cell(s) updated.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Something went wrong while removing spaces." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Possible causes: the sheet is protected or a merged cell was encountered.", _
           vbCritical, "Remove Spaces Error"
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
