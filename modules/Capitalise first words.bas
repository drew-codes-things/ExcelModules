Attribute VB_Name = "CapitaliseFirstWords"
Option Explicit

Sub CapitalizeFirstLetter()
    On Error GoTo ErrorHandler

    If MsgBox("Create a backup sheet before running?" & vbCrLf & _
              "A copy of the active sheet will be saved as Backup_YYYYMMDD.", _
              vbYesNo + vbQuestion, "Backup?") = vbYes Then
        Call CreateBackupSheet(ActiveSheet)
    End If

    Dim cell     As Range
    Dim txt      As String
    Dim newTxt   As String
    Dim changed  As Long
    changed = 0

    For Each cell In Selection
        If VarType(cell.Value) = vbString And Not cell.HasFormula Then
            txt = CStr(cell.Value)
            newTxt = StrConv(txt, vbProperCase)
            If newTxt <> txt Then
                cell.Value = newTxt
                changed = changed + 1
            End If
        End If
    Next cell

    MsgBox changed & " cell(s) converted to Proper Case.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Something went wrong during capitalisation." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Possible causes: the sheet is protected or a merged cell was encountered.", _
           vbCritical, "Capitalise Error"
End Sub

Sub CapitalizeFirstLetterOnly()
    On Error GoTo ErrorHandler

    If MsgBox("Create a backup sheet before running?" & vbCrLf & _
              "A copy of the active sheet will be saved as Backup_YYYYMMDD.", _
              vbYesNo + vbQuestion, "Backup?") = vbYes Then
        Call CreateBackupSheet(ActiveSheet)
    End If

    Dim cell     As Range
    Dim txt      As String
    Dim newTxt   As String
    Dim changed  As Long
    changed = 0

    For Each cell In Selection
        If VarType(cell.Value) = vbString And Not cell.HasFormula Then
            txt = CStr(cell.Value)
            If Len(txt) > 0 Then
                newTxt = UCase(Left(txt, 1)) & Mid(txt, 2)
                If newTxt <> txt Then
                    cell.Value = newTxt
                    changed = changed + 1
                End If
            End If
        End If
    Next cell

    MsgBox changed & " cell(s) updated (first letter only).", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Something went wrong during capitalisation." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Possible causes: the sheet is protected or a merged cell was encountered.", _
           vbCritical, "Capitalise Error"
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
