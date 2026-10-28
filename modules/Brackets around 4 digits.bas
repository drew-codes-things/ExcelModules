Attribute VB_Name = "BracketsAroundFourDigits"
Option Explicit

Sub AddBracketsToFourDigitNumbers()
    On Error GoTo ErrorHandler

    Dim cell      As Range
    Dim regex     As Object
    Dim original  As String
    Dim replaced  As String
    Dim changed   As Long

    If MsgBox("Create a backup sheet before running?" & vbCrLf & _
              "A copy of the active sheet will be saved as Backup_YYYYMMDD.", _
              vbYesNo + vbQuestion, "Backup?") = vbYes Then
        Call CreateBackupSheet(ActiveSheet)
    End If

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "(^|[^\(\d])(\d{4})(?![\)\d])"
    regex.Global = True

    changed = 0

    For Each cell In Selection
        If Not IsError(cell.Value) And Not IsEmpty(cell.Value) And Not cell.HasFormula Then
            If VarType(cell.Value) = vbString Or IsNumeric(cell.Value) Then
                original = CStr(cell.Value)
                If regex.Test(original) Then
                    replaced = regex.Replace(original, "$1($2)")
                    If IsNumeric(replaced) Then replaced = "'" & replaced
                    cell.Value = replaced
                    changed = changed + 1
                End If
            End If
        End If
    Next cell

    Set regex = Nothing

    If changed = 0 Then
        MsgBox "No standalone 4-digit numbers found in the selection.", vbInformation
    Else
        MsgBox changed & " cell(s) updated.", vbInformation
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Something went wrong while adding brackets." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Possible causes: the sheet is protected, VBScript.RegExp is unavailable, " & _
           "or a merged cell was encountered.", _
           vbCritical, "Brackets Error"
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
