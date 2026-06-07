Attribute VB_Name = "BracketsAroundFourDigits"
Option Explicit

Sub AddBracketsToFourDigitNumbers()
    On Error GoTo ErrorHandler

    Dim cell      As Range
    Dim regex     As Object
    Dim original  As String
    Dim changed   As Long

    ' --- Backup prompt ---
    If MsgBox("Create a backup sheet before running?" & vbCrLf & _
              "A copy of the active sheet will be saved as Backup_YYYYMMDD.", _
              vbYesNo + vbQuestion, "Backup?") = vbYes Then
        Call CreateBackupSheet(ActiveSheet)
    End If

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "(?<![\(\d])(\d{4})(?![\)\d])"
    regex.Global = True

    changed = 0

    For Each cell In Selection
        If Not IsEmpty(cell.Value) And Not IsNumeric(cell.Value) = False _
           Or (Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString) Then
            original = CStr(cell.Value)
            If regex.Test(original) Then
                cell.Value = regex.Replace(original, "($1)")
                changed = changed + 1
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
