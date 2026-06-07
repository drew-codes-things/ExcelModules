Attribute VB_Name = "CapitaliseFirstWords"
Option Explicit

Sub CapitalizeFirstLetter()
    On Error GoTo ErrorHandler

    ' --- Backup prompt ---
    If MsgBox("Create a backup sheet before running?" & vbCrLf & _
              "A copy of the active sheet will be saved as Backup_YYYYMMDD.", _
              vbYesNo + vbQuestion, "Backup?") = vbYes Then
        Call CreateBackupSheet(ActiveSheet)
    End If

    Dim cell     As Range
    Dim changed  As Long
    changed = 0

    For Each cell In Selection
        If Not IsEmpty(cell.Value) And Not IsNumeric(cell.Value) Then
            cell.Value = StrConv(CStr(cell.Value), vbProperCase)
            changed = changed + 1
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

    ' --- Backup prompt ---
    If MsgBox("Create a backup sheet before running?" & vbCrLf & _
              "A copy of the active sheet will be saved as Backup_YYYYMMDD.", _
              vbYesNo + vbQuestion, "Backup?") = vbYes Then
        Call CreateBackupSheet(ActiveSheet)
    End If

    Dim cell     As Range
    Dim txt      As String
    Dim changed  As Long
    changed = 0

    For Each cell In Selection
        If Not IsEmpty(cell.Value) And Not IsNumeric(cell.Value) Then
            txt = CStr(cell.Value)
            If Len(txt) > 0 Then
                cell.Value = UCase(Left(txt, 1)) & Mid(txt, 2)
                changed = changed + 1
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
