Attribute VB_Name = "StripWhitespace"
Option Explicit

Sub TrimCells()
    On Error GoTo ErrorHandler

    ' --- Backup prompt ---
    If MsgBox("Create a backup sheet before running?" & vbCrLf & _
              "A copy of the active sheet will be saved as Backup_YYYYMMDD.", _
              vbYesNo + vbQuestion, "Backup?") = vbYes Then
        Call CreateBackupSheet(ActiveSheet)
    End If

    Dim cell    As Range
    Dim changed As Long
    changed = 0

    For Each cell In Selection
        If Not IsEmpty(cell.Value) And Not IsNumeric(cell.Value) Then
            Dim trimmed As String
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

    ' --- Backup prompt ---
    If MsgBox("Create a backup sheet before running?" & vbCrLf & _
              "A copy of the active sheet will be saved as Backup_YYYYMMDD.", _
              vbYesNo + vbQuestion, "Backup?") = vbYes Then
        Call CreateBackupSheet(ActiveSheet)
    End If

    Dim cell    As Range
    Dim changed As Long
    changed = 0

    For Each cell In Selection
        If Not IsEmpty(cell.Value) And Not IsNumeric(cell.Value) Then
            Dim original As String
            original = CStr(cell.Value)
            Dim stripped As String
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
