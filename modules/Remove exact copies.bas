Attribute VB_Name = "RemoveExactDuplicates"
Option Explicit

Sub RemoveExactDuplicates()
    On Error GoTo ErrorHandler

    ' --- Backup prompt ---
    If MsgBox("Create a backup sheet before running?" & vbCrLf & _
              "A copy of the active sheet will be saved as Backup_YYYYMMDD.", _
              vbYesNo + vbQuestion, "Backup?") = vbYes Then
        Call CreateBackupSheet(ActiveSheet)
    End If

    Dim cell        As Range
    Dim compareRng  As Range
    Dim cellVal     As String
    Dim removed     As Long

    Set compareRng = Selection
    removed = 0

    Dim markedCells As Collection
    Set markedCells = New Collection

    For Each cell In compareRng
        If Not IsEmpty(cell.Value) Then
            cellVal = CStr(cell.Value)
            If WorksheetFunction.CountIf(compareRng, cellVal) > 1 Then
                markedCells.Add cell
            End If
        End If
    Next cell

    Dim c As Variant
    For Each c In markedCells
        c.ClearContents
        removed = removed + 1
    Next c

    MsgBox removed & " duplicate cell(s) cleared.", vbInformation, "Remove Exact Duplicates"
    Exit Sub

ErrorHandler:
    MsgBox "Something went wrong while removing duplicates." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Possible causes: the sheet is protected or a merged cell was encountered.", _
           vbCritical, "Remove Duplicates Error"
End Sub


Sub RemoveDuplicatesKeepFirst()
    On Error GoTo ErrorHandler

    ' --- Backup prompt ---
    If MsgBox("Create a backup sheet before running?" & vbCrLf & _
              "A copy of the active sheet will be saved as Backup_YYYYMMDD.", _
              vbYesNo + vbQuestion, "Backup?") = vbYes Then
        Call CreateBackupSheet(ActiveSheet)
    End If

    Dim cell        As Range
    Dim compareRng  As Range
    Dim seen        As Object
    Dim cellVal     As String
    Dim removed     As Long

    Set compareRng = Selection
    Set seen = CreateObject("Scripting.Dictionary")
    removed = 0

    For Each cell In compareRng
        If Not IsEmpty(cell.Value) Then
            cellVal = CStr(cell.Value)
            If seen.exists(cellVal) Then
                cell.ClearContents
                removed = removed + 1
            Else
                seen.Add cellVal, True
            End If
        End If
    Next cell

    MsgBox removed & " duplicate(s) removed. First occurrences kept.", _
           vbInformation, "Remove Duplicates Keep First"
    Exit Sub

ErrorHandler:
    MsgBox "Something went wrong while removing duplicates." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Possible causes: the sheet is protected or a merged cell was encountered.", _
           vbCritical, "Remove Duplicates Error"
End Sub
