Attribute VB_Name = "RemoveExactDuplicates"
Option Explicit

Sub RemoveExactDuplicates()
    On Error GoTo ErrorHandler

    If MsgBox("Create a backup sheet before running?" & vbCrLf & _
              "A copy of the active sheet will be saved as Backup_YYYYMMDD.", _
              vbYesNo + vbQuestion, "Backup?") = vbYes Then
        Call CreateBackupSheet(ActiveSheet)
    End If

    Dim cell        As Range
    Dim compareRng  As Range
    Dim counts      As Object
    Dim cellVal     As String
    Dim removed     As Long

    Set compareRng = Selection
    removed = 0

    Set counts = CreateObject("Scripting.Dictionary")
    For Each cell In compareRng
        If Not IsEmpty(cell.Value) And Not IsError(cell.Value) Then
            cellVal = CStr(cell.Value)
            counts(cellVal) = counts(cellVal) + 1
        End If
    Next cell

    For Each cell In compareRng
        If Not IsEmpty(cell.Value) And Not IsError(cell.Value) Then
            If counts(CStr(cell.Value)) > 1 Then
                cell.ClearContents
                removed = removed + 1
            End If
        End If
    Next cell

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
        If Not IsEmpty(cell.Value) And Not IsError(cell.Value) Then
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
