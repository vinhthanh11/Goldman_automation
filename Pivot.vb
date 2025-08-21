Option Explicit

'===============================
' Pivot Utilities
'===============================

' Refresh every PivotCache once (fast & avoids duplicate refreshes)
Public Sub RefreshAllPivotCaches()
    Dim pc As PivotCache
    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    For Each pc In ThisWorkbook.PivotCaches
        pc.Refresh
    Next pc

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    ' (Optional) Debug.Print "Refresh error: " & Err.Description
    Resume CleanExit
End Sub

' Refresh all pivot tables on a specific worksheet
Public Sub RefreshPivotsOnSheet(ByVal sheetName As String)
    Dim ws As Worksheet, pt As PivotTable
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' not found.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    For Each pt In ws.PivotTables
        pt.PivotCache.Refresh
    Next pt
    Application.ScreenUpdating = True
End Sub

' Refresh a specific pivot table by sheet & pivot name
Public Sub RefreshPivot(ByVal sheetName As String, ByVal pivotName As String)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = ThisWorkbook.Worksheets(sheetName).PivotTables(pivotName)
    On Error GoTo 0
    If pt Is Nothing Then
        MsgBox "Pivot '" & pivotName & "' on sheet '" & sheetName & "' not found.", vbExclamation
        Exit Sub
    End If
    pt.PivotCache.Refresh
End Sub

'========================================
' Copy the currently displayed pivot table
'========================================
' Copies the visible pivot result (respecting current filters) to destSheet as values.
' - includePageFields:=True copies report filter area too (use TableRange2); False omits it (TableRange1).
' - keepFormats:=True will paste values + formats; False = values only (clean snapshot).
' - Clears destSheet first by default.
Public Sub CopyPivotDisplayToSheet( _
    ByVal pivotSheet As String, _
    ByVal pivotName As String, _
    ByVal destSheet As String, _
    Optional ByVal includePageFields As Boolean = True, _
    Optional ByVal keepFormats As Boolean = True, _
    Optional ByVal clearDest As Boolean = True)

    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim pt As PivotTable
    Dim rngCopy As Range
    Dim firstCell As Range

    '--- locate source pivot
    On Error Resume Next
    Set wsSrc = ThisWorkbook.Worksheets(pivotSheet)
    On Error GoTo 0
    If wsSrc Is Nothing Then
        MsgBox "Source sheet '" & pivotSheet & "' not found.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set pt = wsSrc.PivotTables(pivotName)
    On Error GoTo 0
    If pt Is Nothing Then
        MsgBox "Pivot '" & pivotName & "' not found on sheet '" & pivotSheet & "'.", vbExclamation
        Exit Sub
    End If

    '--- range to copy
    ' TableRange1 = entire pivot (NO page fields)
    ' TableRange2 = entire pivot (WITH page fields/report filters)
    If includePageFields Then
        Set rngCopy = pt.TableRange2
    Else
        Set rngCopy = pt.TableRange1
    End If

    If rngCopy Is Nothing Then
        MsgBox "Pivot has no displayable range (is it empty?).", vbExclamation
        Exit Sub
    End If

    '--- get/create destination sheet
    Set wsDst = GetOrCreateSheet(destSheet)
    If clearDest Then wsDst.Cells.Clear

    Set firstCell = wsDst.Range("A1")

    Application.ScreenUpdating = False

    ' Copy & paste
    rngCopy.Copy
    If keepFormats Then
        firstCell.PasteSpecial xlPasteValues
        firstCell.PasteSpecial xlPasteFormats
    Else
        firstCell.PasteSpecial xlPasteValues
    End If
    Application.CutCopyMode = False

    ' tidy
    wsDst.Columns.AutoFit
    Application.ScreenUpdating = True
End Sub

' Helper: get or create worksheet by name in ThisWorkbook
Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function
