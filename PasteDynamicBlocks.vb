Sub PasteDynamicBlocks()
    Dim wsInstr As Worksheet
    Dim wsSrc1 As Worksheet, wsSrc2 As Worksheet, wsDst As Worksheet, wsTs As Worksheet
    Dim srcRange1 As String, srcRange2 As String
    Dim dstCol1 As String, dstCol2 As String
    Dim ts As Variant
    Dim arr1 As Variant, arr2 As Variant
    Dim rows1 As Long, cols1 As Long, rows2 As Long, cols2 As Long
    Dim nextRow As Long, totalRows As Long, lastRow As Long, endCol As Long
    
    Set wsInstr = ThisWorkbook.Worksheets("Instruction")
    
    '--- read all settings from Instruction
    Set wsSrc1 = ThisWorkbook.Worksheets(wsInstr.Range("B2").Value) ' Client sheet name
    srcRange1 = wsInstr.Range("B3").Value ' e.g., "A2:N"
    dstCol1 = wsInstr.Range("B4").Value   ' Client start column
    
    Set wsSrc2 = ThisWorkbook.Worksheets(wsInstr.Range("B5").Value) ' Goldman sheet name
    srcRange2 = wsInstr.Range("B6").Value ' e.g., "A2:N"
    dstCol2 = wsInstr.Range("B7").Value   ' Goldman start column
    
    Set wsDst = ThisWorkbook.Worksheets(wsInstr.Range("B8").Value)  ' Transaction sheet name
    
    Set wsTs = ThisWorkbook.Worksheets(wsInstr.Range("B9").Value)   ' Timestamp sheet name
    ts = wsTs.Range(wsInstr.Range("B10").Value).Value               ' Timestamp cell
    
    '--- find last rows in sources
    Dim last1 As Long, last2 As Long
    last1 = wsSrc1.Cells(wsSrc1.Rows.Count, Left(srcRange1, 1)).End(xlUp).Row
    last2 = wsSrc2.Cells(wsSrc2.Rows.Count, Left(srcRange2, 1)).End(xlUp).Row
    
    '--- load data arrays
    If last1 >= 2 Then
        arr1 = wsSrc1.Range(srcRange1 & last1).Value
        rows1 = UBound(arr1, 1)
        cols1 = UBound(arr1, 2)
    End If
    
    If last2 >= 2 Then
        arr2 = wsSrc2.Range(srcRange2 & last2).Value
        rows2 = UBound(arr2, 1)
        cols2 = UBound(arr2, 2)
    End If
    
    '--- first empty row in destination col A
    nextRow = wsDst.Cells(wsDst.Rows.Count, "A").End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2
    
    '--- total rows to paste
    totalRows = Application.Max(rows1, rows2)
    If totalRows = 0 Then
        MsgBox "No data to paste from either source.", vbInformation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    '--- fill timestamps
    wsDst.Range(wsDst.Cells(nextRow, "A"), wsDst.Cells(nextRow + totalRows - 1, "A")).Value = ts
    
    '--- paste Client data
    If rows1 > 0 Then
        wsDst.Cells(nextRow, dstCol1).Resize(rows1, cols1).Value = arr1
    End If
    
    '--- paste Goldman data
    If rows2 > 0 Then
        wsDst.Cells(nextRow, dstCol2).Resize(rows2, cols2).Value = arr2
    End If
    
    '--- determine last row & end column touched
    lastRow = nextRow + totalRows - 1
    endCol = WorksheetFunction.Max( _
        wsDst.Columns(dstCol1).Column + cols1 - 1, _
        wsDst.Columns(dstCol2).Column + cols2 - 1)
    
    '--- draw top border
    With wsDst.Range(wsDst.Cells(nextRow, 1), wsDst.Cells(nextRow, endCol)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    '--- draw bottom border
    With wsDst.Range(wsDst.Cells(lastRow, 1), wsDst.Cells(lastRow, endCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
