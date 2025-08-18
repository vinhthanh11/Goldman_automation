Option Explicit

' ========= MAIN (runs everything per Config) =========
Sub RunConfigWithDependencies()
    Dim wsConfig As Worksheet, order As Collection
    Dim i As Long, cfgRow As Long
    Dim sheetName As String, sourceName As String, parentName As String
    Dim filterRules As String, keepCols As String, renameMap As String, options As String
    Dim wsInput As Worksheet, wsTarget As Worksheet

    Set wsConfig = ThisWorkbook.Sheets("Config")
    Set order = GetExecutionOrder(wsConfig)
    If order Is Nothing Then Exit Sub

    For i = 1 To order.Count
        sheetName = CStr(order(i))
        cfgRow = FindConfigRow(wsConfig, sheetName)
        If cfgRow = 0 Then GoTo NextItem

        sourceName = Trim(CStr(wsConfig.Cells(cfgRow, 2).Value))   ' Source
        parentName = Trim(CStr(wsConfig.Cells(cfgRow, 3).Value))   ' ParentReport
        filterRules = CStr(wsConfig.Cells(cfgRow, 4).Value)
        keepCols = CStr(wsConfig.Cells(cfgRow, 5).Value)
        renameMap = CStr(wsConfig.Cells(cfgRow, 6).Value)
        options   = CStr(wsConfig.Cells(cfgRow, 7).Value)

        ' Decide input sheet
        Set wsInput = Nothing
        If sourceName <> "" Then
            If SheetExists(sourceName) Then Set wsInput = ThisWorkbook.Sheets(sourceName)
        ElseIf parentName <> "" Then
            If SheetExists(parentName) Then Set wsInput = ThisWorkbook.Sheets(parentName)
        End If
        If wsInput Is Nothing Then GoTo NextItem

        ' Ensure target exists
        Set wsTarget = GetOrCreateSheet(sheetName)

        ' Process
        FilterAndCopy_Flex wsInput, wsTarget, filterRules, keepCols
        ApplyRenameMap wsTarget, renameMap
        ApplyOptions   wsTarget, options
NextItem:
    Next i
End Sub

' ========= DEPENDENCY RESOLVER (topological sort) =========
Function GetExecutionOrder(wsConfig As Worksheet) As Collection
    Dim graph As Object, indegree As Object
    Dim lastRow As Long, i As Long
    Dim sName As String, pName As String
    Dim q As Collection, execOrder As Collection
    Dim key As Variant, child As Variant

    Set graph = CreateObject("Scripting.Dictionary")
    Set indegree = CreateObject("Scripting.Dictionary")
    Set execOrder = New Collection
    Set q = New Collection

    lastRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        sName = CStr(wsConfig.Cells(i, 1).Value)  ' SheetName
        pName = CStr(wsConfig.Cells(i, 3).Value)  ' ParentReport

        If Not graph.Exists(sName) Then Set graph(sName) = CreateObject("Scripting.Dictionary")
        If Not indegree.Exists(sName) Then indegree(sName) = 0

        If Len(pName) > 0 Then
            If Not graph.Exists(pName) Then Set graph(pName) = CreateObject("Scripting.Dictionary")
            If Not indegree.Exists(pName) Then indegree(pName) = 0
            If Not graph(pName).Exists(sName) Then
                graph(pName)(sName) = True
                indegree(sName) = indegree(sName) + 1
            End If
        End If
    Next i

    For Each key In indegree.Keys
        If indegree(key) = 0 Then q.Add CStr(key)
    Next key

    Do While q.Count > 0
        sName = CStr(q(1)) : q.Remove 1
        execOrder.Add sName
        For Each child In graph(sName).Keys
            indegree(child) = indegree(child) - 1
            If indegree(child) = 0 Then q.Add CStr(child)
        Next child
    Loop

    If execOrder.Count <> indegree.Count Then
        MsgBox "❌ Config has circular dependencies!", vbCritical
        Set GetExecutionOrder = Nothing
    Else
        Set GetExecutionOrder = execOrder
    End If
End Function

' ========= FILTER + COPY (ALL visible rows; supports =, <>, >, <, >=, <=, OR via |, contains via ~) =========
Sub FilterAndCopy_Flex(wsSource As Worksheet, wsTarget As Worksheet, _
                       filterRules As String, keepCols As String)

    Dim colDict As Object, rules() As String, rule As Variant
    Dim lastRow As Long, lastCol As Long
    Dim rngData As Range, body As Range, vis As Range
    Dim fieldName As String, op As String, valueExp As String
    Dim critArr() As String
    Dim colArr() As String, pasteCol As Long, c As Long
    Dim srcColIdx As Long, colVis As Range, area As Range
    Dim destRow As Long

    Set colDict = CreateObject("Scripting.Dictionary")

    ' Detect data range
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Or lastCol < 1 Then Exit Sub

    Set rngData = wsSource.Cells(1, 1).Resize(lastRow, lastCol)
    Set body    = rngData.Offset(1).Resize(rngData.Rows.Count - 1)

    ' Header map (case-insensitive)
    For c = 1 To lastCol
        colDict(LCase(Trim(wsSource.Cells(1, c).Value))) = c
    Next c

    ' Prep
    wsTarget.Cells.Clear
    If wsSource.AutoFilterMode Then wsSource.AutoFilterMode = False
    rngData.AutoFilter

    ' Apply rules
    If Len(Trim(filterRules)) > 0 Then
        rules = Split(filterRules, ";")
        For Each rule In rules
            rule = Trim(CStr(rule))
            If Len(rule) = 0 Then GoTo NextRule

            op = ""
            Select Case True
                Case InStr(rule, "<>") > 0: op = "<>"
                Case InStr(rule, ">=") > 0: op = ">="
                Case InStr(rule, "<=") > 0: op = "<="
                Case InStr(rule, ">") > 0:  op = ">"
                Case InStr(rule, "<") > 0:  op = "<"
                Case InStr(rule, "~") > 0:  op = "~"   ' contains (up to 2 terms)
                Case InStr(rule, "=") > 0:  op = "="
            End Select
            If op = "" Then GoTo NextRule

            fieldName = Trim(Split(rule, op)(0))
            valueExp  = Trim(Split(rule, op)(1))
            If Not colDict.Exists(LCase(fieldName)) Then GoTo NextRule

            Select Case op
                Case "="
                    If InStr(valueExp, "|") > 0 Then
                        critArr = Split(valueExp, "|")
                        rngData.AutoFilter Field:=colDict(LCase(fieldName)), _
                                          Criteria1:=critArr, Operator:=xlFilterValues
                    Else
                        rngData.AutoFilter Field:=colDict(LCase(fieldName)), _
                                          Criteria1:=valueExp
                    End If

                Case "<>", ">", "<", ">=", "<="
                    rngData.AutoFilter Field:=colDict(LCase(fieldName)), _
                                      Criteria1:=op & valueExp

                Case "~"
                    critArr = Split(valueExp, "|")
                    If UBound(critArr) = 0 Then
                        rngData.AutoFilter Field:=colDict(LCase(fieldName)), _
                                          Criteria1:="*" & Trim(critArr(0)) & "*"
                    Else
                        ' supports two "contains" terms via OR; for >2 terms consider AdvancedFilter
                        rngData.AutoFilter Field:=colDict(LCase(fieldName)), _
                                          Criteria1:="*" & Trim(critArr(0)) & "*", _
                                          Operator:=xlOr, _
                                          Criteria2:="*" & Trim(critArr(1)) & "*"
                    End If
            End Select
NextRule:
        Next rule
    End If

    ' Visible body (multi-area)
    On Error Resume Next
    Set vis = body.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If vis Is Nothing Then
        wsSource.AutoFilterMode = False
        Exit Sub
    End If

    ' Copy requested columns, appending all visible Areas (no clipboard loss)
    colArr = Split(keepCols, ",")
    pasteCol = 1

    For c = LBound(colArr) To UBound(colArr)
        srcColIdx = 0
        If colDict.Exists(LCase(Trim(colArr(c)))) Then
            srcColIdx = colDict(LCase(Trim(colArr(c))))
        End If
        If srcColIdx = 0 Then GoTo NextKeep

        ' header
        wsTarget.Cells(1, pasteCol).Value = Trim(colArr(c))
        destRow = 2

        ' visible cells for THIS source column
        Set colVis = Application.Intersect(vis, wsSource.Columns(srcColIdx))
        If Not colVis Is Nothing Then
            For Each area In colVis.Areas
                wsTarget.Cells(destRow, pasteCol).Resize(area.Rows.Count, 1).Value = area.Value
                destRow = destRow + area.Rows.Count
            Next area
        End If

        pasteCol = pasteCol + 1
NextKeep:
    Next c

    wsSource.AutoFilterMode = False
End Sub

' ========= RENAME HEADERS =========
Sub ApplyRenameMap(ws As Worksheet, renameMap As String)
    Dim pairs() As String, p As Variant, kv() As String
    Dim lastCol As Long, i As Long
    If renameMap = "" Then Exit Sub

    pairs = Split(renameMap, ",")
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For Each p In pairs
        kv = Split(p, ":")
        If UBound(kv) = 1 Then
            For i = 1 To lastCol
                If LCase(Trim(ws.Cells(1, i).Value)) = LCase(Trim(kv(0))) Then
                    ws.Cells(1, i).Value = Trim(kv(1))
                End If
            Next i
        End If
    Next p
End Sub

' ========= OPTIONS: headers, autofit, freeze, number formats by header =========
Sub ApplyOptions(ws As Worksheet, options As String)
    Dim optArr() As String, kv() As String, i As Long
    Dim opt As Object: Set opt = CreateObject("Scripting.Dictionary")
    If options = "" Then Exit Sub

    optArr = Split(options, ";")
    For i = LBound(optArr) To UBound(optArr)
        If InStr(optArr(i), "=") > 0 Then
            kv = Split(optArr(i), "=")
            opt(LCase(Trim(kv(0)))) = Trim(kv(1))
        ElseIf Len(Trim(optArr(i))) > 0 Then
            opt(LCase(Trim(optArr(i)))) = True
        End If
    Next i

    If opt.Exists("headersbold") Then
        ws.Rows(1).Font.Bold = True
        ws.Rows(1).Interior.Color = RGB(200, 200, 200)
    End If
    If opt.Exists("autofit") Then ws.Cells.EntireColumn.AutoFit
    If opt.Exists("freezetoprow") Then
        ws.Activate: ActiveWindow.FreezePanes = False
        ws.Rows(2).Select: ActiveWindow.FreezePanes = True
    End If

    ' Number formats by header (after renaming)
    ' Syntax: NumFmt=Amount:#,##0.00|UsdEquivalent:#,##0.00|RWA Exposure:0.00%|PaymentDate:yyyy-mm-dd
    If opt.Exists("numfmt") Then
        Dim pairs() As String, p As Variant, colFmt() As String
        Dim colIdx As Long, lastRow As Long
        pairs = Split(opt("numfmt"), "|")
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        For Each p In pairs
            If InStr(p, ":") > 0 Then
                colFmt = Split(p, ":")
                colIdx = FindCol(ws, colFmt(0))
                If colIdx > 0 Then
                    ws.Range(ws.Cells(2, colIdx), ws.Cells(Application.Max(2, lastRow), colIdx)).NumberFormat = colFmt(1)
                End If
            End If
        Next p
    End If
End Sub

' ========= OPTIONAL: RUN SUBSETS (Risk-only / Balance-only) =========
Sub RunRiskPipeline():    RunConfigSubset "Risk Report":            End Sub
Sub RunBalancePipeline(): RunConfigSubset "Balance Break Report":  End Sub

Sub RunConfigSubset(startSheet As String)
    Dim wsConfig As Worksheet, order As Collection
    Dim allowed As Object, i As Long, cfgRow As Long
    Dim sheetName As String, sourceName As String, parentName As String
    Dim filterRules As String, keepCols As String, renameMap As String, options As String
    Dim wsInput As Worksheet, wsTarget As Worksheet

    Set wsConfig = ThisWorkbook.Sheets("Config")
    Set order = GetExecutionOrder(wsConfig)
    If order Is Nothing Then Exit Sub

    Set allowed = CreateObject("Scripting.Dictionary")
    CollectDependents wsConfig, startSheet, allowed

    For i = 1 To order.Count
        sheetName = CStr(order(i))
        If Not allowed.Exists(sheetName) Then GoTo NextItem

        cfgRow = FindConfigRow(wsConfig, sheetName)
        If cfgRow = 0 Then GoTo NextItem

        sourceName = Trim(CStr(wsConfig.Cells(cfgRow, 2).Value))
        parentName = Trim(CStr(wsConfig.Cells(cfgRow, 3).Value))
        filterRules = CStr(wsConfig.Cells(cfgRow, 4).Value)
        keepCols   = CStr(wsConfig.Cells(cfgRow, 5).Value)
        renameMap  = CStr(wsConfig.Cells(cfgRow, 6).Value)
        options    = CStr(wsConfig.Cells(cfgRow, 7).Value)

        Set wsInput = Nothing
        If sourceName <> "" And SheetExists(sourceName) Then
            Set wsInput = ThisWorkbook.Sheets(sourceName)
        ElseIf parentName <> "" And SheetExists(parentName) Then
            Set wsInput = ThisWorkbook.Sheets(parentName)
        End If
        If wsInput Is Nothing Then GoTo NextItem

        Set wsTarget = GetOrCreateSheet(sheetName)
        FilterAndCopy_Flex wsInput, wsTarget, filterRules, keepCols
        ApplyRenameMap wsTarget, renameMap
        ApplyOptions   wsTarget, options
NextItem:
    Next i
End Sub

Sub CollectDependents(wsConfig As Worksheet, root As String, allowed As Object)
    Dim lastRow As Long, i As Long, sName As String, pName As String
    If allowed.Exists(root) Then Exit Sub
    allowed(root) = True

    lastRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        sName = CStr(wsConfig.Cells(i, 1).Value)
        pName = CStr(wsConfig.Cells(i, 3).Value)
        If pName = root Then CollectDependents wsConfig, sName, allowed
    Next i
End Sub

' ========= HELPERS =========
Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    If SheetExists(sheetName) Then
        Set GetOrCreateSheet = ThisWorkbook.Sheets(sheetName)
    Else
        Set GetOrCreateSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Function FindConfigRow(wsConfig As Worksheet, ByVal sheetName As String) As Long
    Dim v As Variant
    On Error Resume Next
    v = Application.Match(sheetName, wsConfig.Columns(1), 0)
    On Error GoTo 0
    If IsError(v) Or Len(v) = 0 Then
        FindConfigRow = 0
    Else
        FindConfigRow = CLng(v)
    End If
End Function

Function FindCol(ws As Worksheet, header As String) As Long
    Dim lastCol As Long, i As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        If LCase(Trim(ws.Cells(1, i).Value)) = LCase(Trim(header)) Then
            FindCol = i
            Exit Function
        End If
    Next i
    FindCol = 0
End Function
