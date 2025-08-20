Option Explicit

' =========================
' CONFIG SHEET COLUMNS
' A: SheetName
' B: Source          (raw sheet name, or blank if dependent)
' C: ParentReport    (upstream sheet name, or blank if base)
' D: FilterRules     (see operator guide)
' E: KeepColumns     (CSV of SOURCE headers, in order)
' F: RenameMap       (CSV of "Original[:Alt1|Alt2]:New"; missing originals ignored; collisions skipped)
' G: Options         (e.g. "FilterUI=All;HeadersBold=True;AutoFit=True;NumFmt=Amount:#,##0.00")
' =========================

' ========= MAIN (runs everything per Config) =========
Sub RunConfigWithDependencies()
    Dim wsConfig As Worksheet, order As Collection
    Dim i As Long, cfgRow As Long
    Dim sheetName As String, sourceName As String, parentName As String
    Dim filterRules As String, keepCols As String, renameMap As String, options As String
    Dim wsInput As Worksheet, wsTarget As Worksheet
    Dim scrn As Boolean, calc As XlCalculation

    ' speed-ups
    scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    calc = Application.Calculation:    Application.Calculation = xlCalculationManual

    Set wsConfig = ThisWorkbook.Sheets("Config")
    Set order = GetExecutionOrder(wsConfig)
    If order Is Nothing Then GoTo Cleanup

    For i = 1 To order.Count
        sheetName = CStr(order(i))
        cfgRow = FindConfigRow(wsConfig, sheetName)
        If cfgRow = 0 Then GoTo NextItem

        sourceName = Trim(CStr(wsConfig.Cells(cfgRow, 2).Value))
        parentName = Trim(CStr(wsConfig.Cells(cfgRow, 3).Value))
        filterRules = CStr(wsConfig.Cells(cfgRow, 4).Value)
        keepCols    = CStr(wsConfig.Cells(cfgRow, 5).Value)
        renameMap   = CStr(wsConfig.Cells(cfgRow, 6).Value)
        options     = CStr(wsConfig.Cells(cfgRow, 7).Value)

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
        FilterAndCopy_Flex wsInput, wsTarget, filterRules, keepCols, options
        ApplyRenameMap     wsTarget, renameMap
        ApplyOptions       wsTarget, options
NextItem:
    Next i

Cleanup:
    Application.ScreenUpdating = scrn
    Application.Calculation = calc
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

' ========= FILTER + COPY (supports UI mode) =========
' Operators (AND across rules; OR within value via "|"):
'   =    equals (CI, OR via "|")
'   <>   not equal (single value, CI)
'   !=   not equal (multi OR, CI). `Status!=` excludes blanks/spaces
'   ~    contains (CI; up to two patterns via "|")
'   !~   does NOT contain (CI; any # terms)
'   >, <, >=, <= numeric/date comparisons
'   =^   equals (case-sensitive, OR via "|")
'   ~^   contains (case-sensitive, any # terms)
'   !=^  not equal (case-sensitive, OR). `Status!=^` excludes blanks/spaces
'   !~^  does NOT contain (case-sensitive, any # terms)
'   ~?   contains (case-insensitive include, unlimited OR, supports `<blank>`)
'
' Options:
'   FilterUI=All  -> paste ALL source columns, then show Excel AutoFilter with rules
'   FilterUI=Keep -> paste only KeepColumns, then show Excel AutoFilter with rules
'
Sub FilterAndCopy_Flex(wsSource As Worksheet, wsTarget As Worksheet, _
                       filterRules As String, keepCols As String, _
                       Optional options As String = "")

    Dim lastRow As Long, lastCol As Long
    Dim rngSrc As Range, bodySrc As Range
    Dim uiMode As String
    Dim opts As Object
    Dim pasteCol As Long

    Set opts = ParseOptions(options)
    uiMode = ""
    If Not opts Is Nothing Then
        If opts.Exists("filterui") Then uiMode = LCase$(Trim$(opts("filterui")))
    End If

    ' --- detect source data ---
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Or lastCol < 1 Then
        wsTarget.Cells.Clear
        Exit Sub
    End If
    Set rngSrc = wsSource.Cells(1, 1).Resize(lastRow, lastCol)
    Set bodySrc = rngSrc.Offset(1).Resize(rngSrc.Rows.Count - 1)

    ' ===== UI mode (copy then filter ON TARGET) =====
    If uiMode = "all" Or uiMode = "keep" Then
        Dim colDictSrc As Object: Set colDictSrc = HeaderDict(wsSource, lastCol)
        Dim colDictTgt As Object
        Dim keepArr() As String, k As Long, srcIdx As Long
        Dim lastColT As Long, lastRowT As Long
        Dim rngTgt As Range

        wsTarget.Cells.Clear

        If uiMode = "all" Then
            wsTarget.Cells(1, 1).Resize(lastRow, lastCol).Value = rngSrc.Value
        Else
            keepArr = Split(keepCols, ",")
            pasteCol = 1
            For k = LBound(keepArr) To UBound(keepArr)
                srcIdx = 0
                If colDictSrc.Exists(LCase$(Trim$(keepArr(k)))) Then
                    srcIdx = colDictSrc(LCase$(Trim$(keepArr(k))))
                End If
                If srcIdx > 0 Then
                    wsTarget.Cells(1, pasteCol).Value = Trim$(keepArr(k))
                    wsTarget.Cells(2, pasteCol).Resize(lastRow - 1, 1).Value = _
                        wsSource.Cells(2, srcIdx).Resize(lastRow - 1, 1).Value
                    pasteCol = pasteCol + 1
                End If
            Next k
        End If

        lastRowT = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
        lastColT = wsTarget.Cells(1, wsTarget.Columns.Count).End(xlToLeft).Column
        If lastRowT < 2 Then Exit Sub
        Set rngTgt = wsTarget.Cells(1, 1).Resize(lastRowT, lastColT)

        If wsTarget.AutoFilterMode Then wsTarget.AutoFilterMode = False
        rngTgt.AutoFilter

        Set colDictTgt = HeaderDict(wsTarget, lastColT)
        ApplyRules_OnTarget rngTgt, colDictTgt, filterRules
        Exit Sub
    End If

    ' ===== non-UI mode (filter source + copy visible subset) =====
    Dim colDictSrc2 As Object: Set colDictSrc2 = HeaderDict(wsSource, lastCol)
    Dim rngVisible As Range, c As Long
    Dim rules() As String, rule As Variant
    Dim fieldName As String, op As String, valueExp As String
    Dim critArr() As String
    Dim colArr() As String, srcColIdx As Long
    Dim area As Range, cell As Range, destRow As Long

    ' collections for row-level checks
    Dim exBlank As Object, exEqCI As Object, exContCI As Object
    Dim exEqCS As Object, exContCS As Object, incEqCS As Object, incContCS As Object, incContCI As Object

    Set exBlank   = CreateObject("Scripting.Dictionary")
    Set exEqCI    = CreateObject("Scripting.Dictionary")
    Set exContCI  = CreateObject("Scripting.Dictionary")
    Set exEqCS    = CreateObject("Scripting.Dictionary")
    Set exContCS  = CreateObject("Scripting.Dictionary")
    Set incEqCS   = CreateObject("Scripting.Dictionary")
    Set incContCS = CreateObject("Scripting.Dictionary")
    Set incContCI = CreateObject("Scripting.Dictionary")

    wsTarget.Cells.Clear
    If wsSource.AutoFilterMode Then wsSource.AutoFilterMode = False
    rngSrc.AutoFilter

    If Len(Trim$(filterRules)) > 0 Then
        rules = Split(filterRules, ";")
        For Each rule In rules
            rule = Trim$(CStr(rule))
            If Len(rule) = 0 Then GoTo NextRule1

            op = DetectOp(rule)
            If op = "" Then GoTo NextRule1

            fieldName = Trim$(Split(rule, op)(0))
            valueExp  = Trim$(Split(rule, op)(1))
            If Not colDictSrc2.Exists(LCase$(fieldName)) Then GoTo NextRule1

            Dim fld As Long: fld = colDictSrc2(LCase$(fieldName))

            Select Case op
                Case "=^"
                    SetDictAdd incEqCS, fld, valueExp, True

                Case "~^"
                    SetDictAdd incContCS, fld, valueExp, True

                Case "~?"
                    SetDictAddCI_WithBlank incContCI, fld, valueExp

                Case "!=^"
                    If valueExp = "" Then
                        exBlank(fld) = True
                    Else
                        SetDictAdd exEqCS, fld, valueExp, True
                    End If

                Case "!~^"
                    SetDictAdd exContCS, fld, valueExp, True

                Case "!="
                    If valueExp = "" Then
                        exBlank(fld) = True
                    ElseIf InStr(valueExp, "|") > 0 Then
                        SetDictAddCI exEqCI, fld, valueExp
                    Else
                        rngSrc.AutoFilter Field:=fld, Criteria1:="<>" & valueExp
                    End If

                Case "!~"
                    SetDictAddCI exContCI, fld, valueExp

                Case "="
                    If InStr(valueExp, "|") > 0 Then
                        critArr = Split(valueExp, "|")
                        rngSrc.AutoFilter Field:=fld, Criteria1:=critArr, Operator:=xlFilterValues
                    Else
                        rngSrc.AutoFilter Field:=fld, Criteria1:=valueExp
                    End If

                Case "<>", ">", "<", ">=", "<="
                    rngSrc.AutoFilter Field:=fld, Criteria1:=op & valueExp

                Case "~"
                    critArr = Split(valueExp, "|")
                    If UBound(critArr) = 0 Then
                        rngSrc.AutoFilter Field:=fld, Criteria1:="*" & Trim$(critArr(0)) & "*"
                    Else
                        rngSrc.AutoFilter Field:=fld, _
                            Criteria1:="*" & Trim$(critArr(0)) & "*", _
                            Operator:=xlOr, _
                            Criteria2:="*" & Trim$(critArr(1)) & "*"
                    End If
            End Select
NextRule1:
        Next rule
    End If

    On Error Resume Next
    Set rngVisible = bodySrc.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If rngVisible Is Nothing Then
        wsSource.AutoFilterMode = False
        Exit Sub
    End If

    ' copy requested columns from visible rows
    colArr = Split(keepCols, ",")
    pasteCol = 1

    For c = LBound(colArr) To UBound(colArr)
        srcColIdx = 0
        If colDictSrc2.Exists(LCase$(Trim$(colArr(c)))) Then
            srcColIdx = colDictSrc2(LCase$(Trim$(colArr(c))))
        End If
        If srcColIdx = 0 Then GoTo NextKeep

        wsTarget.Cells(1, pasteCol).Value = Trim$(colArr(c))
        destRow = 2

        Dim colVis As Range
        Set colVis = Application.Intersect(rngVisible, wsSource.Columns(srcColIdx))
        If Not colVis Is Nothing Then
            For Each area In colVis.Areas
                For Each cell In area.Cells
                    If RowPassesRules(wsSource, cell.Row, _
                                      exBlank, exEqCI, exContCI, _
                                      exEqCS, exContCS, _
                                      incEqCS, incContCS, _
                                      incContCI) Then
                        wsTarget.Cells(destRow, pasteCol).Value = cell.Value
                        destRow = destRow + 1
                    End If
                Next cell
            Next area
        End If
        pasteCol = pasteCol + 1
NextKeep:
    Next c

    wsSource.AutoFilterMode = False
End Sub



' ======== Apply rules on TARGET (UI mode): use native AutoFilter + helper column for residuals ========
Private Sub ApplyRules_OnTarget(rngTgt As Range, colDict As Object, filterRules As String)
    Dim rules() As String, rule As Variant
    Dim fieldName As String, op As String, valueExp As String, fld As Long
    Dim critArr() As String, terms() As String
    Dim helperConds As Collection: Set helperConds = New Collection
    Dim colRef As String, cond As String
    Dim lastRow As Long, lastCol As Long, helperCol As Long, f As String

    If Trim$(filterRules) = "" Then Exit Sub

    lastRow = rngTgt.Rows.Count
    lastCol = rngTgt.Columns.Count

    rules = Split(filterRules, ";")
    For Each rule In rules
        rule = Trim$(CStr(rule))
        If Len(rule) = 0 Then GoTo NextRule2

        op = DetectOp(rule)
        If op = "" Then GoTo NextRule2

        fieldName = Trim$(Split(rule, op)(0))
        valueExp  = Trim$(Split(rule, op)(1))
        If Not colDict.Exists(LCase$(fieldName)) Then GoTo NextRule2

        fld = CLng(colDict(LCase$(fieldName)))
        If fld < 1 Or fld > lastCol Then GoTo NextRule2

        colRef = "$" & ColLetter(fld) & "2"  ' absolute column, relative row

        Select Case op
            ' --- UI-capable directly ---
            Case "="
                If InStr(valueExp, "|") > 0 Then
                    critArr = Split(valueExp, "|")
                    rngTgt.AutoFilter Field:=fld, Criteria1:=critArr, Operator:=xlFilterValues
                Else
                    rngTgt.AutoFilter Field:=fld, Criteria1:=valueExp
                End If

            Case "~"
                critArr = Split(valueExp, "|")
                If UBound(critArr) = 0 Then
                    rngTgt.AutoFilter Field:=fld, Criteria1:="*" & Trim$(critArr(0)) & "*"
                ElseIf UBound(critArr) = 1 Then
                    rngTgt.AutoFilter Field:=fld, _
                        Criteria1:="*" & Trim$(critArr(0)) & "*", _
                        Operator:=xlOr, _
                        Criteria2:="*" & Trim$(critArr(1)) & "*"
                Else
                    ' too many contains terms for UI → helper
                    helperConds.Add BuildContainsCI_OR(colRef, critArr)
                End If

            Case "<>", ">", "<", ">=", "<="
                ' try native UI; if it fails, build helper formula
                If Not TryAutoFilterCompare(rngTgt, fld, op, valueExp) Then
                    helperConds.Add BuildCompareFormula(colRef, op, valueExp)
                End If

            ' --- helper-required (negative / case-sensitive / special) ---
            Case "!="
                If valueExp = "" Then
                    helperConds.Add "LEN(TRIM(" & colRef & "))>0"
                ElseIf InStr(valueExp, "|") > 0 Then
                    terms = Split(valueExp, "|")
                    helperConds.Add BuildNotEqualsCI_AND(colRef, terms)
                Else
                    ' single value: UI usually ok
                    On Error Resume Next
                    rngTgt.AutoFilter Field:=fld, Criteria1:="<>" & valueExp
                    If Err.Number <> 0 Then
                        Err.Clear
                        helperConds.Add BuildNotEqualsCI_AND(colRef, Split(valueExp, "|"))
                    End If
                    On Error GoTo 0
                End If

            Case "!~"
                terms = Split(valueExp, "|")
                helperConds.Add BuildNotContainsCI_AND(colRef, terms)

            Case "=^"
                terms = SplitOrSingle(valueExp)
                helperConds.Add BuildEqualsCS_OR(colRef, terms)

            Case "~^"
                terms = SplitOrSingle(valueExp)
                helperConds.Add BuildContainsCS_OR(colRef, terms)

            Case "!=^"
                If valueExp = "" Then
                    helperConds.Add "LEN(TRIM(" & colRef & "))>0"
                Else
                    terms = SplitOrSingle(valueExp)
                    helperConds.Add BuildNotEqualsCS_AND(colRef, terms)
                End If

            Case "!~^"
                terms = SplitOrSingle(valueExp)
                helperConds.Add BuildNotContainsCS_AND(colRef, terms)

            Case "~?"
                terms = Split(valueExp, "|")
                helperConds.Add BuildContainsCI_OR_WithBlank(colRef, terms)
        End Select

NextRule2:
    Next rule

    ' Add helper if needed
    If helperConds.Count > 0 Then
        Dim helperCol As Long: helperCol = lastCol + 1
        rngTgt.Worksheet.Cells(1, helperCol).Value = "_FilterPass"
        f = "=AND(" & JoinCollection(helperConds, ",") & ")"
        rngTgt.Worksheet.Cells(2, helperCol).Formula = f
        rngTgt.Worksheet.Range(rngTgt.Worksheet.Cells(2, helperCol), _
                               rngTgt.Worksheet.Cells(lastRow, helperCol)).FillDown
        rngTgt.Resize(, helperCol).AutoFilter Field:=helperCol, Criteria1:="=TRUE"
        rngTgt.Worksheet.Columns(helperCol).Hidden = True
    End If
End Sub

' ========= RENAME HEADERS (safe: ignore missing originals; allow multi-origin; avoid collisions) =========
Sub ApplyRenameMap(ws As Worksheet, renameMap As String)
    Dim mapArr() As String, pair As Variant
    Dim leftPart As String, newName As String
    Dim candidates() As String, cand As Variant
    Dim lastCol As Long, i As Long
    Dim hdrDict As Object, colIdx As Variant, foundIdx As Variant

    If Trim$(renameMap) = "" Then Exit Sub

    Set hdrDict = CreateObject("Scripting.Dictionary")
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        hdrDict(LCase$(Trim$(ws.Cells(1, i).Value))) = i
    Next i

    mapArr = Split(renameMap, ",")
    For Each pair In mapArr
        pair = Trim$(CStr(pair))
        If pair <> "" And InStr(pair, ":") > 0 Then
            leftPart = Trim$(Split(pair, ":", 2)(0))  ' may be "Old|AlreadyRenamed"
            newName  = Trim$(Split(pair, ":", 2)(1))
            If newName = "" Then GoTo NextPair

            foundIdx = Empty
            candidates = Split(leftPart, "|")

            For Each cand In candidates
                cand = Trim$(CStr(cand))
                If cand <> "" Then
                    If hdrDict.Exists(LCase$(cand)) Then
                        foundIdx = hdrDict(LCase$(cand))
                        Exit For
                    End If
                End If
            Next cand

            If Not IsEmpty(foundIdx) Then
                colIdx = CLng(foundIdx)
                If hdrDict.Exists(LCase$(newName)) And hdrDict(LCase$(newName)) <> colIdx Then
                    ' collision: skip
                Else
                    ws.Cells(1, colIdx).Value = newName
                    ' update dict
                    For Each cand In candidates
                        cand = Trim$(CStr(cand))
                        If cand <> "" Then
                            If hdrDict.Exists(LCase$(cand)) And hdrDict(LCase$(cand)) = colIdx Then
                                hdrDict.Remove LCase$(cand)
                            End If
                        End If
                    Next cand
                    hdrDict(LCase$(newName)) = colIdx
                End If
            Else
                ' none matched -> ignore
            End If
        End If
NextPair:
    Next pair
End Sub

' ========= OPTIONS: headers, autofit, freeze, number formats, CommaStyle =========
Sub ApplyOptions(ws As Worksheet, options As String)
    Dim opt As Object: Set opt = ParseOptions(options)
    If opt Is Nothing Then Exit Sub

    If opt.Exists("headersbold") Then
        ws.Rows(1).Font.Bold = True
        ws.Rows(1).Interior.Color = RGB(200, 200, 200)
    End If
    If opt.Exists("autofit") Then ws.Cells.EntireColumn.AutoFit
    If opt.Exists("freezetoprow") Then
        ws.Activate: ActiveWindow.FreezePanes = False
        ws.Rows(2).Select: ActiveWindow.FreezePanes = True
    End If

    ' NumFmt=Col:Format|Col2:Format2
    If opt.Exists("numfmt") Then
        Dim pairs() As String, p As Variant, colFmt() As String
        Dim colIdx As Long, lastRow As Long
        pairs = Split(CStr(opt("numfmt")), "|")
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        For Each p In pairs
            If InStr(p, ":") > 0 Then
                colFmt = Split(p, ":", 2)
                colIdx = FindCol(ws, Trim$(colFmt(0)))
                If colIdx > 0 Then
                    ws.Range(ws.Cells(2, colIdx), ws.Cells(Application.Max(2, lastRow), colIdx)).NumberFormat = Trim$(colFmt(1))
                End If
            End If
        Next p
    End If

    ' CommaStyle=HeaderA|HeaderB (2 decimals)  /  CommaStyle0=HeaderC|HeaderD (0 decimals)
    If opt.Exists("commastyle") Then ApplyCommaStyleToHeaders ws, CStr(opt("commastyle")), False
    If opt.Exists("commastyle0") Then ApplyCommaStyleToHeaders ws, CStr(opt("commastyle0")), True
End Sub

' ========= OPTIONAL: RUN SUBSETS =========
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
        keepCols    = CStr(wsConfig.Cells(cfgRow, 5).Value)
        renameMap   = CStr(wsConfig.Cells(cfgRow, 6).Value)
        options     = CStr(wsConfig.Cells(cfgRow, 7).Value)

        Set wsInput = Nothing
        If sourceName <> "" And SheetExists(sourceName) Then
            Set wsInput = ThisWorkbook.Sheets(sourceName)
        ElseIf parentName <> "" And SheetExists(parentName) Then
            Set wsInput = ThisWorkbook.Sheets(parentName)
        End If
        If wsInput Is Nothing Then GoTo NextItem

        Set wsTarget = GetOrCreateSheet(sheetName)
        FilterAndCopy_Flex wsInput, wsTarget, filterRules, keepCols, options
        ApplyRenameMap     wsTarget, renameMap
        ApplyOptions       wsTarget, options
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

Private Function HeaderDict(ws As Worksheet, lastCol As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To lastCol
        d(LCase$(Trim$(ws.Cells(1, c).Value))) = c
    Next c
    Set HeaderDict = d
End Function

Private Function ParseOptions(options As String) As Object
    Dim d As Object, tokens() As String, i As Long, kv() As String, t As String
    If Trim$(options) = "" Then Exit Function
    Set d = CreateObject("Scripting.Dictionary")
    tokens = Split(options, ";")
    For i = LBound(tokens) To UBound(tokens)
        t = Trim$(tokens(i))
        If t <> "" Then
            If InStr(t, "=") > 0 Then
                kv = Split(t, "=", 2)
                d(LCase$(Trim$(kv(0)))) = Trim$(kv(1))
            Else
                d(LCase$(t)) = True
            End If
        End If
    Next i
    Set ParseOptions = d
End Function

Private Function DetectOp(ByVal rule As String) As String
    Select Case True
        Case InStr(rule, "!=^") > 0: DetectOp = "!=^"
        Case InStr(rule, "!~^") > 0: DetectOp = "!~^"
        Case InStr(rule, "=^")  > 0: DetectOp = "=^"
        Case InStr(rule, "~^")  > 0: DetectOp = "~^"
        Case InStr(rule, "~?")  > 0: DetectOp = "~?"
        Case InStr(rule, "!=")  > 0: DetectOp = "!="
        Case InStr(rule, "!~")  > 0: DetectOp = "!~"
        Case InStr(rule, "<>")  > 0: DetectOp = "<>"
        Case InStr(rule, ">=")  > 0: DetectOp = ">="
        Case InStr(rule, "<=")  > 0: DetectOp = "<="
        Case InStr(rule, ">")   > 0: DetectOp = ">"
        Case InStr(rule, "<")   > 0: DetectOp = "<"
        Case InStr(rule, "~")   > 0: DetectOp = "~"
        Case InStr(rule, "=")   > 0: DetectOp = "="
    End Select
End Function

Private Function ColLetter(col As Long) As String
    Dim s As String
    Do
        s = Chr$(((col - 1) Mod 26) + 65) & s
        col = (col - 1) \ 26
    Loop While col > 0
    ColLetter = s
End Function

Private Function EscQ(ByVal s As String) As String
    EscQ = """" & Replace(s, """", """""") & """"
End Function

Private Function JoinCollection(col As Collection, ByVal sep As String) As String
    Dim i As Long, tmp As String
    For i = 1 To col.Count
        tmp = tmp & IIf(i > 1, sep, "") & col(i)
    Next i
    JoinCollection = tmp
End Function

' ----- helper: add dict items -----
Private Sub SetDictAdd(target As Object, fld As Long, valueExp As String, Optional caseSensitive As Boolean = False)
    Dim d As Object, v As Variant
    If Not target.Exists(fld) Then Set target(fld) = CreateObject("Scripting.Dictionary")
    Set d = target(fld)
    If InStr(valueExp, "|") > 0 Then
        For Each v In Split(valueExp, "|")
            If caseSensitive Then
                d(Trim$(CStr(v))) = True
            Else
                d(LCase$(Trim$(CStr(v)))) = True
            End If
        Next v
    Else
        If caseSensitive Then
            d(Trim$(valueExp)) = True
        Else
            d(LCase$(Trim$(valueExp))) = True
        End If
    End If
End Sub

Private Sub SetDictAddCI(target As Object, fld As Long, valueExp As String)
    SetDictAdd target, fld, valueExp, False
End Sub

Private Sub SetDictAddCI_WithBlank(target As Object, fld As Long, valueExp As String)
    Dim d As Object, p As Variant, t As String
    If Not target.Exists(fld) Then Set target(fld) = CreateObject("Scripting.Dictionary")
    Set d = target(fld)
    For Each p In Split(valueExp, "|")
        t = LCase$(Trim$(CStr(p)))
        If t = "<blank>" Then
            d("__BLANK__") = True
        ElseIf t <> "" Then
            d(t) = True
        End If
    Next p
End Sub

' ----- build formula snippets for helper column -----
Private Function SplitOrSingle(s As String) As Variant
    If InStr(s, "|") > 0 Then
        SplitOrSingle = Split(s, "|")
    Else
        Dim a(0 To 0) As String
        a(0) = s
        SplitOrSingle = a
    End If
End Function

Private Function BuildNotEqualsCI_AND(colRef As String, terms() As String) As String
    Dim i As Long, bits() As String
    ReDim bits(LBound(terms) To UBound(terms))
    For i = LBound(terms) To UBound(terms)
        bits(i) = "LOWER(TRIM(" & colRef & "))<>" & EscQ(LCase$(Trim$(terms(i))))
    Next i
    BuildNotEqualsCI_AND = "AND(" & Join(bits, ",") & ")"
End Function

Private Function BuildNotContainsCI_AND(colRef As String, terms() As String) As String
    Dim i As Long, bits() As String
    ReDim bits(LBound(terms) To UBound(terms))
    For i = LBound(terms) To UBound(terms)
        bits(i) = "ISERROR(SEARCH(" & EscQ(Trim$(terms(i))) & "," & colRef & "))"
    Next i
    BuildNotContainsCI_AND = "AND(" & Join(bits, ",") & ")"
End Function

Private Function BuildEqualsCS_OR(colRef As String, terms As Variant) As String
    Dim i As Long, bits() As String
    ReDim bits(LBound(terms) To UBound(terms))
    For i = LBound(terms) To UBound(terms)
        bits(i) = "EXACT(" & colRef & "," & EscQ(Trim$(terms(i))) & ")"
    Next i
    BuildEqualsCS_OR = "OR(" & Join(bits, ",") & ")"
End Function

Private Function BuildNotEqualsCS_AND(colRef As String, terms As Variant) As String
    Dim i As Long, bits() As String
    ReDim bits(LBound(terms) To UBound(terms))
    For i = LBound(terms) To UBound(terms)
        bits(i) = "NOT(EXACT(" & colRef & "," & EscQ(Trim$(terms(i))) & "))"
    Next i
    BuildNotEqualsCS_AND = "AND(" & Join(bits, ",") & ")"
End Function

Private Function BuildContainsCS_OR(colRef As String, terms As Variant) As String
    Dim i As Long, bits() As String
    ReDim bits(LBound(terms) To UBound(terms))
    For i = LBound(terms) To UBound(terms)
        bits(i) = "ISNUMBER(FIND(" & EscQ(Trim$(terms(i))) & "," & colRef & "))"
    Next i
    BuildContainsCS_OR = "OR(" & Join(bits, ",") & ")"
End Function

Private Function BuildNotContainsCS_AND(colRef As String, terms As Variant) As String
    Dim i As Long, bits() As String
    ReDim bits(LBound(terms) To UBound(terms))
    For i = LBound(terms) To UBound(terms)
        bits(i) = "ISERROR(FIND(" & EscQ(Trim$(terms(i))) & "," & colRef & "))"
    Next i
    BuildNotContainsCS_AND = "AND(" & Join(bits, ",") & ")"
End Function

Private Function BuildContainsCI_OR(colRef As String, terms() As String) As String
    Dim i As Long, bits() As String
    ReDim bits(LBound(terms) To UBound(terms))
    For i = LBound(terms) To UBound(terms)
        bits(i) = "ISNUMBER(SEARCH(" & EscQ(Trim$(terms(i))) & "," & colRef & "))"
    Next i
    BuildContainsCI_OR = "OR(" & Join(bits, ",") & ")"
End Function

Private Function BuildContainsCI_OR_WithBlank(colRef As String, terms() As String) As String
    Dim i As Long, bits() As String, n As Long
    n = -1
    ' optionally include blank
    Dim includeBlank As Boolean: includeBlank = False
    Dim t As String
    For i = LBound(terms) To UBound(terms)
        t = LCase$(Trim$(terms(i)))
        If t = "<blank>" Then includeBlank = True Else n = n + 1
    Next i

    ReDim bits(0 To IIf(includeBlank, UBound(terms), UBound(terms)) - IIf(includeBlank, 1, 0))

    Dim idx As Long: idx = 0
    If includeBlank Then
        bits(idx) = "LEN(TRIM(" & colRef & "))=0": idx = idx + 1
    End If
    For i = LBound(terms) To UBound(terms)
        t = Trim$(terms(i))
        If LCase$(t) <> "<blank>" Then
            bits(idx) = "ISNUMBER(SEARCH(" & EscQ(t) & "," & colRef & "))"
            idx = idx + 1
        End If
    Next i

    BuildContainsCI_OR_WithBlank = "OR(" & Join(bits, ",") & ")"
End Function

' ====== formatting helper ======
Private Sub ApplyCommaStyleToHeaders(ws As Worksheet, headersList As String, zeroDecimals As Boolean)
    Dim arr() As String, h As Variant
    Dim colIdx As Long, lastRow As Long
    Dim rng As Range, styleName As String

    If Trim$(headersList) = "" Then Exit Sub
    arr = Split(headersList, "|")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For Each h In arr
        colIdx = FindCol(ws, Trim$(CStr(h)))
        If colIdx > 0 Then
            Set rng = ws.Range(ws.Cells(2, colIdx), ws.Cells(Application.Max(2, lastRow), colIdx))
            styleName = IIf(zeroDecimals, "Comma [0]", "Comma")
            On Error Resume Next
            rng.Style = styleName
            If Err.Number <> 0 Then
                Err.Clear
                If zeroDecimals Then
                    rng.NumberFormat = "#,##0"
                Else
                    rng.NumberFormat = "#,##0.00"
                End If
            End If
            On Error GoTo 0
        End If
    Next h
End Sub

' ========= ROW PASS LOGIC (non-UI branch) =========
Private Function RowPassesRules(ws As Worksheet, r As Long, _
                                exBlank As Object, exEqCI As Object, exContCI As Object, _
                                exEqCS As Object, exContCS As Object, _
                                incEqCS As Object, incContCS As Object, _
                                incContCI As Object) As Boolean
    Dim k As Variant, v As String, lv As String, pat As Variant
    Dim hit As Boolean

    ' Exclude blanks/whitespace
    For Each k In exBlank.Keys
        v = CStr(ws.Cells(r, CLng(k)).Value)
        If Trim$(v) = "" Then RowPassesRules = False: Exit Function
    Next k

    ' Exclude equals (case-insensitive)
    For Each k In exEqCI.Keys
        v = Trim$(CStr(ws.Cells(r, CLng(k)).Value))
        If exEqCI(k).Exists(LCase$(v)) Then RowPassesRules = False: Exit Function
    Next k

    ' Exclude contains (case-insensitive)
    For Each k In exContCI.Keys
        v = Trim$(CStr(ws.Cells(r, CLng(k)).Value))
        lv = LCase$(v)
        For Each pat In exContCI(k).Keys
            If pat <> "" Then
                If InStr(1, lv, CStr(pat), vbTextCompare) > 0 Then
                    RowPassesRules = False: Exit Function
                End If
            End If
        Next pat
    Next k

    ' Exclude equals (case-sensitive)
    For Each k In exEqCS.Keys
        v = Trim$(CStr(ws.Cells(r, CLng(k)).Value))
        For Each pat In exEqCS(k).Keys
            If StrComp(v, CStr(pat), vbBinaryCompare) = 0 Then
                RowPassesRules = False: Exit Function
            End If
        Next pat
    Next k

    ' Exclude contains (case-sensitive)
    For Each k In exContCS.Keys
        v = Trim$(CStr(ws.Cells(r, CLng(k)).Value))
        For Each pat In exContCS(k).Keys
            If pat <> "" Then
                If InStr(1, v, CStr(pat), vbBinaryCompare) > 0 Then
                    RowPassesRules = False: Exit Function
                End If
            End If
        Next pat
    Next k

    ' Include equals (case-sensitive)
    For Each k In incEqCS.Keys
        v = Trim$(CStr(ws.Cells(r, CLng(k)).Value))
        hit = False
        For Each pat In incEqCS(k).Keys
            If StrComp(v, CStr(pat), vbBinaryCompare) = 0 Then
                hit = True: Exit For
            End If
        Next pat
        If Not hit Then RowPassesRules = False: Exit Function
    Next k

    ' Include contains (case-sensitive)
    For Each k In incContCS.Keys
        v = Trim$(CStr(ws.Cells(r, CLng(k)).Value))
        hit = False
        For Each pat In incContCS(k).Keys
            If pat <> "" Then
                If InStr(1, v, CStr(pat), vbBinaryCompare) > 0 Then
                    hit = True: Exit For
                End If
            End If
        Next pat
        If Not hit Then RowPassesRules = False: Exit Function
    Next k

    ' Include contains (case-insensitive ~?)
    For Each k In incContCI.Keys
        v = Trim$(CStr(ws.Cells(r, CLng(k)).Value))
        lv = LCase$(v)
        hit = False
        If incContCI(k).Exists("__BLANK__") And Trim$(v) = "" Then
            hit = True
        Else
            For Each pat In incContCI(k).Keys
                If pat <> "__BLANK__" Then
                    If InStr(1, lv, CStr(pat), vbTextCompare) > 0 Then
                        hit = True: Exit For
                    End If
                End If
            Next pat
        End If
        If Not hit Then RowPassesRules = False: Exit Function
    Next k

    RowPassesRules = True
End Function
                                                                                                                                                                                                    
Private Function TryAutoFilterCompare(rngTgt As Range, fld As Long, op As String, valueExp As String) As Boolean
    On Error GoTo FailFast

    ' Exclude blanks: "<>" (no value) is valid
    If op = "<>" And Len(valueExp) = 0 Then
        rngTgt.AutoFilter Field:=fld, Criteria1:="<>"
        TryAutoFilterCompare = True
        Exit Function
    End If

    ' Try numeric
    If IsNumeric(valueExp) Then
        rngTgt.AutoFilter Field:=fld, Criteria1:=op & CDbl(valueExp)
        TryAutoFilterCompare = True
        Exit Function
    End If

    ' Try date
    If IsDate(valueExp) Then
        rngTgt.AutoFilter Field:=fld, Criteria1:=op & CLng(CDate(valueExp))
        TryAutoFilterCompare = True
        Exit Function
    End If

    ' Fallback as text
    rngTgt.AutoFilter Field:=fld, Criteria1:=op & valueExp
    TryAutoFilterCompare = True
    Exit Function

FailFast:
    TryAutoFilterCompare = False
End Function

Private Function BuildCompareFormula(colRef As String, op As String, valueExp As String) As String
    Dim d As String
    If op = "<>" Then
        If Len(valueExp) = 0 Then
            BuildCompareFormula = "LEN(TRIM(" & colRef & "))>0"
        Else
            BuildCompareFormula = "LOWER(TRIM(" & colRef & "))<>" & EscQ(LCase$(valueExp))
        End If
        Exit Function
    End If

    If IsDate(valueExp) Then
        d = Format$(CDate(valueExp), "yyyy-mm-dd")
        BuildCompareFormula = "IFERROR(DATEVALUE(TRIM(" & colRef & ")),N(" & colRef & "))" & _
                              op & "DATEVALUE(" & EscQ(d) & ")"
    ElseIf IsNumeric(valueExp) Then
        BuildCompareFormula = "VALUE(SUBSTITUTE(TRIM(" & colRef & "),"","""",""""))" & op & CStr(CDbl(valueExp))
    Else
        ' textual comparison (case-insensitive)
        BuildCompareFormula = "LOWER(TRIM(" & colRef & "))" & op & "LOWER(" & EscQ(valueExp) & ")"
    End If
End Function
                                                                                                                                                                            
