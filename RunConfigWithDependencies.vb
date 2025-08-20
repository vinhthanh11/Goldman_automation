Option Explicit

' =========================
' CONFIG SHEET COLUMNS
' A: SheetName
' B: Source          (raw sheet name, or blank if dependent)
' C: ParentReport    (upstream sheet name, or blank if base)
' D: FilterRules     (see operator guide below)
' E: KeepColumns     (CSV of headers to copy, in order; use SOURCE header names)
' F: RenameMap       (CSV of "Original:New" pairs; supports multi-origin "A|B:New")
' G: Options         (e.g. "HeadersBold=True;AutoFit=True;NumFmt=Amount:#,##0.00|RWA Exposure:0.00%;FlagOnly=True;FlagColumn=Include?")
' =========================

' ========= MAIN (runs everything per Config) =========
Sub RunConfigWithDependencies()
    Dim wsConfig As Worksheet, order As Collection
    Dim i As Long, cfgRow As Long
    Dim sheetName As String, sourceName As String, parentName As String
    Dim filterRules As String, keepCols As String, renameMap As String, options As String
    Dim wsInput As Worksheet, wsTarget As Worksheet
    Dim scrn As Boolean, calc As XlCalculation
    Dim opt As Object, isFlagOnly As Boolean

    ' optional speed-ups
    scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    calc = Application.Calculation:    Application.Calculation = xlCalculationManual

    Set wsConfig = ThisWorkbook.Sheets("Config")
    Set order = GetExecutionOrder(wsConfig)
    If order Is Nothing Then GoTo Cleanup

    For i = 1 To order.Count
        sheetName = CStr(order(i))
        cfgRow = FindConfigRow(wsConfig, sheetName)
        If cfgRow = 0 Then GoTo NextItem

        sourceName = Trim(CStr(wsConfig.Cells(cfgRow, 2).Value))   ' Source
        parentName = Trim(CStr(wsConfig.Cells(cfgRow, 3).Value))   ' ParentReport
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

        ' Parse options to check for FlagOnly
        Set opt = ParseOptions(options)
        isFlagOnly = False
        If Not opt Is Nothing Then
            If opt.Exists("flagonly") Then isFlagOnly = CBool(ValueAsBool(opt("flagonly")))
        End If

        ' Process
        FilterAndCopy_Flex wsInput, wsTarget, filterRules, keepCols, options

        ' In FlagOnly mode we only wrote a flag to the source sheet; skip target formatting
        If Not isFlagOnly Then
            ApplyRenameMap     wsTarget, renameMap
            ApplyOptions       wsTarget, options
        End If
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

' ========= FILTER + COPY (ALL visible rows; supports advanced operators and FlagOnly) =========
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
' Example (include blanks OR “Not sent” OR “SLC” in LatestComment, CI):
'   Setts_Input Field_LatestComment~?Not sent|SLC|<blank>
Sub FilterAndCopy_Flex(wsSource As Worksheet, wsTarget As Worksheet, _
                       filterRules As String, keepCols As String, _
                       Optional options As String = "")

    Dim colDict As Object, rules() As String, rule As Variant
    Dim lastRow As Long, lastCol As Long
    Dim rngData As Range, body As Range, vis As Range
    Dim fieldName As String, op As String, valueExp As String
    Dim critArr() As String
    Dim colArr() As String, pasteCol As Long, c As Long
    Dim srcColIdx As Long, colVis As Range, area As Range
    Dim destRow As Long

    ' --- collections for row-level checks ---
    Dim exBlank As Object                         ' colIdx -> True (exclude trimmed blanks)
    Dim exEqCI As Object, exContCI As Object      ' (case-insensitive) excludes
    Dim exEqCS As Object, exContCS As Object      ' (case-sensitive)   excludes
    Dim incEqCS As Object, incContCS As Object    ' (case-sensitive)   includes
    Dim incContCI As Object                       ' (case-insensitive) includes for ~?
    Dim dictTmp As Object
    Dim v As Variant, p As Variant

    ' --- options (FlagOnly) ---
    Dim opt As Object
    Dim flagOnly As Boolean, flagColName As String, flagYes As String, flagNo As String

    Set opt = ParseOptions(options)
    flagOnly = False: flagColName = "Include?": flagYes = "Yes": flagNo = "No"
    If Not opt Is Nothing Then
        If opt.Exists("flagonly") Then flagOnly = CBool(ValueAsBool(opt("flagonly")))
        If opt.Exists("flagcolumn") Then If Trim$(opt("flagcolumn")) <> "" Then flagColName = Trim$(opt("flagcolumn"))
        If opt.Exists("flagyes") Then If Trim$(opt("flagyes")) <> "" Then flagYes = Trim$(opt("flagyes"))
        If opt.Exists("flagn o") Then ' typo-proofing, just in case
            If Trim$(opt("flagn o")) <> "" Then flagNo = Trim$(opt("flagn o"))
        ElseIf opt.Exists("flagno") Then
            If Trim$(opt("flagno")) <> "" Then flagNo = Trim$(opt("flagno"))
        End If
    End If

    Set exBlank   = CreateObject("Scripting.Dictionary")
    Set exEqCI    = CreateObject("Scripting.Dictionary")
    Set exContCI  = CreateObject("Scripting.Dictionary")
    Set exEqCS    = CreateObject("Scripting.Dictionary")
    Set exContCS  = CreateObject("Scripting.Dictionary")
    Set incEqCS   = CreateObject("Scripting.Dictionary")
    Set incContCS = CreateObject("Scripting.Dictionary")
    Set incContCI = CreateObject("Scripting.Dictionary")

    Set colDict = CreateObject("Scripting.Dictionary")

    ' Detect data range
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Or lastCol < 1 Then Exit Sub

    Set rngData = wsSource.Cells(1, 1).Resize(lastRow, lastCol)
    Set body    = rngData.Offset(1).Resize(rngData.Rows.Count - 1)

    ' Header map (case-insensitive lookup)
    For c = 1 To lastCol
        colDict(LCase(Trim(wsSource.Cells(1, c).Value))) = c
    Next c

    ' Prep
    If Not flagOnly Then wsTarget.Cells.Clear
    If wsSource.AutoFilterMode Then wsSource.AutoFilterMode = False
    rngData.AutoFilter

    ' Parse & apply rules
    If Len(Trim(filterRules)) > 0 Then
        rules = Split(filterRules, ";")
        For Each rule In rules
            rule = Trim(CStr(rule))
            If Len(rule) = 0 Then GoTo NextRule

            ' detect operator (check longer tokens first)
            op = ""
            Select Case True
                Case InStr(rule, "!=^") > 0: op = "!=^"
                Case InStr(rule, "!~^") > 0: op = "!~^"
                Case InStr(rule, "=^")  > 0: op = "=^"
                Case InStr(rule, "~^")  > 0: op = "~^"
                Case InStr(rule, "~?")  > 0: op = "~?"
                Case InStr(rule, "!=")  > 0: op = "!="
                Case InStr(rule, "!~")  > 0: op = "!~"
                Case InStr(rule, "<>")  > 0: op = "<>"
                Case InStr(rule, ">=")  > 0: op = ">="
                Case InStr(rule, "<=")  > 0: op = "<="
                Case InStr(rule, ">")   > 0: op = ">"
                Case InStr(rule, "<")   > 0: op = "<"
                Case InStr(rule, "~")   > 0: op = "~"
                Case InStr(rule, "=")   > 0: op = "="
            End Select
            If op = "" Then GoTo NextRule

            fieldName = Trim(Split(rule, op)(0))
            valueExp  = Trim(Split(rule, op)(1))
            If Not colDict.Exists(LCase(fieldName)) Then GoTo NextRule

            Dim fld As Long: fld = colDict(LCase(fieldName))

            Select Case op
                ' ---------- CASE-SENSITIVE includes ----------
                Case "=^"
                    If Not incEqCS.Exists(fld) Then Set incEqCS(fld) = CreateObject("Scripting.Dictionary")
                    Set dictTmp = incEqCS(fld)
                    If InStr(valueExp, "|") > 0 Then
                        For Each v In Split(valueExp, "|"): dictTmp(Trim(CStr(v))) = True: Next v
                    Else
                        dictTmp(Trim$(valueExp)) = True
                    End If

                Case "~^"
                    If Not incContCS.Exists(fld) Then Set incContCS(fld) = CreateObject("Scripting.Dictionary")
                    Set dictTmp = incContCS(fld)
                    If InStr(valueExp, "|") > 0 Then
                        For Each p In Split(valueExp, "|"): dictTmp(Trim(CStr(p))) = True: Next p
                    Else
                        dictTmp(Trim$(valueExp)) = True
                    End If

                ' ---------- CASE-INSENSITIVE includes (NEW) ----------
                Case "~?"
                    If Not incContCI.Exists(fld) Then Set incContCI(fld) = CreateObject("Scripting.Dictionary")
                    Set dictTmp = incContCI(fld)
                    If InStr(valueExp, "|") > 0 Then
                        For Each p In Split(valueExp, "|")
                            p = Trim(CStr(p))
                            If LCase$(p) = "<blank>" Then
                                dictTmp("__BLANK__") = True
                            Else
                                dictTmp(LCase$(p)) = True
                            End If
                        Next p
                    Else
                        If LCase$(valueExp) = "<blank>" Then
                            dictTmp("__BLANK__") = True
                        Else
                            dictTmp(LCase$(valueExp)) = True
                        End If
                    End If

                ' ---------- CASE-SENSITIVE excludes ----------
                Case "!=^"
                    If valueExp = "" Then
                        exBlank(fld) = True
                    Else
                        If Not exEqCS.Exists(fld) Then Set exEqCS(fld) = CreateObject("Scripting.Dictionary")
                        Set dictTmp = exEqCS(fld)
                        If InStr(valueExp, "|") > 0 Then
                            For Each v In Split(valueExp, "|"): dictTmp(Trim(CStr(v))) = True: Next v
                        Else
                            dictTmp(Trim$(valueExp)) = True
                        End If
                    End If

                Case "!~^"
                    If Not exContCS.Exists(fld) Then Set exContCS(fld) = CreateObject("Scripting.Dictionary")
                    Set dictTmp = exContCS(fld)
                    If InStr(valueExp, "|") > 0 Then
                        For Each p In Split(valueExp, "|"): dictTmp(Trim(CStr(p))) = True: Next p
                    Else
                        dictTmp(Trim$(valueExp)) = True
                    End If

                ' ---------- CASE-INSENSITIVE excludes ----------
                Case "!="
                    If valueExp = "" Then
                        exBlank(fld) = True
                    ElseIf InStr(valueExp, "|") > 0 Then
                        If Not exEqCI.Exists(fld) Then Set exEqCI(fld) = CreateObject("Scripting.Dictionary")
                        Set dictTmp = exEqCI(fld)
                        For Each v In Split(valueExp, "|")
                            dictTmp(LCase(Trim(CStr(v)))) = True
                        Next v
                    Else
                        rngData.AutoFilter Field:=fld, Criteria1:="<>" & valueExp
                    End If

                Case "!~"
                    If Not exContCI.Exists(fld) Then Set exContCI(fld) = CreateObject("Scripting.Dictionary")
                    Set dictTmp = exContCI(fld)
                    If InStr(valueExp, "|") > 0 Then
                        For Each p In Split(valueExp, "|")
                            dictTmp(LCase(Trim(CStr(p)))) = True
                        Next p
                    Else
                        dictTmp(LCase(Trim$(valueExp))) = True
                    End If

                ' ---------- POSITIVE CI rules (AutoFilter) ----------
                Case "="
                    If InStr(valueExp, "|") > 0 Then
                        critArr = Split(valueExp, "|")
                        rngData.AutoFilter Field:=fld, Criteria1:=critArr, Operator:=xlFilterValues
                    Else
                        rngData.AutoFilter Field:=fld, Criteria1:=valueExp
                    End If

                Case "<>", ">", "<", ">=", "<="
                    rngData.AutoFilter Field:=fld, Criteria1:=op & valueExp

                Case "~"
                    critArr = Split(valueExp, "|")
                    If UBound(critArr) = 0 Then
                        rngData.AutoFilter Field:=fld, Criteria1:="*" & Trim(critArr(0)) & "*"
                    Else
                        ' supports two contains terms via OR; for >2, prefer ~?
                        rngData.AutoFilter Field:=fld, _
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

    ' ---------- FLAG ONLY MODE ----------
    If flagOnly Then
        Dim visibleRow As Object, ar As Range, rr As Range, r As Long
        Set visibleRow = CreateObject("Scripting.Dictionary")

        ' collect visible rows after AutoFilter
        For Each ar In vis.Areas
            For Each rr In ar.Rows
                visibleRow(rr.Row) = True
            Next rr
        Next ar

        ' ensure/create flag column on SOURCE
        Dim flagColIdx As Long, srcLastCol As Long
        srcLastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
        flagColIdx = FindCol(wsSource, flagColName)
        If flagColIdx = 0 Then
            flagColIdx = srcLastCol + 1
            wsSource.Cells(1, flagColIdx).Value = flagColName
        Else
            ' clear old flags
            wsSource.Range(wsSource.Cells(2, flagColIdx), wsSource.Cells(lastRow, flagColIdx)).ClearContents
        End If

        ' mark Yes/No
        For r = 2 To lastRow
            If visibleRow.Exists(r) And RowPassesRules(wsSource, r, _
                                                       exBlank, exEqCI, exContCI, _
                                                       exEqCS, exContCS, _
                                                       incEqCS, incContCS, _
                                                       incContCI) Then
                wsSource.Cells(r, flagColIdx).Value = flagYes
            Else
                wsSource.Cells(r, flagColIdx).Value = flagNo
            End If
        Next r

        wsSource.AutoFilterMode = False
        Exit Sub
    End If
    ' ---------- END FLAG ONLY MODE ----------

    ' Copy requested columns, enforcing row-level includes/excludes
    colArr = Split(keepCols, ",")
    pasteCol = 1

    For c = LBound(colArr) To UBound(colArr)
        srcColIdx = 0
        If colDict.Exists(LCase(Trim(colArr(c)))) Then
            srcColIdx = colDict(LCase(Trim(colArr(c))))
        Else
            ' Debug.Print "Skip missing keep column in source: [" & Trim(colArr(c)) & "] on sheet " & wsSource.Name
        End If
        If srcColIdx = 0 Then GoTo NextKeep

        wsTarget.Cells(1, pasteCol).Value = Trim(colArr(c))
        destRow = 2

        Set colVis = Application.Intersect(vis, wsSource.Columns(srcColIdx))
        If Not colVis Is Nothing Then
            Dim cell As Range
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

' ========= RENAME HEADERS (safe: ignore missing originals; avoid collisions; supports multi-origin "A|B:New") =========
Sub ApplyRenameMap(ws As Worksheet, renameMap As String)
    Dim mapArr() As String, pair As Variant
    Dim leftPart As String, newName As String
    Dim candidates() As String, cand As Variant
    Dim lastCol As Long, i As Long
    Dim hdrDict As Object, colIdx As Variant, foundIdx As Variant

    If Trim$(renameMap) = "" Then Exit Sub

    ' Build a header index (case-insensitive)
    Set hdrDict = CreateObject("Scripting.Dictionary")
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        hdrDict(LCase$(Trim$(ws.Cells(1, i).Value))) = i
    Next i

    mapArr = Split(renameMap, ",")
    For Each pair In mapArr
        pair = Trim$(CStr(pair))
        If pair <> "" And InStr(pair, ":") > 0 Then
            leftPart = Trim$(Split(pair, ":", 2)(0))  ' may be "A|B|C"
            newName  = Trim$(Split(pair, ":", 2)(1))
            If newName = "" Then GoTo NextPair

            foundIdx = Empty
            candidates = Split(leftPart, "|")

            ' find the first original that actually exists in the current headers
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
                ' Avoid collision: if new header already exists on different col, skip
                If hdrDict.Exists(LCase$(newName)) And hdrDict(LCase$(newName)) <> colIdx Then
                    ' skip silently
                Else
                    ws.Cells(1, colIdx).Value = newName
                    ' update dictionary: remove any matched originals; then add new
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
                ' none of the originals exist → ignore mapping
            End If
        End If
NextPair:
    Next pair
End Sub

' ========= OPTIONS: headers, autofit, freeze, number formats by header, comma styles =========
Sub ApplyOptions(ws As Worksheet, options As String)
    Dim opt As Object, pairs() As String, p As Variant, colFmt() As String
    Dim colIdx As Long, lastRow As Long

    Set opt = ParseOptions(options)
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

    ' Number formats by header (after renaming)
    ' Syntax: NumFmt=Amount:#,##0.00|UsdEquivalent:#,##0.00|RWA Exposure:0.00%|PaymentDate:yyyy-mm-dd
    If opt.Exists("numfmt") Then
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

    ' Comma styles
    If opt.Exists("commastyle") Then
        ApplyCommaStyleToHeaders ws, CStr(opt("commastyle")), False
    End If
    If opt.Exists("commastyle0") Then
        ApplyCommaStyleToHeaders ws, CStr(opt("commastyle0")), True
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
    Dim opt As Object, isFlagOnly As Boolean

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

        ' parse FlagOnly
        Set opt = ParseOptions(options)
        isFlagOnly = False
        If Not opt Is Nothing Then
            If opt.Exists("flagonly") Then isFlagOnly = CBool(ValueAsBool(opt("flagonly")))
        End If

        FilterAndCopy_Flex wsInput, wsTarget, filterRules, keepCols, options
        If Not isFlagOnly Then
            ApplyRenameMap     wsTarget, renameMap
            ApplyOptions       wsTarget, options
        End If
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

    ' Include equals (case-sensitive): must match at least one allowed value per column
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

    ' Include contains (case-sensitive): must contain at least one allowed pattern per column
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

    ' Include contains (case-insensitive, ~?) -> must hit at least one per column
    For Each k In incContCI.Keys
        v = Trim$(CStr(ws.Cells(r, CLng(k)).Value))
        lv = LCase$(v)
        hit = False

        ' allow <blank>
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

' --- Options parsing + helpers ---
Private Function ParseOptions(options As String) As Object
    Dim d As Object, arr() As String, i As Long, kv() As String, tok As String
    If Trim$(options) = "" Then Exit Function
    Set d = CreateObject("Scripting.Dictionary")
    arr = Split(options, ";")
    For i = LBound(arr) To UBound(arr)
        tok = Trim$(arr(i))
        If tok <> "" Then
            If InStr(tok, "=") > 0 Then
                kv = Split(tok, "=", 2)
                d(LCase$(Trim$(kv(0)))) = Trim$(kv(1))
            Else
                d(LCase$(tok)) = True
            End If
        End If
    Next i
    Set ParseOptions = d
End Function

Private Function ValueAsBool(v As Variant) As Boolean
    Dim s As String
    If VarType(v) = vbBoolean Then
        ValueAsBool = v
    Else
        s = LCase$(Trim$(CStr(v)))
        ValueAsBool = (s = "true" Or s = "yes" Or s = "1" Or s = "y")
    End If
End Function

' --- Comma style helper ---
Private Sub ApplyCommaStyleToHeaders(ws As Worksheet, headersList As String, zeroDecimals As Boolean)
    Dim arr() As String, h As Variant
    Dim colIdx As Long, lastRow As Long
    Dim rng As Range, styleName As String

    If Trim$(headersList) = "" Then Exit Sub
    arr = Split(headersList, "|")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For Each h In arr
        colIdx = FindCol(ws, Trim$(CStr(h))) ' uses final (post-rename) headers
        If colIdx > 0 Then
            Set rng = ws.Range(ws.Cells(2, colIdx), ws.Cells(Application.Max(2, lastRow), colIdx))
            styleName = IIf(zeroDecimals, "Comma [0]", "Comma")
            On Error Resume Next
            rng.Style = styleName                     ' try built-in style
            If Err.Number <> 0 Then                  ' fallback (localized Excel)
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
