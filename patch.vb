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

    ' --- collections for row-level checks ---
    Dim exBlank As Object                         ' colIdx -> True (exclude trimmed blanks)
    Dim exEqCI As Object, exContCI As Object      ' (case-insensitive) excludes
    Dim exEqCS As Object, exContCS As Object      ' (case-sensitive)   excludes
    Dim incEqCS As Object, incContCS As Object    ' (case-sensitive)   includes
    Dim incContCI As Object                       ' (case-insensitive) includes for ~?  NEW
    Dim dictTmp As Object
    Dim v As Variant, p As Variant

    Set exBlank   = CreateObject("Scripting.Dictionary")
    Set exEqCI    = CreateObject("Scripting.Dictionary")
    Set exContCI  = CreateObject("Scripting.Dictionary")
    Set exEqCS    = CreateObject("Scripting.Dictionary")
    Set exContCS  = CreateObject("Scripting.Dictionary")
    Set incEqCS   = CreateObject("Scripting.Dictionary")
    Set incContCS = CreateObject("Scripting.Dictionary")
    Set incContCI = CreateObject("Scripting.Dictionary") ' NEW

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
    wsTarget.Cells.Clear
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
                Case InStr(rule, "~?")  > 0: op = "~?"   ' NEW: CI includes, unlimited OR, supports <blank>
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

    ' Copy requested columns, enforcing row-level includes/excludes
    colArr = Split(keepCols, ",")
    pasteCol = 1

    For c = LBound(colArr) To UBound(colArr)
        srcColIdx = 0
        If colDict.Exists(LCase(Trim(colArr(c)))) Then
            srcColIdx = colDict(LCase(Trim(colArr(c))))
        End If
        If srcColIdx = 0 Then GoTo NextKeep

        wsTarget.Cells(1, pasteCol).Value = Trim(colArr(c))
        destRow = 2

        Set colVis = Application.Intersect(vis, wsSource.Columns(srcColIdx))
        If Not colVis Is Nothing Then
            For Each area In colVis.Areas
                Dim cell As Range
                For Each cell In area.Cells
                    If RowPassesRules(wsSource, cell.Row, _
                                      exBlank, exEqCI, exContCI, _
                                      exEqCS, exContCS, _
                                      incEqCS, incContCS, _
                                      incContCI) Then                     ' NEW param
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

    ' Include equals (case-sensitive) -> must match at least one per column
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

    ' Include contains (case-sensitive) -> must hit at least one per column
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

    ' Include contains (case-insensitive, NEW ~?) -> must hit at least one per column
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






