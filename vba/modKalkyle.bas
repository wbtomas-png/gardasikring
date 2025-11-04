Attribute VB_Name = "modKalkyle"
Option Explicit

Private Const POST_BLOCK_HEIGHT As Long = 18
Private Const POST_DATA_ROWS As Long = 13
Private Const POST_FIRST_DATA_ROW As Long = 8
Private Const POST_TOTAL_LENGTH_OFFSET As Long = POST_DATA_ROWS

Private Enum StolpereglerCols
    COL_REG_HOYDE = 1
    COL_REG_FUNDAMENT
    COL_REG_LAKKERT
    COL_REG_MELLOM
    COL_REG_HJORNE
    COL_REG_SKRASTAG
    COL_REG_VARSEL
    COL_REG_ANT_ROR
    COL_REG_OVERLIGGER_TYPE
    COL_REG_OVERLIGGER_PID
    COL_REG_PIGG_AKTIV
    COL_REG_PIGG_RAL
    COL_REG_PIGG_GALV
    COL_REG_NETTING_RAL
    COL_REG_NETTING_GALV
    COL_REG_STREKK_RAL
    COL_REG_STREKK_GALV
    COL_REG_SYSTR_RAL
    COL_REG_SYSTR_GALV
    COL_REG_PANEL_PID
    COL_REG_DIVERSE_PID
End Enum

Private Type ProductInfo
    PID As String
    Beskrivelse As String
    Enhet As String
    Varekost As Double
    Salgspris As Double
End Type

Public Type StolpeResult
    RegelRad As Long
    mellomPID As String
    hjornePID As String
    SkrastagPID As String
    VarselTekst As String
    AntallRor As Long
    OverliggerType As String
    OverliggerPID As String
    PiggtrGalv As String
    PiggtrRAL As String
    NettingGalv As String
    NettingRAL As String
    StrekkGalv As String
    StrekkRAL As String
    SystrGalv As String
    SystrRAL As String
    PanelPID As String
    DiversePID As String
End Type

Public Function HentStolpeRegel(ByVal hoyde As Long, ByVal fundament As String, _
                                ByVal lakkert As Boolean, ByVal antRor As Long, _
                                ByVal overligger As String, Optional ByVal pigg As Boolean = False) As StolpeResult

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim ralTxt As String
    Dim antRorRad As Variant
    Dim overTxt As String
    Dim piggTxt As String

    Set ws = ThisWorkbook.Sheets("Stolperegler")
    ralTxt = IIf(lakkert, "ja", "nei")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If CLng(ws.Cells(r, COL_REG_HOYDE).Value) = hoyde _
           And StrComp(Trim$(ws.Cells(r, COL_REG_FUNDAMENT).Value), fundament, vbTextCompare) = 0 _
           And StrComp(LCase$(Trim$(ws.Cells(r, COL_REG_LAKKERT).Value)), LCase$(ralTxt), vbTextCompare) = 0 Then

            antRorRad = ws.Cells(r, COL_REG_ANT_ROR).Value
            overTxt = LCase$(Trim$(ws.Cells(r, COL_REG_OVERLIGGER_TYPE).Value))
            piggTxt = LCase$(Trim$(ws.Cells(r, COL_REG_PIGG_AKTIV).Value))

            If pigg Then
                If piggTxt <> "ja" Then GoTo NextRow
            ElseIf piggTxt = "ja" Then
                GoTo NextRow
            End If

            If (LCase$(overligger) = "vinkeloverligger" And InStr(overTxt, "vinkel") > 0) _
               Or (LCase$(overligger) = "rør" And (InStr(overTxt, "rør") > 0 Or InStr(overTxt, "ror") > 0) _
                   And antRor = SafeLong(antRorRad)) _
               Or (LCase$(overligger) = "ingen" And overTxt = "" And SafeLong(antRorRad) = 0) Then

                With HentStolpeRegel
                    .RegelRad = r
                    .mellomPID = Trim$(ws.Cells(r, COL_REG_MELLOM).Value)
                    .hjornePID = Trim$(ws.Cells(r, COL_REG_HJORNE).Value)
                    .SkrastagPID = Trim$(ws.Cells(r, COL_REG_SKRASTAG).Value)
                    .VarselTekst = Trim$(ws.Cells(r, COL_REG_VARSEL).Value)
                    .AntallRor = SafeLong(antRorRad)
                    .OverliggerType = Trim$(ws.Cells(r, COL_REG_OVERLIGGER_TYPE).Value)
                    .OverliggerPID = Trim$(ws.Cells(r, COL_REG_OVERLIGGER_PID).Value)
                    .PiggtrRAL = Trim$(ws.Cells(r, COL_REG_PIGG_RAL).Value)
                    .PiggtrGalv = Trim$(ws.Cells(r, COL_REG_PIGG_GALV).Value)
                    .NettingRAL = Trim$(ws.Cells(r, COL_REG_NETTING_RAL).Value)
                    .NettingGalv = Trim$(ws.Cells(r, COL_REG_NETTING_GALV).Value)
                    .StrekkRAL = Trim$(ws.Cells(r, COL_REG_STREKK_RAL).Value)
                    .StrekkGalv = Trim$(ws.Cells(r, COL_REG_STREKK_GALV).Value)
                    .SystrRAL = Trim$(ws.Cells(r, COL_REG_SYSTR_RAL).Value)
                    .SystrGalv = Trim$(ws.Cells(r, COL_REG_SYSTR_GALV).Value)
                    .PanelPID = Trim$(ws.Cells(r, COL_REG_PANEL_PID).Value)
                    .DiversePID = Trim$(ws.Cells(r, COL_REG_DIVERSE_PID).Value)
                End With

                Exit Function
            End If
        End If
NextRow:
    Next r
End Function

Public Sub ResolvePostsWithPigg(wsRegler As Worksheet, _
                                ByVal hoyde As Long, ByVal fundament As String, _
                                ByVal lakkert As Boolean, ByVal antRor As Long, _
                                ByVal overligger As String, ByVal pigg As Boolean, _
                                ByRef mellomPID As String, ByRef hjornePID As String)

    Dim lastRow As Long
    Dim r As Long
    Dim ralTxt As String
    Dim antRorRad As Variant
    Dim overTxt As String
    Dim piggTxt As String

    If Not pigg Then Exit Sub

    ralTxt = IIf(lakkert, "ja", "nei")
    lastRow = wsRegler.Cells(wsRegler.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If CLng(wsRegler.Cells(r, COL_REG_HOYDE).Value) = hoyde _
           And StrComp(Trim$(wsRegler.Cells(r, COL_REG_FUNDAMENT).Value), fundament, vbTextCompare) = 0 _
           And StrComp(LCase$(Trim$(wsRegler.Cells(r, COL_REG_LAKKERT).Value)), LCase$(ralTxt), vbTextCompare) = 0 Then

            antRorRad = wsRegler.Cells(r, COL_REG_ANT_ROR).Value
            overTxt = LCase$(Trim$(wsRegler.Cells(r, COL_REG_OVERLIGGER_TYPE).Value))
            piggTxt = LCase$(Trim$(wsRegler.Cells(r, COL_REG_PIGG_AKTIV).Value))

            If (LCase$(overligger) = "vinkeloverligger" And InStr(overTxt, "vinkel") > 0) _
               Or (LCase$(overligger) = "rør" And (InStr(overTxt, "rør") > 0 Or InStr(overTxt, "ror") > 0) _
                   And antRor = SafeLong(antRorRad)) _
               Or (LCase$(overligger) = "ingen" And overTxt = "" And SafeLong(antRorRad) = 0) Then

                If piggTxt = "ja" Then
                    mellomPID = Trim$(wsRegler.Cells(r, COL_REG_MELLOM).Value)
                    hjornePID = Trim$(wsRegler.Cells(r, COL_REG_HJORNE).Value)
                    Exit Sub
                End If
            End If
        End If
    Next r
End Sub

Public Function SupportsPiggtråd(wsRegler As Worksheet, _
                                 ByVal hoyde As Long, ByVal fundament As String, _
                                 ByVal lakkert As Boolean, ByVal antRor As Long, _
                                 ByVal overligger As String) As Boolean

    Dim lastRow As Long
    Dim r As Long
    Dim ralTxt As String
    Dim antRorRad As Variant
    Dim overTxt As String

    If wsRegler Is Nothing Then Exit Function

    ralTxt = IIf(lakkert, "ja", "nei")
    lastRow = wsRegler.Cells(wsRegler.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If CLng(wsRegler.Cells(r, COL_REG_HOYDE).Value) = hoyde _
           And StrComp(Trim$(wsRegler.Cells(r, COL_REG_FUNDAMENT).Value), fundament, vbTextCompare) = 0 _
           And StrComp(LCase$(Trim$(wsRegler.Cells(r, COL_REG_LAKKERT).Value)), LCase$(ralTxt), vbTextCompare) = 0 Then

            antRorRad = wsRegler.Cells(r, COL_REG_ANT_ROR).Value
            overTxt = LCase$(Trim$(wsRegler.Cells(r, COL_REG_OVERLIGGER_TYPE).Value))

            If (LCase$(overligger) = "vinkeloverligger" And InStr(overTxt, "vinkel") > 0) _
               Or (LCase$(overligger) = "rør" And (InStr(overTxt, "rør") > 0 Or InStr(overTxt, "ror") > 0) _
                   And antRor = SafeLong(antRorRad)) _
               Or (LCase$(overligger) = "ingen" And overTxt = "" And SafeLong(antRorRad) = 0) Then

                SupportsPiggtråd = (LCase$(Trim$(wsRegler.Cells(r, COL_REG_PIGG_AKTIV).Value)) = "ja")
                If SupportsPiggtråd Then Exit Function
            End If
        End If
    Next r
End Function

Public Sub SettInnProdukt(wsCalc As Worksheet, wsData1 As Worksheet, wsData2 As Worksheet, _
                          ByVal blockStartRow As Long, ByRef nextRow As Long, ByVal endRow As Long, _
                          ByVal PID As String, ByVal Antall As Double, ByVal EnhetFallback As String, _
                          ByVal Dekningsgrad As Double, Optional ByRef matRows As Collection)

    Dim info As ProductInfo
    Dim r As Long

    If Len(PID) = 0 Or Antall <= 0 Then Exit Sub

    info = LookupProductInfo(wsData1, wsData2, PID, EnhetFallback)

    For r = blockStartRow To endRow
        If StrComp(Trim$(wsCalc.Cells(r, 2).Value), info.PID, vbTextCompare) = 0 Then
            wsCalc.Cells(r, 4).Value = CDbl(wsCalc.Cells(r, 4).Value) + Antall
            wsCalc.Cells(r, 8).Value = 0
            RegisterRow matRows, r
            Exit Sub
        End If
    Next r

    If nextRow > endRow Then
        Err.Raise vbObjectError + 1000, "SettInnProdukt", "Ingen ledige linjer i valgt post."
    End If

    wsCalc.Cells(nextRow, 2).Value = info.PID
    wsCalc.Cells(nextRow, 3).Value = info.Beskrivelse
    wsCalc.Cells(nextRow, 4).Value = Antall
    wsCalc.Cells(nextRow, 5).Value = info.Enhet
    wsCalc.Cells(nextRow, 6).Value = FormatPris(info.Salgspris, True)
    wsCalc.Cells(nextRow, 7).Value = FormatPris(info.Varekost, False)
    wsCalc.Cells(nextRow, 8).Value = 0

    RegisterRow matRows, nextRow

    nextRow = nextRow + 1
End Sub

Public Sub SettInnArbeid(wsCalc As Worksheet, wsData1 As Worksheet, wsData2 As Worksheet, _
                         ByVal blockStartRow As Long, ByRef nextRow As Long, ByVal endRow As Long, _
                         ByVal ArbeidNavn As String, ByVal TimerPerMontor As Double, _
                         ByVal Montorer As Long, ByVal Dekningsgrad As Double, _
                         Optional ByRef arbeidRows As Collection)

    Dim info As ProductInfo
    Dim r As Long
    Dim idx As Long
    Dim lineDesc As String

    If Montorer <= 0 Then Exit Sub

    info = LookupProductInfo(wsData1, wsData2, "MONTOR-01", "timer")
    For idx = 1 To Montorer
        lineDesc = ArbeidNavn & " - Montør " & idx

        For r = blockStartRow To endRow
            If StrComp(Trim$(wsCalc.Cells(r, 2).Value), info.PID, vbTextCompare) = 0 _
               And StrComp(Trim$(wsCalc.Cells(r, 3).Value), lineDesc, vbTextCompare) = 0 Then
                wsCalc.Cells(r, 4).Value = TimerPerMontor
                wsCalc.Cells(r, 8).Value = 0
                RegisterRow arbeidRows, r
                GoTo NesteMontor
            End If
        Next r

        If nextRow > endRow Then
            Err.Raise vbObjectError + 1001, "SettInnArbeid", "Ingen ledige linjer i valgt post."
        End If

        wsCalc.Cells(nextRow, 2).Value = info.PID
        wsCalc.Cells(nextRow, 3).Value = lineDesc
        wsCalc.Cells(nextRow, 4).Value = TimerPerMontor
        wsCalc.Cells(nextRow, 5).Value = info.Enhet
        wsCalc.Cells(nextRow, 6).Value = FormatPris(info.Salgspris, True)
        wsCalc.Cells(nextRow, 7).Value = FormatPris(info.Varekost, False)
        wsCalc.Cells(nextRow, 8).Value = 0

        RegisterRow arbeidRows, nextRow

        nextRow = nextRow + 1
NesteMontor:
    Next idx
End Sub

Private Function LookupProductInfo(wsData1 As Worksheet, wsData2 As Worksheet, _
                                   ByVal PID As String, ByVal EnhetFallback As String) As ProductInfo

    Dim info As ProductInfo
    Dim lastRow As Long
    Dim r As Long
    Dim candidate As String
    Dim ralSuffix As String

    info.PID = PID
    info.Enhet = EnhetFallback
    ralSuffix = ExtractRalFromPid(PID)

    If Not wsData1 Is Nothing Then
        lastRow = wsData1.Cells(wsData1.Rows.Count, "B").End(xlUp).Row
        For r = 2 To lastRow
            candidate = Trim$(wsData1.Cells(r, "B").Value)
            If InStr(1, candidate, "(RAL)", vbTextCompare) > 0 And Len(ralSuffix) > 0 Then
                candidate = Replace(candidate, "(RAL)", ralSuffix, , , vbTextCompare)
            End If

            If StrComp(candidate, PID, vbTextCompare) = 0 Then
                info.Beskrivelse = ReplaceRalPlaceholder(wsData1.Cells(r, "C").Value, ralSuffix)
                info.Enhet = wsData1.Cells(r, "F").Value
                Exit For
            End If
        Next r
    End If

    If Len(info.Enhet) = 0 Then
        info.Enhet = EnhetFallback
    End If

    If Len(info.Beskrivelse) = 0 Then
        info.Beskrivelse = PID
    End If

    lastRow = wsData2.Cells(wsData2.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        candidate = Trim$(wsData2.Cells(r, "A").Value)

        If InStr(1, candidate, "(RAL)", vbTextCompare) > 0 And Len(ralSuffix) > 0 Then
            candidate = Replace(candidate, "(RAL)", ralSuffix, , , vbTextCompare)
        End If

        If StrComp(candidate, PID, vbTextCompare) = 0 Then
            If PID = "MONTOR-01" Or PID = "HE-STREVERTAPP" Then
                info.Salgspris = Val(wsData2.Cells(r, "B").Value)
                info.Varekost = Val(wsData2.Cells(r, "C").Value)
            Else
                info.Varekost = Val(wsData2.Cells(r, "C").Value)
                info.Salgspris = 0
            End If
            Exit For
        End If
    Next r

    LookupProductInfo = info
End Function

Public Function ResolvePID(ByVal rawPID As String, ByVal RAL As String) As String
    Dim tmp As String
    Dim erstatning As String

    tmp = Trim$(rawPID)

    If Len(tmp) = 0 Then Exit Function
    If InStr(1, tmp, "(RAL)", vbTextCompare) > 0 Then
        erstatning = ResolveRalPlaceholder(tmp, RAL)

        If Len(erstatning) > 0 Then
            tmp = Replace(tmp, "(RAL)", erstatning, , , vbTextCompare)
        Else
            tmp = Replace(tmp, "-(RAL)", "", , , vbTextCompare)
            tmp = Replace(tmp, "(RAL)", "", , , vbTextCompare)
        End If
    End If

    If Right$(tmp, 1) = "-" Then
        tmp = Left$(tmp, Len(tmp) - 1)
    End If

    ResolvePID = tmp
End Function

Private Function ResolveRalPlaceholder(ByVal rawPID As String, ByVal RAL As String) As String
    Dim lowerPid As String
    Dim normalizedRal As String

    lowerPid = LCase$(rawPID)
    normalizedRal = Trim$(RAL)

    If Len(normalizedRal) = 0 Then Exit Function

    Select Case True
        Case InStr(lowerPid, "he-sy200") > 0
            Select Case normalizedRal
                Case "9005"
                    ResolveRalPlaceholder = "SV"
                Case "6009"
                    ResolveRalPlaceholder = "GR"
                Case Else
                    ResolveRalPlaceholder = normalizedRal
            End Select
        Case Else
            ResolveRalPlaceholder = normalizedRal
    End Select
End Function

Private Function ReplaceRalPlaceholder(ByVal tekst As String, ByVal ralSuffix As String) As String
    If InStr(1, tekst, "(RAL)", vbTextCompare) > 0 Then
        ReplaceRalPlaceholder = Replace(tekst, "(RAL)", ralSuffix, , , vbTextCompare)
    Else
        ReplaceRalPlaceholder = tekst
    End If
End Function

Private Function ExtractRalFromPid(ByVal PID As String) As String
    Dim pos As Long
    Dim suffix As String
    Dim lastTwo As String

    pos = InStrRev(PID, "-")
    If pos > 0 And pos < Len(PID) Then
        suffix = Mid$(PID, pos + 1)
    End If

    If Len(suffix) > 0 Then
        lastTwo = UCase$(Right$(suffix, 2))
        If lastTwo = "SV" Or lastTwo = "GR" Then
            ExtractRalFromPid = lastTwo
        Else
            ExtractRalFromPid = suffix
        End If
        Exit Function
    End If

    If Len(PID) >= 2 Then
        lastTwo = UCase$(Right$(PID, 2))
        If lastTwo = "SV" Or lastTwo = "GR" Then
            ExtractRalFromPid = lastTwo
            Exit Function
        End If
    End If

    ExtractRalFromPid = vbNullString
End Function

Public Function SafeLong(ByVal value As Variant) As Long
    If IsNumeric(value) Then
        SafeLong = CLng(value)
    Else
        SafeLong = 0
    End If
End Function

Public Function NormaliserDekningsgrad(ByVal rawValue As Double) As Double
    If rawValue > 1 Then
        NormaliserDekningsgrad = rawValue / 100#
    Else
        NormaliserDekningsgrad = rawValue
    End If
End Function

Private Function FormatPris(ByVal value As Double, Optional ByVal blankIfZero As Boolean = False) As Variant
    If blankIfZero And Abs(value) < 0.0001 Then
        FormatPris = vbNullString
    Else
        FormatPris = Round(value, 2)
    End If
End Function

Private Sub RegisterRow(ByRef rows As Collection, ByVal rowNumber As Long)
    If rows Is Nothing Then Exit Sub

    On Error Resume Next
    rows.Add rowNumber, CStr(rowNumber)
    If Err.Number = 457 Then
        Err.Clear
    ElseIf Err.Number <> 0 Then
        Err.Raise Err.Number, "RegisterRow", Err.Description
    End If
    On Error GoTo 0
End Sub

Public Sub ApplyMaterialMarkup(wsCalc As Worksheet, ByVal matRows As Collection, _
                               ByVal arbeidRows As Collection, ByVal Dekningsgrad As Double)

    Dim target As Double
    Dim sumCm As Double
    Dim sumSe As Double
    Dim sumCe As Double
    Dim idx As Long
    Dim rowNumber As Long
    Dim markupFactor As Double
    Dim serviceContribution As Double

    If wsCalc Is Nothing Then Exit Sub
    If matRows Is Nothing Then Exit Sub
    If matRows.Count = 0 Then Exit Sub

    target = NormaliserDekningsgrad(Dekningsgrad)

    For idx = 1 To matRows.Count
        rowNumber = CLng(matRows(idx))
        sumCm = sumCm + ToDouble(wsCalc.Cells(rowNumber, 7).Value) * _
                          ToDouble(wsCalc.Cells(rowNumber, 4).Value)
    Next idx

    If Not arbeidRows Is Nothing Then
        For idx = 1 To arbeidRows.Count
            rowNumber = CLng(arbeidRows(idx))
            sumSe = sumSe + ToDouble(wsCalc.Cells(rowNumber, 6).Value) * _
                              ToDouble(wsCalc.Cells(rowNumber, 4).Value)
            sumCe = sumCe + ToDouble(wsCalc.Cells(rowNumber, 7).Value) * _
                              ToDouble(wsCalc.Cells(rowNumber, 4).Value)
        Next idx
    End If

    If sumCm <= 0 Or target >= 0.9999 Then
        markupFactor = 0
    Else
        serviceContribution = sumSe - sumCe
        markupFactor = (target * sumSe + target * sumCm - serviceContribution) / _
                       (sumCm * (1 - target))
        If markupFactor < 0 Then markupFactor = 0
    End If

    For idx = 1 To matRows.Count
        rowNumber = CLng(matRows(idx))
        wsCalc.Cells(rowNumber, 8).Value = Round(markupFactor * 100#, 2)
    Next idx
End Sub

Private Function ToDouble(ByVal value As Variant) As Double
    If IsNumeric(value) Then
        ToDouble = CDbl(value)
    Else
        ToDouble = 0#
    End If
End Function

Public Function PostStartRow(ByVal postNr As Long) As Long
    PostStartRow = POST_FIRST_DATA_ROW + (postNr - 1) * POST_BLOCK_HEIGHT
End Function

Public Function PostEndRow(ByVal postNr As Long) As Long
    PostEndRow = PostStartRow(postNr) + POST_DATA_ROWS - 1
End Function

Public Function PostTextRow(ByVal postNr As Long) As Long
    PostTextRow = PostStartRow(postNr) + 15
End Function

Public Function PostTotalLengthRow(ByVal postNr As Long) As Long
    PostTotalLengthRow = PostStartRow(postNr) + POST_TOTAL_LENGTH_OFFSET
End Function

Public Sub ClearPostArea(wsCalc As Worksheet, ByVal startRow As Long, ByVal endRow As Long)
    wsCalc.Range(wsCalc.Cells(startRow, 2), wsCalc.Cells(endRow, 8)).ClearContents
End Sub

Public Function BuildTilbudstekst(ByVal hoyde As Long, ByVal TraadType As String, _
                                  ByVal RAL As String, ByVal OverliggerType As String, _
                                  ByVal AntallRor As Long, ByVal Fundament As String, _
                                  ByVal Piggrad As Boolean) As String

    Dim fargeTxt As String
    Dim overTxt As String
    Dim piggTxt As String
    Dim typeTxt As String

    ' === Tilpass tilbudsteksten her ===
    Select Case LCase$(TraadType)
        Case "panel"
            typeTxt = "panelgjerde"
            If Len(RAL) > 0 Then
                fargeTxt = "RAL " & RAL
            Else
                fargeTxt = "galvanisert"
            End If
        Case "ral"
            typeTxt = "flettverksgjerde"
            fargeTxt = "lakkert (" & RAL & ")"
        Case Else
            typeTxt = "flettverksgjerde"
            fargeTxt = "galvanisert"
    End Select

    Select Case LCase$(OverliggerType)
        Case "rør"
            overTxt = AntallRor & " rør i topp"
        Case "vinkeloverligger"
            overTxt = "vinkeloverligger"
        Case Else
            overTxt = "uten overligger"
    End Select

    If Piggrad Then
        piggTxt = " med piggtråd (3 løp)"
    Else
        piggTxt = ""
    End If

    BuildTilbudstekst = "Levering og montering av " & typeTxt & " H=" & hoyde & " mm, " & _
                        fargeTxt & ", " & overTxt & piggTxt & ". " & _
                        "Fundamentering: " & Fundament & ". Inkluderer stolper, netting, stag og " & _
                        "nødvendig festemateriell."
    ' === Slutt på tilbudstekst ===
End Function
