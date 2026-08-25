Attribute VB_Name = "MonitorDuracionF0"
Option Explicit

'==================================================================
'  DURACION FONDO 0  -  v11
'  - fCorte: los vencimientos se miden desde la fecha de la FMS,
'    no desde hoy (asi los depositos que vencen entre FMS y hoy
'    no se pierden).
'  - Rating ponderado (corto y largo) integrado.
'  - Flujos por Asset Class (3 categorias) con graficos apilados.
'  Config: B1 carpeta FMS | B2 carpeta vector | B3 fecha FMS (AUTO) |
'          B4 fecha vector (AUTO) | B5 carpeta archivo historico
'==================================================================

' --- Cartera (FMS) ---
Private Const C_FOND As Long = 1, C_ASSET As Long = 3, C_SBS As Long = 4
Private Const C_EMISOR As Long = 5, C_NEMO As Long = 7, C_MONEDA As Long = 8
Private Const C_CANT As Long = 9, C_RATING As Long = 11, C_VCTO As Long = 12
' --- Anexo I ---
Private Const A_FOND As Long = 1, A_DESC As Long = 4, A_MONTO As Long = 6
' --- Vector SBS ---
Private Const V_CODE As Long = 1, V_YTW As Long = 14, V_SPR As Long = 15, V_DUR As Long = 24
' --- Base F0 ---
Private Const B_HDR As Long = 3, B_DAT As Long = 4
Private Const B_ISIN As Long = 3, B_SBS As Long = 5

Sub CalcularDuracionFondo0()
    Dim wbHome As Workbook: Set wbHome = ThisWorkbook
    Dim wbFMS As Workbook, wbVec As Workbook
    Dim openedFMS As Boolean, openedVec As Boolean, okCalc As Boolean
    Dim i As Long, msgErr As String, saveMsg As String
    Dim hoy As Date: hoy = Date
    Dim paso As String

    On Error GoTo Limpieza
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    '--- 1. CONFIG ---
    paso = "1-Config"
    Dim cFMS As String, cVec As String, rFMS As String, rVec As String, archivo As String
    LeerConfig wbHome, cFMS, cVec, rFMS, rVec, archivo
    Dim dFMS As String, dVec As String
    dFMS = ResolverFecha(rFMS, 2, hoy)
    dVec = ResolverFecha(rVec, 1, hoy)

    ' Fecha de corte = fecha de la FMS (MANDA para vencimientos, NO hoy)
    Dim fCorte As Date
    If Len(dFMS) = 8 Then fCorte = CDate(ToFecha(dFMS)) Else fCorte = hoy

    Dim pFMS As String, pVec As String
    If Len(cFMS) > 0 And Len(dFMS) = 8 Then
        If Right(cFMS, 1) <> "\" Then cFMS = cFMS & "\"
        pFMS = cFMS & "FMS_" & dFMS & ".xlsx"
    End If
    If Len(cVec) > 0 And Len(dVec) = 8 Then
        If Right(cVec, 1) <> "\" Then cVec = cVec & "\"
        pVec = cVec & dVec & " RFL.xls"
    End If

    '--- 2. CURVA CP Duracion ---
    paso = "2-Curva CP Duracion"
    Dim cvDias() As Double, cvDur() As Double, nC As Long
    LeerCurvaDur wbHome, cvDias, cvDur, nC

    '--- 3. ABRIR ARCHIVOS ---
    paso = "3-Abrir FMS"
    Set wbFMS = AbrirLibro(pFMS, "Selecciona el FMS (t-2)", openedFMS)
    If wbFMS Is Nothing Then GoTo Limpieza
    paso = "3-Abrir Vector"
    Set wbVec = AbrirLibro(pVec, "Selecciona el vector SBS (AAAAMMDD RFL)", openedVec)
    If wbVec Is Nothing Then GoTo Limpieza

    '--- 4. CONSERVAR ISIN MANUAL ---
    paso = "4-Capturar ISIN"
    Dim prevISIN As Object: Set prevISIN = CapturarISIN(wbHome)

    '--- 5. VOLCAR VECTOR -> HOJA "SBS" ---
    paso = "5-Volcar vector"
    Dim wsSBS As Worksheet: Set wsSBS = VolcarVectorASBS(wbVec, wbHome)

    '--- 6. DICCIONARIOS ---
    paso = "6-Diccionarios"
    Dim dYTW As Object, dSPR As Object, dDUR As Object
    Set dYTW = CreateObject("Scripting.Dictionary")
    Set dSPR = CreateObject("Scripting.Dictionary")
    Set dDUR = CreateObject("Scripting.Dictionary")
    Dim lastS As Long: lastS = wsSBS.Cells(wsSBS.Rows.Count, 5).End(xlUp).Row
    If lastS >= 2 Then
        Dim datS As Variant, kk As String
        datS = wsSBS.Range(wsSBS.Cells(2, 1), wsSBS.Cells(lastS, 5)).Value
        For i = 1 To UBound(datS, 1)
            kk = CStr(datS(i, 5))
            If Len(kk) > 0 And Not dDUR.Exists(kk) Then
                dYTW(kk) = datS(i, 2): dSPR(kk) = datS(i, 3): dDUR(kk) = datS(i, 4)
            End If
        Next i
    End If

    '--- 7. VALOR CARTERA FONDO 0 ---
    paso = "7-Valor cartera"
    Dim valorCartera As Double, hayValor As Boolean, desc As String
    Dim wsA As Worksheet: Set wsA = wbFMS.Worksheets("Anexo I")
    Dim lastA As Long: lastA = wsA.Cells(wsA.Rows.Count, A_FOND).End(xlUp).Row
    If lastA >= 2 Then
        Dim datA As Variant
        datA = wsA.Range(wsA.Cells(2, 1), wsA.Cells(lastA, A_MONTO)).Value
        For i = 1 To UBound(datA, 1)
            If NormFond(datA(i, A_FOND)) = "0" Then
                desc = Trim(CStr(datA(i, A_DESC)))
                If Len(desc) > 0 Then
                    If Split(desc, " ")(0) = "1.1.1" And IsNumeric(datA(i, A_MONTO)) Then
                        valorCartera = CDbl(datA(i, A_MONTO)): hayValor = True: Exit For
                    End If
                End If
            End If
        Next i
    End If

    '--- 8. LEER CARTERA FONDO 0 ---
    paso = "8-Leer cartera"
    Dim wsC As Worksheet: Set wsC = wbFMS.Worksheets("Cartera")
    Dim lastC As Long: lastC = wsC.Cells(wsC.Rows.Count, C_FOND).End(xlUp).Row
    Dim datC As Variant
    If lastC >= 2 Then datC = wsC.Range(wsC.Cells(2, 1), wsC.Cells(lastC, C_VCTO)).Value

    '--- 9. HOJA "Base F0" ---
    paso = "9-Encabezados Base"
    Dim wsB As Worksheet: Set wsB = HojaFija(wbHome, "Base F0")
    wsB.Range("A1").Value = "BASE FONDO 0": wsB.Range("A1").Font.Bold = True
    wsB.Range("C1").Value = "FMS (t-2): " & dFMS & "   |   Vector (t): " & dVec & "   |   Calculo: " & Format(hoy, "dd/mm/yyyy")
    Dim hdr As Variant, c As Long
    hdr = Array("Asset Class", "Emisor", "ISIN", "Nemonico", "Codigo SBS", _
                "Moneda", "Monto (mn)", "Vcto", "YTW", "Duracion", "Spread", "Rating", "Origen")
    For c = 0 To UBound(hdr): wsB.Cells(B_HDR, c + 1).Value = hdr(c): Next c
    wsB.Range(wsB.Cells(B_HDR, 1), wsB.Cells(B_HDR, 13)).Font.Bold = True

    '--- 10. LLENAR BASE + ACUMULAR ---
    paso = "10-Llenar base"
    Dim r As Long: r = B_DAT
    Dim nInstr As Long, nSinDur As Long, nCup As Long
    Dim cod As String, cant As Double, monMn As Double
    Dim ytw As Variant, spr As Variant, dur As Variant, origen As String
    Dim totMonto As Double
    Dim swYTW As Double, sumYTW As Double, swSPR As Double, sumSPR As Double, swDUR As Double, sumDUR As Double
    Dim nMax As Long: nMax = 0
    If Not IsEmpty(datC) Then nMax = UBound(datC, 1)
    Dim mx As Long: mx = Application.Max(1, nMax)
    Dim aNemo() As String, aCod() As String, aMon() As Double
    Dim aCY() As Variant, aCD() As Variant, aCS() As Variant
    ReDim aNemo(1 To mx): ReDim aCod(1 To mx): ReDim aMon(1 To mx)
    ReDim aCY(1 To mx): ReDim aCD(1 To mx): ReDim aCS(1 To mx)
    Dim cpNemo() As String, cpCod() As String, cpVcto() As Variant, cpDias() As Double, cpDur() As Variant
    ReDim cpNemo(1 To mx): ReDim cpCod(1 To mx): ReDim cpVcto(1 To mx)
    ReDim cpDias(1 To mx): ReDim cpDur(1 To mx)
    Dim idx As Long: idx = 0

    If Not IsEmpty(datC) Then
        For i = 1 To UBound(datC, 1)
            If NormFond(datC(i, C_FOND)) = "0" Then
                nInstr = nInstr + 1
                cod = NormCode(datC(i, C_SBS))
                cant = ANum(datC(i, C_CANT))
                monMn = cant / 1000000

                If dDUR.Exists(cod) Then
                    ytw = dYTW(cod): spr = dSPR(cod): dur = dDUR(cod): origen = "Vector"
                Else
                    Dim vd As Variant: vd = ToFecha(datC(i, C_VCTO))
                    If IsDate(vd) And nC >= 1 Then
                        Dim diasB As Double: diasB = CDbl(CDate(vd) - fCorte)
                        dur = InterpLineal(diasB, cvDias, cvDur, nC)
                        ytw = "": spr = "": origen = "Deposito (curva)"
                        nCup = nCup + 1
                        cpNemo(nCup) = CStr(datC(i, C_NEMO)): cpCod(nCup) = cod
                        cpVcto(nCup) = CDate(vd): cpDias(nCup) = diasB: cpDur(nCup) = dur
                    Else
                        ytw = "": spr = "": dur = "NO CALCULABLE": origen = "Sin vcto/curva"
                        nSinDur = nSinDur + 1
                    End If
                End If

                wsB.Cells(r, 1).Value = datC(i, C_ASSET)
                wsB.Cells(r, 2).Value = datC(i, C_EMISOR)
                If prevISIN.Exists(cod) Then wsB.Cells(r, 3).Value = prevISIN(cod)
                wsB.Cells(r, 4).Value = datC(i, C_NEMO)
                wsB.Cells(r, 5).Value = cod
                wsB.Cells(r, 6).Value = datC(i, C_MONEDA)
                wsB.Cells(r, 7).Value = monMn
                wsB.Cells(r, 8).Value = ToFecha(datC(i, C_VCTO))
                wsB.Cells(r, 9).Value = ytw
                wsB.Cells(r, 10).Value = dur
                wsB.Cells(r, 11).Value = spr
                wsB.Cells(r, 12).Value = datC(i, C_RATING)
                wsB.Cells(r, 13).Value = origen

                totMonto = totMonto + monMn
                idx = idx + 1
                aNemo(idx) = CStr(datC(i, C_NEMO)): aCod(idx) = cod: aMon(idx) = monMn
                If IsNumeric(ytw) And monMn <> 0 Then
                    sumYTW = sumYTW + monMn * CDbl(ytw): swYTW = swYTW + monMn: aCY(idx) = monMn * CDbl(ytw)
                Else: aCY(idx) = ""
                End If
                If IsNumeric(dur) And monMn <> 0 Then
                    sumDUR = sumDUR + monMn * CDbl(dur): swDUR = swDUR + monMn: aCD(idx) = monMn * CDbl(dur)
                Else: aCD(idx) = ""
                End If
                If IsNumeric(spr) And monMn <> 0 Then
                    sumSPR = sumSPR + monMn * CDbl(spr): swSPR = swSPR + monMn: aCS(idx) = monMn * CDbl(spr)
                Else: aCS(idx) = ""
                End If
                r = r + 1
            End If
        Next i
    End If

    paso = "10b-Formato Base"
    wsB.Columns("G").NumberFormat = "#,##0.00"
    wsB.Columns("H").NumberFormat = "dd/mm/yyyy"
    wsB.Columns("I").NumberFormat = "0.00"
    wsB.Columns("J").NumberFormat = "0.0000"
    wsB.Columns("K").NumberFormat = "0"
    wsB.Columns("A:M").AutoFit
    FormatoInstitucional wsB, B_HDR, 1, 13

    '--- 11. HOJA "Cupon F0" ---
    paso = "11-Cupon F0"
    Dim wsCup As Worksheet: Set wsCup = HojaFija(wbHome, "Cupon F0")
    wsCup.Range("A1").Value = "INSTRUMENTOS SIN VECTOR (duracion por interpolacion)"
    wsCup.Range("A1").Font.Bold = True
    Dim hCup As Variant: hCup = Array("Nemonico", "Codigo SBS", "Vcto", "Dias (FMS->vcto)", "Duracion interp.")
    For c = 0 To UBound(hCup): wsCup.Cells(3, c + 1).Value = hCup(c): Next c
    wsCup.Range(wsCup.Cells(3, 1), wsCup.Cells(3, 5)).Font.Bold = True
    Dim j As Long
    For j = 1 To nCup
        wsCup.Cells(3 + j, 1).Value = cpNemo(j)
        wsCup.Cells(3 + j, 2).Value = cpCod(j)
        wsCup.Cells(3 + j, 3).Value = cpVcto(j)
        wsCup.Cells(3 + j, 4).Value = cpDias(j)
        wsCup.Cells(3 + j, 5).Value = cpDur(j)
    Next j
    wsCup.Columns("C").NumberFormat = "dd/mm/yyyy"
    wsCup.Columns("E").NumberFormat = "0.0000"
    wsCup.Columns("A:E").AutoFit

    '--- 12. HOJA "Calculo F0" ---
    paso = "12-Calculo F0"
    Dim wsK As Worksheet: Set wsK = HojaFija(wbHome, "Calculo F0")
    Dim wYTW As Double, wDUR As Double, wSPR As Double
    If swYTW > 0 Then wYTW = sumYTW / swYTW
    If swDUR > 0 Then wDUR = sumDUR / swDUR
    If swSPR > 0 Then wSPR = sumSPR / swSPR

    wsK.Range("A1").Value = "CALCULO FONDO 0": wsK.Range("A1").Font.Bold = True
    wsK.Range("A2").Value = "FMS (t-2):": wsK.Range("B2").Value = dFMS
    wsK.Range("C2").Value = "Vector (t):": wsK.Range("D2").Value = dVec
    wsK.Range("A4").Value = "Valor cartera F0 (Anexo I 1.1.1):"
    wsK.Range("D4").Value = IIf(hayValor, valorCartera, "NO ENCONTRADO")
    wsK.Range("D4").NumberFormat = "#,##0.00"
    wsK.Range("A5").Value = "Total instrumentos (S monto, mn):": wsK.Range("D5").Value = totMonto
    wsK.Range("D5").NumberFormat = "#,##0.00"
    wsK.Range("A6").Value = "YTW ponderado (excl. cupon):": wsK.Range("D6").Value = wYTW
    wsK.Range("D6").NumberFormat = "0.00"
    wsK.Range("A7").Value = "Duracion ponderada:": wsK.Range("D7").Value = wDUR
    wsK.Range("D7").NumberFormat = "0.0000": wsK.Range("D7").Font.Bold = True
    If swDUR > 0 Then
        If wDUR >= 1 Then
            wsK.Range("D7").Interior.Color = RGB(255, 150, 150)
            wsK.Range("E7").Value = "EXCEDE LIMITE (>1)": wsK.Range("E7").Font.Color = RGB(192, 0, 0)
        Else
            wsK.Range("D7").Interior.Color = RGB(180, 230, 180)
        End If
    End If
    wsK.Range("A8").Value = "Spread ponderado (excl. cupon):": wsK.Range("D8").Value = wSPR
    wsK.Range("D8").NumberFormat = "0.00"
    wsK.Range("A9").Value = "Rating ponderado Corto"
    wsK.Range("A10").Value = "Instrumentos:"
    wsK.Range("D10").Value = nInstr & " total  |  " & nCup & " sin vector (curva)  |  " & nSinDur & " sin calcular"
    wsK.Range("A11").Value = "Cobertura duracion:"
    If totMonto > 0 Then
        wsK.Range("D11").Value = Format(swDUR / totMonto, "0.0%") & " del monto"
    Else
        wsK.Range("D11").Value = "s/d"
    End If

    Dim h2 As Variant, kr As Long: kr = 13
    h2 = Array("Nemonico", "Codigo SBS", "Monto (mn)", "MontoxYTW", "MontoxDuracion", "MontoxSpread")
    For c = 0 To UBound(h2): wsK.Cells(kr, c + 1).Value = h2(c): Next c
    wsK.Range(wsK.Cells(kr, 1), wsK.Cells(kr, 6)).Font.Bold = True
    For j = 1 To idx
        wsK.Cells(kr + j, 1).Value = aNemo(j): wsK.Cells(kr + j, 2).Value = aCod(j)
        wsK.Cells(kr + j, 3).Value = aMon(j): wsK.Cells(kr + j, 4).Value = aCY(j)
        wsK.Cells(kr + j, 5).Value = aCD(j): wsK.Cells(kr + j, 6).Value = aCS(j)
    Next j
    Dim tr As Long: tr = kr + idx + 1
    wsK.Cells(tr, 1).Value = "TOTAL": wsK.Cells(tr, 1).Font.Bold = True
    wsK.Cells(tr, 3).Value = totMonto: wsK.Cells(tr, 4).Value = sumYTW
    wsK.Cells(tr, 5).Value = sumDUR: wsK.Cells(tr, 6).Value = sumSPR
    wsK.Range(wsK.Cells(tr, 3), wsK.Cells(tr, 6)).Font.Bold = True
    wsK.Range(wsK.Cells(kr + 1, 3), wsK.Cells(tr, 6)).NumberFormat = "#,##0.00"
    wsK.Columns("A:F").AutoFit
    FormatoInstitucional wsK, kr, 1, 6

    '--- 12b. RATING ---
    paso = "12b-Rating"
    CrearHojaRating wbHome
    CalcularRatingEnCalculo wbHome

    '--- 12c. FLUJOS ---
    paso = "12c-Flujos"
    GenerarFlujos wbHome

    '--- 13. GUARDAR COPIA ---
    paso = "13-Guardar copia"
    If swDUR > 0 And Len(archivo) > 0 Then
        If Right(archivo, 1) <> "\" Then archivo = archivo & "\"
        If Len(Dir(archivo, vbDirectory)) > 0 Then
            Dim outPath As String: outPath = archivo & "F0_Duracion_" & BestDateCode(dVec, dFMS) & ".xlsx"
            wbHome.Sheets(Array("Base F0", "Calculo F0")).Copy
            Dim wbOut As Workbook: Set wbOut = ActiveWorkbook
            wbOut.SaveAs Filename:=outPath, FileFormat:=51
            wbOut.Close SaveChanges:=False
            saveMsg = "Archivado: " & outPath
        Else
            saveMsg = "Carpeta archivo (B5) no existe: no se guardo copia."
        End If
    ElseIf Len(archivo) = 0 Then
        saveMsg = "Sin carpeta archivo en Config (B5): no se guardo copia."
    End If
    okCalc = True

Limpieza:
    If Err.Number <> 0 Then msgErr = "ERROR " & Err.Number & " en [" & paso & "]: " & Err.Description
    If openedVec And Not wbVec Is Nothing Then wbVec.Close SaveChanges:=False
    If openedFMS And Not wbFMS Is Nothing Then wbFMS.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If Len(msgErr) > 0 Then
        MsgBox msgErr, vbCritical, "Duracion Fondo 0"
    ElseIf okCalc Then
        MsgBox "Listo." & vbCrLf & _
               "FMS: " & dFMS & "   Vector: " & dVec & vbCrLf & _
               "Instrumentos F0: " & nInstr & "  (curva: " & nCup & ", sin calcular: " & nSinDur & ")" & vbCrLf & _
               "Total monto (mn): " & Format(totMonto, "#,##0.00") & vbCrLf & _
               "Duracion del fondo: " & IIf(swDUR > 0, Format(wDUR, "0.0000") & IIf(wDUR >= 1, "  EXCEDE 1", ""), "s/d") & vbCrLf & _
               saveMsg, vbInformation, "Duracion Fondo 0"
    End If
End Sub

'==================================================================
'  AUXILIARES GENERALES
'==================================================================
Private Function ANum(v As Variant) As Double
    On Error GoTo fallo
    If IsNumeric(v) Then ANum = CDbl(v) Else ANum = 0
    Exit Function
fallo:
    ANum = 0
End Function

'------------------------------------------------------------------
Private Function VolcarVectorASBS(wbVec As Workbook, wbHome As Workbook) As Worksheet
    Dim wsV As Worksheet: Set wsV = wbVec.Worksheets(1)
    Dim lastV As Long: lastV = wsV.Cells(wsV.Rows.Count, V_CODE).End(xlUp).Row
    Dim lastCol As Long: lastCol = wsV.Cells(1, wsV.Columns.Count).End(xlToLeft).Column
    Dim wsS As Worksheet: Set wsS = HojaFija(wbHome, "SBS")
    wsS.Range("A1:E1").Value = Array("Codigo (orig)", "YTW", "Spread", "Duracion", "Codigo sin guiones")
    wsS.Range("A1:E1").Font.Bold = True
    Dim i As Long, outRow As Long: outRow = 2
    Dim cod As Variant
    For i = 2 To lastV
        cod = LeerCelda(wsV, i, V_CODE, lastCol)
        If Len(Trim(CStr(cod))) > 0 Then
            wsS.Cells(outRow, 1).Value = cod
            wsS.Cells(outRow, 2).Value = LeerCelda(wsV, i, V_YTW, lastCol)
            Dim sprRaw As Variant: sprRaw = LeerCelda(wsV, i, V_SPR, lastCol)
            wsS.Cells(outRow, 3).Value = IIf(IsNumeric(sprRaw), sprRaw * 100, sprRaw)
            wsS.Cells(outRow, 4).Value = LeerCelda(wsV, i, V_DUR, lastCol)
            wsS.Cells(outRow, 5).Value = NormCode(cod)
            outRow = outRow + 1
        End If
    Next i
    wsS.Columns("B").NumberFormat = "0.00"
    wsS.Columns("C").NumberFormat = "0"
    wsS.Columns("D").NumberFormat = "0.0000"
    wsS.Columns("A:E").AutoFit
    Set VolcarVectorASBS = wsS
End Function

Private Function LeerCelda(ws As Worksheet, ByVal fila As Long, ByVal col As Long, ByVal maxCol As Long) As Variant
    If col > maxCol Then LeerCelda = "" Else LeerCelda = ws.Cells(fila, col).Value
End Function

'------------------------------------------------------------------
Private Function ResolverFecha(ByVal raw As String, ByVal offHabiles As Long, ByVal hoy As Date) As String
    Dim s As String: s = Trim(raw)
    If Len(s) = 0 Or UCase(s) = "AUTO" Then
        ResolverFecha = Format(RestarHabiles(hoy, offHabiles), "yyyymmdd")
    ElseIf Len(s) = 8 And IsNumeric(s) Then
        ResolverFecha = s
    ElseIf IsDate(s) Then
        ResolverFecha = Format(CDate(s), "yyyymmdd")
    Else
        ResolverFecha = ""
    End If
End Function

Private Function RestarHabiles(ByVal d As Date, ByVal n As Long) As Date
    Dim cnt As Long
    Do While cnt < n
        d = d - 1
        Select Case Weekday(d, vbMonday)
            Case 1 To 5: cnt = cnt + 1
        End Select
    Loop
    RestarHabiles = d
End Function

'------------------------------------------------------------------
Private Sub FormatoInstitucional(ws As Worksheet, ByVal filaHdr As Long, _
                                 ByVal colIni As Long, ByVal colFin As Long)
    Dim ult As Long
    ult = ws.Cells(ws.Rows.Count, colIni).End(xlUp).Row
    If ult < filaHdr Then ult = filaHdr
    ws.Cells.Font.Name = "Arial"
    With ws.Range(ws.Cells(filaHdr, colIni), ws.Cells(ult, colFin)).Font
        .Name = "Arial": .Size = 8
    End With
    With ws.Range(ws.Cells(filaHdr, colIni), ws.Cells(filaHdr, colFin))
        .Interior.Color = RGB(212, 12, 12)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(filaHdr, colIni), ws.Cells(ult, colFin))
    rng.Borders(xlEdgeLeft).LineStyle = xlNone
    rng.Borders(xlEdgeRight).LineStyle = xlNone
    rng.Borders(xlInsideVertical).LineStyle = xlNone
    rng.Borders(xlInsideHorizontal).LineStyle = xlNone
    With rng.Borders(xlEdgeTop): .LineStyle = xlContinuous: .Weight = xlThin: End With
    With rng.Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlThin: End With
    With ws.Range(ws.Cells(filaHdr, colIni), ws.Cells(filaHdr, colFin)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(0, 0, 0)
    End With
End Sub

'------------------------------------------------------------------
Private Sub LeerCurvaDur(wb As Workbook, ByRef cvDias() As Double, ByRef cvDur() As Double, ByRef nC As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("CP Duracion")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "CP Duracion"
        ws.Range("A1").Value = "Curva CP Duracion (pegar de SBS). Interpola C3:C8 (dias) y D3:D8 (duracion)."
        ws.Range("A2").Value = "Nodo": ws.Range("C2").Value = "Dias": ws.Range("D2").Value = "Duracion"
        ws.Range("A2:D2").Font.Bold = True
        Dim nod As Variant, di As Variant, k As Long
        nod = Array("CD1D", "CD1M", "CD3M", "CD6M", "CD9M", "CD1A")
        di = Array(1, 30, 90, 180, 270, 360)
        For k = 0 To 5
            ws.Cells(3 + k, 1).Value = nod(k)
            ws.Cells(3 + k, 3).Value = di(k)
            ws.Cells(3 + k, 4).Value = di(k) / 360
        Next k
        ws.Columns("A:D").AutoFit
    End If
    Dim i As Long, cnt As Long: cnt = 0
    ReDim cvDias(1 To 6): ReDim cvDur(1 To 6)
    For i = 3 To 8
        If IsNumeric(ws.Cells(i, 3).Value) And Len(Trim(CStr(ws.Cells(i, 3).Value))) > 0 _
           And IsNumeric(ws.Cells(i, 4).Value) Then
            cnt = cnt + 1
            cvDias(cnt) = CDbl(ws.Cells(i, 3).Value)
            cvDur(cnt) = CDbl(ws.Cells(i, 4).Value)
        End If
    Next i
    nC = cnt
End Sub

'------------------------------------------------------------------
Private Function InterpLineal(ByVal x As Double, xs() As Double, ys() As Double, ByVal n As Long) As Variant
    Dim i As Long
    If n < 1 Then InterpLineal = "": Exit Function
    If n = 1 Then InterpLineal = ys(1): Exit Function
    If x <= xs(1) Then InterpLineal = ys(1): Exit Function
    If x >= xs(n) Then InterpLineal = ys(n): Exit Function
    For i = 1 To n - 1
        If x >= xs(i) And x <= xs(i + 1) Then
            If xs(i + 1) = xs(i) Then
                InterpLineal = ys(i)
            Else
                InterpLineal = ys(i) + (ys(i + 1) - ys(i)) * (x - xs(i)) / (xs(i + 1) - xs(i))
            End If
            Exit Function
        End If
    Next i
    InterpLineal = ys(n)
End Function

'------------------------------------------------------------------
Private Function CapturarISIN(wb As Workbook) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("Base F0")
    On Error GoTo 0
    If Not ws Is Nothing Then
        Dim last As Long: last = ws.Cells(ws.Rows.Count, B_SBS).End(xlUp).Row
        Dim i As Long, cod As String, isin As String
        For i = B_DAT To last
            cod = Trim(CStr(ws.Cells(i, B_SBS).Value))
            isin = Trim(CStr(ws.Cells(i, B_ISIN).Value))
            If Len(cod) > 0 And Len(isin) > 0 Then
                If Not d.Exists(cod) Then d(cod) = isin
            End If
        Next i
    End If
    Set CapturarISIN = d
End Function

'------------------------------------------------------------------
Private Sub LeerConfig(wb As Workbook, ByRef cFMS As String, ByRef cVec As String, _
                       ByRef rFMS As String, ByRef rVec As String, ByRef archivo As String)
    Dim ws As Worksheet, nueva As Boolean
    On Error Resume Next
    Set ws = wb.Worksheets("Config")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "Config"
        nueva = True
    End If
    ws.Range("A1").Value = "Carpeta del FMS:"
    ws.Range("A2").Value = "Carpeta del vector SBS:"
    ws.Range("A3").Value = "Fecha FMS (AUTO = t-2 hab., o escribe fecha):"
    ws.Range("A4").Value = "Fecha Vector (AUTO = t-1 hab., o escribe fecha):"
    ws.Range("A5").Value = "Carpeta archivo historico:"
    If nueva Then
        ws.Range("B3").Value = "AUTO": ws.Range("B4").Value = "AUTO"
    End If
    ws.Columns("A").AutoFit
    cFMS = Trim(CStr(ws.Range("B1").Value))
    cVec = Trim(CStr(ws.Range("B2").Value))
    rFMS = Trim(CStr(ws.Range("B3").Value))
    rVec = Trim(CStr(ws.Range("B4").Value))
    archivo = Trim(CStr(ws.Range("B5").Value))
End Sub

'------------------------------------------------------------------
Private Function AbrirLibro(ByVal ruta As String, ByVal titulo As String, ByRef abierto As Boolean) As Workbook
    Dim wb As Workbook, nombre As String, f As Variant
    abierto = False
    If Len(ruta) > 0 Then
        nombre = Dir(ruta)
        If Len(nombre) > 0 Then
            On Error Resume Next
            Set wb = Workbooks(nombre)
            On Error GoTo 0
            If Not wb Is Nothing Then Set AbrirLibro = wb: Exit Function
            Set wb = Workbooks.Open(ruta, ReadOnly:=True, UpdateLinks:=False)
            abierto = True: Set AbrirLibro = wb: Exit Function
        End If
    End If
    Dim tit As String: tit = titulo
    If Len(ruta) > 0 Then tit = tit & "  [no encontrado: " & ruta & "]"
    f = Application.GetOpenFilename("Excel (*.xls*),*.xls*", , tit)
    If VarType(f) = vbBoolean Then Exit Function
    nombre = Dir(CStr(f))
    On Error Resume Next
    Set wb = Workbooks(nombre)
    On Error GoTo 0
    If wb Is Nothing Then
        Set wb = Workbooks.Open(CStr(f), ReadOnly:=True, UpdateLinks:=False)
        abierto = True
    End If
    Set AbrirLibro = wb
End Function

'------------------------------------------------------------------
Private Function HojaFija(wb As Workbook, nombre As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nombre)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = nombre
    Else
        ws.Cells.Clear
    End If
    Set HojaFija = ws
End Function

'------------------------------------------------------------------
Private Function NormCode(v As Variant) As String
    Dim s As String
    If IsNumeric(v) Then s = Format(v, "0") Else s = Trim(CStr(v))
    NormCode = UCase(Trim(Replace(s, "-", "")))
End Function
Private Function NormFond(v As Variant) As String
    Dim s As String
    s = Trim(CStr(v))
    If Len(s) = 0 Then NormFond = "": Exit Function
    If IsNumeric(s) Then NormFond = CStr(CLng(Val(s))) Else NormFond = UCase(s)
End Function
Private Function ToFecha(v As Variant) As Variant
    Dim s As String
    If IsNumeric(v) Then s = Format(v, "0") Else s = Trim(CStr(v))
    If Len(s) = 8 Then ToFecha = DateSerial(CLng(Left(s, 4)), CLng(Mid(s, 5, 2)), CLng(Mid(s, 7, 2))) Else ToFecha = ""
End Function
Private Function BestDateCode(dVec As String, dFMS As String) As String
    If Len(dVec) = 8 Then BestDateCode = dVec _
    ElseIf Len(dFMS) = 8 Then BestDateCode = dFMS _
    Else BestDateCode = Format(Now, "yyyymmdd")
End Function

'==================================================================
'  RATING
'==================================================================
Private Sub CrearHojaRating(wb As Workbook)
    Dim ws As Worksheet
    Set ws = HojaFija(wb, "Rating F0")
    ws.Range("A1").Value = "METODOLOGIA DE CALIFICACION - FONDO 0"
    ws.Range("A1").Font.Bold = True

    ws.Range("A3").Value = "CORTO PLAZO"
    ws.Range("A4").Value = "Rating"
    ws.Range("B4").Value = "Nota"
    ws.Range("A5").Value = "CPL 1+"
    ws.Range("B5").Value = 10
    ws.Range("A6").Value = "CPL 1"
    ws.Range("B6").Value = 7
    ws.Range("A7").Value = "CPL 1-"
    ws.Range("B7").Value = 4

    ws.Range("D3").Value = "LARGO PLAZO"
    ws.Range("D4").Value = "Local"
    ws.Range("E4").Value = "Internacional"
    ws.Range("F4").Value = "Nota"
    Dim t As Variant, i As Long
    t = Array( _
        Array("Soberanos", "BBB-", 10), _
        Array("AAA", "BB+", 7), _
        Array("AA+/AA", "BB", 6), _
        Array("AA-", "BB-", 5), _
        Array("A+", "B+", 4), _
        Array("A/A-", "B", 3), _
        Array("BBB+/BBB", "B-", 2))
    For i = 0 To UBound(t)
        ws.Cells(5 + i, 4).Value = t(i)(0)
        ws.Cells(5 + i, 5).Value = t(i)(1)
        ws.Cells(5 + i, 6).Value = t(i)(2)
    Next i

    FormatoHeadRating ws, "A3:B3"
    FormatoHeadRating ws, "A4:B4"
    FormatoHeadRating ws, "D3:F3"
    FormatoHeadRating ws, "D4:F4"
    ws.Columns("A:F").AutoFit
End Sub

Private Sub FormatoHeadRating(ws As Worksheet, direccion As String)
    With ws.Range(direccion)
        .Interior.Color = RGB(212, 12, 12)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True: .Font.Name = "Arial": .Font.Size = 8
        .HorizontalAlignment = xlCenter
    End With
End Sub

'------------------------------------------------------------------
Private Sub CalcularRatingEnCalculo(wb As Workbook)
    Dim wsB As Worksheet, wsK As Worksheet
    On Error Resume Next
    Set wsB = wb.Worksheets("Base F0")
    Set wsK = wb.Worksheets("Calculo F0")
    On Error GoTo 0
    If wsB Is Nothing Or wsK Is Nothing Then Exit Sub

    Dim last As Long
    last = wsB.Cells(wsB.Rows.Count, 8).End(xlUp).Row
    Dim i As Long

    Dim sumNMc As Double, sumMc As Double, sumNMl As Double, sumMl As Double

    Dim nemos() As String, rLoc() As String, esc() As String
    Dim notas() As Double, montos() As Double, nDatos As Long
    ReDim nemos(1 To 5000)
    ReDim rLoc(1 To 5000)
    ReDim esc(1 To 5000)
    ReDim notas(1 To 5000)
    ReDim montos(1 To 5000)
    nDatos = 0

    For i = B_DAT To last
        Dim rating As String
        rating = Trim(CStr(wsB.Cells(i, 12).Value))
        Dim nemo As String
        nemo = Trim(CStr(wsB.Cells(i, 1).Value))
        Dim monto As Double
        monto = 0
        If IsNumeric(wsB.Cells(i, 7).Value) Then monto = CDbl(wsB.Cells(i, 7).Value)
        Dim escala As String, nota As Double, locRat As String
        NotaRating rating, escala, nota, locRat
        If escala <> "" And monto > 0 And nota > 0 Then
            nDatos = nDatos + 1
            nemos(nDatos) = nemo
            rLoc(nDatos) = locRat
            esc(nDatos) = escala
            notas(nDatos) = nota
            montos(nDatos) = monto
            If escala = "Corto" Then
                sumNMc = sumNMc + nota * monto
                sumMc = sumMc + monto
            Else
                sumNMl = sumNMl + nota * monto
                sumMl = sumMl + monto
            End If
        End If
    Next i

    Dim pc As Double, pl As Double
    If sumMc > 0 Then pc = sumNMc / sumMc
    If sumMl > 0 Then pl = sumNMl / sumMl

    ' Corto en la fila 9
    wsK.Range("D9").Value = pc
    wsK.Range("D9").NumberFormat = "0.00"
    wsK.Range("E9").Value = EquivCorto(pc)

    ' Largo: actualiza si existe, si no inserta en la fila 10
    Dim filaLargo As Long
    filaLargo = 0
    Dim f As Long
    For f = 1 To 30
        If InStr(1, CStr(wsK.Cells(f, 1).Value), "Rating ponderado Largo", vbTextCompare) > 0 Then
            filaLargo = f
            Exit For
        End If
    Next f
    If filaLargo = 0 Then
        wsK.Rows("10:10").Insert Shift:=xlDown
        filaLargo = 10
        wsK.Range("A10").Value = "Rating ponderado Largo"
    End If
    wsK.Cells(filaLargo, 4).Value = pl
    wsK.Cells(filaLargo, 4).NumberFormat = "0.00"
    wsK.Cells(filaLargo, 5).Value = EquivLargo(pl)

    ' Tabla detalle al final (limpia anterior si existe)
    Dim busca As Range
    Set busca = wsK.Columns(1).Find("DETALLE CALIFICACION POR INSTRUMENTO", LookIn:=xlValues, LookAt:=xlWhole)
    If Not busca Is Nothing Then
        wsK.Range(wsK.Rows(busca.Row), wsK.Rows(wsK.Rows.Count)).ClearContents
    End If

    Dim r0 As Long
    r0 = wsK.Cells(wsK.Rows.Count, 1).End(xlUp).Row + 3
    wsK.Cells(r0, 1).Value = "DETALLE CALIFICACION POR INSTRUMENTO"
    wsK.Cells(r0, 1).Font.Bold = True

    Dim hr As Long
    hr = r0 + 1
    wsK.Cells(hr, 1).Value = "Instrumento"
    wsK.Cells(hr, 2).Value = "Rating local"
    wsK.Cells(hr, 3).Value = "Escala"
    wsK.Cells(hr, 4).Value = "Nota"
    wsK.Cells(hr, 5).Value = "Monto (mn)"
    wsK.Cells(hr, 6).Value = "Nota x Monto"

    Dim r As Long
    r = hr + 1
    For i = 1 To nDatos
        wsK.Cells(r, 1).Value = nemos(i)
        wsK.Cells(r, 2).Value = rLoc(i)
        wsK.Cells(r, 3).Value = esc(i)
        wsK.Cells(r, 4).Value = notas(i)
        wsK.Cells(r, 5).Value = montos(i)
        wsK.Cells(r, 6).Value = notas(i) * montos(i)
        r = r + 1
    Next i
    wsK.Range(wsK.Cells(hr + 1, 5), wsK.Cells(r - 1, 6)).NumberFormat = "#,##0.00"

    With wsK.Range(wsK.Cells(hr, 1), wsK.Cells(hr, 6))
        .Interior.Color = RGB(212, 12, 12)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Name = "Arial"
        .Font.Size = 8
    End With
End Sub

'------------------------------------------------------------------
' Rating LOCAL (de la Base) -> escala y nota. Gobierno Central = 10.
Private Sub NotaRating(ByVal rating As String, ByRef escala As String, ByRef nota As Double, _
                       ByRef locRat As String)
    Dim s As String
    s = UCase(Trim(rating))
    escala = ""
    nota = 0
    locRat = ""

    If InStr(s, "GOBIERNO") > 0 Or InStr(s, "SOBERAN") > 0 Then
        escala = "Largo"
        nota = 10
        locRat = "Soberanos"
        Exit Sub
    End If

    If Left(s, 3) = "CPL" Then
        escala = "Corto"
        Dim c As String
        c = Replace(Trim(Mid(s, 4)), " ", "")
        Select Case c
            Case "1+"
                nota = 10
                locRat = "CPL 1+"
            Case "1"
                nota = 7
                locRat = "CPL 1"
            Case "1-"
                nota = 4
                locRat = "CPL 1-"
        End Select
    ElseIf Left(s, 3) = "LPI" Or Left(s, 2) = "LP" Then
        escala = "Largo"
        Dim x As String
        If Left(s, 3) = "LPI" Then
            x = Trim(Mid(s, 4))
        Else
            x = Trim(Mid(s, 3))
        End If
        x = Replace(x, " ", "")
        Select Case x
            Case "SOBERANOS", "SOBERANO"
                nota = 10
                locRat = "Soberanos"
            Case "AAA"
                nota = 7
                locRat = "AAA"
            Case "AA+", "AA"
                nota = 6
                locRat = "AA+/AA"
            Case "AA-"
                nota = 5
                locRat = "AA-"
            Case "A+"
                nota = 4
                locRat = "A+"
            Case "A", "A-"
                nota = 3
                locRat = "A/A-"
            Case "BBB+", "BBB"
                nota = 2
                locRat = "BBB+/BBB"
            Case "BB+"
                nota = 6
                locRat = "BB+"
        End Select
    End If
End Sub

'------------------------------------------------------------------
Private Function EquivCorto(v As Double) As String
    Dim d10 As Double, d7 As Double, d4 As Double
    d10 = Abs(v - 10)
    d7 = Abs(v - 7)
    d4 = Abs(v - 4)
    If d10 <= d7 And d10 <= d4 Then
        EquivCorto = "CPL 1+"
    ElseIf d7 <= d4 Then
        EquivCorto = "CPL 1"
    Else
        EquivCorto = "CPL 1-"
    End If
End Function

'------------------------------------------------------------------
Private Function EquivLargo(v As Double) As String
    Dim notasArr As Variant, localesArr As Variant, intArr As Variant
    notasArr = Array(10, 7, 6, 5, 4, 3, 2)
    localesArr = Array("Soberanos", "AAA", "AA+/AA", "AA-", "A+", "A/A-", "BBB+/BBB")
    intArr = Array("BBB-", "BB+", "BB", "BB-", "B+", "B", "B-")
    Dim mejor As Long, dmin As Double, i As Long
    mejor = 0
    dmin = 1E+30
    For i = 0 To UBound(notasArr)
        If Abs(v - notasArr(i)) < dmin Then
            dmin = Abs(v - notasArr(i))
            mejor = i
        End If
    Next i
    EquivLargo = localesArr(mejor) & " (" & intArr(mejor) & ")"
End Function

'==================================================================
'  FLUJOS POR ASSET CLASS (3 categorias) + graficos apilados
'==================================================================
Private Sub GenerarFlujos(wb As Workbook)
    Dim wsB As Worksheet, wsF As Worksheet
    On Error Resume Next
    Set wsB = wb.Worksheets("Base F0")
    On Error GoTo 0
    If wsB Is Nothing Then Exit Sub

    Dim last As Long: last = wsB.Cells(wsB.Rows.Count, 8).End(xlUp).Row
    Dim i As Long

    Dim dAnio As Object: Set dAnio = CreateObject("Scripting.Dictionary")
    Dim dMes As Object: Set dMes = CreateObject("Scripting.Dictionary")
    Dim setAnio As Object: Set setAnio = CreateObject("Scripting.Dictionary")
    Dim setMes As Object: Set setMes = CreateObject("Scripting.Dictionary")

    Dim aClases(0 To 2) As String
    aClases(0) = "Bonos"
    aClases(1) = "Depositos a Plazo"
    aClases(2) = "Papeles Comerciales"

    For i = B_DAT To last
        Dim vctoRaw As Variant: vctoRaw = wsB.Cells(i, 8).Value
        Dim acRaw As String: acRaw = Trim(CStr(wsB.Cells(i, 1).Value))
        Dim grp As String: grp = GrupoAC(acRaw)
        Dim monto As Double: monto = 0
        If IsNumeric(wsB.Cells(i, 7).Value) Then monto = CDbl(wsB.Cells(i, 7).Value)

        ' Fecha robusta: acepta fecha real o texto AAAAMMDD
        Dim vcto As Date, vctoOK As Boolean: vctoOK = False
        If IsDate(vctoRaw) Then
            vcto = CDate(vctoRaw): vctoOK = True
        Else
            Dim fConv As Variant: fConv = ToFecha(vctoRaw)
            If IsDate(fConv) Then vcto = CDate(fConv): vctoOK = True
        End If

        If vctoOK And Len(grp) > 0 And Left(grp, 5) <> "OTRO:" Then
            Dim aa As String: aa = CStr(Year(vcto))
            Dim am As String: am = Format(vcto, "yyyymm")
            If Not setAnio.Exists(aa) Then setAnio(aa) = True
            If Not setMes.Exists(am) Then setMes(am) = True
            If Not dAnio.Exists(aa) Then Set dAnio(aa) = CreateObject("Scripting.Dictionary")
            If dAnio(aa).Exists(grp) Then dAnio(aa)(grp) = dAnio(aa)(grp) + monto Else dAnio(aa)(grp) = monto
            If Not dMes.Exists(am) Then Set dMes(am) = CreateObject("Scripting.Dictionary")
            If dMes(am).Exists(grp) Then dMes(am)(grp) = dMes(am)(grp) + monto Else dMes(am)(grp) = monto
        End If
    Next i

    Dim aAnios() As String: aAnios = OrdenaClaves(setAnio, False)
    Dim aMeses() As String: aMeses = OrdenaClaves(setMes, False)

    Set wsF = HojaFija(wb, "Flujos F0")
    Application.DisplayAlerts = False
    Dim ch As ChartObject
    For Each ch In wsF.ChartObjects: ch.Delete: Next ch
    Application.DisplayAlerts = True

    wsF.Range("A1").Value = "FLUJOS DE VENCIMIENTO POR ASSET CLASS (Val_total, mn)"
    wsF.Range("A1").Font.Bold = True

    ' CUADRO 1: ANUAL
    Dim r0 As Long: r0 = 3
    wsF.Cells(r0, 1).Value = "Cuadro 1 - Por ano y Asset Class (mn)"
    wsF.Cells(r0, 1).Font.Bold = True
    Dim hr As Long: hr = r0 + 1
    Dim finAnual As Long
    finAnual = EscribeMatriz(wsF, hr, aClases, aAnios, dAnio, False)

    Dim gAnualTop As Long: gAnualTop = finAnual + 2
    CrearGraficoApilado wsF, hr, finAnual, 1, UBound(aAnios) + 2, _
                        "Vencimientos por ano - Asset Class", gAnualTop

    ' CUADRO 2: MENSUAL
    Dim r2 As Long: r2 = gAnualTop + 16
    wsF.Cells(r2, 1).Value = "Cuadro 2 - Por mes y Asset Class (mn)"
    wsF.Cells(r2, 1).Font.Bold = True
    Dim hr2 As Long: hr2 = r2 + 1
    Dim finMens As Long
    finMens = EscribeMatriz(wsF, hr2, aClases, aMeses, dMes, True)

    Dim gMensTop As Long: gMensTop = finMens + 2
    CrearGraficoApilado wsF, hr2, finMens, 1, UBound(aMeses) + 2, _
                        "Vencimientos por mes - Asset Class", gMensTop

    wsF.Columns("A").ColumnWidth = 24
End Sub

'------------------------------------------------------------------
Private Function GrupoAC(ByVal ac As String) As String
    Dim s As String: s = LCase(Trim(ac))
    Select Case s
        Case "bonos de gobierno central de la republica", _
             "bonos de gobierno central de la república", _
             "bonos de empresas privadas", _
             "otros bonos del sistema financiero", _
             "titulos con derecho crediticio", _
             "títulos con derecho crediticio"
            GrupoAC = "Bonos"
        Case "depositos a plazo", "depósitos a plazo"
            GrupoAC = "Depositos a Plazo"
        Case "cd seriados subastado bancos", _
             "papeles comerciales", _
             "papeles comerciales titulizados"
            GrupoAC = "Papeles Comerciales"
        Case Else
            GrupoAC = "OTRO: " & ac
    End Select
End Function

'------------------------------------------------------------------
Private Function EscribeMatriz(ws As Worksheet, hr As Long, aClases() As String, _
                               aPer() As String, dPer As Object, _
                               Optional periodoEsMes As Boolean = False) As Long
    Dim c As Long, r As Long, j As Long
    ws.Cells(hr, 1).Value = "Asset Class"
    For c = 0 To UBound(aPer)
        If periodoEsMes Then
            ws.Cells(hr, c + 2).Value = MesCorto(aPer(c))
        Else
            ws.Cells(hr, c + 2).Value = "'" & aPer(c)
        End If
    Next c

    r = hr + 1
    For j = 0 To UBound(aClases)
        ws.Cells(r, 1).Value = aClases(j)
        For c = 0 To UBound(aPer)
            Dim v As Double: v = 0
            If dPer.Exists(aPer(c)) Then
                If dPer(aPer(c)).Exists(aClases(j)) Then v = dPer(aPer(c))(aClases(j))
            End If
            ws.Cells(r, c + 2).Value = v
        Next c
        r = r + 1
    Next j

    ws.Cells(r, 1).Value = "TOTAL": ws.Cells(r, 1).Font.Bold = True
    For c = 0 To UBound(aPer)
        Dim tot As Double: tot = 0
        For j = 0 To UBound(aClases)
            tot = tot + ws.Cells(hr + 1 + j, c + 2).Value
        Next j
        ws.Cells(r, c + 2).Value = tot
        ws.Cells(r, c + 2).Font.Bold = True
    Next c
    ws.Range(ws.Cells(hr + 1, 2), ws.Cells(r, UBound(aPer) + 2)).NumberFormat = "#,##0.00"

    With ws.Range(ws.Cells(hr, 1), ws.Cells(hr, UBound(aPer) + 2))
        .Interior.Color = RGB(212, 12, 12)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Name = "Arial": .Font.Size = 8
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(hr, 1), ws.Cells(r, UBound(aPer) + 2))
        .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlThin
        .Font.Name = "Arial": .Font.Size = 8
    End With
    EscribeMatriz = r
End Function

'------------------------------------------------------------------
Private Sub CrearGraficoApilado(ws As Worksheet, hr As Long, filaTotal As Long, _
                                colIni As Long, colFin As Long, titulo As String, filaTop As Long)
    Dim rngDatos As Range
    Set rngDatos = ws.Range(ws.Cells(hr, colIni), ws.Cells(filaTotal - 1, colFin))

    Dim izq As Double: izq = ws.Cells(filaTop, 1).Left
    Dim arr As Double: arr = ws.Cells(filaTop, 1).Top
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(Left:=izq, Top:=arr, Width:=460, Height:=220)

    Dim pal(0 To 2) As Long
    pal(0) = RGB(212, 12, 12)
    pal(1) = RGB(89, 89, 89)
    pal(2) = RGB(191, 168, 128)

    With co.Chart
        .ChartType = xlColumnStacked
        .SetSourceData Source:=rngDatos, PlotBy:=xlRows
        .HasTitle = True
        .ChartTitle.Text = titulo
        .ChartTitle.Font.Size = 10
        .ChartTitle.Font.Name = "Arial"
        On Error Resume Next
        .ApplyDataLabels Type:=xlDataLabelsShowNone
        On Error GoTo 0
        .ChartGroups(1).GapWidth = 60

        Dim s As Long
        For s = 1 To .SeriesCollection.Count
            .SeriesCollection(s).Format.Fill.ForeColor.RGB = pal((s - 1) Mod 3)
            .SeriesCollection(s).Format.Line.ForeColor.RGB = RGB(255, 255, 255)
            .SeriesCollection(s).Format.Line.Weight = 0.75
        Next s

        On Error Resume Next
        .Axes(xlCategory).TickLabels.Font.Name = "Arial"
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlValue).TickLabels.Font.Name = "Arial"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Legend.Font.Name = "Arial"
        .Legend.Font.Size = 8
        .Legend.Position = xlLegendPositionBottom
        On Error GoTo 0
    End With
End Sub

'------------------------------------------------------------------
Private Function OrdenaClaves(d As Object, desc As Boolean) As String()
    Dim n As Long: n = d.Count
    Dim arr() As String
    If n = 0 Then ReDim arr(0 To 0): OrdenaClaves = arr: Exit Function
    ReDim arr(0 To n - 1)
    Dim k As Variant, i As Long: i = 0
    For Each k In d.Keys: arr(i) = CStr(k): i = i + 1: Next k
    Dim a As Long, b As Long, tmp As String
    For a = 0 To n - 2
        For b = a + 1 To n - 1
            Dim swap As Boolean
            If desc Then swap = (arr(b) > arr(a)) Else swap = (arr(b) < arr(a))
            If swap Then tmp = arr(a): arr(a) = arr(b): arr(b) = tmp
        Next b
    Next a
    OrdenaClaves = arr
End Function

'------------------------------------------------------------------
Private Function MesCorto(aaaamm As String) As String
    Dim aa As Integer, mm As Integer
    Dim s As String: s = Trim(aaaamm)
    ' Si no tiene 6 digitos numericos, devuelve el texto tal cual (no revienta)
    If Len(s) <> 6 Or Not IsNumeric(s) Then
        MesCorto = s
        Exit Function
    End If
    aa = CInt(Left(s, 4))
    mm = CInt(Mid(s, 5, 2))
    If mm < 1 Or mm > 12 Then
        MesCorto = s
        Exit Function
    End If
    MesCorto = Format(DateSerial(aa, mm, 1), "mmm-yy")
End Function
