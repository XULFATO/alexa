Attribute VB_Name = "Modulo_Importador"
Option Explicit

' ============================================================
'  IMPORTADOR MAESTRO / ESCLAVO  v5
'
'  ESTRUCTURA ESCLAVO (confirmada por fotos):
'    Fila 1 : labels PT   (Numero empregado interno, NIC Code, Nome, Apelido...)
'    Fila 2 : labels EN   (Employee ID, NIC Code (Client...), Name, Family Name...)
'    Fila 3 : codigos     (vacia en A/B, A002, A001, AL11, A013...)
'    Fila 4 : tipos dato  (9(7), 9(12), X(30)...)
'    Fila 5 : descripciones largas
'    Fila 6+: DATOS
'
'  MAPEO de columnas:
'    Header maestro NISS    -> columna del esclavo que tenga "NIC Code" en fila 1 (col B)
'    Header maestro CA001   -> quitar C -> A001 -> buscar en fila 3 del esclavo
'
'  HEADERS del resultado:
'    Ambas hojas usan los labels de la FILA 1 del esclavo (reordenados segun maestro)
' ============================================================

Public Sub ImportarMaestroEsclavo()

    Dim sRutaMaestro As String
    Dim sRutaEsclavo As String

    sRutaMaestro = SeleccionarFichero("Selecciona el CSV Maestro", "CSV (*.csv),*.csv")
    If sRutaMaestro = "" Then MsgBox "Cancelado.", vbInformation: Exit Sub

    sRutaEsclavo = SeleccionarFichero("Selecciona el Excel Esclavo", "Excel (*.xlsx;*.xls;*.xlsm),*.xlsx;*.xls;*.xlsm")
    If sRutaEsclavo = "" Then MsgBox "Cancelado.", vbInformation: Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrHandler

    ' ================================================================
    ' PASO 1: Leer CSV Maestro
    ' ================================================================
    Dim arrMaestro() As String
    Dim nFilasMaestro As Long
    arrMaestro = LeerCSV(sRutaMaestro, nFilasMaestro)

    Dim iHeaderMaestro As Long
    iHeaderMaestro = BuscarFilaHeader(arrMaestro, nFilasMaestro)
    If iHeaderMaestro < 0 Then
        MsgBox "No se encontro fila de headers en el Maestro.", vbCritical: GoTo Salir
    End If

    ' Limpiar BOM y parsear headers
    arrMaestro(iHeaderMaestro) = LimpiarBOM(arrMaestro(iHeaderMaestro))
    Dim arrHeaderMaestro() As String
    arrHeaderMaestro = SplitCSVLine(arrMaestro(iHeaderMaestro))
    Dim nColsMaestro As Long
    nColsMaestro = UBound(arrHeaderMaestro) + 1
    Dim hh As Long
    For hh = 0 To nColsMaestro - 1
        arrHeaderMaestro(hh) = LimpiarBOM(Trim(arrHeaderMaestro(hh)))
    Next hh

    ' ================================================================
    ' PASO 2: Abrir Excel Esclavo
    ' ================================================================
    Dim wbEsclavo As Workbook
    Set wbEsclavo = Workbooks.Open(sRutaEsclavo, ReadOnly:=True)
    Dim wsEsclavo As Worksheet
    Set wsEsclavo = wbEsclavo.Sheets(1)  ' hoja "Data"

    ' Buscar fila de codigos Axxx (normalmente fila 3)
    Dim iFilaCodigos As Long
    iFilaCodigos = BuscarFilaCodigos(wsEsclavo)
    If iFilaCodigos < 0 Then
        wbEsclavo.Close False
        MsgBox "No se encontro fila de codigos (Axxx) en el Esclavo.", vbCritical: GoTo Salir
    End If

    ' Fila 1 = labels PT (cabeceras del resultado)
    ' Fila de datos = fija en 6 (confirmado por imagen)
    Dim iFilaLabelsPT As Long
    iFilaLabelsPT = 1
    Dim iInicioData As Long
    iInicioData = 6

    Dim nColsEsclavo As Long
    nColsEsclavo = wsEsclavo.UsedRange.Columns.Count

    ' Leer arrays: codigos (fila 3), labels PT (fila 1)
    Dim arrCodEsclavo() As String    ' A002, A001, AL11... (vacio en col A y B)
    Dim arrLblPT() As String         ' Numero empregado interno, NIC Code, Nome...
    ReDim arrCodEsclavo(1 To nColsEsclavo)
    ReDim arrLblPT(1 To nColsEsclavo)
    Dim c As Long
    For c = 1 To nColsEsclavo
        arrCodEsclavo(c) = Trim(CStr(wsEsclavo.Cells(iFilaCodigos, c).Value))
        arrLblPT(c) = Trim(CStr(wsEsclavo.Cells(iFilaLabelsPT, c).Value))
    Next c

    ' Buscar columna NIC Code en fila 1 PT (col B segun imagen)
    Dim iNicCodeEsclavo As Long
    iNicCodeEsclavo = BuscarColumnaFlexible(arrLblPT, "NICCODE", nColsEsclavo)
    If iNicCodeEsclavo < 0 Then iNicCodeEsclavo = BuscarColumnaEnArrayBase1(arrLblPT, "NIC", nColsEsclavo)
    If iNicCodeEsclavo < 0 Then iNicCodeEsclavo = 2  ' fallback col B

    ' ================================================================
    ' PASO 3: Construir mapa de reordenacion
    '   NISS del maestro  -> columna NIC Code del esclavo (col B)
    '   CA001 del maestro -> A001 -> buscar en arrCodEsclavo (fila 3)
    ' ================================================================
    Dim arrMapEsclavo() As Long   ' columna en esclavo para cada col del maestro (-1 = no hay)
    Dim arrHeaderSalida() As String  ' label PT correspondiente a cada col del maestro
    ReDim arrMapEsclavo(0 To nColsMaestro - 1)
    ReDim arrHeaderSalida(0 To nColsMaestro - 1)

    Dim i As Long
    For i = 0 To nColsMaestro - 1
        Dim sCodMaestro As String
        sCodMaestro = Trim(arrHeaderMaestro(i))

        Dim iColE As Long
        If UCase(sCodMaestro) = "NISS" Then
            ' NISS maestro -> NIC Code esclavo (col B)
            iColE = iNicCodeEsclavo
        Else
            ' CA001 -> A001 -> buscar en fila codigos
            Dim sBuscado As String
            If UCase(Left(sCodMaestro, 1)) = "C" Then
                sBuscado = Mid(sCodMaestro, 2)
            Else
                sBuscado = sCodMaestro
            End If
            iColE = BuscarColumnaEnArrayBase1(arrCodEsclavo, sBuscado, nColsEsclavo)
        End If

        arrMapEsclavo(i) = iColE
        ' Header de salida = label PT de esa columna del esclavo
        If iColE > 0 Then
            arrHeaderSalida(i) = arrLblPT(iColE)
        Else
            arrHeaderSalida(i) = sCodMaestro & " [N/D]"
        End If
    Next i

    ' ================================================================
    ' PASO 4: Preparar hojas de salida
    ' ================================================================
    Dim wsM As Worksheet, wsE As Worksheet
    Set wsM = ObtenerOCrearHoja(ThisWorkbook, "MAESTRO")
    Set wsE = ObtenerOCrearHoja(ThisWorkbook, "ESCLAVO")
    wsM.Cells.ClearContents
    wsE.Cells.ClearContents
    wsM.Cells.NumberFormat = "@"
    wsE.Cells.NumberFormat = "@"

    ' Headers = labels PT del esclavo (reordenados segun maestro)
    Dim j As Long
    For j = 0 To nColsMaestro - 1
        wsM.Cells(1, j + 1).NumberFormat = "General"
        wsM.Cells(1, j + 1).Value = arrHeaderSalida(j)
        wsE.Cells(1, j + 1).NumberFormat = "General"
        wsE.Cells(1, j + 1).Value = arrHeaderSalida(j)
    Next j
    wsM.Rows(1).Font.Bold = True
    wsE.Rows(1).Font.Bold = True

    ' ================================================================
    ' PASO 5: Datos Maestro
    '   La columna NISS del maestro se vuelca tal cual (es el NIC Code equivalente)
    ' ================================================================
    Dim iFilaSalidaM As Long
    iFilaSalidaM = 2
    Dim r As Long
    For r = iHeaderMaestro + 1 To nFilasMaestro - 1
        If Len(Trim(arrMaestro(r))) = 0 Then GoTo SiguienteFilaMaestro
        Dim arrFila() As String
        arrFila = SplitCSVLine(arrMaestro(r))
        For j = 0 To nColsMaestro - 1
            If j <= UBound(arrFila) Then
                wsM.Cells(iFilaSalidaM, j + 1).Value = arrFila(j)
            End If
        Next j
        iFilaSalidaM = iFilaSalidaM + 1
SiguienteFilaMaestro:
    Next r

    ' ================================================================
    ' PASO 6: Datos Esclavo reordenado (desde fila 6)
    ' ================================================================
    Dim iUltimaFilaE As Long
    iUltimaFilaE = wsEsclavo.Cells(wsEsclavo.Rows.Count, iNicCodeEsclavo).End(xlUp).Row

    Dim iFilaSalidaE As Long
    iFilaSalidaE = 2
    Dim rE As Long
    For rE = iInicioData To iUltimaFilaE
        If Len(Trim(CStr(wsEsclavo.Cells(rE, iNicCodeEsclavo).Value))) = 0 Then GoTo SiguienteFilaEsclavo
        For j = 0 To nColsMaestro - 1
            iColE = arrMapEsclavo(j)
            If iColE > 0 Then
                wsE.Cells(iFilaSalidaE, j + 1).Value = _
                    Trim(CStr(wsEsclavo.Cells(rE, iColE).Value))
            End If
        Next j
        iFilaSalidaE = iFilaSalidaE + 1
SiguienteFilaEsclavo:
    Next rE

    wbEsclavo.Close False

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Importacion completada." & vbCrLf & _
           "  Maestro : " & iFilaSalidaM - 2 & " filas  ->  hoja MAESTRO" & vbCrLf & _
           "  Esclavo : " & iFilaSalidaE - 2 & " filas  ->  hoja ESCLAVO" & vbCrLf & vbCrLf & _
           "  Col. NIC Code  (Esclavo) : col. " & iNicCodeEsclavo & vbCrLf & _
           "  Fila codigos   (Esclavo) : fila " & iFilaCodigos & vbCrLf & _
           "  Inicio datos   (Esclavo) : fila " & iInicioData, vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
Salir:
End Sub

' ============================================================
'  HELPERS
' ============================================================

Private Function LimpiarBOM(s As String) As String
    Dim sBOM As String
    sBOM = Chr(239) & Chr(187) & Chr(191)
    If Left(s, 3) = sBOM Then s = Mid(s, 4)
    Do While Left(s, 1) = Chr(255) Or Left(s, 1) = Chr(254) Or _
             Left(s, 1) = Chr(239) Or Left(s, 1) = Chr(187) Or Left(s, 1) = Chr(191)
        s = Mid(s, 2)
    Loop
    LimpiarBOM = s
End Function

Private Function SeleccionarFichero(sTitulo As String, sFiltro As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = sTitulo
        .Filters.Clear
        .Filters.Add sFiltro, Split(sFiltro, ",")(1)
        .AllowMultiSelect = False
        If .Show = -1 Then
            SeleccionarFichero = .SelectedItems(1)
        Else
            SeleccionarFichero = ""
        End If
    End With
End Function

Private Function LeerCSV(sRuta As String, ByRef nFilas As Long) As String()
    Dim iFile As Integer
    iFile = FreeFile
    Open sRuta For Input As #iFile
    Dim sLinea As String
    Dim arr() As String
    ReDim arr(0 To 9999)
    nFilas = 0
    Do While Not EOF(iFile)
        Line Input #iFile, sLinea
        If nFilas > UBound(arr) Then ReDim Preserve arr(0 To UBound(arr) + 5000)
        arr(nFilas) = sLinea
        nFilas = nFilas + 1
    Loop
    Close #iFile
    LeerCSV = arr
End Function

Private Function SplitCSVLine(sLinea As String) As String()
    Dim arrResult() As String
    ReDim arrResult(0 To 199)
    Dim nCols As Long
    nCols = 0
    Dim sCampo As String
    sCampo = ""
    Dim bEnComillas As Boolean
    bEnComillas = False
    Dim i As Long, ch As String
    For i = 1 To Len(sLinea)
        ch = Mid(sLinea, i, 1)
        If ch = """" Then
            bEnComillas = Not bEnComillas
        ElseIf ch = ";" And Not bEnComillas Then
            If nCols > UBound(arrResult) Then ReDim Preserve arrResult(0 To UBound(arrResult) + 100)
            arrResult(nCols) = sCampo
            nCols = nCols + 1
            sCampo = ""
        Else
            sCampo = sCampo & ch
        End If
    Next i
    If nCols > UBound(arrResult) Then ReDim Preserve arrResult(0 To nCols)
    arrResult(nCols) = sCampo
    ReDim Preserve arrResult(0 To nCols)
    SplitCSVLine = arrResult
End Function

Private Function BuscarFilaHeader(arr() As String, nFilas As Long) As Long
    Dim i As Long
    For i = 0 To nFilas - 1
        If Len(Trim(arr(i))) > 0 Then
            BuscarFilaHeader = i
            Exit Function
        End If
    Next i
    BuscarFilaHeader = -1
End Function

Private Function BuscarFilaCodigos(ws As Worksheet) As Long
    ' Devuelve la fila que tiene >= 3 celdas con patron A+digitos
    Dim nCols As Long
    nCols = ws.UsedRange.Columns.Count
    Dim r As Long, c As Long
    For r = 1 To 10
        Dim nMatch As Long
        nMatch = 0
        For c = 1 To nCols
            Dim sVal As String
            sVal = Trim(CStr(ws.Cells(r, c).Value))
            If Len(sVal) >= 2 Then
                If UCase(Left(sVal, 1)) = "A" And IsNumeric(Mid(sVal, 2)) Then
                    nMatch = nMatch + 1
                End If
            End If
        Next c
        If nMatch >= 3 Then
            BuscarFilaCodigos = r
            Exit Function
        End If
    Next r
    BuscarFilaCodigos = -1
End Function

Private Function BuscarColumnaEnArray(arr() As String, sBuscar As String) As Long
    Dim i As Long
    For i = 0 To UBound(arr)
        If InStr(1, UCase(arr(i)), UCase(sBuscar)) > 0 Then
            BuscarColumnaEnArray = i
            Exit Function
        End If
    Next i
    BuscarColumnaEnArray = -1
End Function

Private Function BuscarColumnaEnArrayBase1(arr() As String, sBuscar As String, nCols As Long) As Long
    Dim i As Long
    For i = 1 To nCols
        If UCase(Trim(arr(i))) = UCase(Trim(sBuscar)) Then
            BuscarColumnaEnArrayBase1 = i
            Exit Function
        End If
    Next i
    For i = 1 To nCols
        If InStr(1, UCase(arr(i)), UCase(sBuscar)) > 0 Then
            BuscarColumnaEnArrayBase1 = i
            Exit Function
        End If
    Next i
    BuscarColumnaEnArrayBase1 = -1
End Function

Private Function BuscarColumnaFlexible(arr() As String, sBuscar As String, nCols As Long) As Long
    Dim i As Long
    For i = 1 To nCols
        If UCase(Replace(Trim(arr(i)), " ", "")) = UCase(sBuscar) Then
            BuscarColumnaFlexible = i
            Exit Function
        End If
    Next i
    BuscarColumnaFlexible = -1
End Function

Private Function ObtenerOCrearHoja(wb As Workbook, sNombre As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sNombre)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = sNombre
    End If
    Set ObtenerOCrearHoja = ws
End Function
