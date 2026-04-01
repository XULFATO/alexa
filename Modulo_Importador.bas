Attribute VB_Name = "Modulo_Importador"
Option Explicit

' ============================================================
'  IMPORTADOR MAESTRO / ESCLAVO
'  - Maestro : CSV  (headers fila buscada din�micamente)
'              campo clave: columna que contenga "NISS"
'  - Esclavo : XLSX (dos filas header buscadas din�micamente)
'              fila c�digos: A001, A002...
'              fila labels : Name, Nif...
'              campo clave: columna que contenga "NIKCODE"
'  - Output  : libro activo, dos hojas ("MAESTRO" y "ESCLAVO")
'              el esclavo se reordena seg�n columnas del maestro
'              columnas sin equivalente en maestro se descartan
' ============================================================

Public Sub ImportarMaestroEsclavo()

    Dim sRutaMaestro As String
    Dim sRutaEsclavo As String

    ' --- Seleccionar CSV Maestro ---
    sRutaMaestro = SeleccionarFichero("Selecciona el CSV Maestro", "CSV (*.csv),*.csv")
    If sRutaMaestro = "" Then MsgBox "Cancelado.", vbInformation: Exit Sub

    ' --- Seleccionar Excel Esclavo ---
    sRutaEsclavo = SeleccionarFichero("Selecciona el Excel Esclavo", "Excel (*.xlsx;*.xls;*.xlsm),*.xlsx;*.xls;*.xlsm")
    If sRutaEsclavo = "" Then MsgBox "Cancelado.", vbInformation: Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrHandler

    ' ---- PASO 1: Cargar CSV Maestro en memoria ----
    Dim arrMaestro() As String
    Dim nFilasMaestro As Long
    arrMaestro = LeerCSV(sRutaMaestro, nFilasMaestro)

    ' Buscar fila de headers en maestro (primera fila no vac�a)
    Dim iHeaderMaestro As Long
    iHeaderMaestro = BuscarFilaHeader(arrMaestro, nFilasMaestro)
    If iHeaderMaestro < 0 Then
        MsgBox "No se encontr� fila de headers en el Maestro.", vbCritical: GoTo Salir
    End If

    ' Construir diccionario de columnas maestro: c�digo (CA001) -> �ndice columna (0-based)
    Dim nColsMaestro As Long
    Dim arrHeaderMaestro() As String
    arrHeaderMaestro = Split(arrMaestro(iHeaderMaestro), ";")
    nColsMaestro = UBound(arrHeaderMaestro) + 1

    ' Buscar columna NISS en maestro
    Dim iNissMaestro As Long
    iNissMaestro = BuscarColumnaEnArray(arrHeaderMaestro, "NISS")
    If iNissMaestro < 0 Then
        MsgBox "No se encontr� columna 'NISS' en el Maestro.", vbCritical: GoTo Salir
    End If

    ' ---- PASO 2: Cargar Excel Esclavo ----
    Dim wbEsclavo As Workbook
    Set wbEsclavo = Workbooks.Open(sRutaEsclavo, ReadOnly:=True)
    Dim wsEsclavo As Worksheet
    Set wsEsclavo = wbEsclavo.Sheets(1)

    ' Buscar fila de c�digos (A001, A002...) y fila de labels
    Dim iFilaCodigos As Long, iFilaLabels As Long
    BuscarFilasEsclavo wsEsclavo, iFilaCodigos, iFilaLabels
    If iFilaCodigos < 0 Then
        wbEsclavo.Close False
        MsgBox "No se encontr� fila de c�digos (Axxx) en el Esclavo.", vbCritical: GoTo Salir
    End If

    ' Leer headers del esclavo
    Dim nColsEsclavo As Long
    nColsEsclavo = wsEsclavo.Cells(iFilaCodigos, wsEsclavo.Columns.Count).End(xlToLeft).Column
    Dim arrCodEsclavo() As String
    Dim arrLblEsclavo() As String
    ReDim arrCodEsclavo(1 To nColsEsclavo)
    ReDim arrLblEsclavo(1 To nColsEsclavo)
    Dim c As Long
    For c = 1 To nColsEsclavo
        arrCodEsclavo(c) = Trim(CStr(wsEsclavo.Cells(iFilaCodigos, c).Value))
        If iFilaLabels > 0 Then
            arrLblEsclavo(c) = Trim(CStr(wsEsclavo.Cells(iFilaLabels, c).Value))
        End If
    Next c

    ' Buscar columna NIKCODE en esclavo
    Dim iNikkodeEsclavo As Long
    iNikkodeEsclavo = BuscarColumnaEnArrayBase1(arrCodEsclavo, "NIKCODE", nColsEsclavo)
    If iNikkodeEsclavo < 0 Then
        ' Intentar tambi�n en fila labels
        iNikkodeEsclavo = BuscarColumnaEnArrayBase1(arrLblEsclavo, "NIKCODE", nColsEsclavo)
    End If
    If iNikkodeEsclavo < 0 Then
        wbEsclavo.Close False
        MsgBox "No se encontr� columna 'NIKCODE' en el Esclavo.", vbCritical: GoTo Salir
    End If

    ' ---- PASO 3: Construir mapa de reordenaci�n ----
    ' Para cada columna del maestro (CA001), quitar C -> A001, buscar en esclavo
    ' Resultado: arrMapEsclavo(i) = columna en esclavo que corresponde a columna i del maestro
    '            -1 si no existe
    Dim arrMapEsclavo() As Long
    ReDim arrMapEsclavo(0 To nColsMaestro - 1)
    Dim i As Long
    For i = 0 To nColsMaestro - 1
        Dim sCodMaestro As String
        sCodMaestro = Trim(arrHeaderMaestro(i))
        ' Quitar la C inicial: CA001 -> A001
        Dim sBuscado As String
        If Left(sCodMaestro, 1) = "C" Or Left(sCodMaestro, 1) = "c" Then
            sBuscado = Mid(sCodMaestro, 2)
        Else
            sBuscado = sCodMaestro
        End If
        arrMapEsclavo(i) = BuscarColumnaEnArrayBase1(arrCodEsclavo, sBuscado, nColsEsclavo)
    Next i

    ' ---- PASO 4: Preparar hojas de salida en el libro activo ----
    Dim wbSalida As Workbook
    Set wbSalida = ThisWorkbook

    Dim wsM As Worksheet, wsE As Worksheet
    Set wsM = ObtenerOCrearHoja(wbSalida, "MAESTRO")
    Set wsE = ObtenerOCrearHoja(wbSalida, "ESCLAVO")

    wsM.Cells.ClearContents
    wsE.Cells.ClearContents

    ' ---- PASO 5: Volcar Maestro ----
    ' Header maestro en fila 1
    Dim j As Long
    For j = 0 To nColsMaestro - 1
        wsM.Cells(1, j + 1).Value = arrHeaderMaestro(j)
    Next j

    ' Datos maestro
    Dim iFilaSalidaM As Long
    iFilaSalidaM = 2
    Dim r As Long
    For r = iHeaderMaestro + 1 To nFilasMaestro - 1
        If Len(Trim(arrMaestro(r))) = 0 Then GoTo SiguienteFilaMaestro
        Dim arrFila() As String
        arrFila = Split(arrMaestro(r), ";")
        For j = 0 To nColsMaestro - 1
            If j <= UBound(arrFila) Then
                wsM.Cells(iFilaSalidaM, j + 1).Value = arrFila(j)
            End If
        Next j
        iFilaSalidaM = iFilaSalidaM + 1
SiguienteFilaMaestro:
    Next r

    ' ---- PASO 6: Volcar Esclavo reordenado ----
    ' Header esclavo en fila 1 (labels si existen, si no c�digos)
    For j = 0 To nColsMaestro - 1
        Dim iColE As Long
        iColE = arrMapEsclavo(j)
        If iColE > 0 Then
            If iFilaLabels > 0 And Len(arrLblEsclavo(iColE)) > 0 Then
                wsE.Cells(1, j + 1).Value = arrLblEsclavo(iColE)
            Else
                wsE.Cells(1, j + 1).Value = arrCodEsclavo(iColE)
            End If
        Else
            wsE.Cells(1, j + 1).Value = arrHeaderMaestro(j) & " [N/D]"
        End If
    Next j

    ' Datos esclavo: primera fila de datos es la siguiente despu�s del mayor de iFilaCodigos/iFilaLabels
    Dim iInicioData As Long
    iInicioData = iFilaCodigos
    If iFilaLabels > iInicioData Then iInicioData = iFilaLabels
    iInicioData = iInicioData + 1

    Dim iUltimaFilaE As Long
    iUltimaFilaE = wsEsclavo.Cells(wsEsclavo.Rows.Count, iNikkodeEsclavo).End(xlUp).Row

    Dim iFilaSalidaE As Long
    iFilaSalidaE = 2
    Dim rE As Long
    For rE = iInicioData To iUltimaFilaE
        For j = 0 To nColsMaestro - 1
            iColE = arrMapEsclavo(j)
            If iColE > 0 Then
                wsE.Cells(iFilaSalidaE, j + 1).Value = wsEsclavo.Cells(rE, iColE).Value
            End If
        Next j
        iFilaSalidaE = iFilaSalidaE + 1
    Next rE

    wbEsclavo.Close False

    ' ---- Formato m�nimo ----
    With wsM.Rows(1).Font: .Bold = True: End With
    With wsE.Rows(1).Font: .Bold = True: End With

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Importaci�n completada." & vbCrLf & _
           "  Maestro : " & iFilaSalidaM - 2 & " filas volcadas en hoja MAESTRO" & vbCrLf & _
           "  Esclavo : " & iFilaSalidaE - 2 & " filas volcadas en hoja ESCLAVO", vbInformation
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
    Dim sContenido As String
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

Private Function BuscarFilaHeader(arr() As String, nFilas As Long) As Long
    ' Devuelve el �ndice (0-based) de la primera fila no vac�a
    Dim i As Long
    For i = 0 To nFilas - 1
        If Len(Trim(arr(i))) > 0 Then
            BuscarFilaHeader = i
            Exit Function
        End If
    Next i
    BuscarFilaHeader = -1
End Function

Private Function BuscarColumnaEnArray(arr() As String, sBuscar As String) As Long
    ' B�squeda 0-based, case-insensitive, coincidencia parcial
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
    ' B�squeda 1-based, case-insensitive, coincidencia exacta primero, luego parcial
    Dim i As Long
    ' Exacta
    For i = 1 To nCols
        If UCase(Trim(arr(i))) = UCase(Trim(sBuscar)) Then
            BuscarColumnaEnArrayBase1 = i
            Exit Function
        End If
    Next i
    ' Parcial
    For i = 1 To nCols
        If InStr(1, UCase(arr(i)), UCase(sBuscar)) > 0 Then
            BuscarColumnaEnArrayBase1 = i
            Exit Function
        End If
    Next i
    BuscarColumnaEnArrayBase1 = -1
End Function

Private Sub BuscarFilasEsclavo(ws As Worksheet, ByRef iFilaCodigos As Long, ByRef iFilaLabels As Long)
    ' Busca la fila que contiene c�digos tipo Axxx (patr�n: letra A + d�gitos)
    ' y la fila adyacente con labels de texto
    iFilaCodigos = -1
    iFilaLabels = -1
    Dim r As Long
    Dim nCols As Long
    nCols = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If nCols < 2 Then nCols = ws.UsedRange.Columns.Count

    For r = 1 To 10  ' Buscar en las primeras 10 filas
        Dim nMatchCodigos As Long
        nMatchCodigos = 0
        Dim c As Long
        For c = 1 To nCols
            Dim sVal As String
            sVal = Trim(CStr(ws.Cells(r, c).Value))
            ' Patr�n: empieza por A seguido de d�gitos (A001, A002...)
            If Len(sVal) >= 2 Then
                If UCase(Left(sVal, 1)) = "A" And IsNumeric(Mid(sVal, 2)) Then
                    nMatchCodigos = nMatchCodigos + 1
                End If
            End If
        Next c
        If nMatchCodigos >= 3 Then  ' al menos 3 coincidencias para confirmar
            iFilaCodigos = r
            ' La fila de labels es la adyacente (anterior o posterior) con texto no num�rico
            If r > 1 Then
                iFilaLabels = r - 1
            Else
                iFilaLabels = r + 1
            End If
            Exit For
        End If
    Next r
End Sub

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
