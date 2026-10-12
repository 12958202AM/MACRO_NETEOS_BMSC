Attribute VB_Name = "Módulo1"
Sub Account_Offset_Engine_AM_v6()

    ' 1. APAGADO DE MOTORES DE EXCEL (Máxima Velocidad)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    If ActiveSheet.FilterMode Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        MsgBox "Existen filtros activos en la cédula de trabajo. Por favor, elimine los filtros antes de ejecutar el motor de compensación.", vbCritical, "Alerta de Auditoría"
        Exit Sub
    End If

    On Error GoTo ManejoError

    Dim ws As Worksheet
    Dim ultimaFila As Long, colDatos As Long, colResultado As Long
    Dim filaInicio As Long
    Dim i As Long, j As Long
    Dim valor As Double, valorTemp As Double

    Set ws = ActiveSheet

    ' ==========================================
    ' 1. SELECCIÓN DINÁMICA DE LA PRIMERA CELDA
    ' ==========================================
    Dim rangoSeleccionado As Range
    On Error Resume Next
    Set rangoSeleccionado = Application.InputBox( _
        Prompt:="Seleccione la PRIMERA CELDA que contiene el primer importe de la base (excluya el encabezado):", _
        Title:="Punto de Inicio de Auditoría", Type:=8)
    On Error GoTo ManejoError

    If rangoSeleccionado Is Nothing Then GoTo SalidaSegura

    colDatos = rangoSeleccionado.Column
    filaInicio = rangoSeleccionado.Row
    ultimaFila = ws.Cells(ws.Rows.Count, colDatos).End(xlUp).Row

    If filaInicio > ultimaFila Then
        MsgBox "No se detectaron partidas debajo de la celda seleccionada.", vbExclamation
        GoTo SalidaSegura
    End If

    ' ==========================================
    ' 2. CARGAR RAM Y LIMPIEZA INICIAL
    ' ==========================================
    Dim arrDatos As Variant
    Dim arrResultados As Variant

    arrDatos = ws.Range(ws.Cells(1, colDatos), ws.Cells(ultimaFila, colDatos)).Value
    ReDim arrResultados(1 To ultimaFila, 1 To 1)

    If filaInicio > 1 Then arrResultados(filaInicio - 1, 1) = "Clasificación"

    ' ==========================================
    ' FASE 1: CRUCE 1 A 1 DIRECTO
    ' ==========================================
    Dim dictPos As Object, dictNeg As Object
    Set dictPos = CreateObject("Scripting.Dictionary")
    Set dictNeg = CreateObject("Scripting.Dictionary")

    For i = filaInicio To ultimaFila
        If Not IsError(arrDatos(i, 1)) Then
            If Trim(CStr(arrDatos(i, 1))) <> "" And IsNumeric(Trim(CStr(arrDatos(i, 1)))) Then
                valor = Round(CDbl(Trim(CStr(arrDatos(i, 1)))), 2)
                If valor > 0 Then
                    If Not dictPos.Exists(valor) Then Set dictPos(valor) = New Collection
                    dictPos(valor).Add i
                ElseIf valor < 0 Then
                    If Not dictNeg.Exists(Abs(valor)) Then Set dictNeg(Abs(valor)) = New Collection
                    dictNeg(Abs(valor)).Add i
                End If
            End If
        End If
    Next i

    Dim key As Variant, limite As Long
    For Each key In dictPos.Keys
        If dictNeg.Exists(key) Then
            limite = Application.WorksheetFunction.Min(dictPos(key).Count, dictNeg(key).Count)
            For j = 1 To limite
                arrResultados(dictPos(key)(j), 1) = "Neto.1-1"
                arrResultados(dictNeg(key)(j), 1) = "Neto.1-1"
            Next j
        End If
    Next key

    Set dictPos = Nothing
    Set dictNeg = Nothing

    ' ==========================================
    ' FASE 2: CONCILIACIÓN POR REFERENCIA
    ' ==========================================
    Dim colDocNum As Long
    Dim dictDoc As Object
    Dim rangoDoc As Range
    Dim nombreColumnaDoc As String
    Dim arrDocs As Variant

    Do
        Application.ScreenUpdating = True
        On Error Resume Next
        Set rangoDoc = Application.InputBox( _
            Prompt:="FASE 2 (Por Referencia):" & vbCrLf & _
                    "Seleccione cualquier celda de la columna de apoyo (Ej. Factura/NIT)." & vbCrLf & vbCrLf & _
                    "-> Para omitir o finalizar esta fase, presione CANCELAR.", _
            Title:="Agrupación por Referencia Cruzada", Type:=8)
        On Error GoTo ManejoError
        Application.ScreenUpdating = False

        If rangoDoc Is Nothing Then Exit Do
        colDocNum = rangoDoc.Column

        If colDocNum > 0 Then
            nombreColumnaDoc = ""
            If filaInicio > 1 Then nombreColumnaDoc = Trim(CStr(ws.Cells(filaInicio - 1, colDocNum).Value))
            If nombreColumnaDoc = "" Then nombreColumnaDoc = Split(ws.Cells(1, colDocNum).Address, "$")(1)

            arrDocs = ws.Range(ws.Cells(1, colDocNum), ws.Cells(ultimaFila, colDocNum)).Value
            Set dictDoc = CreateObject("Scripting.Dictionary")

            For i = filaInicio To ultimaFila
                If arrResultados(i, 1) = "" Then
                    If Trim(CStr(arrDatos(i, 1))) <> "" And IsNumeric(Trim(CStr(arrDatos(i, 1)))) Then
                        Dim docKey As String
                        docKey = Trim(CStr(arrDocs(i, 1)))
                        If docKey <> "" Then
                            If Not dictDoc.Exists(docKey) Then Set dictDoc(docKey) = New Collection
                            dictDoc(docKey).Add i
                        End If
                    End If
                End If
            Next i

            Dim docItem As Variant, sumaDoc As Double, itemFila As Variant
            For Each docItem In dictDoc.Keys
                sumaDoc = 0
                For Each itemFila In dictDoc(docItem)
                    sumaDoc = sumaDoc + Round(CDbl(Trim(CStr(arrDatos(itemFila, 1)))), 2)
                Next itemFila
                If Round(sumaDoc, 2) = 0 Then
                    For Each itemFila In dictDoc(docItem)
                        arrResultados(itemFila, 1) = "Neto-" & nombreColumnaDoc
                    Next itemFila
                End If
            Next docItem

            Set dictDoc = Nothing
            Erase arrDocs
        End If
        Set rangoDoc = Nothing
    Loop

    ' ==========================================
    ' INSERCIÓN DE COLUMNA Y ORDENAMIENTO VISUAL
    ' ==========================================
    colResultado = colDatos + 1
    ws.Columns(colResultado).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Range(ws.Cells(1, colResultado), ws.Cells(ultimaFila, colResultado)).Value = arrResultados

    Dim ultimaColumna As Long, filaEncabezado As Long
    ultimaColumna = ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1
    If filaInicio > 1 Then filaEncabezado = filaInicio - 1 Else filaEncabezado = 1

    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=ws.Cells(filaInicio, colDatos), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange ws.Range(ws.Cells(filaEncabezado, 1), ws.Cells(ultimaFila, ultimaColumna))
        If filaInicio > 1 Then .Header = xlYes Else .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    arrDatos = ws.Range(ws.Cells(1, colDatos), ws.Cells(ultimaFila, colDatos)).Value
    arrResultados = ws.Range(ws.Cells(1, colResultado), ws.Cells(ultimaFila, colResultado)).Value

    ' ==========================================
    ' FASE 3: COMPENSACIÓN COMPLEJA (Fuerza Bruta)
    ' ==========================================
    Dim umbralInput As Variant
    Dim umbralMax As Double
    Dim sumaAcumulada As Double
    Dim colGrupo As Collection, indicesEnGrupo As Collection

    Application.ScreenUpdating = True
    ' Type:=1 permite teclear el número o hacer CLIC en una celda para capturar su importe
    umbralInput = Application.InputBox( _
        Prompt:="FASE 3 (Compensación Compleja):" & vbCrLf & _
                "Indique el límite de Materialidad (Importe máximo a evaluar)." & vbCrLf & _
                "Puede escribir el valor o HACER CLIC en cualquier celda de su hoja:", _
        Title:="Ajuste de Materialidad", Default:=90000, Type:=1)
    Application.ScreenUpdating = False

    If TypeName(umbralInput) <> "Boolean" Then
        ' SE APLICA VALOR ABSOLUTO PARA EVITAR ERRORES SI SE SELECCIONA UN NEGATIVO
        umbralMax = Abs(CDbl(umbralInput))

        Dim idxFiltrado() As Long, valFiltrado() As Double
        ReDim idxFiltrado(1 To ultimaFila)
        ReDim valFiltrado(1 To ultimaFila)
        Dim cantFiltrado As Long: cantFiltrado = 0

        For i = filaInicio To ultimaFila
            If arrResultados(i, 1) = "" Then
                If Trim(CStr(arrDatos(i, 1))) <> "" And IsNumeric(Trim(CStr(arrDatos(i, 1)))) Then
                    valor = Round(CDbl(Trim(CStr(arrDatos(i, 1)))), 2)
                    If Abs(valor) <= umbralMax And valor <> 0 Then
                        cantFiltrado = cantFiltrado + 1
                        idxFiltrado(cantFiltrado) = i
                        valFiltrado(cantFiltrado) = valor
                    End If
                End If
            End If
        Next i

        If cantFiltrado > 0 Then
            Dim usadoEnFase3() As Boolean
            ReDim usadoEnFase3(1 To cantFiltrado)
            Dim correlativoV4 As Long: correlativoV4 = 1

            For i = 1 To cantFiltrado
                If Not usadoEnFase3(i) And arrResultados(idxFiltrado(i), 1) = "" Then
                    Set colGrupo = New Collection
                    Set indicesEnGrupo = New Collection
                    
                    colGrupo.Add idxFiltrado(i)
                    indicesEnGrupo.Add i
                    sumaAcumulada = valFiltrado(i)

                    For j = i + 1 To cantFiltrado
                        If Not usadoEnFase3(j) And arrResultados(idxFiltrado(j), 1) = "" Then
                            valorTemp = valFiltrado(j)
                            If (sumaAcumulada > 0 And valorTemp < 0) Or (sumaAcumulada < 0 And valorTemp > 0) Then
                                sumaAcumulada = Round(sumaAcumulada + valorTemp, 2)
                                colGrupo.Add idxFiltrado(j)
                                indicesEnGrupo.Add j
                                If sumaAcumulada = 0 Then Exit For
                            End If
                        End If
                    Next j

                    If sumaAcumulada = 0 And colGrupo.Count > 1 Then
                        Dim idx As Variant
                        For Each idx In colGrupo
                            arrResultados(idx, 1) = "Neto.v4-" & correlativoV4
                        Next idx
                        Dim idxG As Variant
                        For Each idxG In indicesEnGrupo
                            usadoEnFase3(idxG) = True
                        Next idxG
                        correlativoV4 = correlativoV4 + 1
                    End If
                    Set colGrupo = Nothing
                    Set indicesEnGrupo = Nothing
                End If
            Next i
            Erase usadoEnFase3
        End If
        Erase idxFiltrado
        Erase valFiltrado
    End If

    ' ==========================================
    ' REPORTE PRELIMINAR (Corte 1)
    ' ==========================================
    Dim sumaPosBlancos As Double, sumaNegBlancos As Double
    Dim cantPosBlancos As Long, cantNegBlancos As Long
    Dim cantNeteados As Long: cantNeteados = 0

    For i = filaInicio To ultimaFila
        If Trim(CStr(arrDatos(i, 1))) <> "" And IsNumeric(Trim(CStr(arrDatos(i, 1)))) Then
            valor = Round(CDbl(Trim(CStr(arrDatos(i, 1)))), 2)
            If arrResultados(i, 1) = "" Then
                If valor > 0 Then
                    sumaPosBlancos = sumaPosBlancos + valor
                    cantPosBlancos = cantPosBlancos + 1
                ElseIf valor < 0 Then
                    sumaNegBlancos = sumaNegBlancos + valor
                    cantNegBlancos = cantNegBlancos + 1
                End If
            ElseIf arrResultados(i, 1) <> "Clasificación" Then
                cantNeteados = cantNeteados + 1
            End If
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "=== ESTADO DE PARTIDAS PREVIO A COMPENSACIÓN ===" & vbCrLf & vbCrLf & _
           "Partidas Conciliadas Exactas: " & Format(cantNeteados, "#,##0") & vbCrLf & _
           "--------------------------------------------------------" & vbCrLf & _
           "SALDOS RESIDUALES (Por revisar):" & vbCrLf & _
           "Saldos Deudores (+):  " & Format(sumaPosBlancos, "#,##0.00") & "  (" & cantPosBlancos & " partidas)" & vbCrLf & _
           "Saldos Acreedores (-): " & Format(sumaNegBlancos, "#,##0.00") & "  (" & cantNegBlancos & " partidas)" & vbCrLf & _
           "--------------------------------------------------------" & vbCrLf & _
           "Diferencia Neta Actual: " & Format(sumaPosBlancos + sumaNegBlancos, "#,##0.00"), _
           vbInformation, "Corte Preliminar de Auditoría"

    ' =========================================================================
    ' FASE 4: ABSORCIÓN ASIMÉTRICA CON TOLERANCIA (CAZA-CENTAVOS)
    ' =========================================================================
    If cantNegBlancos > 0 And cantPosBlancos > 0 Then
        Dim respAbsorcion As Integer
        respAbsorcion = MsgBox("Existen " & cantNegBlancos & " partidas acreedoras sin compensar." & vbCrLf & vbCrLf & _
                               "¿Desea activar el Motor de Absorción para liquidar estos saldos negativos contra el bolsón de saldos deudores?", _
                               vbYesNo + vbQuestion, "Validación de Absorción (1 de 2)")
        
        If respAbsorcion = vbYes Then
            Application.ScreenUpdating = False
            
            Dim colNegativos As New Collection
            Dim colPositivosUsados As New Collection
            Dim sumaAsimetrica As Double: sumaAsimetrica = 0
            
            ' Agrupar todos los negativos
            For i = filaInicio To ultimaFila
                If arrResultados(i, 1) = "" And IsNumeric(Trim(CStr(arrDatos(i, 1)))) Then
                    valor = Round(CDbl(Trim(CStr(arrDatos(i, 1)))), 2)
                    If valor < 0 Then
                        colNegativos.Add i
                        sumaAsimetrica = Round(sumaAsimetrica + valor, 2)
                    End If
                End If
            Next i
            
            ' Algoritmo Greedy: Tomar los mayores que encajen sin pasarse
            For i = filaInicio To ultimaFila
                If arrResultados(i, 1) = "" And IsNumeric(Trim(CStr(arrDatos(i, 1)))) Then
                    valor = Round(CDbl(Trim(CStr(arrDatos(i, 1)))), 2)
                    If valor > 0 Then
                        If Round(sumaAsimetrica + valor, 2) <= 0 Then
                            colPositivosUsados.Add i
                            arrResultados(i, 1) = "EnProceso" ' Marca temporal
                            sumaAsimetrica = Round(sumaAsimetrica + valor, 2)
                            If sumaAsimetrica = 0 Then Exit For
                        End If
                    End If
                End If
            Next i
            
            ' Si aún queda brecha negativa, buscar el positivo más pequeño desde abajo para cruzar el 0 por centavos
            If sumaAsimetrica < 0 Then
                Dim gapRestante As Double
                gapRestante = Abs(sumaAsimetrica)
                Dim bestIdx As Long: bestIdx = 0
                Dim bestVal As Double: bestVal = 9999999999#
                
                For i = ultimaFila To filaInicio Step -1
                    If arrResultados(i, 1) = "" And IsNumeric(Trim(CStr(arrDatos(i, 1)))) Then
                        valor = Round(CDbl(Trim(CStr(arrDatos(i, 1)))), 2)
                        If valor > 0 Then
                            If valor >= gapRestante And valor < bestVal Then
                                bestVal = valor
                                bestIdx = i
                            End If
                        End If
                    End If
                Next i
                
                If bestIdx > 0 Then
                    colPositivosUsados.Add bestIdx
                    arrResultados(bestIdx, 1) = "EnProceso"
                    sumaAsimetrica = Round(sumaAsimetrica + bestVal, 2)
                Else
                    ' Si no hay un único grande, meter los pequeños de abajo hacia arriba hasta cruzar
                    For i = ultimaFila To filaInicio Step -1
                        If arrResultados(i, 1) = "" And IsNumeric(Trim(CStr(arrDatos(i, 1)))) Then
                            valor = Round(CDbl(Trim(CStr(arrDatos(i, 1)))), 2)
                            If valor > 0 Then
                                colPositivosUsados.Add i
                                arrResultados(i, 1) = "EnProceso"
                                sumaAsimetrica = Round(sumaAsimetrica + valor, 2)
                                If sumaAsimetrica >= 0 Then Exit For
                            End If
                        End If
                    Next i
                End If
            End If
            
            Application.ScreenUpdating = True
            
            ' Validar Varianza
            If sumaAsimetrica <> 0 Then
                Dim respVarianza As Integer
                respVarianza = MsgBox("Atención: La compensación forzada ha generado una varianza técnica (diferencia de ajuste) de " & Format(sumaAsimetrica, "#,##0.00") & "." & vbCrLf & vbCrLf & _
                                      "¿Autoriza registrar este bloque de partidas bajo la referencia 'Neto.Asimetrico' para su posterior revisión?", _
                                      vbYesNo + vbExclamation, "Validación de Varianza (2 de 2)")
                
                If respVarianza = vbYes Then
                    For Each idx In colNegativos
                        arrResultados(idx, 1) = "Neto.Asimetrico"
                    Next idx
                    For Each idx In colPositivosUsados
                        arrResultados(idx, 1) = "Neto.Asimetrico"
                    Next idx
                Else
                    ' Rollback (Revertir marca temporal)
                    For Each idx In colPositivosUsados
                        arrResultados(idx, 1) = ""
                    Next idx
                End If
            Else
                ' Cuadre perfecto, se aplica directo
                For Each idx In colNegativos
                    arrResultados(idx, 1) = "Neto.Asimetrico"
                Next idx
                For Each idx In colPositivosUsados
                    arrResultados(idx, 1) = "Neto.Asimetrico"
                Next idx
            End If
            
            Set colNegativos = Nothing
            Set colPositivosUsados = Nothing
        End If
    End If

    ' ==========================================
    ' ETIQUETADO DE SALDOS RESIDUALES
    ' ==========================================
    Application.ScreenUpdating = True
    Dim respNombrar As Integer
    Dim textoBlanco As String
    Dim respSigno As String
    Dim huboEtiquetado As Boolean: huboEtiquetado = False

    respNombrar = MsgBox("¿Desea asignar una marca de auditoría a las partidas que permanecerán como Saldos Residuales (en blanco)?", vbYesNo + vbQuestion, "Etiquetar Remanentes")

    If respNombrar = vbYes Then
        textoBlanco = InputBox("Ingrese la marca o etiqueta para las partidas residuales:", "Cédula de Trabajo")
        If Trim(textoBlanco) <> "" Then
            respSigno = UCase(Trim(InputBox( _
                "¿A qué tipo de saldo desea aplicar la etiqueta '" & textoBlanco & "'?" & vbCrLf & vbCrLf & _
                "P = Solo a Saldos Deudores (+)" & vbCrLf & _
                "N = Solo a Saldos Acreedores (-)" & vbCrLf & _
                "T = A Todos (Ambos signos)", _
                "Clasificación del Saldo", "T")))

            huboEtiquetado = True
            For i = filaInicio To ultimaFila
                If arrResultados(i, 1) = "" Then
                    If Trim(CStr(arrDatos(i, 1))) <> "" And IsNumeric(Trim(CStr(arrDatos(i, 1)))) Then
                        Dim valS As Double
                        valS = Round(CDbl(Trim(CStr(arrDatos(i, 1)))), 2)
                        If (respSigno = "P" And valS > 0) Or _
                           (respSigno = "N" And valS < 0) Or _
                           (respSigno = "T" Or respSigno = "") Then
                            arrResultados(i, 1) = textoBlanco
                        End If
                    End If
                End If
            Next i
        End If
    End If

    ws.Range(ws.Cells(1, colResultado), ws.Cells(ultimaFila, colResultado)).Value = arrResultados

    ' ==========================================
    ' DICTAMEN DE CIERRE DE AUDITORÍA
    ' ==========================================
    sumaPosBlancos = 0: sumaNegBlancos = 0
    cantPosBlancos = 0: cantNegBlancos = 0
    Dim sumaPosEtiq As Double, sumaNegEtiq As Double
    Dim cantPosEtiq As Long, cantNegEtiq As Long

    For i = filaInicio To ultimaFila
        If Trim(CStr(arrDatos(i, 1))) <> "" And IsNumeric(Trim(CStr(arrDatos(i, 1)))) Then
            valor = Round(CDbl(Trim(CStr(arrDatos(i, 1)))), 2)
            If arrResultados(i, 1) = "" Then
                If valor > 0 Then
                    sumaPosBlancos = sumaPosBlancos + valor
                    cantPosBlancos = cantPosBlancos + 1
                ElseIf valor < 0 Then
                    sumaNegBlancos = sumaNegBlancos + valor
                    cantNegBlancos = cantNegBlancos + 1
                End If
            ElseIf huboEtiquetado And arrResultados(i, 1) = textoBlanco Then
                If valor > 0 Then
                    sumaPosEtiq = sumaPosEtiq + valor
                    cantPosEtiq = cantPosEtiq + 1
                ElseIf valor < 0 Then
                    sumaNegEtiq = sumaNegEtiq + valor
                    cantNegEtiq = cantNegEtiq + 1
                End If
            End If
        End If
    Next i

    Dim mensajeReporte As String
    mensajeReporte = "RESUMEN DE SALDOS RESIDUALES (Sin compensar):" & vbCrLf & _
                     "--------------------------------------------------------" & vbCrLf & _
                     "Saldos Deudores (+):  " & Format(sumaPosBlancos, "#,##0.00") & "  (" & cantPosBlancos & " partidas)" & vbCrLf & _
                     "Saldos Acreedores (-): " & Format(sumaNegBlancos, "#,##0.00") & "  (" & cantNegBlancos & " partidas)" & vbCrLf & _
                     "--------------------------------------------------------" & vbCrLf & _
                     "Diferencia Neta Restante: " & Format(sumaPosBlancos + sumaNegBlancos, "#,##0.00")

    If huboEtiquetado Then
        mensajeReporte = mensajeReporte & vbCrLf & vbCrLf & _
                         "PARTIDAS ETIQUETADAS COMO: '" & textoBlanco & "'" & vbCrLf & _
                         "--------------------------------------------------------" & vbCrLf & _
                         "Deudores Etiquetados: " & Format(sumaPosEtiq, "#,##0.00") & "  (" & cantPosEtiq & " partidas)" & vbCrLf & _
                         "Acreedores Etiquetados: " & Format(sumaNegEtiq, "#,##0.00") & "  (" & cantNegEtiq & " partidas)" & vbCrLf & _
                         "--------------------------------------------------------" & vbCrLf & _
                         "Diferencia Neta Etiquetada: " & Format(sumaPosEtiq + sumaNegEtiq, "#,##0.00")
    End If

    mensajeReporte = mensajeReporte & vbCrLf & vbCrLf & _
                     "_________________________________________________" & vbCrLf & _
                     "                                                                                A.M."

    MsgBox mensajeReporte, vbInformation, "Dictamen Final de Compensación"

SalidaSegura:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ManejoError:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "Ocurrió un error en la ejecución: " & Err.Description, vbCritical
End Sub

