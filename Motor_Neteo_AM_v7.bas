Attribute VB_Name = "Motor_Neteo_AM_v7"
'==================================================================================================
' MOTOR DE NETEO Y CLASIFICACION DE PARTIDAS  -  A.M.  v7
'--------------------------------------------------------------------------------------------------
' Cruza partidas de uno o varios mayores contables puestos en UNA sola hoja (por ejemplo el gasto
' 45902 y el pasivo 24399), del criterio mas solido al mas debil y todo en memoria (RAM):
'
'   Nivel REF  : la glosa cita el numero de otro comprobante (reversos: "REV CPTE 250725340").
'   Niveles 2+ : combinaciones de columnas (llave) + importe. Ej.: NRO_TRANSACCION + GLOSA.
'                En cada nivel: cruce 1 a 1, grupos N a N que suman 0 y subconjuntos del grupo.
'   Final      : opcional, Fuerza bruta (FB) o Absorcion asimetrica (ASI).
'   Sin cruce  : Pagos / Ajuste / Otras Provisiones/Reversiones (nombres editables).
'
' El universo se separa por cuenta (primer digito: 2 Pasivo, 4 Gasto) y signo (Debe + / Haber -):
'   - Cruce entre cuentas distintas   -> Provisiones/Reversiones  (ID PREV-0001)
'   - Cruce dentro de la misma cuenta -> Neto                     (ID NET-0001)
'
' Resultado: 4 columnas -> Clasificacion | Detalle Clasificacion | ID Cruce | Criterio
' No reordena la hoja y no necesita columnas auxiliares (ABS, CONCAT, BUSCARV).
'
' Ejecutar la macro: Account_Offset_Engine_AM_v7   (Windows: usa Scripting.Dictionary)
'==================================================================================================
Option Explicit

'------------------------------------ Parametros ajustables ---------------------------------------
Private Const TITULO As String = "Motor de Neteo AM v7"
Private Const PFX_PREV As String = "PREV"
Private Const PFX_NET As String = "NET"
Private Const PFX_FB As String = "FB"
Private Const PFX_ASI As String = "ASI"
Private Const CAT_PREV As String = "Provisiones/Reversiones"
Private Const CAT_NET As String = "Neto"
Private Const CAT_ASI As String = "Neto asimétrico"
Private Const CAT_OTRAS As String = "Otras Provisiones/Reversiones"
Private Const DEF_RES_POS As String = "Pagos"
Private Const DEF_RES_NEG As String = "Ajuste"
Private Const DEF_BAL_POS As String = "Pagos"
Private Const DEF_BAL_NEG As String = "Otras Provisiones/Reversiones"
Private Const CRIT_SIN As String = "Sin cruce"
Private Const CRIT_INVALIDO As String = "Importe no válido"
Private Const HDR_CAT As String = "Clasificación"
Private Const HDR_DET As String = "Detalle Clasificación"
Private Const HDR_ID As String = "ID Cruce"
Private Const HDR_CRIT As String = "Criterio"
Private Const MAX_NAN_GROUP As Long = 500       ' tamano maximo de un grupo N a N
Private Const MAX_SUB_GROUP As Long = 40        ' tamano maximo de grupo para buscar subconjuntos
Private Const SUB_K As Long = 4                 ' subconjuntos dentro del grupo: 1 contra hasta 3
Private Const MIN_REF_DIGITS As Long = 5        ' digitos minimos de un numero de comprobante en la glosa
Private Const DFS_BUDGET As Long = 200000       ' combinaciones maximas evaluadas por partida objetivo
Private Const FB_TIME_LIMIT As Double = 120     ' segundos maximos para la fuerza bruta
Private Const TIE_MAX As Long = 50              ' candidatos revisados para desempatar por glosa
Private Const REG_APP As String = "MotorNeteoAM"
Private Const ACC_FMT As String = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"

'------------------------------------ Estado del motor (RAM) --------------------------------------
Private ws As Worksheet
Private gN As Long, gR1 As Long, gR2 As Long, gHdrRow As Long
Private gAmtCol1 As Long, gAmtCol2 As Long
Private gAccCol As Long, gDigits As Long, gGlosaCol As Long
Private gAmt() As Currency          ' importe con signo (Debe + / Haber -), exacto a 2 decimales
Private gIncl() As Boolean          ' fila incluida (visible, si se eligio "solo visibles")
Private gOk() As Boolean            ' participa: incluida, importe valido y distinto de 0
Private gBad() As Boolean           ' importe con error o texto no numerico
Private gUsed() As Boolean          ' ya cruzada
Private gAcc() As String            ' digitos de la cuenta contable
Private gGi() As Long               ' grupo de cuenta (1..gG)
Private gGrpName() As String
Private gG As Long
Private gCat() As String, gDet() As String, gCrit() As String, gPfx() As String
Private gNum() As Long
Private gCnt As Object              ' prefijo -> ultimo numero de ID
Private gRaw As Object              ' columna -> valores crudos (Value2)
Private gNorm As Object             ' columna -> valores normalizados para llaves
Private gWords() As Variant         ' palabras de la glosa (desempate de la fuerza bruta)
Private gGlosaNorm As Variant
Private gMap(0 To 65535) As Integer ' mapa de normalizacion de caracteres
Private gMapReady As Boolean
Private gLvCount As Long
Private gLvType() As Long           ' 1 = referencia en glosa, 2 = combinacion de columnas
Private gLvCols() As Variant
Private gLvGlosa() As Long
Private gLvName() As String
Private gLvCross() As Long, gLvItems() As Long
Private gLvExtra() As String
Private sVal() As Currency, sIdx() As Long, sTaken() As Boolean, sChosen() As Long
Private sM As Long, sK As Long, sNodes As Long, sSolLen As Long
Private sRun As Object
Private gTimeLimited As Boolean, gTimeStart As Double, gTimedOut As Boolean
Private gProcSecs As Double, gClock As Double
Private gNText As Long, gNErr As Long, gNEmpty As Long, gNZero As Long, gNHidden As Long, gNFixed As Long
Private gFinalMsg As String

'==================================================================================================
'  PROCEDIMIENTO PRINCIPAL
'==================================================================================================
Public Sub Account_Offset_Engine_AM_v7()
    Dim oldCalc As Long, oldEvt As Boolean, saved As Boolean
    Dim nErr As Long, dErr As String

    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Active la hoja que contiene el mayor contable y vuelva a ejecutar el motor.", vbExclamation, TITULO
        Exit Sub
    End If

    On Error GoTo ManejoError
    Set ws = ActiveSheet
    If ws.ProtectContents Then
        MsgBox "La hoja está protegida. Desprotéjala antes de ejecutar el motor.", vbExclamation, TITULO
        Exit Sub
    End If

    oldCalc = Application.Calculation
    oldEvt = Application.EnableEvents
    saved = True
    Application.EnableCancelKey = xlErrorHandler
    ResetState
    InitCharMap

    ' 1. Columnas de importe y cuenta, bloque de datos y filas visibles
    If Not AskAmountColumns() Then GoTo SalidaSegura
    If Not AskAccountColumn() Then GoTo SalidaSegura
    If Not DetectDataBlock() Then GoTo SalidaSegura
    If Not AskVisibility() Then GoTo SalidaSegura

    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ClockStart
    LoadData
    ClockStop
    If CountOk() = 0 Then
        Application.ScreenUpdating = True
        MsgBox "No se encontraron importes válidos (distintos de cero) en la base seleccionada.", vbExclamation, TITULO
        GoTo SalidaSegura
    End If

    ' 2. Niveles de cruce por llave
    Application.ScreenUpdating = True
    If Not ConfigureLevels() Then GoTo SalidaSegura
    Application.ScreenUpdating = False
    ClockStart
    RunLevels
    ClockStop

    ' 3. Corte preliminar y opcion final (fuerza bruta / absorcion asimetrica)
    Application.ScreenUpdating = True
    ShowPreliminaryReport
    FinalOption

    ' 4. Partidas sin cruce
    Application.ScreenUpdating = True
    ResidualLabels

    ' 5. Escritura de resultados y dictamen
    If Not WriteResults() Then GoTo SalidaSegura
    Application.ScreenUpdating = True
    ShowFinalReport

SalidaSegura:
    On Error Resume Next
    Application.StatusBar = False
    If saved Then
        Application.Calculation = oldCalc
        Application.EnableEvents = oldEvt
    End If
    Application.ScreenUpdating = True
    Application.EnableCancelKey = xlInterrupt
    ReleaseState
    Exit Sub

ManejoError:
    nErr = Err.Number
    dErr = Err.Description
    Application.ScreenUpdating = True
    If nErr = 18 Then
        MsgBox "Proceso interrumpido por el usuario (Esc). No se escribieron resultados.", vbExclamation, TITULO
    Else
        MsgBox "Ocurrió un error en la ejecución:" & vbCrLf & dErr & " (" & nErr & ")", vbCritical, TITULO
    End If
    Resume SalidaSegura
End Sub

'==================================================================================================
'  1. CONFIGURACION: IMPORTE, CUENTA, BLOQUE DE DATOS Y FILAS VISIBLES
'==================================================================================================
Private Function AskAmountColumns() As Boolean
    Dim r As Range, cols As Variant, defAddr As String, sugRow As Long, msg As String
    Dim cSaldo As Long, cDebe As Long, cHaber As Long, tmp As Long

    sugRow = ScanHeaderRow()
    If sugRow > 0 Then
        cSaldo = FindHdrCol(sugRow, "SALDO", "")
        If cSaldo = 0 Then cSaldo = FindHdrCol(sugRow, "IMPORTE", "")
        If cSaldo = 0 Then cSaldo = FindHdrCol(sugRow, "MONTO", "")
        cDebe = FindHdrCol(sugRow, "DEBE", "")
        cHaber = FindHdrCol(sugRow, "HABER", "")
        If cSaldo > 0 Then
            defAddr = ws.Cells(sugRow, cSaldo).Address
        ElseIf cDebe > 0 And cHaber > 0 Then
            defAddr = ws.Cells(sugRow, cDebe).Address & Application.International(xlListSeparator) & _
                      ws.Cells(sugRow, cHaber).Address
        End If
    End If

    msg = "PASO 1/3 - IMPORTE" & vbLf & _
          "Clic en el ENCABEZADO del saldo / importe." & vbLf & _
          "Si hay DEBE y HABER: Ctrl+clic en los dos (primero DEBE), aunque no estén juntos." & vbLf & _
          "Los números guardados como texto se convertirán a número."
    Do
        Set r = AskRange(msg, defAddr)
        If r Is Nothing Then Exit Function
        cols = RangeColumns(r)
        If IsArray(cols) Then
            If UBound(cols) = 0 Then
                gAmtCol1 = cols(0): gAmtCol2 = 0
                Exit Do
            ElseIf UBound(cols) = 1 Then
                gAmtCol1 = cols(0): gAmtCol2 = cols(1)
                Exit Do
            End If
        End If
        MsgBox "Seleccione UN encabezado (saldo) o DOS (Debe y Haber).", vbExclamation, TITULO
    Loop
    gHdrRow = r.Areas(1).Row

    ' Si se marco primero el HABER, se invierte el orden
    If gAmtCol2 > 0 Then
        If IsHaberHeader(HdrKey(gAmtCol1)) And Not IsHaberHeader(HdrKey(gAmtCol2)) Then
            tmp = gAmtCol1: gAmtCol1 = gAmtCol2: gAmtCol2 = tmp
        End If
    End If
    AskAmountColumns = True
End Function

Private Function AskAccountColumn() As Boolean
    Dim r As Range, defAddr As String, v As Variant, c As Long

    c = FindHdrCol(gHdrRow, "CUENTACONTABLE", "")
    If c = 0 Then c = FindHdrCol(gHdrRow, "CUENTA", "")
    If c > 0 Then defAddr = ws.Cells(gHdrRow, c).Address

    Set r = AskRange("PASO 2/3 - CUENTA CONTABLE" & vbLf & _
                     "Clic en el ENCABEZADO de la cuenta (separa, por ejemplo, gasto 4 de pasivo 2)." & vbLf & _
                     "Cancelar = no separar por cuenta.", defAddr)
    gDigits = 1
    If r Is Nothing Then
        gAccCol = 0
    Else
        gAccCol = r.Areas(1).Column
        v = Application.InputBox(Prompt:="PASO 3/3 - GRUPO DE CUENTA" & vbLf & _
                                 "¿Cuántos dígitos iniciales de la cuenta forman el grupo?" & vbLf & _
                                 "1 = primer dígito (2 Pasivo / 4 Gasto)", _
                                 Title:=TITULO, Default:=1, Type:=1)
        If VarType(v) <> vbBoolean Then
            gDigits = CLng(v)
            If gDigits < 1 Then gDigits = 1
            If gDigits > 15 Then gDigits = 15
        End If
    End If
    AskAccountColumn = True
End Function

Private Function DetectDataBlock() As Boolean
    Dim lastUsed As Long, lr As Long, i As Long, n As Long, r2 As Long, below As Long
    Dim a1 As Variant, a2 As Variant, a3 As Variant, msg As String, resp As VbMsgBoxResult, r As Range

    lastUsed = LastRowOf(gAmtCol1)
    If gAmtCol2 > 0 Then
        lr = LastRowOf(gAmtCol2)
        If lr > lastUsed Then lastUsed = lr
    End If
    If gAccCol > 0 Then
        lr = LastRowOf(gAccCol)
        If lr > lastUsed Then lastUsed = lr
    End If
    If lastUsed <= gHdrRow Then
        MsgBox "No hay datos debajo de la fila de encabezados (" & gHdrRow & ").", vbExclamation, TITULO
        Exit Function
    End If

    ' El bloque termina en la primera fila vacia (debajo suele haber totales, leyendas o tablas)
    n = lastUsed - gHdrRow
    a1 = ReadColumn(gAmtCol1, gHdrRow + 1, lastUsed)
    If gAmtCol2 > 0 Then a2 = ReadColumn(gAmtCol2, gHdrRow + 1, lastUsed)
    If gAccCol > 0 Then a3 = ReadColumn(gAccCol, gHdrRow + 1, lastUsed)
    r2 = lastUsed
    For i = 1 To n
        If IsBlankV(a1(i, 1)) Then
            If IsBlankAt(a2, i) Then
                If IsBlankAt(a3, i) Then
                    r2 = gHdrRow + i - 1
                    Exit For
                End If
            End If
        End If
    Next
    ' Filas finales de total (SUMA / SUBTOTALES sin cuenta)
    Do While r2 > gHdrRow + 1
        If IsTotalRow(r2) Then
            r2 = r2 - 1
        Else
            Exit Do
        End If
    Loop
    If r2 <= gHdrRow Then
        MsgBox "No se detectaron partidas debajo de los encabezados.", vbExclamation, TITULO
        Exit Function
    End If
    gR1 = gHdrRow + 1: gR2 = r2: gN = gR2 - gR1 + 1

    With ws.UsedRange
        below = .Row + .Rows.Count - 1
    End With
    If lastUsed > below Then below = lastUsed

    msg = "BASE DE DATOS DETECTADA" & vbCrLf & vbCrLf & _
          "Encabezados: fila " & gHdrRow & vbCrLf & _
          "Datos: filas " & FmtN(gR1) & " a " & FmtN(gR2) & " (" & FmtN(gN) & " filas)" & vbCrLf & _
          "Importe: " & AmountDesc() & vbCrLf & "Cuenta: "
    If gAccCol > 0 Then
        msg = msg & HdrName(gAccCol) & " (grupo = " & gDigits & IIf(gDigits = 1, " dígito)", " dígitos)")
    Else
        msg = msg & "(no se separa por cuenta)"
    End If
    If below > gR2 Then
        msg = msg & vbCrLf & vbCrLf & "Debajo hay contenido hasta la fila " & FmtN(below) & _
              " (totales, leyendas o tablas dinámicas) que NO se incluirá."
    End If
    msg = msg & vbCrLf & vbCrLf & "¿Es correcto?" & vbCrLf & "Sí = continuar" & vbCrLf & _
          "No = indicar la última fila de datos" & vbCrLf & "Cancelar = salir"
    resp = MsgBox(msg, vbYesNoCancel + vbQuestion, TITULO)
    If resp = vbCancel Then Exit Function
    If resp = vbNo Then
        Set r = AskRange("Clic en una celda de la ÚLTIMA fila de datos.", ws.Cells(gR2, gAmtCol1).Address)
        If r Is Nothing Then Exit Function
        If r.Row <= gHdrRow Then
            MsgBox "La última fila debe estar debajo de los encabezados.", vbExclamation, TITULO
            Exit Function
        End If
        gR2 = r.Row: gN = gR2 - gR1 + 1
    End If
    DetectDataBlock = True
End Function

Private Function AskVisibility() As Boolean
    Dim vis As Range, ar As Range, nVis As Long, i As Long, r As Long, resp As VbMsgBoxResult

    ReDim gIncl(1 To gN)
    For i = 1 To gN
        gIncl(i) = True
    Next
    gNHidden = 0
    If gN > 1 Then
        On Error Resume Next
        Set vis = ws.Range(ws.Cells(gR1, gAmtCol1), ws.Cells(gR2, gAmtCol1)).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        If Not vis Is Nothing Then nVis = vis.Cells.Count
        If nVis < gN Then
            resp = MsgBox("Hay " & FmtN(gN - nVis) & " filas ocultas o filtradas dentro de la base." & vbCrLf & vbCrLf & _
                          "Sí = procesar SOLO las filas visibles" & vbCrLf & _
                          "No = procesar TODAS las filas" & vbCrLf & "Cancelar = salir", _
                          vbYesNoCancel + vbQuestion, TITULO)
            If resp = vbCancel Then Exit Function
            If resp = vbYes Then
                For i = 1 To gN
                    gIncl(i) = False
                Next
                If Not vis Is Nothing Then
                    For Each ar In vis.Areas
                        For r = ar.Row To ar.Row + ar.Rows.Count - 1
                            gIncl(r - gR1 + 1) = True
                        Next
                    Next
                End If
                gNHidden = gN - nVis
            End If
        End If
    End If
    AskVisibility = True
End Function

'==================================================================================================
'  CARGA EN MEMORIA
'==================================================================================================
Private Sub LoadData()
    Dim a1 As Variant, a2 As Variant, a3 As Variant, i As Long, k1 As Long, k2 As Long
    Dim ok1 As Boolean, ok2 As Boolean, v1 As Currency, v2 As Currency, dG As Object, key As String, ng As Long

    ReDim gAmt(1 To gN): ReDim gOk(1 To gN): ReDim gBad(1 To gN): ReDim gUsed(1 To gN)
    ReDim gAcc(1 To gN): ReDim gGi(1 To gN)
    ReDim gCat(1 To gN): ReDim gDet(1 To gN): ReDim gCrit(1 To gN): ReDim gPfx(1 To gN): ReDim gNum(1 To gN)
    ReDim gWords(1 To gN)

    a1 = RawCol(gAmtCol1)
    If gAmtCol2 > 0 Then a2 = RawCol(gAmtCol2)
    If gAccCol > 0 Then a3 = RawCol(gAccCol)
    Set dG = NewDict()
    ng = 0

    For i = 1 To gN
        ok1 = ParseCell(a1(i, 1), v1, k1)
        ok2 = False: v2 = 0: k2 = 3
        If gAmtCol2 > 0 Then ok2 = ParseCell(a2(i, 1), v2, k2)
        If gIncl(i) Then
            If k1 = 1 Then gNText = gNText + 1
            If k2 = 1 Then gNText = gNText + 1
        End If
        If k1 = 2 Or k1 = 4 Or k2 = 2 Or k2 = 4 Then
            gBad(i) = True
            If gIncl(i) Then gNErr = gNErr + 1
        ElseIf ok1 Or ok2 Then
            If Not ok1 Then v1 = 0
            If Not ok2 Then v2 = 0
            gAmt(i) = v1 - v2
            If gAmt(i) = 0 Then
                If gIncl(i) Then gNZero = gNZero + 1
            Else
                gOk(i) = gIncl(i)
            End If
        Else
            If gIncl(i) Then gNEmpty = gNEmpty + 1
        End If

        If gAccCol > 0 Then gAcc(i) = AccDigits(a3(i, 1))
        key = Left$(gAcc(i), gDigits)
        If dG.Exists(key) Then
            gGi(i) = dG.Item(key)
        Else
            ng = ng + 1
            ReDim Preserve gGrpName(1 To ng)
            gGrpName(ng) = key
            dG.Add key, ng
            gGi(i) = ng
        End If
    Next
    gG = ng
    Set dG = Nothing
End Sub

'==================================================================================================
'  2. NIVELES DE CRUCE
'==================================================================================================
Private Function ConfigureLevels() As Boolean
    Dim cG As Long, cT As Long, cD As Long, cTp As Long, cMo As Long, cU As Long, cF As Long, cTarget As Long
    Dim resp As VbMsgBoxResult, k As Long, nDef As Long, fromSaved As Boolean, r As Range, cols As Variant
    Dim msg As String, keep() As Boolean

    ' Columnas estandar del mayor (se reconocen por el nombre del encabezado)
    cG = FindHdrCol(gHdrRow, "GLOSA", "")
    If cG = 0 Then cG = FindHdrCol(gHdrRow, "DESCRIPCION", "")
    If cG = 0 Then cG = FindHdrCol(gHdrRow, "CONCEPTO", "")
    cT = FindHdrCol(gHdrRow, "TRANSACC", "TIPO")
    If cT = 0 Then cT = FindHdrCol(gHdrRow, "COMPROBANTE", "TIPO|FECHA")
    If cT = 0 Then cT = FindHdrCol(gHdrRow, "ASIENTO", "TIPO|FECHA")
    cD = FindHdrCol(gHdrRow, "DOCUMENTO", "TIPO|FECHA")
    cTp = FindHdrCol(gHdrRow, "TIPODOCUMENTO", "")
    cMo = FindHdrCol(gHdrRow, "MODULO", "")
    cU = FindHdrCol(gHdrRow, "NOMBREUSUARIO", "")
    If cU = 0 Then cU = FindHdrCol(gHdrRow, "USUARIO", "CODIGO|CARGO|CI")
    cF = FindHdrCol(gHdrRow, "FECHADOCUMENTO", "")
    If cF = 0 Then cF = FindHdrCol(gHdrRow, "FECHA", "PROCESO|HORA")
    gGlosaCol = cG

    ' Candidatos: la configuracion de la ultima ejecucion (si existe) y los niveles estandar
    gLvCount = 0
    fromSaved = LoadSavedLevels()
    nDef = gLvCount
    cTarget = cT
    If cTarget = 0 Then cTarget = cD
    If cG > 0 And cTarget > 0 Then AddLevel 1, Array(cTarget), cG
    If cT > 0 And cG > 0 Then AddLevel 2, Array(cT, cG), 0
    If cMo > 0 And cTp > 0 And cD > 0 Then AddLevel 2, Array(cMo, cTp, cD), 0
    If cU > 0 And cF > 0 Then AddLevel 2, Array(cU, cF), 0
    If Not fromSaved Then nDef = gLvCount

    If gLvCount > 0 Then
        msg = "NIVELES DE CRUCE PROPUESTOS"
        If fromSaved Then msg = msg & " (última configuración usada)"
        msg = msg & vbCrLf & "Sí = usar estos niveles  |  No = elegir uno por uno  |  Cancelar = salir" & vbCrLf & vbCrLf & _
              "Del más sólido al más débil:" & vbCrLf & LevelList(1, nDef) & vbCrLf & _
              "En cada nivel: cruce 1 a 1 (llave + importe), grupos N a N que suman 0 y " & _
              "subconjuntos dentro del grupo."
        resp = MsgBox(msg, vbYesNoCancel + vbQuestion, TITULO)
        If resp = vbCancel Then Exit Function
        ReDim keep(1 To gLvCount)
        If resp = vbYes Then
            For k = 1 To nDef
                keep(k) = True
            Next
        Else
            For k = 1 To gLvCount
                resp = MsgBox("¿Usar este nivel?" & vbCrLf & vbCrLf & gLvName(k) & vbCrLf & vbCrLf & _
                              "Sí = usar   No = omitir   Cancelar = terminar la selección", _
                              vbYesNoCancel + vbQuestion, TITULO & " - nivel " & k & " de " & gLvCount)
                If resp = vbCancel Then Exit For
                keep(k) = (resp = vbYes)
            Next
        End If
        KeepLevels keep
    Else
        MsgBox "No se reconocieron columnas estándar (GLOSA, NRO_TRANSACCION...)." & vbCrLf & _
               "Indique sus propias combinaciones de columnas.", vbInformation, TITULO
    End If

    ' Combinaciones adicionales elegidas por el usuario
    Do
        msg = "¿Desea AGREGAR una combinación de columnas como nivel adicional?" & vbCrLf & vbCrLf
        If gLvCount > 0 Then
            msg = msg & "Niveles actuales:" & vbCrLf & LevelList(1, gLvCount)
        Else
            msg = msg & "(todavía no hay niveles)"
        End If
        resp = MsgBox(msg, vbYesNo + vbQuestion, TITULO)
        If resp <> vbYes Then Exit Do
        Set r = AskRange("NIVEL ADICIONAL" & vbLf & "Ctrl+clic en los ENCABEZADOS de las columnas a combinar." & vbLf & _
                         "Ej.: NRO_DOCUMENTO + NOMBRE_USUARIO" & vbLf & "(el importe se agrega solo)")
        If Not r Is Nothing Then
            cols = RangeColumns(r)
            If IsArray(cols) Then AddLevel 2, cols, 0
        End If
    Loop

    If gLvCount = 0 Then
        resp = MsgBox("No hay niveles por llave." & vbCrLf & _
                      "¿Continuar solo con la opción final (fuerza bruta / absorción asimétrica)?", _
                      vbYesNo + vbQuestion, TITULO)
        If resp <> vbYes Then Exit Function
    End If
    If gGlosaCol = 0 Then
        For k = 1 To gLvCount
            If gLvType(k) = 1 Then
                gGlosaCol = gLvGlosa(k)
                Exit For
            End If
        Next
    End If
    SaveLevels
    ConfigureLevels = True
End Function

Private Sub RunLevels()
    Dim k As Long, before As Long, c11 As Long, cNaN As Long, cSub As Long

    If gLvCount = 0 Then Exit Sub
    ReDim gLvCross(1 To gLvCount): ReDim gLvItems(1 To gLvCount): ReDim gLvExtra(1 To gLvCount)
    For k = 1 To gLvCount
        Application.StatusBar = TITULO & " - nivel " & k & " de " & gLvCount & ": " & Left$(gLvName(k), 80)
        before = CountUsed()
        If gLvType(k) = 1 Then
            RunRefLevel gLvGlosa(k), gLvCols(k), c11, "Ref. en " & HdrName(gLvGlosa(k)) & " -> " & JoinHdr(gLvCols(k))
            gLvCross(k) = c11
        Else
            RunKeyLevel gLvCols(k), gLvName(k), c11, cNaN, cSub
            gLvCross(k) = c11 + cNaN + cSub
            If cNaN + cSub > 0 Then gLvExtra(k) = " (N a N: " & cNaN & ", subconj.: " & cSub & ")"
        End If
        gLvItems(k) = CountUsed() - before
    Next
    Application.StatusBar = False
End Sub

' Nivel de referencia: la glosa de una partida cita el numero de comprobante de otra con importe
' exactamente opuesto (reversos y regularizaciones). Prefiere la misma cuenta (reverso).
Private Sub RunRefLevel(ByVal glosaCol As Long, ByVal targets As Variant, ByRef made As Long, ByVal crit As String)
    Dim raw As Variant, d As Object, tvals() As Variant, nt As Long, t As Long, i As Long, j As Long
    Dim e As Long, ne As Long, entRow() As Long, entNx() As Long, s As String, tok As Variant
    Dim k As Long, best As Long, bestX As Long, own As Boolean

    made = 0
    raw = RawCol(glosaCol)
    nt = UBound(targets) - LBound(targets) + 1
    ReDim tvals(1 To nt)
    For t = 1 To nt
        tvals(t) = NormCol(CLng(targets(LBound(targets) + t - 1)))
    Next

    ' Indice: numero de comprobante -> filas (lista enlazada en orden de fila)
    Set d = NewDict()
    ReDim entRow(1 To gN * nt): ReDim entNx(1 To gN * nt)
    For t = 1 To nt
        For i = gN To 1 Step -1
            If gOk(i) Then
                s = tvals(t)(i)
                If Len(s) > 0 Then
                    If IsAllDigits(s) Then
                        ne = ne + 1: entRow(ne) = i
                        If d.Exists(s) Then
                            entNx(ne) = d.Item(s)
                            d.Item(s) = ne
                        Else
                            d.Add s, ne
                        End If
                    End If
                End If
            End If
        Next
    Next
    If ne = 0 Then Exit Sub

    For i = 1 To gN
        If gOk(i) Then
            If Not gUsed(i) Then
                If VarType(raw(i, 1)) = vbString Then tok = DigitTokens(CStr(raw(i, 1)), MIN_REF_DIGITS) Else tok = Empty
                If IsArray(tok) Then
                    For k = 1 To UBound(tok)
                        own = False
                        For t = 1 To nt
                            If tvals(t)(i) = tok(k) Then own = True
                        Next
                        If Not own Then
                            If d.Exists(tok(k)) Then
                                best = 0: bestX = 0
                                e = d.Item(tok(k))
                                Do While e <> 0
                                    j = entRow(e)
                                    If j <> i Then
                                        If Not gUsed(j) Then
                                            If gAmt(j) = -gAmt(i) Then
                                                If gGi(j) = gGi(i) Then
                                                    best = j
                                                    Exit Do
                                                End If
                                                If bestX = 0 Then bestX = j
                                            End If
                                        End If
                                    End If
                                    e = entNx(e)
                                Loop
                                If best = 0 Then best = bestX
                                If best > 0 Then
                                    LabelPair i, best, "", crit
                                    made = made + 1
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        End If
    Next
    Set d = Nothing
End Sub

' Nivel por combinacion de columnas: (a) 1 a 1 con llave + importe, (b) grupos N a N con la misma
' llave que suman 0, (c) subconjuntos que suman 0 dentro de los grupos que no cuadran completos.
Private Sub RunKeyLevel(ByVal cols As Variant, ByVal nm As String, ByRef c11 As Long, ByRef cNaN As Long, ByRef cSub As Long)
    Dim nc As Long, p As Long, i As Long, b As Long, nb As Long, g As Long, ng As Long, cnt As Long
    Dim s As String, bk As String, anyVal As Boolean
    Dim ca() As Variant, keys() As String, d As Object
    Dim head() As Long, nxt() As Long, gh() As Long, gnx() As Long, gc() As Long
    Dim gs() As Currency, gp() As Boolean, gneg() As Boolean, mem() As Long

    c11 = 0: cNaN = 0: cSub = 0
    nc = UBound(cols) - LBound(cols) + 1
    ReDim ca(1 To nc)
    For p = 1 To nc
        ca(p) = NormCol(CLng(cols(LBound(cols) + p - 1)))
    Next
    ReDim keys(1 To gN)
    For i = 1 To gN
        If gOk(i) Then
            If Not gUsed(i) Then
                s = vbNullString: anyVal = False
                For p = 1 To nc
                    If Len(ca(p)(i)) > 0 Then anyVal = True
                    s = s & ca(p)(i) & ChrW$(31)
                Next
                If anyVal Then keys(i) = s
            End If
        End If
    Next

    ' (a) cruce 1 a 1: misma llave y mismo importe absoluto, signos opuestos
    Set d = NewDict()
    ReDim head(1 To gN): ReDim nxt(1 To gN)
    For i = gN To 1 Step -1
        If Len(keys(i)) > 0 Then
            bk = keys(i) & AmtKey(gAmt(i))
            If d.Exists(bk) Then
                b = d.Item(bk)
            Else
                nb = nb + 1: b = nb
                d.Add bk, b
            End If
            nxt(i) = head(b): head(b) = i
        End If
    Next
    For b = 1 To nb
        If nxt(head(b)) <> 0 Then PairChain head(b), nxt, "", nm, c11
    Next
    Set d = Nothing
    Erase head, nxt

    ' (b) y (c) grupos con la misma llave
    Set d = NewDict()
    ReDim gh(1 To gN): ReDim gnx(1 To gN): ReDim gc(1 To gN)
    ReDim gs(1 To gN): ReDim gp(1 To gN): ReDim gneg(1 To gN)
    For i = gN To 1 Step -1
        If Len(keys(i)) > 0 Then
            If Not gUsed(i) Then
                If d.Exists(keys(i)) Then
                    g = d.Item(keys(i))
                Else
                    ng = ng + 1: g = ng
                    d.Add keys(i), g
                End If
                gnx(i) = gh(g): gh(g) = i
                gc(g) = gc(g) + 1
                gs(g) = gs(g) + gAmt(i)
                If gAmt(i) > 0 Then
                    gp(g) = True
                Else
                    gneg(g) = True
                End If
            End If
        End If
    Next
    For g = 1 To ng
        If gc(g) >= 3 Then
            If gp(g) And gneg(g) Then
                If gs(g) = 0 And gc(g) <= MAX_NAN_GROUP Then
                    ChainToArray gh(g), gnx, mem, cnt
                    If cnt >= 2 Then
                        LabelSet mem, cnt, "", nm & " (N a N)"
                        cNaN = cNaN + 1
                    End If
                ElseIf gc(g) <= MAX_SUB_GROUP Then
                    ChainToArray gh(g), gnx, mem, cnt
                    If cnt >= 3 Then SubsetSearch mem, cnt, SUB_K, False, 0, "", nm & " (subconjunto)", cSub, False
                End If
            End If
        End If
    Next
    Set d = Nothing
End Sub

' Empareja positivos y negativos de una misma cadena (bucket). Prefiere la contrapartida en OTRA
' cuenta (gasto contra pasivo); si no hay, dentro de la misma cuenta. Orden de fila estable.
Private Sub PairChain(ByVal first As Long, nxt() As Long, ByVal pfx As String, ByVal crit As String, ByRef made As Long)
    Dim i As Long, j As Long, k As Long, g As Long, gi As Long, nP As Long, nNg As Long
    Dim posL() As Long, negL() As Long, srt() As Long, cntG() As Long, ptr() As Long, segE() As Long, fillP() As Long

    i = first
    Do While i <> 0
        If Not gUsed(i) Then
            If gAmt(i) > 0 Then
                nP = nP + 1
            Else
                nNg = nNg + 1
            End If
        End If
        i = nxt(i)
    Loop
    If nP = 0 Or nNg = 0 Then Exit Sub

    ReDim posL(1 To nP): ReDim negL(1 To nNg): ReDim srt(1 To nNg)
    ReDim cntG(1 To gG): ReDim ptr(1 To gG): ReDim segE(1 To gG): ReDim fillP(1 To gG)
    nP = 0: nNg = 0
    i = first
    Do While i <> 0
        If Not gUsed(i) Then
            If gAmt(i) > 0 Then
                nP = nP + 1: posL(nP) = i
            Else
                nNg = nNg + 1: negL(nNg) = i
                cntG(gGi(i)) = cntG(gGi(i)) + 1
            End If
        End If
        i = nxt(i)
    Loop
    ' negativos ordenados por grupo de cuenta (estable)
    k = 1
    For g = 1 To gG
        ptr(g) = k: fillP(g) = k
        k = k + cntG(g)
        segE(g) = k - 1
    Next
    For j = 1 To nNg
        g = gGi(negL(j))
        srt(fillP(g)) = negL(j)
        fillP(g) = fillP(g) + 1
    Next
    For k = 1 To nP
        i = posL(k): gi = gGi(i): j = 0
        For g = 1 To gG
            If g <> gi Then
                If ptr(g) <= segE(g) Then
                    j = srt(ptr(g)): ptr(g) = ptr(g) + 1
                    Exit For
                End If
            End If
        Next
        If j = 0 Then
            If ptr(gi) <= segE(gi) Then
                j = srt(ptr(gi)): ptr(gi) = ptr(gi) + 1
            End If
        End If
        If j > 0 Then
            LabelPair i, j, pfx, crit
            made = made + 1
        End If
    Next
End Sub

'==================================================================================================
'  BUSQUEDA DE COMBINACIONES QUE SUMAN 0 (pares y 1 contra varios)
'==================================================================================================
Private Sub SubsetSearch(items() As Long, ByVal cnt As Long, ByVal K As Long, ByVal doPairs As Boolean, _
                         ByVal thr As Currency, ByVal pfx As String, ByVal crit As String, _
                         ByRef found As Long, ByVal useSim As Boolean)
    Dim live() As Long, nl As Long, i As Long, t As Long, q As Long, p As Long
    Dim nPos As Long, nNeg As Long, sg As Long, nt As Long, tg() As Long, tv() As Currency
    Dim mem() As Long, key As String, m As Long

    If cnt < 2 Then Exit Sub
    ReDim live(1 To cnt)
    For i = 1 To cnt
        t = items(i)
        If gOk(t) Then
            If Not gUsed(t) Then
                If thr = 0 Or Abs(gAmt(t)) <= thr Then
                    nl = nl + 1: live(nl) = t
                End If
            End If
        End If
    Next
    If nl < 2 Then Exit Sub
    If doPairs Then PairsByAmount live, nl, pfx, crit, found, useSim
    If K < 3 Then Exit Sub

    ' 1 contra varios: los objetivos son las partidas del signo minoritario
    For i = 1 To nl
        t = live(i)
        If Not gUsed(t) Then
            If gAmt(t) > 0 Then
                nPos = nPos + 1
            Else
                nNeg = nNeg + 1
            End If
        End If
    Next
    If nPos = 0 Or nNeg = 0 Then Exit Sub
    If nPos <= nNeg Then sg = 1 Else sg = -1
    ReDim tg(1 To nl): ReDim tv(1 To nl): ReDim sIdx(1 To nl): ReDim sVal(1 To nl)
    nt = 0: m = 0
    For i = 1 To nl
        t = live(i)
        If Not gUsed(t) Then
            If Sgn(gAmt(t)) = sg Then
                nt = nt + 1: tg(nt) = t: tv(nt) = Abs(gAmt(t))
            Else
                m = m + 1: sIdx(m) = t: sVal(m) = Abs(gAmt(t))
            End If
        End If
    Next
    sM = m
    If sM < 2 Or nt = 0 Then Exit Sub
    QSortDesc sVal, sIdx, 1, sM
    If nt > 1 Then QSortDesc tv, tg, 1, nt
    Set sRun = NewDict()
    For p = 1 To sM
        key = AmtKey(sVal(p))
        If Not sRun.Exists(key) Then sRun.Add key, p
    Next
    ReDim sTaken(1 To sM): ReDim sChosen(1 To K)
    sK = K
    For q = 1 To nt
        If TimeUp() Then Exit For
        If (q Mod 200) = 0 And gTimeLimited Then
            Application.StatusBar = TITULO & " - fuerza bruta: " & FmtN(q) & " de " & FmtN(nt)
        End If
        t = tg(q)
        If Not gUsed(t) Then
            sNodes = 0: sSolLen = 0
            If Dfs(1, tv(q), 1) Then
                ReDim mem(1 To sSolLen + 1)
                mem(1) = t
                For p = 1 To sSolLen
                    mem(p + 1) = sIdx(sChosen(p))
                    sTaken(sChosen(p)) = True
                Next
                LabelSet mem, sSolLen + 1, pfx, crit
                found = found + 1
            End If
        End If
    Next
    Set sRun = Nothing
End Sub

' Busca entre los candidatos (ordenados de mayor a menor) 2..K-1 partidas que sumen remv.
Private Function Dfs(ByVal start As Long, ByVal remv As Currency, ByVal depth As Long) As Boolean
    Dim p As Long, lo As Long, hi As Long, md As Long, v As Currency, key As String

    sNodes = sNodes + 1
    If sNodes > DFS_BUDGET Then Exit Function
    If (sNodes And 4095) = 0 Then
        If TimeUp() Then
            sNodes = DFS_BUDGET + 1
            Exit Function
        End If
    End If
    If start > sM Then Exit Function

    If depth >= sK - 1 Then
        ' ultima partida: busqueda directa del importe exacto
        key = AmtKey(remv)
        If sRun.Exists(key) Then
            p = sRun.Item(key)
            If p < start Then p = start
            Do While p <= sM
                If sVal(p) <> remv Then Exit Do
                If Not sTaken(p) Then
                    sChosen(depth) = p: sSolLen = depth
                    Dfs = True
                    Exit Function
                End If
                p = p + 1
            Loop
        End If
        Exit Function
    End If

    ' primer candidato con valor <= remv (busqueda binaria)
    lo = start: hi = sM + 1
    Do While lo < hi
        md = (lo + hi) \ 2
        If sVal(md) > remv Then
            lo = md + 1
        Else
            hi = md
        End If
    Loop
    For p = lo To sM
        If Not sTaken(p) Then
            v = sVal(p)
            If v * (sK - depth) < remv Then Exit For
            If v = remv Then
                If depth >= 2 Then
                    sChosen(depth) = p: sSolLen = depth
                    Dfs = True
                    Exit Function
                End If
            Else
                sChosen(depth) = p
                If Dfs(p + 1, remv - v, depth + 1) Then
                    Dfs = True
                    Exit Function
                End If
                If sNodes > DFS_BUDGET Then Exit Function
            End If
        End If
    Next
End Function

' Pares 1 a 1 por importe exacto. Si hay varios candidatos del mismo importe, elige el de glosa
' mas parecida (palabras en comun) cuando useSim = True.
Private Sub PairsByAmount(live() As Long, ByVal nl As Long, ByVal pfx As String, ByVal crit As String, _
                          ByRef found As Long, ByVal useSim As Boolean)
    Dim d As Object, k As Long, t As Long, b As Long, nb As Long, key As String
    Dim bHead() As Long, nx() As Long, cand As Long, best As Long, bestSim As Long, sm As Long, seen As Long

    Set d = NewDict()
    ReDim bHead(1 To nl): ReDim nx(1 To nl)
    For k = nl To 1 Step -1
        t = live(k)
        If Not gUsed(t) Then
            If gAmt(t) < 0 Then
                key = AmtKey(gAmt(t))
                If d.Exists(key) Then
                    b = d.Item(key)
                Else
                    nb = nb + 1: b = nb
                    d.Add key, b
                End If
                nx(k) = bHead(b): bHead(b) = k
            End If
        End If
    Next
    If nb = 0 Then Exit Sub
    For k = 1 To nl
        If TimeUp() Then Exit For
        t = live(k)
        If Not gUsed(t) Then
            If gAmt(t) > 0 Then
                key = AmtKey(gAmt(t))
                If d.Exists(key) Then
                    b = d.Item(key)
                    Do While bHead(b) <> 0
                        If gUsed(live(bHead(b))) Then
                            bHead(b) = nx(bHead(b))
                        Else
                            Exit Do
                        End If
                    Loop
                    If bHead(b) <> 0 Then
                        best = bHead(b)
                        If useSim And gGlosaCol > 0 Then
                            bestSim = Similarity(t, live(best)): seen = 1
                            cand = nx(best)
                            Do While cand <> 0
                                If seen >= TIE_MAX Then Exit Do
                                If Not gUsed(live(cand)) Then
                                    sm = Similarity(t, live(cand)): seen = seen + 1
                                    If sm > bestSim Then
                                        best = cand: bestSim = sm
                                    End If
                                End If
                                cand = nx(cand)
                            Loop
                        End If
                        LabelPair t, live(best), pfx, crit
                        found = found + 1
                    End If
                End If
            End If
        End If
    Next
    Set d = Nothing
End Sub

'==================================================================================================
'  3. OPCION FINAL: FUERZA BRUTA (FB) O ABSORCION ASIMETRICA (ASI)
'==================================================================================================
Private Sub FinalOption()
    Dim v As Variant, opt As Long, thr As Variant, kv As Variant, K As Long, r As Range, grpCols As Variant
    Dim found As Long, resp As VbMsgBoxResult, nPend As Long, thrC As Currency

    nPend = CountPending()
    If nPend < 2 Then Exit Sub
    v = Application.InputBox(Prompt:="OPCIÓN FINAL (" & FmtN(nPend) & " partidas pendientes)" & vbLf & _
                             "1 = Fuerza bruta (FB): combinaciones que suman 0" & vbLf & _
                             "2 = Absorción asimétrica (ASI): negativos contra positivos" & vbLf & _
                             "0 = Ninguna", Title:=TITULO, Default:=0, Type:=1)
    If VarType(v) = vbBoolean Then Exit Sub
    opt = CLng(v)

    If opt = 1 Then
        thr = Application.InputBox(Prompt:="FUERZA BRUTA - Materialidad" & vbLf & _
                                   "Importe máximo por partida (escriba o haga clic en una celda)." & vbLf & _
                                   "0 = sin límite.", Title:=TITULO, Default:=RegGet("FB_Umbral", "90000"), Type:=1)
        If VarType(thr) = vbBoolean Then Exit Sub
        kv = Application.InputBox(Prompt:="FUERZA BRUTA - Partidas por combinación" & vbLf & _
                                  "2 = solo pares 1 a 1 (recomendado)" & vbLf & _
                                  "3 a 5 = una contra varias (más lento y con más coincidencias casuales)", _
                                  Title:=TITULO, Default:=RegGet("FB_K", "2"), Type:=1)
        If VarType(kv) = vbBoolean Then Exit Sub
        K = CLng(kv)
        If K < 2 Then K = 2
        If K > 6 Then K = 6
        resp = MsgBox("¿Restringir la fuerza bruta a partidas que compartan una columna?" & vbCrLf & _
                      "(ej. la misma CUENTA_CONTABLE o el mismo proveedor)" & vbCrLf & vbCrLf & _
                      "Sí = elegir columna(s)" & vbCrLf & "No = toda la base pendiente", vbYesNo + vbQuestion, TITULO)
        If resp = vbYes Then
            Set r = AskRange("Ctrl+clic en el/los ENCABEZADOS que deben coincidir.")
            If Not r Is Nothing Then grpCols = RangeColumns(r)
        End If
        If Abs(CDbl(thr)) >= 9E+11 Then
            thrC = 0
        Else
            thrC = CCur(Abs(CDbl(thr)))
        End If
        RegSet "FB_Umbral", CStr(thrC)
        RegSet "FB_K", CStr(K)
        Application.ScreenUpdating = False
        ClockStart
        RunBruteForce thrC, K, grpCols, found
        ClockStop
        Application.StatusBar = False
        Application.ScreenUpdating = True
        gFinalMsg = "Fuerza bruta: " & FmtN(found) & " cruces"
        If gTimedOut Then gFinalMsg = gFinalMsg & " (se alcanzó el tiempo máximo de " & FB_TIME_LIMIT & " s)"
    ElseIf opt = 2 Then
        RunAsymmetric
    End If
End Sub

Private Sub RunBruteForce(ByVal thr As Currency, ByVal K As Long, ByVal grpCols As Variant, ByRef found As Long)
    Dim crit As String, pend() As Long, n As Long, i As Long, nc As Long, p As Long, g As Long, ng As Long
    Dim ca() As Variant, s As String, d As Object, gh() As Long, gnx() As Long, mem() As Long, cnt As Long

    gTimeLimited = True: gTimeStart = Timer: gTimedOut = False
    If gGlosaCol > 0 Then gGlosaNorm = NormCol(gGlosaCol)
    If IsArray(grpCols) Then
        crit = "Fuerza bruta por " & JoinHdr(grpCols)
        nc = UBound(grpCols) - LBound(grpCols) + 1
        ReDim ca(1 To nc)
        For p = 1 To nc
            ca(p) = NormCol(CLng(grpCols(LBound(grpCols) + p - 1)))
        Next
        Set d = NewDict()
        ReDim gh(1 To gN): ReDim gnx(1 To gN)
        For i = gN To 1 Step -1
            If gOk(i) Then
                If Not gUsed(i) Then
                    s = vbNullString
                    For p = 1 To nc
                        s = s & ca(p)(i) & ChrW$(31)
                    Next
                    If d.Exists(s) Then
                        g = d.Item(s)
                    Else
                        ng = ng + 1: g = ng
                        d.Add s, g
                    End If
                    gnx(i) = gh(g): gh(g) = i
                End If
            End If
        Next
        For g = 1 To ng
            If TimeUp() Then Exit For
            ChainToArray gh(g), gnx, mem, cnt
            If cnt >= 2 Then SubsetSearch mem, cnt, K, True, thr, PFX_FB, crit, found, True
        Next
        Set d = Nothing
    Else
        crit = "Fuerza bruta (solo importe)"
        ReDim pend(1 To gN)
        For i = 1 To gN
            If gOk(i) Then
                If Not gUsed(i) Then
                    n = n + 1: pend(n) = i
                End If
            End If
        Next
        If n >= 2 Then SubsetSearch pend, n, K, True, thr, PFX_FB, crit, found, True
    End If
    gTimeLimited = False
End Sub

' Absorcion asimetrica (como la Fase 4 de la v6), por grupo de cuenta: todos los negativos
' pendientes contra los positivos mayores que caben, mas el menor que cruza el 0. Pide autorizacion.
Private Sub RunAsymmetric()
    Dim g As Long, i As Long, k As Long, nNeg As Long, nPos As Long, nTook As Long, cnt As Long, best As Long
    Dim sumNeg As Currency, acc As Currency, gap As Currency, sumTook As Currency
    Dim pos() As Long, pv() As Currency, took() As Boolean, mem() As Long
    Dim msg As String, resp As VbMsgBoxResult, nDone As Long, catAsi As String

    For g = 1 To gG
        nNeg = 0: nPos = 0: sumNeg = 0
        For i = 1 To gN
            If gOk(i) Then
                If Not gUsed(i) Then
                    If gGi(i) = g Then
                        If gAmt(i) < 0 Then
                            nNeg = nNeg + 1: sumNeg = sumNeg + gAmt(i)
                        Else
                            nPos = nPos + 1
                        End If
                    End If
                End If
            End If
        Next
        If nNeg > 0 And nPos > 0 Then
            ReDim pos(1 To nPos): ReDim pv(1 To nPos): ReDim took(1 To nPos)
            k = 0
            For i = 1 To gN
                If gOk(i) Then
                    If Not gUsed(i) Then
                        If gGi(i) = g And gAmt(i) > 0 Then
                            k = k + 1: pos(k) = i: pv(k) = gAmt(i)
                        End If
                    End If
                End If
            Next
            If nPos > 1 Then QSortDesc pv, pos, 1, nPos
            acc = sumNeg: nTook = 0: sumTook = 0
            For k = 1 To nPos
                If acc + pv(k) <= 0 Then
                    took(k) = True: acc = acc + pv(k)
                    nTook = nTook + 1: sumTook = sumTook + pv(k)
                End If
                If acc = 0 Then Exit For
            Next
            If acc < 0 Then
                gap = -acc: best = 0
                For k = nPos To 1 Step -1
                    If Not took(k) Then
                        If pv(k) >= gap Then
                            best = k
                            Exit For
                        End If
                    End If
                Next
                If best > 0 Then
                    took(best) = True: acc = acc + pv(best)
                    nTook = nTook + 1: sumTook = sumTook + pv(best)
                End If
            End If
            If nTook > 0 Then
                msg = "ABSORCIÓN ASIMÉTRICA - " & GrpTag(g) & vbCrLf & vbCrLf & _
                      "Acreedores (-) a liquidar: " & FmtN(nNeg) & " partidas   Bs " & FmtM(sumNeg) & vbCrLf & _
                      "Deudores (+) utilizados: " & FmtN(nTook) & " partidas   Bs " & FmtM(sumTook) & vbCrLf & _
                      "Varianza técnica: Bs " & FmtM(acc) & vbCrLf & vbCrLf & _
                      "¿Registrar este bloque con ID ASI (" & IIf(acc = 0, CAT_NET, CAT_ASI) & ")?"
                resp = MsgBox(msg, vbYesNo + IIf(acc = 0, vbQuestion, vbExclamation), TITULO)
                If resp = vbYes Then
                    ReDim mem(1 To nNeg + nTook)
                    cnt = 0
                    For i = 1 To gN
                        If gOk(i) Then
                            If Not gUsed(i) Then
                                If gGi(i) = g And gAmt(i) < 0 Then
                                    cnt = cnt + 1: mem(cnt) = i
                                End If
                            End If
                        End If
                    Next
                    For k = 1 To nPos
                        If took(k) Then
                            cnt = cnt + 1: mem(cnt) = pos(k)
                        End If
                    Next
                    If acc = 0 Then catAsi = CAT_NET Else catAsi = CAT_ASI
                    LabelSet mem, cnt, PFX_ASI, "Absorción asimétrica (varianza Bs " & FmtM(acc) & ")", catAsi
                    nDone = nDone + 1
                End If
            End If
        End If
    Next
    gFinalMsg = "Absorción asimétrica: " & nDone & " bloque(s) registrado(s)"
End Sub

'==================================================================================================
'  4. PARTIDAS SIN CRUCE
'==================================================================================================
Private Sub ResidualLabels()
    Dim lb() As String, nm() As String, cnts() As Long, i As Long, k As Long
    Dim resp As VbMsgBoxResult, v As String, msg As String

    ReDim lb(1 To 4): ReDim nm(1 To 4): ReDim cnts(1 To 4)
    lb(1) = RegGet("Etq_ResPos", DEF_RES_POS): lb(2) = RegGet("Etq_ResNeg", DEF_RES_NEG)
    lb(3) = RegGet("Etq_BalPos", DEF_BAL_POS): lb(4) = RegGet("Etq_BalNeg", DEF_BAL_NEG)
    nm(1) = "Gasto / Resultado (+)": nm(2) = "Gasto / Resultado (-)"
    nm(3) = "Pasivo / Balance (+)": nm(4) = "Pasivo / Balance (-)"
    For i = 1 To gN
        If gOk(i) Then
            If Not gUsed(i) Then
                k = ResIdx(i)
                cnts(k) = cnts(k) + 1
            End If
        End If
    Next
    If cnts(1) + cnts(2) + cnts(3) + cnts(4) > 0 Then
        msg = "PARTIDAS SIN CRUCE - nombre que recibirán:" & vbCrLf & vbCrLf
        For k = 1 To 4
            If cnts(k) > 0 Then msg = msg & nm(k) & ": " & FmtN(cnts(k)) & " partidas -> " & lb(k) & vbCrLf
        Next
        If gGlosaCol > 0 And gAccCol > 0 Then
            msg = msg & vbCrLf & "En Pasivo / Balance, si la glosa cita otra cuenta de gasto (ej. 45299101), " & _
                  "se clasifica como '" & CAT_OTRAS & " 452'." & vbCrLf
        End If
        msg = msg & vbCrLf & "Sí = usar estos nombres" & vbCrLf & "No = cambiarlos"
        resp = MsgBox(msg, vbYesNo + vbQuestion, TITULO)
        If resp = vbNo Then
            For k = 1 To 4
                If cnts(k) > 0 Then
                    v = InputBox("Nombre para las partidas sin cruce de " & nm(k) & ":", TITULO, lb(k))
                    If Len(Trim$(v)) > 0 Then lb(k) = Trim$(v)
                End If
            Next
            RegSet "Etq_ResPos", lb(1): RegSet "Etq_ResNeg", lb(2)
            RegSet "Etq_BalPos", lb(3): RegSet "Etq_BalNeg", lb(4)
        End If
    End If
    Application.ScreenUpdating = False
    ClockStart
    LabelResiduals lb
    ClockStop
End Sub

Private Sub LabelResiduals(lb() As String)
    Dim i As Long, k As Long, raw As Variant, resPfx As String, accLen As Long, other As String, f As String

    ' prefijos (3 digitos) de las cuentas de resultado presentes en la base, ej. |459|
    For i = 1 To gN
        If gOk(i) Then
            If Len(gAcc(i)) > 0 Then
                If accLen = 0 Then accLen = Len(gAcc(i))
                f = Left$(gAcc(i), 1)
                If f = "4" Or f = "5" Then
                    f = Left$(gAcc(i), 3)
                    If InStr(1, resPfx, "|" & f & "|") = 0 Then resPfx = resPfx & "|" & f & "|"
                End If
            End If
        End If
    Next
    If gGlosaCol > 0 Then raw = RawCol(gGlosaCol)
    For i = 1 To gN
        If gOk(i) Then
            If Not gUsed(i) Then
                k = ResIdx(i)
                gCat(i) = lb(k): gDet(i) = lb(k): gCrit(i) = CRIT_SIN
                If k >= 3 And gGlosaCol > 0 And accLen > 0 Then
                    other = CitedOtherAccounts(raw(i, 1), resPfx, accLen)
                    If Len(other) > 0 Then
                        gCat(i) = CAT_OTRAS
                        gDet(i) = CAT_OTRAS & " " & other
                    End If
                End If
            End If
        ElseIf gIncl(i) And gBad(i) Then
            gCrit(i) = CRIT_INVALIDO
        End If
    Next
End Sub

Private Function ResIdx(ByVal i As Long) As Long
    Dim f As String
    f = Left$(gAcc(i), 1)
    If Len(f) = 0 Or f = "4" Or f = "5" Then
        If gAmt(i) > 0 Then ResIdx = 1 Else ResIdx = 2
    Else
        If gAmt(i) > 0 Then ResIdx = 3 Else ResIdx = 4
    End If
End Function

' Cuentas de resultado citadas en la glosa que NO estan en la base (ej. 45299101 -> "452").
Private Function CitedOtherAccounts(ByVal v As Variant, ByVal resPfx As String, ByVal accLen As Long) As String
    Dim tok As Variant, k As Long, p3 As String, res As String, f As String

    If VarType(v) <> vbString Then Exit Function
    tok = DigitTokens(CStr(v), accLen)
    If Not IsArray(tok) Then Exit Function
    For k = 1 To UBound(tok)
        If Len(tok(k)) = accLen Then
            f = Left$(tok(k), 1)
            If f = "4" Or f = "5" Then
                p3 = Left$(tok(k), 3)
                If InStr(1, resPfx, "|" & p3 & "|") = 0 Then
                    If InStr(1, "/" & res & "/", "/" & p3 & "/") = 0 Then
                        If Len(res) > 0 Then res = res & "/"
                        res = res & p3
                    End If
                End If
            End If
        End If
    Next
    CitedOtherAccounts = res
End Function

'==================================================================================================
'  5. ESCRITURA DE RESULTADOS
'==================================================================================================
Private Function WriteResults() As Boolean
    Dim outCol As Long, exist As Long, resp As VbMsgBoxResult, r As Range, lastHdr As Long
    Dim i As Long, k As Long, w As Long, maxNum As Long, fmtId As String, ins As Boolean
    Dim outArr() As Variant, prevVals As Variant, af As Range, afFirst As Long, afLast As Long, afRows As Long

    Application.ScreenUpdating = True
    exist = FindOutputColumns()
    If exist > 0 Then
        resp = MsgBox("Ya existen columnas de resultado de una ejecución anterior (" & ColLetter(exist) & ":" & _
                      ColLetter(exist + 3) & ")." & vbCrLf & vbCrLf & "Sí = sobrescribirlas" & vbCrLf & _
                      "No = crear columnas nuevas" & vbCrLf & "Cancelar = salir sin escribir", _
                      vbYesNoCancel + vbQuestion, TITULO)
        If resp = vbCancel Then Exit Function
        If resp = vbYes Then outCol = exist
    End If
    If outCol = 0 Then
        resp = MsgBox("¿Dónde ubicar las 4 columnas de resultado?" & vbCrLf & "(" & HDR_CAT & " | " & HDR_DET & _
                      " | " & HDR_ID & " | " & HDR_CRIT & ")" & vbCrLf & vbCrLf & _
                      "Sí = al FINAL de la tabla" & vbCrLf & _
                      "No = a la DERECHA de una columna que usted elija" & vbCrLf & _
                      "Cancelar = salir sin escribir", vbYesNoCancel + vbQuestion, TITULO)
        If resp = vbCancel Then Exit Function
        If resp = vbYes Then
            lastHdr = ws.Cells(gHdrRow, ws.Columns.Count).End(xlToLeft).Column
            outCol = lastHdr + 1
            If outCol + 3 > ws.Columns.Count Then
                MsgBox "No hay columnas libres a la derecha de la tabla.", vbExclamation, TITULO
                Exit Function
            End If
            If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(gHdrRow, outCol), ws.Cells(gR2, outCol + 3))) > 0 Then ins = True
        Else
            Set r = AskRange("Clic en una celda de la columna a cuya DERECHA irán las 4 columnas nuevas.", _
                             ws.Cells(gHdrRow, gAmtCol1).Address)
            If r Is Nothing Then Exit Function
            outCol = r.Areas(1).Column + 1
            ins = True
        End If
    End If

    Application.ScreenUpdating = False
    ClockStart
    ' formato contable y conversion de numeros-texto (antes de insertar columnas)
    gNFixed = FixAmountColumn(gAmtCol1)
    If gAmtCol2 > 0 Then gNFixed = gNFixed + FixAmountColumn(gAmtCol2)
    If ins Then ws.Columns(outCol).Resize(, 4).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    For i = 1 To gN
        If gNum(i) > maxNum Then maxNum = gNum(i)
    Next
    w = Len(CStr(maxNum))
    If w < 4 Then w = 4
    fmtId = String$(w, "0")
    ReDim outArr(1 To gN, 1 To 4)
    If outCol = exist And gNHidden > 0 Then prevVals = ws.Cells(gR1, outCol).Resize(gN, 4).Value
    For i = 1 To gN
        If gIncl(i) Then
            If Len(gCat(i)) > 0 Then outArr(i, 1) = gCat(i)
            If Len(gDet(i)) > 0 Then outArr(i, 2) = gDet(i)
            If gNum(i) > 0 Then outArr(i, 3) = gPfx(i) & "-" & Format$(gNum(i), fmtId)
            If Len(gCrit(i)) > 0 Then outArr(i, 4) = gCrit(i)
        ElseIf IsArray(prevVals) Then
            For k = 1 To 4
                outArr(i, k) = prevVals(i, k)
            Next
        End If
    Next
    ws.Cells(gHdrRow, outCol).Resize(1, 4).Value = Array(HDR_CAT, HDR_DET, HDR_ID, HDR_CRIT)
    ws.Cells(gR1, outCol).Resize(gN, 4).Value = outArr

    ' formato de encabezados, ancho de columnas y autofiltro (no criticos)
    On Error Resume Next
    If outCol <> exist And outCol > 1 Then
        ws.Cells(gHdrRow, outCol - 1).Copy
        ws.Cells(gHdrRow, outCol).Resize(1, 4).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
    End If
    ws.Range(ws.Cells(gHdrRow, outCol), ws.Cells(Application.WorksheetFunction.Min(gR2, gHdrRow + 500), outCol + 3)).Columns.AutoFit
    If ws.AutoFilterMode And Not ws.FilterMode Then
        Set af = ws.AutoFilter.Range
        If af.Row = gHdrRow Then
            afFirst = af.Column: afLast = af.Column + af.Columns.Count - 1
            If outCol + 3 > afLast And outCol <= afLast + 1 Then
                afRows = af.Row + af.Rows.Count - 1
                If afRows < gR2 Then afRows = gR2
                ws.AutoFilterMode = False
                ws.Range(ws.Cells(gHdrRow, afFirst), ws.Cells(afRows, outCol + 3)).AutoFilter
            End If
        End If
    End If
    On Error GoTo 0
    ClockStop
    WriteResults = True
End Function

' Convierte a numero las celdas de texto con importes (ej. "1.234,56") y aplica formato contable.
' Las formulas no se tocan.
Private Function FixAmountColumn(ByVal c As Long) As Long
    Dim rg As Range, txt As Range, ar As Range, cell As Range, v As Variant, fmt As Variant
    Dim i As Long, n As Long, amt As Currency, allOk As Boolean

    Set rg = ws.Range(ws.Cells(gR1, c), ws.Cells(gR2, c))
    On Error Resume Next
    If rg.Cells.Count > 1 Then
        Set txt = rg.SpecialCells(xlCellTypeConstants, xlTextValues)
    ElseIf VarType(rg.Value2) = vbString Then
        Set txt = rg
    End If
    fmt = rg.NumberFormat
    If Not txt Is Nothing Or IsNull(fmt) Then
        rg.NumberFormat = ACC_FMT
    ElseIf fmt = "General" Or fmt = "@" Then
        rg.NumberFormat = ACC_FMT
    End If
    On Error GoTo 0
    If txt Is Nothing Then Exit Function
    For Each ar In txt.Areas
        If ar.Cells.Count = 1 Then
            If ParseAmountText(CStr(ar.Value2), amt) Then
                ar.Value2 = CDbl(amt)
                n = n + 1
            End If
        Else
            v = ar.Value2
            allOk = True
            For i = 1 To UBound(v, 1)
                If ParseAmountText(CStr(v(i, 1)), amt) Then
                    v(i, 1) = CDbl(amt)
                    n = n + 1
                Else
                    allOk = False
                End If
            Next
            If allOk Then
                ar.Value2 = v
            Else
                i = 0
                For Each cell In ar.Cells
                    i = i + 1
                    If VarType(v(i, 1)) = vbDouble Then cell.Value2 = v(i, 1)
                Next
            End If
        End If
    Next
    FixAmountColumn = n
End Function

Private Function FindOutputColumns() As Long
    Dim lastC As Long, c As Long
    lastC = ws.Cells(gHdrRow, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastC - 3
        If SafeStr(ws.Cells(gHdrRow, c).Value) = HDR_CAT Then
            If SafeStr(ws.Cells(gHdrRow, c + 1).Value) = HDR_DET And SafeStr(ws.Cells(gHdrRow, c + 2).Value) = HDR_ID _
               And SafeStr(ws.Cells(gHdrRow, c + 3).Value) = HDR_CRIT Then
                FindOutputColumns = c
                Exit Function
            End If
        End If
    Next
End Function

'==================================================================================================
'  REPORTES
'==================================================================================================
Private Sub ShowPreliminaryReport()
    Dim msg As String, k As Long, i As Long, g As Long, nOk As Long, nUsed As Long, gShow As Long
    Dim pc() As Long, nc() As Long, ps() As Currency, ns() As Currency

    ReDim pc(1 To gG): ReDim nc(1 To gG): ReDim ps(1 To gG): ReDim ns(1 To gG)
    For i = 1 To gN
        If gOk(i) Then
            nOk = nOk + 1
            If gUsed(i) Then
                nUsed = nUsed + 1
            ElseIf gAmt(i) > 0 Then
                pc(gGi(i)) = pc(gGi(i)) + 1: ps(gGi(i)) = ps(gGi(i)) + gAmt(i)
            Else
                nc(gGi(i)) = nc(gGi(i)) + 1: ns(gGi(i)) = ns(gGi(i)) + gAmt(i)
            End If
        End If
    Next
    msg = "=== CORTE PRELIMINAR (cruces por llave) ===" & vbCrLf & _
          "Partidas con importe: " & FmtN(nOk) & "    Cruzadas: " & FmtN(nUsed) & _
          " (" & Format$(nUsed / nOk, "0.0%") & ")" & vbCrLf & vbCrLf
    For k = 1 To gLvCount
        msg = msg & k & ". " & Left$(gLvName(k), 48) & ": " & FmtN(gLvCross(k)) & " cruces, " & _
              FmtN(gLvItems(k)) & " partidas" & gLvExtra(k) & vbCrLf
    Next
    msg = msg & vbCrLf & "PENDIENTES DE CRUCE:" & vbCrLf
    gShow = gG
    If gShow > 6 Then gShow = 6
    For g = 1 To gShow
        msg = msg & GrpTag(g) & " (+): " & FmtN(pc(g)) & " partidas   Bs " & FmtM(ps(g)) & vbCrLf
        msg = msg & GrpTag(g) & " (-): " & FmtN(nc(g)) & " partidas   Bs " & FmtM(ns(g)) & vbCrLf
    Next
    msg = msg & vbCrLf & "Tiempo de proceso: " & Format$(gProcSecs, "0.0") & " s"
    MsgBox Left$(msg, 1020), vbInformation, "Corte preliminar - " & TITULO
End Sub

Private Sub ShowFinalReport()
    Dim d As Object, i As Long, k As Long, g As Long, nCat As Long, gShow As Long, totN As Long
    Dim catName() As String, cnt() As Long, sm() As Currency, tot() As Currency, ctrl As Currency
    Dim msg As String, ln As String

    gShow = gG
    If gShow > 3 Then gShow = 3
    ReDim catName(1 To 20): ReDim cnt(1 To 20): ReDim sm(1 To 20, 1 To gG): ReDim tot(1 To gG)
    Set d = NewDict()
    For i = 1 To gN
        If gOk(i) Then
            If Len(gCat(i)) > 0 Then
                If d.Exists(gCat(i)) Then
                    k = d.Item(gCat(i))
                ElseIf nCat < 20 Then
                    nCat = nCat + 1: k = nCat
                    d.Add gCat(i), k
                    catName(k) = gCat(i)
                Else
                    k = 20
                End If
                cnt(k) = cnt(k) + 1
                sm(k, gGi(i)) = sm(k, gGi(i)) + gAmt(i)
            End If
            tot(gGi(i)) = tot(gGi(i)) + gAmt(i)
            totN = totN + 1
            If gCat(i) = CAT_PREV Or gCat(i) = CAT_NET Then ctrl = ctrl + gAmt(i)
        End If
    Next
    Set d = Nothing

    msg = "=== DICTAMEN FINAL DE CLASIFICACIÓN ===" & vbCrLf & "(partidas | Bs por grupo de cuenta)" & vbCrLf & vbCrLf
    For k = 1 To nCat
        ln = catName(k) & ": " & FmtN(cnt(k))
        For g = 1 To gShow
            ln = ln & " | " & GrpTag(g) & " " & FmtM(sm(k, g))
        Next
        msg = msg & ln & vbCrLf
    Next
    ln = "TOTAL: " & FmtN(totN)
    For g = 1 To gShow
        ln = ln & " | " & GrpTag(g) & " " & FmtM(tot(g))
    Next
    msg = msg & "----------------------------------------------" & vbCrLf & ln & vbCrLf & vbCrLf
    msg = msg & "Control Neto + Provisiones/Reversiones: Bs " & FmtM(ctrl)
    If ctrl = 0 Then
        msg = msg & " (cuadra)" & vbCrLf
    Else
        msg = msg & " (revisar)" & vbCrLf
    End If
    msg = msg & "Cruces: PREV " & FmtN(CntOf(PFX_PREV)) & " | NET " & FmtN(CntOf(PFX_NET)) & _
          " | FB " & FmtN(CntOf(PFX_FB)) & " | ASI " & FmtN(CntOf(PFX_ASI)) & vbCrLf
    If Len(gFinalMsg) > 0 Then msg = msg & gFinalMsg & vbCrLf
    If gNErr + gNEmpty + gNZero + gNHidden > 0 Then
        msg = msg & "No procesadas: importe no válido " & FmtN(gNErr) & " | vacías " & FmtN(gNEmpty) & _
              " | importe 0 " & FmtN(gNZero) & " | ocultas " & FmtN(gNHidden) & vbCrLf
    End If
    If gNFixed > 0 Then msg = msg & "Celdas de texto convertidas a número: " & FmtN(gNFixed) & vbCrLf
    msg = msg & "Tiempo de proceso: " & Format$(gProcSecs, "0.0") & " s" & vbCrLf & vbCrLf & _
          "_________________________________" & vbCrLf & "A.M."
    MsgBox Left$(msg, 1020), vbInformation, "Dictamen final - " & TITULO
End Sub

'==================================================================================================
'  ETIQUETADO E IDS
'==================================================================================================
' Marca un conjunto de partidas que suman 0 con un mismo ID. Cuentas distintas -> PREV, misma -> NET.
Private Sub LabelSet(m() As Long, ByVal cnt As Long, ByVal pfxOv As String, ByVal crit As String, _
                     Optional ByVal catOv As String = "")
    Dim i As Long, cross As Boolean, g0 As Long, sfx As String, a5 As String
    Dim pfx As String, num As Long, cat As String, det As String

    g0 = gGi(m(1))
    For i = 2 To cnt
        If gGi(m(i)) <> g0 Then
            cross = True
            Exit For
        End If
    Next
    If cross Then
        cat = CAT_PREV: pfx = PFX_PREV
        For i = 1 To cnt
            a5 = Left$(gAcc(m(i)), 5)
            If Len(a5) > 0 Then
                If Left$(a5, 1) = "4" Or Left$(a5, 1) = "5" Then
                    If InStr(1, "/" & sfx & "/", "/" & a5 & "/") = 0 Then
                        If Len(sfx) > 0 Then sfx = sfx & "/"
                        sfx = sfx & a5
                    End If
                End If
            End If
        Next
        det = cat
        If Len(sfx) > 0 Then det = cat & " " & sfx
    Else
        cat = CAT_NET: pfx = PFX_NET: det = CAT_NET
    End If
    If Len(pfxOv) > 0 Then pfx = pfxOv
    If Len(catOv) > 0 Then
        cat = catOv: det = catOv
    End If
    num = NextNum(pfx)
    For i = 1 To cnt
        gUsed(m(i)) = True
        gCat(m(i)) = cat: gDet(m(i)) = det: gPfx(m(i)) = pfx: gNum(m(i)) = num: gCrit(m(i)) = crit
    Next
End Sub

Private Sub LabelPair(ByVal i As Long, ByVal j As Long, ByVal pfx As String, ByVal crit As String)
    Dim m() As Long
    ReDim m(1 To 2)
    m(1) = i: m(2) = j
    LabelSet m, 2, pfx, crit
End Sub

Private Function NextNum(ByVal pfx As String) As Long
    If gCnt.Exists(pfx) Then
        gCnt.Item(pfx) = gCnt.Item(pfx) + 1
    Else
        gCnt.Add pfx, 1
    End If
    NextNum = gCnt.Item(pfx)
End Function

Private Function CntOf(ByVal pfx As String) As Long
    If gCnt.Exists(pfx) Then CntOf = gCnt.Item(pfx)
End Function

'==================================================================================================
'  CONFIGURACION DE NIVELES (lista, guardado y recuperacion)
'==================================================================================================
Private Sub AddLevel(ByVal typ As Long, ByVal cols As Variant, ByVal glosaCol As Long)
    Dim nm As String, k As Long, n As Long
    If typ = 1 Then
        nm = "Ref. de comprobante en " & HdrName(glosaCol) & " -> " & JoinHdr(cols)
    Else
        nm = JoinHdr(cols)
    End If
    For k = 1 To gLvCount
        If gLvName(k) = nm Then Exit Sub
    Next
    n = gLvCount + 1
    ReDim Preserve gLvType(1 To n)
    ReDim Preserve gLvCols(1 To n)
    ReDim Preserve gLvGlosa(1 To n)
    ReDim Preserve gLvName(1 To n)
    gLvType(n) = typ
    gLvCols(n) = cols
    gLvGlosa(n) = glosaCol
    gLvName(n) = nm
    gLvCount = n
End Sub

Private Sub KeepLevels(keep() As Boolean)
    Dim k As Long, n As Long, t() As Long, c() As Variant, gl() As Long, nm() As String
    For k = 1 To gLvCount
        If keep(k) Then n = n + 1
    Next
    If n = 0 Then
        gLvCount = 0
        Exit Sub
    End If
    ReDim t(1 To n): ReDim c(1 To n): ReDim gl(1 To n): ReDim nm(1 To n)
    n = 0
    For k = 1 To gLvCount
        If keep(k) Then
            n = n + 1
            t(n) = gLvType(k): c(n) = gLvCols(k): gl(n) = gLvGlosa(k): nm(n) = gLvName(k)
        End If
    Next
    gLvType = t: gLvCols = c: gLvGlosa = gl: gLvName = nm
    gLvCount = n
End Sub

Private Function LevelList(ByVal a As Long, ByVal b As Long) As String
    Dim k As Long, s As String
    For k = a To b
        If k - a >= 10 Then
            s = s & "..." & vbCrLf
            Exit For
        End If
        s = s & k & ". " & Left$(gLvName(k), 62) & vbCrLf
    Next
    LevelList = s
End Function

Private Sub SaveLevels()
    Dim k As Long, j As Long, s As String, part As String, cols As Variant
    For k = 1 To gLvCount
        cols = gLvCols(k)
        If gLvType(k) = 1 Then
            part = "R:" & HdrKey(gLvGlosa(k)) & ">"
        Else
            part = "K:"
        End If
        For j = LBound(cols) To UBound(cols)
            If j > LBound(cols) Then part = part & "+"
            part = part & HdrKey(CLng(cols(j)))
        Next
        If Len(s) > 0 Then s = s & "|"
        s = s & part
    Next
    RegSet "Niveles", s
End Sub

Private Function LoadSavedLevels() As Boolean
    Dim s As String, parts() As String, k As Long, typ As Long, glo As Long, cols As Variant
    s = RegGet("Niveles", "")
    If Len(s) = 0 Then Exit Function
    parts = Split(s, "|")
    For k = 0 To UBound(parts)
        If Not ParseLevelSpec(parts(k), typ, glo, cols) Then Exit Function
    Next
    For k = 0 To UBound(parts)
        If ParseLevelSpec(parts(k), typ, glo, cols) Then AddLevel typ, cols, glo
    Next
    LoadSavedLevels = (gLvCount > 0)
End Function

Private Function ParseLevelSpec(ByVal spec As String, ByRef typ As Long, ByRef glo As Long, ByRef cols As Variant) As Boolean
    Dim body As String, parts() As String, k As Long, arr() As Long, c As Long, gp As Long
    glo = 0
    If Left$(spec, 2) = "R:" Then
        typ = 1: body = Mid$(spec, 3)
        gp = InStr(1, body, ">")
        If gp = 0 Then Exit Function
        glo = ColByKey(Left$(body, gp - 1))
        If glo = 0 Then Exit Function
        body = Mid$(body, gp + 1)
    ElseIf Left$(spec, 2) = "K:" Then
        typ = 2: body = Mid$(spec, 3)
    Else
        Exit Function
    End If
    If Len(body) = 0 Then Exit Function
    parts = Split(body, "+")
    ReDim arr(0 To UBound(parts))
    For k = 0 To UBound(parts)
        c = ColByKey(parts(k))
        If c = 0 Then Exit Function
        arr(k) = c
    Next
    cols = arr
    ParseLevelSpec = True
End Function

Private Function ColByKey(ByVal key As String) As Long
    Dim lastC As Long, c As Long
    If Len(key) = 0 Then Exit Function
    lastC = ws.Cells(gHdrRow, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastC
        If HdrKey(c) = key Then
            ColByKey = c
            Exit Function
        End If
    Next
End Function

'==================================================================================================
'  ENCABEZADOS
'==================================================================================================
' Fila de encabezados sugerida: la que tiene mas nombres tipicos de un mayor (primeras 40 filas).
Private Function ScanHeaderRow() As Long
    Dim lastR As Long, lastC As Long, a As Variant, r As Long, c As Long, sc As Long, best As Long, bestR As Long
    With ws.UsedRange
        lastR = .Row + .Rows.Count - 1
        lastC = .Column + .Columns.Count - 1
    End With
    If lastR > 40 Then lastR = 40
    If lastC > 300 Then lastC = 300
    If lastR < 2 And lastC < 2 Then Exit Function
    a = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value2
    For r = 1 To lastR
        sc = 0
        For c = 1 To lastC
            If VarType(a(r, c)) = vbString Then
                If IsStdHeader(Replace(NormText(CStr(a(r, c))), " ", "")) Then sc = sc + 1
            End If
        Next
        If sc > best Then
            best = sc: bestR = r
        End If
    Next
    If best >= 3 Then ScanHeaderRow = bestR
End Function

Private Function IsStdHeader(ByVal k As String) As Boolean
    Dim w As Variant
    If Len(k) = 0 Or Len(k) > 40 Then Exit Function
    For Each w In Array("SALDO", "DEBE", "HABER", "CUENTA", "GLOSA", "TRANSACC", "DOCUMENTO", "USUARIO", _
                        "FECHA", "MODULO", "IMPORTE", "MONTO", "DESCRIPCION", "COMPROBANTE")
        If InStr(1, k, w) > 0 Then
            IsStdHeader = True
            Exit Function
        End If
    Next
End Function

Private Function FindHdrCol(ByVal hr As Long, ByVal must As String, ByVal excl As String) As Long
    Dim lastC As Long, c As Long, j As Long, k As String, ex() As String, bad As Boolean
    lastC = ws.Cells(hr, ws.Columns.Count).End(xlToLeft).Column
    If Len(excl) > 0 Then ex = Split(excl, "|")
    For c = 1 To lastC
        k = HdrKeyAt(hr, c)
        If InStr(1, k, must) > 0 Then
            bad = IsOwnHeader(k)
            If Len(excl) > 0 Then
                For j = 0 To UBound(ex)
                    If InStr(1, k, ex(j)) > 0 Then bad = True
                Next
            End If
            If Not bad Then
                FindHdrCol = c
                Exit Function
            End If
        End If
    Next
End Function

Private Function IsOwnHeader(ByVal k As String) As Boolean
    IsOwnHeader = (k = "CLASIFICACION" Or k = "DETALLECLASIFICACION" Or k = "IDCRUCE" Or k = "CRITERIO")
End Function

Private Function IsHaberHeader(ByVal k As String) As Boolean
    IsHaberHeader = (InStr(1, k, "HABER") > 0 Or InStr(1, k, "CREDITO") > 0 Or InStr(1, k, "ABONO") > 0)
End Function

Private Function HdrKeyAt(ByVal hr As Long, ByVal c As Long) As String
    Dim v As Variant
    v = ws.Cells(hr, c).Value2
    If VarType(v) = vbString Then HdrKeyAt = Replace(NormText(CStr(v)), " ", "")
End Function

Private Function HdrKey(ByVal c As Long) As String
    HdrKey = HdrKeyAt(gHdrRow, c)
End Function

Private Function HdrName(ByVal c As Long) As String
    Dim v As Variant, s As String
    v = ws.Cells(gHdrRow, c).Value2
    If Not IsError(v) Then s = CStr(v)
    s = Replace(Replace(Replace(s, vbCrLf, " "), vbLf, " "), vbCr, " ")
    Do While InStr(1, s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Trim$(s)
    If Len(s) = 0 Then s = "Col " & ColLetter(c)
    HdrName = s
End Function

Private Function JoinHdr(ByVal cols As Variant) As String
    Dim k As Long, s As String
    For k = LBound(cols) To UBound(cols)
        If Len(s) > 0 Then s = s & " + "
        s = s & HdrName(CLng(cols(k)))
    Next
    JoinHdr = s
End Function

Private Function AmountDesc() As String
    If gAmtCol2 > 0 Then
        AmountDesc = HdrName(gAmtCol1) & " - " & HdrName(gAmtCol2) & " (Debe - Haber)"
    Else
        AmountDesc = HdrName(gAmtCol1)
    End If
End Function

'==================================================================================================
'  LECTURA Y NORMALIZACION
'==================================================================================================
Private Function ReadColumn(ByVal c As Long, ByVal r1 As Long, ByVal r2 As Long) As Variant
    Dim a() As Variant
    If r2 = r1 Then
        ReDim a(1 To 1, 1 To 1)
        a(1, 1) = ws.Cells(r1, c).Value2
        ReadColumn = a
    Else
        ReadColumn = ws.Range(ws.Cells(r1, c), ws.Cells(r2, c)).Value2
    End If
End Function

Private Function RawCol(ByVal c As Long) As Variant
    Dim k As String
    k = CStr(c)
    If Not gRaw.Exists(k) Then gRaw.Add k, ReadColumn(c, gR1, gR2)
    RawCol = gRaw.Item(k)
End Function

Private Function NormCol(ByVal c As Long) As Variant
    Dim k As String, raw As Variant, arr() As String, i As Long
    k = CStr(c)
    If Not gNorm.Exists(k) Then
        raw = RawCol(c)
        ReDim arr(1 To gN)
        For i = 1 To gN
            arr(i) = NormKey(raw(i, 1))
        Next
        gNorm.Add k, arr
    End If
    NormCol = gNorm.Item(k)
End Function

' Valor de celda -> texto normalizado para llaves: mayusculas, sin tildes, sin signos, espacios
' simples; "000123" = 123; las fechas se comparan por su numero de serie (Value2).
Private Function NormKey(ByVal v As Variant) As String
    Select Case VarType(v)
        Case vbString
            NormKey = NormText(CStr(v))
        Case vbDouble, vbSingle, vbCurrency, vbDecimal, vbLong, vbInteger, vbByte
            If v = Int(v) And Abs(v) < 1E+15 Then
                NormKey = Format$(v, "0")
            Else
                NormKey = Trim$(Str$(v))
            End If
        Case vbBoolean
            If v Then NormKey = "VERDADERO" Else NormKey = "FALSO"
        Case vbDate
            NormKey = Format$(CDbl(v), "0")
    End Select
End Function

Private Function NormText(ByVal s As String) As String
    Dim b() As Byte, o() As Byte, i As Long, L As Long, k As Long, code As Long, m As Integer
    Dim pend As Boolean, allDig As Boolean

    L = Len(s)
    If L = 0 Then Exit Function
    b = s
    ReDim o(0 To 2 * L - 1)
    allDig = True
    For i = 0 To 2 * L - 2 Step 2
        code = CLng(b(i)) + CLng(b(i + 1)) * 256
        m = gMap(code)
        If m = 0 Then
            If k > 0 Then pend = True
        Else
            If pend Then
                o(k) = 32: k = k + 2
                pend = False: allDig = False
            End If
            o(k) = m: k = k + 2
            If m < 48 Or m > 57 Then allDig = False
        End If
    Next
    If k = 0 Then Exit Function
    ReDim Preserve o(0 To k - 1)
    NormText = o
    If allDig Then NormText = StripZeros(NormText)
End Function

Private Sub InitCharMap()
    Dim i As Long
    If gMapReady Then Exit Sub
    For i = 48 To 57
        gMap(i) = i
    Next
    For i = 65 To 90
        gMap(i) = i
        gMap(i + 32) = i
    Next
    SetMap "ÀÁÂÃÄÅàáâãäåª", 65
    SetMap "Çç", 67
    SetMap "ÈÉÊËèéêë", 69
    SetMap "ÌÍÎÏìíîï", 73
    SetMap "Ññ", 78
    SetMap "ÒÓÔÕÖØòóôõöøº", 79
    SetMap "Šš", 83
    SetMap "ÙÚÛÜùúûü", 85
    SetMap "ÝýÿŸ", 89
    SetMap "Žž", 90
    gMapReady = True
End Sub

Private Sub SetMap(ByVal chars As String, ByVal target As Integer)
    Dim k As Long
    For k = 1 To Len(chars)
        gMap(AscW(Mid$(chars, k, 1)) And &HFFFF&) = target
    Next
End Sub

' Numeros de 'minLen' o mas digitos dentro de un texto (sin ceros a la izquierda).
' Devuelve un arreglo 1..n de textos, o Empty si no hay ninguno.
Private Function DigitTokens(ByVal s As String, ByVal minLen As Long) As Variant
    Dim b() As Byte, k As Long, L As Long, st As Long, n As Long, runLen As Long, isD As Boolean, tok() As String

    L = Len(s)
    If L < minLen Then Exit Function
    b = s
    ReDim tok(1 To 8)
    For k = 0 To 2 * L Step 2
        isD = False
        If k < 2 * L Then
            If b(k + 1) = 0 Then
                If b(k) >= 48 Then
                    If b(k) <= 57 Then isD = True
                End If
            End If
        End If
        If isD Then
            If st = 0 Then st = k \ 2 + 1
        ElseIf st > 0 Then
            runLen = k \ 2 + 1 - st
            If runLen >= minLen Then
                n = n + 1
                If n > UBound(tok) Then
                    ReDim Preserve tok(1 To n * 2)
                End If
                tok(n) = StripZeros(Mid$(s, st, runLen))
            End If
            st = 0
        End If
    Next
    If n > 0 Then
        ReDim Preserve tok(1 To n)
        DigitTokens = tok
    End If
End Function

Private Function StripZeros(ByVal s As String) As String
    Dim i As Long, L As Long
    L = Len(s)
    i = 1
    Do While i < L
        If Mid$(s, i, 1) <> "0" Then Exit Do
        i = i + 1
    Loop
    StripZeros = Mid$(s, i)
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    If Len(s) > 0 Then IsAllDigits = Not (s Like "*[!0-9]*")
End Function

Private Function AccDigits(ByVal v As Variant) As String
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbDecimal, vbLong, vbInteger
            If v >= 0 And v = Int(v) And v < 1E+15 Then
                AccDigits = Format$(v, "0")
            Else
                AccDigits = DigitsOnly(CStr(v))
            End If
        Case vbString
            AccDigits = DigitsOnly(CStr(v))
    End Select
End Function

Private Function DigitsOnly(ByVal s As String) As String
    Dim b() As Byte, o() As Byte, i As Long, k As Long, L As Long
    L = Len(s)
    If L = 0 Then Exit Function
    b = s
    ReDim o(0 To 2 * L - 1)
    For i = 0 To 2 * L - 2 Step 2
        If b(i + 1) = 0 Then
            If b(i) >= 48 Then
                If b(i) <= 57 Then
                    o(k) = b(i): k = k + 2
                End If
            End If
        End If
    Next
    If k = 0 Then Exit Function
    ReDim Preserve o(0 To k - 1)
    DigitsOnly = o
End Function

'==================================================================================================
'  IMPORTES
'==================================================================================================
' kind: 0 = numero, 1 = texto convertido, 2 = error (#N/A...), 3 = vacio, 4 = texto no numerico
Private Function ParseCell(ByVal v As Variant, ByRef amt As Currency, ByRef kind As Long) As Boolean
    amt = 0: kind = 3
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbDecimal, vbLong, vbInteger, vbByte
            amt = RoundCents(CDbl(v)): kind = 0
            ParseCell = True
        Case vbString
            If ParseAmountText(CStr(v), amt) Then
                kind = 1
                ParseCell = True
            ElseIf Len(Trim$(v)) > 0 And Trim$(v) <> "-" Then
                kind = 4
            End If
        Case vbError
            kind = 2
    End Select
End Function

' Texto -> importe. Acepta 1.234,56 / 1,234.56 / 1234,56 / (1.234,56) / 1234- / Bs 1.234,56
Private Function ParseAmountText(ByVal s As String, ByRef amt As Currency) As Boolean
    Dim neg As Boolean, i As Long, ch As String, nDot As Long, nCom As Long, lastDot As Long, lastCom As Long
    Dim dec As String, t As String, sysDec As String

    s = UCase$(s)
    s = Replace(s, ChrW$(160), ""): s = Replace(s, " ", ""): s = Replace(s, vbTab, "")
    s = Replace(s, "BS.", ""): s = Replace(s, "BS", ""): s = Replace(s, "USD", "")
    s = Replace(s, "$US", ""): s = Replace(s, "$", "")
    If Len(s) = 0 Then Exit Function
    If Left$(s, 1) = "(" And Right$(s, 1) = ")" And Len(s) > 2 Then
        neg = True: s = Mid$(s, 2, Len(s) - 2)
    End If
    If Right$(s, 1) = "-" And Len(s) > 1 Then
        neg = Not neg: s = Left$(s, Len(s) - 1)
    End If
    If Left$(s, 1) = "-" And Len(s) > 1 Then
        neg = Not neg: s = Mid$(s, 2)
    ElseIf Left$(s, 1) = "+" And Len(s) > 1 Then
        s = Mid$(s, 2)
    End If
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "0" To "9"
            Case "."
                nDot = nDot + 1: lastDot = i
            Case ","
                nCom = nCom + 1: lastCom = i
            Case Else
                Exit Function
        End Select
    Next
    If nDot + nCom = Len(s) Then Exit Function

    sysDec = Application.International(xlDecimalSeparator)
    If nDot > 0 And nCom > 0 Then
        If lastDot > lastCom Then dec = "." Else dec = ","
    ElseIf nCom = 1 Then
        If Len(s) - lastCom <> 3 Then
            dec = ","
        ElseIf sysDec = "," Then
            dec = ","
        End If
    ElseIf nDot = 1 Then
        If Len(s) - lastDot <> 3 Then
            dec = "."
        ElseIf sysDec = "." Then
            dec = "."
        End If
    End If
    If dec = "." Then
        If nDot > 1 Then Exit Function
        t = Replace(s, ",", "")
    ElseIf dec = "," Then
        If nCom > 1 Then Exit Function
        t = Replace(Replace(s, ".", ""), ",", ".")
    Else
        t = Replace(Replace(s, ".", ""), ",", "")
    End If
    If neg Then
        amt = RoundCents(-Val(t))
    Else
        amt = RoundCents(Val(t))
    End If
    ParseAmountText = True
End Function

Private Function RoundCents(ByVal d As Double) As Currency
    Dim c As Currency
    If d > 9E+11 Then d = 9E+11
    If d < -9E+11 Then d = -9E+11
    c = CCur(d)
    If c >= 0 Then
        RoundCents = Int(c * 100 + 0.5) / 100
    Else
        RoundCents = -Int(-c * 100 + 0.5) / 100
    End If
End Function

Private Function AmtKey(ByVal c As Currency) As String
    AmtKey = CStr(Abs(c))
End Function

'==================================================================================================
'  UTILIDADES
'==================================================================================================
Private Function AskRange(ByVal prompt As String, Optional ByVal defAddr As String = "") As Range
    Dim r As Range
    On Error Resume Next
    If Len(defAddr) > 0 Then
        Set r = Application.InputBox(Prompt:=prompt, Title:=TITULO, Default:=defAddr, Type:=8)
    Else
        Set r = Application.InputBox(Prompt:=prompt, Title:=TITULO, Type:=8)
    End If
    On Error GoTo 0
    If Not r Is Nothing Then
        If Not r.Worksheet Is ws Then
            MsgBox "Seleccione celdas de la hoja activa (" & ws.Name & ").", vbExclamation, TITULO
            Set r = Nothing
        End If
    End If
    Set AskRange = r
End Function

Private Function RangeColumns(ByVal r As Range) As Variant
    Dim ar As Range, c As Long, n As Long, cols() As Long, seen As String
    For Each ar In r.Areas
        If ar.Columns.Count > 60 Then Exit Function
        For c = ar.Column To ar.Column + ar.Columns.Count - 1
            If InStr(1, seen, "|" & c & "|") = 0 Then
                seen = seen & "|" & c & "|"
                ReDim Preserve cols(0 To n)
                cols(n) = c
                n = n + 1
            End If
        Next
    Next
    If n > 0 Then RangeColumns = cols
End Function

' Similitud de glosas: cantidad de palabras (4+ letras) en comun.
Private Function Similarity(ByVal i As Long, ByVal j As Long) As Long
    Dim wi As Variant, wj As Variant, a As Long, b As Long, n As Long
    If IsEmpty(gWords(i)) Then gWords(i) = GlosaWords(i)
    If IsEmpty(gWords(j)) Then gWords(j) = GlosaWords(j)
    wi = gWords(i): wj = gWords(j)
    For a = LBound(wi) To UBound(wi)
        For b = LBound(wj) To UBound(wj)
            If wi(a) = wj(b) Then
                n = n + 1
                Exit For
            End If
        Next
    Next
    Similarity = n
End Function

Private Function GlosaWords(ByVal i As Long) As Variant
    Dim parts() As String, w() As String, k As Long, n As Long, s As String, dup As String
    GlosaWords = Array()
    If Not IsArray(gGlosaNorm) Then Exit Function
    s = gGlosaNorm(i)
    If Len(s) = 0 Then Exit Function
    parts = Split(s, " ")
    ReDim w(0 To UBound(parts))
    For k = 0 To UBound(parts)
        If Len(parts(k)) >= 4 Then
            If Not IsAllDigits(parts(k)) Then
                If InStr(1, " " & dup & " ", " " & parts(k) & " ") = 0 Then
                    w(n) = parts(k): n = n + 1
                    dup = dup & " " & parts(k)
                End If
            End If
        End If
    Next
    If n > 0 Then
        ReDim Preserve w(0 To n - 1)
        GlosaWords = w
    End If
End Function

' Ordena de mayor a menor; a igual importe, por fila ascendente (resultado estable y trazable).
Private Sub QSortDesc(v() As Currency, ix() As Long, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Currency, pIx As Long, tv As Currency, ti As Long
    Do While lo < hi
        i = lo: j = hi
        p = v((lo + hi) \ 2): pIx = ix((lo + hi) \ 2)
        Do While i <= j
            Do While v(i) > p Or (v(i) = p And ix(i) < pIx)
                i = i + 1
            Loop
            Do While v(j) < p Or (v(j) = p And ix(j) > pIx)
                j = j - 1
            Loop
            If i <= j Then
                tv = v(i): v(i) = v(j): v(j) = tv
                ti = ix(i): ix(i) = ix(j): ix(j) = ti
                i = i + 1: j = j - 1
            End If
        Loop
        If (j - lo) < (hi - i) Then
            If lo < j Then QSortDesc v, ix, lo, j
            lo = i
        Else
            If i < hi Then QSortDesc v, ix, i, hi
            hi = j
        End If
    Loop
End Sub

Private Sub ChainToArray(ByVal first As Long, nx() As Long, mem() As Long, ByRef cnt As Long)
    Dim i As Long
    cnt = 0
    i = first
    Do While i <> 0
        If Not gUsed(i) Then
            cnt = cnt + 1
            If cnt = 1 Then
                ReDim mem(1 To 16)
            ElseIf cnt > UBound(mem) Then
                ReDim Preserve mem(1 To cnt * 2)
            End If
            mem(cnt) = i
        End If
        i = nx(i)
    Loop
End Sub

Private Function TimeUp() As Boolean
    Dim e As Double
    If Not gTimeLimited Then Exit Function
    e = Timer - gTimeStart
    If e < 0 Then e = e + 86400
    If e > FB_TIME_LIMIT Then
        gTimedOut = True
        TimeUp = True
    End If
End Function

Private Sub ClockStart()
    gClock = Timer
End Sub

Private Sub ClockStop()
    Dim e As Double
    e = Timer - gClock
    If e < 0 Then e = e + 86400
    gProcSecs = gProcSecs + e
End Sub

Private Function CountOk() As Long
    Dim i As Long, n As Long
    For i = 1 To gN
        If gOk(i) Then n = n + 1
    Next
    CountOk = n
End Function

Private Function CountUsed() As Long
    Dim i As Long, n As Long
    For i = 1 To gN
        If gUsed(i) Then n = n + 1
    Next
    CountUsed = n
End Function

Private Function CountPending() As Long
    Dim i As Long, n As Long
    For i = 1 To gN
        If gOk(i) Then
            If Not gUsed(i) Then n = n + 1
        End If
    Next
    CountPending = n
End Function

Private Function LastRowOf(ByVal c As Long) As Long
    LastRowOf = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
End Function

Private Function IsTotalRow(ByVal r As Long) As Boolean
    Dim f As String
    f = UCase$(CStr(ws.Cells(r, gAmtCol1).Formula))
    If Left$(f, 1) <> "=" Then Exit Function
    If InStr(1, f, "SUM(") = 0 And InStr(1, f, "SUBTOTAL(") = 0 And InStr(1, f, "AGGREGATE(") = 0 Then Exit Function
    If gAccCol > 0 Then
        If Not IsBlankV(ws.Cells(r, gAccCol).Value2) Then Exit Function
    End If
    IsTotalRow = True
End Function

Private Function IsBlankV(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        IsBlankV = True
    ElseIf VarType(v) = vbString Then
        IsBlankV = (Len(Trim$(v)) = 0)
    End If
End Function

Private Function IsBlankAt(ByVal arr As Variant, ByVal i As Long) As Boolean
    If IsArray(arr) Then
        IsBlankAt = IsBlankV(arr(i, 1))
    Else
        IsBlankAt = True
    End If
End Function

Private Function SafeStr(ByVal v As Variant) As String
    If Not IsError(v) Then SafeStr = CStr(v)
End Function

Private Function ColLetter(ByVal c As Long) As String
    ColLetter = Split(ws.Cells(1, c).Address(True, False), "$")(0)
End Function

Private Function GrpLabel(ByVal g As Long) As String
    If Len(gGrpName(g)) = 0 Then GrpLabel = "(sin cuenta)" Else GrpLabel = gGrpName(g)
End Function

Private Function GrpTag(ByVal g As Long) As String
    GrpTag = "Cta " & GrpLabel(g)
End Function

Private Function FmtN(ByVal n As Double) As String
    FmtN = Format$(n, "#,##0")
End Function

Private Function FmtM(ByVal c As Currency) As String
    FmtM = Format$(c, "#,##0.00")
End Function

Private Function RegGet(ByVal nm As String, ByVal def As String) As String
    On Error Resume Next
    RegGet = def
    RegGet = GetSetting(REG_APP, "Config", nm, def)
    On Error GoTo 0
End Function

Private Sub RegSet(ByVal nm As String, ByVal v As String)
    On Error Resume Next
    SaveSetting REG_APP, "Config", nm, v
    On Error GoTo 0
End Sub

Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

Private Sub ResetState()
    Set gCnt = NewDict(): Set gRaw = NewDict(): Set gNorm = NewDict()
    gN = 0: gG = 0: gLvCount = 0
    gAmtCol1 = 0: gAmtCol2 = 0: gAccCol = 0: gDigits = 1: gGlosaCol = 0
    gNText = 0: gNErr = 0: gNEmpty = 0: gNZero = 0: gNHidden = 0: gNFixed = 0
    gProcSecs = 0: gFinalMsg = vbNullString
    gTimeLimited = False: gTimedOut = False
    gGlosaNorm = Empty
End Sub

Private Sub ReleaseState()
    Set gCnt = Nothing: Set gRaw = Nothing: Set gNorm = Nothing: Set sRun = Nothing
    Erase gAmt, gIncl, gOk, gBad, gUsed, gAcc, gGi, gCat, gDet, gCrit, gPfx, gNum, gWords
    Erase sVal, sIdx, sTaken, sChosen
    gGlosaNorm = Empty
    Set ws = Nothing
End Sub
