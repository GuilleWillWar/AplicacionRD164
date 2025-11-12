Attribute VB_Name = "Funciones"

Private Const CELDA_SUPERFICIE As String = "B3"

Private Const celdaActividad As String = "D3"
Private Const celdaAltura As String = "B7"
Private Const celdaAutonomia As String = "A53"
Private Const celdaCaudalMinimo As String = "F51"
Private Const celdaConfig As String = "B2"
Private Const celdaConfiguracion As String = "B2"
Private Const celdaCubiertaEvacua As String = "B46"
Private Const celdaCubiertaLigera As String = "B47"
Private Const celdaDeteccionAuto As String = "A51"
Private Const celdaDeteccionManual As String = "B51"
Private Const celdaDistancia As String = "B24"
Private Const celdaDistancia1 As String = "F30"
Private Const celdaDistancia2 As String = "F31"
Private Const celdaDistanciaIntroducida As String = "D28"
Private Const celdaEdificio As String = "G3"
Private Const celdaEficaciaMinima As String = "B53"
Private Const celdaExigeAuto As String = "A51"
Private Const celdaExigeCam As String = "D51"
Private Const celdaExigeDir As String = "E51"
Private Const celdaExigeMan As String = "B51"
Private Const celdaExtintoresPortatiles As String = "D53"
Private Const celdaFachada As String = "B8"
Private Const celdaFlagBIE As String = "B21"
Private Const celdaFlagBaterias As String = "B16"
Private Const celdaFlagColapso As String = "B19"
Private Const celdaFlagDeteccion As String = "B17"
Private Const celdaFlagEstructura As String = "B18"
Private Const celdaFlagExtincionAuto As String = "B10"
Private Const celdaFlagHumos As String = "B20"
Private Const celdaFlagSeguridad As String = "B22"
Private Const celdaHidrantesCamiones As String = "D51"
Private Const celdaHidrantesDirecta As String = "E51"
Private Const celdaMaxSeguridad As String = "B12"
Private Const celdaMegafonia As String = "C51"
Private Const celdaNombre As String = "G5"
Private Const celdaNri As String = "F2"
Private Const celdaNumero As String = "C53"
Private Const celdaNecesitaBIES As String = "E53"
Private Const celdaNecesitaCol As String = "A55"
Private Const celdaNecesitaExtAuto As String = "B55"
Private Const celdaNecesitaHumos As String = "C55"
Private Const celdaOcupacion As String = "F28"
Private Const celdaPCV As String = "F11"
Private Const celdaPGCV As String = "F13"
Private Const celdaPGSV As String = "F12"
Private Const celdaPSV As String = "F10"
Private Const celdaParedes As String = "F14"
Private Const celdaPersonas As String = "B28"
Private Const celdaPersonasEntrada As String = "B28"
Private Const celdaPersonasSalida As String = "F28"
Private Const celdaRasante As String = "B3"
Private Const celdaResistencia As String = "F9"
Private Const celdaResultado As String = "D29"
Private Const celdaResultadoCamara As String = "F26"
Private Const celdaResultadoEstado As String = "F7"
Private Const celdaResultadoFachada As String = "F25"
Private Const celdaResultadoResistencia As String = "F24"
Private Const celdaResultadoResistenciaEstructural As String = "F46"
Private Const celdaResultadoSuperficie As String = "F8"
Private Const celdaSalida1Distancia As String = "F30"
Private Const celdaSalida2ConAlternativa As String = "F32"
Private Const celdaSalida2SinAlternativa As String = "F31"
Private Const celdaSalidaCantidad As String = "F29"
Private Const celdaSalidas As String = "F29"
Private Const celdaSalidasIntroducidas As String = "B29"
Private Const celdaSituacion As String = "B14"
Private Const celdaSuelos As String = "F15"
Private Const celdaSuperficie As String = "F3"
Private Const celdaSuperficieEstablecimiento As String = "B49"
Private Const celdaSuperficieTotal As String = "B49"
Private Const celdaTechos As String = "F14"
Private Const celdaTipo As String = "F53"
Private Const celdaTipoBie As String = "F53"
Private Const celdaSuperficieViable As String = "D7"
Private Const celdaResistenciaViable As String = "D8"

Private Const ADDR_A1 As String = "A1"
Private Const ADDR_A1_D1 As String = "A1:D1"
Private Const ADDR_C56 As String = "C56"
Private Const ADDR_C57 As String = "C57"
Private Const ADDR_C58 As String = "C58"
Private Const ADDR_C59 As String = "C59"
Private Const ADDR_E56 As String = "E56"
Private Const ADDR_E57 As String = "E57"
Private Const ADDR_E58 As String = "E58"
Private Const ADDR_G6 As String = "G6"
Private Const ADDR_H2_H55 As String = "H2:H55"

Public Sub VerificarViabilidad()

    ' Constantes de celdas en hoja "Interfaz"
    Const celdaFactor1 As String = "B9"
    Const celdaFactor2 As String = "B10"
    Const celdaFactor3 As String = "B11"
    Const celdaComentario As String = "F7"
    
    Dim wsHoja As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Superficie NRI")
    If wsHoja Is Nothing Then Set wsHoja = ActiveSheet
    wsHoja.Activate

    ' Entradas
    Dim configBuscado As String, nriBuscado As String, rasanteBuscado As String
    Dim alturaEvaBuscada As Double, longitudFachadaBuscada As Double
    Dim superficieMinima As Double

    ' Estado
    Dim fachadaOk As Boolean, alturaOk As Boolean, encontrado As Boolean
    Dim NOA As Boolean, tC As Boolean
    Dim superficieCruda As Variant, textoCelda As String
    Dim superficieMax As Double                    ' <-- tope permitido base (tabla × factor)
    Dim factor As Double
    Dim fila As Long, ultimaFila As Long
    Dim errores As String

    errores = "": superficieMax = 0

    ' Validaciones
    If wsHoja.Range(celdaConfig).Value = "" Then errores = errores & "- Falta tipo de configuracion" & vbCrLf
    If IsError(wsHoja.Range(celdaNri).Value) Then errores = errores & "- Falta NRI" & vbCrLf
    If wsHoja.Range(celdaSuperficie).Value = "" Then errores = errores & "- Falta superficie" & vbCrLf
    If wsHoja.Range(celdaRasante).Value = "" Then errores = errores & "- Falta tipo de rasante" & vbCrLf
    If wsHoja.Range(celdaAltura).Value = "" Then errores = errores & "- Falta altura de evacuacion" & vbCrLf
    If wsHoja.Range(celdaFachada).Value = "" Then errores = errores & "- Falta longitud de fachada accesible" & vbCrLf

    With wsHoja.Range(celdaComentario)
        If .Comment Is Nothing Then
            If errores <> "" Then .AddComment text:=errores
        Else
            If errores <> "" Then .Comment.text text:=errores Else .Comment.Delete
        End If
    End With
    If errores <> "" Then
        wsHoja.Range(celdaResultadoEstado).Value = "Error"
        wsHoja.Range(celdaResultadoSuperficie).Value = "Error"
        Exit Sub
    End If

    ' Captura entradas
    configBuscado = Trim(LCase(wsHoja.Range(celdaConfig).Value))
    nriBuscado = Trim(LCase(wsHoja.Range(celdaNri).Value))
    superficieMinima = CDbl(wsHoja.Range(celdaSuperficie).Value)
    rasanteBuscado = Trim(LCase(wsHoja.Range(celdaRasante).Value))
    alturaEvaBuscada = CDbl(wsHoja.Range(celdaAltura).Value)
    longitudFachadaBuscada = CDbl(wsHoja.Range(celdaFachada).Value)

    ' Factor
    factor = 1
    If wsHoja.Range(celdaFactor1).Value = True And wsHoja.Range(celdaFactor2).Value = True Then
        factor = 2.5
    ElseIf wsHoja.Range(celdaFactor1).Value = True Then
        factor = 1.25
    ElseIf wsHoja.Range(celdaFactor2).Value = True Then
        factor = 2
    ElseIf wsHoja.Range(celdaFactor3).Value = True Then
        factor = 2.5
    End If

    ' Búsqueda en tabla
    ultimaFila = 104
    For fila = 17 To ultimaFila
        If Trim(LCase(ws.Cells(fila, 2).Value)) = configBuscado _
        And Trim(LCase(ws.Cells(fila, 3).Value)) = nriBuscado _
        And Trim(LCase(ws.Cells(fila, 4).Value)) = rasanteBuscado Then

            superficieCruda = ws.Cells(fila, 5).Value

            If Not IsNumeric(superficieCruda) Then
                If CStr(superficieCruda) = "NO/A" Then
                    NOA = True
                Else
                    tC = True
                    textoCelda = CStr(superficieCruda)
                End If
            Else
                superficieMax = CDbl(superficieCruda) * factor   ' <-- TOPE BASE
                ' Viabilidad
                If superficieMax >= superficieMinima Then
                    If longitudFachadaBuscada >= ws.Cells(fila, 6).Value Then fachadaOk = True
                    If ws.Cells(fila, 8).Value = 0 Or alturaEvaBuscada <= ws.Cells(fila, 8).Value Then alturaOk = True
                    If fachadaOk And alturaOk Then encontrado = True
                End If
            End If
            Exit For
        End If
    Next fila

    ' Escribir SIEMPRE el tope base calculado por esta vía
    If tC Then
        wsHoja.Range(celdaResultadoSuperficie).Value = textoCelda
        wsHoja.Range(celdaSuperficieViable).Value = 0
    ElseIf NOA Then
        wsHoja.Range(celdaResultadoSuperficie).Value = "NO/A"
        wsHoja.Range(celdaSuperficieViable).Value = 0
    Else
        wsHoja.Range(celdaResultadoSuperficie).Value = superficieMax
        wsHoja.Range(celdaSuperficieViable).Value = superficieCruda
    End If

    ' Estado + comentario
    Dim msg As String: msg = ""
    If tC Then
        wsHoja.Range(celdaResultadoEstado).Value = "NO ADMITIDO"
        msg = "En esta configuracion la superficie no es significativa. " & textoCelda
    ElseIf NOA Then
        wsHoja.Range(celdaResultadoEstado).Value = "NO ADMITIDO"
        msg = "Resultado NO/A en tabla base. Valorar NOTA 5 si procede."
    ElseIf encontrado Then
        wsHoja.Range(celdaResultadoEstado).Value = "ADMITIDO"
        If Not wsHoja.Range(celdaComentario).Comment Is Nothing Then wsHoja.Range(celdaComentario).Comment.Delete
    Else
        wsHoja.Range(celdaResultadoEstado).Value = "NO ADMITIDO"
        msg = "Los criterios introducidos no cumplen con los requisitos minimos."
        If Not wsHoja.Range(celdaFactor1).Value And Not wsHoja.Range(celdaFactor2).Value And Not wsHoja.Range(celdaFactor3).Value Then
            msg = msg & vbCrLf & "Analice factores multiplicadores (Notas 2 y 3)."
        End If
        If Not fachadaOk Then msg = msg & vbCrLf & "Rechazo por fachada."
        If Not alturaOk Then msg = msg & vbCrLf & "Rechazo por altura."
    End If

    With wsHoja.Range(celdaComentario)
        If .Comment Is Nothing Then
            If msg <> "" Then .AddComment text:=msg
        Else
            If msg <> "" Then .Comment.text text:=msg Else .Comment.Delete
        End If
    End With
End Sub

Public Sub CalcularNota5()

    ' === Constantes de celdas en hoja Interfaz ===
    Const celdaComentario As String = "F7"
    ' Importante: declaramos aquí también los factores para poder leerlos
    Const celdaFactor1 As String = "B9"   ' Nota 2 (x1,25)
    Const celdaFactor2 As String = "B10"  ' Nota 3 (x2)
    Const celdaFactor3 As String = "B11"  ' Nota 2+3 (x2,5)

    Dim wsHoja As Worksheet
    Set wsHoja = ActiveSheet
    wsHoja.Activate

    ' --- 1) Validación de entradas ---
    Dim errores As String: errores = ""
    If wsHoja.Range(celdaConfig).Value = "" Then errores = errores & "- Falta tipo de configuracion" & vbCrLf
    If IsError(wsHoja.Range(celdaNri).Value) Then errores = errores & "- Falta NRI" & vbCrLf
    If wsHoja.Range(celdaSuperficie).Value = "" Then errores = errores & "- Falta superficie" & vbCrLf
    If wsHoja.Range(celdaRasante).Value = "" Then errores = errores & "- Falta tipo de rasante" & vbCrLf
    If wsHoja.Range(celdaAltura).Value = "" Then errores = errores & "- Falta altura de evacuacion" & vbCrLf
    If wsHoja.Range(celdaFachada).Value = "" Then errores = errores & "- Falta longitud de fachada accesible" & vbCrLf

    If errores <> "" Then
        ' No borramos comentarios existentes; añadimos
        With wsHoja.Range(celdaComentario)
            If .Comment Is Nothing Then
                .AddComment text:=errores
            Else
                .Comment.text text:=.Comment.text & IIf(.Comment.text <> "", vbCrLf, "") & errores
            End If
        End With
        wsHoja.Range(celdaResultadoEstado).Value = "Error"
        ' No tocamos la superficie aquí
        Exit Sub
    End If

    ' --- 2) Entradas y flags ---
    Dim configBuscado As String
    Dim superficieBuscada As Double
    Dim alturaEvaBuscada As Double
    Dim longitudFachadaBuscada As Double

    configBuscado = UCase$(CStr(wsHoja.Range(celdaConfig).Value))
    superficieBuscada = CDbl(wsHoja.Range(celdaSuperficie).Value)
    alturaEvaBuscada = CDbl(wsHoja.Range(celdaAltura).Value)
    longitudFachadaBuscada = CDbl(wsHoja.Range(celdaFachada).Value)

    Dim fExt As Boolean, fBat As Boolean, fDet As Boolean, fEstr As Boolean
    Dim fCol As Boolean, fHum As Boolean, fBIE As Boolean, fSeg As Boolean

    fExt = (wsHoja.Range(celdaFlagExtincionAuto).Value = True)
    fBat = (wsHoja.Range(celdaFlagBaterias).Value = True)
    fDet = (wsHoja.Range(celdaFlagDeteccion).Value = True)
    fEstr = (wsHoja.Range(celdaFlagEstructura).Value = True)
    fCol = (wsHoja.Range(celdaFlagColapso).Value = True)
    fHum = (wsHoja.Range(celdaFlagHumos).Value = True)
    fBIE = (wsHoja.Range(celdaFlagBIE).Value = True)
    fSeg = (wsHoja.Range(celdaFlagSeguridad).Value = True)

    ' --- 3) Prerrequisitos y cálculo del tope Nota 5 (máximo alcanzable con los flags) ---
    Dim preOk As Boolean
    preOk = (fBat And fBIE And alturaEvaBuscada <= 15 And longitudFachadaBuscada >= 5)

    Dim tope As Double: tope = 0

    If preOk Then
        ' 100 m² (dos caminos)
        If fEstr And (configBuscado = "AH" Or configBuscado = "AV") Then tope = Application.Max(tope, 100)
        If fCol And configBuscado = "AH" Then tope = Application.Max(tope, 100)

        ' 300 / 1500 / 3000
        If fEstr And fDet And fHum Then
            Select Case configBuscado
                Case "AV": tope = Application.Max(tope, 300)
                Case "AH": tope = Application.Max(tope, 1500)
                Case "B":  tope = Application.Max(tope, 3000)
            End Select
            ' 5000 / 6000 con Seguridad
            If fSeg Then
                If configBuscado = "AH" Then tope = Application.Max(tope, 5000)
                If configBuscado = "B" Then tope = Application.Max(tope, 6000)
            End If
        End If
    End If

    ' --- 4) Viabilidad por Nota 5 ---
    Dim viab As String
    Dim admitePorNota5 As Boolean
    admitePorNota5 = (preOk And tope > 0 And superficieBuscada <= tope)

    If admitePorNota5 Then
        viab = "ADMITIDO CON NOTA 5"
    Else
        viab = "NO ADMITIDO CON NOTA 5"
    End If

    wsHoja.Range(celdaResultadoEstado).Value = viab

    ' --- 5) Escritura de superficie ---
    ' Solo sobrescribimos la superficie mostrada si Nota 5 ha calculado un tope válido (>0).
    ' Si tope=0 (no aplica o flags insuficientes), NO tocamos la superficie: se mantiene la de VerificarViabilidad.
    If tope > 0 Then
        wsHoja.Range(celdaResultadoSuperficie).Value = tope
    End If

    ' --- 6) Comentarios: NO borrar los existentes; solo añadir Nota 5 ---
    Dim existente As String, add As String, finalTxt As String
    existente = ""
    If Not wsHoja.Range(celdaComentario).Comment Is Nothing Then
        existente = wsHoja.Range(celdaComentario).Comment.text
    End If

    If admitePorNota5 Then
        add = "Sector viable por NOTA 5."
    Else
        If Not preOk Then
            add = "NOTA 5 no aplicable: faltan prerrequisitos (BIE y baterias activas, altura<=15 m, fachada>=5 m)."
        ElseIf tope = 0 Then
            add = "Con los flags actuales no se alcanza un tope por NOTA 5."
        Else
            add = "La superficie solicitada (" & CStr(superficieBuscada) & " m²) excede el tope de NOTA 5 (" & CStr(tope) & " m²)."
        End If
        ' Si no hay Notas 2/3 activas, reforzamos sugerencia (se mantiene lo ya comentado por VerificarViabilidad)
        Dim n2 As Boolean, n3 As Boolean, n23 As Boolean
        n2 = (wsHoja.Range(celdaFactor1).Value = True)
        n3 = (wsHoja.Range(celdaFactor2).Value = True)
        n23 = (wsHoja.Range(celdaFactor3).Value = True)
        If Not (n2 Or n3 Or n23) Then
            add = add & vbCrLf & "Analice si se pueden incluir los factores multiplicadores de superficie (Notas 2 y 3)."
        End If
    End If

    finalTxt = Trim(existente & IIf(Len(existente) > 0 And Len(add) > 0, vbCrLf, "") & add)

    With wsHoja.Range(celdaComentario)
        If .Comment Is Nothing Then
            If finalTxt <> "" Then .AddComment text:=finalTxt
        Else
            If finalTxt <> "" Then .Comment.text text:=finalTxt
        End If
    End With

End Sub



Public Sub ObtenerRevestimientos()

    ' Constantes de celdas en hoja "Interfaz"
    Const celdaComentario As String = "F14"
    Dim wsHoja As Worksheet
    
    ' --- Hoja de trabajo: si no llega, uso la activa (Interfaz o 001/002/...)
    If wsHoja Is Nothing Then Set wsHoja = ActiveSheet
    wsHoja.Activate

    ' Variables
    Dim revestimientoParedes As String
    Dim revestimientoSuelos As String
    Dim situacion As String
    Dim errores As String

    errores = ""

    ' Validacion de entradas
    If wsHoja.Range(celdaSituacion).Value = "" Then errores = "- Falta situacion del elemento" & vbCrLf

    ' Gestion de comentarios y errores
    With wsHoja.Range(celdaComentario)
        If .Comment Is Nothing Then
            If errores <> "" Then .AddComment text:=errores
        Else
            If errores <> "" Then
                .Comment.text text:=errores
            Else
                .Comment.Delete
            End If
        End If
    End With

    ' Si hay errores, ponemos Error en resultados y salimos
    If errores <> "" Then
        wsHoja.Range(celdaParedes).Value = "Error"
        wsHoja.Range(celdaSuelos).Value = "Error"
        Exit Sub
    End If

    ' Captura de entradas
    situacion = Trim(wsHoja.Range(celdaSituacion).Value)

    Select Case LCase(situacion)
    Case "zonas ocupables, en general"
        revestimientoParedes = "C-s2,d0"
        revestimientoSuelos = "CFL-s1"
    Case "pasillos y escaleras protegidos"
        revestimientoParedes = "B-s1,d0"
        revestimientoSuelos = "CFL-s1"
    Case "aparcamientos y sectores de nivel de riesgo intrnseco alto"
        revestimientoParedes = "B-s1,d0"
        revestimientoSuelos = "BFL-s1"
    Case "espacios ocultos no estancos o susceptibles de provocar un encendio"
        revestimientoParedes = "B-s3,d0"
        revestimientoSuelos = "BFL-s2"
    Case Else
        ' Situacin no reconocida: colocar comentario, poner "Error" en resultados y salir
        With wsHoja.Range(celdaComentario)
            If .Comment Is Nothing Then
                .AddComment text:="Situacion no reconocida"
            Else
                .Comment.text text:="Situacion no reconocida"
            End If
        End With
        wsHoja.Range(celdaParedes).Value = "Error"
        wsHoja.Range(celdaSuelos).Value = "Error"
        Exit Sub
    End Select

    ' Asignar resultados y borrar comentario si exista
    wsHoja.Range(celdaParedes).Value = revestimientoParedes
    wsHoja.Range(celdaSuelos).Value = revestimientoSuelos

    If Not wsHoja.Range(celdaComentario).Comment Is Nothing Then
        wsHoja.Range(celdaComentario).Comment.Delete
    End If

End Sub

Public Sub VerificarResistencia()

    ' Constantes de celdas en hoja "Interfaz"
    Const celdaComentario As String = "F9"

    Dim wsHoja As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("RF Particiones y Puertas")
    
    If wsHoja Is Nothing Then Set wsHoja = ActiveSheet
    wsHoja.Activate

    ' Variables de búsqueda
    Dim configBuscado As String
    Dim nriBuscado As String
    Dim rasanteBuscado As String
    Dim configActual As String
    Dim nriActual As String
    Dim rasanteActual As String

    ' Resultados
    Dim resistencia As String
    Dim resistenciaCorregida As String
    resistenciaCorregida = ""
    Dim PSV As String, PCV As String, PGSV As String, PGCV As String

    ' Control
    Dim maximaSeguridad As Boolean
    Dim encontrado As Boolean
    Dim fila As Long
    Dim ultimaFila As Long
    Dim errores As String

    ' Validación de entradas
    errores = ""
    If wsHoja.Range(celdaConfig).Value = "" Then errores = errores & "- Falta tipo de configuracion" & vbCrLf
    If IsError(wsHoja.Range(celdaNri).Value) Then errores = errores & "- Falta NRI" & vbCrLf
    If wsHoja.Range(celdaRasante).Value = "" Then errores = errores & "- Falta tipo de rasante" & vbCrLf

    ' Comentario con errores
    With wsHoja.Range(celdaComentario)
        If .Comment Is Nothing Then
            If errores <> "" Then .AddComment text:=errores
        Else
            If errores <> "" Then
                .Comment.text text:=errores
            Else
                .Comment.Delete
            End If
        End If
    End With

    If errores <> "" Then
        wsHoja.Range(celdaResistencia).Value = "Error"
        wsHoja.Range(celdaPSV).Value = "Error"
        wsHoja.Range(celdaPCV).Value = "Error"
        wsHoja.Range(celdaPGSV).Value = "Error"
        wsHoja.Range(celdaPGCV).Value = "Error"
        Exit Sub
    End If

    ' Captura de entradas
    configBuscado = Trim$(wsHoja.Range(celdaConfig).Value)
    nriBuscado = Trim$(wsHoja.Range(celdaNri).Value)
    rasanteBuscado = Trim$(wsHoja.Range(celdaRasante).Value)
    maximaSeguridad = (wsHoja.Range(celdaMaxSeguridad).Value = True)

    ' Búsqueda estándar
    ultimaFila = 89
    encontrado = False

    For fila = 10 To ultimaFila
        configActual = Trim$(ws.Cells(fila, 2).Value)
        nriActual = Trim$(ws.Cells(fila, 4).Value)
        rasanteActual = Trim$(ws.Cells(fila, 3).Value)

        If configActual = configBuscado And nriActual = nriBuscado And rasanteActual = rasanteBuscado Then
            resistencia = ws.Cells(fila, 5).Value
            PSV = ws.Cells(fila, 6).Value
            PCV = ws.Cells(fila, 7).Value
            PGSV = ws.Cells(fila, 8).Value
            PGCV = ws.Cells(fila, 9).Value
            encontrado = True
            Exit For
        End If
    Next fila

    ' Búsqueda reforzada si aplica (solo ajusta resistencia)
    If maximaSeguridad Then
        ultimaFila = 126
        For fila = 95 To ultimaFila
            configActual = Trim$(ws.Cells(fila, 2).Value)
            nriActual = Trim$(ws.Cells(fila, 4).Value)
            rasanteActual = Trim$(ws.Cells(fila, 3).Value)

            If configActual = configBuscado And nriActual = nriBuscado And rasanteActual = rasanteBuscado Then
                resistenciaCorregida = ws.Cells(fila, 5).Value
                encontrado = True
                Exit For
            End If
        Next fila
    End If

    ' Salida
    If encontrado Then
        ' Resistencia base
        wsHoja.Range(celdaResistencia).Value = resistencia

        ' Resistencia viable (si tienes esa celda definida)
        If resistenciaCorregida <> "" Then
            wsHoja.Range(celdaResistenciaViable).Value = resistencia
            wsHoja.Range(celdaResistencia).Value = resistenciaCorregida
        Else
            wsHoja.Range(celdaResistenciaViable).Value = resistencia
        End If

        wsHoja.Range(celdaPSV).Value = PSV
        wsHoja.Range(celdaPCV).Value = PCV
        wsHoja.Range(celdaPGSV).Value = PGSV
        wsHoja.Range(celdaPGCV).Value = PGCV

        If Not wsHoja.Range(celdaComentario).Comment Is Nothing Then
            wsHoja.Range(celdaComentario).Comment.Delete
        End If
    Else
        wsHoja.Range(celdaResistencia).Value = "Error"
        wsHoja.Range(celdaPSV).Value = "Error"
        wsHoja.Range(celdaPCV).Value = "Error"
        wsHoja.Range(celdaPGSV).Value = "Error"
        wsHoja.Range(celdaPGCV).Value = "Error"

        With wsHoja.Range(celdaComentario)
            If .Comment Is Nothing Then
                .AddComment text:="No se encontro ninguna coincidencia con esos criterios, es posible que con las variables introducidas no haya elementos viables"
            Else
                .Comment.text text:="No se encontro ninguna coincidencia con esos criterios, es posible que con las variables introducidas no haya elementos viables"
            End If
        End With
    End If

End Sub

Public Sub VerificarResistenciaSeparadores()

    ' Constantes de celdas en hoja "Interfaz"
    Const celdaComentario As String = "F24"
    Dim wsHoja As Worksheet
    ' --- Hoja de trabajo: si no llega, uso la activa (Interfaz o 001/002/...)
    If wsHoja Is Nothing Then Set wsHoja = ActiveSheet
    wsHoja.Activate

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Resist. Estable")

    ' Variables de bsqueda
    Dim configBuscado As String
    Dim nriBuscado As String
    Dim configActual As String
    Dim nriActual As String

    ' Variables de resultados
    Dim resistencia As String

    ' Variables de control
    Dim distancia As Boolean
    Dim encontrado As Boolean
    Dim fila As Long
    Dim ultimaFila As Long
    Dim errores As String

    errores = ""

    ' Validacion de entradas acumulada
    If wsHoja.Range(celdaConfig).Value = "" Then errores = errores & "- Falta tipo de configuracion" & vbCrLf
    If IsError(wsHoja.Range(celdaNri).Value) Then errores = errores & "- Falta NRI" & vbCrLf

    ' Gestion comentario de errores
    With wsHoja.Range(celdaComentario)
        If .Comment Is Nothing Then
            If errores <> "" Then .AddComment text:=errores
        Else
            If errores <> "" Then
                .Comment.text text:=errores
            Else
                .Comment.Delete
            End If
        End If
    End With

    ' Si hay errores, poner "Error" y salir
    If errores <> "" Then
        wsHoja.Range(celdaResultadoResistencia).Value = "Error"
        Exit Sub
    End If

    ' Captura de entradas
    configBuscado = Trim(wsHoja.Range(celdaConfig).Value)
    nriBuscado = Trim(wsHoja.Range(celdaNri).Value)
    distancia = (wsHoja.Range(celdaDistancia).Value = True)

    ' Busqueda en la hoja de datos
    ultimaFila = 53
    encontrado = False
    For fila = 14 To ultimaFila
        configActual = Trim(ws.Cells(fila, 2).Value)
        nriActual = Trim(ws.Cells(fila, 3).Value)

        If configActual = configBuscado And nriActual = nriBuscado Then
            If distancia Then
                resistencia = ws.Cells(fila, 6).Value
            Else
                resistencia = ws.Cells(fila, 4).Value
            End If
            encontrado = True
            Exit For
        End If
    Next fila

    ' Mostrar resultado o error si no encontrado
    If encontrado Then
        wsHoja.Range(celdaResultadoResistencia).Value = resistencia
        ' Borrar comentario si exista
        If Not wsHoja.Range(celdaComentario).Comment Is Nothing Then
            wsHoja.Range(celdaComentario).Comment.Delete
        End If
    Else
        wsHoja.Range(celdaResultadoResistencia).Value = "Error"
        With wsHoja.Range(celdaComentario)
            If .Comment Is Nothing Then
                .AddComment text:="Esta configuracion no admite calculo de resistencia de los elementos separadores porque no esta admitida"
            Else
                .Comment.text text:="Esta configuracion no admite calculo de resistencia de los elementos separadores porque no esta admitida"
            End If
        End With
    End If

End Sub

Public Sub reaccionElementos()

    ' Constantes de celdas
    Const celdaComentario As String = "F25"
    Dim wsHoja As Worksheet
    ' --- Hoja de trabajo: si no llega, uso la activa (Interfaz o 001/002/...)
    If wsHoja Is Nothing Then Set wsHoja = ActiveSheet
    wsHoja.Activate

    ' Variables de entrada
    Dim alturaEva As Double

    ' Variables de salida
    Dim fachada As String
    Dim camara As String

    Dim errores As String
    errores = ""

    ' Validacion de entrada
    If wsHoja.Range(celdaAltura).Value = "" Then errores = "- Debe introducir altura de evacuacion" & vbCrLf

    ' Gestion comentario de errores
    With wsHoja.Range(celdaComentario)
        If .Comment Is Nothing Then
            If errores <> "" Then
                .AddComment text:=errores
            End If
        Else
            If errores <> "" Then
                .Comment.text text:=errores
            Else
                .Comment.Delete
            End If
        End If
    End With

    ' Si hay error, marcar error en resultados y salir
    If errores <> "" Then
        wsHoja.Range(celdaResultadoFachada).Value = "Error"
        wsHoja.Range(celdaResultadoCamara).Value = "Error"
        Exit Sub
    End If

    ' Captura de valor
    alturaEva = CDbl(wsHoja.Range(celdaAltura).Value)

    ' Asignacin de clase de reaccin al fuego segn altura
    Select Case True
    Case alturaEva <= 10
        fachada = "D-s3,d0"
        camara = "D-s3,d0"
    Case alturaEva <= 18
        fachada = "C-s3,d0"
        camara = "B-s3,d0"
    Case alturaEva <= 28
        fachada = "B-s3,d0"
        camara = "B-s3,d0"
    Case Else
        fachada = "B-s3,d0"
        camara = "A2-s3,d0"
    End Select

    ' Mostrar resultados
    wsHoja.Range(celdaResultadoFachada).Value = fachada
    wsHoja.Range(celdaResultadoCamara).Value = camara

End Sub

Public Sub ocupacion()

    ' Constantes de celdas
    Const celdaComentario As String = "F28"      ' Celda para comentarios/errores
    Dim wsHoja As Worksheet
    ' --- Hoja de trabajo: si no llega, uso la activa (Interfaz o 001/002/...)
    If wsHoja Is Nothing Then Set wsHoja = ActiveSheet
    wsHoja.Activate

    ' Variable de entrada y clculo
    Dim personas As Double
    Dim errores As String
    errores = ""

    ' Validacion de entrada
    If wsHoja.Range(celdaPersonasEntrada).Value = "" Then
        errores = "- Debe introducir el numero de personas que se espera que ocupen el espacio" & vbCrLf
    End If

    ' Gestion comentario
    With wsHoja.Range(celdaComentario)
        If .Comment Is Nothing Then
            If errores <> "" Then
                .AddComment text:=errores
            End If
        Else
            If errores <> "" Then
                .Comment.text text:=errores
            Else
                .Comment.Delete
            End If
        End If
    End With

    ' Si hay error, marcar error en resultado y salir
    If errores <> "" Then
        wsHoja.Range(celdaPersonasSalida).Value = "Error"
        Exit Sub
    End If

    ' Captura de valor
    personas = CDbl(wsHoja.Range(celdaPersonasEntrada).Value)

    ' Ajuste segn tramos con factores multiplicadores
    Select Case True
    Case personas > 0 And personas < 100
        personas = 1.1 * personas
    Case personas >= 100 And personas < 200
        personas = 110 + 1.05 * (personas - 100)
    Case personas >= 200 And personas < 500
        personas = 215 + 1.03 * (personas - 200)
    Case personas >= 500
        personas = 524 + 1.01 * (personas - 500)
    Case Else
        ' En caso de 0 o valores negativos, se deja como est o se podra avisar
    End Select

    ' Escribir resultado en hoja
    wsHoja.Range(celdaPersonasSalida).Value = personas

End Sub

Public Sub calcularSalidas()

    ' === Constantes solo de este procedimiento (ajusta si cambian) ===
    Const celdaFactor1 As String = "B30"
    Const celdaFactor2 As String = "B31"
    Const celdaFactor3 As String = "B32"
    Const celdaComentario As String = "F29"   ' Comentarios/errores

    ' NOTA: Se asume que en tu proyecto ya existen:
    '  celdaNri, celdaSuperficie, celdaOcupacion, celdaDistanciaIntroducida, celdaSalidasIntroducidas,
    '  celdaSalidaCantidad, celdaSalida1Distancia, celdaSalida2SinAlternativa, celdaSalida2ConAlternativa, celdaResultado

    Dim wsHoja As Worksheet
    If wsHoja Is Nothing Then Set wsHoja = ActiveSheet
    wsHoja.Activate

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Recorridos Evacuacion")

    ' Entradas
    Dim nriBuscado As String
    Dim superficie As Double
    Dim ocupacion As Double
    Dim distanciaIntroducida As Double

    ' Control
    Dim errores As String
    Dim encontrado As Boolean
    Dim fila As Long, ultimaFila As Long
    Dim factor As Double
    Dim distanciaResultado As Double

    errores = ""
    encontrado = False
    factor = 1
    distanciaResultado = 0

    ' === Validación básica de entradas ===
    If IsError(wsHoja.Range(celdaNri).Value) Or Trim$(CStr(wsHoja.Range(celdaNri).Value)) = "" Then _
        errores = errores & "- Falta NRI" & vbCrLf

    If Trim$(CStr(wsHoja.Range(celdaSuperficie).Value)) = "" Or Not IsNumeric(wsHoja.Range(celdaSuperficie).Value) Then _
        errores = errores & "- Falta superficie" & vbCrLf

    If Trim$(CStr(wsHoja.Range(celdaDistanciaIntroducida).Value)) = "" Or Not IsNumeric(wsHoja.Range(celdaDistanciaIntroducida).Value) Then _
        errores = errores & "- Falta distancia" & vbCrLf

    ' Gestiona comentario de errores
    With wsHoja.Range(celdaComentario)
        If errores <> "" Then
            If .Comment Is Nothing Then
                .AddComment text:=errores
            Else
                .Comment.text text:=errores
            End If
        Else
            If Not .Comment Is Nothing Then .ClearComments
        End If
    End With

    ' Si hay errores, deja salida coherente y termina
    If errores <> "" Then
        wsHoja.Range(celdaSalidaCantidad).Value = "Error"
        wsHoja.Range(celdaResultado).ClearContents
        wsHoja.Range(celdaSalida1Distancia).ClearContents
        wsHoja.Range(celdaSalida2SinAlternativa).ClearContents
        wsHoja.Range(celdaSalida2ConAlternativa).ClearContents
        Exit Sub
    End If

    ' === Captura defensiva de valores ===
    nriBuscado = Trim$(CStr(wsHoja.Range(celdaNri).Value))

    If IsNumeric(wsHoja.Range(celdaSuperficie).Value) Then
        superficie = CDbl(wsHoja.Range(celdaSuperficie).Value)
    Else
        superficie = 0
    End If

    If IsNumeric(wsHoja.Range(celdaOcupacion).Value) Then
        ocupacion = CDbl(wsHoja.Range(celdaOcupacion).Value)
    Else
        ocupacion = 0
    End If

    If IsNumeric(wsHoja.Range(celdaDistanciaIntroducida).Value) Then
        distanciaIntroducida = CDbl(wsHoja.Range(celdaDistanciaIntroducida).Value)
    Else
        distanciaIntroducida = 0
    End If

    Dim flagDosSalidasHoja As Boolean
    flagDosSalidasHoja = (wsHoja.Range(celdaSalidasIntroducidas).Value = True)

    ' === Factor multiplicador (no tocar aquí el flag de salidas) ===
    factor = 1
    If wsHoja.Range(celdaFactor1).Value = True Then factor = factor * 1.25
    If wsHoja.Range(celdaFactor2).Value = True Then factor = factor * 1.25

    Dim factor3Activo As Boolean
    factor3Activo = (wsHoja.Range(celdaFactor3).Value = True)
    If factor3Activo Then factor = factor * 2

    ' === Decisión consolidada: ¿1 salida o 2 salidas? ===
    Dim dosSalidas As Boolean
    dosSalidas = False

    ' Regla 1: superficie > 50 => dos salidas
    If superficie > 50 Then dosSalidas = True

    ' Regla 2: ocupación > 50 Y exista valor en col 3 para ese NRI => dos salidas
    Dim hayCol3ParaNri As Boolean
    hayCol3ParaNri = False
    ultimaFila = 24
    For fila = 17 To ultimaFila
        If Trim$(CStr(ws.Cells(fila, 1).Value)) = nriBuscado Then
            If UCase$(Trim$(CStr(ws.Cells(fila, 3).Value))) <> "N/A" And CStr(ws.Cells(fila, 3).Value) <> "" Then
                hayCol3ParaNri = True
                Exit For
            End If
        End If
    Next fila
    If ocupacion > 50 And hayCol3ParaNri Then dosSalidas = True

    ' Regla 3: Factor3 fuerza dos salidas
    If factor3Activo Then dosSalidas = True

    ' Regla 4: Flag manual de la hoja
    If flagDosSalidasHoja Then dosSalidas = True

    ' === Búsqueda de fila de criterios (tabla superior: filas 4..11) ===
    encontrado = False
    ultimaFila = 11
    For fila = 4 To ultimaFila
        If Trim$(CStr(ws.Cells(fila, 1).Value)) = nriBuscado Then
            encontrado = True

            If Not dosSalidas Then
                ' -------- CASO 1 SALIDA --------
                wsHoja.Range(celdaSalidaCantidad).Value = 1

                distanciaResultado = 0
                If IsNumeric(ws.Cells(fila, 2).Value) Then
                    distanciaResultado = CDbl(ws.Cells(fila, 2).Value) * factor
                End If
                If distanciaResultado > 90 Then distanciaResultado = 90

                wsHoja.Range(celdaSalida1Distancia).Value = distanciaResultado
                wsHoja.Range(celdaSalida2SinAlternativa).Value = 0
                wsHoja.Range(celdaSalida2ConAlternativa).Value = 0

                If distanciaResultado > distanciaIntroducida Then
                    wsHoja.Range(celdaResultado).Value = "ADMITIDO"
                Else
                    wsHoja.Range(celdaResultado).Value = "NO ADMITIDO"
                End If

            Else
                ' -------- CASO 2 SALIDAS --------
                wsHoja.Range(celdaSalidaCantidad).Value = 2
                wsHoja.Range(celdaSalida1Distancia).Value = 0

                distanciaResultado = 0
                If IsNumeric(ws.Cells(fila, 3).Value) Then
                    distanciaResultado = CDbl(ws.Cells(fila, 3).Value) * factor
                End If
                If distanciaResultado > 90 Then distanciaResultado = 90
                wsHoja.Range(celdaSalida2SinAlternativa).Value = distanciaResultado

                Dim dAlt As Double
                dAlt = 0
                If IsNumeric(ws.Cells(fila, 4).Value) Then
                    dAlt = CDbl(ws.Cells(fila, 4).Value) * factor
                End If
                If dAlt > 90 Then dAlt = 90
                wsHoja.Range(celdaSalida2ConAlternativa).Value = dAlt

                If distanciaResultado > distanciaIntroducida Then
                    wsHoja.Range(celdaResultado).Value = "ADMITIDO"
                Else
                    wsHoja.Range(celdaResultado).Value = "NO ADMITIDO"
                End If
            End If

            Exit For
        End If
    Next fila

    ' === Post-procesado de comentarios ===
    If Not encontrado Then
        errores = "- No se encontró ninguna coincidencia con esos criterios."
        With wsHoja.Range(celdaComentario)
            If .Comment Is Nothing Then
                .AddComment text:=errores
            Else
                .Comment.text text:=errores
            End If
        End With
        wsHoja.Range(celdaSalidaCantidad).Value = "Error"
        wsHoja.Range(celdaSalida1Distancia).ClearContents
        wsHoja.Range(celdaSalida2SinAlternativa).ClearContents
        wsHoja.Range(celdaSalida2ConAlternativa).ClearContents
    Else
        With wsHoja.Range(celdaComentario)
            If Not .Comment Is Nothing Then .ClearComments
        End With
    End If

End Sub


Public Sub calcularResistenciaEstructural()

    ' Constantes de celdas en hoja "Interfaz"
    Const celdaComentario As String = "F46"      ' Celda para comentarios/errores
    Dim wsHoja As Worksheet
    ' --- Hoja de trabajo: si no llega, uso la activa (Interfaz o 001/002/...)
    If wsHoja Is Nothing Then Set wsHoja = ActiveSheet
    wsHoja.Activate

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Resistencia Estructural")

    ' Variables de entrada
    Dim configBuscado As String
    Dim nriBuscado As String
    Dim rasanteBuscado As String
    Dim cubiertaEvacua As Boolean
    Dim cubiertaLigera As Boolean

    ' Variables de bsqueda
    Dim configActual As String
    Dim nriActual As String
    Dim rasanteActual As String
    Dim resistencia As String
    resistencia = ""

    ' Variables de control
    Dim fila As Long
    Dim ultimaFila As Long

    ' Validacion de entradas
    Dim errores As String
    errores = ""

    If wsHoja.Range(celdaConfig).Value = "" Then errores = errores & "- Falta tipo de configuracion" & vbCrLf
    If IsError(wsHoja.Range(celdaNri).Value) Then errores = errores & "- Falta NRI" & vbCrLf
    If wsHoja.Range(celdaRasante).Value = "" Then errores = errores & "- Falta tipo de rasante" & vbCrLf

    ' Gestion comentario para errores
    With wsHoja.Range(celdaComentario)
        If errores <> "" Then
            If .Comment Is Nothing Then
                .AddComment text:=errores
            Else
                .Comment.text text:=errores
            End If
            ' Poner "Error" en resultado y salir
            wsHoja.Range(celdaResultadoResistenciaEstructural).Value = "Error"
            Exit Sub
        Else
            If Not .Comment Is Nothing Then .Comment.Delete
        End If
    End With

    ' Captura de entradas
    configBuscado = Trim(wsHoja.Range(celdaConfig).Value)
    nriBuscado = Trim(wsHoja.Range(celdaNri).Value)
    rasanteBuscado = Trim(wsHoja.Range(celdaRasante).Value)
    cubiertaEvacua = (wsHoja.Range(celdaCubiertaEvacua).Value = True)
    cubiertaLigera = (wsHoja.Range(celdaCubiertaLigera).Value = True)

    ' Busqueda segn tipo de cubierta
    If Not cubiertaLigera And Not cubiertaEvacua Then
        ' Busqueda normal (sin cubiertas especiales)
        ultimaFila = 82
        For fila = 3 To ultimaFila
            configActual = Trim(ws.Cells(fila, 2).Value)
            rasanteActual = Trim(ws.Cells(fila, 3).Value)
            nriActual = Trim(ws.Cells(fila, 4).Value)

            If configActual = configBuscado And _
               nriActual = nriBuscado And _
               rasanteActual = rasanteBuscado Then
                resistencia = ws.Cells(fila, 5).Value
                Exit For
            End If
        Next fila

    ElseIf cubiertaLigera Then
        ' Busqueda para cubiertas ligeras
        ultimaFila = 165
        For fila = 134 To ultimaFila
            configActual = Trim(ws.Cells(fila, 2).Value)
            nriActual = Trim(ws.Cells(fila, 3).Value)

            If configActual = configBuscado And nriActual = nriBuscado Then
                resistencia = ws.Cells(fila, 4).Value
                Exit For
            End If
        Next fila

    ElseIf cubiertaEvacua Then
        ' Busqueda para cubiertas con evacuacin
        ultimaFila = 127
        For fila = 88 To ultimaFila
            configActual = Trim(ws.Cells(fila, 2).Value)
            nriActual = Trim(ws.Cells(fila, 3).Value)

            If configActual = configBuscado And nriActual = nriBuscado Then
                resistencia = ws.Cells(fila, 4).Value
                Exit For
            End If
        Next fila
    End If

    ' Mostrar resultado o error
    If resistencia = "" Then
        ' No se encontr resultado, aadir comentario y poner "Error"
        errores = "- No se encontr resistencia para esa combinacion."
        With wsHoja.Range(celdaComentario)
            If .Comment Is Nothing Then
                .AddComment text:=errores
            Else
                .Comment.text text:=errores
            End If
        End With
        wsHoja.Range(celdaResultadoResistenciaEstructural).Value = "Error"
    Else
        ' Resultado encontrado, eliminar comentario si existe
        wsHoja.Range(celdaResultadoResistenciaEstructural).Value = resistencia
        With wsHoja.Range(celdaComentario)
            If Not .Comment Is Nothing Then .Comment.Delete
        End With
    End If

End Sub

Public Sub calcularSistemasDeteccion()

    ' Constantes de celdas en hoja "Interfaz"

    Const celdaComentario As String = "A51"
    Dim wsHoja As Worksheet
    ' --- Hoja de trabajo: si no llega, uso la activa (Interfaz o 001/002/...)
    If wsHoja Is Nothing Then Set wsHoja = ActiveSheet
    wsHoja.Activate
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Deteccion")

    ' Variables de entrada
    Dim configBuscado As String
    Dim nriBuscado As String
    Dim actividadBuscada As String
    Dim superficie As Double
    Dim ocupacion As Variant
    Dim ocupacionSuperficie As Double
    Dim superficieEstablecimiento As Double

    ' Variables de comparacin
    Dim configActual As String
    Dim nriActual As String
    Dim actividadActual As String
    Dim fila As Long
    Dim ultimaFila As Long

    ' Inicializacion de resultados (limpiar)
    wsHoja.Range(celdaDeteccionAuto).Value = "No"
    wsHoja.Range(celdaDeteccionManual).Value = "No"
    wsHoja.Range(celdaMegafonia).Value = "No"

    ' Validacion de entradas
    Dim errores As String
    errores = ""

    If wsHoja.Range(celdaConfig).Value = "" Then errores = errores & "- Falta tipo de configuracion" & vbCrLf
    If IsError(wsHoja.Range(celdaNri).Value) Then errores = errores & "- Falta NRI" & vbCrLf
    If wsHoja.Range(celdaSuperficie).Value = "" Then errores = errores & "- Falta superficie" & vbCrLf
    If wsHoja.Range(celdaActividad).Value = "" Then errores = errores & "- Falta tipo de actividad" & vbCrLf
    If wsHoja.Range(celdaSuperficieEstablecimiento).Value = "" Then errores = errores & "- Falta superficie total del establecimiento" & vbCrLf
    If wsHoja.Range(celdaOcupacion).Value = "" Then errores = errores & "- Falta ocupacion" & vbCrLf
    
    
    ocupacion = wsHoja.Range(celdaOcupacion).Value 'Comprobacion rara para ocupacion
    
    If Not IsNumeric(ocupacion) Then
        errores = errores & "Ocupacion debe ser un valor numerico" & vbCrLf
    End If
    
    ' Gestion comentarios para errores
    With wsHoja.Range(celdaComentario)
        If errores <> "" Then
            If .Comment Is Nothing Then
                .AddComment text:=errores
            Else
                .Comment.text text:=errores
            End If
            ' Poner "Error" en los resultados y salir
            wsHoja.Range(celdaDeteccionAuto).Value = "Error"
            wsHoja.Range(celdaDeteccionManual).Value = "Error"
            wsHoja.Range(celdaMegafonia).Value = "Error"
            Exit Sub
        Else
            If Not .Comment Is Nothing Then .Comment.Delete
        End If
    End With

    ' Captura de valores
    configBuscado = Trim(wsHoja.Range(celdaConfig).Value)
    nriBuscado = Trim(wsHoja.Range(celdaNri).Value)
    actividadBuscada = Trim(wsHoja.Range(celdaActividad).Value)
    superficie = CDbl(wsHoja.Range(celdaSuperficie).Value)
    superficieEstablecimiento = CDbl(wsHoja.Range(celdaSuperficieEstablecimiento).Value)
    ocupacion = wsHoja.Range(celdaOcupacion).Value
    
    If Not IsNumeric(ocupacion) Then
        MsgBox "El valor de la ocupacion debe ser un numero"
        Exit Sub
    End If
    
    ' Calculo de ocupacin por superficie
    If superficie > 0 Then
        ocupacionSuperficie = ocupacion / superficie
    Else
        ocupacionSuperficie = 0
    End If

    ' Busqueda de coincidencia
    ultimaFila = 82
    For fila = 3 To ultimaFila
        configActual = Trim(ws.Cells(fila, 2).Value)
        nriActual = Trim(ws.Cells(fila, 3).Value)
        actividadActual = Trim(ws.Cells(fila, 4).Value)

        If configActual = configBuscado And _
           nriActual = nriBuscado And _
           actividadActual = actividadBuscada Then

            ' Evaluacin deteccin manual
            If superficie >= 400 Then
      
                wsHoja.Range(celdaDeteccionManual).Value = "Si"
            Else
                wsHoja.Range(celdaDeteccionManual).Value = "No"
   
            End If

            ' Evaluacin megafona
            If superficieEstablecimiento >= 10000 And ocupacionSuperficie >= 3 Then
                wsHoja.Range(celdaMegafonia).Value = "Si"
            Else
                wsHoja.Range(celdaMegafonia).Value = "No"
            End If

            ' Evaluacin deteccin automtica
            If IsNumeric(ws.Cells(fila, 5).Value) And superficie >= ws.Cells(fila, 5).Value Then
                wsHoja.Range(celdaDeteccionAuto).Value = ws.Cells(fila, 6).Value
            Else
                wsHoja.Range(celdaDeteccionAuto).Value = "No"
            End If

            Exit For
        End If
    Next fila

End Sub

Public Sub calcularHidrantes()
    ' Constantes de celdas en hoja "Interfaz"
    Const celdaComentario As String = "D51"
    Dim wsHoja As Worksheet
    ' --- Hoja de trabajo: si no llega, uso la activa (Interfaz o 001/002/...)
    If wsHoja Is Nothing Then Set wsHoja = ActiveSheet
    wsHoja.Activate

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Hidrantes")

    ' Variables de entrada
    Dim configBuscado As String
    Dim nriBuscado As String
    Dim superficie As Variant
    Dim superficieTotal As Variant

    ' Variables de bsqueda
    Dim configActual As String
    Dim nriActual As String
    Dim fila As Long
    Dim ultimaFila As Long

    ' Variable para acumular textos de errores
    Dim errores As String
    errores = ""

    ' Inicializacion (limpieza de resultados y comentarios)
    With wsHoja
        .Range(celdaComentario).ClearComments
        .Range(celdaCaudalMinimo).Value = ""
        .Range(celdaAutonomia).Value = ""
        .Range(celdaHidrantesCamiones).Value = ""
        .Range(celdaHidrantesDirecta).Value = ""
    End With

    ' Validacion de entradas: comprobamos cada dato y vamos concatenando mensajes
    If Trim(wsHoja.Range(celdaConfig).Value) = "" Then errores = errores & "- Falta tipo de configuracion" & vbCrLf
    If IsError(wsHoja.Range(celdaNri).Value) Then errores = errores & "- Falta NRI" & vbCrLf
    If wsHoja.Range(celdaSuperficie).Value = "" Or Not IsNumeric(wsHoja.Range(celdaSuperficie).Value) Then errores = errores & "- Superficie no valida" & vbCrLf
    
    ' Si hay errores, escribimos "Error" en resultados y aadimos comentario, luego salimos
    With wsHoja
        If errores <> "" Then
            .Range(celdaHidrantesCamiones).Value = "Error"
            .Range(celdaHidrantesDirecta).Value = "Error"
            .Range(celdaCaudalMinimo).Value = "Error"
            .Range(celdaAutonomia).Value = "Error"

            ' Aadir o actualizar comentario con errores
            If .Range(celdaComentario).Comment Is Nothing Then
                .Range(celdaComentario).AddComment text:=errores
            Else
                .Range(celdaComentario).Comment.text text:=errores
            End If

            Exit Sub
        Else
            ' Si no hay errores, eliminar comentario si existiera
            If Not .Range(celdaHidrantesCamiones).Comment Is Nothing Then
                .Range(celdaHidrantesCamiones).Comment.Delete
            End If
        End If
    End With

    ' Lectura de valores
    configBuscado = Trim(wsHoja.Range(celdaConfig).Value)
    nriBuscado = Trim(wsHoja.Range(celdaNri).Value)
    superficie = wsHoja.Range(celdaSuperficie).Value
    superficieTotal = wsHoja.Range(celdaSuperficieTotal).Value
    
    ' -------------------------
    ' Busqueda de hidrantes exteriores para camiones (bloque 1)
    ' -------------------------
    ultimaFila = 91
    For fila = 3 To ultimaFila
        configActual = Trim(ws.Cells(fila, 2).Value)
        nriActual = Trim(ws.Cells(fila, 3).Value)

        If configActual = configBuscado And nriActual = nriBuscado Then

            If superficieTotal >= 5000 Then
                wsHoja.Range(celdaHidrantesCamiones).Value = ws.Cells(fila, 7).Value
            End If

            If superficie >= ws.Cells(fila, 4).Value And wsHoja.Range(celdaHidrantesCamiones).Value <> "Si" Then
                wsHoja.Range(celdaHidrantesCamiones).Value = ws.Cells(fila, 5).Value
                Exit For
            End If
        End If
    Next fila
    If wsHoja.Range(celdaHidrantesCamiones).Value = "" Then wsHoja.Range(celdaHidrantesCamiones).Value = "No"

    ' -------------------------
    ' Busqueda de hidrantes por alimentacin directa (bloque 2)
    ' -------------------------
    ultimaFila = 159
    For fila = 96 To ultimaFila
        configActual = Trim(ws.Cells(fila, 2).Value)
        nriActual = Trim(ws.Cells(fila, 3).Value)

        If configActual = configBuscado And nriActual = nriBuscado Then
            If superficie >= ws.Cells(fila, 4).Value Then
                wsHoja.Range(celdaHidrantesDirecta).Value = ws.Cells(fila, 5).Value
            End If
        End If
    Next fila
    If wsHoja.Range(celdaHidrantesDirecta).Value = "" Then wsHoja.Range(celdaHidrantesDirecta).Value = "No"
    ' -------------------------
    ' Busqueda de caudal mnimo y autonoma (bloque 3)
    ' -------------------------
    ultimaFila = 203
    For fila = 164 To ultimaFila
        configActual = Trim(ws.Cells(fila, 2).Value)
        nriActual = Trim(ws.Cells(fila, 3).Value)

        If configActual = configBuscado And nriActual = nriBuscado Then
            wsHoja.Range(celdaCaudalMinimo).Value = ws.Cells(fila, 4).Value
            wsHoja.Range(celdaAutonomia).Value = ws.Cells(fila, 5).Value
            Exit For
        End If
    Next fila

End Sub

Public Sub calcularExtintores()

    Const celdaComentario As String = "B53"
    Dim wsHoja As Worksheet
    ' --- Hoja de trabajo: si no llega, uso la activa (Interfaz o 001/002/...)
    If wsHoja Is Nothing Then Set wsHoja = ActiveSheet
    wsHoja.Activate

    Dim nri As String
    Dim Config As String
    Dim superficie As Variant
    Dim eficacia As String
    Dim numeroExtintores As Integer

    Dim errores As String
    errores = ""


    ' Limpieza inicial de resultados y comentarios
    With wsHoja
        .Range(celdaEficaciaMinima).Value = ""
        .Range(celdaNumero).Value = ""
        .Range(celdaExtintoresPortatiles).Value = ""
        .Range(celdaComentario).ClearComments
    End With

    ' Validacion de entradas: acumulamos mensajes de error si faltan datos
    If Trim(wsHoja.Range(celdaConfig).Value) = "" Then errores = errores & "- Falta tipo de configuracion" & vbCrLf
    If IsError(wsHoja.Range(celdaNri).Value) Then errores = errores & "- Falta NRI" & vbCrLf
    If wsHoja.Range(celdaSuperficie).Value = "" Or Not IsNumeric(wsHoja.Range(celdaSuperficie).Value) Then errores = errores & "- Superficie no valida" & vbCrLf

    ' Gestion de errores: si hay errores, escribimos "Error" y comentamos, luego salimos
    With wsHoja
        If errores <> "" Then
            .Range(celdaEficaciaMinima).Value = "Error"
            .Range(celdaNumero).Value = "Error"
            .Range(celdaExtintoresPortatiles).Value = "Error"

            ' Aadir o actualizar comentario en celdaEficaciaMinima (puedes cambiarla si prefieres otra)
            If .Range(celdaComentario).Comment Is Nothing Then
                .Range(celdaComentario).AddComment text:=errores
            Else
                .Range(celdaComentario).Comment.text text:=errores
            End If

            Exit Sub
        Else
            ' Si no hay errores, borrar comentario previo
            If Not .Range(celdaComentario).Comment Is Nothing Then
                .Range(celdaComentario).Comment.Delete
            End If
        End If
    End With
    
    ' Lectura de datos
    nri = Trim(wsHoja.Range(celdaNri).Value)
    superficie = wsHoja.Range(celdaSuperficie).Value
    Config = Trim(wsHoja.Range(celdaConfig).Value)

    ' Calculo de eficacia segn NRI
    Select Case LCase(nri)
    Case "bajo 1", "bajo 2", "medio 3", "medio 4", "medio 5"
        eficacia = "21A"
    Case "alto 6", "alto 7", "alto 8"
        eficacia = "34A"
    Case Else
        ' En vez de MsgBox ponemos error en salida y comentario
        errores = "- NRI no reconocido: " & nri
        With wsHoja
            .Range(celdaEficaciaMinima).Value = "Error"
            .Range(celdaNumero).Value = "Error"
            .Range(celdaExtintoresPortatiles).Value = "Error"

            If .Range(celdaComentario).Comment Is Nothing Then
                .Range(celdaComentario).AddComment text:=errores
            Else
                .Range(celdaComentario).Comment.text text:=errores
            End If
        End With
        Exit Sub
    End Select

    ' Calculo del nmero de extintores
    Select Case LCase(nri)
    Case "bajo 1", "bajo 2"
        If superficie > 600 Then
            numeroExtintores = Application.WorksheetFunction.RoundUp((superficie - 600) / 200, 0) + 1
        Else
            numeroExtintores = 1
        End If
    Case "medio 3", "medio 4", "medio 5"
        If superficie > 400 Then
            numeroExtintores = Application.WorksheetFunction.RoundUp((superficie - 400) / 200, 0) + 1
        Else
            numeroExtintores = 1
        End If
    Case "alto 6", "alto 7", "alto 8"
        If superficie > 300 Then
            numeroExtintores = Application.WorksheetFunction.RoundUp((superficie - 300) / 200, 0) + 1
        Else
            numeroExtintores = 1
        End If
    End Select

    ' Evaluar si se requieren extintores porttiles
    If Config = "D" And LCase(nri) <> "bajo 1" Then
        wsHoja.Range(celdaExtintoresPortatiles).Value = "Si necesita extintores moviles en funcion de la geometria del establecimiento"
    Else
        wsHoja.Range(celdaExtintoresPortatiles).Value = "En caso de no existir combustibles liquidos en el establecimiento, no necesita extintores porttiles"
    End If

    ' Salida de datos
    With wsHoja
        .Range(celdaEficaciaMinima).Value = eficacia
        .Range(celdaNumero).Value = numeroExtintores
    End With

End Sub

Public Sub calcularBies()
    Const celdaComentario As String = "E53"
    Dim wsHoja As Worksheet
    ' --- Hoja de trabajo: si no llega, uso la activa (Interfaz o 001/002/...)
    If wsHoja Is Nothing Then Set wsHoja = ActiveSheet
    wsHoja.Activate

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BIES")

    Dim nri As String
    Dim Config As String
    Dim superficie As Variant
    Dim configActual As String
    Dim nriActual As String
    Dim superficieActual As Double
    Dim fila As Long, ultimaFila As Long
    Dim errores As String
    errores = ""

    ' Limpieza inicial de resultados y comentarios
    With wsHoja
        .Range(celdaNecesitaBIES).Value = ""
        .Range(celdaTipoBie).Value = ""
        .Range(celdaComentario).ClearComments
    End With

    ' Validacion de entradas
    If Trim(wsHoja.Range(celdaConfig).Value) = "" Then errores = errores & "- Falta tipo de configuracion" & vbCrLf
    If IsError(wsHoja.Range(celdaNri).Value) Then errores = errores & "- Falta NRI" & vbCrLf
    If wsHoja.Range(celdaSuperficie).Value = "" Or Not IsNumeric(wsHoja.Range(celdaSuperficie).Value) Then errores = errores & "- Superficie no valida" & vbCrLf

    ' Si hay errores, ponemos "Error" en resultados, aadimos comentario y salimos
    With wsHoja
        If errores <> "" Then
            .Range(celdaNecesitaBIES).Value = "Error"
            .Range(celdaTipoBie).Value = "Error"
            If .Range(celdaComentario).Comment Is Nothing Then
                .Range(celdaComentario).AddComment text:=errores
            Else
                .Range(celdaComentario).Comment.text text:=errores
            End If
            Exit Sub
        Else
            ' Si no hay errores, borramos comentario previo
            If Not .Range(celdaComentario).Comment Is Nothing Then
                .Range(celdaComentario).Comment.Delete
            End If
        End If
    End With

    ' Entradas limpias
    nri = Trim(wsHoja.Range(celdaNri).Value)
    superficie = CDbl(wsHoja.Range(celdaSuperficie).Value)
    Config = Trim(wsHoja.Range(celdaConfig).Value)

    ' Valores por defecto
    With wsHoja
        .Range(celdaNecesitaBIES).Value = "No"
        .Range(celdaTipoBie).Value = "No"
    End With

    ' Buscar en tabla BIES
    ultimaFila = 42                              ' Ajustar si necesario

    For fila = 3 To ultimaFila
        configActual = Trim(ws.Cells(fila, 2).Value)
        nriActual = Trim(ws.Cells(fila, 3).Value)
        superficieActual = ws.Cells(fila, 4).Value

        If configActual = Config And nriActual = nri And superficie >= superficieActual Then
            wsHoja.Range(celdaNecesitaBIES).Value = ws.Cells(fila, 5).Value
            wsHoja.Range(celdaTipoBie).Value = ws.Cells(fila, 6).Value
            Exit For
        End If
    Next fila

End Sub

Public Sub calcularColumnaSeca()
    Const celdaComentario As String = "A55"
    Dim wsHoja As Worksheet
    ' --- Hoja de trabajo: si no llega, uso la activa (Interfaz o 001/002/...)
    If wsHoja Is Nothing Then Set wsHoja = ActiveSheet
    wsHoja.Activate

    Dim Config As String
    Dim altura As Variant
    Dim errores As String
    errores = ""

    ' Limpieza inicial
    With wsHoja
        .Range(celdaComentario).ClearComments
        .Range(celdaComentario).Value = ""
    End With

    ' Validaciones de entrada
    If Trim(wsHoja.Range(celdaConfig).Value) = "" Then errores = errores & "- Falta tipo de configuracion" & vbCrLf
    If Not IsNumeric(wsHoja.Range(celdaAltura).Value) Then errores = errores & "- Altura de evacuacin no valida" & vbCrLf

    ' Gestion de errores
    With wsHoja
        If errores <> "" Then
            .Range(celdaNecesitaCol).Value = "Error"
            If .Range(celdaComentario).Comment Is Nothing Then
                .Range(celdaComentario).AddComment text:=errores
            Else
                .Range(celdaComentario).Comment.text text:=errores
            End If
            Exit Sub
        Else
            ' Si no hay errores, borramos comentario previo si existiera
            If Not .Range(celdaComentario).Comment Is Nothing Then
                .Range(celdaComentario).Comment.Delete
            End If
        End If
    End With

    ' Captura de valores
    altura = CDbl(wsHoja.Range(celdaAltura).Value)
    Config = Trim(wsHoja.Range(celdaConfig).Value)

    ' Lgica original
    wsHoja.Range(celdaNecesitaCol).Value = "No"

    If altura >= 15 Then
        wsHoja.Range(celdaNecesitaCol).Value = "Si"
    End If

    If Config = "D" Then
        wsHoja.Range(celdaNecesitaCol).Value = "No"
    End If
End Sub

Public Sub calcularExtincioAuto()
    Const celdaComentario       As String = "B55"
    ' Usa SIEMPRE la misma celda de salida para este cÃ¤lculo
    Const celdaSalida           As String = celdaNecesitaExtAuto ' <-- ya la tienes declarada como Const en tu modulo

    Dim wsHoja As Worksheet
    If wsHoja Is Nothing Then Set wsHoja = ActiveSheet
    wsHoja.Activate

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Extincion Auto")

    Dim nri As String, Config As String, actividad As String
    Dim superficie As Double

    Dim fila As Long, ultimaFila As Long
    Dim configActual As String, nriActual As String, actividadActual As String
    Dim superficieActual As Double, valorTabla As Variant
    Dim errores As String, hayMatch As Boolean

    ' === Helpers locales ===
    Dim FnNorm As Object
    Set FnNorm = CreateObject("Scripting.Dictionary")

    ' FnNorm("x") -> UCase(Trim(CStr(x)))
    ' Uso: FnNorm.Item(x)
    FnNorm.CompareMode = 1                       ' TextCompare

    ' --- Cargar entradas desde Interfaz ---
    Config = UCase$(Trim$(CStr(wsHoja.Range(celdaConfig).Value)))
    actividad = UCase$(Trim$(CStr(wsHoja.Range(celdaActividad).Value)))
    nri = UCase$(Trim$(CStr(wsHoja.Range(celdaNri).Value)))

    If IsError(wsHoja.Range(celdaSuperficie).Value) Or _
       Len(Trim$(CStr(wsHoja.Range(celdaSuperficie).Value))) = 0 Or _
       Not IsNumeric(wsHoja.Range(celdaSuperficie).Value) Then
        superficie = -1
    Else
        superficie = CDbl(wsHoja.Range(celdaSuperficie).Value)
    End If

    ' --- Limpieza inicial ---
    With wsHoja
        .Range(celdaSalida).ClearComments
        .Range(celdaSalida).Value = ""
    End With

    ' --- Validaciones ---
    If Len(Config) = 0 Then errores = errores & "- Falta tipo de configuracion" & vbCrLf
    If Len(actividad) = 0 Then errores = errores & "- Falta tipo de actividad" & vbCrLf
    If Len(nri) = 0 Then errores = errores & "- Falta NRI" & vbCrLf
    If superficie < 0 Then errores = errores & "- Superficie no valida" & vbCrLf

    With wsHoja
        If errores <> "" Then
            .Range(celdaSalida).Value = "Error"
            If .Range(celdaComentario).Comment Is Nothing Then
                .Range(celdaComentario).AddComment text:=errores
            Else
                .Range(celdaComentario).Comment.text text:=errores
            End If
            Exit Sub
        Else
            If Not .Range(celdaComentario).Comment Is Nothing Then
                .Range(celdaComentario).Comment.Delete
            End If
        End If
    End With

    ' Valor por defecto
    wsHoja.Range(celdaSalida).Value = "No"


    ultimaFila = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
    If ultimaFila < 3 Then ultimaFila = 3

    hayMatch = False

    For fila = 3 To ultimaFila
        configActual = UCase$(Trim$(CStr(ws.Cells(fila, 2).Value)))
        nriActual = UCase$(Trim$(CStr(ws.Cells(fila, 3).Value)))
        actividadActual = UCase$(Trim$(CStr(ws.Cells(fila, 4).Value)))

        ' Coincidencia por las tres claves
        If (configActual = Config) And (nriActual = nri) And (actividadActual = actividad) Then
            hayMatch = True

            If IsNumeric(ws.Cells(fila, 5).Value) Then
                superficieActual = CDbl(ws.Cells(fila, 5).Value)
                ' Regla: si la superficie introducida >= umbral de tabla -> toma el valor de la tabla (col 6)
                If superficie >= superficieActual Then
                    valorTabla = ws.Cells(fila, 6).Value
                    If Len(Trim$(CStr(valorTabla))) = 0 Then valorTabla = "Si" ' por si la tabla deja vacio y significa "aplica"
                    wsHoja.Range(celdaSalida).Value = valorTabla
                    Exit For
                Else
                    wsHoja.Range(celdaSalida).Value = "No"
                    ' Ojo: puede haber otra fila para la misma config con umbral menor;
                    '
                End If
            Else
                ' Fila
                wsHoja.Range(celdaSalida).Value = "En esta configuracion no aplica"
                
            End If
        End If
    Next fila

        
    If Not hayMatch Then
        wsHoja.Range(celdaSalida).Value = "Sin coincidencia en tabla"
    End If
End Sub

Public Sub calcularHumos()
    Const celdaComentario As String = "C55"

    Dim wsHoja As Worksheet
    If wsHoja Is Nothing Then Set wsHoja = ActiveSheet
    wsHoja.Activate

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Control Humos")

    Dim superficie As Double
    Dim nri As String, actividad As String
    Dim fila As Long, ultimaFila As Long
    Dim nriActual As String, actividadActual As String
    Dim superficieActual As Double
    Dim necesita As String
    Dim errores As String
    errores = ""

    ' --- Limpieza inicial
    With wsHoja
        .Range(celdaNecesitaHumos).ClearComments
        .Range(celdaNecesitaHumos).Value = ""
    End With

    ' --- Validaciones
    If IsError(wsHoja.Range(celdaNri).Value) Or Trim$(CStr(wsHoja.Range(celdaNri).Value)) = "" Then _
       errores = errores & "- Falta NRI" & vbCrLf
    If Trim$(CStr(wsHoja.Range(celdaActividad).Value)) = "" Then _
       errores = errores & "- Falta tipo de actividad" & vbCrLf
    If Len(Trim$(CStr(wsHoja.Range(celdaSuperficie).Value))) = 0 Or _
       Not IsNumeric(wsHoja.Range(celdaSuperficie).Value) Then _
       errores = errores & "- Superficie no valida" & vbCrLf

    ' --- Gestion de errores
    With wsHoja
        If errores <> "" Then
            .Range(celdaNecesitaHumos).Value = "Error"
            If .Range(celdaComentario).Comment Is Nothing Then
                .Range(celdaComentario).AddComment text:=errores
            Else
                .Range(celdaComentario).Comment.text text:=errores
            End If
            Exit Sub
        Else
            If Not .Range(celdaComentario).Comment Is Nothing Then
                .Range(celdaComentario).Comment.Delete
            End If
        End If
    End With

    ' --- Lectura de entradas
    nri = LCase$(Trim$(CStr(wsHoja.Range(celdaNri).Value)))
    actividad = LCase$(Trim$(CStr(wsHoja.Range(celdaActividad).Value)))
    superficie = CDbl(wsHoja.Range(celdaSuperficie).Value)

    ' Valor por defecto
    wsHoja.Range(celdaNecesitaHumos).Value = "No"

    ' Ultima fila dinamica en la columna B (NRI)
    ultimaFila = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
    If ultimaFila < 2 Then ultimaFila = 2

    For fila = 2 To ultimaFila
        nriActual = LCase$(Trim$(CStr(ws.Cells(fila, 2).Value)))
        actividadActual = LCase$(Trim$(CStr(ws.Cells(fila, 3).Value)))

        If nriActual = nri And actividadActual = actividad Then
            If IsNumeric(ws.Cells(fila, 4).Value) Then
                superficieActual = CDbl(ws.Cells(fila, 4).Value)


                If superficie < superficieActual Then

                    wsHoja.Range(celdaNecesitaHumos).Value = "No"
                    Exit Sub
                Else
                    ' Mayor o igual que el umbral -> usa el valor de la tabla (Si/No)
                    necesita = CStr(ws.Cells(fila, 5).Value)
                    If Trim$(necesita) = "" Then necesita = "Si"
                    wsHoja.Range(celdaNecesitaHumos).Value = necesita
                    Exit Sub
                End If
            Else
                wsHoja.Range(celdaNecesitaHumos).Value = "Valor no numerico en superficie"
                Exit Sub
            End If
        End If
    Next fila

    ' Si encontro NRI/actividad pero no umbral, ya habriamos salido.
    ' Si no hubo coincidencia NRI+actividad:
    wsHoja.Range(celdaNecesitaHumos).Value = _
                                           "No se ha encontrado coincidencia exacta en la tabla para ese NRI y actividad."
End Sub

Public Function CopiarDatosSup(Optional mostrarMensajes As Boolean = True) As Boolean

    Dim wsInterfaz As Worksheet
    Dim wsDestino As Worksheet
    Dim filaDestino As Long

    Dim nombre As Variant, configuracion As Variant, nri As Variant, edificio As Variant
    Dim superficie As Variant, superficieCorregida As Variant, viabilidad As Variant

    Const celdaSuperficieCorregida As String = "F8"
    Const celdaViabilidad As String = "F7"

    ' Referencias a hojas
    Set wsInterfaz = ThisWorkbook.Sheets("Interfaz")
    Set wsDestino = ThisWorkbook.Sheets("Anx. II Sup")

    ' Obtener los valores de Interfaz
    edificio = wsInterfaz.Range(celdaEdificio).Value
    nombre = wsInterfaz.Range(celdaNombre).Value
    configuracion = wsInterfaz.Range(celdaConfiguracion).Value
    nri = wsInterfaz.Range(celdaNri).Value
    superficie = wsInterfaz.Range(celdaSuperficie).Value
    superficieCorregida = wsInterfaz.Range(celdaSuperficieCorregida).Value
    viabilidad = wsInterfaz.Range(celdaViabilidad).Value

    ' Validar que todos los valores estn rellenos
    If IsEmpty(nombre) Or IsEmpty(configuracion) Or IsEmpty(nri) _
       Or IsEmpty(superficie) Or IsEmpty(superficieCorregida) Or IsEmpty(viabilidad) Then
        
        If mostrarMensajes Then
            MsgBox "Por favor, rellena todos los campos obligatorios en la hoja Interfaz antes de exportar." & vbCrLf & _
                   "Campos requeridos: Nombre, Configuracion, NRI, Superficie, Superficie Corregida, Viabilidad", vbExclamation
        End If
        
        CopiarDatosSup = False
        Exit Function
    End If

    ' Comprobar si alguno de los campos contiene la palabra "Error"
    If UCase(nombre) = "ERROR" Or UCase(configuracion) = "ERROR" Or UCase(nri) = "ERROR" _
       Or UCase(superficie) = "ERROR" Or UCase(superficieCorregida) = "ERROR" Or UCase(viabilidad) = "ERROR" Then
        
        If mostrarMensajes Then
            MsgBox "No se puede exportar porque los datos contienen errores", vbExclamation
        End If
        
        CopiarDatosSup = False
        Exit Function
    End If

    ' Buscar la primera fila realmente vaca (desde fila 2)
    filaDestino = 2
    Do While Application.WorksheetFunction.CountA(wsDestino.Rows(filaDestino)) > 0
        filaDestino = filaDestino + 1
    Loop

    ' Copiar los datos
    With wsDestino
        .Cells(filaDestino, "A").Value = edificio
        .Cells(filaDestino, "B").Value = nombre
        .Cells(filaDestino, "C").Value = configuracion
        .Cells(filaDestino, "D").Value = nri
        .Cells(filaDestino, "E").Value = superficie
        .Cells(filaDestino, "F").Value = superficieCorregida
        .Cells(filaDestino, "G").Value = viabilidad
    End With

    If mostrarMensajes Then
        MsgBox "Datos copiados correctamente en la fila " & filaDestino & " de la hoja Anx II. Sup", vbInformation
    End If

    CopiarDatosSup = True
End Function

Public Function CopiarDatosReac(Optional mostrarMensajes As Boolean = True) As Boolean

    Dim wsInterfaz As Worksheet
    Dim wsDestino As Worksheet
    Dim filaDestino As Long
    
    Dim nombre As Variant, situacion As Variant, techos As Variant
    Dim superficie As Variant, suelos As Variant, edificio As Variant
    

    ' Referencias a hojas
    Set wsInterfaz = ThisWorkbook.Sheets("Interfaz")
    Set wsDestino = ThisWorkbook.Sheets("Anx. II Reac")

    ' Obtener los valores de Interfaz
    edificio = wsInterfaz.Range(celdaEdificio).Value
    nombre = wsInterfaz.Range(celdaNombre).Value
    superficie = wsInterfaz.Range(celdaSuperficie).Value
    situacion = wsInterfaz.Range(celdaSituacion).Value
    techos = wsInterfaz.Range(celdaTechos).Value
    suelos = wsInterfaz.Range(celdaSuelos).Value

    ' Validar que todos los valores estn rellenos
    If IsEmpty(nombre) Or IsEmpty(situacion) Or IsEmpty(techos) _
       Or IsEmpty(superficie) Or IsEmpty(suelos) Then
        
        If mostrarMensajes Then
            MsgBox "Por favor, rellena todos los campos obligatorios en la hoja Interfaz antes de exportar." & vbCrLf & _
                   "Campos requeridos: Nombre, Superficie, Situacion del elemento, Reaccion al fuego de techos y paredes, Reaccion al fuego de suelos", vbExclamation
        End If
        
        CopiarDatosReac = False
        Exit Function
    End If

    ' Comprobar si alguno de los campos contiene la palabra "Error"
    If UCase(nombre) = "ERROR" Or UCase(situacion) = "ERROR" Or UCase(techos) = "ERROR" _
       Or UCase(superficie) = "ERROR" Or UCase(suelos) = "ERROR" Then
        
        If mostrarMensajes Then
            MsgBox "No se puede exportar porque los datos contienen errores", vbExclamation
        End If
        
        CopiarDatosReac = False
        Exit Function
    End If

    ' Buscar la primera fila realmente vaca (desde fila 2)
    filaDestino = 2
    Do While Application.WorksheetFunction.CountA(wsDestino.Rows(filaDestino)) > 0
        filaDestino = filaDestino + 1
    Loop

    ' Copiar los datos
    With wsDestino
        .Cells(filaDestino, "A").Value = edificio
        .Cells(filaDestino, "B").Value = nombre
        .Cells(filaDestino, "C").Value = superficie
        .Cells(filaDestino, "D").Value = situacion
        .Cells(filaDestino, "E").Value = techos
        .Cells(filaDestino, "F").Value = suelos
    End With

    If mostrarMensajes Then
        MsgBox "Datos copiados correctamente en la fila " & filaDestino & " de la hoja Anx II. Reac", vbInformation
    End If

    CopiarDatosReac = True
End Function

Public Function CopiarDatosResSec(Optional mostrarMensajes As Boolean = True) As Boolean

    Dim wsInterfaz As Worksheet
    Dim wsDestino As Worksheet
    Dim filaDestino As Long

    Dim edificio As Variant, nombre As Variant, configuracion As Variant, nri As Variant
    Dim superficie As Variant, resistencia As Variant

    ' Referencias a hojas
    Set wsInterfaz = ThisWorkbook.Sheets("Interfaz")
    Set wsDestino = ThisWorkbook.Sheets("Anx. II Res Sec")

    ' Obtener los valores de Interfaz
    edificio = wsInterfaz.Range(celdaEdificio).Value
    nombre = wsInterfaz.Range(celdaNombre).Value
    configuracion = wsInterfaz.Range(celdaConfiguracion).Value
    nri = wsInterfaz.Range(celdaNri).Value
    superficie = wsInterfaz.Range(celdaSuperficie).Value
    resistencia = wsInterfaz.Range(celdaResistencia).Value

    ' Validar que todos los valores estn rellenos
    If IsEmpty(nombre) Or IsEmpty(configuracion) Or IsEmpty(nri) _
       Or IsEmpty(superficie) Or IsEmpty(resistencia) Then

        If mostrarMensajes Then
            MsgBox "Por favor, rellena todos los campos obligatorios en la hoja Interfaz antes de exportar." & vbCrLf & _
                   "Campos requeridos: Nombre, Configuracion, NRI, Superficie, Resistencia de elementos separadores de sectores", vbExclamation
        End If

        CopiarDatosResSec = False
        Exit Function
    End If

    ' Comprobar si alguno de los campos contiene la palabra "Error"
    If UCase(nombre) = "ERROR" Or UCase(configuracion) = "ERROR" Or UCase(nri) = "ERROR" _
       Or UCase(superficie) = "ERROR" Or UCase(resistencia) = "ERROR" Then
        
        If mostrarMensajes Then
            MsgBox "No se puede exportar porque los datos contienen errores", vbExclamation
        End If
        
        CopiarDatosResSec = False
        Exit Function
    End If

    ' Buscar la primera fila realmente vaca (desde fila 2)
    filaDestino = 2
    Do While Application.WorksheetFunction.CountA(wsDestino.Rows(filaDestino)) > 0
        filaDestino = filaDestino + 1
    Loop

    ' Copiar los datos
    With wsDestino
        .Cells(filaDestino, "A").Value = edificio
        .Cells(filaDestino, "B").Value = nombre
        .Cells(filaDestino, "C").Value = configuracion
        .Cells(filaDestino, "D").Value = nri
        .Cells(filaDestino, "E").Value = superficie
        .Cells(filaDestino, "F").Value = resistencia
    End With

    If mostrarMensajes Then
        MsgBox "Datos copiados correctamente en la fila " & filaDestino & " de la hoja Anx II. Res Sec", vbInformation
    End If

    CopiarDatosResSec = True

End Function

Public Function CopiarDatosResExt(Optional mostrarMensajes As Boolean = True) As Boolean

    Dim wsInterfaz As Worksheet
    Dim wsDestino As Worksheet
    Dim filaDestino As Long
    
    Dim nombre As Variant, configuracion As Variant, nri As Variant
    Dim superficie As Variant, resistencia As Variant, edificio As Variant
    

    ' Referencias a hojas
    Set wsInterfaz = ThisWorkbook.Sheets("Interfaz")
    Set wsDestino = ThisWorkbook.Sheets("Anx. II Res Ext")

    ' Obtener los valores de Interfaz
    edificio = wsInterfaz.Range(celdaEdificio).Value
    nombre = wsInterfaz.Range(celdaNombre).Value
    configuracion = wsInterfaz.Range(celdaConfiguracion).Value
    nri = wsInterfaz.Range(celdaNri).Value
    superficie = wsInterfaz.Range(celdaSuperficie).Value
    resistencia = wsInterfaz.Range(celdaResistencia).Value

    ' Validar que todos los valores estn rellenos
    If IsEmpty(nombre) Or IsEmpty(configuracion) Or IsEmpty(nri) _
       Or IsEmpty(superficie) Or IsEmpty(resistencia) Then
        
        If mostrarMensajes Then
            MsgBox "Por favor, rellena todos los campos obligatorios en la hoja Interfaz antes de exportar." & vbCrLf & _
                   "Campos requeridos: Nombre, Configuracion, NRI, Superficie, Resistencia de elementos separadores de establecimientos", vbExclamation
        End If
        
        CopiarDatosResExt = False
        Exit Function
    End If

    ' Comprobar si alguno de los campos contiene la palabra "Error"
    If UCase(nombre) = "ERROR" Or UCase(configuracion) = "ERROR" Or UCase(nri) = "ERROR" _
       Or UCase(superficie) = "ERROR" Or UCase(resistencia) = "ERROR" Then
        
        If mostrarMensajes Then
            MsgBox "No se puede exportar porque los datos contienen errores", vbExclamation
        End If
        
        CopiarDatosResExt = False
        Exit Function
    End If

    ' Buscar la primera fila realmente vaca (desde fila 2)
    filaDestino = 2
    Do While Application.WorksheetFunction.CountA(wsDestino.Rows(filaDestino)) > 0
        filaDestino = filaDestino + 1
    Loop

    ' Copiar los datos
    With wsDestino
        .Cells(filaDestino, "A").Value = edificio
        .Cells(filaDestino, "B").Value = nombre
        .Cells(filaDestino, "C").Value = configuracion
        .Cells(filaDestino, "D").Value = nri
        .Cells(filaDestino, "E").Value = superficie
        .Cells(filaDestino, "F").Value = resistencia
    End With

    If mostrarMensajes Then
        MsgBox "Datos copiados correctamente en la fila " & filaDestino & " de la hoja Anx II. Res Ext", vbInformation
    End If

    CopiarDatosResExt = True
End Function

Public Function CopiarDatosOcupacion(Optional mostrarMensajes As Boolean = True) As Boolean

    Dim wsInterfaz As Worksheet
    Dim wsDestino As Worksheet
    Dim filaDestino As Long

    Dim nombre As Variant, personas As Variant, edificio As Variant
    Dim superficie As Variant, ocupacion As Variant


    ' Referencias a hojas
    Set wsInterfaz = ThisWorkbook.Sheets("Interfaz")
    Set wsDestino = ThisWorkbook.Sheets("Anx. II Ocu")

    ' Obtener los valores de Interfaz
    edificio = wsInterfaz.Range(celdaEdificio).Value
    nombre = wsInterfaz.Range(celdaNombre).Value
    personas = wsInterfaz.Range(celdaPersonas).Value
    superficie = wsInterfaz.Range(celdaSuperficie).Value
    ocupacion = wsInterfaz.Range(celdaOcupacion).Value

    ' Validar que todos los valores estn rellenos
    If IsEmpty(nombre) Or IsEmpty(personas) _
       Or IsEmpty(superficie) Or IsEmpty(ocupacion) Then

        If mostrarMensajes Then
            MsgBox "Por favor, rellena todos los campos obligatorios en la hoja Interfaz antes de exportar." & vbCrLf & _
                   "Campos requeridos: Nombre, Superficie, Personas, Ocupacion", vbExclamation
        End If

        CopiarDatosOcupacion = False
        Exit Function
    End If

    ' Comprobar si alguno de los campos contiene la palabra "Error"
    If UCase(nombre) = "ERROR" Or UCase(personas) = "ERROR" _
       Or UCase(superficie) = "ERROR" Or UCase(ocupacion) = "ERROR" Then

        If mostrarMensajes Then
            MsgBox "No se puede exportar porque los datos contienen errores", vbExclamation
        End If

        CopiarDatosOcupacion = False
        Exit Function
    End If

    ' Buscar la primera fila realmente vaca (desde fila 2)
    filaDestino = 2
    Do While Application.WorksheetFunction.CountA(wsDestino.Rows(filaDestino)) > 0
        filaDestino = filaDestino + 1
    Loop

    ' Copiar los datos
    With wsDestino
        .Cells(filaDestino, "A").Value = edificio
        .Cells(filaDestino, "B").Value = nombre
        .Cells(filaDestino, "C").Value = superficie
        .Cells(filaDestino, "D").Value = personas
        .Cells(filaDestino, "E").Value = ocupacion
    End With

    If mostrarMensajes Then
        MsgBox "Datos copiados correctamente en la fila " & filaDestino & " de la hoja Anx II. Ocu", vbInformation
    End If

    CopiarDatosOcupacion = True

End Function

Public Function CopiarDatosSalidas(Optional mostrarMensajes As Boolean = True) As Boolean

    Dim wsInterfaz As Worksheet
    Dim wsDestino As Worksheet
    Dim filaDestino As Long

    Dim nombre As Variant, salidas As Variant, nri As Variant, longitud As Variant
    Dim superficie As Variant, distancia1 As Variant, distancia2 As Variant, edificio As Variant


    ' Referencias a hojas
    Set wsInterfaz = ThisWorkbook.Sheets("Interfaz")
    Set wsDestino = ThisWorkbook.Sheets("Anx. II Sal")

    ' Obtener los valores de Interfaz
    edificio = wsInterfaz.Range(celdaEdificio).Value
    nombre = wsInterfaz.Range(celdaNombre).Value
    nri = wsInterfaz.Range(celdaNri).Value
    superficie = wsInterfaz.Range(celdaSuperficie).Value
    salidas = wsInterfaz.Range(celdaSalidas).Value
    distancia1 = wsInterfaz.Range(celdaDistancia1).Value
    distancia2 = wsInterfaz.Range(celdaDistancia2).Value

    ' Validar que todos los valores obligatorios estn rellenos
    If IsEmpty(nombre) Or IsEmpty(nri) _
       Or IsEmpty(superficie) Or IsEmpty(salidas) Then

        If mostrarMensajes Then
            MsgBox "Por favor, rellena todos los campos obligatorios en la hoja Interfaz antes de exportar." & vbCrLf & _
                   "Campos requeridos: Nombre, Superficie, Nivel de riesgo, Numero mnimo de salidas", vbExclamation
        End If

        CopiarDatosSalidas = False
        Exit Function
    End If

    ' Comprobar si alguno de los campos contiene la palabra "Error"
    If UCase(nombre) = "ERROR" Or UCase(nri) = "ERROR" Or UCase(superficie) = "ERROR" _
       Or UCase(salidas) = "ERROR" Or UCase(distancia1) = "ERROR" Or UCase(distancia2) = "ERROR" Then
        
        If mostrarMensajes Then
            MsgBox "No se puede exportar porque los datos contienen errores", vbExclamation
        End If
        
        CopiarDatosSalidas = False
        Exit Function
    End If

    ' Asignar longitud segn distancia
    If distancia1 = 0 Then
        longitud = distancia2
    Else
        longitud = distancia1
    End If

    ' Buscar la primera fila realmente vaca (desde fila 2)
    filaDestino = 2
    Do While Application.WorksheetFunction.CountA(wsDestino.Rows(filaDestino)) > 0
        filaDestino = filaDestino + 1
    Loop

    ' Copiar los datos
    With wsDestino
        .Cells(filaDestino, "A").Value = edificio
        .Cells(filaDestino, "B").Value = nombre
        .Cells(filaDestino, "C").Value = nri
        .Cells(filaDestino, "D").Value = superficie
        .Cells(filaDestino, "E").Value = salidas
        .Cells(filaDestino, "F").Value = longitud
    End With

    If mostrarMensajes Then
        MsgBox "Datos copiados correctamente en la fila " & filaDestino & " de la hoja Anx II. Sal", vbInformation
    End If

    CopiarDatosSalidas = True

End Function

Public Function CopiarDatosRes(Optional mostrarMensajes As Boolean = True) As Boolean

    Dim wsInterfaz As Worksheet
    Dim wsDestino As Worksheet
    Dim filaDestino As Long

    Dim nombre As Variant, configuracion As Variant, nri As Variant
    Dim superficie As Variant, resistencia As Variant, edificio As Variant

    ' Referencias a hojas
    Set wsInterfaz = ThisWorkbook.Sheets("Interfaz")
    Set wsDestino = ThisWorkbook.Sheets("Anx. II Res")

    ' Obtener los valores de Interfaz
    edificio = wsInterfaz.Range(celdaEdificio).Value
    nombre = wsInterfaz.Range(celdaNombre).Value
    configuracion = wsInterfaz.Range(celdaConfiguracion).Value
    nri = wsInterfaz.Range(celdaNri).Value
    superficie = wsInterfaz.Range(celdaSuperficie).Value
    resistencia = wsInterfaz.Range(celdaResultadoResistenciaEstructural).Value

    ' Validar que todos los valores estn rellenos
    If IsEmpty(nombre) Or IsEmpty(configuracion) Or IsEmpty(nri) _
       Or IsEmpty(superficie) Or IsEmpty(resistencia) Then

        If mostrarMensajes Then
            MsgBox "Por favor, rellena todos los campos obligatorios en la hoja Interfaz antes de exportar." & vbCrLf & _
                   "Campos requeridos: Nombre, Configuracion, NRI, Superficie, Resistencia de elementos estructurales", vbExclamation
        End If

        CopiarDatosRes = False
        Exit Function
    End If

    ' Comprobar si alguno de los campos contiene la palabra "Error"
    If UCase(nombre) = "ERROR" Or UCase(configuracion) = "ERROR" Or UCase(nri) = "ERROR" _
       Or UCase(superficie) = "ERROR" Or UCase(resistencia) = "ERROR" Then

        If mostrarMensajes Then
            MsgBox "No se puede exportar porque los datos contienen errores", vbExclamation
        End If

        CopiarDatosRes = False
        Exit Function
    End If

    ' Buscar la primera fila realmente vaca (desde fila 2)
    filaDestino = 2
    Do While Application.WorksheetFunction.CountA(wsDestino.Rows(filaDestino)) > 0
        filaDestino = filaDestino + 1
    Loop

    ' Copiar los datos
    With wsDestino
        .Cells(filaDestino, "A").Value = edificio
        .Cells(filaDestino, "B").Value = nombre
        .Cells(filaDestino, "C").Value = configuracion
        .Cells(filaDestino, "D").Value = nri
        .Cells(filaDestino, "E").Value = superficie
        .Cells(filaDestino, "F").Value = resistencia
    End With

    If mostrarMensajes Then
        MsgBox "Datos copiados correctamente en la fila " & filaDestino & " de la hoja Anx II. Res", vbInformation
    End If

    CopiarDatosRes = True

End Function

Public Function CopiarDatosDetAuto(Optional mostrarMensajes As Boolean = True) As Boolean

    Dim wsInterfaz As Worksheet
    Dim wsDestino As Worksheet
    Dim filaDestino As Long

    Dim nombre As Variant, nri As Variant, exige As Variant
    Dim superficie As Variant, actividad As Variant, edificio As Variant

    Const celdaDeteccionAuto As String = "A51"

    ' Referencias a hojas
    Set wsInterfaz = ThisWorkbook.Sheets("Interfaz")
    Set wsDestino = ThisWorkbook.Sheets("Anx. III Det Auto")

    ' Obtener los valores de Interfaz
    edificio = wsInterfaz.Range(celdaEdificio).Value
    nombre = wsInterfaz.Range(celdaNombre).Value
    nri = wsInterfaz.Range(celdaNri).Value
    superficie = wsInterfaz.Range(celdaSuperficie).Value
    exige = wsInterfaz.Range(celdaDeteccionAuto).Value
    actividad = wsInterfaz.Range(celdaActividad).Value

    ' Validar que todos los valores estn rellenos
    If IsEmpty(nombre) Or IsEmpty(nri) _
       Or IsEmpty(superficie) Or IsEmpty(exige) Or IsEmpty(actividad) Then

        If mostrarMensajes Then
            MsgBox "Por favor, rellena todos los campos obligatorios en la hoja Interfaz antes de exportar." & vbCrLf & _
                   "Campos requeridos: Nombre, Superficie, Nivel de riesgo, Actividad, Deteccion automatica", vbExclamation
        End If

        CopiarDatosDetAuto = False
        Exit Function
    End If

    ' Comprobar si alguno de los campos contiene la palabra "Error"
    If UCase(nombre) = "ERROR" Or UCase(nri) = "ERROR" Or UCase(superficie) = "ERROR" _
       Or UCase(exige) = "ERROR" Or UCase(actividad) = "ERROR" Then

        If mostrarMensajes Then
            MsgBox "No se puede exportar porque los datos contienen errores", vbExclamation
        End If

        CopiarDatosDetAuto = False
        Exit Function
    End If

    ' Buscar la primera fila realmente vaca (desde fila 2)
    filaDestino = 2
    Do While Application.WorksheetFunction.CountA(wsDestino.Rows(filaDestino)) > 0
        filaDestino = filaDestino + 1
    Loop

    ' Copiar los datos
    With wsDestino
        .Cells(filaDestino, "A").Value = edificio
        .Cells(filaDestino, "B").Value = nombre
        .Cells(filaDestino, "C").Value = superficie
        .Cells(filaDestino, "D").Value = nri
        .Cells(filaDestino, "E").Value = actividad
        .Cells(filaDestino, "F").Value = exige
    End With

    If mostrarMensajes Then
        MsgBox "Datos copiados correctamente en la fila " & filaDestino & " de la hoja Anx III. Det Auto", vbInformation
    End If

    CopiarDatosDetAuto = True

End Function

Public Function CopiarDatosMan(Optional mostrarMensajes As Boolean = True) As Boolean

    Dim wsInterfaz As Worksheet
    Dim wsDestino As Worksheet
    Dim filaDestino As Long

    Dim nombre As Variant, nri As Variant, exige As Variant
    Dim superficie As Variant, actividad As Variant, edificio As Variant

    ' Referencias a hojas
    Set wsInterfaz = ThisWorkbook.Sheets("Interfaz")
    Set wsDestino = ThisWorkbook.Sheets("Anx. III Det Man")

    ' Obtener los valores de Interfaz
    edificio = wsInterfaz.Range(celdaEdificio).Value
    nombre = wsInterfaz.Range(celdaNombre).Value
    nri = wsInterfaz.Range(celdaNri).Value
    superficie = wsInterfaz.Range(celdaSuperficie).Value
    exige = wsInterfaz.Range(celdaDeteccionManual).Value
    actividad = wsInterfaz.Range(celdaActividad).Value

    ' Validar que todos los valores estn rellenos
    If IsEmpty(nombre) Or IsEmpty(nri) _
       Or IsEmpty(superficie) Or IsEmpty(exige) Or IsEmpty(actividad) Then

        If mostrarMensajes Then
            MsgBox "Por favor, rellena todos los campos obligatorios en la hoja Interfaz antes de exportar." & vbCrLf & _
                   "Campos requeridos: Nombre, Superficie, Nivel de riesgo, Actividad, Deteccion manual", vbExclamation
        End If

        CopiarDatosMan = False
        Exit Function
    End If

    ' Comprobar si alguno de los campos contiene la palabra "Error"
    If UCase(nombre) = "ERROR" Or UCase(nri) = "ERROR" Or UCase(superficie) = "ERROR" _
       Or UCase(exige) = "ERROR" Or UCase(actividad) = "ERROR" Then

        If mostrarMensajes Then
            MsgBox "No se puede exportar porque los datos contienen errores", vbExclamation
        End If

        CopiarDatosMan = False
        Exit Function
    End If

    ' Buscar la primera fila realmente vaca (desde fila 2)
    filaDestino = 2
    Do While Application.WorksheetFunction.CountA(wsDestino.Rows(filaDestino)) > 0
        filaDestino = filaDestino + 1
    Loop

    ' Copiar los datos
    With wsDestino
        .Cells(filaDestino, "A").Value = edificio
        .Cells(filaDestino, "B").Value = nombre
        .Cells(filaDestino, "C").Value = superficie
        .Cells(filaDestino, "D").Value = nri
        .Cells(filaDestino, "E").Value = actividad
        .Cells(filaDestino, "F").Value = exige
    End With

    If mostrarMensajes Then
        MsgBox "Datos copiados correctamente en la fila " & filaDestino & " de la hoja Anx III. Man", vbInformation
    End If

    CopiarDatosMan = True

End Function

Public Function CopiarDatosAlr(Optional mostrarMensajes As Boolean = True) As Boolean

    Dim wsInterfaz As Worksheet
    Dim wsDestino As Worksheet
    Dim filaDestino As Long

    Dim nombre As Variant, nri As Variant, exigeAuto As Variant, exigeMan As Variant
    Dim superficie As Variant, actividad As Variant, exigeAlr As Variant, edificio As Variant


    ' Referencias a hojas
    Set wsInterfaz = ThisWorkbook.Sheets("Interfaz")
    Set wsDestino = ThisWorkbook.Sheets("Anx. III Alr")

    ' Obtener los valores de Interfaz
    edificio = wsInterfaz.Range(celdaEdificio).Value
    nombre = wsInterfaz.Range(celdaNombre).Value
    nri = wsInterfaz.Range(celdaNri).Value
    superficie = wsInterfaz.Range(celdaSuperficie).Value
    exigeAuto = wsInterfaz.Range(celdaExigeAuto).Value
    exigeMan = wsInterfaz.Range(celdaExigeMan).Value
    actividad = wsInterfaz.Range(celdaActividad).Value

    ' Validar que todos los valores estn rellenos
    If IsEmpty(nombre) Or IsEmpty(nri) _
       Or IsEmpty(superficie) Or IsEmpty(exigeAuto) Or IsEmpty(exigeMan) Or IsEmpty(actividad) Then

        If mostrarMensajes Then
            MsgBox "Por favor, rellena todos los campos obligatorios en la hoja Interfaz antes de exportar." & vbCrLf & _
                   "Campos requeridos: Nombre, Superficie, Nivel de riesgo, Actividad, Detecciin automtica, Deteccion Manual", vbExclamation
        End If

        CopiarDatosAlr = False
        Exit Function
    End If

    ' Validar que ningn valor sea "Error"
    If nombre = "Error" Or nri = "Error" Or superficie = "Error" _
       Or exigeAuto = "Error" Or exigeMan = "Error" Or actividad = "Error" Then

        If mostrarMensajes Then
            MsgBox "No se puede exportar porque los datos contienen errores", vbExclamation
        End If

        CopiarDatosAlr = False
        Exit Function
    End If

    ' Evaluar exigeAlr segn condiciones
    If exigeAuto = "Si" Or exigeMan = "Si" Then
        exigeAlr = "Si"
    Else
        exigeAlr = "No"
    End If

    ' Buscar la primera fila realmente vaca (desde fila 2)
    filaDestino = 2
    Do While Application.WorksheetFunction.CountA(wsDestino.Rows(filaDestino)) > 0
        filaDestino = filaDestino + 1
    Loop

    ' Copiar los datos
    With wsDestino
        .Cells(filaDestino, "A").Value = edificio
        .Cells(filaDestino, "B").Value = nombre
        .Cells(filaDestino, "C").Value = superficie
        .Cells(filaDestino, "D").Value = nri
        .Cells(filaDestino, "E").Value = actividad
        .Cells(filaDestino, "F").Value = exigeAlr
    End With

    If mostrarMensajes Then
        MsgBox "Datos copiados correctamente en la fila " & filaDestino & " de la hoja Anx III. Alr", vbInformation
    End If

    CopiarDatosAlr = True

End Function

Public Function CopiarDatosMeg(Optional mostrarMensajes As Boolean = True) As Boolean

    Dim wsInterfaz As Worksheet
    Dim wsDestino As Worksheet
    Dim filaDestino As Long

    Dim nombre As Variant, nri As Variant, exige As Variant
    Dim superficie As Variant, actividad As Variant, edificio As Variant

    ' Referencias a hojas
    Set wsInterfaz = ThisWorkbook.Sheets("Interfaz")
    Set wsDestino = ThisWorkbook.Sheets("Anx. III Meg")

    ' Obtener los valores de Interfaz
    edificio = wsInterfaz.Range(celdaEdificio).Value
    nombre = wsInterfaz.Range(celdaNombre).Value
    nri = wsInterfaz.Range(celdaNri).Value
    superficie = wsInterfaz.Range(celdaSuperficie).Value
    exige = wsInterfaz.Range(celdaMegafonia).Value
    actividad = wsInterfaz.Range(celdaActividad).Value

    ' Validar que todos los valores estn rellenos
    If IsEmpty(nombre) Or IsEmpty(nri) _
       Or IsEmpty(superficie) Or IsEmpty(exige) Or IsEmpty(actividad) Then

        If mostrarMensajes Then
            MsgBox "Por favor, rellena todos los campos obligatorios en la hoja Interfaz antes de exportar." & vbCrLf & _
                   "Campos requeridos: Nombre, Superficie, Nivel de riesgo, Actividad, Megafona alarma local y alarma general", vbExclamation
        End If

        CopiarDatosMeg = False
        Exit Function
    End If

    ' Validar que ningn valor sea "Error"
    If nombre = "Error" Or nri = "Error" Or superficie = "Error" _
       Or exige = "Error" Or actividad = "Error" Then

        If mostrarMensajes Then
            MsgBox "No se puede exportar porque los datos contienen errores", vbExclamation
        End If

        CopiarDatosMeg = False
        Exit Function
    End If

    ' Buscar la primera fila realmente vaca (desde fila 2)
    filaDestino = 2
    Do While Application.WorksheetFunction.CountA(wsDestino.Rows(filaDestino)) > 0
        filaDestino = filaDestino + 1
    Loop

    ' Copiar los datos
    With wsDestino
        .Cells(filaDestino, "A").Value = edificio
        .Cells(filaDestino, "B").Value = nombre
        .Cells(filaDestino, "C").Value = superficie
        .Cells(filaDestino, "D").Value = nri
        .Cells(filaDestino, "E").Value = actividad
        .Cells(filaDestino, "F").Value = exige
    End With

    If mostrarMensajes Then
        MsgBox "Datos copiados correctamente en la fila " & filaDestino & " de la hoja Anx III. Meg", vbInformation
    End If

    CopiarDatosMeg = True

End Function

Public Function CopiarDatosHidrantes(Optional mostrarMensajes As Boolean = True) As Boolean

    Dim wsInterfaz As Worksheet
    Dim wsDestino As Worksheet
    Dim filaDestino As Long

    Dim nombre As Variant, nri As Variant, exigeCam As Variant, exigeDir As Variant
    Dim superficie As Variant, actividad As Variant, edificio As Variant


    ' Referencias a hojas
    Set wsInterfaz = ThisWorkbook.Sheets("Interfaz")
    Set wsDestino = ThisWorkbook.Sheets("Anx. III Hid")

    ' Obtener los valores de Interfaz
    edificio = wsInterfaz.Range(celdaEdificio).Value
    nombre = wsInterfaz.Range(celdaNombre).Value
    nri = wsInterfaz.Range(celdaNri).Value
    superficie = wsInterfaz.Range(celdaSuperficie).Value
    exigeCam = wsInterfaz.Range(celdaExigeCam).Value
    exigeDir = wsInterfaz.Range(celdaExigeDir).Value
    actividad = wsInterfaz.Range(celdaActividad).Value

    ' Validar que todos los valores estn rellenos
    If IsEmpty(nombre) Or IsEmpty(nri) _
       Or IsEmpty(superficie) Or IsEmpty(exigeCam) Or IsEmpty(exigeDir) Or IsEmpty(actividad) Then

        If mostrarMensajes Then
            MsgBox "Por favor, rellena todos los campos obligatorios en la hoja Interfaz antes de exportar." & vbCrLf & _
                   "Campos requeridos: Nombre, Superficie, Nivel de riesgo, Actividad, Hidrantes de llenado de camiones, Hidrantes impulsion directa", vbExclamation
        End If

        CopiarDatosHidrantes = False
        Exit Function
    End If

    ' Validar que no haya valores con texto "Error"
    If LCase(nombre) = "error" Or LCase(nri) = "error" _
       Or LCase(superficie) = "error" Or LCase(exigeCam) = "error" _
       Or LCase(exigeDir) = "error" Or LCase(actividad) = "error" Then

        If mostrarMensajes Then
            MsgBox "No se puede exportar porque los datos contienen errores", vbExclamation
        End If

        CopiarDatosHidrantes = False
        Exit Function
    End If

    ' Buscar la primera fila realmente vaca (desde fila 2)
    filaDestino = 2
    Do While Application.WorksheetFunction.CountA(wsDestino.Rows(filaDestino)) > 0
        filaDestino = filaDestino + 1
    Loop

    ' Copiar los datos
    With wsDestino
        .Cells(filaDestino, "A").Value = edificio
        .Cells(filaDestino, "B").Value = nombre
        .Cells(filaDestino, "C").Value = superficie
        .Cells(filaDestino, "D").Value = nri
        .Cells(filaDestino, "E").Value = actividad
        .Cells(filaDestino, "F").Value = exigeCam
        .Cells(filaDestino, "G").Value = exigeDir
    End With

    If mostrarMensajes Then
        MsgBox "Datos copiados correctamente en la fila " & filaDestino & " de la hoja Anx III. Hid", vbInformation
    End If

    CopiarDatosHidrantes = True

End Function

Public Function CopiarDatosExt(Optional mostrarMensajes As Boolean = True) As Boolean

    Dim wsInterfaz As Worksheet
    Dim wsDestino As Worksheet
    Dim filaDestino As Long

    Dim nombre As Variant, nri As Variant, exige As Variant, edificio As Variant
    Dim superficie As Variant, actividad As Variant, numero As Variant


    ' Referencias a hojas
    Set wsInterfaz = ThisWorkbook.Sheets("Interfaz")
    Set wsDestino = ThisWorkbook.Sheets("Anx. III Ext Ext")

    ' Obtener los valores de Interfaz
    edificio = wsInterfaz.Range(celdaEdificio).Value
    nombre = wsInterfaz.Range(celdaNombre).Value
    nri = wsInterfaz.Range(celdaNri).Value
    superficie = wsInterfaz.Range(celdaSuperficie).Value
    numero = wsInterfaz.Range(celdaNumero).Value
    actividad = wsInterfaz.Range(celdaActividad).Value

    ' Validar que todos los valores estn rellenos
    If IsEmpty(nombre) Or IsEmpty(nri) _
       Or IsEmpty(superficie) Or IsEmpty(numero) Or IsEmpty(actividad) Then

        If mostrarMensajes Then
            MsgBox "Por favor, rellena todos los campos obligatorios en la hoja Interfaz antes de exportar." & vbCrLf & _
                   "Campos requeridos: Nombre, Superficie, Nivel de riesgo, Actividad, Numero de extintores", vbExclamation
        End If

        CopiarDatosExt = False
        Exit Function
    End If

    ' Comprobar si algn valor contiene el texto "Error"
    If LCase(nombre) = "error" Or LCase(nri) = "error" _
       Or LCase(superficie) = "error" Or LCase(numero) = "error" _
       Or LCase(actividad) = "error" Then

        If mostrarMensajes Then
            MsgBox "No se puede exportar porque los datos contienen errores", vbCritical
        End If

        CopiarDatosExt = False
        Exit Function
    End If

    ' Evaluar exige segn nmero
    If numero = 0 Then
        exige = "No"
    Else
        exige = "Si"
    End If

    ' Buscar la primera fila realmente vaca (desde fila 2)
    filaDestino = 2
    Do While Application.WorksheetFunction.CountA(wsDestino.Rows(filaDestino)) > 0
        filaDestino = filaDestino + 1
    Loop

    ' Copiar los datos
    With wsDestino
        .Cells(filaDestino, "A").Value = edificio
        .Cells(filaDestino, "B").Value = nombre
        .Cells(filaDestino, "C").Value = superficie
        .Cells(filaDestino, "D").Value = nri
        .Cells(filaDestino, "E").Value = actividad
        .Cells(filaDestino, "F").Value = exige
        .Cells(filaDestino, "G").Value = numero
    End With

    If mostrarMensajes Then
        MsgBox "Datos copiados correctamente en la fila " & filaDestino & " de la hoja Anx III. Ext Ext", vbInformation
    End If

    CopiarDatosExt = True

End Function

Public Function CopiarDatosBie(Optional mostrarMensajes As Boolean = True) As Boolean

    Dim wsInterfaz As Worksheet
    Dim wsDestino As Worksheet
    Dim filaDestino As Long

    Dim nombre As Variant, nri As Variant, exige As Variant, edificio As Variant
    Dim superficie As Variant, actividad As Variant, tipo As Variant


    ' Referencias a hojas
    Set wsInterfaz = ThisWorkbook.Sheets("Interfaz")
    Set wsDestino = ThisWorkbook.Sheets("Anx. III BIES")

    ' Obtener los valores de Interfaz
    edificio = wsInterfaz.Range(celdaEdificio).Value
    nombre = wsInterfaz.Range(celdaNombre).Value
    nri = wsInterfaz.Range(celdaNri).Value
    superficie = wsInterfaz.Range(celdaSuperficie).Value
    exige = wsInterfaz.Range(celdaNecesitaBIES).Value
    tipo = wsInterfaz.Range(celdaTipo).Value
    actividad = wsInterfaz.Range(celdaActividad).Value

    ' Validar que todos los valores estn rellenos
    If IsEmpty(nombre) Or IsEmpty(nri) Or IsEmpty(exige) _
       Or IsEmpty(superficie) Or IsEmpty(tipo) Or IsEmpty(actividad) Then

        If mostrarMensajes Then
            MsgBox "Por favor, rellena todos los campos obligatorios en la hoja Interfaz antes de exportar." & vbCrLf & _
                   "Campos requeridos: Nombre, Superficie, Nivel de riesgo, Actividad, BIES, Tipo de BIE", vbExclamation
        End If

        CopiarDatosBie = False
        Exit Function
    End If

    ' Validar que no haya errores en los datos
    If LCase(nombre) = "error" Or LCase(nri) = "error" Or LCase(superficie) = "error" _
       Or LCase(actividad) = "error" Or LCase(exige) = "error" Or LCase(tipo) = "error" Then

        If mostrarMensajes Then
            MsgBox "No se puede exportar porque los datos contienen errores", vbCritical
        End If

        CopiarDatosBie = False
        Exit Function
    End If

    ' Buscar la primera fila realmente vaca (desde fila 2)
    filaDestino = 2
    Do While Application.WorksheetFunction.CountA(wsDestino.Rows(filaDestino)) > 0
        filaDestino = filaDestino + 1
    Loop

    ' Copiar los datos
    With wsDestino
        .Cells(filaDestino, "A").Value = edificio
        .Cells(filaDestino, "B").Value = nombre
        .Cells(filaDestino, "C").Value = superficie
        .Cells(filaDestino, "D").Value = nri
        .Cells(filaDestino, "E").Value = actividad
        .Cells(filaDestino, "F").Value = exige
        .Cells(filaDestino, "G").Value = tipo
    End With

    If mostrarMensajes Then
        MsgBox "Datos copiados correctamente en la fila " & filaDestino & " de la hoja Anx III. BIES", vbInformation
    End If

    CopiarDatosBie = True

End Function

Public Function CopiarDatosHumos(Optional mostrarMensajes As Boolean = True) As Boolean

    Dim wsInterfaz As Worksheet
    Dim wsDestino As Worksheet
    Dim filaDestino As Long

    Dim nombre As Variant, nri As Variant, exige As Variant, edificio As Variant
    Dim superficie As Variant


    ' Referencias a hojas
    Set wsInterfaz = ThisWorkbook.Sheets("Interfaz")
    Set wsDestino = ThisWorkbook.Sheets("Anx. III Hum")

    ' Obtener los valores de Interfaz
    edificio = wsInterfaz.Range(celdaEdificio).Value
    nombre = wsInterfaz.Range(celdaNombre).Value
    nri = wsInterfaz.Range(celdaNri).Value
    superficie = wsInterfaz.Range(celdaSuperficie).Value
    exige = wsInterfaz.Range(celdaNecesitaHumos).Value

    ' Validar que todos los valores estn rellenos
    If IsEmpty(nombre) Or IsEmpty(nri) Or IsEmpty(exige) _
       Or IsEmpty(superficie) Then

        If mostrarMensajes Then
            MsgBox "Por favor, rellena todos los campos obligatorios en la hoja Interfaz antes de exportar." & vbCrLf & _
                   "Campos requeridos: Nombre, Superficie, Nivel de riesgo, Extincion Automatica", vbExclamation
        End If

        CopiarDatosHumos = False
        Exit Function
    End If

    ' Validar que no haya errores en los campos (insensible a mayusculas)
    If LCase(nombre) = "error" Or LCase(nri) = "error" _
       Or LCase(exige) = "error" Or LCase(superficie) = "error" Then

        If mostrarMensajes Then
            MsgBox "No se puede exportar porque los datos contienen errores.", vbCritical
        End If
        CopiarDatosHumos = False
        Exit Function
    End If

    ' Buscar la primera fila realmente vaca (desde fila 2)
    filaDestino = 2
    Do While Application.WorksheetFunction.CountA(wsDestino.Rows(filaDestino)) > 0
        filaDestino = filaDestino + 1
    Loop

    ' Copiar los datos
    With wsDestino
        .Cells(filaDestino, "A").Value = edificio
        .Cells(filaDestino, "B").Value = nombre
        .Cells(filaDestino, "C").Value = superficie
        .Cells(filaDestino, "D").Value = nri
        .Cells(filaDestino, "E").Value = exige
    End With

    If mostrarMensajes Then
        MsgBox "Datos copiados correctamente en la fila " & filaDestino & " de la hoja Anx III. Hum", vbInformation
    End If

    CopiarDatosHumos = True

End Function

Public Function CopiarDatosAuto(Optional mostrarMensajes As Boolean = True) As Boolean

    Dim wsInterfaz As Worksheet
    Dim wsDestino As Worksheet
    Dim filaDestino As Long

    Dim nombre As Variant, nri As Variant, exige As Variant, edificio As Variant
    Dim superficie As Variant, actividad As Variant


    ' Referencias a hojas
    Set wsInterfaz = ThisWorkbook.Sheets("Interfaz")
    Set wsDestino = ThisWorkbook.Sheets("Anx. III Ext Auto")

    ' Obtener los valores de Interfaz
    edificio = wsInterfaz.Range(celdaEdificio).Value
    nombre = wsInterfaz.Range(celdaNombre).Value
    nri = wsInterfaz.Range(celdaNri).Value
    superficie = wsInterfaz.Range(celdaSuperficie).Value
    exige = wsInterfaz.Range(celdaNecesitaExtAuto).Value
    actividad = wsInterfaz.Range(celdaActividad).Value

    ' Validar que todos los valores estn rellenos
    If IsEmpty(nombre) Or IsEmpty(nri) Or IsEmpty(exige) _
       Or IsEmpty(superficie) Or IsEmpty(actividad) Then

        If mostrarMensajes Then
            MsgBox "Por favor, rellena todos los campos obligatorios en la hoja Interfaz antes de exportar." & vbCrLf & _
                   "Campos requeridos: Nombre, Superficie, Nivel de riesgo, Actividad, Extincion Automatica", vbExclamation
        End If

        CopiarDatosAuto = False
        Exit Function
    End If

    ' Validar que no haya errores en los campos (insensible a mayusculas)
    If LCase(nombre) = "error" Or LCase(nri) = "error" _
       Or LCase(exige) = "error" Or LCase(superficie) = "error" _
       Or LCase(actividad) = "error" Then

        If mostrarMensajes Then
            MsgBox "No se puede exportar porque los datos contienen errores", vbCritical
        End If

        CopiarDatosAuto = False
        Exit Function
    End If

    ' Buscar la primera fila realmente vaca (desde fila 2)
    filaDestino = 2
    Do While Application.WorksheetFunction.CountA(wsDestino.Rows(filaDestino)) > 0
        filaDestino = filaDestino + 1
    Loop

    ' Copiar los datos
    With wsDestino
        .Cells(filaDestino, "A").Value = edificio
        .Cells(filaDestino, "B").Value = nombre
        .Cells(filaDestino, "C").Value = superficie
        .Cells(filaDestino, "D").Value = nri
        .Cells(filaDestino, "E").Value = actividad
        .Cells(filaDestino, "F").Value = exige
    End With

    If mostrarMensajes Then
        MsgBox "Datos copiados correctamente en la fila " & filaDestino & " de la hoja Anx III. Ext Auto", vbInformation
    End If

    CopiarDatosAuto = True

End Function

Public Function CopiarDatosCol(Optional mostrarMensajes As Boolean = True) As Boolean

    Dim wsInterfaz As Worksheet
    Dim wsDestino As Worksheet
    Dim filaDestino As Long

    Dim nombre As Variant, nri As Variant, exige As Variant, edificio As Variant
    Dim superficie As Variant, actividad As Variant

    ' Referencias a hojas
    Set wsInterfaz = ThisWorkbook.Sheets("Interfaz")
    Set wsDestino = ThisWorkbook.Sheets("Anx. III Col")

    ' Obtener los valores de Interfaz
    edificio = wsInterfaz.Range(celdaEdificio).Value
    nombre = wsInterfaz.Range(celdaNombre).Value
    nri = wsInterfaz.Range(celdaNri).Value
    superficie = wsInterfaz.Range(celdaSuperficie).Value
    exige = wsInterfaz.Range(celdaNecesitaCol).Value
    actividad = wsInterfaz.Range(celdaActividad).Value

    ' Validar que todos los valores estn rellenos
    If IsEmpty(nombre) Or IsEmpty(nri) Or IsEmpty(exige) _
       Or IsEmpty(superficie) Or IsEmpty(actividad) Then

        If mostrarMensajes Then
            MsgBox "Por favor, rellena todos los campos obligatorios en la hoja Interfaz antes de exportar." & vbCrLf & _
                   "Campos requeridos: Nombre, Superficie, Nivel de riesgo, Actividad, Columna Seca", vbExclamation
        End If

        CopiarDatosCol = False
        Exit Function
    End If

    ' Validar que no haya errores en los campos (insensible a mayusculas)
    If LCase(nombre) = "error" Or LCase(nri) = "error" Or LCase(exige) = "error" _
       Or LCase(superficie) = "error" Or LCase(actividad) = "error" Then

        If mostrarMensajes Then
            MsgBox "No se puede exportar porque los datos contienen errores", vbCritical
        End If

        CopiarDatosCol = False
        Exit Function
    End If

    ' Buscar la primera fila realmente vaca (desde fila 2)
    filaDestino = 2
    Do While Application.WorksheetFunction.CountA(wsDestino.Rows(filaDestino)) > 0
        filaDestino = filaDestino + 1
    Loop

    ' Copiar los datos
    With wsDestino
        .Cells(filaDestino, "A").Value = edificio
        .Cells(filaDestino, "B").Value = nombre
        .Cells(filaDestino, "C").Value = superficie
        .Cells(filaDestino, "D").Value = nri
        .Cells(filaDestino, "E").Value = actividad
        .Cells(filaDestino, "F").Value = exige
    End With

    If mostrarMensajes Then
        MsgBox "Datos copiados correctamente en la fila " & filaDestino & " de la hoja Anx III. Col", vbInformation
    End If

    CopiarDatosCol = True

End Function

Public Sub ExportarTablasAWord()

    Const RUTA_LOGO_IZQ As String = "\Recursos\logo_izq.png"
    Const RUTA_LOGO_DER As String = "\Recursos\logo_der.png"
    Const wdPageBreak As Long = 7
    Const wdCollapseEnd As Long = 0
    Const RUTA_TEXTOS_XLSX As String = "\Recursos\Textos.xlsx"
    Const NOMBRE_HOJA_TEXTOS As String = "Textos"

    Dim wbTextos As Workbook
    Dim wsTextosExt As Worksheet
    Dim wordApp As Object, wordDoc As Object
    Dim ws As Worksheet, wsTextos As Worksheet
    Dim ruta As String, nombreArchivo As String, nombreFinal As String
    Dim nExp As String, cliente As String, nombreProyecto As String
    Dim HojaNombres As Variant, hojaTitulos As Variant
    Dim idx As Long

    ' --- VALIDACIONES EN INTERFAZ ---
    With ThisWorkbook.Worksheets("Interfaz")
        If IsError(.Range(celdaExigeDir).Value) Or Len(Trim$(CStr(.Range(celdaExigeDir).Value))) = 0 Or Val(.Range(celdaExigeDir).Value) = 0 Then
            .Range(celdaExigeDir).Value = "No tiene"
        End If
        If CStr(.Range(ADDR_C56).Value) = "" Or .Range(ADDR_C56).Value = 0 _
           Or CStr(.Range(ADDR_C57).Value) = "" Or .Range(ADDR_C57).Value = 0 _
           Or CStr(.Range(ADDR_C58).Value) = "" Or .Range(ADDR_C58).Value = 0 _
           Or CStr(.Range(celdaExigeDir).Value) = "" Or .Range(celdaExigeDir).Value = 0 Then
            MsgBox "Faltan datos esenciales en la hoja Interfaz.", vbExclamation
            Exit Sub
        End If
        If CStr(.Range(ADDR_E56).Value) = "" Or .Range(ADDR_E56).Value = 0 _
           Or CStr(.Range(ADDR_E57).Value) = "" Or .Range(ADDR_E57).Value = 0 _
           Or CStr(.Range(ADDR_E58).Value) = "" Or .Range(ADDR_E58).Value = 0 Then
            MsgBox "Faltan datos esenciales de la definicion del nombre del proyecto", vbExclamation
            Exit Sub
        End If
    End With

    ' --- ARRANCAR WORD ---
    Call CerrarTodasInstanciasWord
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    If wordApp Is Nothing Then Set wordApp = CreateObject("Word.Application")
    On Error GoTo 0
    wordApp.Visible = True

    Set wordDoc = wordApp.Documents.add
    AsegurarEstilosBase wordDoc
    AsegurarEstilosBaseConSiguiente wordDoc
    ConstruirEncabezado wordDoc, ThisWorkbook.path & RUTA_LOGO_IZQ, ThisWorkbook.path & RUTA_LOGO_DER, TextoCentralEncabezado()

    ' --- TEXTOS EXTERNOS ---
    Dim wsTextosInt As Worksheet
    On Error Resume Next
    Set wsTextosInt = ThisWorkbook.Worksheets("Textos")
    On Error GoTo 0

    Set wsTextosExt = AbrirHojaTextosExterna(ThisWorkbook.path & RUTA_TEXTOS_XLSX, NOMBRE_HOJA_TEXTOS)
    If wsTextosExt Is Nothing Then
        MsgBox "No se pudo abrir '" & ThisWorkbook.path & RUTA_TEXTOS_XLSX & "'. " & vbCrLf & _
               "Se utilizara la hoja interna 'Textos' si existe.", vbInformation
    Else
        Set wbTextos = wsTextosExt.Parent
    End If

    ' --- NOMBRE DE SALIDA ---
    ruta = Environ$("USERPROFILE") & "\Desktop\"
    nExp = ThisWorkbook.Worksheets("Interfaz").Range(ADDR_E56).Value
    cliente = ThisWorkbook.Worksheets("Interfaz").Range(ADDR_E57).Value
    nombreProyecto = ThisWorkbook.Worksheets("Interfaz").Range(ADDR_E58).Value
    nombreArchivo = "PXL_" & nExp & "_" & cliente & "_" & nombreProyecto
    nombreFinal = nombreArchivo & ".docx"
    Dim i As Long: i = 1
    Do While Dir$(ruta & nombreFinal) <> ""
        nombreFinal = nombreArchivo & "_" & i & ".docx"
        i = i + 1
    Loop

    ' --- LISTA DE HOJAS/TiTULOS ---
    HojaNombres = Array("Anx. II Sup", "Anx. II Reac", "Anx. II Res Sec", "Anx. II Res Ext", _
                        "Anx. II Ocu", "Anx. II Sal", "Anx. II Res", _
                        "Anx. III Det Auto", "Anx. III Det Man", "Anx. III Alr", _
                        "Anx. III Meg", "Anx. III Hid", "Anx. III Ext Ext", "Anx. III BIES", _
                        "Anx. III Ext Auto", "Anx. III Col", "Anx. III Hum")

    hojaTitulos = Array("Superficie Admisible", _
                        "Reaccion al fuego de techos y paredes", _
                        "Resistencia al fuego de los elementos constructivos separadores de sectores", _
                        "Resistencia al fuego de los elementos separadores con otros establecimientos", _
                        "Ocupacion", _
                        "Salidas de Emergencia", _
                        "Resistencia al fuego de los elementos estructurales", _
                        "Deteccion Automatica", _
                        "Deteccion Manual", _
                        "Alarmas", _
                        "Megafonia, Alarma general y alarma local", _
                        "Hidrantes", _
                        "Extintores", _
                        "Bocas de incendio equipadas", _
                        "Extincion automatica", _
                        "Columna seca", _
                        "Extraccion de humos")

    Dim hayDatosSecundarios As Boolean: hayDatosSecundarios = False

    ' --- RECORRIDO DE HOJAS ---
    For idx = LBound(HojaNombres) To UBound(HojaNombres)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(HojaNombres(idx)))
        On Error GoTo 0
        If ws Is Nothing Then GoTo SaltarHoja

        ' Detectar ultima celda
        Dim cLast As Range, rLast As Long, cLastCol As Long
        On Error Resume Next
        Set cLast = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, LookAt:=xlPart, _
                                  SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
        If Not cLast Is Nothing Then rLast = cLast.row
        Set cLast = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, LookAt:=xlPart, _
                                  SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        If Not cLast Is Nothing Then cLastCol = cLast.Column
        On Error GoTo 0

        If rLast < 2 Or cLastCol < 1 Then GoTo SaltarHoja
        hayDatosSecundarios = True

        ' TiTULO
        With wordDoc.content
            .Collapse wdCollapseEnd
            .InsertParagraphAfter
            .Collapse wdCollapseEnd
            .InsertAfter CStr(hojaTitulos(idx))
            ApplyStyleSafe wordDoc, .Paragraphs(.Paragraphs.Count).Range, "EstiloTituloTabla", "Normal"
            .InsertParagraphAfter
            .Collapse wdCollapseEnd
        End With

        ' DESCRIPCIoN â€” ðŸ”§ PARCHE DE COLOR AZUL
        Dim textoDescripcion As String
        textoDescripcion = ObtenerDescripcionDesdeTextos(CStr(HojaNombres(idx)), wsTextosExt, wsTextosInt)

        If Len(textoDescripcion) > 0 Then
            With wordDoc.content
                .Collapse wdCollapseEnd
                .InsertAfter textoDescripcion

                ' Forzamos el formato a negro/estilo normal
                With .Paragraphs(.Paragraphs.Count).Range
                    .Style = wordDoc.Styles("Normal")
                    .Font.Reset
                    .Font.ColorIndex = wdAuto
                    .ParagraphFormat.Reset
                End With

                .InsertParagraphAfter
                .InsertParagraphAfter
                .Collapse wdCollapseEnd
            End With
        End If

        ' edificios
        Dim edificiosArr As Variant, e As Long
        edificiosArr = UniqueBuildingsSorted(ws, rLast)

        If IsEmpty(edificiosArr) Then
            PegarTablaDesdeBloqueContiguo ws, rLast, cLastCol, vbNullString, wordDoc, True
        Else
            For e = LBound(edificiosArr) To UBound(edificiosArr)
                Dim nombreBonito As String: nombreBonito = CStr(edificiosArr(e))
                With wordDoc.content
                    .Collapse wdCollapseEnd
                    .InsertParagraphAfter
                    .Collapse wdCollapseEnd
                    .InsertAfter nombreBonito
                    ApplyStyleSafe wordDoc, .Paragraphs(.Paragraphs.Count).Range, "EstiloSubtituloEdificio", "Normal"
                    .InsertParagraphAfter
                    .Collapse wdCollapseEnd
                End With
                PegarTablaDesdeBloqueContiguo ws, rLast, cLastCol, nombreBonito, wordDoc, True
                With wordDoc.content
                    .Collapse wdCollapseEnd
                    .InsertParagraphAfter
                    .Collapse wdCollapseEnd
                End With
            Next e
        End If

        ' Salto de pagina
        If idx < UBound(HojaNombres) Then
            With wordDoc.content
                .Collapse wdCollapseEnd
                .InsertBreak wdPageBreak
                .Collapse wdCollapseEnd
            End With
        End If

SaltarHoja:
    Next idx

    If Not hayDatosSecundarios Then
        MsgBox "No se han encontrado datos en las hojas secundarias.", vbExclamation
        GoTo GuardarCerrar
    End If

    ' --- NOTAS DEL TeCNICO ---
    Dim wsInterfaz As Worksheet, rngNotas As Range, celdaNota As Range
    Dim contadorNotas As Long
    Set wsInterfaz = ThisWorkbook.Worksheets("Interfaz")
    Set rngNotas = wsInterfaz.Range(ADDR_H2_H55)
    contadorNotas = 0
    For Each celdaNota In rngNotas
        If Trim$(CStr(celdaNota.Value)) <> "" Then contadorNotas = contadorNotas + 1
    Next celdaNota

    If contadorNotas > 0 Then
        With wordDoc.content
            .Collapse wdCollapseEnd
            .InsertBreak wdPageBreak
            .Collapse wdCollapseEnd
            .InsertAfter "Notas del Tecnico"
            ApplyStyleSafe wordDoc, .Paragraphs(.Paragraphs.Count).Range, "EstiloTituloTabla", "Normal"
            .InsertParagraphAfter
            .Collapse wdCollapseEnd
        End With

        contadorNotas = 1
        For Each celdaNota In rngNotas
            If Trim$(CStr(celdaNota.Value)) <> "" Then
                With wordDoc.content
                    .Collapse wdCollapseEnd
                    .InsertAfter "Nota " & contadorNotas & ":"
                    With .Paragraphs(.Paragraphs.Count).Range.Font
                        .Underline = True
                        .bold = False
                        .Color = RGB(0, 0, 0)
                    End With
                    .InsertParagraphAfter
                    .Collapse wdCollapseEnd

                    .InsertAfter CStr(celdaNota.Value)
                    ApplyStyleSafe wordDoc, .Paragraphs(.Paragraphs.Count).Range, "Normal", "Normal"
                    .InsertParagraphAfter
                    .Collapse wdCollapseEnd
                End With
                contadorNotas = contadorNotas + 1
            End If
        Next celdaNota
    End If

GuardarCerrar:
    wordDoc.SaveAs ruta & nombreFinal
    wordDoc.Close
    wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
    
    If Not wbTextos Is Nothing Then
        On Error Resume Next
        wbTextos.Close SaveChanges:=False
        On Error GoTo 0
        Set wbTextos = Nothing
        Set wsTextosExt = Nothing
    End If

    MsgBox "Documento Word generado en el escritorio:" & vbCrLf & nombreFinal, vbInformation

End Sub

' =======================
'   HELPERS
' =======================
Public Sub AsegurarEstilosBaseConSiguiente(ByVal wordDoc As Object)
    On Error Resume Next
    wordDoc.Styles("EstiloTituloTabla").NextParagraphStyle = wordDoc.Styles("Normal")
    wordDoc.Styles("EstiloSubtituloEdificio").NextParagraphStyle = wordDoc.Styles("Normal")
    On Error GoTo 0
End Sub

Public Function AbrirHojaTextosExterna(ByVal rutaCompleta As String, ByVal nombreHoja As String) As Worksheet
    On Error Resume Next
    Dim wb As Workbook, ws As Worksheet, oldAlerts As Boolean
    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Set wb = Application.Workbooks.Open(fileName:=rutaCompleta, ReadOnly:=True, _
                                        UpdateLinks:=False, AddToMru:=False, Local:=True)
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
    If Not wb Is Nothing Then
        On Error Resume Next
        Set ws = wb.Worksheets(nombreHoja)
        On Error GoTo 0
        If Not ws Is Nothing Then
            Set AbrirHojaTextosExterna = ws
            Exit Function
        Else
            wb.Close SaveChanges:=False
        End If
    End If
    Set AbrirHojaTextosExterna = Nothing
End Function

Public Function ObtenerDescripcionDesdeTextos(ByVal clave As String, _
                                               ByVal wsExt As Worksheet, _
                                               ByVal wsInt As Worksheet) As String
    Dim celda As Range, txt As String

    ' 1) Intentar en externo
    If Not wsExt Is Nothing Then
        Set celda = wsExt.Columns(1).Find(What:=CStr(clave), LookIn:=xlValues, LookAt:=xlWhole)
        If Not celda Is Nothing Then
            txt = CStr(celda.Offset(0, 1).Value)
        End If
    End If

    ' 2) Si no hubo suerte, intentar en interno
    If Len(txt) = 0 And Not wsInt Is Nothing Then
        Set celda = wsInt.Columns(1).Find(What:=CStr(clave), LookIn:=xlValues, LookAt:=xlWhole)
        If Not celda Is Nothing Then
            txt = CStr(celda.Offset(0, 1).Value)
        End If
    End If

    ' 3) Reemplazos (si hay texto)
    If Len(txt) > 0 Then
        With ThisWorkbook.Worksheets("Interfaz")
            txt = Replace(txt, "{superficie}", .Range(ADDR_C59).Value)
            txt = Replace(txt, "{nri}", .Range(ADDR_C57).Value)
            txt = Replace(txt, "{tipo}", .Range(ADDR_C56).Value)
            txt = Replace(txt, "{hidrantes}", "HIDRANTES")
        End With
    End If

    ObtenerDescripcionDesdeTextos = txt
End Function

' Crea/ajusta los estilos usados en el informe
Public Sub AsegurarEstilosBase(ByVal wordDoc As Object)
    EnsureParagraphStyle wordDoc, "EstiloTituloTabla", "Helvetica Neue Thin", 12, True, RGB(0, 112, 192), 6
    EnsureParagraphStyle wordDoc, "EstiloSubtituloEdificio", "Helvetica Neue Thin", 10, False, RGB(112, 112, 112), 4
End Sub

Private Sub EnsureParagraphStyle(ByVal wordDoc As Object, ByVal nombre As String, _
                                 ByVal fuente As String, ByVal tam As Single, ByVal bold As Boolean, _
                                 ByVal colorRGB As Long, ByVal spaceAfter As Single)
    Dim s As Object
    On Error Resume Next
    Set s = wordDoc.Styles(nombre)
    On Error GoTo 0
    If s Is Nothing Then
        Set s = wordDoc.Styles.add(Name:=nombre, Type:=1) ' 1 = wdStyleTypeParagraph
    End If
    With s.Font
        .Name = fuente
        .Size = tam
        .bold = bold
        .Color = colorRGB
    End With
    With s.ParagraphFormat
        .spaceAfter = spaceAfter
        .keepTogether = True
        .keepWithNext = True
        .WordWrap = True
        .WidowControl = False
        .Hyphenation = False
    End With
End Sub

Public Function UniqueBuildingsSorted(ws As Worksheet, ByVal rLast As Long) As Variant
    Dim dict As Object, r As Long, k As String, arr(), i As Long
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1                         ' textcompare

    For r = 2 To rLast
        k = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(k) > 0 Then If Not dict.Exists(k) Then dict.add k, k
    Next r

    If dict.Count = 0 Then Exit Function
    arr = dict.Keys
    ' ordenar (burbuja simple: n pequeo)
    Dim j As Long, tmp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If UCase$(arr(j)) < UCase$(arr(i)) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
    UniqueBuildingsSorted = arr
End Function

' Aplica un estilo de forma segura. Si no existe o da error, intenta crearlo
' y si an falla, aplica el estilo de reserva (fallbackName, p.ej. "Normal").
Public Sub ApplyStyleSafe(ByVal wordDoc As Object, ByVal rng As Object, _
                           ByVal styleName As String, ByVal fallbackName As String)
    On Error GoTo TryCreate
    rng.Style = styleName
    Exit Sub
TryCreate:
    Err.Clear
    ' Intentar crearlo (por si no estaba)
    AsegurarEstilosBase wordDoc
    On Error Resume Next
    rng.Style = styleName
    If Err.Number <> 0 Then
        Err.Clear
        rng.Style = fallbackName
    End If
End Sub

' Pega de forma robusta el contenido del portapapeles como tabla.
' Devuelve True si no hubo error en el pegado.
Private Function SafePasteTable(ByVal wordDoc As Object) As Boolean
    On Error GoTo ErrHandler
    Const wdCollapseEnd As Long = 0

    With wordDoc.content
        .Collapse wdCollapseEnd
        ' 1) Intento recomendado: PasteSpecial RTF (DataType:=1)
        .PasteSpecial DataType:=1
        SafePasteTable = True
        Exit Function
    End With

ErrHandler:
    ' 2) Fallback: pegado normal (por si RTF falla)
    On Error Resume Next
    wordDoc.content.Collapse wdCollapseEnd
    wordDoc.Application.Selection.Paste
    SafePasteTable = (Err.Number = 0)
    Err.Clear
End Function

' Texto central del encabezado (puedes ajustarlo aqu a tu gusto)
Public Function TextoCentralEncabezado() As String
    On Error Resume Next
    Dim itf As Worksheet
    Set itf = ThisWorkbook.Worksheets("Interfaz")
    If Not itf Is Nothing Then
        TextoCentralEncabezado = "Informe de Proteccion Contra Incendios" & vbCr & _
                                 "Tipo: " & CStr(itf.Range(celdaConfig).Value) & "   NRI: " & CStr(itf.Range(celdaNri).Value)
    Else
        TextoCentralEncabezado = "Informe de Proteccion Contra Incendios"
    End If
End Function

Public Sub ConstruirEncabezado(ByVal wordDoc As Object, _
                                ByVal rutaLogoIzq As String, _
                                ByVal rutaLogoDer As String, _
                                ByVal textoCentral As String)
    Const wdHeaderFooterPrimary As Long = 1
    Const wdAlignRowCenter As Long = 1
    Const wdAlignParagraphLeft As Long = 0
    Const wdAlignParagraphCenter As Long = 1
    Const wdParagraphAlignmentRight As Long = 2
    Const wdRowHeightExactly As Long = 2
    Const wdCellAlignVerticalCenter As Long = 1

    Dim hdr As Object, tb As Object, rng As Object
    Dim rutaI As String, rutaD As String
    rutaI = rutaLogoIzq: rutaD = rutaLogoDer

    ' Distancia del encabezado al borde superior (ajusta si lo ves muy arriba)
    wordDoc.PageSetup.HeaderDistance = wordDoc.Application.CentimetersToPoints(1#)

    Set hdr = wordDoc.Sections(1).Headers(wdHeaderFooterPrimary)
    hdr.Range.text = ""

    ' Tabla 1x3 ocupando 100% de ancho entre mrgenes
    Set tb = hdr.Range.Tables.add(Range:=hdr.Range, NumRows:=1, NumColumns:=3)
    With tb
        .Rows.Alignment = wdAlignRowCenter
        .PreferredWidthType = 1                  ' porcentaje
        .PreferredWidth = 100
        ' Anchos relativos: 25%  50%  25%
        .Columns(1).PreferredWidthType = 1: .Columns(1).PreferredWidth = 25
        .Columns(2).PreferredWidthType = 1: .Columns(2).PreferredWidth = 50
        .Columns(3).PreferredWidthType = 1: .Columns(3).PreferredWidth = 25
        .Borders.Enable = False

        ' Altura exacta y sin espacio extra
        .Rows.HeightRule = wdRowHeightExactly
        .Rows.Height = 40                        ' puntos; sube/baja a tu gusto (ej. 4555)
        .Range.ParagraphFormat.SpaceBefore = 0
        .Range.ParagraphFormat.spaceAfter = 0

        ' Quitar acolchados internos de celdas para pegar a los mrgenes
        .TopPadding = 0: .BottomPadding = 0
        .LeftPadding = 0: .RightPadding = 0
    End With

    ' ======== CELDA IZQUIERDA: LOGO ========
    Set rng = tb.Cell(1, 1).Range
    rng.text = ""
    If Len(Dir$(rutaI, vbNormal)) > 0 Then
        rng.InlineShapes.AddPicture fileName:=rutaI, LinkToFile:=False, SaveWithDocument:=True
        ' Fija ancho visual del logo (ajusta 95120 pt segn tu imagen)
        Call AjustarAnchuraExacta(rng.InlineShapes(1), 105)
        tb.Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Else
        rng.text = ""
    End If

    ' ======== CELDA CENTRAL: TEXTO ========
    With tb.Cell(1, 2).Range
        .text = textoCentral
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Font
            .Name = "Helvetica Neue Thin"
            .Size = 9
            .bold = False
            .Color = RGB(128, 128, 128)          ' un gris similar al de tu ejemplo
        End With
    End With

    ' ======== CELDA DERECHA: LOGO ========
    Set rng = tb.Cell(1, 3).Range
    rng.text = ""
    If Len(Dir$(rutaD, vbNormal)) > 0 Then
        rng.InlineShapes.AddPicture fileName:=rutaD, LinkToFile:=False, SaveWithDocument:=True
        Call AjustarAnchuraExacta(rng.InlineShapes(1), 105)
        tb.Cell(1, 3).Range.ParagraphFormat.Alignment = wdParagraphAlignmentRight
    Else
        rng.text = ""
        tb.Cell(1, 3).Range.ParagraphFormat.Alignment = wdParagraphAlignmentRight
    End If

    ' Alinear verticalmente al centro los tres contenidos
    On Error Resume Next
    tb.Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    If Err.Number <> 0 Then
        ' Fallback por si alguna versin no admite asignacin en bloque
        Err.Clear
        Dim cel As Object
        For Each cel In tb.Range.Cells
            cel.VerticalAlignment = wdCellAlignVerticalCenter
        Next cel
    End If
    On Error GoTo 0
End Sub

Private Sub AjustarAnchuraExacta(ByVal ish As Object, ByVal anchoPuntos As Single)
    On Error Resume Next
    If anchoPuntos > 0 Then
        Dim factor As Double
        factor = anchoPuntos / ish.Width
        ish.Width = ish.Width * factor
        ish.Height = ish.Height * factor
    End If
End Sub

' Escala un InlineShape para no superar un ancho mximo (puntos)
Private Sub AjustarAnchuraMax(ByVal ish As Object, ByVal anchoMaxPuntos As Single)
    On Error Resume Next
    If ish.Width > anchoMaxPuntos Then
        Dim factor As Double
        factor = anchoMaxPuntos / ish.Width
        ish.Width = ish.Width * factor
        ish.Height = ish.Height * factor
    End If
End Sub

' Normaliza claves: quita dobles espacios, trim y pasa a mayusculas para comparar
Private Function NormKey(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    NormKey = UCase$(t)
End Function

Public Sub PegarTablaDesdeBloqueContiguo( _
        ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long, _
        ByVal edificio As String, ByVal wordDoc As Object, _
        Optional ByVal sortRows As Boolean = True)
    
    Dim tmp As Worksheet, nombreTmp As String
    Dim r As Long, c As Long, outRow As Long
    Dim colLastLetter As String
    Dim KFiltro As String, KFila As String
    Dim t As Object, tablasAntes As Long
    Dim ur As Range

    Application.ScreenUpdating = False

    nombreTmp = NombreTemporalLibre(ThisWorkbook, "TMP_Contiguo")
    Set tmp = ThisWorkbook.Worksheets.add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    tmp.Name = nombreTmp

    colLastLetter = Split(ws.Cells(1, lastCol).Address(False, False), "$")(0)
    KFiltro = NormKey(edificio)

    ' 1) Copiar encabezado como valores
    For c = 1 To lastCol
        tmp.Cells(1, c).Value = ws.Cells(1, c).Value
    Next c
    outRow = 2

    ' 2) Copiar filas: todas o filtradas por edificio (col A)
    If Len(KFiltro) = 0 Then
        For r = 2 To lastRow
            If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))) > 0 Then
                For c = 1 To lastCol
                    tmp.Cells(outRow, c).Value = ws.Cells(r, c).Value
                Next c
                outRow = outRow + 1
            End If
        Next r
    Else
        For r = 2 To lastRow
            KFila = NormKey(CStr(ws.Cells(r, 1).Value))
            If KFila = KFiltro Then
                For c = 1 To lastCol
                    tmp.Cells(outRow, c).Value = ws.Cells(r, c).Value
                Next c
                outRow = outRow + 1
            End If
        Next r
    End If

    Set ur = tmp.UsedRange
    If ur.Rows.Count <= 1 Then GoTo fin

    ur.Copy

    ' 3) Pegado robusto en Word
    tablasAntes = wordDoc.Tables.Count
    If SafePasteTable(wordDoc) And wordDoc.Tables.Count > tablasAntes Then
        Set t = wordDoc.Tables(wordDoc.Tables.Count)
        With t
            .Range.Font.Name = "Arial"
            .Range.Font.Size = 9
            .Rows(1).Range.Font.bold = True
            .Rows(1).Shading.BackgroundPatternColor = RGB(217, 217, 217)
            .Range.ParagraphFormat.spaceAfter = 6
            .Range.ParagraphFormat.Alignment = 1
            .Range.ParagraphFormat.Hyphenation = False
            .Range.ParagraphFormat.WordWrap = True

            Dim rObj As Object, cObj As Object
            For Each rObj In .Rows
                rObj.AllowBreakAcrossPages = False
            Next rObj
            For Each cObj In .Range.Cells
                cObj.Range.ParagraphFormat.Hyphenation = False
                cObj.Range.ParagraphFormat.WordWrap = True
            Next cObj

            With .Borders
                .InsideLineStyle = 1
                .OutsideLineStyle = 1
            End With
            .AutoFitBehavior 1
            .PreferredWidthType = 1
            .PreferredWidth = 100
        End With
    End If

fin:
    Application.CutCopyMode = False
    Application.DisplayAlerts = False
    On Error Resume Next
    tmp.Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Public Sub CerrarTodasInstanciasWord()
    Dim objWMI As Object, colProcesses As Object, objProcess As Object
    On Error Resume Next
    Set objWMI = GetObject("winmgmts:\\.\root\CIMV2")
    Set colProcesses = objWMI.ExecQuery("Select * from Win32_Process Where Name='WINWORD.EXE'")
    For Each objProcess In colProcesses
        objProcess.Terminate
    Next
    On Error GoTo 0
End Sub

Function HayProcesosWord() As Boolean
    Dim objWMI As Object, colProcesses As Object
    On Error Resume Next
    Set objWMI = GetObject("winmgmts:\\.\root\CIMV2")
    Set colProcesses = objWMI.ExecQuery("Select * from Win32_Process Where Name='WINWORD.EXE'")
    HayProcesosWord = (colProcesses.Count > 0)
    On Error GoTo 0
End Function

' Ajusta la lista de entradas que debe limpiar la hoja Interfaz
Private Function RangosEntradaInterfaz() As Variant
    RangosEntradaInterfaz = Array( _
                            "B2", "D2", "F3", "B3", "D3", "B7", "B8", "B9", "G3", "G5", _
                            "B10", "B11", "B12", "B14", "B16", "B17", "B18", "B19", "B20", "B21", "B22", _
                            "F7", "F8", "F9", "F10", "F11", "F12", "F13", "F14", "F15", "F16", "F17", _
                            "B24", "F24", "F25", "F26", "B28", "B29", "B30", "B31", "B32", "F28", "F29", "F30", _
                            "F31", "F32", "B46", "B47", "F46", "B49", "A51", "B51", "C51", "D51", "E51", "F51", _
                            "A53", "B53", "C53", "D53", "E53", "F53", "A55", "B55", "C55", "D55", "E55", "F55", "D29", "D28")
End Function

Private Function HojaExiste(ByVal nombreHoja As String, Optional ByVal wb As Workbook) As Boolean
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set ws = wb.Worksheets(nombreHoja)
    HojaExiste = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0
End Function

Private Function SiguienteNombreCaso() As String
    Dim n As Long, nm As String
    n = 1
    Do
        nm = Format$(n, "000")
        If Not HojaExiste(nm, ThisWorkbook) Then
            SiguienteNombreCaso = nm
            Exit Function
        End If
        n = n + 1
    Loop
End Function

Private Function EsNombreCaso(ByVal nombre As String) As Boolean
    EsNombreCaso = (Trim$(nombre) Like "###")
End Function

Private Function HojaDestino(ByVal wb As Workbook, ByVal wsInterfaz As Worksheet) As Worksheet
    Dim ws As Worksheet, idxUlt As Long: idxUlt = -1
    For Each ws In wb.Worksheets
        If EsNombreCaso(ws.Name) Then
            If ws.Index > idxUlt Then
                idxUlt = ws.Index
                Set HojaDestino = ws
            End If
        End If
    Next ws
    If HojaDestino Is Nothing Then Set HojaDestino = wsInterfaz
End Function

Private Sub PegarSoloValores(ByVal rng As Range)
    If Not rng Is Nothing Then rng.Value = rng.Value
End Sub

Private Sub LimpiarEntradasEnInterfaz(ByVal ws As Worksheet)
    Dim arr As Variant, i As Long
    arr = RangosEntradaInterfaz()
    For i = LBound(arr) To UBound(arr)
        On Error Resume Next
        ws.Range(arr(i)).ClearContents
        On Error GoTo 0
    Next i
End Sub

' ===== Archivar desde interfaz

Public Sub ArchivarCasoDesdeInterfaz(Optional ByVal wsBase As Worksheet)
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsCaso As Worksheet
    Dim nombreCaso As String
    Dim prevCalc As XlCalculation
    Dim wsAfter As Worksheet

    If wsBase Is Nothing Then
        Set wsBase = ActiveSheet
    End If
    If wsBase Is Nothing Then
        MsgBox "No se ha podido determinar la hoja Interfaz.", vbExclamation
        Exit Sub
    End If

    On Error GoTo fin_limpio
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    prevCalc = Application.Calculation
    Application.Calculation = xlCalculationManual

    ' Crear nombre de caso
    nombreCaso = SiguienteNombreCaso()

    ' Insertar copia a continuacion de la hoja destino (uUltimo caso o la Interfaz)
    Set wsAfter = HojaDestino(wb, wsBase)
    wsBase.Copy After:=wsAfter

    Set wsCaso = ActiveSheet
    wsCaso.Name = nombreCaso


    ' Modo "solo recalcular" y limitar la vista a A1:H55
    PrepararHojaSoloRecalculo wsCaso

    ' Marcar (opcional)
    On Error Resume Next
    wsCaso.Tab.Color = RGB(52, 152, 219)
    wsCaso.Range(ADDR_A1).AddComment "Caso archivado el " & Format(Now, "yyyy-mm-dd hh:nn")
    On Error GoTo 0

    ' Limpiar entradas en la Interfaz base (si asi lo quieres)
    LimpiarEntradasEnInterfaz wsBase
    EstablecerPredeterminadosInterfaz wsBase
    
    MsgBox "Caso archivado: " & nombreCaso & vbCrLf, vbInformation

fin_limpio:
    On Error Resume Next
    Application.Calculation = prevCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Public Sub PrepararHojaSoloRecalculo(ByVal ws As Worksheet)

    ws.Columns("I:XFD").EntireColumn.Hidden = True
    ws.Rows("56:" & ws.Rows.Count).EntireRow.Hidden = True

    ws.ScrollArea = "A1:H55"

    ws.Range(ADDR_A1).Select
End Sub

Public Sub LimpiarAnexos()
    
    Const PASS_HOJA As String = ""
    Dim anexos As Variant, i As Long
    Dim ws As Worksheet, lo As ListObject
    Dim lr As Long, hizoAlgoHoja As Boolean
    Dim acciones As Long, noEncontradas As Long, yaVacias As Long
    Dim calcPrev As XlCalculation, evPrev As Boolean, updPrev As Boolean

    anexos = Array("Anx. II Sup", "Anx. II Reac", "Anx. II Res Sec", "Anx. II Res Ext", _
                   "Anx. II Ocu", "Anx. II Sal", "Anx. II Res", _
                   "Anx. III Det Auto", "Anx. III Det Man", "Anx. III Alr", _
                   "Anx. III Meg", "Anx. III Hid", "Anx. III Ext Ext", "Anx. III BIES", _
                   "Anx. III Ext Auto", "Anx. III Col", "Anx. III Hum")

    ' --- congelar UI, recordar estado
    updPrev = Application.ScreenUpdating: Application.ScreenUpdating = False
    evPrev = Application.EnableEvents:    Application.EnableEvents = False
    calcPrev = Application.Calculation:   Application.Calculation = xlCalculationManual

    On Error GoTo ErrTrap

    For i = LBound(anexos) To UBound(anexos)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(anexos(i)))
        On Error GoTo ErrTrap

        If ws Is Nothing Then
            noEncontradas = noEncontradas + 1
        Else
            ' Desproteger si procede
            On Error Resume Next: ws.Unprotect Password:=PASS_HOJA: On Error GoTo ErrTrap
            hizoAlgoHoja = False

            ' 1) Vaciar filas de datos de las Tablas (si hay)
            If ws.ListObjects.Count > 0 Then
                For Each lo In ws.ListObjects
                    If Not lo.DataBodyRange Is Nothing Then
                        lo.DataBodyRange.Delete  ' elimina todas las filas de datos
                        acciones = acciones + 1
                        hizoAlgoHoja = True
                    End If
                Next lo
            End If

            ' 2) Si no hay Tablas (o ya estaban vacas), limpiar filas 2..ltima usada
            If Not hizoAlgoHoja Then
                lr = LastUsedRow(ws)
                If lr > 1 Then
                    ws.Range(ws.Rows(2), ws.Rows(lr)).ClearContents
                    acciones = acciones + 1
                    hizoAlgoHoja = True
                Else
                    yaVacias = yaVacias + 1
                End If
            End If

            ' Reproteger si quieres
            On Error Resume Next: ws.Protect Password:=PASS_HOJA, UserInterfaceOnly:=True: On Error GoTo ErrTrap
        End If
    Next i

Finalizar:
    ' Restaurar estado UI
    Application.Calculation = calcPrev
    Application.EnableEvents = evPrev
    Application.ScreenUpdating = updPrev
    Exit Sub

ErrTrap:
    ' Asegura salida limpia y muestra el error
    Application.Calculation = calcPrev
    Application.EnableEvents = evPrev
    Application.ScreenUpdating = updPrev
    MsgBox "LimpiarAnexos: error " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Function LastUsedRow(ws As Worksheet) As Long
    Dim r As Range
    On Error Resume Next
    Set r = ws.Cells.Find(What:="*", After:=ws.Range(ADDR_A1), LookIn:=xlFormulas, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If r Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = r.row
    End If
End Function

Public Sub BorrarCasosNumericos()
    Dim wb As Workbook
    Dim i As Long
    Dim ws As Worksheet
    Dim nm As String
    Dim borradas As Long
    Dim saltadas As String

    Set wb = ActiveWorkbook                      ' <- clave: trabajamos sobre el libro activo

    ' Estructura protegida?
    If wb.ProtectStructure Then
        MsgBox "No se pueden eliminar hojas porque la ESTRUCTURA del libro esta protegida." & vbCrLf & _
               "Revisa: Revisar > Proteger libro (estructura).", vbExclamation
        Exit Sub
    End If

    Application.DisplayAlerts = False

    ' Recorremos al revs para poder borrar sin problemas
    For i = wb.Worksheets.Count To 1 Step -1
        Set ws = wb.Worksheets(i)
        nm = ws.Name

        ' Normalizamos: quitamos NBSP y espacios
        nm = Replace(nm, Chr(160), "")           ' NBSP
        nm = Trim$(nm)

        If EsHojaCaso(nm) Then
            On Error Resume Next
            ws.Delete
            If Err.Number <> 0 Then
                saltadas = saltadas & vbCrLf & " - " & ws.Name & " (" & Err.Description & ")"
                Err.Clear
            Else
                borradas = borradas + 1
            End If
            On Error GoTo 0
        End If
    Next i

    Application.DisplayAlerts = True

    Dim msg As String
    msg = borradas & " hoja(s) de caso eliminadas."
    If Len(saltadas) > 0 Then
        msg = msg & vbCrLf & "No se pudieron borrar:" & saltadas
        MsgBox msg, vbExclamation
    Else
        MsgBox msg, vbInformation
    End If
End Sub

Private Function EsHojaCaso(ByVal nm As String) As Boolean
    ' Coincide con 1-3 dgitos (1, 02, 123) o con tus nombres antiguos
    EsHojaCaso = (nm Like "#") Or (nm Like "##") Or (nm Like "###") _
                 Or (LCase$(Left$(nm, 14)) = "interfaz_caso_")
End Function

Public Sub ExtraerCasosATablas()

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim baseName As String: baseName = "Interfaz"
    Dim tmpName As String
    Dim wsBase As Worksheet
    Dim casos As Collection
    Dim nm As Variant
    Dim resumen As String
    Dim usarSup As Boolean
    Dim prevCalc As XlCalculation
    Dim respuesta As VbMsgBoxResult

    ' Hoja Interfaz
    Set wsBase = HojaExisteLocal(wb, baseName)
    If wsBase Is Nothing Then
        MsgBox "No se encuentra la hoja '" & baseName & "'.", vbCritical
        Exit Sub
    End If

    ' Listar hojas de caso (001, 002, ...)
    Set casos = ListarHojasCasoNumericas(wb)
    If casos Is Nothing Or casos.Count = 0 Then
        MsgBox "No se han encontrado hojas de caso (001, 002, 003...).", vbInformation
        Exit Sub
    End If

    ' Viabilidad desde Interfaz!F7 (true si numero <> 0, o texto si/true/verdadero)
    Dim vF7 As Variant, s As String
    vF7 = wsBase.Range(celdaResultadoEstado).Value
    usarSup = False
    If IsNumeric(vF7) Then
        usarSup = (Val(vF7) <> 0)
    Else
        s = LCase$(CStr(vF7))
        usarSup = (s = "si" Or s = "true" Or s = "verdadero")
    End If


    ' Aparcar Interfaz con nombre temporal seguro
    tmpName = NombreTemporalLibreSeguro(wb, baseName & "_BASE")

    ' Optimizaciones
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    prevCalc = Application.Calculation
    Application.Calculation = xlCalculationManual

    On Error GoTo FIN_SEGURO

    ' Renombrar Interfaz original a temporal
    wsBase.Name = tmpName

    ' Procesar cada caso (001, 002, ...)
    For Each nm In casos
        Dim wsCaso As Worksheet
        Set wsCaso = HojaExisteLocal(wb, CStr(nm))
        If wsCaso Is Nothing Then
            resumen = resumen & "Caso " & CStr(nm) & ": no encontrado" & vbCrLf
        Else
            ' Montar el caso como "Interfaz"
            wsCaso.Name = baseName

            ' Recalculo completo por dependencias
            Application.CalculateFull

            ' Ejecutar copias silenciosas; devuelve texto con incidencias
            Dim errores As String
            On Error Resume Next
            errores = EjecutarCopiasSilenciosas(usarSup)
            On Error GoTo FIN_SEGURO

            If Len(errores) = 0 Then
                resumen = resumen & "Caso " & CStr(nm) & ": OK" & vbCrLf
            Else
                resumen = resumen & "Caso " & CStr(nm) & ": ERRORES" & vbCrLf & errores & vbCrLf
            End If

            ' Devolver nombre del caso
            wb.Worksheets(baseName).Name = CStr(nm)
        End If
    Next nm

    ' Restaurar Interfaz original
    wb.Worksheets(tmpName).Name = baseName

FIN_SEGURO:
    ' Restaurar entorno
    Application.Calculation = prevCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' Intento de restauracion segura si hubo error entre renombres
    If Err.Number <> 0 Then
        On Error Resume Next
        If Not HojaExisteLocal(wb, baseName) Is Nothing Then
            ' ok
        ElseIf Not HojaExisteLocal(wb, tmpName) Is Nothing Then
            wb.Worksheets(tmpName).Name = baseName
        End If
        On Error GoTo 0
        MsgBox "Se produjo un error: " & Err.Description, vbExclamation
    End If

    If Len(resumen) = 0 Then resumen = "No se proceso ningun caso."
    MsgBox resumen, vbInformation, "Extraccion por casos"
End Sub

Public Function HojaExisteLocal(ByVal wb As Workbook, ByVal nombre As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nombre)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        Set HojaExisteLocal = ws
    Else
        Set HojaExisteLocal = Nothing
    End If
End Function

' Devuelve True si la hoja esta protegida para contenido/dibujo/escenarios
Public Function HojaEstaProtegida(ByVal ws As Worksheet) As Boolean
    HojaEstaProtegida = (ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios)
End Function

' Lista estandar de anexos (ajusta si hace falta)
Public Function ListaAnexosEstandar() As Variant
    ListaAnexosEstandar = Array( _
                          "Anx. II Sup", "Anx. II Reac", "Anx. II Res Sec", "Anx. II Res Ext", _
                          "Anx. II Ocu", "Anx. II Sal", "Anx. II Res", _
                          "Anx. III Det Auto", "Anx. III Det Man", "Anx. III Alr", _
                          "Anx. III Meg", "Anx. III Hid", "Anx. III Ext Ext", "Anx. III BIES", _
                          "Anx. III Ext Auto", "Anx. III Col", "Anx. III Hum" _
                                                              )
End Function

' Genera un nombre temporal garantizado libre.
Public Function NombreTemporalLibreSeguro(ByVal wb As Workbook, ByVal base As String) As String
    Dim i As Long, candidato As String
    If Len(base) = 0 Then base = "TMP"
    candidato = base
    i = 1
    Do While Not (HojaExisteLocal(wb, candidato) Is Nothing)
        candidato = base & "_" & CStr(i)
        i = i + 1
    Loop
    NombreTemporalLibreSeguro = candidato
End Function

' --- Devuelve un nombre temporal que no exista aun ---
Private Function NombreTemporalLibre(wb As Workbook, ByVal base As String) As String
    Dim i As Long, candidato As String
    candidato = base
    i = 1
    ' OJO: HojaExiste(nombreHoja, wb). Si lo llamas al reves fallara por tipos.
    Do While HojaExiste(candidato, wb)
        candidato = base & "_" & CStr(i)
        i = i + 1
    Loop
    NombreTemporalLibre = candidato
End Function

' --- Wrapper seguro: existe hoja con ese nombre? ---
Private Function HojaExisteNombre(wb As Workbook, ByVal nombreHoja As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nombreHoja)
    HojaExisteNombre = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0
End Function

' --- Ejecuta el mismo bloque que tu EjecutarExtracciones, pero en silencio y sin preguntar ---
Private Function EjecutarCopiasSilenciosas(ByVal usarSupNormal As Boolean) As String
    Dim errores As String
    Dim ok As Boolean

    ' Resto (mismo orden que usas en EjecutarExtracciones)
    If Not CopiarDatosSup(False) Then errores = errores & "- Viabilidad de la superficie." & vbCrLf
    If Not CopiarDatosReac(False) Then errores = errores & "- Reaccion al fuego." & vbCrLf
    If Not CopiarDatosResSec(False) Then errores = errores & "- Resistencia separadores de sectores." & vbCrLf
    If Not CopiarDatosResExt(False) Then errores = errores & "- Resistencia separadores con otros establecimientos." & vbCrLf
    If Not CopiarDatosOcupacion(False) Then errores = errores & "- Ocupacion." & vbCrLf
    If Not CopiarDatosSalidas(False) Then errores = errores & "- Salidas de emergencia." & vbCrLf
    If Not CopiarDatosRes(False) Then errores = errores & "- Resistencia elementos estructurales." & vbCrLf
    If Not CopiarDatosDetAuto(False) Then errores = errores & "- Deteccion automtica." & vbCrLf
    If Not CopiarDatosMan(False) Then errores = errores & "- Deteccion manual." & vbCrLf
    If Not CopiarDatosAlr(False) Then errores = errores & "- Alarmas." & vbCrLf
    If Not CopiarDatosMeg(False) Then errores = errores & "- Megafonia / alarmas." & vbCrLf
    If Not CopiarDatosHidrantes(False) Then errores = errores & "- Hidrantes." & vbCrLf
    If Not CopiarDatosExt(False) Then errores = errores & "- Extintores." & vbCrLf
    If Not CopiarDatosBie(False) Then errores = errores & "- BIEs." & vbCrLf
    If Not CopiarDatosCol(False) Then errores = errores & "- Columna seca." & vbCrLf
    If Not CopiarDatosAuto(False) Then errores = errores & "- Extincion automatica." & vbCrLf
    If Not CopiarDatosHumos(False) Then errores = errores & "- Extraccion de humos." & vbCrLf

    EjecutarCopiasSilenciosas = errores
End Function

' --- Lista hojas cuyo nombre es 1-3 dgitos: 1, 02, 123 ---
Private Function ListarHojasCasoNumericas(wb As Workbook) As Collection
    Dim Col As New Collection
    Dim ws As Worksheet
    Dim nm As String
    For Each ws In wb.Worksheets
        nm = ws.Name
        nm = Replace(nm, Chr(160), "")           ' NBSP
        nm = Replace(nm, " ", "")
        nm = Trim$(nm)
        If (nm Like "#") Or (nm Like "##") Or (nm Like "###") Then
            Col.add ws.Name
        End If
    Next ws
    Set ListarHojasCasoNumericas = Col
End Function

Public Sub CalcularTodo(Optional ByVal ws As Worksheet)
    Dim tareas As Variant, i As Long
    Dim resumen As String
    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean, prevScreen As Boolean
    Dim t0 As Double: t0 = Timer
    Dim huboErrores As Boolean: huboErrores = False
    Dim estado As String
    
    If ws Is Nothing Then Set ws = ActiveSheet
    ws.Activate                                  ' <- hace activa la hoja de trabajo
    Application.GoTo ws.Range("A1"), False       ' opcional: asegura el foco en esa hoja

    ' Orden SIN CalcularNota5 (la lanzamos nosotros condicionalmente)
    tareas = Array( _
             "VerificarViabilidad", _
             "VerificarResistencia", _
             "ObtenerRevestimientos", _
             "VerificarResistenciaSeparadores", _
             "reaccionElementos", _
             "ocupacion", _
             "calcularSalidas", _
             "calcularResistenciaEstructural", _
             "calcularSistemasDeteccion", _
             "calcularHidrantes", _
             "calcularExtintores", _
             "calcularBies", _
             "calcularColumnaSeca", _
             "calcularExtincioAuto", _
             "calcularHumos" _
             )

    ' Preparacion segura
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    On Error GoTo FIN_SEGURO
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Asegurar que las rutinas trabajan sobre la hoja correcta
    ws.Activate

    ' --- 1) Ejecutar bloque principal (sin Nota 5)
    For i = LBound(tareas) To UBound(tareas)
        On Error Resume Next
        Err.Clear
        Application.run tareas(i)                ' Estas rutinas operan sobre ActiveSheet
        If Err.Number = 0 Then
            resumen = resumen & " " & tareas(i) & ": OK" & vbCrLf
        Else
            resumen = resumen & " " & tareas(i) & ": ERROR -> " & Err.Description & vbCrLf
            huboErrores = True
        End If
        On Error GoTo 0
    Next i

    ' --- 2) Si NO esta admitido por viabilidad normal, intentar Nota 5
    ' OJO: celdaResultadoEstado debe ser tu constante de rango, p.ej. "I1"
    estado = CStr(ws.Range(celdaResultadoEstado).Value)
    If UCase$(estado) <> "ADMITIDO" Then
        On Error Resume Next
        Err.Clear
        Application.run "CalcularNota5"          ' Trabaja sobre la hoja activa (ws)
        If Err.Number = 0 Then
            resumen = resumen & " CalcularNota5: OK" & vbCrLf
        Else
            resumen = resumen & " CalcularNota5: ERROR -> " & Err.Description & vbCrLf
            huboErrores = True
        End If
        On Error GoTo 0

        ' Releer estado tras Nota 5 (en la misma hoja ws)
        estado = CStr(ws.Range(celdaResultadoEstado).Value)
        If UCase$(estado) <> "ADMITIDO" And UCase$(estado) <> "ADMITIDO CON NOTA 5" Then
            ws.Range(celdaResultadoEstado).Value = "NO ADMITIDO"
        End If
    End If

FIN_SEGURO:
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen

    ' Pintar estado en la hoja que llama
    If huboErrores Then
        EstadoCaso_MarcarModificado ws
    Else
        EstadoCaso_MarcarCalculo ws
        EstadoCaso_SetInicializado ws, True
    End If

    ' Mensaje resumen
    resumen = "Calculo completo (" & Format(Timer - t0, "0.0") & " s):" & vbCrLf & resumen
    MsgBox resumen, IIf(huboErrores, vbExclamation, vbInformation), "Resumen de calculo - " & ws.Name
    
    On Error Resume Next
    ' GenerarHojaResumenVGP
    On Error GoTo 0
End Sub

Public Function RangoUnionDeEntradas(ByVal ws As Worksheet) As Range
    Dim arr, i As Long, r As Range
    arr = EntradasVigiladas(ws)
    On Error GoTo Salir

    For i = LBound(arr) To UBound(arr)
        If r Is Nothing Then
            Set r = ws.Range(CStr(arr(i)))
        Else
            Set r = Application.Union(r, ws.Range(CStr(arr(i))))
        End If
    Next i
Salir:
    Set RangoUnionDeEntradas = r
End Function

' ===== Meta (huella por hoja) =====
' Guardamos, comparamos y actualizamos huella en una hoja oculta "_MetaCasos"
Private Sub Meta_Asegurar()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("_MetaCasos")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.add(Before:=wb.Worksheets(1))
        ws.Name = "_MetaCasos"
        ws.Visible = xlSheetVeryHidden
        ws.Range(ADDR_A1_D1).Value = Array("Hoja", "UUltimoCalculo", "Huella", "Notas")
        ws.Columns("A:D").ColumnWidth = 28
    End If
End Sub

Private Function Meta_CambioRespectoHuella(ByVal ws As Worksheet) As Boolean
    Call Meta_Asegurar
    Dim META As Worksheet: Set META = ThisWorkbook.Worksheets("_MetaCasos")
    Dim fila As Long: fila = Meta_Fila(ws.Name, False)
    If fila = 0 Then
        Meta_CambioRespectoHuella = True         ' nunca calculado antes
        Exit Function
    End If

    Dim hOld As String: hOld = CStr(META.Cells(fila, 3).Value)
    Dim hNow As String: hNow = CalcularHuella(ws)
    Meta_CambioRespectoHuella = (hOld <> hNow)
End Function

Private Function Meta_Fila(ByVal nombreHoja As String, ByVal crear As Boolean) As Long
    Dim META As Worksheet: Set META = ThisWorkbook.Worksheets("_MetaCasos")
    Dim lastRow As Long: lastRow = META.Cells(META.Rows.Count, 1).End(xlUp).row
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(CStr(META.Cells(r, 1).Value), nombreHoja, vbTextCompare) = 0 Then
            Meta_Fila = r
            Exit Function
        End If
    Next r
    If crear Then
        Meta_Fila = lastRow + 1
    Else
        Meta_Fila = 0
    End If
End Function

' Huella simple y robusta (hash 33x XOR sobre valores de entradas)
Private Function CalcularHuella(ByVal ws As Worksheet) As String
    Dim arr As Variant: arr = EntradasVigiladas()
    Dim i As Long, s As String
    Dim v As Variant

    On Error Resume Next
    For i = LBound(arr) To UBound(arr)
        v = ws.Range(CStr(arr(i))).Value
        s = s & "|" & FormatearValor(v)
    Next i
    On Error GoTo 0

    Dim h As Long: h = 5381
    Dim j As Long, ch As Integer
    For j = 1 To Len(s)
        ch = AscW(Mid$(s, j, 1))
        h = ((h * 33) Xor ch) And &H7FFFFFFF
    Next j
    CalcularHuella = CStr(h)
End Function

Private Function FormatearValor(ByVal v As Variant) As String
    If IsError(v) Then
        FormatearValor = "#ERR!"
    ElseIf IsDate(v) Then
        FormatearValor = Format$(v, "yyyymmdd\THH:nn:ss")
    ElseIf IsNumeric(v) Then
        FormatearValor = Format$(CDbl(v), "0.############")
    ElseIf VarType(v) = vbBoolean Then
        FormatearValor = IIf(CBool(v), "TRUE", "FALSE")
    ElseIf IsEmpty(v) Then
        FormatearValor = ""
    Else
        FormatearValor = CStr(v)
    End If
End Function

Public Sub EstadoCaso_MarcarCalculo(ByVal ws As Worksheet)
    Dim r As Range, prot As Boolean
    Set r = ObtenerCeldaEstado(ws)
    If r Is Nothing Then Exit Sub

    prot = ws.ProtectContents
    If prot Then On Error Resume Next: ws.Unprotect "0000": On Error GoTo 0

    With r
        .FormatConditions.Delete
        .Value = "CALCULADO: " & Format(Now, "dd/mm/yyyy hh:nn:ss")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.bold = True
        .Font.Color = vbWhite
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(67, 160, 71)       ' verde
        .Borders.LineStyle = xlContinuous
        .locked = False                          ' opcional: que no bloquee escritura
    End With

    If prot Then On Error Resume Next: ws.Protect "0000": On Error GoTo 0
End Sub

Public Sub EstadoCaso_MarcarModificado(ByVal ws As Worksheet)
    Dim r As Range, prot As Boolean
    Set r = ObtenerCeldaEstado(ws)
    If r Is Nothing Then Exit Sub

    prot = ws.ProtectContents
    If prot Then On Error Resume Next: ws.Unprotect "0000": On Error GoTo 0

    With r
        .FormatConditions.Delete
        .Value = "MODIFICADO: " & Format(Now, "dd/mm/yyyy hh:nn:ss")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.bold = True
        .Font.Color = vbWhite
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(211, 47, 47)       ' rojo
        .Borders.LineStyle = xlContinuous
        .locked = False
    End With

    If prot Then On Error Resume Next: ws.Protect "0000": On Error GoTo 0
End Sub

Public Sub EstadoCaso_OnChange(ByVal ws As Worksheet, ByVal Target As Range)
    Dim rWatch As Range, rEstado As Range

    ' Si estamos en un calculo, no reaccionar
    If g_BloquearEventos Then Exit Sub

    ' No reaccionar si el propio estado (G6) cambia
    Set rEstado = ObtenerCeldaEstado(ws)
    If Not rEstado Is Nothing Then
        If Not Intersect(Target, rEstado) Is Nothing Then Exit Sub
    End If

    Set rWatch = RangoUnionDeEntradas(ws)
    If rWatch Is Nothing Then Exit Sub

    If Not Intersect(Target, rWatch) Is Nothing Then
        EstadoCaso_MarcarModificado ws
    End If
End Sub

Public Function ObtenerCeldaEstado(ByVal ws As Worksheet) As Range
    Dim r As Range

    ' 1) Intentar nombre local "EstadoCaso"
    On Error Resume Next
    Set r = Nothing
    Set r = ws.Names("EstadoCaso").RefersToRange
    On Error GoTo 0

    ' 2) Si no existe el nombre local, usar la celda por defecto (G6)
    If r Is Nothing Then
        On Error Resume Next
        Set r = ws.Range(CELDA_ESTADO_POR_DEFECTO) ' "G6"
        On Error GoTo 0
    End If

    ' 3) Si aun es Nothing (G6 no existe por algun motivo), intenta crear el nombre local y devolver G6
    If r Is Nothing Then
        On Error Resume Next
        ws.Names.add Name:="EstadoCaso", RefersTo:=ws.Range(ADDR_G6)
        Set r = ws.Range(ADDR_G6)
        On Error GoTo 0
    End If

    Set ObtenerCeldaEstado = r
End Function

Private Function EntradasVigiladas() As Variant
    Dim arr As Variant
    On Error Resume Next
    arr = Application.run("RangosEntradaInterfaz")
    On Error GoTo 0
    If IsEmpty(arr) Then
        arr = Array( _
              "B2", "D2", "F3", "B3", "D3", "B7", "B8", "B9", "G3", "G5", _
              "B10", "B11", "B12", "B14", "B16", "B17", "B18", "B19", "B20", "B21", "B22", _
              "F7", "F8", "F9", "F10", "F11", "F12", "F13", "F14", "F15", "F16", "F17", _
              "B24", "F24", "F25", "F26", "B28", "B29", "B30", "B31", "B32", "F28", "F29", "F30", _
              "F31", "F32", "B46", "B47", "F46", "B49", "A51", "B51", "C51", "D51", "E51", "F51", _
              "A53", "B53", "C53", "D53", "E53", "F53", "A55", "B55", "C55", "D55", "E55", "F55", "D29", "D28")
    End If
    EntradasVigiladas = arr
End Function

Private Function EntradasUnion_Cache(ByVal ws As Worksheet) As Range
    Static dict As Object                        ' Scripting.Dictionary
    If dict Is Nothing Then Set dict = CreateObject("Scripting.Dictionary")
    Dim key As String: key = CStr(ws.codeName)
    If dict.Exists(key) Then
        Set EntradasUnion_Cache = dict(key)
        Exit Function
    End If
    Dim arr As Variant: arr = EntradasVigiladas()
    Dim i As Long
    Dim u As Range, r As Range
    On Error Resume Next
    For i = LBound(arr) To UBound(arr)
        Set r = ws.Range(CStr(arr(i)))
        If Not r Is Nothing Then
            If u Is Nothing Then
                Set u = r
            Else
                Set u = Application.Union(u, r)
            End If
        End If
    Next i
    On Error GoTo 0
    dict [key] = u
    Set EntradasUnion_Cache = u
End Function

Private Sub Meta_GuardarHuella(ByVal ws As Worksheet)
    ' Ligero: preparado para persistir hash si se necesita
End Sub

Public Sub ResetEntorno()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    g_BloquearEventos = False
    MsgBox "Entorno reactivado."
End Sub

Public Sub Test_EstadoCaso_Resolucion()
    Dim ws As Worksheet, r As Range
    Set ws = ActiveSheet
    Set r = ObtenerCeldaEstado(ws)
    If r Is Nothing Then
        MsgBox "No se pudo resolver la celda de estado." & vbCrLf & _
               "Revisar CELDA_ESTADO_POR_DEFECTO y nombre local 'EstadoCaso'."
    Else
        r.Interior.Color = RGB(255, 235, 59)     ' amarillo temporal
        MsgBox "EstadoCaso -> " & r.Address(0, 0) & " en hoja: " & ws.Name
    End If
End Sub

Public Sub DiagnosticarEstadoCaso()
    Dim ws As Worksheet, r As Range, nm As Name, info As String
    Set ws = ActiveSheet

    On Error Resume Next
    Set r = Nothing
    Set r = ObtenerCeldaEstado(ws)
    On Error GoTo 0

    info = "Hoja activa: " & ws.Name & vbCrLf
    info = info & "Protegida: " & ws.ProtectContents & vbCrLf

    ' Existe un nombre 'EstadoCaso' local y/o global?
    Dim tieneLocal As Boolean, tieneGlobal As Boolean, refLocal As String, refGlobal As String
    For Each nm In ThisWorkbook.Names
        If LCase$(nm.Name) = LCase$(ws.Name & "!EstadoCaso") Then
            tieneLocal = True: refLocal = nm.RefersTo
        ElseIf LCase$(nm.Name) = "estadocaso" Then
            tieneGlobal = True: refGlobal = nm.RefersTo
        End If
    Next nm

    info = info & "Nombre local EstadoCaso: " & IIf(tieneLocal, refLocal, "NO") & vbCrLf
    info = info & "Nombre global EstadoCaso: " & IIf(tieneGlobal, refGlobal, "NO") & vbCrLf

    If r Is Nothing Then
        info = info & "ObtenerCeldaEstado -> Nothing" & vbCrLf
    Else
        info = info & "ObtenerCeldaEstado -> " & r.Address(0, 0) & vbCrLf
        info = info & "CF en esa celda: " & r.FormatConditions.Count & vbCrLf
        info = info & "Color actual (Interior.Color): " & r.Interior.Color & vbCrLf
    End If

    MsgBox info, vbInformation, "Diagnostico EstadoCaso"
End Sub

Public Sub CrearNombreLocal_EstadoCaso_G6()
    Dim ws As Worksheet: Set ws = ActiveSheet
    On Error Resume Next
    ws.Names("EstadoCaso").Delete
    On Error GoTo 0
    ws.Names.add Name:="EstadoCaso", RefersTo:=ws.Range(ADDR_G6)
    MsgBox "Nombre local 'EstadoCaso' creado en " & ws.Name & " -> G6"
End Sub

Public Sub ProbarPintadoEstado()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Range: Set r = ObtenerCeldaEstado(ws)
    If r Is Nothing Then
        MsgBox "Sigue sin resolverse la celda de estado en " & ws.Name
        Exit Sub
    End If
    EstadoCaso_MarcarCalculo ws                  ' deberia pintar VERDE G6
    Application.Wait Now + TimeValue("0:00:01")
    EstadoCaso_MarcarModificado ws               ' deberia cambiar a ROJO G6
End Sub

' --- Flag por hoja: ya se ha pulsado Calcular alguna vez? ---
Public Sub EstadoCaso_SetInicializado(ByVal ws As Worksheet, ByVal valor As Boolean)
    On Error Resume Next
    ws.Names("EstadoCaso_Inicializado").Delete
    On Error GoTo 0

    ' Guardamos un nombre local que evalua a TRUE/FALSE
    ws.Names.add Name:="EstadoCaso_Inicializado", RefersTo:="=" & UCase$(CStr(valor))
End Sub

Public Function EstadoCaso_EstaInicializado(ByVal ws As Worksheet) As Boolean
    Dim nm As Name
    On Error Resume Next
    Set nm = ws.Names("EstadoCaso_Inicializado")
    If Not nm Is Nothing Then
        EstadoCaso_EstaInicializado = CBool(Evaluate(nm.RefersTo))
    Else
        EstadoCaso_EstaInicializado = False
    End If
End Function

Public Sub Fix_Estado_G6()
    Dim ws As Worksheet: Set ws = ActiveSheet

    ' 1) Borrar nombres en conflicto
    On Error Resume Next
    ThisWorkbook.Names("EstadoCaso").Delete
    ws.Names("EstadoCaso").Delete
    On Error GoTo 0

    ' 2) Crear nombre LOCAL apuntando a G6
    ws.Names.add Name:="EstadoCaso", RefersTo:=ws.Range(ADDR_G6)

    ' 3) Pintar para comprobar
    EstadoCaso_MarcarCalculo ws
    MsgBox "EstadoCaso -> " & ObtenerCeldaEstado(ws).Address(0, 0) & " (debe ser G6)."
End Sub

Public Sub LimpiarEntradas(Optional ByVal ws As Worksheet)
    Dim celdas As Variant, i As Long
    Dim rng As Range
    Dim tieneVal As Boolean, tipoVal As Long
    Dim f As String, ref As String
    Dim vals() As String, arrTmp() As String
    Dim r As Range, c As Range
    Dim k As Long, j As Long
    Dim esBooleana As Boolean
    Dim vChk As String, falsoPreferido As Variant
    Dim sep As String
    
    If ws Is Nothing Then Set ws = ActiveSheet
    
    celdas = Array( _
        "B2", "D2", "F3", "B3", "D3", "B7", "B8", "D8", "B9", "G3", "G5", _
        "B10", "B11", "B12", "B14", "B16", "B17", "B18", "B19", "B20", "B21", "B22", _
        "F7", "F8", "F9", "F10", "F11", "F12", "F13", "F14", "F15", "F16", "F17", _
        "B24", "F24", "F25", "F26", "B28", "B29", "B30", "B31", "B32", "F28", "F29", "F30", _
        "F31", "F32", "B46", "B47", "F46", "B49", "A51", "B51", "C51", "D51", "E51", "F51", _
        "A53", "B53", "C53", "D53", "E53", "F53", "A55", "B55", "C55", "D55", "E55", "F55", "D29", "D28")
    
    For i = LBound(celdas) To UBound(celdas)
        Set rng = ws.Range(celdas(i))
        
        '--- Detectar validación y tipo ---
        On Error Resume Next
        tipoVal = rng.Validation.Type
        If Err.Number <> 0 Then
            Err.Clear
            tieneVal = False
        Else
            tieneVal = True
        End If
        On Error GoTo 0
        
        If tieneVal And tipoVal = xlValidateList Then
            '--- Obtener valores de la lista de validación (inline o rango) ---
            ReDim vals(-1 To -1)
            On Error Resume Next
            f = rng.Validation.Formula1
            On Error GoTo 0
            
            If Len(f) > 0 Then
                If Left$(f, 1) = "=" Then
                    ref = Mid$(f, 2)
                    Set r = Nothing
                    On Error Resume Next
                    Set r = rng.Parent.Range(ref)    ' si es un rango/nombre local
                    On Error GoTo 0
                    If r Is Nothing Then
                        On Error Resume Next
                        Set r = rng.Parent.Evaluate(ref) ' si es un nombre definido
                        On Error GoTo 0
                    End If
                    If Not r Is Nothing Then
                        ReDim vals(0 To r.Count - 1)
                        k = 0
                        For Each c In r.Cells
                            vals(k) = CStr(c.Value)
                            k = k + 1
                        Next
                    Else
                        ' No se pudo resolver: intentar como inline sin "="
                        sep = IIf(InStr(1, ref, ";") > 0, ";", IIf(InStr(1, ref, ",") > 0, ",", ""))
                        If sep <> "" Then
                            arrTmp = Split(ref, sep)
                            ReDim vals(0 To UBound(arrTmp))
                            For k = LBound(arrTmp) To UBound(arrTmp)
                                vals(k) = arrTmp(k)
                            Next
                        End If
                    End If
                Else
                    ' Inline "A;B;C" o "A,B,C"
                    sep = IIf(InStr(1, f, ";") > 0, ";", IIf(InStr(1, f, ",") > 0, ",", ""))
                    If sep <> "" Then
                        arrTmp = Split(f, sep)
                        ReDim vals(0 To UBound(arrTmp))
                        For k = LBound(arrTmp) To UBound(arrTmp)
                            vals(k) = arrTmp(k)
                        Next
                    End If
                End If
            End If
            
            '--- ¿La lista es booleana? (verdadero/falso/true/false/sí/si/no) ---
            esBooleana = False
            If (Not Not vals) <> 0 Then ' array dimensionado
                For j = LBound(vals) To UBound(vals)
                    vChk = UCase$(Trim$(vals(j)))
                    If vChk = "VERDADERO" Or vChk = "FALSO" Or vChk = "TRUE" Or vChk = "FALSE" _
                       Or vChk = "SÍ" Or vChk = "SI" Or vChk = "NO" Then
                        esBooleana = True
                        Exit For
                    End If
                Next j
            End If
            
            If esBooleana Then
                '--- Elegir el "falso" preferido: FALSO > NO > FALSE > (booleano False) ---
                falsoPreferido = Empty
                If (Not Not vals) <> 0 Then
                    For j = LBound(vals) To UBound(vals)
                        If UCase$(Trim$(vals(j))) = "FALSO" Then falsoPreferido = "FALSO": Exit For
                    Next j
                    If IsEmpty(falsoPreferido) Then
                        For j = LBound(vals) To UBound(vals)
                            If UCase$(Trim$(vals(j))) = "NO" Then falsoPreferido = "NO": Exit For
                        Next j
                    End If
                    If IsEmpty(falsoPreferido) Then
                        For j = LBound(vals) To UBound(vals)
                            If UCase$(Trim$(vals(j))) = "FALSE" Then falsoPreferido = "FALSE": Exit For
                        Next j
                    End If
                End If
                If IsEmpty(falsoPreferido) Then
                    rng.Value = False
                Else
                    rng.Value = falsoPreferido
                End If
            Else
                ' Tiene otra validación de datos (no booleana): dejar vacía sin tocar formato/validación
                rng.ClearContents
            End If
        
        Else
            ' Sin validación o validación distinta de lista: dejar vacía
            rng.ClearContents
        End If
    Next i
    
End Sub

Private Function EsHojaAnexo(ByVal nombreHoja As String) As Boolean
    ' Ajusta el patron si tienes otros prefijos
    EsHojaAnexo = (LCase$(Left$(Trim$(nombreHoja), 4)) = "anx.")
End Function

Public Sub DesprotegerTodoElArchivo()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        On Error GoTo 0
    Next ws
    MsgBox "Todas las hojas han sido desprotegidas.", vbInformation
End Sub

Public Sub SOS_RearmarExcel()
    On Error Resume Next
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .CutCopyMode = False
        .StatusBar = False
        .AskToUpdateLinks = False
    End With
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.EnableCalculation = True
    Next ws
    
    ' Desproteger tambin la estructura del libro
    
    On Error GoTo 0
    MsgBox "He reactivado eventos, calculo automatico y desprotegido todo el libro/hojas.", vbInformation
End Sub

' ==== Diagnstico rpido del entorno ====
Public Sub DiagnosticarEstado()
    Dim msg As String, wb As Workbook, ws As Worksheet
    Set wb = ThisWorkbook
    msg = "== DIAGNSTICO ==" & vbCrLf
    msg = msg & "EnableEvents: " & CStr(Application.EnableEvents) & vbCrLf
    msg = msg & "Calculation: " & IIf(Application.Calculation = xlCalculationAutomatic, "Automtico", "Manual") & vbCrLf
    msg = msg & "Libro protegido (estructura): " & CStr(wb.ProtectStructure) & vbCrLf

    Dim protHojas As Long, sinMacro As Long, totalBtns As Long
    For Each ws In wb.Worksheets
        If ws.ProtectContents Then protHojas = protHojas + 1
        ' Controles de formularios con macro OnAction
        Dim shp As Shape
        For Each shp In ws.Shapes
            If shp.Type = msoFormControl Then
                totalBtns = totalBtns + 1
                If shp.OnAction = "" Then sinMacro = sinMacro + 1
            End If
        Next shp
    Next ws

    msg = msg & "Hojas protegidas (contenido): " & protHojas & vbCrLf
    msg = msg & "Botones (Form Controls): " & totalBtns & " | sin macro: " & sinMacro & vbCrLf
    msg = msg & vbCrLf & "Si aparece un error al compilar, ve a VBA: Depurar > Compilar VBAProject."
    MsgBox msg, vbInformation, "Estado"
End Sub

Public Sub DiagnosticarCeldaInterfazF3()
    Dim ws As Worksheet, r As Range, msg As String
    Dim hasDV As Boolean, isMerged As Boolean, locked As Boolean
    Dim protContenido As Boolean, protEstructura As Boolean
    Dim shp As Shape, solapes As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Interfaz")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "No encuentro la hoja 'Interfaz'.", vbCritical: Exit Sub
    End If

    Set r = ws.Range(celdaSuperficie)
    protContenido = ws.ProtectContents
    protEstructura = ThisWorkbook.ProtectStructure
    locked = r.locked
    isMerged = r.MergeCells

    ' Hay validacin?
    On Error Resume Next
    hasDV = Not r.Validation Is Nothing
    If Err.Number <> 0 Then hasDV = False
    Err.Clear
    On Error GoTo 0

    ' Hay formas encima?
    For Each shp In ws.Shapes
        On Error Resume Next
        If Not Intersect(r, ws.Range(shp.TopLeftCell.Address, shp.BottomRightCell.Address)) Is Nothing Then
            solapes = solapes + 1
        End If
        On Error GoTo 0
    Next shp

    msg = "DIAGNSTICO F3 (Interfaz)" & vbCrLf & String(28, "-") & vbCrLf
    msg = msg & "Libro Protegido (estructura): " & protEstructura & vbCrLf
    msg = msg & "Hoja Protegida (contenido):  " & protContenido & vbCrLf
    msg = msg & "F3 Locked:                   " & locked & vbCrLf
    msg = msg & "F3 MergeCells:               " & isMerged & vbCrLf
    msg = msg & "F3 con Validacin:           " & hasDV & vbCrLf
    msg = msg & "Formas solapando F3:         " & solapes & vbCrLf
    MsgBox msg, vbInformation, "Diagnstico F3"
End Sub

' =========================
' DEFAULTS PARA INTERFAZ
' =========================


Private Function DefaultsInterfaz() As Variant
    ' {direccion, valor_por_defecto}; si valor_por_defecto = "" -> toma primer item del desplegable
    DefaultsInterfaz = Array( _
                       Array(celdaSituacion, "Zonas ocupables, en general"), _
                       Array("B9", ""), _
                       Array("B10", ""), _
                       Array("B11", ""), _
                       Array("B12", ""), _
                       Array("B16", ""), _
                       Array("B17", ""), _
                       Array("B18", ""), _
                       Array("B19", ""), _
                       Array("B20", ""), _
                       Array("B21", ""), _
                       Array("B22", ""), _
                       Array("B24", ""), _
                       Array("B28", ""), _
                       Array("B29", ""), _
                       Array("B30", ""), _
                       Array("B31", ""), _
                       Array("B32", ""), _
                       Array("B46", ""), _
                       Array("B47", "") _
                       )
End Function

Public Sub EstablecerPredeterminadosInterfaz(ByVal ws As Worksheet)
    Dim defs As Variant, i As Long
    defs = DefaultsInterfaz()
    On Error Resume Next
    For i = LBound(defs) To UBound(defs)
        SetDefaultFromValidation ws, CStr(defs(i)(0)), CStr(defs(i)(1))
    Next i
    On Error GoTo 0
End Sub

Private Sub SetDefaultFromValidation(ByVal ws As Worksheet, ByVal addr As String, ByVal desiredDefault As String)
    Dim r As Range, f As String, items As Variant, rng As Range

    Set r = ws.Range(addr)

    ' Si defines un valor fijo, aplicalo
    If Len(desiredDefault) > 0 Then
        r.Value = desiredDefault
        Exit Sub
    End If

    ' Si no defines valor, toma el primero del origen de la validacion (lista)
    On Error Resume Next
    If r.Validation.Type = xlValidateList Then
        f = r.Validation.Formula1
    Else
        f = ""
    End If
    On Error GoTo 0

    If Len(f) = 0 Then Exit Sub

    If Left$(f, 1) = "=" Then
        ' Rango o nombre
        On Error Resume Next
        Set rng = ws.Evaluate(f)
        If rng Is Nothing Then
            Set rng = ws.Parent.Evaluate(f)      ' prueba a nivel libro
        End If
        If rng Is Nothing Then
            Set rng = ws.Range(Replace$(f, "=", ""))
        End If
        On Error GoTo 0
        If Not rng Is Nothing Then r.Value = rng.Cells(1, 1).Value
    Else
        ' Lista literal "A,B,C"
        items = Split(f, ",")
        If UBound(items) >= 0 Then r.Value = Trim$(items(0))
    End If
End Sub

Public Sub Interfaz_DesprotegerYDesbloquear()
    Dim ws As Worksheet, arr As Variant, i As Long
    Dim prevEvents As Boolean

    ' Asegurar eventos ACTIVOS
    prevEvents = Application.EnableEvents
    Application.EnableEvents = True

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Interfaz")
    On Error GoTo 0
    If ws Is Nothing Then
        Application.EnableEvents = True
        MsgBox "No encuentro la hoja 'Interfaz'.", vbCritical
        Exit Sub
    End If

    ' 1) Quitar protecciones (sin contrasea)
    On Error Resume Next
    ThisWorkbook.Unprotect
    ws.Unprotect
    ws.EnableSelection = xlNoRestrictions
    On Error GoTo 0

    If ws.ProtectContents Then
        Application.EnableEvents = True
        MsgBox "La hoja 'Interfaz' sigue protegida. Abre Revisar > Desproteger hoja e intenta con contrasea en blanco. " & _
               "Hazlo una sola vez; el cdigo ya no la volver a proteger.", vbExclamation
        Exit Sub
    End If

    ' 2) Desbloquear F3 y las entradas de la interfaz
    ws.Range(celdaSuperficie).locked = False
    On Error Resume Next
    arr = RangosEntradaInterfaz()
    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            ws.Range(CStr(arr(i))).locked = False
        Next i
    End If
    On Error GoTo 0

    ' Garantizar eventos activados antes de salir
    Application.EnableEvents = True

    MsgBox "Interfaz desprotegida y celdas de entrada (incluida F3) desbloqueadas.", vbInformation
End Sub

Public Sub RepararBotonesInterfaz()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim ole As OLEObject
    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets("Interfaz")
    If ws Is Nothing Then
        MsgBox "No existe la hoja 'Interfaz'.", vbCritical
        Exit Sub
    End If

    ' Asegurar interactividad global
    Application.EnableEvents = True
    Application.Interactive = True
    Application.ScreenUpdating = True

    ' Quitar protecciones y limites de seleccion
    ThisWorkbook.Unprotect
    ws.Unprotect
    ws.EnableSelection = xlNoRestrictions
    ws.ScrollArea = ""                           ' <- MUY CLAVE si antes estaba restringida

    ' Desbloquear y asegurar visibilidad de todos los botones/objetos
    For Each shp In ws.Shapes
        shp.locked = False
        shp.Visible = msoTrue
    Next shp

    ' Para ActiveX: habilitar y desbloquear
    For Each ole In ws.OLEObjects
        ole.locked = False
        ole.Enabled = True
        If Not ole.Object Is Nothing Then
            On Error Resume Next
            ole.Object.Enabled = True
            On Error GoTo 0
        End If
        ole.Visible = True
    Next ole

    ' Nota: el Modo DisÃ±o de ActiveX es global del libro; verifica que esta DESACTIVADO en Desarrollador.
    MsgBox "Interfaz reseteada: protecciones fuera, ScrollArea limpia, objetos habilitados.", vbInformation
End Sub

Public Sub SOS_RestaurarEstado_Interfaz()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Interfaz")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "No encuentro la hoja 'Interfaz'.", vbCritical
        Exit Sub
    End If

    ' --- Estado global de Excel ---
    Application.EnableEvents = True              ' Reactiva eventos (SelectionChange, etc.)
    Application.Calculation = xlCalculationAutomatic '
    Application.Interactive = True
    Application.ScreenUpdating = True

    ' --- Estado de la hoja Interfaz ---
    On Error Resume Next
    ThisWorkbook.Unprotect
    ws.Unprotect
    ws.EnableSelection = xlNoRestrictions
    ws.ScrollArea = ""
    ws.EnableCalculation = True                  ' CLAVE: si estaba en False, no recalcula nada
    On Error GoTo 0

    ' Recalcular a fondo
    Application.CalculateFullRebuild             ' Fuerza dependencia completa
    ws.Calculate

    MsgBox "Restaurado: eventos ON, calculo AUTOMATICO, hoja habilitada y recalculo forzado.", vbInformation
End Sub

' --- Ejecuta: Diagnostico_Interfaz_InteractiveCells ---

Public Sub Diagnostico_Interfaz_InteractiveCells()
    Dim ws As Worksheet, repWs As Worksheet
    Dim row As Long, ok As Boolean, msg As String
    Dim arr As Variant, arr2 As Variant
    Dim i As Long, addr As String
    Dim nm As Name, r As Range
    Dim fails As Long, warns As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Interfaz")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "No encuentro la hoja 'Interfaz'.", vbCritical
        Exit Sub
    End If

    ' Crear (o limpiar) hoja de informe
    On Error Resume Next
    Set repWs = ThisWorkbook.Worksheets("Diag_Interfaz")
    On Error GoTo 0
    If repWs Is Nothing Then
        Set repWs = ThisWorkbook.Worksheets.add(After:=ws)
        repWs.Name = "Diag_Interfaz"
    Else
        repWs.Cells.Clear
    End If

    ' Cabecera
    row = 1
    repWs.Range("A" & row).Value = "Chequeo"
    repWs.Range("B" & row).Value = "Estado"
    repWs.Range("C" & row).Value = "Detalle / Sugerencia"
    repWs.Rows(row).Font.bold = True
    row = row + 1

    ' --- 1) Estado global de Excel ---
    AddCheck repWs, row, "Application.EnableEvents", IIf(Application.EnableEvents, "OK", "FAIL"), _
             IIf(Application.EnableEvents, "Eventos activos", "Eventos desactivados -> reactivar")
    If Not Application.EnableEvents Then fails = fails + 1

    AddCheck repWs, row, "Application.Calculation", IIf(Application.Calculation = xlCalculationAutomatic, "OK", "FAIL"), _
             "Modo actual: " & ModeToText(Application.Calculation) & " -> se recomienda Automatico"
    If Not (Application.Calculation = xlCalculationAutomatic) Then warns = warns + 1

    AddCheck repWs, row, "Application.Interactive", IIf(Application.Interactive, "OK", "WARN"), _
             IIf(Application.Interactive, "Interactivo", "No interactivo -> puede bloquear clicks")
    If Not Application.Interactive Then warns = warns + 1

    AddCheck repWs, row, "Application.ScreenUpdating", IIf(Application.ScreenUpdating, "OK", "WARN"), _
             IIf(Application.ScreenUpdating, "ScreenUpdating ON", "OFF -> solo impacto visual")

    ' --- 2) Estado del libro/hoja ---
    AddCheck repWs, row, "ThisWorkbook.ProtectStructure", IIf(ThisWorkbook.ProtectStructure, "WARN", "OK"), _
             IIf(ThisWorkbook.ProtectStructure, "Estructura protegida (no afecta clicks, pero ojo duplicados/renombres)", "Estructura sin proteger")
    If ThisWorkbook.ProtectStructure Then warns = warns + 1

    AddCheck repWs, row, "Interfaz.ProtectContents", IIf(ws.ProtectContents, "FAIL", "OK"), _
             IIf(ws.ProtectContents, "Contenido protegido -> puede impedir seleccion/edicion", "Contenido no protegido")
    If ws.ProtectContents Then fails = fails + 1

    AddCheck repWs, row, "Interfaz.EnableSelection", _
             IIf(ws.EnableSelection = xlNoRestrictions, "OK", "FAIL"), _
             "Valor: " & EnableSelToText(ws.EnableSelection) & " -> debe ser 'Sin restricciones'"
    If Not (ws.EnableSelection = xlNoRestrictions) Then fails = fails + 1

    AddCheck repWs, row, "Interfaz.ScrollArea", IIf(Len(ws.ScrollArea) = 0, "OK", "FAIL"), _
             IIf(Len(ws.ScrollArea) = 0, "Sin limites de Scroll", "ScrollArea=" & ws.ScrollArea & " -> limpia con ws.ScrollArea = """"")
    If Len(ws.ScrollArea) <> 0 Then fails = fails + 1

    AddCheck repWs, row, "Interfaz.EnableCalculation", IIf(ws.EnableCalculation, "OK", "FAIL"), _
             IIf(ws.EnableCalculation, "Calculo de hoja habilitado", "Calculo deshabilitado -> no actualiza formulas")
    If Not ws.EnableCalculation Then fails = fails + 1

    ' --- 3) Validaciones y celdas "boton" (si existen funciones auxiliares) ---
    ' Intentamos obtener los rangos de entrada conocidos si existen en tu proyecto
    If TryRunFunc("RangosEntradaInterfaz", arr) Then
        AddCheck repWs, row, "RangosEntradaInterfaz()", "OK", "Encontrada funcion; se inspeccionan rangos"
        Call CheckRangesArray(ws, arr, repWs, row, fails, warns)
    Else
        AddCheck repWs, row, "RangosEntradaInterfaz()", "WARN", "Funcion no encontrada; se omite este chequeo"
        warns = warns + 1
    End If

    If TryRunFunc("EntradasVigiladas", arr2) Then
        AddCheck repWs, row, "EntradasVigiladas()", "OK", "Encontrada funcion; se inspeccionan rangos"
        Call CheckRangesArray(ws, arr2, repWs, row, fails, warns)
    Else
        AddCheck repWs, row, "EntradasVigiladas()", "WARN", "Funcion no encontrada; se omite este chequeo"
        warns = warns + 1
    End If

    ' --- 4) Nombres que apuntan a Interfaz (posibles validaciones o listas) ---
    For Each nm In ThisWorkbook.Names
        On Error Resume Next
        Set r = Nothing
        Set r = nm.RefersToRange
        On Error GoTo 0
        If Not r Is Nothing Then
            If r.Worksheet.Name = ws.Name Then
                AddCheck repWs, row, "Nombre: " & nm.Name, "INFO", "Refiere a " & r.Address(External:=False)
            End If
        End If
    Next nm

    ' --- 5) Resumen ---
    msg = "Diagnostico 'Interfaz' completado." & vbCrLf & _
          "FAILS: " & fails & " | WARNS: " & warns & vbCrLf & _
          "Ver hoja 'Diag_Interfaz' para detalle."
    MsgBox msg, IIf(fails > 0, vbExclamation, vbInformation)

    ' Formato basico del reporte
    With repWs
        .Columns("A:C").EntireColumn.AutoFit
        .Rows(1).Interior.ColorIndex = 36
    End With
End Sub

' --- Helpers ---

Private Sub AddCheck(ByVal ws As Worksheet, ByRef row As Long, ByVal titulo As String, ByVal estado As String, ByVal detalle As String)
    ws.Cells(row, 1).Value = titulo
    ws.Cells(row, 2).Value = estado
    ws.Cells(row, 3).Value = detalle
    If UCase$(estado) = "FAIL" Then ws.Cells(row, 2).Interior.Color = RGB(255, 200, 200)
    If UCase$(estado) = "WARN" Then ws.Cells(row, 2).Interior.Color = RGB(255, 245, 200)
    row = row + 1
End Sub

Private Function ModeToText(ByVal m As XlCalculation) As String
    Select Case m
    Case xlCalculationAutomatic: ModeToText = "Automatico"
    Case xlCalculationManual:    ModeToText = "Manual"
    Case xlCalculationSemiautomatic: ModeToText = "Semiautomatico"
    Case Else: ModeToText = CStr(m)
    End Select
End Function

Private Function EnableSelToText(ByVal v As XlEnableSelection) As String
    Select Case v
    Case xlNoRestrictions: EnableSelToText = "Sin restricciones"
    Case xlUnlockedCells:  EnableSelToText = "Solo celdas desbloqueadas"
    Case xlNoSelection:    EnableSelToText = "Sin seleccion"
    Case Else: EnableSelToText = CStr(v)
    End Select
End Function

Private Function TryRunFunc(ByVal funcName As String, ByRef ret As Variant) As Boolean
    On Error Resume Next
    ret = Application.run(funcName)
    TryRunFunc = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub CheckRangesArray(ByVal ws As Worksheet, ByVal arr As Variant, ByVal repWs As Worksheet, ByRef row As Long, ByRef fails As Long, ByRef warns As Long)
    Dim i As Long, addr As String, r As Range, hasVal As Boolean, s As String

    If Not IsArray(arr) Then Exit Sub

    For i = LBound(arr) To UBound(arr)
        addr = CStr(arr(i))
        On Error Resume Next
        Set r = ws.Range(addr)
        On Error GoTo 0

        If r Is Nothing Then
            AddCheck repWs, row, "Rango listado no existe", "FAIL", "Direccion: " & addr
            fails = fails + 1
        Else
            ' Estado de bloqueo / validacion / texto vs formula
            s = ""
            s = s & "Locked=" & r.locked & "; "
            s = s & "HasValidation=" & HasValidation(r) & "; "
            s = s & "Formula=" & (Left$(r.Formula, 1) = "=") & "; "
            s = s & "NumberFormat=" & r.NumberFormat

            ' Si la hoja esta protegida y EnableSelection <> xlNoRestrictions, un Locked=True puede impedir clicks
            If r.locked And ws.ProtectContents Then
                AddCheck repWs, row, addr, "FAIL", "Celda bloqueada con hoja protegida -> impide interaccion"
                fails = fails + 1
            Else
                AddCheck repWs, row, addr, "OK", s
            End If
        End If

        Set r = Nothing
    Next i
End Sub

Private Function HasValidation(ByVal r As Range) As Boolean
    On Error Resume Next
    HasValidation = (r.Validation.Type <> xlValidateInputOnly)
    If Err.Number <> 0 Then
        HasValidation = False
        Err.Clear
    End If
    On Error GoTo 0
End Function


Private Function FindHeaderIndex(ByVal arr As Variant, ByVal headerName As String) As Long
    Dim c&, maxC&
    On Error GoTo fin
    maxC = UBound(arr, 2)
    For c = 1 To maxC
        If StrComp(CStr(arr(1, c)), headerName, vbTextCompare) = 0 Then
            FindHeaderIndex = c
            Exit Function
        End If
    Next c
fin:
End Function
