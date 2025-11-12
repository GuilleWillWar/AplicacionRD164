Attribute VB_Name = "modVS"
Option Explicit
'Esta cargando

'============== CONFIG ==============
Private Const DEFAULT_SRC_SUBFOLDER As String = "src"
'====================================

'---- Ruta carpeta src
Public Function SrcPath() As String
    Dim p As String
    p = ThisWorkbook.path & "\" & DEFAULT_SRC_SUBFOLDER
    If Right$(p, 1) <> "\" Then p = p & "\"
    SrcPath = p
End Function

'---- Normaliza texto para el editor VBA
Private Function NormalizeVbaText(ByVal s As String) As String
    If Left$(s, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then s = Mid$(s, 4) ' BOM UTF-8
    s = Replace(s, vbCrLf, vbLf): s = Replace(s, vbCr, vbLf): s = Replace(s, vbLf, vbCrLf)
    s = Replace$(s, ChrW$(&HFEFF), "")   ' ZWNBSP
    s = Replace$(s, ChrW$(&H200B), "")   ' ZWSP
    s = Replace$(s, ChrW$(&H200C), "")   ' ZWNJ
    s = Replace$(s, ChrW$(&H200D), "")   ' ZWJ
    s = Replace$(s, ChrW$(&HA0), " ")    ' NBSP -> espacio normal
    NormalizeVbaText = s
End Function

'---- Lee texto de archivo (ANSI/UTF-8)
Private Function ReadAllText(ByVal fullPath As String) As String
    Dim f As Integer, bytes() As Byte
    If Dir$(fullPath) = "" Then Exit Function
    f = FreeFile
    Open fullPath For Binary As #f
    If LOF(f) > 0 Then
        ReDim bytes(1 To LOF(f)) As Byte
        Get #f, , bytes
    End If
    Close #f
    ReadAllText = StrConv(bytes, vbUnicode)
End Function

'---- Texto actual de un componente
Private Function GetCompText(ByVal comp As VBIDE.VBComponent) As String
    Dim cm As VBIDE.CodeModule: Set cm = comp.CodeModule
    If cm.CountOfLines > 0 Then
        GetCompText = cm.lines(1, cm.CountOfLines)
    Else
        GetCompText = vbNullString
    End If
End Function

'---- Busca SOLO módulos de documento por CodeName
Private Function FindDocumentByCodeName(vbProj As VBIDE.VBProject, ByVal codeName As String) As VBIDE.VBComponent
    Dim c As VBIDE.VBComponent
    For Each c In vbProj.VBComponents
        If c.Type = vbext_ct_Document Then
            If StrComp(c.Name, codeName, vbBinaryCompare) = 0 Then
                Set FindDocumentByCodeName = c
                Exit Function
            End If
        End If
    Next c
End Function

'================= EXPORTA TODO =================
Public Sub ExportAllVBA()
Attribute ExportAllVBA.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim vbProj As VBIDE.VBProject: Set vbProj = ThisWorkbook.VBProject
    Dim path As String: path = SrcPath()
    Dim comp As VBIDE.VBComponent
    Dim outFile As String
    Dim nBas As Long, nCls As Long, nFrm As Long, nDoc As Long

    If Dir$(path, vbDirectory) = "" Then MkDir path

    For Each comp In vbProj.VBComponents
        Select Case comp.Type
            Case vbext_ct_StdModule
                outFile = path & comp.Name & ".bas"
                comp.Export outFile: nBas = nBas + 1

            Case vbext_ct_ClassModule
                outFile = path & comp.Name & ".cls"
                comp.Export outFile: nCls = nCls + 1

            Case vbext_ct_MSForm
                outFile = path & comp.Name & ".frm"
                comp.Export outFile: nFrm = nFrm + 1  ' genera también .frx

            Case vbext_ct_Document
                outFile = path & comp.Name & ".doccls"
                ' --- PARCHE ROBUSTO: guarda UTF-8 con BOM vía ADODB.Stream y fallback puro VBA ---
                SaveTextUTF8 outFile, GetCompText(comp)
                nDoc = nDoc + 1
        End Select
    Next comp

    MsgBox "Exportación a " & path & vbCrLf & _
           ".bas: " & nBas & " | .cls: " & nCls & " | .frm: " & nFrm & " | .doccls: " & nDoc, vbInformation
End Sub
Private Function IsDocumentComponent(ByVal c As VBIDE.VBComponent) As Boolean
    On Error Resume Next
    IsDocumentComponent = (c.Type = vbext_ct_Document)
    If Err.Number <> 0 Then
        ' Si falla acceder a .Type, por seguridad lo tratamos como Document (no borrar)
        IsDocumentComponent = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

'-----IMPORTAR-------

Public Sub ImportAllVBA()
    Dim vbProj As VBIDE.VBProject
    Dim path As String
    Dim file As String, base As String, ext As String
    Dim comp As VBIDE.VBComponent
    Dim cm As VBIDE.CodeModule
    Dim txt As String
    Dim nBas As Long, nCls As Long, nFrm As Long, nDoc As Long
    Dim testCount As Long

    On Error GoTo FailFast

    Set vbProj = ThisWorkbook.VBProject
    path = SrcPath()

    ' Carpeta origen
    If Dir$(path, vbDirectory) = "" Then
        MsgBox "No existe la carpeta: " & path, vbExclamation
        Exit Sub
    End If

    ' Proyecto protegido
    If vbProj.Protection = vbext_pp_locked Then
        MsgBox "El proyecto VBA esta bloqueado. Desprotege antes de importar.", vbExclamation
        Exit Sub
    End If

    ' Permiso al modelo de objetos del VBIDE
    On Error Resume Next
    testCount = vbProj.VBComponents.Count
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        MsgBox "Excel no permite acceder al modelo de objetos de VBA." & vbCrLf & _
               "Activa: Archivo > Opciones > Centro de confianza > Configuracion... > " & _
               "'Confiar en el modelo de objetos de proyectos de VBA'.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' --- .bas/.cls/.frm: reemplazar si existe e importar SIEMPRE
    file = Dir$(path & "*.*")
    Do While Len(file) > 0
        ext = LCase$(Mid$(file, InStrRev(file, ".") + 1))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            base = Left$(file, InStrRev(file, ".") - 1)

            On Error Resume Next
            Set comp = vbProj.VBComponents(base)
            On Error GoTo 0

            If Not comp Is Nothing Then
                If Not IsDocumentComponent(comp) Then
                    vbProj.VBComponents.Remove comp
                Else
                    GoTo NextFile ' seguridad: no tocar documentos aqui
                End If
            End If

            vbProj.VBComponents.Import path & file
            Select Case ext
                Case "bas": nBas = nBas + 1
                Case "cls": nCls = nCls + 1
                Case "frm": nFrm = nFrm + 1
            End Select
        End If
NextFile:
        file = Dir$()
    Loop

    ' --- .doccls: escribir SIEMPRE en el DOCUMENTO cuyo CodeName = nombre del archivo
    file = Dir$(path & "*.doccls")
    Do While Len(file) > 0
        base = Left$(file, InStrRev(file, ".") - 1)

        Set comp = FindDocumentByCodeName(vbProj, base)
        If comp Is Nothing Then
            MsgBox "No se encontro documento con CodeName='" & base & "' para " & file & vbCrLf & _
                   "Abre VBE > selecciona la hoja > Propiedades (F4) > (Name) y renombra el .doccls para que coincida.", _
                   vbExclamation
        Else
            txt = NormalizeVbaText(ReadAllText(path & file))
            Set cm = comp.CodeModule
            If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
            If Len(txt) > 0 Then cm.AddFromString txt
            nDoc = nDoc + 1
        End If
        file = Dir$()
    Loop

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Importacion completada:" & vbCrLf & _
           " .bas: " & nBas & " | .cls: " & nCls & " | .frm: " & nFrm & " | .doccls: " & nDoc, vbInformation
    Exit Sub

FailFast:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "ImportAllVBA"
End Sub
'============= DIAGNÓSTICO: lista CodeName de documentos ============
Public Sub Debug_ListDocumentCodeNames()
    Dim c As VBIDE.VBComponent
    Debug.Print "=== Document CodeNames ==="
    For Each c In ThisWorkbook.VBProject.VBComponents
        If c.Type = vbext_ct_Document Then
            Debug.Print " - "; c.Name
        End If
    Next c
End Sub

'====================== UTILIDADES DE EXPORTACIÓN ====================

' Guarda texto en UTF-8 con BOM. Usa ADODB.Stream; si falla, fallback puro VBA.
Private Sub SaveTextUTF8(ByVal outFile As String, ByVal txt As String)
    On Error GoTo AdoFailed

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2                 ' adTypeText
        .Mode = 3                 ' adModeReadWrite
        .Charset = "utf-8"        ' ¡charset antes de Open!
        .Open
        If LenB(txt) > 0 Then .WriteText txt, 0  ' adWriteChar
        .Position = 0
        .SaveToFile outFile, 2    ' adSaveCreateOverWrite
        .Close
    End With
    Exit Sub

AdoFailed:
    ' Fallback: escribir UTF-8 con BOM sin ADODB
    Dim fh As Integer, bytes() As Byte
    bytes = Utf8EncodeWithBOM(txt)
    fh = FreeFile
    Open outFile For Binary As #fh
    Put #fh, , bytes
    Close #fh
End Sub

' Convierte String (UTF-16 de VBA) a bytes UTF-8 con BOM.
Private Function Utf8EncodeWithBOM(ByVal s As String) As Byte()
    Dim i As Long, cp As Long, low As Long
    Dim b() As Byte, k As Long

    ' Reservar BOM
    ReDim b(0 To 2)
    b(0) = &HEF: b(1) = &HBB: b(2) = &HBF
    k = 3

    For i = 1 To Len(s)
        cp = AscW(Mid$(s, i, 1))

        ' Pares sustitutos ? codepoint de 4 bytes
        If cp >= &HD800 And cp <= &HDBFF Then
            i = i + 1
            low = AscW(Mid$(s, i, 1))
            cp = &H10000 + ((cp - &HD800) * &H400) + (low - &HDC00)
        End If

        If cp < &H80 Then
            ReDim Preserve b(0 To k)
            b(k) = cp
            k = k + 1
        ElseIf cp < &H800 Then
            ReDim Preserve b(0 To k + 1)
            b(k) = &HC0 Or (cp \ &H40)
            b(k + 1) = &H80 Or (cp And &H3F)
            k = k + 2
        ElseIf cp < &H10000 Then
            ReDim Preserve b(0 To k + 2)
            b(k) = &HE0 Or (cp \ &H1000)
            b(k + 1) = &H80 Or ((cp \ &H40) And &H3F)
            b(k + 2) = &H80 Or (cp And &H3F)
            k = k + 3
        Else
            ReDim Preserve b(0 To k + 3)
            b(k) = &HF0 Or (cp \ &H40000)
            b(k + 1) = &H80 Or ((cp \ &H1000) And &H3F)
            b(k + 2) = &H80 Or ((cp \ &H40) And &H3F)
            b(k + 3) = &H80 Or (cp And &H3F)
            k = k + 4
        End If
    Next

    Utf8EncodeWithBOM = b
End Function



