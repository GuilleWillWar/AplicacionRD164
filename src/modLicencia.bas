Attribute VB_Name = "modLicencia"
Option Explicit

' Devuelve True si el equipo/usuario actual está autorizado según el fichero.
Public Function IsAuthorized(ByVal licensePath As String) As Boolean
    Dim fso As Object, ts As Object
    Dim linea As String
    Dim equipo As String, usuario As String

    equipo = UCase$(Trim$(Environ$("COMPUTERNAME")))
    usuario = UCase$(Trim$(Environ$("USERNAME")))

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Comprobación de existencia
    If Not fso.FileExists(licensePath) Then
        Err.Raise vbObjectError + 16401, "IsAuthorized", _
                  "No se encuentra el fichero de licencias: " & licensePath
    End If

    Set ts = fso.OpenTextFile(licensePath, 1, False) ' ForReading

    On Error GoTo salida

    Do While Not ts.AtEndOfStream
        linea = ts.ReadLine

        ' Normaliza / ignora comentarios y vacías
        linea = Trim$(linea)
        If Len(linea) = 0 Then GoTo siguiente
        If Left$(linea, 1) = "#" Then GoTo siguiente

        ' Comparación case-insensitive
        If UCase$(Left$(linea, 5)) = "USER:" Then
            ' Validación por usuario de Windows
            If UCase$(Trim$(Mid$(linea, 6))) = usuario Then
                IsAuthorized = True
                Exit Do
            End If
        Else
            ' Validación por nombre de equipo
            If UCase$(linea) = equipo Then
                IsAuthorized = True
                Exit Do
            End If
        End If

siguiente:
    Loop

salida:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Function



