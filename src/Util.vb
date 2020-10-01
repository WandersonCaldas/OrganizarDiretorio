Imports System.IO

Public Class Util

    Public Shared Function CompletaZero(ByVal txt_texto As String, ByVal cod_quantidade As Integer) As String
        While Len(txt_texto) < cod_quantidade
            txt_texto = "0" & txt_texto
        End While
        CompletaZero = txt_texto

    End Function


    Public Shared Function CriarDiretorio(ByVal txt_caminho As String, ByVal fi As FileInfo, ByVal caminho_destino As String) As String
        Dim retorno As String = ""
        Dim dt As DateTime = fi.LastWriteTime

        retorno = caminho_destino & "\" & Year(dt) & "\" & CompletaZero(Month(dt), 2) & "\"

        If Not System.IO.Directory.Exists(retorno) Then
            System.IO.Directory.CreateDirectory(retorno)
        End If

        Return retorno
    End Function
End Class
