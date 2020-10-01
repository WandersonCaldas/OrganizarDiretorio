Imports System.Configuration
Imports System.Globalization
Imports System.IO
Imports System.Threading

Module Module1

    Sub Main()
        Dim culture As New CultureInfo("pt-BR")
        Thread.CurrentThread.CurrentCulture = culture
        Thread.CurrentThread.CurrentUICulture = culture

        Dim _caminho_origem As String = ConfigurationManager.AppSettings.Get("caminho_origem")
        Dim _caminho_destino As String = ConfigurationManager.AppSettings.Get("caminho_destino")
        Dim _extensao As String = ConfigurationManager.AppSettings.Get("extensao")

        Dim erro As String = ""
        Dim listaExtensao As New List(Of Result)

        If Not String.IsNullOrWhiteSpace(_extensao) Then
            Dim a_extensao = _extensao.Split(",")

            For i = LBound(a_extensao) To UBound(a_extensao)
                Dim result As New Result
                result.txt_extensao = a_extensao(i)

                If listaExtensao.Count = 0 Then
                    listaExtensao.Add(result)
                ElseIf listaExtensao.Where(Function(x) x.txt_extensao = a_extensao(i)).Count = 0 Then
                    listaExtensao.Add(result)
                End If
            Next
        End If

        Console.WriteLine("ROTINA INICIADA")

        If String.IsNullOrWhiteSpace(erro) AndAlso String.IsNullOrWhiteSpace(_caminho_origem) Then erro = "ERRO: CAMINHO DE ORIGEM DEVE SER INFORMADO."
        If String.IsNullOrWhiteSpace(erro) AndAlso String.IsNullOrWhiteSpace(_caminho_destino) Then erro = "ERRO: CAMINHO DE DESTINO DEVE SER INFORMADO."

        If Not String.IsNullOrWhiteSpace(erro) Then
            Console.WriteLine(erro)
            Console.ReadKey()
        Else
            Dim raizDir As String = ""

            'For Each raizDir In Directory.GetDirectories(_caminho_origem)
            For Each nome In Directory.GetFiles(_caminho_origem)
                Try
                    Dim extensao As String = Path.GetExtension(nome)
                    Dim nomeArquivo As String = Path.GetFileNameWithoutExtension(nome)
                    Dim fi As FileInfo = New FileInfo(nome)

                    If listaExtensao.Count > 0 Then
                        If listaExtensao.Where(Function(x) x.txt_extensao.ToString.Trim.ToLower = extensao.ToString.ToLower).Count > 0 Then
                            Dim retorno As String = Util.CriarDiretorio(nomeArquivo & extensao, fi, _caminho_destino)

                            Dim gravar As String = retorno & nomeArquivo & extensao
                            gravar = gravar.Replace("\\", "\")

                            File.Copy(nome, gravar, True)
                            Console.WriteLine("SUCESSO -> ORIGEM: " & nome & " DESTINO: " & gravar)
                        End If
                    Else
                        Dim retorno As String = Util.CriarDiretorio(nomeArquivo & extensao, fi, _caminho_destino)

                        Dim gravar As String = retorno & nomeArquivo & extensao
                        gravar = gravar.Replace("\\", "\")

                        File.Copy(nome, gravar, True)
                        Console.WriteLine("SUCESSO -> ORIGEM: " & nome & " DESTINO: " & gravar)
                    End If
                Catch ex As Exception
                    Console.WriteLine("ERRO: " & ex.Source & " - " & ex.Message & " - " & ex.StackTrace)
                End Try

            Next
            'Next

            Console.WriteLine("ROTINA FINALIZADA")
            Console.ReadKey()
        End If
    End Sub

End Module
