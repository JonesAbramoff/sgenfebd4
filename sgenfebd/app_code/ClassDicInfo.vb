Imports System
Imports System.IO
Imports System.Text
Imports System.Data.Odbc
Imports Microsoft.Win32

Public Class ClassDicInfo

    Public Sub New()
    End Sub

    Public Function Iniciar(ByVal sEmpresa As String, ByVal iFilialEmpresa As Integer, ByVal sADM100INI As String, ByRef sDirXml As String, ByRef sDirXsd As String, ByRef sConexaoDados As String, ByRef iDebug As Integer, ByRef sRazaoSocial As String, ByRef sNomeFantasia As String) As Long

        Dim iTamanho As Integer
        Dim sRetorno As String
        Dim dic As DataClasses2DataContext = New DataClasses2DataContext
        Dim db1 As DataClassesDataContext = New DataClassesDataContext
        Dim resdicControle As IEnumerable(Of Controle)
        Dim objControle As Controle
        Dim odbc As OdbcConnection = New OdbcConnection
        Dim sConexaoDic As String

        Dim iInicio As Integer
        Dim iFim As Integer
        Dim iPos5 As Integer
        Dim iPos6 As Integer
        Dim sFile As String
        Dim iPos As Integer
        Dim iEmpresa As Integer
        Dim resdicEmpresa As IEnumerable(Of Empresa)

        Try

            iTamanho = 255
            sRetorno = StrDup(iTamanho, Chr(0))

            iTamanho = GetPrivateProfileString("Geral", "DicDCConStr", "", sRetorno, iTamanho, sADM100INI)

            If gobjApp.iDebug = 1 Then MsgBox("0.1")
            If gobjApp.iDebug = 1 Then MsgBox(sRetorno)


            If iTamanho > 0 Then

                sRetorno = Mid(sRetorno, 1, iTamanho)

                sConexaoDic = sRetorno

                If gobjApp.iDebug = 1 Then MsgBox("0.11")
                If gobjApp.iDebug = 1 Then MsgBox(sConexaoDic)


            Else

                '***** coloca a string de conexao apontando para o SGEDic *****
                odbc.ConnectionString = "DSN=SGEDic;UID=admin;PWD=cacareco"

                iTamanho = 255
                sRetorno = StrDup(iTamanho, Chr(0))

                iTamanho = GetPrivateProfileString("Geral", "DicConStr", "", sRetorno, iTamanho, sADM100INI)

                If gobjApp.iDebug = 1 Then MsgBox("0.2")
                If gobjApp.iDebug = 1 Then MsgBox(sRetorno)

                If iTamanho > 0 Then

                    sRetorno = Mid(sRetorno, 1, iTamanho)

                    odbc.ConnectionString = sRetorno

                    If gobjApp.iDebug = 1 Then MsgBox("0.3")
                    If gobjApp.iDebug = 1 Then MsgBox(odbc.ConnectionString)

                End If


                odbc.Open()

                sConexaoDic = dic.Connection.ConnectionString

                iInicio = InStr(sConexaoDic, "Data Source=")

                iFim = InStr(iInicio, sConexaoDic, ";")

                sConexaoDic = Mid(sConexaoDic, 1, iInicio + 11) & odbc.DataSource & Mid(sConexaoDic, iFim)

                iInicio = InStr(sConexaoDic, "Initial Catalog=")

                iFim = InStr(iInicio, sConexaoDic, ";")

                sConexaoDic = Mid(sConexaoDic, 1, iInicio + 15) & odbc.Database & Mid(sConexaoDic, iFim)



                iTamanho = 255
                sRetorno = StrDup(iTamanho, Chr(0))

                iDebug = 0

                iTamanho = GetPrivateProfileString("Geral", "DataSource_Dic", "", sRetorno, iTamanho, sADM100INI)

                If iTamanho > 0 Then

                    sRetorno = Mid(sRetorno, 1, iTamanho)

                    iPos5 = InStr(sConexaoDic, "Data Source=")
                    iPos6 = InStr(sConexaoDic, ";Initial Catalog")


                    sConexaoDic = Mid(sConexaoDic, 1, iPos5 + 11) & sRetorno & Mid(sConexaoDic, iPos6)

                End If

                odbc.Close()

            End If

            dic.Connection.ConnectionString = sConexaoDic

            '        MsgBox(sConexaoDic)

            dic.Connection.Open()

            iDebug = 0

            resdicControle = dic.ExecuteQuery(Of Controle) _
            ("SELECT * FROM Controle WHERE Codigo = {0}", 100)

            For Each objControle In resdicControle
                iDebug = objControle.Conteudo
                Exit For
            Next

            '********** pega o diretorio do log para colocar os arquivos xml *************

            sDirXml = ""

            'verifica se tem no adm100.ini o diretorio especifico para a filial
            Dim lCodigo As Long

            lCodigo = CLng("9" & Format(CInt(sEmpresa), "00") & Format(iFilialEmpresa, "00"))

            resdicControle = dic.ExecuteQuery(Of Controle) _
            ("SELECT * FROM Controle WHERE Codigo = {0}", lCodigo)

            For Each objControle In resdicControle
                sDirXml = objControle.Conteudo
                Exit For
            Next

            If sDirXml = "" Then

                'verifica se tem no adm100.ini o diretorio especifico para a empresa

                iEmpresa = CInt(sEmpresa) + 900

                resdicControle = dic.ExecuteQuery(Of Controle) _
                ("SELECT * FROM Controle WHERE Codigo = {0}", iEmpresa)

                For Each objControle In resdicControle
                    sDirXml = objControle.Conteudo
                    Exit For
                Next

                If sDirXml = "" Then

                    resdicControle = dic.ExecuteQuery(Of Controle) _
                    ("SELECT * FROM Controle WHERE Codigo = {0}", 101)

                    For Each objControle In resdicControle
                        sDirXml = objControle.Conteudo
                        Exit For
                    Next

                    If sDirXml = "" Then

                        resdicControle = dic.ExecuteQuery(Of Controle) _
                        ("SELECT * FROM Controle WHERE Codigo = {0}", 1)

                        For Each objControle In resdicControle
                            sRetorno = objControle.Conteudo

                            sFile = Dir(sRetorno)

                            iPos = InStr(sRetorno, sFile)

                            sDirXml = Mid(sRetorno, 1, iPos - 1)

                            Exit For
                        Next

                    End If

                End If

            End If

            '********** pega o diretorio dos executaveis para ler os arquivos xsd *************
            sDirXsd = ""

            resdicControle = dic.ExecuteQuery(Of Controle) _
            ("SELECT * FROM Controle WHERE Codigo = {0}", 102)

            For Each objControle In resdicControle
                sDirXsd = objControle.Conteudo
                Exit For
            Next

            If sDirXsd = "" Then
                sDirXsd = "c:\sge\programa\xsd_310\"
            End If

            iTamanho = 255
            sRetorno = StrDup(iTamanho, Chr(0))

            iTamanho = GetPrivateProfileString("Geral", "DadosDCConStr" & sEmpresa, "", sRetorno, iTamanho, sADM100INI)

            If gobjApp.iDebug = 1 Then MsgBox("0.1")
            If gobjApp.iDebug = 1 Then MsgBox(sRetorno)


            If iTamanho > 0 Then

                sRetorno = Mid(sRetorno, 1, iTamanho)

                sConexaoDados = sRetorno

                If gobjApp.iDebug = 1 Then MsgBox("0.11")
                If gobjApp.iDebug = 1 Then MsgBox(sConexaoDados)


            Else


                '***** coloca a string de conexao apontando para o SGEDados em questao *****
                odbc.ConnectionString = "DSN=SGEDados" & sEmpresa & ";UID=sa;PWD=SAPWD"
                '            odbc.ConnectionString = "DSN=SGEDados" & sEmpresa & ";UID=sa;PWD=SAPWD"


                If gobjApp.iDebug = 1 Then MsgBox("0.1")
                If gobjApp.iDebug = 1 Then MsgBox(odbc.ConnectionString)

                iTamanho = 255
                sRetorno = StrDup(iTamanho, Chr(0))

                iTamanho = GetPrivateProfileString("Geral", "DadosConStr" & sEmpresa, "", sRetorno, iTamanho, sADM100INI)

                If gobjApp.iDebug = 1 Then MsgBox("0.2")
                If gobjApp.iDebug = 1 Then MsgBox(sRetorno)

                If iTamanho > 0 Then

                    sRetorno = Mid(sRetorno, 1, iTamanho)

                    odbc.ConnectionString = sRetorno

                    If gobjApp.iDebug = 1 Then MsgBox("0.3")
                    If gobjApp.iDebug = 1 Then MsgBox(odbc.ConnectionString)

                End If


                If gobjApp.iDebug = 1 Then MsgBox("0.4")
                If gobjApp.iDebug = 1 Then MsgBox(odbc.ConnectionString)

                odbc.Open()

                sConexaoDados = db1.Connection.ConnectionString

                iInicio = InStr(sConexaoDados, "Data Source=")

                iFim = InStr(iInicio, sConexaoDados, ";")

                sConexaoDados = Mid(sConexaoDados, 1, iInicio + 11) & odbc.DataSource & Mid(sConexaoDados, iFim)

                iInicio = InStr(sConexaoDados, "Initial Catalog=")

                iFim = InStr(iInicio, sConexaoDados, ";")

                sConexaoDados = Mid(sConexaoDados, 1, iInicio + 15) & odbc.Database & Mid(sConexaoDados, iFim)


                iTamanho = 255
                sRetorno = StrDup(iTamanho, Chr(0))

                iTamanho = GetPrivateProfileString("Geral", "DataSource" & sEmpresa, "", sRetorno, iTamanho, sADM100INI)

                If iTamanho > 0 Then

                    sRetorno = Mid(sRetorno, 1, iTamanho)

                    iPos5 = InStr(sConexaoDados, "Data Source=")
                    iPos6 = InStr(sConexaoDados, ";Initial Catalog")


                    sConexaoDados = Mid(sConexaoDados, 1, iPos5 + 11) & sRetorno & Mid(sConexaoDados, iPos6)

                End If

                odbc.Close()

            End If

            resdicEmpresa = dic.ExecuteQuery(Of Empresa) _
            ("SELECT * FROM Empresas WHERE Codigo = {0}", CLng(sEmpresa))

            For Each objEmpresa In resdicEmpresa
                sRazaoSocial = DesacentuaTexto(Trim(objEmpresa.Nome))
                sNomeFantasia = DesacentuaTexto(Trim(objEmpresa.NomeReduzido))
                Exit For
            Next

            Iniciar = SUCESSO

        Catch ex As Exception

            Iniciar = 1

            Form1.Msg.Items.Add("Erro na conexão ao dicionario de dados. " & ex.Message)

        End Try

    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
