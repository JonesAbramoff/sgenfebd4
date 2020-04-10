Option Strict Off
Option Explicit On

Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml
Imports System.Xml.Schema
Imports System.Xml.XPath
Imports System.Data.Odbc
Imports Microsoft.Win32
Imports System.Net
Imports sgenfebd4.NFeXsd

Public Class ClassGlobalApp

    Public iDebug As Integer
    Public sErro As String
    Public sMsg1 As String

    Public sADM100INI As String
    Public sCertificado As String
    Public iNFeAmbiente As Integer
    Public cert As New X509Certificate2
    Public iFilialEmpresa As Integer
    Public dbDadosNfe As New DataClassesDataContext
    Public dbLog As New DataClassesDataContext
    Public lEndereco As Long
    Public objEstado As New Estado
    Public sCGC As String
    Public sRazaoSocial As String
    Public sNomeFantasia As String
    Public objFilialEmpresa As New FiliaisEmpresa
    Public objEndereco As New Endereco
    Public objPais As New Paise
    Public objCidade As New Cidade
    Public sVersaoMsg As String

    Public sSistemaContingencia As String

    Public sDirXml As String
    Public sDirXsd As String

    Private sConexaoDados As String

    Private objDicInfo As New ClassDicInfo

    Public iScan As Integer

    Private dUltHoraLog As Double

    Public gsModelo As String

    Public Sub DadosCommit()
        'confirma a transacao e abre uma nova

        dbDadosNfe.Transaction.Commit()

        dbDadosNfe.Transaction = dbDadosNfe.Connection.BeginTransaction()

    End Sub

    Public Sub Terminar()

        dbDadosNfe.Connection.Close()
        dbLog.Connection.Close()

        dbDadosNfe.Dispose()
        dbLog.Dispose()

    End Sub

    Public Sub GravarLog(ByVal sTexto As String, ByVal lLote As Long, ByVal lNumIntNF As Long)

        Dim dHora As Double

        Form1.Msg.Items.Add(sTexto)

        Application.DoEvents()

        dHora = TimeOfDay.ToOADate

        If dHora - dUltHoraLog < DELTA_HORA Then dHora = dUltHoraLog + DELTA_HORA

        dUltHoraLog = dHora

        dbLog.Transaction = dbLog.Connection.BeginTransaction()

        Call dbLog.ExecuteCommand("INSERT INTO NFeFedLoteLog ( FilialEmpresa, Lote, Data, Hora, Status, NumIntNF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5})", _
            iFilialEmpresa, lLote, Now.Date, dUltHoraLog, Replace(Left(sTexto, 255), "'", "*"), lNumIntNF)

        dbLog.Transaction.Commit()

    End Sub

    Private Function AbrirBDsDados() As Long

        Dim resFilialEmpresa As IEnumerable(Of FiliaisEmpresa)
        Dim resEndereco As IEnumerable(Of Endereco)
        Dim resCidade As IEnumerable(Of Cidade)
        Dim resPais As IEnumerable(Of Paise)
        Dim resEstado As IEnumerable(Of Estado)
        Dim objAdmConfig1 As AdmConfig
        Dim resAdmConfig1 As IEnumerable(Of AdmConfig)
        Dim resAdmConfig2 As IEnumerable(Of AdmConfig)
        Dim objFilEmp As New FiliaisEmpresa

        Try


            'abrir bd de dados e obter informacoes gerais
            dbDadosNfe.Connection.ConnectionString = sConexaoDados
            dbDadosNfe.Connection.Open()

            dbLog.Connection.ConnectionString = sConexaoDados
            dbLog.Connection.Open()

            '***** inicia a trasacao *******************
            dbDadosNfe.Transaction = dbDadosNfe.Connection.BeginTransaction()

            resFilialEmpresa = dbDadosNfe.ExecuteQuery(Of FiliaisEmpresa) _
            ("SELECT * FROM FiliaisEmpresa WHERE FilialEmpresa = {0} ", iFilialEmpresa)

            For Each objFilEmp In resFilialEmpresa
                sErro = "8"
                sCertificado = objFilEmp.CertificadoA1A3
                sErro = "9"
                iNFeAmbiente = objFilEmp.NFeAmbiente
                sErro = "10"
                lEndereco = objFilEmp.Endereco
                sErro = "11"
                sCGC = objFilEmp.CGC
                sErro = "12"
                Exit For
            Next

            If iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then Form1.EmHomologacao.Visible = True

            objFilialEmpresa = objFilEmp

            resEndereco = dbDadosNfe.ExecuteQuery(Of Endereco) _
            ("SELECT * FROM Enderecos WHERE Codigo = {0}", lEndereco)
            objEndereco = resEndereco(0)

            resEstado = dbDadosNfe.ExecuteQuery(Of Estado) _
                ("SELECT * FROM Estados WHERE Sigla = {0}", objEndereco.SiglaEstado)
            objEstado = resEstado(0)

            resCidade = dbDadosNfe.ExecuteQuery(Of Cidade) _
            ("SELECT * FROM Cidades WHERE Descricao = {0}", objEndereco.Cidade)
            objCidade = resCidade(0)

            resPais = dbDadosNfe.ExecuteQuery(Of Paise) _
            ("SELECT * FROM Paises WHERE Codigo = {0}", objEndereco.CodigoPais)
            objPais = resPais(0)

            resAdmConfig1 = dbDadosNfe.ExecuteQuery(Of AdmConfig) _
            ("SELECT * FROM AdmConfig WHERE  Codigo = {0} AND FilialEmpresa = {1}", "EMPRESA_RAZAO_SOCIAL", iFilialEmpresa)
            If resAdmConfig1.Count > 0 Then

                resAdmConfig1 = dbDadosNfe.ExecuteQuery(Of AdmConfig) _
                ("SELECT * FROM AdmConfig WHERE  Codigo = {0} AND FilialEmpresa = {1} ", "EMPRESA_RAZAO_SOCIAL", iFilialEmpresa)

                objAdmConfig1 = resAdmConfig1(0)
                sRazaoSocial = objAdmConfig1.Conteudo

            End If

            resAdmConfig1 = dbDadosNfe.ExecuteQuery(Of AdmConfig) _
            ("SELECT * FROM AdmConfig WHERE  Codigo = {0} AND FilialEmpresa = {1}", "EMPRESA_NOME_FANTASIA", iFilialEmpresa)
            If resAdmConfig1.Count > 0 Then

                resAdmConfig1 = dbDadosNfe.ExecuteQuery(Of AdmConfig) _
                ("SELECT * FROM AdmConfig WHERE  Codigo = {0} AND FilialEmpresa = {1} ", "EMPRESA_NOME_FANTASIA", iFilialEmpresa)

                objAdmConfig1 = resAdmConfig1(0)
                sNomeFantasia = objAdmConfig1.Conteudo

            End If

            resAdmConfig2 = dbDadosNfe.ExecuteQuery(Of AdmConfig) _
            ("SELECT * FROM AdmConfig WHERE  Codigo = {0} ", "VERSAO_MSG")
            If resAdmConfig2.Count > 0 Then

                resAdmConfig2 = dbDadosNfe.ExecuteQuery(Of AdmConfig) _
                ("SELECT * FROM AdmConfig WHERE  Codigo = {0} ", "VERSAO_MSG")

                sVersaoMsg = resAdmConfig2(0).Conteudo
            Else
                sVersaoMsg = ""
            End If

            sErro = ""

            AbrirBDsDados = SUCESSO

        Catch ex As Exception

            AbrirBDsDados = 1

            Form1.Msg.Items.Add("Erro na abertura dos bancos de dados. " & ex.Message & " " & sErro)

        End Try

    End Function

    Public Function Iniciar(ByVal sEmpresa As String, ByVal iFilialEmp As Integer) As Long

        Dim lErro As Long

        Try

            Call ObterAdm100Ini()

            iFilialEmpresa = iFilialEmp

            'obter informacoes do dicionario de dados
            lErro = objDicInfo.Iniciar(sEmpresa, iFilialEmpresa, sADM100INI, sDirXml, sDirXsd, sConexaoDados, iDebug, sRazaoSocial, sNomeFantasia)
            If lErro <> SUCESSO Then Throw New System.Exception("Erro obtendo informações do dicionario de dados.")

            lErro = AbrirBDsDados()
            If lErro <> SUCESSO Then Throw New System.Exception("")

            Call GravarLog("Iniciando processamento...", 0, 0)

            '
            '  seleciona certificado do repositório MY do windows
            '
            lErro = ObtemCertificado()
            If lErro <> SUCESSO Then Throw New System.Exception("Erro obtendo certificado digital.")

            Iniciar = SUCESSO

        Catch ex As Exception

            Iniciar = 1

            Form1.Msg.Items.Add("Erro na inicialização dos bancos de dados. " & ex.Message)

        End Try

    End Function

    Private Sub ObterAdm100Ini()

        Dim iTamanho As Integer, sRetorno As String
        sADM100INI = Environment.GetEnvironmentVariable("USERPROFILE") & "\windows\adm100.ini"

        iTamanho = 255
        sRetorno = StrDup(iTamanho, Chr(0))
        iTamanho = GetPrivateProfileString("Geral", "LockFile", "", sRetorno, iTamanho, sADM100INI)
        If iTamanho = 0 Then sADM100INI = Environment.GetEnvironmentVariable("WINDIR") & "\adm100.ini"

    End Sub

    Private Function ObtemCertificado() As Long
        '
        '  seleciona certificado do repositório MY do windows
        '
        Try

            Dim certificado As Certificado = New Certificado

            If UCase(sCertificado) = "FORPRINT" Then
                iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO
            End If

            Dim lErro As Long
            lErro = certificado.BuscaNome(sCertificado, cert)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            If DateDiff(DateInterval.Day, Now.Date, cert.NotAfter) < 15 And DateDiff(DateInterval.Day, Now.Date, cert.NotAfter) >= 0 Then

                GravarLog("ATENÇÃO: FALTAM " & DateDiff(DateInterval.Day, Now.Date, cert.NotAfter) & " DIAS PARA TERMINAR A VALIDADE DO SEU CERTIFICADO. FAVOR RENOVA-LO.", 0, 0)

            ElseIf DateDiff(DateInterval.Day, Now.Date, cert.NotAfter) < 0 Then

                Throw New System.Exception("ATENÇÃO: O CERTIFICADO ESTÁ COM O PRAZO DE VALIDADE VENCIDO. FAVOR RENOVA-LO.")

            End If

            ObtemCertificado = SUCESSO

        Catch ex As Exception

            ObtemCertificado = 1

            Form1.Msg.Items.Add("Erro na seleção do certificado digital. " & ex.Message)

        End Try

    End Function

    Public Function Verifica_Status_Servico() As Long

        Try

            Dim XMLStreamRet As MemoryStream = New MemoryStream(10000)
            Dim xRet As Byte()
            Dim xmlNode1 As XmlNode
            Dim iPos As Integer
            Dim objStatServ As TConsStatServ = New TConsStatServ
            Dim objRetStatServ As TRetConsStatServ = New TRetConsStatServ
            Dim XMLStringRetStatServ As String
            Dim XMLString4 As String
            Dim XMLStream As MemoryStream = New MemoryStream(10000)
            Dim NfeStatusServico As New nfestatusservico2.NFeStatusServico4
            'Dim NFecabec_StatusServ As New nfestatusservico2.nfeCabecMsg

            objStatServ.cUF = GetCode(Of TCodUfIBGE)(CStr(objEstado.CodIBGE))

            objStatServ.versao = NFE_VERSAO_XML

            If iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then
                objStatServ.tpAmb = TAmb.Item2
            Else
                objStatServ.tpAmb = TAmb.Item1
            End If

            objStatServ.xServ = TConsStatServXServ.STATUS

            Dim mySerializerZ As New XmlSerializer(GetType(TConsStatServ))

            XMLStream = New MemoryStream(10000)

            mySerializerZ.Serialize(XMLStream, objStatServ)

            Dim xmz As Byte()
            xmz = XMLStream.ToArray

            XMLString4 = System.Text.Encoding.UTF8.GetString(xmz)

            XMLString4 = Mid(XMLString4, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLString4, 20)

            sErro = "39.3"
            sMsg1 = "vai gravar NFeFedLoteLot"

            Call GravarLog("Iniciando a verificação do status do serviço", 0, 0)

            '****************  salva o arquivo 

            iPos = InStr(XMLString4, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")


            If iPos <> 0 Then

                XMLString4 = Mid(XMLString4, 1, iPos - 1) & Mid(XMLString4, iPos + 99)

            End If


            Dim DocDados1 As XmlDocument = New XmlDocument

            Call Salva_Arquivo(DocDados1, XMLString4)

            'NFecabec_StatusServ.cUF = CStr(objEstado.CodIBGE)
            'NFecabec_StatusServ.versaoDados = NFE_VERSAO_XML

            'NfeStatusServico.nfeCabecMsgValue = NFecabec_StatusServ

            sErro = "39.4"
            sMsg1 = "vai fazer consulta a Status do Servico"

            Dim sURL As String
            sURL = ""
            Call WS_Obter_URL(sURL, iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO, objEstado.Sigla, "NfeStatusServico", gobjApp.gsModelo)

            NfeStatusServico.Url = sURL

            NfeStatusServico.ClientCertificates.Add(cert)
            xmlNode1 = NfeStatusServico.nfeStatusServicoNF(DocDados1)

            XMLStringRetStatServ = xmlNode1.OuterXml

            If iDebug = 1 Then
                MsgBox("39.5")
                MsgBox(XMLStringRetStatServ)
            End If
            sErro = "39.5"
            sMsg1 = "consultou o Status do Servico"

            xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringRetStatServ)

            XMLStreamRet = New MemoryStream(10000)
            XMLStreamRet.Write(xRet, 0, xRet.Length)

            Dim mySerializerRetConsStatServ1 As New XmlSerializer(GetType(TRetConsStatServ))

            Dim objRetConsStatServ1 As TRetConsStatServ = New TRetConsStatServ

            XMLStreamRet.Position = 0

            objRetConsStatServ1 = mySerializerRetConsStatServ1.Deserialize(XMLStreamRet)

            If objRetConsStatServ1.cStat <> "107" Then
                Form1.Msg.Items.Add(XMLString4)
                Form1.Msg.Items.Add(XMLStringRetStatServ)


                If sSistemaContingencia = "SCAN" Then
                    If iScan = -1 Then
                        Throw New System.Exception("Serviço não está em operação. Tente usar notas da serie 900 a 999.")
                    Else
                        Throw New System.Exception("Serviço Scan não está em operação. Tente usar notas da serie abaixo de 900.")
                    End If
                Else
                    If iScan = -1 Then
                        Throw New System.Exception("O serviço não está em operação. Verifique se estão operando pelo sistema de contingência.")
                    Else
                        Throw New System.Exception("O serviço de contingência não está em operação. Verifique se o sistema normal está em funcionamento.")
                    End If
                End If

            End If

            Verifica_Status_Servico = SUCESSO

        Catch ex As Exception

            Verifica_Status_Servico = 1

            Call GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Public Function Local_ObterBD(ByVal objLocal As TLocal, ByVal objEndereco As Endereco) As Long

        Dim resCidade As IEnumerable(Of Cidade)
        Dim resPais As IEnumerable(Of Paise)
        Dim resEstado As IEnumerable(Of Estado)
        Dim objCidade As New Cidade
        Dim objEstado As New Estado
        Dim objPais As New Paise

        Try

            If Len(objEndereco.Logradouro) > 0 Or Len(objEndereco.Endereco) > 0 Then

                resEstado = dbDadosNfe.ExecuteQuery(Of Estado) _
                ("SELECT * FROM Estados WHERE Sigla = {0}", objEndereco.SiglaEstado)

                For Each objEstado In resEstado

                    Exit For
                Next

                resCidade = dbDadosNfe.ExecuteQuery(Of Cidade) _
                ("SELECT * FROM Cidades WHERE Descricao = {0}", objEndereco.Cidade)

                For Each objCidade In resCidade
                    Exit For
                Next

                resPais = dbDadosNfe.ExecuteQuery(Of Paise) _
                ("SELECT * FROM Paises WHERE Codigo = {0}", objEndereco.CodigoPais)

                For Each objPais In resPais
                    Exit For
                Next

                If Len(objEndereco.Logradouro) > 0 Then
                    objLocal.xLgr = DesacentuaTexto(Left(IIf(Len(objEndereco.TipoLogradouro) > 0, objEndereco.TipoLogradouro & " ", "") & objEndereco.Logradouro, 60))
                    objLocal.nro = objEndereco.Numero
                    If Len(DesacentuaTexto(objEndereco.Complemento)) > 0 Then objLocal.xCpl = DesacentuaTexto(objEndereco.Complemento)
                Else
                    objLocal.xLgr = DesacentuaTexto(objEndereco.Endereco)
                    objLocal.nro = "0"
                End If

                If Len(objEndereco.Bairro) = 0 Then
                    objLocal.xBairro = "a"
                Else
                    objLocal.xBairro = Trim(DesacentuaTexto(objEndereco.Bairro))
                End If

                'se for Brasil 
                If objPais.CodBacen = 1058 Then
                    objLocal.cMun = objCidade.CodIBGE
                    objLocal.xMun = DesacentuaTexto(objCidade.Descricao)
                    objLocal.UF = GetCode(Of TUf)(objEndereco.SiglaEstado)

                Else
                    objLocal.cMun = "9999999"
                    objLocal.xMun = "EXTERIOR"
                    objLocal.UF = GetCode(Of TUf)("EX")

                End If

            End If

            Local_ObterBD = SUCESSO

        Catch ex As Exception

            Local_ObterBD = 1

            Call GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Public Function Endereco_ObterBD(ByRef objEndereco As TEndereco, ByVal lEndereco As Long, ByRef sEmail As String) As Long

        Dim resEndereco As IEnumerable(Of Endereco)
        Dim resCidade As IEnumerable(Of Cidade)
        Dim resPais As IEnumerable(Of Paise)
        Dim resEstado As IEnumerable(Of Estado)
        Dim objCidade As New Cidade
        Dim objEstado As New Estado
        Dim objPais As New Paise

        Try

            sErro = "15"
            sMsg1 = "vai acessar a tabela Endereco do Destinatario, Estado, Cidades e Paises"

            resEndereco = dbDadosNfe.ExecuteQuery(Of Endereco) _
            ("SELECT * FROM Enderecos WHERE Codigo = {0}", lEndereco)

            For Each objEndBD In resEndereco

                sEmail = objEndBD.Email

                resEstado = dbDadosNfe.ExecuteQuery(Of Estado) _
                ("SELECT * FROM Estados WHERE Sigla = {0}", objEndBD.SiglaEstado)

                For Each objEstado In resEstado

                    Exit For
                Next

                resCidade = dbDadosNfe.ExecuteQuery(Of Cidade) _
                ("SELECT * FROM Cidades WHERE Descricao = {0}", objEndBD.Cidade)

                For Each objCidade In resCidade
                    Exit For
                Next

                resPais = dbDadosNfe.ExecuteQuery(Of Paise) _
                ("SELECT * FROM Paises WHERE Codigo = {0}", objEndBD.CodigoPais)

                For Each objPais In resPais
                    '    objEndereco.cPais = GetCode(Of NFeXsd.Tpais)(objPais.CodBacen)
                    '    'objEndereco.cPaisSpecified = True
                    Exit For
                Next
                objEndereco.cPais = objPais.CodBacen
                objEndereco.xPais = objPais.Nome

                If Len(objEndBD.Logradouro) > 0 Then
                    objEndereco.xLgr = DesacentuaTexto(Left(IIf(Len(objEndBD.TipoLogradouro) > 0, objEndBD.TipoLogradouro & " ", "") & objEndBD.Logradouro, 60))
                    objEndereco.nro = objEndBD.Numero
                    If Len(DesacentuaTexto(objEndBD.Complemento)) > 0 Then objEndereco.xCpl = DesacentuaTexto(objEndBD.Complemento)
                Else
                    objEndereco.xLgr = DesacentuaTexto(objEndBD.Endereco)
                    objEndereco.nro = "0"
                End If

                If Len(objEndBD.TelNumero1) > 0 Then
                    Call Formata_String_Numero(IIf(Len(CStr(objEndBD.TelDDD1)) > 0, CStr(objEndBD.TelDDD1), "") + objEndBD.TelNumero1, objEndereco.fone)
                ElseIf Len(objEndBD.Telefone1) > 0 Then
                    Call Formata_String_Numero(objEndBD.Telefone1, objEndereco.fone)
                End If

                If Len(objEndBD.Bairro) = 0 Then
                    objEndereco.xBairro = "NAO INFORMADO"
                Else
                    objEndereco.xBairro = Trim(DesacentuaTexto(objEndBD.Bairro))
                End If

                'se for Brasil 
                If objPais.CodBacen = 1058 Then

                    objEndereco.cMun = objCidade.CodIBGE
                    objEndereco.xMun = DesacentuaTexto(objCidade.Descricao)
                    objEndereco.UF = GetCode(Of TUf)(objEndBD.SiglaEstado)

                    If Len(objEndBD.CEP) > 0 Then
                        objEndereco.CEP = objEndBD.CEP
                    End If
                Else

                    objEndereco.cMun = "9999999"
                    objEndereco.xMun = "EXTERIOR"
                    objEndereco.UF = GetCode(Of TUf)("EX")

                End If

                Exit For
            Next

            Endereco_ObterBD = SUCESSO

        Catch ex As Exception

            Endereco_ObterBD = 1

            Call GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Public Function NFCe_Trata_QRCode(ByRef sQRCode As String, ByVal objprotNFe As TProtNFe, ByVal DocDados As XmlDocument) As Long

        Try

            Dim sDigVal As String, scDest As String = "", sdhEmi As String, svNF As String = "", svICMS As String = ""

            Dim ns As New XmlNamespaceManager(DocDados.NameTable)
            ns.AddNamespace("nfe", "http://www.portalfiscal.inf.br/nfe")
            Dim xpathNav As XPathNavigator = DocDados.CreateNavigator()
            Dim node As XPathNavigator = xpathNav.SelectSingleNode("//nfe:protNFe/nfe:infProt/nfe:digVal", ns)
            sDigVal = node.InnerXml

            node = xpathNav.SelectSingleNode("//nfe:NFe/nfe:infNFe/nfe:ide/nfe:dhEmi", ns)
            sdhEmi = node.InnerXml

            node = xpathNav.SelectSingleNode("//nfe:NFe/nfe:infNFe/nfe:dest/nfe:CNPJ", ns)
            If Not (node Is Nothing) Then
                scDest = node.InnerXml
            Else
                node = xpathNav.SelectSingleNode("//nfe:NFe/nfe:infNFe/nfe:dest/nfe:CPF", ns)
                If Not (node Is Nothing) Then
                    scDest = node.InnerXml
                Else
                    node = xpathNav.SelectSingleNode("//nfe:NFe/nfe:infNFe/nfe:dest/nfe:idEstrangeiro", ns)
                    If Not (node Is Nothing) Then
                        scDest = node.InnerXml
                    End If
                End If
            End If

            node = xpathNav.SelectSingleNode("//nfe:NFe/nfe:infNFe/nfe:total/nfe:ICMSTot/nfe:vICMS", ns)
            svICMS = node.InnerXml

            node = xpathNav.SelectSingleNode("//nfe:NFe/nfe:infNFe/nfe:total/nfe:ICMSTot/nfe:vNF", ns)
            svNF = node.InnerXml

            'MsgBox("cdest: " & scDest & " dhemi: " & sdhEmi & " vnf: " & svNF & " vicms: " & svICMS & " digval: " & sDigVal)

            sQRCode = NFCE_Gera_QRCode(objprotNFe.infProt.chNFe, "100", GetXmlAttrNameFromEnumValue(Of TAmb)(objprotNFe.infProt.tpAmb), scDest, sdhEmi, svNF, svICMS, sDigVal, objFilialEmpresa.idNFCECSC, objFilialEmpresa.NFCECSC)

            NFCe_Trata_QRCode = SUCESSO

        Catch ex As Exception

            NFCe_Trata_QRCode = 1

            Call GravarLog(ex.Message, 0, 0)

        Finally

        End Try

    End Function

    Public Function Evento_Grava_Xml(ByVal objprocEventoNFe As Object) As Long

        Dim DocDados2 As XmlDocument = New XmlDocument
        Dim XMLStreamDados1 = New MemoryStream(10000)
        Dim XMLStreamDados2 = New MemoryStream(10000)
        Dim sArquivo As String
        Dim iPos As Integer
        Dim XMLStreamDados = New MemoryStream(10000)

        Try

            Dim mySerializerProcEvento As New XmlSerializer(GetType(EventoXsd.TProcEvento))

            Dim XMLStream1 = New MemoryStream(10000)

            Dim objprocEvento As New EventoXsd.TProcEvento

            With objprocEvento

                .evento = New EventoXsd.TEvento
                Call ClonarEstruturasSerializaveis(.evento, objprocEventoNFe.evento)

                .retEvento = New EventoXsd.TRetEvento
                Call ClonarEstruturasSerializaveis(.retEvento, objprocEventoNFe.retEvento)

                .versao = objprocEventoNFe.versao

            End With

            mySerializerProcEvento.Serialize(XMLStream1, objprocEvento)

            Dim xmw1 As Byte()
            Dim XMLStringProc As String

            xmw1 = XMLStream1.ToArray

            XMLStringProc = System.Text.Encoding.UTF8.GetString(xmw1)

            XMLStringProc = Mid(XMLStringProc, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLStringProc, 20)

            iPos = InStr(XMLStringProc, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")

            If iPos <> 0 Then

                XMLStringProc = Mid(XMLStringProc, 1, iPos - 1) & Mid(XMLStringProc, iPos + 99)

            End If

            Dim xDadosCanc As Byte()

            xDadosCanc = System.Text.Encoding.UTF8.GetBytes(XMLStringProc)

            XMLStreamDados = New MemoryStream(10000)

            XMLStreamDados.Write(xDadosCanc, 0, xDadosCanc.Length)

            sErro = "25"
            sMsg1 = "vai gravar o xml"

            Dim DocDadosCanc As XmlDocument = New XmlDocument
            XMLStreamDados.Position = 0
            DocDadosCanc.Load(XMLStreamDados)
            sArquivo = sDirXml & objprocEventoNFe.retEvento.infEvento.tpEvento.ToString & "-" & objprocEventoNFe.retEvento.infEvento.chNFe & "-" & objprocEventoNFe.retEvento.infEvento.nSeqEvento & "-procEventoNfe.xml"

            Dim writer As New XmlTextWriter(sArquivo, Nothing)

            writer.Formatting = Formatting.None
            DocDadosCanc.WriteTo(writer)
            writer.Close()

            Evento_Grava_Xml = 0

        Catch ex As Exception

            Evento_Grava_Xml = 1

            Call GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Public Sub New()

        iScan = -1
        sSistemaContingencia = ""
        gsModelo = ""

    End Sub

End Class
