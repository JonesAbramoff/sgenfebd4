Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml
Imports System.Xml.Schema
Imports System.Data.Odbc
Imports Microsoft.Win32
Imports sgenfebd4.NFeXsd

Public Class ClassConsultaNFe

    Public Function Consulta_NFe(ByVal sEmpresa As String, ByVal sChaveNFe As String, ByVal iFilialEmpresa As Integer, ByVal iOperacao As Integer, Optional ByVal iScan As Integer = -1, Optional ByVal iRenomearArq As Integer = 0) As Long
        'iOperacao (0=nao especificado, 1=cancelamento)

        Dim iSerie As Integer
        Dim sSerie As String
        Dim lNumNotaFiscal As Long
        Dim objAdmConfig As AdmConfig
        Dim a5 As TNFe

        Dim dbDadosNfe As DataClassesDataContext = New DataClassesDataContext
        Dim dbLog As DataClassesDataContext = New DataClassesDataContext
        Dim iResult As Integer
        Dim XMLStream1 As MemoryStream = New MemoryStream(10000)
        'Dim XMLStreamCabec As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamRet As MemoryStream = New MemoryStream(10000)


        Dim XMLString1 As String

        Dim resNFeFedRetCancNfe As IEnumerable(Of NFeFedRetCancNFe)
        Dim resNFeFedRetEnvCCe As IEnumerable(Of NFeFedRetEnvCCe)
        Dim resNFeFedProtNFe As IEnumerable(Of NFeFedProtNFe)
        Dim resNFeNFiscal As IEnumerable(Of NFeNFiscal)
        Dim results As IEnumerable(Of NFeNFiscal)
        Dim resAdmConfig As IEnumerable(Of AdmConfig)
        Dim resNFeFedDenegada As IEnumerable(Of NFeFedDenegada)

        Dim objNFeNFiscal As NFeNFiscal
        Dim objNFeFedLote As NFeFedLote = New NFeFedLote
        Dim objNFeFedRetEnvi As NFeFedRetEnvi = New NFeFedRetEnvi
        Dim objNFeFedProtNFe As NFeFedProtNFe = New NFeFedProtNFe
        Dim xRet As Byte()
        Dim sErro1 As String

        Dim colNumIntNFiscal As Collection = New Collection
        'Dim XMLStringCabec As String

        Dim iPos As Integer
        Dim lErro As Long

        Dim XMLStringNFes As String

        Dim sArquivo As String
        'Dim xmcabec As Byte()

        Dim colNFiscal As Collection = New Collection

        Dim xmlNode1 As XmlNode
        Dim iProcessado As Integer

        Try

            If gobjApp.iDebug = 1 Then MsgBox("6")

            gobjApp.sErro = "6"
            gobjApp.sMsg1 = "vai inserir na tabela NFeFedLoteLog"

            Call gobjApp.GravarLog("Iniciando a consulta da nota fiscal com chave " & sChaveNFe, 0, 0)

            If gobjApp.iDebug = 1 Then MsgBox("13")
            gobjApp.sErro = "13"
            gobjApp.sMsg1 = "vai montar o cabecalho"

            'Dim mySerializercabec As Object
            'Dim objCabecMsg As Object
            Dim NfeConsulta As New nfeconsulta2.NFeConsultaProtocolo4
            'Dim NfeCabec As Object

            '
            'XMLStreamCabec = New MemoryStream(10000)

            'mySerializercabec.Serialize(XMLStreamCabec, objCabecMsg)

            'XMLStreamCabec.Position = 0

            'xmcabec = XMLStreamCabec.ToArray

            'XMLStringCabec = System.Text.Encoding.UTF8.GetString(xmcabec)

            'XMLStringCabec = Mid(XMLStringCabec, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLStringCabec, 20)

            If gobjApp.iDebug = 1 Then MsgBox("19")
            gobjApp.sErro = "19"
            gobjApp.sMsg1 = "vai montar a estrutura TConsSitNFe"

            Dim objconsSitNFe As TConsSitNFe = New TConsSitNFe

            objconsSitNFe.chNFe = sChaveNFe

            gobjApp.gsModelo = chNFe_Retorna_Modelo(sChaveNFe)

            If gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then
                objconsSitNFe.tpAmb = TAmb.Item2
            Else
                objconsSitNFe.tpAmb = TAmb.Item1
            End If

            objconsSitNFe.versao = TVerConsSitNFe.Item400
            objconsSitNFe.xServ = TConsSitNFeXServ.CONSULTAR

            If gobjApp.iDebug = 1 Then MsgBox("20")
            gobjApp.sErro = "20"
            gobjApp.sMsg1 = "vai serializar TConsSitNFe"

            Dim mySerializerw As New XmlSerializer(GetType(TConsSitNFe))

            If gobjApp.iDebug = 1 Then MsgBox("20.a")
            gobjApp.sErro = "20.a"
            gobjApp.sMsg1 = "vai serializar TConsSitNFe"

            Dim XMLStream2 = New MemoryStream(10000)
            mySerializerw.Serialize(XMLStream2, objconsSitNFe)

            If gobjApp.iDebug = 1 Then MsgBox("20.b")
            gobjApp.sErro = "20.b"
            gobjApp.sMsg1 = "vai serializar TConsSitNFe"

            Dim xm2 As Byte()
            xm2 = XMLStream2.ToArray

            If gobjApp.iDebug = 1 Then MsgBox("20.c")
            gobjApp.sErro = "20.c"
            gobjApp.sMsg1 = "vai serializar TConsSitNFe"

            XMLString1 = System.Text.Encoding.UTF8.GetString(xm2)

            If gobjApp.iDebug = 1 Then MsgBox("20.d")
            gobjApp.sErro = "20.d" & " " & XMLString1
            gobjApp.sMsg1 = "vai serializar TConsSitNFe"

            XMLString1 = Mid(XMLString1, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLString1, 20)

            iPos = InStr(XMLString1, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")

            If iPos <> 0 Then

                XMLString1 = Mid(XMLString1, 1, iPos - 1) & Mid(XMLString1, iPos + 99)

            End If


            Dim XMLStringRetConsultaNF As String

            Dim DocDados1 As New XmlDocument

            If gobjApp.iDebug = 1 Then MsgBox("20.1")


            Call Salva_Arquivo(DocDados1, XMLString1)

            If gobjApp.iDebug = 1 Then MsgBox("20.2")

            'NfeCabec.cUF = CStr(gobjApp.objEstado.CodIBGE)
            'NfeCabec.versaoDados = NFE_VERSAO_XML

            'NfeConsulta.nfeCabecMsgValue = NfeCabec

            Dim sURL As String
            sURL = ""
            Call WS_Obter_URL(sURL, gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO, gobjApp.objEstado.Sigla, "NfeConsultaProtocolo", gobjApp.gsModelo)

            NfeConsulta.Url = sURL

            NfeConsulta.ClientCertificates.Add(gobjApp.cert)

            Select Case WS_Obter_Autorizador(gobjApp.objEstado.Sigla)

                'Case "PR", "BA"
                'xmlNode1 = NfeConsulta.nfeConsultaNF(DocDados1)

                Case Else
                    xmlNode1 = NfeConsulta.nfeConsultaNF(DocDados1)

            End Select

            If gobjApp.iDebug = 1 Then MsgBox("20.3")

            XMLStringRetConsultaNF = xmlNode1.OuterXml

            '**** codigo criado para retirar os procEnventoNFe q estao impedindo a desserializacao do xml ****

            If gobjApp.iDebug = 1 Then MsgBox(XMLStringRetConsultaNF)

            Dim XMLStreamDadosDebug = New MemoryStream(10000)

            Dim xDadosDebug As Byte()

            xDadosDebug = System.Text.Encoding.UTF8.GetBytes(XMLStringRetConsultaNF)

            XMLStreamDadosDebug.Write(xDadosDebug, 0, xDadosDebug.Length)

            If gobjApp.iDebug = 1 Then MsgBox("20.4")


            Dim DocDadosDebug3 As XmlDocument = New XmlDocument

            XMLStreamDadosDebug.Position = 0
            DocDadosDebug3.Load(XMLStreamDadosDebug)

            sArquivo = gobjApp.sDirXml & sChaveNFe & "-consultanfe.xml"


            Dim writerDebug As New XmlTextWriter(sArquivo, Nothing)

            writerDebug.Formatting = Formatting.None
            DocDadosDebug3.WriteTo(writerDebug)

            writerDebug.Close()

            If gobjApp.iDebug = 1 Then MsgBox("20.5")

            xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringRetConsultaNF)

            XMLStreamRet = New MemoryStream(10000)
            XMLStreamRet.Write(xRet, 0, xRet.Length)

            Dim mySerializerRetConsultaNF As New XmlSerializer(GetType(TRetConsSitNFe))

            Dim objRetConsSitNFe As TRetConsSitNFe = New TRetConsSitNFe

            XMLStreamRet.Position = 0

            If gobjApp.iDebug = 1 Then MsgBox("20.6")

            objRetConsSitNFe = mySerializerRetConsultaNF.Deserialize(XMLStreamRet)

            If gobjApp.iDebug = 1 Then MsgBox("21")
            gobjApp.sErro = "21"
            gobjApp.sMsg1 = "trata a nota fiscal consultada"

            Call gobjApp.GravarLog("CStat = " & objRetConsSitNFe.cStat & "  " & objRetConsSitNFe.xMotivo, 0, 0)

            iProcessado = 0

            If Not objRetConsSitNFe.protNFe Is Nothing Then

                If gobjApp.iDebug = 1 Then MsgBox("21.1")

                If objRetConsSitNFe.protNFe.infProt.cStat = "100" Or objRetConsSitNFe.protNFe.infProt.cStat = "150" Or objRetConsSitNFe.protNFe.infProt.cStat = "205" Or objRetConsSitNFe.protNFe.infProt.cStat = "110" Or objRetConsSitNFe.protNFe.infProt.cStat = "301" Or objRetConsSitNFe.protNFe.infProt.cStat = "302" Then

                    If gobjApp.iDebug = 1 Then MsgBox("22")

                    iSerie = CInt(Mid(sChaveNFe, 23, 3))
                    sSerie = iSerie & "-e"
                    lNumNotaFiscal = CLng(Mid(sChaveNFe, 26, 9))

                    Dim iAno As Integer
                    Dim iMes As Integer

                    iAno = CInt(Mid(sChaveNFe, 3, 2)) + 2000
                    iMes = CInt(Mid(sChaveNFe, 5, 2))

                    '                lcNF = CLng(Mid(sChaveNFe, 36, 8))

                    results = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeNFiscal) _
                    ("SELECT * FROM NFeNFiscal WHERE Serie = {0} AND NumNotaFiscal = {1} AND FilialEmpresa = {2} AND YEAR(DataEmissao) = {3} AND MONTH(DataEmissao) = {4}", sSerie, lNumNotaFiscal, iFilialEmpresa, iAno, iMes)

                    For Each objNFeNFiscal In results

                        If gobjApp.iDebug = 1 Then MsgBox("23")

                        resNFeFedProtNFe = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedProtNFe) _
                        ("SELECT * FROM NFeFedProtNFe WHERE (cStat = '100' Or cStat = '150' Or cStat = '205' Or cStat = '110' Or cStat = '301' Or cStat = '302') AND FilialEmpresa = {0} AND chNFe = {1}", iFilialEmpresa, sChaveNFe)

                        If resNFeFedProtNFe.Count = 0 Then

                            iResult = gobjApp.dbDadosNfe.ExecuteCommand("INSERT INTO NFeFedProtNFe ( FilialEmpresa, NumIntNF, versao, nRec, tpAmb, verAplic, chNFe, nProt, cStat, xMotivo, Data, Hora, DataRegistro, HoraRegistro) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13} )", _
                            iFilialEmpresa, objNFeNFiscal.NumIntDoc, objRetConsSitNFe.protNFe.versao, "", objRetConsSitNFe.tpAmb, objRetConsSitNFe.verAplic, objRetConsSitNFe.chNFe, objRetConsSitNFe.protNFe.infProt.nProt, objRetConsSitNFe.cStat, objRetConsSitNFe.xMotivo, UTCParaDate(objRetConsSitNFe.protNFe.infProt.dhRecbto), UTCParaHora(objRetConsSitNFe.protNFe.infProt.dhRecbto), Now.Date, TimeOfDay.ToOADate)

                        Else

                            resNFeFedProtNFe = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedProtNFe) _
                            ("SELECT * FROM NFeFedProtNFe WHERE (cStat = '100' Or cStat = '150' Or cStat = '205' Or cStat = '110' Or cStat = '301' Or cStat = '302') AND FilialEmpresa = {0} AND chNFe = {1}", iFilialEmpresa, sChaveNFe)

                            For Each objNFeFedProtNFe In resNFeFedProtNFe
                                If objNFeFedProtNFe.NumIntNF <> objNFeNFiscal.NumIntDoc Then
                                    iResult = gobjApp.dbDadosNfe.ExecuteCommand("UPDATE NFeFedProtNFe SET NumIntNF = {0} WHERE (cStat = '100' Or cStat = '150' Or cStat = '205' Or cStat = '110' Or cStat = '301' Or cStat = '302') AND FilialEmpresa = {1} AND chNFe = {2}", _
                                    objNFeNFiscal.NumIntDoc, iFilialEmpresa, sChaveNFe)

                                End If
                                Exit For
                            Next


                        End If

                        Exit For

                    Next

                    If gobjApp.iDebug = 1 Then MsgBox("sDir = " & gobjApp.sDirXml)


                    sArquivo = gobjApp.sDirXml & objRetConsSitNFe.chNFe & "-pre.xml"


                    If iRenomearArq = 1 Then
                        If My.Computer.FileSystem.FileExists(gobjApp.sDirXml & objRetConsSitNFe.chNFe & "-pre.xml") Then
                            If My.Computer.FileSystem.FileExists(gobjApp.sDirXml & objRetConsSitNFe.chNFe & "-pre_old.xml") Then My.Computer.FileSystem.DeleteFile(gobjApp.sDirXml & objRetConsSitNFe.chNFe & "-pre_old.xml")
                            My.Computer.FileSystem.RenameFile(gobjApp.sDirXml & objRetConsSitNFe.chNFe & "-pre.xml", objRetConsSitNFe.chNFe & "-pre_old.xml")
                        End If
                    End If


                    'se o arquivo nao existir  ==> cria o arquivo
                    If Dir(sArquivo) <> objRetConsSitNFe.chNFe & "-pre.xml" Then


                        resAdmConfig = gobjApp.dbDadosNfe.ExecuteQuery(Of AdmConfig) _
                        ("SELECT * FROM AdmConfig WHERE  Codigo = {0} ", "VERSAO_MSG")

                        If resAdmConfig.Count > 0 Then

                            resAdmConfig = gobjApp.dbDadosNfe.ExecuteQuery(Of AdmConfig) _
                            ("SELECT * FROM AdmConfig WHERE  Codigo = {0} ", "VERSAO_MSG")

                            objAdmConfig = resAdmConfig(0)
                        Else
                            objAdmConfig = New AdmConfig
                        End If

                        Dim X As New ClassMontaXmlNFe
                        lErro = X.Monta_NFiscal_Xml(a5, objNFeNFiscal, sSerie, 0, colNFiscal, sChaveNFe, XMLStringNFes)
                        If lErro <> SUCESSO Then Throw New System.Exception("Erro na rotina Monta_Fiscal_Xml.")


                    End If


                    If iRenomearArq = 1 Then
                        If My.Computer.FileSystem.FileExists(gobjApp.sDirXml & objRetConsSitNFe.protNFe.infProt.chNFe & "-procNfe.xml") Then
                            If My.Computer.FileSystem.FileExists(gobjApp.sDirXml & objRetConsSitNFe.protNFe.infProt.chNFe & "-procNfe_old.xml") Then My.Computer.FileSystem.DeleteFile(gobjApp.sDirXml & objRetConsSitNFe.protNFe.infProt.chNFe & "-procNfe_old.xml")
                            My.Computer.FileSystem.RenameFile(gobjApp.sDirXml & objRetConsSitNFe.protNFe.infProt.chNFe & "-procNfe.xml", objRetConsSitNFe.protNFe.infProt.chNFe & "-procNfe_old.xml")
                        End If
                    End If


                    'se o arquivo ainda nao existir
                    If ((objRetConsSitNFe.protNFe.infProt.cStat = "100" Or objRetConsSitNFe.protNFe.infProt.cStat = "150") And Dir(gobjApp.sDirXml & objRetConsSitNFe.protNFe.infProt.chNFe & "-procNfe.xml") <> objRetConsSitNFe.protNFe.infProt.chNFe & "-procNfe.xml") Or _
                        ((objRetConsSitNFe.protNFe.infProt.cStat = "205" Or objRetConsSitNFe.protNFe.infProt.cStat = "110" Or objRetConsSitNFe.protNFe.infProt.cStat = "301" Or objRetConsSitNFe.protNFe.infProt.cStat = "302") And _
                         Dir(gobjApp.sDirXml & objRetConsSitNFe.protNFe.infProt.chNFe & "-den.xml") <> objRetConsSitNFe.protNFe.infProt.chNFe & "-den.xml") Then


                        Dim DocDados2 As XmlDocument = New XmlDocument
                        Dim XMLStreamDados1 = New MemoryStream(10000)
                        Dim XMLString As String
                        Dim XMLString3 As String
                        Dim XMLStreamDados = New MemoryStream(10000)

                        sArquivo = gobjApp.sDirXml & objRetConsSitNFe.protNFe.infProt.chNFe & "-pre.xml"
                        DocDados2.Load(sArquivo)
                        DocDados2.Save(XMLStreamDados1)

                        Dim xm As Byte()

                        'pega a parte do xml que fica entre <NFe> e </NFe>
                        xm = XMLStreamDados1.ToArray
                        XMLString1 = System.Text.Encoding.UTF8.GetString(xm)

                        'cria uma versao do que vai ser armazenado somente com o que é possivel, ou seja,
                        'versao e protNFE. O <NFe> vai ser inserido depois (XMLString1)
                        Dim objNFeProc As TNfeProc = New TNfeProc

                        'objNFeProc.NFe = objNFe
                        objNFeProc.versao = NFE_VERSAO_XML
                        objNFeProc.protNFe = New TProtNFe
                        objNFeProc.protNFe.versao = objRetConsSitNFe.protNFe.versao
                        objNFeProc.protNFe.infProt = New TProtNFeInfProt
                        objNFeProc.protNFe.infProt.chNFe = objRetConsSitNFe.protNFe.infProt.chNFe
                        objNFeProc.protNFe.infProt.cStat = objRetConsSitNFe.protNFe.infProt.cStat
                        objNFeProc.protNFe.infProt.dhRecbto = objRetConsSitNFe.protNFe.infProt.dhRecbto
                        objNFeProc.protNFe.infProt.digVal = objRetConsSitNFe.protNFe.infProt.digVal
                        objNFeProc.protNFe.infProt.Id = objRetConsSitNFe.protNFe.infProt.Id
                        objNFeProc.protNFe.infProt.nProt = objRetConsSitNFe.protNFe.infProt.nProt
                        objNFeProc.protNFe.infProt.tpAmb = objRetConsSitNFe.protNFe.infProt.tpAmb
                        objNFeProc.protNFe.infProt.verAplic = objRetConsSitNFe.protNFe.infProt.verAplic
                        objNFeProc.protNFe.infProt.xMotivo = objRetConsSitNFe.protNFe.infProt.xMotivo


                        Dim objProtNFe As TProtNFe = New TProtNFe

                        Dim mySerializer As New XmlSerializer(GetType(TNfeProc))

                        XMLStream1 = New MemoryStream(10000)

                        mySerializer.Serialize(XMLStream1, objNFeProc)

                        Dim xm3 As Byte()
                        xm3 = XMLStream1.ToArray

                        XMLString3 = System.Text.Encoding.UTF8.GetString(xm3)

                        iPos = InStr(XMLString3, "<protNFe")

                        'criacao da string completa
                        XMLString = Mid(XMLString3, 1, iPos - 1) & XMLString1 & Mid(XMLString3, iPos)

                        XMLString = Mid(XMLString, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLString, 20)

                        iPos = InStr(XMLString, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")

                        If iPos <> 0 Then

                            XMLString1 = Mid(XMLString, 1, iPos - 1) & Mid(XMLString, iPos + 99)

                        End If

                        XMLStreamDados = New MemoryStream(10000)

                        Dim xDados1 As Byte()

                        xDados1 = System.Text.Encoding.UTF8.GetBytes(XMLString)

                        XMLStreamDados.Write(xDados1, 0, xDados1.Length)


                        Dim DocDados3 As XmlDocument = New XmlDocument

                        XMLStreamDados.Position = 0
                        DocDados3.Load(XMLStreamDados)

                        If objRetConsSitNFe.protNFe.infProt.cStat = "100" Or objRetConsSitNFe.protNFe.infProt.cStat = "150" Then

                            sArquivo = gobjApp.sDirXml & objRetConsSitNFe.protNFe.infProt.chNFe & "-procNfe.xml"

                        ElseIf objRetConsSitNFe.protNFe.infProt.cStat = "205" Or objRetConsSitNFe.protNFe.infProt.cStat = "110" Or objRetConsSitNFe.protNFe.infProt.cStat = "301" Or objRetConsSitNFe.protNFe.infProt.cStat = "302" Then

                            sArquivo = gobjApp.sDirXml & objRetConsSitNFe.protNFe.infProt.chNFe & "-den.xml"

                        End If

                        Dim writer As New XmlTextWriter(sArquivo, Nothing)

                        writer.Formatting = Formatting.None
                        DocDados3.WriteTo(writer)

                        writer.Close()

                    End If

                End If

                iProcessado = 1

            End If

            If Not objRetConsSitNFe.procEventoNFe Is Nothing Then


                For iIndice = 0 To objRetConsSitNFe.procEventoNFe.Count - 1

                    If (objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.cStat = "135" Or objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.cStat = "136") And objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.tpEvento = "110110" Then

                        iSerie = CInt(Mid(objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.chNFe, 23, 3))
                        sSerie = iSerie & "-e"
                        lNumNotaFiscal = CLng(Mid(objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.chNFe, 26, 9))

                        Dim iAno As Integer
                        Dim iMes As Integer

                        iAno = CInt(Mid(objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.chNFe, 3, 2)) + 2000
                        iMes = CInt(Mid(objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.chNFe, 5, 2))

                        resNFeNFiscal = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeNFiscal) _
                        ("SELECT * FROM NFeNFiscal WHERE Serie = {0} AND NumNotaFiscal = {1} AND FilialEmpresa = {2} AND YEAR(DataEmissao) = {3} AND MONTH(DataEmissao) = {4}", sSerie, lNumNotaFiscal, iFilialEmpresa, iAno, iMes)

                        For Each objNFeNFiscal In resNFeNFiscal

                            resNFeFedRetEnvCCe = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedRetEnvCCe) _
                            ("SELECT * FROM NFeFedRetEnvCCe WHERE chNFe = {0} AND nSeqEvento = {1}", objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.chNFe, objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.nSeqEvento)

                            If resNFeFedRetEnvCCe.Count = 0 Then

                                iResult = gobjApp.dbDadosNfe.ExecuteCommand("INSERT INTO NFeFedRetEnvCCe ( FilialEmpresa, NumIntNF, versao, idLote, tpAmb, verAplic, cOrgao, xMotivo, Id, tpAmb1, verAplic1, cOrgao1, cStat1, xMotivo1, chNFe, tpEvento, xEvento, nSeqEvento, CPFCNPJ, email, dataRegEvento, horaRegEvento, nProt, cStat) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23} )", _
                                iFilialEmpresa, objNFeNFiscal.NumIntDoc, objRetConsSitNFe.versao, 0, objRetConsSitNFe.tpAmb, objRetConsSitNFe.verAplic, objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.cOrgao, objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.xMotivo, IIf(objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.Id Is Nothing, "", objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.Id), objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.tpAmb, objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.verAplic, objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.cOrgao, objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.cStat, objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.xMotivo, objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.chNFe, _
                                objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.tpEvento, IIf(objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.xEvento Is Nothing, "", objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.xEvento), objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.nSeqEvento, IIf(objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.Item Is Nothing, "", objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.Item), IIf(objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.emailDest Is Nothing, "", objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.emailDest), IIf(objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.dhRegEvento Is Nothing, Now.Date, CDate(objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.dhRegEvento).Date), IIf(objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.dhRegEvento Is Nothing, TimeOfDay.ToOADate, CDate(objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.dhRegEvento).TimeOfDay.TotalDays), IIf(objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.nProt Is Nothing, "", objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.nProt), objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.cStat)

                            End If

                            Form1.Msg.Items.Add("Retorno do evento " & IIf(objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.xEvento Is Nothing, "carta de correção", objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.xEvento) & " nSeqEvento = " & objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.nSeqEvento & " chave NFe = " & objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.chNFe)
                            Form1.Msg.Items.Add("cStat = " & objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.cStat & " xMotivo1 = " & objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.xMotivo)


                            sArquivo = "110110-" & objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.chNFe & "-" & objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.nSeqEvento & "-procEventoNfe.xml"

                            If iRenomearArq = 1 Then
                                If My.Computer.FileSystem.FileExists(gobjApp.sDirXml & "110110-" & objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.chNFe & "-" & objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.nSeqEvento & "-procEventoNfe.xml") Then
                                    If My.Computer.FileSystem.FileExists(gobjApp.sDirXml & "110110-" & objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.chNFe & "-" & objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.nSeqEvento & "-procEventoNfe_old.xml") Then My.Computer.FileSystem.DeleteFile(gobjApp.sDirXml & objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.chNFe & "-" & objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.nSeqEvento & "-procEventoNfe_old.xml")
                                    My.Computer.FileSystem.RenameFile(gobjApp.sDirXml & "110110-" & objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.chNFe & "-" & objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.nSeqEvento & "-procEventoNfe.xml", "110110-" & objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.chNFe & "-" & objRetConsSitNFe.procEventoNFe(iIndice).retEvento.infEvento.nSeqEvento & "-procEventoNfe_old.xml")
                                End If
                            End If



                            'se o arquivo ainda nao existir
                            If Dir(gobjApp.sDirXml & sArquivo) <> sArquivo Then

                                Call gobjApp.Evento_Grava_Xml(objRetConsSitNFe.procEventoNFe(iIndice))

                            End If

                            Exit For

                        Next

                    End If

                Next

                iProcessado = 1

            End If


            'se for uma nota denegada vai guardar a informacao em NFeFedDenegada
            If (objRetConsSitNFe.cStat = "205" Or objRetConsSitNFe.cStat = "110" Or objRetConsSitNFe.cStat = "301" Or objRetConsSitNFe.cStat = "302") Then

                iSerie = CInt(Mid(sChaveNFe, 23, 3))
                sSerie = iSerie & "-e"
                lNumNotaFiscal = CLng(Mid(sChaveNFe, 26, 9))

                Dim iAno As Integer
                Dim iMes As Integer

                iAno = CInt(Mid(sChaveNFe, 3, 2)) + 2000
                iMes = CInt(Mid(sChaveNFe, 5, 2))

                results = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeNFiscal) _
                ("SELECT * FROM NFeNFiscal WHERE Serie = {0} AND NumNotaFiscal = {1} AND FilialEmpresa = {2} AND YEAR(DataEmissao) = {3} AND MONTH(DataEmissao) = {4}", sSerie, lNumNotaFiscal, iFilialEmpresa, iAno, iMes)

                For Each objNFeNFiscal In results

                    resNFeFedDenegada = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedDenegada) _
                    ("SELECT * FROM NFeFedDenegada WHERE NumIntNF = {0} ", objNFeNFiscal.NumIntDoc)

                    If resNFeFedDenegada.Count = 0 Then

                        iResult = gobjApp.dbDadosNfe.ExecuteCommand("INSERT INTO NFeFedDenegada ( NumIntNF, Processada) VALUES ( {0}, {1})", _
                        objNFeNFiscal.NumIntDoc, 0)

                    End If


                    Exit For

                Next

                iProcessado = 1

            End If

            If Not objRetConsSitNFe.procEventoNFe Is Nothing Then

                For Each objprocEventoNFe In objRetConsSitNFe.procEventoNFe

                    If objprocEventoNFe.retEvento.infEvento.tpEvento = "110111" Then

                        iSerie = CInt(Mid(sChaveNFe, 23, 3))
                        sSerie = iSerie & "-e"
                        lNumNotaFiscal = CLng(Mid(sChaveNFe, 26, 9))

                        Dim iAno As Integer
                        Dim iMes As Integer

                        iAno = CInt(Mid(sChaveNFe, 3, 2)) + 2000
                        iMes = CInt(Mid(sChaveNFe, 5, 2))

                        '??? deveria usar tb a data de emissao
                        results = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeNFiscal) _
                        ("SELECT * FROM NFeNFiscal WHERE Serie = {0} AND NumNotaFiscal = {1} AND FilialEmpresa = {2} AND YEAR(DataEmissao) = {3} AND MONTH(DataEmissao) = {4}", sSerie, lNumNotaFiscal, iFilialEmpresa, iAno, iMes)

                        For Each objNFeNFiscal In results

                            If gobjApp.iDebug = 1 Then MsgBox("33")

                            sErro1 = ""

                            If iRenomearArq = 1 Then
                                If My.Computer.FileSystem.FileExists(gobjApp.sDirXml & "110111-" & sChaveNFe & "-1-procEventoNfe.xml") Then
                                    If My.Computer.FileSystem.FileExists(gobjApp.sDirXml & "110111-" & sChaveNFe & "-1-procEventoNfe_old.xml") Then My.Computer.FileSystem.DeleteFile(gobjApp.sDirXml & "110111-" & sChaveNFe & "-1-procEventoNfe_old.xml")
                                    My.Computer.FileSystem.RenameFile(gobjApp.sDirXml & "110111-" & sChaveNFe & "-1-procEventoNfe.xml", "110111-" & objRetConsSitNFe.chNFe & "-1-procEventoNfe_old.xml")
                                End If
                            End If

                            'se o arquivo ainda nao existir
                            '                            If Not My.Computer.FileSystem.FileExists(gobjApp.sDirXml & "110111-" & sChaveNFe & "-1-procEventoNfe.xml") Then

                            lErro = ClassCancelaNFe.Cancelamento_NFe_Gravar_XML(iFilialEmpresa, gobjApp.sDirXml, objprocEventoNFe, sErro1, dbDadosNfe, objNFeNFiscal.NumIntDoc)
                            If lErro <> 0 Then Throw New System.Exception(sErro1)

                            '                            End If

                            Exit For

                        Next

                    End If

                Next

                If gobjApp.iDebug = 1 Then MsgBox("35")

                iProcessado = 1

            End If

            If gobjApp.iDebug = 1 Then MsgBox("50")

            If Not objNFeNFiscal Is Nothing And iOperacao = 0 Then

                resNFeFedRetCancNfe = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedRetCancNFe) _
                ("SELECT * FROM NFeFedRetCancNFe WHERE NumIntNF = {0} AND (cStat = 101 Or cStat = 151 Or cStat = 135 Or cStat = 155)", objNFeNFiscal.NumIntDoc)

                If resNFeFedRetCancNfe.Count > 0 Then

                    Call gobjApp.GravarLog("A NOTA FISCAL " & objNFeNFiscal.NumNotaFiscal & " TEM O SEU CANCELAMENTO HOMOLOGADO.", 0, objNFeNFiscal.NumIntDoc)

                    iProcessado = 1

                Else

                    resNFeFedProtNFe = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedProtNFe) _
                    ("SELECT * FROM NFeFedProtNFe WHERE (cStat = '205' Or cStat = '110' Or cStat = '301' Or cStat = '302') AND FilialEmpresa = {0} AND NumIntNF = {1}", iFilialEmpresa, objNFeNFiscal.NumIntDoc)

                    If resNFeFedProtNFe.Count > 0 Then

                        Call gobjApp.GravarLog("A NOTA FISCAL " & objNFeNFiscal.NumNotaFiscal & " FOI DENEGADA.", 0, objNFeNFiscal.NumIntDoc)

                        iProcessado = 1

                    Else

                        resNFeFedProtNFe = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedProtNFe) _
                        ("SELECT * FROM NFeFedProtNFe WHERE (cStat = '100' Or cStat = '150') AND FilialEmpresa = {0} AND NumIntNF = {1}", iFilialEmpresa, objNFeNFiscal.NumIntDoc)

                        If resNFeFedProtNFe.Count > 0 Then

                            Call gobjApp.GravarLog("A NOTA FISCAL " & objNFeNFiscal.NumNotaFiscal & " ESTA AUTORIZADA.", 0, objNFeNFiscal.NumIntDoc)

                            iProcessado = 1

                        End If

                    End If

                End If

                'se for uma consulta seguindo um cancelamento
            ElseIf Not objNFeNFiscal Is Nothing And iOperacao = 1 Then

                resNFeFedRetCancNfe = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedRetCancNFe) _
                ("SELECT * FROM NFeFedRetCancNFe WHERE NumIntNF = {0} AND (cStat = 101 Or cStat = 151 Or cStat = 135 Or cStat = 155)", objNFeNFiscal.NumIntDoc)

                If resNFeFedRetCancNfe.Count > 0 Then

                    Call gobjApp.GravarLog("A NOTA FISCAL " & objNFeNFiscal.NumNotaFiscal & " TEM O SEU CANCELAMENTO HOMOLOGADO.", 0, objNFeNFiscal.NumIntDoc)

                    iProcessado = 1

                Else

                    Call gobjApp.GravarLog("O CANCELAMENTO DA NOTA FISCAL " & objNFeNFiscal.NumNotaFiscal & " AINDA NÃO FOI HOMOLOGADO. TENTE REFAZER OU CONSULTAR A NFE MAIS TARDE.", 0, objNFeNFiscal.NumIntDoc)

                    iProcessado = 1

                End If

            End If

            '            If objRetConsSitNFe.cStat = "217" Then
            If iProcessado = 0 And Not objRetConsSitNFe Is Nothing Then
                Throw New System.Exception(" cStat = " & objRetConsSitNFe.cStat & " motivo = " & objRetConsSitNFe.xMotivo)
            End If


            Call gobjApp.GravarLog("Consulta da chave " & sChaveNFe & " realizada com sucesso", 0, 0)

            gobjApp.DadosCommit()

            Consulta_NFe = SUCESSO

        Catch ex As Exception

            Consulta_NFe = 1

            Call gobjApp.GravarLog("ERRO - a consulta a chave " & sChaveNFe & " foi encerrado por erro. Erro = " & ex.Message, 0, 0)

            Call gobjApp.GravarLog(gobjApp.sErro & " " & gobjApp.sMsg1, 0, 0)

            Call gobjApp.GravarLog("ERRO - Encerrada a consulta da chave", 0, 0)

        Finally

            If gobjApp.iDebug = 1 Then Call MsgBox(gobjApp.sErro)

        End Try

    End Function

End Class


