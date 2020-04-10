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
Imports sgenfebd4.CCEXsd

Public Class ClassCCeNFe

    Public Function CCe_NFe(ByVal sEmpresa As String, ByVal iFilialEmpresa As Integer, ByVal lIdLote As Long, Optional ByVal iScan As Integer = -1) As Long

        Dim NfeRecEvento As New recepcaoevento.NFeRecepcaoEvento4
        Dim NfeRecEventoP As New recepcaoevento.NFeRecepcaoEvento4

        Dim iResult As Integer
        Dim XMLStream1 As MemoryStream = New MemoryStream(10000)
        'Dim XMLStreamCabec As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamRet As MemoryStream = New MemoryStream(10000)


        Dim XMLString1 As String

        Dim objNFeNFiscal As NFeNFiscal
        Dim objNFeFedLote As NFeFedLote = New NFeFedLote
        Dim objNFeFedRetEnvi As NFeFedRetEnvi = New NFeFedRetEnvi
        Dim objNFeFedProtNFe As NFeFedProtNFe = New NFeFedProtNFe
        Dim xRet As Byte()
        'Dim mySerializercabec As New XmlSerializer(GetType(recepcaoevento.nfeCabecMsg))

        Dim colNumIntNFiscal As Collection = New Collection

        'Dim objCabecMsg As recepcaoevento.nfeCabecMsg = New recepcaoevento.nfeCabecMsg

        Dim iPos As Integer
        Dim lErro As Long


        Dim sArquivo As String

        Dim colNFiscal As Collection = New Collection

        Dim xmlNode1 As XmlNode

        Dim objEnvEvento As TEnvEvento = New TEnvEvento
        Dim sUF As String
        Dim AD As AssinaturaDigital = New AssinaturaDigital
        Dim XMLStreamDados As MemoryStream = New MemoryStream(10000)
        Dim objValidaXML As ClassValidaXML = New ClassValidaXML
        Dim objNFeFedRetEnvCCe As NFeFedRetEnvCCe = New NFeFedRetEnvCCe
        Dim resNFeFedLoteCCe As IEnumerable(Of NFeFedLoteCce)
        Dim resCCe As IEnumerable(Of Cce)

        Dim objCCe As Cce
        Dim objNFeFedLoteCCe As NFeFedLoteCce

        Dim iIndice As Integer

        Dim objEvento As CCEXsd.TEvento
        Dim XMLStringEventos As String
        Dim XMLString2 As String
        Dim iSerie As Integer
        Dim sSerie As String
        Dim lNumNotaFiscal As Long
        Dim results As IEnumerable(Of NFeNFiscal)
        Dim xMLStringEvento(20) As String
        Dim objProcEvento As CCEXsd.TProcEvento
        Dim aobjEvento(20) As CCEXsd.TEvento

        gobjApp.sMsg1 = ""
        gobjApp.sErro = ""

        Try
            XMLStringEventos = ""

            If gobjApp.iDebug = 1 Then gobjApp.sErro = "6"

            gobjApp.sErro = "6"
            gobjApp.sMsg1 = "vai inserir na tabela NFeFedLoteLog"

            Call gobjApp.GravarLog("Iniciando o envio do evento de carta de correção eletronica", 0, 0)

            iIndice = -1

            resNFeFedLoteCCe = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedLoteCce) _
            ("SELECT * FROM NFeFedLoteCce WHERE FilialEmpresa = {0} AND idLote = {1}", iFilialEmpresa, lIdLote)

            For Each objNFeFedLoteCCe In resNFeFedLoteCCe

                resCCe = gobjApp.dbDadosNfe.ExecuteQuery(Of Cce) _
                ("SELECT * FROM Cce WHERE NumIntDoc = {0} ", objNFeFedLoteCCe.NumIntCce)

                For Each objCCe In resCCe



                    iIndice = iIndice + 1

                    objEvento = New CCEXsd.TEvento

                    objEvento.versao = "1.00"

                    objEvento.infEvento = New CCEXsd.TEventoInfEvento

                    objEvento.infEvento.Id = "ID" & "110110" & objCCe.chNFe & Format(objCCe.nSeqEvento, "00")

                    sUF = Left(objCCe.chNFe, 2)

                    objEvento.infEvento.cOrgao = GetCode(Of CCEXsd.TCOrgaoIBGE)(sUF)


                    If gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then
                        objEvento.infEvento.tpAmb = CCEXsd.TAmb.Item2
                    Else
                        objEvento.infEvento.tpAmb = CCEXsd.TAmb.Item1
                    End If

                    objEvento.infEvento.ItemElementName = CCEXsd.ItemChoiceType.CNPJ
                    objEvento.infEvento.Item = Mid(objCCe.chNFe, 7, 14)

                    objEvento.infEvento.chNFe = objCCe.chNFe


                    objEvento.infEvento.dhEvento = Format(Now.Date, "yyyy-MM-dd") & "T" & Format(TimeOfDay, "HH:mm:ss") & Format(TimeOfDay, "zzz")

                    objEvento.infEvento.tpEvento = TEventoInfEventoTpEvento.Item110110

                    objEvento.infEvento.nSeqEvento = objCCe.nSeqEvento

                    objEvento.infEvento.verEvento = TEventoInfEventoVerEvento.Item100

                    objEvento.infEvento.detEvento = New CCEXsd.TEventoInfEventoDetEvento

                    objEvento.infEvento.detEvento.versao = TEventoInfEventoDetEventoVersao.Item100

                    objEvento.infEvento.detEvento.descEvento = TEventoInfEventoDetEventoDescEvento.CartadeCorreção

                    objEvento.infEvento.detEvento.xCorrecao = Trim(DesacentuaTexto(objCCe.Correcao))

                    objEvento.infEvento.detEvento.xCondUso = TEventoInfEventoDetEventoXCondUso.ACartadeCorreçãoédisciplinadapelo1ºAdoart7ºdoConvênioSNde15dedezembrode1970epodeserutilizadapararegularizaçãodeerroocorridonaemissãodedocumentofiscaldesdequeoerronãoestejarelacionadocomIasvariáveisquedeterminamovalordoimpostotaiscomobasedecálculoalíquotadiferençadepreçoquantidadevalordaoperaçãooudaprestaçãoIIacorreçãodedadoscadastraisqueimpliquemudançadoremetenteoudodestinatárioIIIadatadeemissãooudesaída

                    If gobjApp.iDebug = 1 Then gobjApp.sErro = "20"
                    gobjApp.sErro = "20"
                    gobjApp.sMsg1 = "vai serializar TEvento"

                    Dim mySerializer As New XmlSerializer(GetType(CCEXsd.TEvento))

                    Dim XMLStream2 = New MemoryStream(10000)
                    mySerializer.Serialize(XMLStream2, objEvento)

                    Dim xm2 As Byte()
                    xm2 = XMLStream2.ToArray

                    XMLString1 = System.Text.Encoding.UTF8.GetString(xm2)

                    '                    XMLString1 = Mid(XMLString1, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLString1, 20)

                    iPos = InStr(XMLString1, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")

                    If iPos <> 0 Then

                        XMLString1 = Mid(XMLString1, 1, iPos - 1) & Mid(XMLString1, iPos + 99)

                    End If



                    lErro = AD.Assinar(XMLString1, "infEvento", gobjApp.cert, gobjApp.iDebug)
                    If lErro <> SUCESSO Then Throw New System.Exception("")




                    Dim xMlD As XmlDocument

                    xMlD = AD.XMLDocAssinado()

                    Dim xString As String
                    xString = AD.XMLStringAssinado


                    XMLStringEventos = XMLStringEventos & Mid(xString, 22)

                    xRet = System.Text.Encoding.UTF8.GetBytes(xString)

                    XMLStreamRet = New MemoryStream(10000)
                    XMLStreamRet.Write(xRet, 0, xRet.Length)

                    Dim mySerializerEvento As New XmlSerializer(GetType(CCEXsd.TEvento))


                    XMLStreamRet.Position = 0

                    aobjEvento(iIndice) = New CCEXsd.TEvento

                    aobjEvento(iIndice) = mySerializerEvento.Deserialize(XMLStreamRet)

                    Exit For

                Next

            Next


            objEnvEvento.idLote = lIdLote
            objEnvEvento.versao = "1.00"

            Dim mySerializerw As New XmlSerializer(GetType(TEnvEvento))

            XMLStream1 = New MemoryStream(10000)

            mySerializerw.Serialize(XMLStream1, objEnvEvento)

            Dim xmw As Byte()
            xmw = XMLStream1.ToArray

            XMLString1 = System.Text.Encoding.UTF8.GetString(xmw)

            XMLString2 = Mid(XMLString1, 1, Len(XMLString1) - 12) & XMLStringEventos & Mid(XMLString1, Len(XMLString1) - 12)

            XMLString2 = Mid(XMLString2, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLString2, 20)

            iPos = InStr(XMLString2, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")

            If iPos <> 0 Then

                XMLString2 = Mid(XMLString2, 1, iPos - 1) & Mid(XMLString2, iPos + 99)

            End If

            '************* valida dados antes do envio **********************
            Dim xDados As Byte()

            xDados = System.Text.Encoding.UTF8.GetBytes(XMLString2)

            XMLStreamDados = New MemoryStream(10000)

            XMLStreamDados.Write(xDados, 0, xDados.Length)

            If gobjApp.iDebug = 1 Then MsgBox("21")
            gobjApp.sErro = "21"
            gobjApp.sMsg1 = "vai gravar o xml"


            Dim DocDados As XmlDocument = New XmlDocument
            XMLStreamDados.Position = 0
            DocDados.Load(XMLStreamDados)
            sArquivo = gobjApp.sDirXml & objEnvEvento.idLote & "-env-evento.xml"
            '            DocDados.Save(sArquivo)
            Dim writer1 As New XmlTextWriter(sArquivo, Nothing)

            writer1.Formatting = Formatting.None
            DocDados.WriteTo(writer1)
            writer1.Close()


            If gobjApp.iDebug = 1 Then MsgBox("22")
            gobjApp.sErro = "22"
            gobjApp.sMsg1 = "vai validar o arquivo xml de envio de evento"

            lErro = objValidaXML.validaXML(sArquivo, gobjApp.sDirXsd & "envCCe_v1.00.xsd", 0, 0, iFilialEmpresa)
            If lErro = 1 Then

                Call gobjApp.GravarLog("ERRO - o envio do evento foi encerrado por erro.", 0, 0)

                Exit Try
            End If

            If gobjApp.iDebug = 1 Then MsgBox("23")
            gobjApp.sErro = "23"
            gobjApp.sMsg1 = "vai enviar o evento"

            Dim DocDados1 As New XmlDocument


            Call Salva_Arquivo(DocDados1, XMLString2)

            'NfeCabec.cUF = CStr(gobjApp.objEstado.CodIBGE)
            'NfeCabec.versaoDados = "1.00"

            'NfeRecEvento.nfeCabecMsgValue = NfeCabec

            Dim sURL As String
            sURL = ""
            Call WS_Obter_URL(sURL, gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO, gobjApp.objEstado.Sigla, "RecepcaoEvento", "NFe")

            NfeRecEvento.Url = sURL

            Dim XMLStringRetEnvCCE As String

            NfeRecEvento.ClientCertificates.Add(gobjApp.cert)
            xmlNode1 = NfeRecEvento.nfeRecepcaoEvento(DocDados1)

            XMLStringRetEnvCCE = xmlNode1.OuterXml

            xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringRetEnvCCE)

            XMLStreamRet = New MemoryStream(10000)
            XMLStreamRet.Write(xRet, 0, xRet.Length)

            Dim mySerializerRetEnvCCE As New XmlSerializer(GetType(TRetEnvEvento))

            Dim objRetEnvCCE As TRetEnvEvento = New TRetEnvEvento

            XMLStreamRet.Position = 0

            objRetEnvCCE = mySerializerRetEnvCCE.Deserialize(XMLStreamRet)

            If gobjApp.iDebug = 1 Then gobjApp.sErro = "24"
            gobjApp.sErro = "24"
            gobjApp.sMsg1 = "trata o retorno da carta de correcao"


            '            objNFeFedRetEnvCCe = New NFeFedRetEnvCCe

            If Not objRetEnvCCE.retEvento Is Nothing Then

                For iIndice = 0 To objRetEnvCCE.retEvento.Count - 1


                    iSerie = CInt(Mid(objRetEnvCCE.retEvento(iIndice).infEvento.chNFe, 23, 3))
                    sSerie = iSerie & "-e"
                    lNumNotaFiscal = CLng(Mid(objRetEnvCCE.retEvento(iIndice).infEvento.chNFe, 26, 9))

                    results = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeNFiscal) _
                    ("SELECT * FROM NFeNFiscal WHERE Serie = {0} AND NumNotaFiscal = {1} AND FilialEmpresa = {2}", sSerie, lNumNotaFiscal, iFilialEmpresa)

                    For Each objNFeNFiscal In results


                        iResult = gobjApp.dbDadosNfe.ExecuteCommand("INSERT INTO NFeFedRetEnvCCe ( FilialEmpresa, NumIntNF, versao, idLote, tpAmb, verAplic, cOrgao, xMotivo, Id, tpAmb1, verAplic1, cOrgao1, cStat1, xMotivo1, chNFe, tpEvento, xEvento, nSeqEvento, CPFCNPJ, email, dataRegEvento, horaRegEvento, nProt, cStat) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23} )", _
                        iFilialEmpresa, objNFeNFiscal.NumIntDoc, objRetEnvCCE.versao, objRetEnvCCE.idLote, objRetEnvCCE.tpAmb, objRetEnvCCE.verAplic, objRetEnvCCE.cOrgao, objRetEnvCCE.xMotivo, IIf(objRetEnvCCE.retEvento(iIndice).infEvento.Id Is Nothing, "", objRetEnvCCE.retEvento(iIndice).infEvento.Id), objRetEnvCCE.retEvento(iIndice).infEvento.tpAmb, objRetEnvCCE.retEvento(iIndice).infEvento.verAplic, objRetEnvCCE.retEvento(iIndice).infEvento.cOrgao, objRetEnvCCE.retEvento(iIndice).infEvento.cStat, objRetEnvCCE.retEvento(iIndice).infEvento.xMotivo, objRetEnvCCE.retEvento(iIndice).infEvento.chNFe, _
                        objRetEnvCCE.retEvento(iIndice).infEvento.tpEvento, IIf(objRetEnvCCE.retEvento(iIndice).infEvento.xEvento Is Nothing, "", objRetEnvCCE.retEvento(iIndice).infEvento.xEvento), objRetEnvCCE.retEvento(iIndice).infEvento.nSeqEvento, IIf(objRetEnvCCE.retEvento(iIndice).infEvento.Item Is Nothing, "", objRetEnvCCE.retEvento(iIndice).infEvento.Item), IIf(objRetEnvCCE.retEvento(iIndice).infEvento.emailDest Is Nothing, "", objRetEnvCCE.retEvento(iIndice).infEvento.emailDest), IIf(objRetEnvCCE.retEvento(iIndice).infEvento.dhRegEvento Is Nothing, Now.Date, CDate(objRetEnvCCE.retEvento(iIndice).infEvento.dhRegEvento).Date), IIf(objRetEnvCCE.retEvento(iIndice).infEvento.dhRegEvento Is Nothing, TimeOfDay.ToOADate, CDate(objRetEnvCCE.retEvento(iIndice).infEvento.dhRegEvento).TimeOfDay.TotalDays), IIf(objRetEnvCCE.retEvento(iIndice).infEvento.nProt Is Nothing, "", objRetEnvCCE.retEvento(iIndice).infEvento.nProt), objRetEnvCCE.retEvento(iIndice).infEvento.cStat)

                        Form1.Msg.Items.Add("Retorno do evento " & IIf(objRetEnvCCE.retEvento(iIndice).infEvento.xEvento Is Nothing, "carta de correção", objRetEnvCCE.retEvento(iIndice).infEvento.xEvento) & " nSeqEvento = " & objRetEnvCCE.retEvento(iIndice).infEvento.nSeqEvento & " chave NFe = " & objRetEnvCCE.retEvento(iIndice).infEvento.chNFe)
                        Form1.Msg.Items.Add("cStat = " & objRetEnvCCE.retEvento(iIndice).infEvento.cStat & " xMotivo1 = " & objRetEnvCCE.retEvento(iIndice).infEvento.xMotivo)

                        'se o evento tiver sido registrado
                        If objRetEnvCCE.retEvento(iIndice).infEvento.cStat = 135 Or objRetEnvCCE.retEvento(iIndice).infEvento.cStat = 136 Then

                            objProcEvento = New CCEXsd.TProcEvento

                            objProcEvento.versao = "1.00"

                            objProcEvento.evento = aobjEvento(iIndice)

                            objProcEvento.retEvento = objRetEnvCCE.retEvento(iIndice)

                            Dim mySerializerProcEvento As New XmlSerializer(GetType(CCEXsd.TProcEvento))

                            XMLStream1 = New MemoryStream(10000)

                            mySerializerProcEvento.Serialize(XMLStream1, objProcEvento)

                            Dim xmw1 As Byte()
                            Dim XMLStringProc As String

                            xmw1 = XMLStream1.ToArray

                            XMLStringProc = System.Text.Encoding.UTF8.GetString(xmw1)

                            XMLStringProc = Mid(XMLStringProc, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLStringProc, 20)

                            iPos = InStr(XMLStringProc, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")

                            If iPos <> 0 Then

                                XMLStringProc = Mid(XMLStringProc, 1, iPos - 1) & Mid(XMLStringProc, iPos + 99)

                            End If


                            Dim xDadosCce As Byte()

                            xDadosCce = System.Text.Encoding.UTF8.GetBytes(XMLStringProc)

                            XMLStreamDados = New MemoryStream(10000)

                            XMLStreamDados.Write(xDadosCce, 0, xDadosCce.Length)

                            If gobjApp.iDebug = 1 Then MsgBox("25")
                            gobjApp.sErro = "25"
                            gobjApp.sMsg1 = "vai gravar o xml"


                            Dim DocDadosCce As XmlDocument = New XmlDocument
                            XMLStreamDados.Position = 0
                            DocDadosCce.Load(XMLStreamDados)
                            sArquivo = gobjApp.sDirXml & objRetEnvCCE.retEvento(iIndice).infEvento.tpEvento.ToString & "-" & objRetEnvCCE.retEvento(iIndice).infEvento.chNFe & "-" & objRetEnvCCE.retEvento(iIndice).infEvento.nSeqEvento & "-procEventoNfe.xml"

                            Dim writer As New XmlTextWriter(sArquivo, Nothing)

                            writer.Formatting = Formatting.None
                            DocDadosCce.WriteTo(writer)
                            writer.Close()


                        End If


                        Exit For

                    Next

                Next

            Else

                iResult = gobjApp.dbDadosNfe.ExecuteCommand("INSERT INTO NFeFedRetEnvCCe ( FilialEmpresa, NumIntNF, versao, idLote, tpAmb, verAplic, cOrgao, xMotivo) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7} )", _
                iFilialEmpresa, 0, objRetEnvCCE.versao, objRetEnvCCE.idLote, objRetEnvCCE.tpAmb, objRetEnvCCE.verAplic, objRetEnvCCE.cOrgao, objRetEnvCCE.xMotivo)

                Form1.Msg.Items.Add("Ocorreu um erro no envio do evento, cStat = " & objRetEnvCCE.cStat & " xMotivo = " & objRetEnvCCE.xMotivo)

            End If

            If gobjApp.iDebug = 1 Then MsgBox("26")
            gobjApp.sErro = "26"
            gobjApp.sMsg1 = "vai gravar o xml"


            Dim DocDados2 As XmlDocument = New XmlDocument
            XMLStreamRet.Position = 0
            DocDados2.Load(XMLStreamRet)
            sArquivo = gobjApp.sDirXml & objRetEnvCCE.idLote & "-ret-env-evento.xml"
            DocDados2.Save(sArquivo)

            gobjApp.DadosCommit()

            Form1.Msg.Items.Add("Envio do evento encerrado.")

            CCe_NFe = 1

        Catch ex As Exception

            CCe_NFe = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

            Call gobjApp.GravarLog(gobjApp.sErro & " " & gobjApp.sMsg1, 0, 0)

            Call gobjApp.GravarLog("ERRO - Encerrado o envio do evento", 0, 0)

        Finally

        End Try

    End Function

End Class

