Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml
Imports System.Xml.Schema
Imports System.Data.Odbc
Imports sgenfebd4.NFeXsd
Imports sgenfebd4.CancXsd
Imports sgenfebd4.EventoXsd

Public Class ClassCancelaNFe

    Public Shared Function Cancelamento_NFe_Gravar_XML(ByVal iFilialEmpresa As Integer, ByVal sDir As String, ByVal objprocEventoNFe As NFeXsd.TProcEvento, ByRef sErro As String, ByVal dbDadosNfe As DataClassesDataContext, ByVal lNumIntNF As Long) As Long

        Dim DocDados2 As XmlDocument = New XmlDocument
        Dim XMLStreamDados1 = New MemoryStream(10000)
        Dim XMLStreamDados2 = New MemoryStream(10000)
        Dim sArquivo As String
        Dim iPos As Integer
        Dim XMLStreamDados = New MemoryStream(10000)
        Dim iResult As Integer
        Dim resNFeFedRecCancNFe As IEnumerable(Of NFeFedRetCancNFe)

        Try

            If gobjApp.iDebug = 1 Then MsgBox("101")


            resNFeFedRecCancNFe = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedRetCancNFe) _
            ("SELECT * FROM NFeFedRetCancNFe WHERE NumIntNF = {0} AND data = {1} AND ABS(hora - {2}) < 0.001", lNumIntNF, UTCParaDate(objprocEventoNFe.retEvento.infEvento.dhRegEvento), CDate(objprocEventoNFe.retEvento.infEvento.dhRegEvento).TimeOfDay.TotalDays)

            If resNFeFedRecCancNFe.Count = 0 Then

                If gobjApp.iDebug = 1 Then MsgBox("102")

                iResult = gobjApp.dbDadosNfe.ExecuteCommand("INSERT INTO NFeFedRetCancNFe ( FilialEmpresa, NumIntNF, versao, tpAmb, verAplic, cStat, xMotivo, cUF, chNFe, nProt, Data, Hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, (9), {10}, {11} )", _
                    iFilialEmpresa, lNumIntNF, objprocEventoNFe.retEvento.versao, objprocEventoNFe.retEvento.infEvento.tpAmb, objprocEventoNFe.retEvento.infEvento.verAplic, objprocEventoNFe.retEvento.infEvento.cStat, IIf(objprocEventoNFe.retEvento.infEvento.xMotivo Is Nothing, "", objprocEventoNFe.retEvento.infEvento.xMotivo), objprocEventoNFe.retEvento.infEvento.cOrgao, objprocEventoNFe.retEvento.infEvento.chNFe, IIf(objprocEventoNFe.retEvento.infEvento.nProt = "", "0", objprocEventoNFe.retEvento.infEvento.nProt), IIf(objprocEventoNFe.retEvento.infEvento.dhRegEvento = Nothing, Now.Date, UTCParaDate(objprocEventoNFe.retEvento.infEvento.dhRegEvento)), CDate(objprocEventoNFe.retEvento.infEvento.dhRegEvento).TimeOfDay.TotalDays)

            End If

            If gobjApp.iDebug = 1 Then MsgBox("103")

            'se o arquivo ainda nao existir
            If Not My.Computer.FileSystem.FileExists(gobjApp.sDirXml & "110111-" & objprocEventoNFe.evento.infEvento.chNFe & "-1-procEventoNfe.xml") Then

                Call gobjApp.Evento_Grava_Xml(objprocEventoNFe)

            End If

            Cancelamento_NFe_Gravar_XML = SUCESSO

        Catch ex As Exception

            Cancelamento_NFe_Gravar_XML = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        Finally

        End Try

    End Function

    Public Function Evento_Cancela_NFe(ByVal sEmpresa As String, ByVal lNumIntNF As Long, ByVal sMotivo As String, ByVal iFilialEmpresa As Integer, ByRef sChNFe As String, Optional ByVal iScan As Integer = -1, Optional ByVal sSistemaContingencia As String = "") As Long

        Dim iResult As Integer
        Dim XMLStream1 As MemoryStream = New MemoryStream(10000)
        'Dim XMLStreamCabec As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamRet As MemoryStream = New MemoryStream(10000)


        Dim XMLString1 As String

        Dim resNFeNFiscal As IEnumerable(Of NFeNFiscal)
        Dim resNFeFedProtNFE As IEnumerable(Of NFeFedProtNFe)

        Dim objNFeNFiscal As New NFeNFiscal
        Dim objNFeFedLote As NFeFedLote = New NFeFedLote
        Dim objNFeFedRetEnvi As NFeFedRetEnvi = New NFeFedRetEnvi
        Dim objNFeFedProtNFe As NFeFedProtNFe = New NFeFedProtNFe
        Dim xRet As Byte()
        'Dim mySerializercabec As New XmlSerializer(GetType(recepcaoevento.nfeCabecMsg))

        Dim objDescEvento As TEventoInfEventoDetEventoDescEvento

        Dim colNumIntNFiscal As Collection = New Collection

        'Dim objCabecMsg As recepcaoevento.nfeCabecMsg = New recepcaoevento.nfeCabecMsg

        Dim iPos As Integer
        Dim lErro As Long


        Dim sArquivo As String

        Dim colNFiscal As Collection = New Collection

        Dim xmlNode1 As XmlNode


        Dim objEnvCancEvento As CancXsd.TEnvEvento = New CancXsd.TEnvEvento
        Dim sUF As String
        Dim AD As AssinaturaDigital = New AssinaturaDigital
        Dim XMLStreamDados As MemoryStream = New MemoryStream(10000)
        Dim objValidaXML As ClassValidaXML = New ClassValidaXML
        Dim resFATConfig As IEnumerable(Of FATConfig)

        Dim objFATConfig As FATConfig

        Dim iIndice As Integer

        Dim objCancEvento As CancXsd.TEvento
        Dim XMLStringEvento As String
        Dim XMLString2 As String
        Dim iSerie As Integer
        Dim sSerie As String
        Dim lNumNotaFiscal As Long
        Dim results As IEnumerable(Of NFeNFiscal)
        Dim aobjEvento(20) As CancXsd.TEvento
        Dim lLote As Long
        '        Dim iSeqEvento As Integer
        Dim objConsultaNFe As ClassConsultaNFe = New ClassConsultaNFe

        Try

            Dim sModelo As String

            Dim NfeRecEvento As New recepcaoevento.NFeRecepcaoEvento4

            gobjApp.sErro = "6"
            gobjApp.sMsg1 = "vai inserir na tabela NFeFedLoteLog"
            If gobjApp.iDebug = 1 Then MsgBox(gobjApp.sErro)

            Call gobjApp.GravarLog("Iniciando o envio do evento de cancelamento", 0, lNumIntNF)

            resFATConfig = gobjApp.dbDadosNfe.ExecuteQuery(Of FATConfig) _
            ("SELECT * FROM FatConfig WHERE Codigo = {0} ", "NUM_PROX_LOTE_EVENTO_CANC")

            objFATConfig = resFATConfig(0)

            lLote = CLng(objFATConfig.Conteudo)

            iResult = gobjApp.dbDadosNfe.ExecuteCommand("UPDATE FatConfig Set Conteudo = {0} WHERE Codigo = {1}", lLote + 1, "NUM_PROX_LOTE_EVENTO_CANC")


            If gobjApp.iDebug = 1 Then MsgBox("NumIntNF =" & CStr(lNumIntNF))
            gobjApp.sErro = "6.1"
            gobjApp.sMsg1 = "vai acessar NFEFedProtNFE. NumIntNF = " & CStr(lNumIntNF)
            If gobjApp.iDebug = 1 Then MsgBox(gobjApp.sErro)


            resNFeFedProtNFE = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedProtNFe) _
            ("SELECT * FROM NFeFedProtNFE WHERE NumIntNF = {0} AND Len(nProt) > 0 AND (cStat = '100' Or cStat = '150') ORDER BY Data DESC, Hora DESC", lNumIntNF)

            gobjApp.sErro = "6.2"
            gobjApp.sMsg1 = "vai pegar a resposta do acesso a NFeFedProtNFe"
            If gobjApp.iDebug = 1 Then MsgBox(gobjApp.sErro)

            objNFeFedProtNFe = resNFeFedProtNFE(0)

            sChNFe = objNFeFedProtNFe.chNFe

            'lNumIntNFiscalParam
            resNFeNFiscal = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeNFiscal) _
            ("SELECT * FROM NFeNFiscal WHERE NumIntDoc = {0} ", lNumIntNF)

            If gobjApp.iDebug = 1 Then MsgBox("NumIntNF = " & lNumIntNF)
            gobjApp.sErro = "6.3"
            gobjApp.sMsg1 = "vai pegar a resposta do acesso a NFeNFiscal"
            If gobjApp.iDebug = 1 Then MsgBox(gobjApp.sErro)

            objNFeNFiscal = resNFeNFiscal(0)

            objCancEvento = New CancXsd.TEvento

            objCancEvento.versao = "1.00"

            objCancEvento.infEvento = New CancXsd.TEventoInfEvento

            objCancEvento.infEvento.Id = "ID" & "110111" & objNFeFedProtNFe.chNFe & "01"

            sUF = Left(objNFeFedProtNFe.chNFe, 2)

            If objNFeNFiscal.ModDocFisE = 35 Then
                sModelo = "NFCe"
            Else
                sModelo = "NFe"
            End If

            gobjApp.gsModelo = sModelo


            objCancEvento.infEvento.cOrgao = GetCode(Of CancXsd.TCOrgaoIBGE)(sUF)

            If gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then
                objCancEvento.infEvento.tpAmb = CancXsd.TAmb.Item2
            Else
                objCancEvento.infEvento.tpAmb = CancXsd.TAmb.Item1
            End If

            objCancEvento.infEvento.ItemElementName = CancXsd.ItemChoiceType.CNPJ
            objCancEvento.infEvento.Item = Mid(objNFeFedProtNFe.chNFe, 7, 14)

            objCancEvento.infEvento.chNFe = objNFeFedProtNFe.chNFe

            objCancEvento.infEvento.dhEvento = Format(Now.Date, "yyyy-MM-dd") & "T" & Format(TimeOfDay, "HH:mm:ss") & Format(TimeOfDay, "zzz")

            objCancEvento.infEvento.tpEvento = TEventoInfEventoTpEvento.Item110111

            objCancEvento.infEvento.nSeqEvento = "1"

            objCancEvento.infEvento.verEvento = TEventoInfEventoVerEvento.Item100

            objCancEvento.infEvento.detEvento = New CancXsd.TEventoInfEventoDetEvento

            objCancEvento.infEvento.detEvento.versao = TEventoInfEventoDetEventoVersao.Item100

            objDescEvento = New TEventoInfEventoDetEventoDescEvento
            objCancEvento.infEvento.detEvento.descEvento = objDescEvento

            objCancEvento.infEvento.detEvento.xJust = Trim(DesacentuaTexto(sMotivo))

            objCancEvento.infEvento.detEvento.nProt = objNFeFedProtNFe.nProt


            gobjApp.sErro = "20"
            gobjApp.sMsg1 = "vai serializar TEvento"
            If gobjApp.iDebug = 1 Then MsgBox(gobjApp.sErro)

            Dim mySerializer As New XmlSerializer(GetType(CancXsd.TEvento))

            Dim XMLStream2 = New MemoryStream(10000)
            mySerializer.Serialize(XMLStream2, objCancEvento)

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

            XMLStringEvento = Mid(xString, 22)

            objEnvCancEvento.idLote = lLote
            objEnvCancEvento.versao = "1.00"

            Dim mySerializerw As New XmlSerializer(GetType(CancXsd.TEnvEvento))

            XMLStream1 = New MemoryStream(10000)

            mySerializerw.Serialize(XMLStream1, objEnvCancEvento)

            Dim xmw As Byte()
            xmw = XMLStream1.ToArray

            XMLString1 = System.Text.Encoding.UTF8.GetString(xmw)

            XMLString2 = Mid(XMLString1, 1, Len(XMLString1) - 12) & XMLStringEvento & Mid(XMLString1, Len(XMLString1) - 12)

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
            sArquivo = gobjApp.sDirXml & objEnvCancEvento.idLote & "-env-evento-canc.xml"
            '            DocDados.Save(sArquivo)
            Dim writer1 As New XmlTextWriter(sArquivo, Nothing)

            writer1.Formatting = Formatting.None
            DocDados.WriteTo(writer1)
            writer1.Close()


            If gobjApp.iDebug = 1 Then MsgBox("22")
            gobjApp.sErro = "22"
            gobjApp.sMsg1 = "vai validar o arquivo xml de envio de evento de cancelamento"

            lErro = objValidaXML.validaXML(sArquivo, gobjApp.sDirXsd & "envEventoCancNFe_v1.00.xsd", 0, 0, iFilialEmpresa)
            If lErro = 1 Then

                Call gobjApp.GravarLog("ERRO - Encerrado o envio do evento de cancelamento", 0, 0)

                Exit Try
            End If

            If gobjApp.iDebug = 1 Then MsgBox("23")
            gobjApp.sErro = "23"
            gobjApp.sMsg1 = "vai enviar o evento de cancelamento"

            Dim DocDados1 As New XmlDocument

            Call Salva_Arquivo(DocDados1, XMLString2)

            ''grava o arquivo com final ped-can.xml que vai ser necessário para a posterior geração do arquivo can.xml
            'lErro = Cancelamento_NFe_Gravar_PED_CANC_XML(gobjApp.iNFeAmbiente, Trim(DesacentuaTexto(sMotivo)), gobjApp.iDebug, objNFeFedProtNFe, gobjApp.sDirXml, gobjApp.sErro)
            'If lErro <> SUCESSO Then Throw New System.Exception("Erro na rotina Cancelamento_NFe_Gravar_PED_CANC_XML.")

            'NfeCabec.cUF = CStr(gobjApp.objEstado.CodIBGE)
            'NfeCabec.versaoDados = "1.00"

            'NfeRecEvento.nfeCabecMsgValue = NfeCabec

            Dim sURL As String
            sURL = ""
            Call WS_Obter_URL(sURL, gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO, gobjApp.objEstado.Sigla, "RecepcaoEvento", sModelo)

            NfeRecEvento.Url = sURL

            Dim XMLStringRetEnvCanc As String

            NfeRecEvento.ClientCertificates.Add(gobjApp.cert)
            xmlNode1 = NfeRecEvento.nfeRecepcaoEvento(DocDados1)

            XMLStringRetEnvCanc = xmlNode1.OuterXml

            xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringRetEnvCanc)

            XMLStreamRet = New MemoryStream(10000)
            XMLStreamRet.Write(xRet, 0, xRet.Length)

            Dim mySerializerRetEnvCCE As New XmlSerializer(GetType(CancXsd.TRetEnvEvento))

            Dim objRetEnvCanc As CancXsd.TRetEnvEvento = New CancXsd.TRetEnvEvento

            XMLStreamRet.Position = 0

            objRetEnvCanc = mySerializerRetEnvCCE.Deserialize(XMLStreamRet)

            gobjApp.sErro = "24"
            gobjApp.sMsg1 = "trata o retorno do evento de cancelamento"
            If gobjApp.iDebug = 1 Then MsgBox(gobjApp.sErro)


            If Not objRetEnvCanc.retEvento Is Nothing Then

                For iIndice = 0 To objRetEnvCanc.retEvento.Count - 1

                    iSerie = CInt(Mid(objRetEnvCanc.retEvento(iIndice).infEvento.chNFe, 23, 3))
                    sSerie = iSerie & "-e"
                    lNumNotaFiscal = CLng(Mid(objRetEnvCanc.retEvento(iIndice).infEvento.chNFe, 26, 9))

                    results = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeNFiscal) _
                    ("SELECT * FROM NFeNFiscal WHERE Serie = {0} AND NumNotaFiscal = {1} AND FilialEmpresa = {2}", sSerie, lNumNotaFiscal, iFilialEmpresa)

                    For Each objNFeNFiscal In results

                        iResult = gobjApp.dbDadosNfe.ExecuteCommand("INSERT INTO NFeFedRetEnvEventoCanc ( FilialEmpresa, NumIntNF, versao, idLote, tpAmb, verAplic, cOrgao, xMotivo, Id, tpAmb1, verAplic1, cOrgao1, cStat1, xMotivo1, chNFe, tpEvento, xEvento, nSeqEvento, CPFCNPJ, email, dataRegEvento, horaRegEvento, nProt, cStat) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23} )", _
                            iFilialEmpresa, objNFeNFiscal.NumIntDoc, objRetEnvCanc.versao, objRetEnvCanc.idLote, objRetEnvCanc.tpAmb, objRetEnvCanc.verAplic, objRetEnvCanc.cOrgao, IIf(objRetEnvCanc.xMotivo Is Nothing, "", objRetEnvCanc.xMotivo), IIf(objRetEnvCanc.retEvento(iIndice).infEvento.Id Is Nothing, "", objRetEnvCanc.retEvento(iIndice).infEvento.Id), IIf(objRetEnvCanc.retEvento(iIndice).infEvento.tpAmb = Nothing, 0, objRetEnvCanc.retEvento(iIndice).infEvento.tpAmb), IIf(objRetEnvCanc.retEvento(iIndice).infEvento.verAplic Is Nothing, "", objRetEnvCanc.retEvento(iIndice).infEvento.verAplic), IIf(objRetEnvCanc.retEvento(iIndice).infEvento.cOrgao = Nothing, 0, objRetEnvCanc.retEvento(iIndice).infEvento.cOrgao), IIf(objRetEnvCanc.retEvento(iIndice).infEvento.cStat Is Nothing, "", objRetEnvCanc.retEvento(iIndice).infEvento.cStat), IIf(objRetEnvCanc.retEvento(iIndice).infEvento.xMotivo Is Nothing, "", objRetEnvCanc.retEvento(iIndice).infEvento.xMotivo), IIf(objRetEnvCanc.retEvento(iIndice).infEvento.chNFe Is Nothing, "", objRetEnvCanc.retEvento(iIndice).infEvento.chNFe), _
                            IIf(objRetEnvCanc.retEvento(iIndice).infEvento.tpEvento Is Nothing, 0, objRetEnvCanc.retEvento(iIndice).infEvento.tpEvento), IIf(objRetEnvCanc.retEvento(iIndice).infEvento.xEvento Is Nothing, "", objRetEnvCanc.retEvento(iIndice).infEvento.xEvento), IIf(objRetEnvCanc.retEvento(iIndice).infEvento.nSeqEvento Is Nothing, 0, objRetEnvCanc.retEvento(iIndice).infEvento.nSeqEvento), IIf(objRetEnvCanc.retEvento(iIndice).infEvento.Item Is Nothing, "", objRetEnvCanc.retEvento(iIndice).infEvento.Item), IIf(objRetEnvCanc.retEvento(iIndice).infEvento.emailDest Is Nothing, "", objRetEnvCanc.retEvento(iIndice).infEvento.emailDest), IIf(objRetEnvCanc.retEvento(iIndice).infEvento.dhRegEvento Is Nothing, Now.Date, _
                            CDate(objRetEnvCanc.retEvento(iIndice).infEvento.dhRegEvento).Date), IIf(objRetEnvCanc.retEvento(iIndice).infEvento.dhRegEvento Is Nothing, TimeOfDay.ToOADate, CDate(objRetEnvCanc.retEvento(iIndice).infEvento.dhRegEvento).TimeOfDay.TotalDays), IIf(objRetEnvCanc.retEvento(iIndice).infEvento.nProt Is Nothing, "", objRetEnvCanc.retEvento(iIndice).infEvento.nProt), IIf(objRetEnvCanc.retEvento(iIndice).infEvento.cStat Is Nothing, "", objRetEnvCanc.retEvento(iIndice).infEvento.cStat))

                        Form1.Msg.Items.Add("Retorno do evento " & IIf(objRetEnvCanc.retEvento(iIndice).infEvento.xEvento Is Nothing, "cancelamento", objRetEnvCanc.retEvento(iIndice).infEvento.xEvento) & " nSeqEvento = " & objRetEnvCanc.retEvento(iIndice).infEvento.nSeqEvento & " chave NFe = " & objRetEnvCanc.retEvento(iIndice).infEvento.chNFe)
                        Form1.Msg.Items.Add("cStat = " & objRetEnvCanc.retEvento(iIndice).infEvento.cStat & " xMotivo1 = " & objRetEnvCanc.retEvento(iIndice).infEvento.xMotivo)

                        Exit For

                    Next

                Next

            Else

                iResult = gobjApp.dbDadosNfe.ExecuteCommand("INSERT INTO NFeFedRetEnvEventoCanc ( FilialEmpresa, NumIntNF, versao, idLote, tpAmb, verAplic, cOrgao, xMotivo, chNFe, nSeqEvento, dataRegEvento, horaRegEvento) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11} )", _
                iFilialEmpresa, 0, objRetEnvCanc.versao, objRetEnvCanc.idLote, objRetEnvCanc.tpAmb, objRetEnvCanc.verAplic, objRetEnvCanc.cOrgao, objRetEnvCanc.xMotivo, objNFeFedProtNFe.chNFe, 1, Now.Date, TimeOfDay.ToOADate)

                Form1.Msg.Items.Add("Ocorreu um erro no envio do evento de cancelamento, cStat = " & objRetEnvCanc.cStat & " xMotivo = " & objRetEnvCanc.xMotivo)

            End If

            If gobjApp.iDebug = 1 Then MsgBox("26")
            gobjApp.sErro = "26"
            gobjApp.sMsg1 = "vai gravar o xml"

            Dim DocDados2 As XmlDocument = New XmlDocument
            XMLStreamRet.Position = 0
            DocDados2.Load(XMLStreamRet)
            sArquivo = gobjApp.sDirXml & objRetEnvCanc.idLote & "-ret-env-evento-canc.xml"
            DocDados2.Save(sArquivo)

            gobjApp.DadosCommit()

            Call gobjApp.GravarLog("Envio do evento encerrado.", 0, 0)

            Evento_Cancela_NFe = 1

        Catch ex As Exception

            Evento_Cancela_NFe = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

            Call gobjApp.GravarLog(gobjApp.sErro & " " & gobjApp.sMsg1, 0, 0)

            Call gobjApp.GravarLog("ERRO - Encerrado o envio do evento de cancelamento", 0, 0)

            Call gobjApp.GravarLog("O CANCELAMENTO DA NOTA FISCAL " & objNFeNFiscal.NumNotaFiscal & " NÃO FOI HOMOLOGADO. TENTE REFAZER O CANCELAMENTO.", 0, lNumIntNF)

        Finally


        End Try

    End Function

End Class
