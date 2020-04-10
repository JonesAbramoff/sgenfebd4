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

Public Class ClassEnvioNFe

    Public Function Envia_Lote_NFe(ByVal sEmpresa As String, ByVal lLote As Long, ByVal iFilialEmpresa As Integer) As Long

        Dim sArquivo As String

        Dim XMLStream As MemoryStream = New MemoryStream(10000)
        Dim XMLStream1 As MemoryStream = New MemoryStream(10000)
        'Dim XMLStreamCabec As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamRet As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamDados As MemoryStream = New MemoryStream(10000)


        Dim results As IEnumerable(Of NFeNFiscal)
        Dim resNFeFedLote As IEnumerable(Of NFeFedLote)
        Dim resNFeFedDenegada As IEnumerable(Of NFeFedDenegada)
        Dim objNFeFedLote As NFeFedLote
        '        Dim objEndereco As Endereco
        Dim objNFiscal As NFeNFiscal

        Dim objConsultaNFe As ClassConsultaNFe = New ClassConsultaNFe

        Dim objStatServ As TConsStatServ = New TConsStatServ
        Dim objRetStatServ As TRetConsStatServ = New TRetConsStatServ
        Dim lErro As Long

        Dim colNFiscal As Collection = New Collection

        Dim envioNFe As TEnviNFe = New TEnviNFe


        Dim XMLString As String
        Dim XMLString1 As String
        Dim XMLString2 As String
        Dim XMLString3 As String
        Dim XMLStringNFes As String


        Dim iResult As Integer

        '        Dim certificado As Certificado = New Certificado

        Dim lNumIntNF As Long

        Dim objValidaXML As ClassValidaXML = New ClassValidaXML

        Dim iPos As Integer
        Dim objconsRetReciNFe As TRetConsReciNFe = New TRetConsReciNFe

        'Dim a4 As nfeautorizacao.nfeCabecMsg = New nfeautorizacao.nfeCabecMsg

        'Dim NFeAutorizacao As New nfeautorizacao.NfeAutorizacao
        'Dim NFeRetAutorizacao As New nferetautorizacao.NfeRetAutorizacao
        'Dim NFecabec_EnvioLote As New nfeautorizacao.nfeCabecMsg
        'Dim NFecabec_RetEnvioLote As New nferetautorizacao.nfeCabecMsg

        'Dim NFeAutorizacaoPR As New PR_NFeAutorizacao3.NfeAutorizacao3
        'Dim NFeRetAutorizacaoPR As New PR_NFeRetAutorizacao3.NfeRetAutorizacao3
        'Dim NFecabec_EnvioLote As New nfeautorizacao.nfeCabecMsg
        'Dim NFecabec_RetEnvioLote As New nferetautorizacao.nfeCabecMsg

        Dim sSerie As String
        Dim lNumNotaFiscal As Long

        Dim IPITrib As TIpi = New TIpi

        Dim iScan As Integer
        Dim xRet As Byte()

        Dim uniEncoding As New UnicodeEncoding()

        Dim sMsg As String

        Dim a5 As TNFe

        Dim sArq As String
        Dim sChaveNFe As String
        Dim sChaveNFe1 As String
        Dim iAchou As Integer

        Dim xmlNode1 As XmlNode

        Try

            sSerie = ""
            XMLStringNFes = ""

            lNumIntNF = 0

            resNFeFedLote = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedLote) _
            ("SELECT * FROM NFeFedLote WHERE Lote = {0} AND FilialEmpresa = {1} ORDER BY NumIntNF", lLote, iFilialEmpresa)

            Form1.ProgressBar1.Maximum = resNFeFedLote.Count
            Form1.ProgressBar1.Minimum = 0
            Form1.ProgressBar1.Value = 0

            iScan = 0

            If gobjApp.iDebug = 1 Then MsgBox("3")
            gobjApp.sErro = "3"
            gobjApp.sMsg1 = "para cada nota fiscal do lote vai chamar Monta_NFiscal_Xml"

            resNFeFedLote = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedLote) _
            ("SELECT * FROM NFeFedLote WHERE Lote = {0} AND FilialEmpresa = {1} ORDER BY NumIntNF", lLote, iFilialEmpresa)

            For Each objNFeFedLote In resNFeFedLote

                sMsg = ""

                lNumIntNF = objNFeFedLote.NumIntNF

                If gobjApp.iDebug = 1 Then MsgBox("NumIntNF = " & lNumIntNF)

                results = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeNFiscal) _
                ("SELECT * FROM NFeNFiscal WHERE NumIntDoc = {0} ", objNFeFedLote.NumIntNF)

                For Each objNFiscal In results

                    sChaveNFe1 = ""
                    Dim X As New ClassMontaXmlNFe
                    lErro = X.Monta_NFiscal_Xml(a5, objNFiscal, sSerie, lLote, colNFiscal, sChaveNFe1, XMLStringNFes)
                    If lErro <> SUCESSO Then Throw New System.Exception("Erro na rotina Monta_Fiscal_Xml.")


                    If gobjApp.iDebug = 1 Then MsgBox("38")
                    gobjApp.sErro = "38"
                    gobjApp.sMsg1 = "vai continuar a montar o arquivo LoteNFe<n. do lote>.txt - setor B"

                Next

                Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 1

            Next

            If gobjApp.iDebug = 1 Then MsgBox("39")
            gobjApp.sErro = "39"
            gobjApp.sMsg1 = "vai atualizar FatConfig e inserir NFeFedLoteLog"

            sSerie = ""
            lNumNotaFiscal = 0

            Call gobjApp.GravarLog("Iniciando a validação do lote", lLote, 0)

            lNumIntNF = 0

            envioNFe.versao = NFE_VERSAO_XML
            envioNFe.idLote = lLote

            Dim mySerializerw As New XmlSerializer(GetType(TEnviNFe))

            XMLStream1 = New MemoryStream(10000)

            mySerializerw.Serialize(XMLStream1, envioNFe)

            Dim xmw As Byte()
            xmw = XMLStream1.ToArray

            XMLString1 = System.Text.Encoding.UTF8.GetString(xmw)

            XMLString2 = Mid(XMLString1, 1, Len(XMLString1) - 10) & XMLStringNFes & Mid(XMLString1, Len(XMLString1) - 10)

            XMLString2 = Mid(XMLString2, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLString2, 20)

            iPos = InStr(XMLString2, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")

            If iPos <> 0 Then

                XMLString2 = Mid(XMLString2, 1, iPos - 1) & Mid(XMLString2, iPos + 99)

            End If


            Dim XMLStringRetEnvNFE As String

            '************* valida dados antes do envio **********************
            Dim xDados As Byte()

            xDados = System.Text.Encoding.UTF8.GetBytes(XMLString2)

            XMLStreamDados = New MemoryStream(10000)

            XMLStreamDados.Write(xDados, 0, xDados.Length)

            Dim DocDados As XmlDocument = New XmlDocument
            XMLStreamDados.Position = 0
            DocDados.Load(XMLStreamDados)
            sArquivo = gobjApp.sDirXml & envioNFe.idLote & "-env-lot.xml"
            DocDados.Save(sArquivo)

            If gobjApp.iDebug = 1 Then MsgBox("39.1")
            gobjApp.sErro = "39.1"
            gobjApp.sMsg1 = "vai validar o XML"


            'lErro = objValidaXML.validaXML(sArquivo, gobjApp.sDirXsd & "enviNFe_v4.00.xsd", lLote, lNumIntNF, iFilialEmpresa)
            'If lErro = 1 Then

            '    Call gobjApp.GravarLog("ERRO - o envio do lote " & CStr(lLote) & " foi encerrado por erro." & " " & sArquivo & " " & gobjApp.sDirXsd & "enviNFe_v4.00.xsd", lLote, 0)

            '    Exit Try

            'End If

            If gobjApp.iDebug = 1 Then MsgBox("39.2")
            gobjApp.sErro = "39.2"
            gobjApp.sMsg1 = "validou o XML"


            Dim DocDados1 = New XmlDocument

            Call Salva_Arquivo(DocDados1, XMLString2)

            Dim NFeAutorizacao As New nfeautorizacao.NFeAutorizacao4
            Dim NFeRetAutorizacao As New nferetautorizacao.NFeRetAutorizacao4

            If gobjApp.iDebug = 1 Then MsgBox("39.6")
            gobjApp.sErro = "39.6"
            gobjApp.sMsg1 = "vai enviar a nota"

            Dim sURL As String
            sURL = ""
            Call WS_Obter_URL(sURL, gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO, gobjApp.objEstado.Sigla, "NFeAutorizacao", gobjApp.gsModelo)

            If gobjApp.iDebug = 1 Then MsgBox("39.6.1")
            gobjApp.sErro = "39.6.1"
            gobjApp.sMsg1 = "vai enviar a nota"

            NFeAutorizacao.Url = sURL

            If gobjApp.iDebug = 1 Then MsgBox("39.6.2")
            gobjApp.sErro = "39.6.2"
            gobjApp.sMsg1 = "vai enviar a nota"

            NFeAutorizacao.ClientCertificates.Add(gobjApp.cert)

            If gobjApp.iDebug = 1 Then MsgBox("39.6.3")
            gobjApp.sErro = "39.6.3"
            gobjApp.sMsg1 = "vai enviar a nota"

            xmlNode1 = NFeAutorizacao.nfeAutorizacaoLote(DocDados1)

            If gobjApp.iDebug = 1 Then MsgBox("39.6.4")
            gobjApp.sErro = "39.6.4" & sURL
            gobjApp.sMsg1 = "vai enviar a nota para " & sURL

            XMLStringRetEnvNFE = xmlNode1.OuterXml

            If gobjApp.iDebug = 1 Then
                MsgBox("39.7")
                MsgBox(XMLStringRetEnvNFE)

            End If
            gobjApp.sErro = "39.7"
            gobjApp.sMsg1 = "enviou a nota"

            xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringRetEnvNFE)

            XMLStreamRet = New MemoryStream(10000)

            XMLStreamRet.Write(xRet, 0, xRet.Length)

            Dim mySerializerRetEnvNFe As New XmlSerializer(GetType(TRetEnviNFe))

            Dim objRetEnviNFE As TRetEnviNFe = New TRetEnviNFe

            XMLStreamRet.Position = 0

            objRetEnviNFE = mySerializerRetEnvNFe.Deserialize(XMLStreamRet)

            Dim snRec As String
            Dim infRec As TRetEnviNFeInfRec

            If objRetEnviNFE.Item.GetType.FullName.ToString = "sgenfebd4.NFeXsd.TRetEnviNFeInfRec" Then

                infRec = objRetEnviNFE.Item

            Else

                gobjApp.sErro = "39.7a"
                gobjApp.sMsg1 = objRetEnviNFE.Item.GetType.FullName.ToString

            End If

            If infRec Is Nothing Then
                snRec = ""
            Else
                snRec = infRec.nRec
            End If

            Dim dthRecbto As Date

            If infRec Is Nothing Then
                dthRecbto = Now
            Else
                dthRecbto = objRetEnviNFE.dhRecbto
            End If

            Dim stMed As String

            If infRec Is Nothing Then
                stMed = ""
            Else
                stMed = infRec.tMed
            End If

            iResult = gobjApp.dbLog.ExecuteCommand("INSERT INTO NFeFedRetEnvi ( FilialEmpresa, Lote, tpAmb, verAplic, versao, cStat, xMotivo, cUF, nRec, tMed, data, hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10},{11} )", _
            iFilialEmpresa, envioNFe.idLote, objRetEnviNFE.tpAmb, objRetEnviNFE.verAplic, objRetEnviNFE.versao, objRetEnviNFE.cStat, objRetEnviNFE.xMotivo, objRetEnviNFE.cUF, snRec, stMed, dthRecbto.Date, dthRecbto.TimeOfDay.TotalDays)

            Call gobjApp.GravarLog("Retorno do envio do lote - " & objRetEnviNFE.xMotivo, lLote, 0)

            If snRec <> "" Then

                Dim objconsReciNFe As TConsReciNFe = New TConsReciNFe

                If gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then
                    objconsReciNFe.tpAmb = TAmb.Item2
                Else
                    objconsReciNFe.tpAmb = TAmb.Item1
                End If
                objconsReciNFe.versao = NFE_VERSAO_XML
                objconsReciNFe.nRec = infRec.nRec

                Dim mySerializerx As New XmlSerializer(GetType(TConsReciNFe))

                XMLStream1 = New MemoryStream(10000)
                mySerializerx.Serialize(XMLStream1, objconsReciNFe)

                Dim xm1 As Byte()
                xm1 = XMLStream1.ToArray

                XMLString1 = System.Text.Encoding.UTF8.GetString(xm1)

                XMLString1 = Mid(XMLString1, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLString1, 20)

                Dim XMLStringRetConsReciNFE As String

                iPos = InStr(XMLString1, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")

                If iPos <> 0 Then

                    XMLString1 = Mid(XMLString1, 1, iPos - 1) & Mid(XMLString1, iPos + 99)

                End If

                DocDados1 = New XmlDocument

                Call Salva_Arquivo(DocDados1, XMLString1)

                'NFecabec_RetEnvioLote.cUF = CStr(gobjApp.objEstado.CodIBGE)
                'NFecabec_RetEnvioLote.versaoDados = NFE_VERSAO_XML

                'NFeRetAutorizacao.nfeCabecMsgValue = NFecabec_RetEnvioLote

                Call gobjApp.GravarLog("Iniciando a consulta do status do lote - Aguarde", lLote, 0)

                Dim i1 As Integer

                i1 = 1

                Do While i1 < 11

                    Sleep(2000)

                    Call WS_Obter_URL(sURL, gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO, gobjApp.objEstado.Sigla, "NFeRetAutorizacao", gobjApp.gsModelo)

                    NFeRetAutorizacao.Url = sURL

                    NFeRetAutorizacao.ClientCertificates.Add(gobjApp.cert)

                    Select Case WS_Obter_Autorizador(gobjApp.objEstado.Sigla)

                        'Case "PR"
                        'xmlNode1 = NFeRetAutorizacao.nfeRetAutorizacao(DocDados1)

                        Case Else
                            xmlNode1 = NFeRetAutorizacao.nfeRetAutorizacaoLote(DocDados1)

                    End Select

                    XMLStringRetConsReciNFE = xmlNode1.OuterXml


                    xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringRetConsReciNFE)

                    XMLStreamRet = New MemoryStream(10000)
                    XMLStreamRet.Write(xRet, 0, xRet.Length)

                    Dim mySerializerRetConsReciNFe As New XmlSerializer(GetType(TRetConsReciNFe))

                    objconsRetReciNFe = New TRetConsReciNFe


                    XMLStreamRet.Position = 0

                    objconsRetReciNFe = mySerializerRetConsReciNFe.Deserialize(XMLStreamRet)

                    If objconsRetReciNFe.cStat = "105" Then

                        Call gobjApp.GravarLog("Lote em Processamento - Aguarde - Tentativa " & i1 & "/10", lLote, 0)

                    Else

                        iResult = gobjApp.dbLog.ExecuteCommand("INSERT INTO NFeFedRetConsReci ( Lote, FilialEmpresa,versao, tpAmb, verAplic, nRec, cStat, xMotivo, cUF, data, hora, cMsg, xMsg) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}  )", _
                        lLote, iFilialEmpresa, objconsRetReciNFe.versao, objconsRetReciNFe.tpAmb, objconsRetReciNFe.verAplic, objconsRetReciNFe.nRec, objconsRetReciNFe.cStat, objconsRetReciNFe.xMotivo, objconsRetReciNFe.cUF, Now.Date, TimeOfDay.ToOADate, IIf(objconsRetReciNFe.cMsg Is Nothing, "", objconsRetReciNFe.cMsg), IIf(objconsRetReciNFe.xMsg Is Nothing, "", objconsRetReciNFe.xMsg))

                        Call gobjApp.GravarLog("Retorno da consulta do lote - " & objconsRetReciNFe.xMotivo & IIf(objconsRetReciNFe.cMsg <> "0", " - código de Msg = " & objconsRetReciNFe.cMsg & " - " & objconsRetReciNFe.xMsg, ""), lLote, 0)

                        If Not objconsRetReciNFe.protNFe Is Nothing Then

                            For i = 0 To objconsRetReciNFe.protNFe.Count - 1

                                If String.IsNullOrEmpty(objconsRetReciNFe.protNFe(i).infProt.nProt) Then
                                    objconsRetReciNFe.protNFe(i).infProt.nProt = ""
                                End If

                                iAchou = 0

                                For Each objNFiscal In colNFiscal
                                    If Format(CInt(Serie_Sem_E(objNFiscal.Serie)), "000") = Mid(objconsRetReciNFe.protNFe(i).infProt.chNFe, 23, 3) And _
                                       Format(objNFiscal.NumNotaFiscal, "000000000") = Mid(objconsRetReciNFe.protNFe(i).infProt.chNFe, 26, 9) Then
                                        iAchou = 1
                                        Exit For
                                    End If
                                Next

                                If iAchou = 0 Then
                                    Throw New System.Exception("A nota nao corresponde a nenhuma nota enviada. Serie Consulta = " & Mid(objconsRetReciNFe.protNFe(i).infProt.chNFe, 23, 3) & " Numero Consulta = " & Mid(objconsRetReciNFe.protNFe(i).infProt.chNFe, 26, 9))
                                End If

                                Call gobjApp.GravarLog("Nota Fiscal  - " & objconsRetReciNFe.protNFe(i).infProt.chNFe & " - " & objconsRetReciNFe.protNFe(i).infProt.xMotivo, lLote, 0)

                                If objconsRetReciNFe.protNFe(i).infProt.cStat = "100" Or objconsRetReciNFe.protNFe(i).infProt.cStat = "150" Then

                                    If objconsRetReciNFe.protNFe(i).versao <> "1.10" And objconsRetReciNFe.protNFe(i).versao <> "2.00" And objconsRetReciNFe.protNFe(i).versao <> "3.10" And objconsRetReciNFe.protNFe(i).versao <> "4.00" Then
                                        Throw New System.Exception("Versao não tratada. Versao = " & objconsRetReciNFe.protNFe(i).versao)
                                    End If

                                    Dim DocDados2 As XmlDocument = New XmlDocument
                                    Dim XMLStreamDados1 = New MemoryStream(10000)

                                    sArquivo = gobjApp.sDirXml & objconsRetReciNFe.protNFe(i).infProt.chNFe & "-pre.xml"
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
                                    objNFeProc.protNFe = objconsRetReciNFe.protNFe(i)

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

                                        XMLString = Mid(XMLString, 1, iPos - 1) & Mid(XMLString, iPos + 99)

                                    End If


                                    XMLStreamDados = New MemoryStream(10000)

                                    Dim xDados1 As Byte()

                                    xDados1 = System.Text.Encoding.UTF8.GetBytes(XMLString)

                                    XMLStreamDados.Write(xDados1, 0, xDados1.Length)


                                    Dim DocDados3 As XmlDocument = New XmlDocument

                                    XMLStreamDados.Position = 0
                                    DocDados3.Load(XMLStreamDados)
                                    'sArquivo = sDir & objNFe.infNFe.Id & ".xml"
                                    sArquivo = gobjApp.sDirXml & objconsRetReciNFe.protNFe(i).infProt.chNFe & "-procNfe.xml"
                                    '                                    DocDados3.Save(sArquivo)

                                    Dim writer As New XmlTextWriter(sArquivo, Nothing)

                                    writer.Formatting = Formatting.None
                                    DocDados3.WriteTo(writer)
                                    writer.Close()

                                    Dim sQRCode As String = ""

                                    'se for NFCe vai gerar o QRCode
                                    If Mid(objconsRetReciNFe.protNFe(i).infProt.chNFe, 21, 2) = "65" Then

                                        Call gobjApp.NFCe_Trata_QRCode(sQRCode, objconsRetReciNFe.protNFe(i), DocDados3)

                                    End If

                                    iResult = gobjApp.dbLog.ExecuteCommand("INSERT INTO NFeFedProtNFe ( FilialEmpresa, NumIntNF, versao, nRec, tpAmb, verAplic, chNFe, nProt, cStat, xMotivo, Data, Hora, DataRegistro, HoraRegistro, QRCode) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14} )", _
                                            iFilialEmpresa, objNFiscal.NumIntDoc, objconsRetReciNFe.protNFe(i).versao, objconsRetReciNFe.nRec, objconsRetReciNFe.protNFe(i).infProt.tpAmb, objconsRetReciNFe.protNFe(i).infProt.verAplic, objconsRetReciNFe.protNFe(i).infProt.chNFe, objconsRetReciNFe.protNFe(i).infProt.nProt, objconsRetReciNFe.protNFe(i).infProt.cStat, Left(objconsRetReciNFe.protNFe(i).infProt.xMotivo, 255), UTCParaDate(objconsRetReciNFe.protNFe(i).infProt.dhRecbto), UTCParaHora(objconsRetReciNFe.protNFe(i).infProt.dhRecbto), Now.Date, TimeOfDay.ToOADate, sQRCode)

                                    'se for duplicidade pesquisa a chave e faz consulta
                                ElseIf objconsRetReciNFe.protNFe(i).infProt.cStat = "204" Or objconsRetReciNFe.protNFe(i).infProt.cStat = "539" Then

                                    Dim colArq As New Collection

                                    sChaveNFe = Left(objconsRetReciNFe.protNFe(i).infProt.chNFe, 35)

                                    sArq = My.Computer.FileSystem.GetName(gobjApp.sDirXml & sChaveNFe & "*-pre.xml")

                                    sArq = Dir(gobjApp.sDirXml & sArq)

                                    Do While sArq <> ""

                                        colArq.Add(sArq)

                                        sArq = Dir()

                                    Loop


                                    For Each sArq In colArq

                                        sChaveNFe = Left(sArq, 44)

                                        lErro = objConsultaNFe.Consulta_NFe(sEmpresa, sChaveNFe, iFilialEmpresa, 0)

                                        If lErro = SUCESSO Then Exit For

                                    Next

                                    'se for uma nota denegada vai guardar a informacao em NFeFedDenegada
                                ElseIf (objconsRetReciNFe.protNFe(i).infProt.cStat = "205" Or objconsRetReciNFe.protNFe(i).infProt.cStat = "110" Or objconsRetReciNFe.protNFe(i).infProt.cStat = "301" Or objconsRetReciNFe.protNFe(i).infProt.cStat = "302") Then


                                    resNFeFedDenegada = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedDenegada) _
                                    ("SELECT * FROM NFeFedDenegada WHERE NumIntNF = {0} ", objNFiscal.NumIntDoc)

                                    If resNFeFedDenegada.Count = 0 Then

                                        iResult = gobjApp.dbLog.ExecuteCommand("INSERT INTO NFeFedDenegada ( NumIntNF, Processada) VALUES ( {0}, {1})", _
                                        objNFiscal.NumIntDoc, 0)

                                        If objconsRetReciNFe.protNFe(i).versao <> "1.10" And objconsRetReciNFe.protNFe(i).versao <> "2.00" And objconsRetReciNFe.protNFe(i).versao <> "3.10" And objconsRetReciNFe.protNFe(i).versao <> "4.00" Then
                                            Throw New System.Exception("Versao não tratada. Versao = " & objconsRetReciNFe.protNFe(i).versao)
                                        End If

                                        iResult = gobjApp.dbLog.ExecuteCommand("INSERT INTO NFeFedProtNFe ( FilialEmpresa, NumIntNF, versao, nRec, tpAmb, verAplic, chNFe, nProt, cStat, xMotivo, Data, Hora, DataRegistro, HoraRegistro) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13} )", _
                                            iFilialEmpresa, objNFiscal.NumIntDoc, objconsRetReciNFe.protNFe(i).versao, objconsRetReciNFe.nRec, objconsRetReciNFe.protNFe(i).infProt.tpAmb, objconsRetReciNFe.protNFe(i).infProt.verAplic, objconsRetReciNFe.protNFe(i).infProt.chNFe, objconsRetReciNFe.protNFe(i).infProt.nProt, objconsRetReciNFe.protNFe(i).infProt.cStat, Left(objconsRetReciNFe.protNFe(i).infProt.xMotivo, 255), UTCParaDate(objconsRetReciNFe.protNFe(i).infProt.dhRecbto), UTCParaHora(objconsRetReciNFe.protNFe(i).infProt.dhRecbto), Now.Date, TimeOfDay.ToOADate)

                                        lErro = objConsultaNFe.Consulta_NFe(sEmpresa, objconsRetReciNFe.protNFe(i).infProt.chNFe, iFilialEmpresa, 0)


                                    End If

                                Else

                                    'cstat nao tratados acima, provavelmente de erros

                                    iResult = gobjApp.dbLog.ExecuteCommand("INSERT INTO NFeFedProtNFe ( FilialEmpresa, NumIntNF, versao, nRec, tpAmb, verAplic, chNFe, nProt, cStat, xMotivo, Data, Hora, DataRegistro, HoraRegistro) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13} )", _
                                        iFilialEmpresa, objNFiscal.NumIntDoc, objconsRetReciNFe.protNFe(i).versao, objconsRetReciNFe.nRec, objconsRetReciNFe.protNFe(i).infProt.tpAmb, objconsRetReciNFe.protNFe(i).infProt.verAplic, objconsRetReciNFe.protNFe(i).infProt.chNFe, objconsRetReciNFe.protNFe(i).infProt.nProt, objconsRetReciNFe.protNFe(i).infProt.cStat, Left(objconsRetReciNFe.protNFe(i).infProt.xMotivo, 255), UTCParaDate(objconsRetReciNFe.protNFe(i).infProt.dhRecbto), UTCParaHora(objconsRetReciNFe.protNFe(i).infProt.dhRecbto), Now.Date, TimeOfDay.ToOADate)

                                End If

                            Next

                        End If

                        Exit Do

                    End If

                    i1 = i1 + 1

                Loop

                If i1 = 11 Then

                    iResult = gobjApp.dbLog.ExecuteCommand("INSERT INTO NFeFedRetConsReci ( Lote, FilialEmpresa, versao, tpAmb, verAplic, nRec, cStat, xMotivo, cUF, Data, Hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10} )", _
                    lLote, iFilialEmpresa, objconsRetReciNFe.versao, objconsRetReciNFe.tpAmb, objconsRetReciNFe.verAplic, objconsRetReciNFe.nRec, objconsRetReciNFe.cStat, objconsRetReciNFe.xMotivo, objconsRetReciNFe.cUF, Now.Date, TimeOfDay.ToOADate)

                    Call gobjApp.GravarLog("Retorno da consulta do lote - " & objconsRetReciNFe.xMotivo & " . Tente a consulta a este lote mais tarde.", lLote, 0)

                End If

            End If

            gobjApp.DadosCommit()

            Call gobjApp.GravarLog("Encerrado o processamento do lote " & CStr(lLote), lLote, 0)

            Envia_Lote_NFe = SUCESSO

        Catch ex As Exception

            Envia_Lote_NFe = 1

            Dim sMsg2 As String

            If ex.InnerException Is Nothing Then
                sMsg2 = ""
            Else
                sMsg2 = " - " & ex.InnerException.Message
            End If

            Call gobjApp.GravarLog("ERRO - " & ex.Message & sMsg2 & IIf(lNumNotaFiscal <> 0, "Serie = " & sSerie & " Nota Fiscal = " & lNumNotaFiscal, ""), lLote, lNumIntNF)

            Call gobjApp.GravarLog("ERRO - " & gobjApp.sErro & " - " & gobjApp.sMsg1 & IIf(lNumNotaFiscal <> 0, " Serie = " & sSerie & " Nota Fiscal = " & lNumNotaFiscal, ""), lLote, lNumIntNF)

            Call gobjApp.GravarLog("ERRO - o envio do lote " & CStr(lLote) & " foi encerrado por erro.", lLote, 0)

        Finally

        End Try

    End Function

End Class

