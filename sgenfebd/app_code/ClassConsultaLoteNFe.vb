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

Public Class ClassConsultaLoteNFe

    Public Function Consulta_Lote_NFe(ByVal sEmpresa As String, ByVal lLote As Long, ByVal iFilialEmpresa As Integer, Optional ByVal iScan As Integer = -1) As Long

        Dim dic As DataClasses2DataContext = New DataClasses2DataContext
        Dim odbc As OdbcConnection = New OdbcConnection
        Dim xmlNode1 As XmlNode

        Dim dbDadosNfe As DataClassesDataContext = New DataClassesDataContext
        Dim dbLog As DataClassesDataContext = New DataClassesDataContext
        Dim iResult As Integer
        Dim lNumIntNFiscal As Long
        Dim XMLStream1 As MemoryStream = New MemoryStream(10000)
        'Dim XMLStreamCabec As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamRet As MemoryStream = New MemoryStream(10000)


        Dim XMLString1 As String

        Dim resNfeFedDenegada As IEnumerable(Of NFeFedDenegada)
        Dim resNFeFedRetEnvi As IEnumerable(Of NFeFedRetEnvi)
        Dim resNFeFedLote As IEnumerable(Of NFeFedLote)
        Dim resNFeFedProtNFe As IEnumerable(Of NFeFedProtNFe)
        Dim resNFiscal As IEnumerable(Of NFiscal)

        Dim objNFiscal As NFiscal
        Dim objNFeFedLote As NFeFedLote = New NFeFedLote
        Dim objNFeFedRetEnvi As NFeFedRetEnvi = New NFeFedRetEnvi
        Dim objNFeFedProtNFe As NFeFedProtNFe = New NFeFedProtNFe
        Dim xRet As Byte()


        Dim colNumIntNFiscal As Collection = New Collection

        'Dim XMLStringCabec As String

        Dim iPos As Integer
        Dim iAchou As Integer
        Dim j As Integer
        Dim lErro As Long
        Dim sChaveNFe As String
        Dim sArq As String

        Dim objConsultaNFe As ClassConsultaNFe = New ClassConsultaNFe

        Try

            Dim results As IEnumerable(Of NFeNFiscal)
            Dim objNF As NFeNFiscal

            gobjApp.sErro = "6"
            gobjApp.sMsg1 = "vai acessar a tabela NFeFedRetEnvi"

            Dim sModelo As String

            '******** pega o ultimo retorno de envio do lote em questao *************
            resNFeFedRetEnvi = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedRetEnvi) _
            ("SELECT * FROM NFeFedRetEnvi WHERE Lote = {0} ORDER BY data DESC, hora DESC ", lLote)

            objNFeFedRetEnvi = resNFeFedRetEnvi(0)

            gobjApp.sErro = "7"
            gobjApp.sMsg1 = "vai tratar o lote "

            'se o lote tiver sido recebido com sucesso 
            If objNFeFedRetEnvi.cStat = "103" Then

                gobjApp.sErro = "8"
                gobjApp.sMsg1 = "vai montar o cabecalho da mensagem"

                'Dim mySerializercabec As Object
                'Dim objCabecMsg As Object
                Dim NfeConsulta As New nfeconsulta2.NFeConsultaProtocolo4
                'Dim NfeCabec As Object
                Dim NFeRetAutorizacao As New nferetautorizacao.NFeRetAutorizacao4

                'objCabecMsg.versaoDados = NFE_VERSAO_XML


                'gobjApp.sErro = "9"
                'gobjApp.sMsg1 = "vai serializar o cabecalho"

                'XMLStreamCabec = New MemoryStream(10000)

                'mySerializercabec.Serialize(XMLStreamCabec, objCabecMsg)

                'XMLStreamCabec.Position = 0

                'Dim xmcabec As Byte()

                'xmcabec = XMLStreamCabec.ToArray

                'XMLStringCabec = System.Text.Encoding.UTF8.GetString(xmcabec)

                'XMLStringCabec = Mid(XMLStringCabec, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLStringCabec, 20)

                gobjApp.sErro = "10"
                gobjApp.sMsg1 = "vai acessar a tabela NFeFedLote"

                resNFeFedLote = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedLote) _
                ("SELECT * FROM NFeFedLote WHERE Lote = {0} ORDER BY NumIntNF", lLote)

                For Each objNFeFedLote In resNFeFedLote

                    colNumIntNFiscal.Add(objNFeFedLote.NumIntNF)

                    If sModelo = "" Then

                        results = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeNFiscal) _
                        ("SELECT * FROM NFeNFiscal WHERE NumIntDoc = {0} ", objNFeFedLote.NumIntNF)

                        For Each objNF In results

                            If objNF.ModDocFisE = 35 Then
                                sModelo = "NFCe"
                            Else
                                sModelo = "NFe"
                            End If

                            gobjApp.gsModelo = sModelo

                        Next

                    End If

                Next

                gobjApp.sErro = "11"
                gobjApp.sMsg1 = "vai montar a estrutura TConsReciNFe"

                Dim objconsReciNFe As TConsReciNFe = New TConsReciNFe

                If gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then
                    objconsReciNFe.tpAmb = TAmb.Item2
                Else
                    objconsReciNFe.tpAmb = TAmb.Item1
                End If

                objconsReciNFe.versao = NFE_VERSAO_XML
                objconsReciNFe.nRec = objNFeFedRetEnvi.nRec

                Dim mySerializerx As New XmlSerializer(GetType(TConsReciNFe))

                gobjApp.sErro = "12"
                gobjApp.sMsg1 = "serializa a estrutura TConsReciNFe"

                XMLStream1 = New MemoryStream(10000)
                mySerializerx.Serialize(XMLStream1, objconsReciNFe)

                Dim xm1 As Byte()
                xm1 = XMLStream1.ToArray

                XMLString1 = System.Text.Encoding.UTF8.GetString(xm1)

                XMLString1 = Mid(XMLString1, 1, 19) & " encoding=""UTF-8"" " & Mid(XMLString1, 20)

                iPos = InStr(XMLString1, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")

                If iPos <> 0 Then

                    XMLString1 = Mid(XMLString1, 1, iPos - 1) & Mid(XMLString1, iPos + 99)

                End If


                Dim XMLStringRetConsReciNFE As String

                gobjApp.sErro = "12"
                gobjApp.sMsg1 = "insere os dados na tabela NFeFedLoteLog"


                Call gobjApp.GravarLog("Iniciando a consulta do status do lote", lLote, 0)

                gobjApp.sErro = "13"
                gobjApp.sMsg1 = "vai enviar a msg de consulta do lote"


                Dim DocDados1 As New XmlDocument

                Call Salva_Arquivo(DocDados1, XMLString1)

                'NfeCabec.cUF = CStr(gobjApp.objEstado.CodIBGE)
                'NfeCabec.versaoDados = NFE_VERSAO_XML

                'NFeRetAutorizacao.nfeCabecMsgValue = NfeCabec

                Dim sURL As String
                sURL = ""
                Call WS_Obter_URL(sURL, gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO, gobjApp.objEstado.Sigla, "NFeRetAutorizacao", gobjApp.gsModelo)

                NFeRetAutorizacao.Url = sURL

                NFeRetAutorizacao.ClientCertificates.Add(gobjApp.cert)

                Select Case WS_Obter_Autorizador(gobjApp.objEstado.Sigla)

                    'Case "PR"
                    '    xmlNode1 = NFeRetAutorizacao.NfeRetAutorizacao(DocDados1)

                    Case Else
                        xmlNode1 = NFeRetAutorizacao.nfeRetAutorizacaoLote(DocDados1)

                End Select

                XMLStringRetConsReciNFE = xmlNode1.OuterXml

                gobjApp.sErro = "14"
                gobjApp.sMsg1 = "vai tratar a resposta a consulta"

                xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringRetConsReciNFE)

                XMLStreamRet = New MemoryStream(10000)
                XMLStreamRet.Write(xRet, 0, xRet.Length)

                Dim mySerializerRetConsReciNFe As New XmlSerializer(GetType(TRetConsReciNFe))

                Dim objconsRetReciNFe As TRetConsReciNFe = New TRetConsReciNFe

                gobjApp.sErro = "15"
                gobjApp.sMsg1 = "vai deserializar a resposta da consulta do lote"

                XMLStreamRet.Position = 0

                objconsRetReciNFe = mySerializerRetConsReciNFe.Deserialize(XMLStreamRet)

                iResult = gobjApp.dbDadosNfe.ExecuteCommand("INSERT INTO NFeFedRetConsReci ( Lote, FilialEmpresa, versao, tpAmb, verAplic, nRec, cStat, xMotivo, cUF, Data, Hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10} )", _
                lLote, iFilialEmpresa, objconsRetReciNFe.versao, objconsRetReciNFe.tpAmb, objconsRetReciNFe.verAplic, objconsRetReciNFe.nRec, objconsRetReciNFe.cStat, objconsRetReciNFe.xMotivo, objconsRetReciNFe.cUF, Now.Date, TimeOfDay.ToOADate)

                gobjApp.sErro = "16"
                gobjApp.sMsg1 = "vai tratar as notas do lote"

                '                If 1 = 2 Then
                If Not objconsRetReciNFe.protNFe Is Nothing Then


                    gobjApp.sErro = "17"
                    gobjApp.sMsg1 = "vai tratar as notas consultadas do lote"

                    For i = 0 To objconsRetReciNFe.protNFe.Count - 1

                        If String.IsNullOrEmpty(objconsRetReciNFe.protNFe(i).infProt.nProt) Then
                            objconsRetReciNFe.protNFe(i).infProt.nProt = ""
                        End If

                        iAchou = 0

                        For j = 1 To colNumIntNFiscal.Count

                            lNumIntNFiscal = colNumIntNFiscal(j)

                            resNFiscal = gobjApp.dbDadosNfe.ExecuteQuery(Of NFiscal) _
                            ("SELECT * FROM NFiscal WHERE NumINtDoc = {0}", lNumIntNFiscal)

                            objNFiscal = resNFiscal(0)

                            If Format(CInt(Serie_Sem_E(objNFiscal.Serie)), "000") = Mid(objconsRetReciNFe.protNFe(i).infProt.chNFe, 23, 3) And _
                               Format(objNFiscal.NumNotaFiscal, "000000000") = Mid(objconsRetReciNFe.protNFe(i).infProt.chNFe, 26, 9) Then
                                iAchou = 1
                                Exit For
                            End If

                        Next

                        If iAchou = 0 Then
                            Throw New System.Exception("A nota consultada nao foi encontrada nas notas que estao no lote. Serie = " & Mid(objconsRetReciNFe.protNFe(i).infProt.chNFe, 23, 3) & " Numero = " & Mid(objconsRetReciNFe.protNFe(i).infProt.chNFe, 26, 9))
                        End If

                        Call gobjApp.GravarLog("Iniciando a consulta da Nota Fiscal " & objNFiscal.NumNotaFiscal & ".", lLote, 0)

                        If objconsRetReciNFe.protNFe(i).infProt.cStat = "100" Or objconsRetReciNFe.protNFe(i).infProt.cStat = "150" Then


                            '******** pega o ultimo retorno de envio do lote em questao *************
                            resNFeFedProtNFe = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedProtNFe) _
                            ("SELECT * FROM NFeFedProtNFe WHERE (cStat = '100' Or cStat = '150') AND  chNFe = {0} AND FilialEmpresa = {1} AND NumIntNF = {2}", objconsRetReciNFe.protNFe(i).infProt.chNFe, iFilialEmpresa, lNumIntNFiscal)

                            If resNFeFedProtNFe.Count = 0 Then

                                If objconsRetReciNFe.protNFe(i).versao <> "1.10" And objconsRetReciNFe.protNFe(i).versao <> "2.00" And objconsRetReciNFe.protNFe(i).versao <> "3.10" Then
                                    Throw New System.Exception("Versao não tratada. Versao = " & objconsRetReciNFe.protNFe(i).versao)
                                End If

                                iResult = gobjApp.dbDadosNfe.ExecuteCommand("INSERT INTO NFeFedProtNFe ( FilialEmpresa, NumIntNF, versao, nRec, tpAmb, verAplic, chNFe, nProt, cStat, xMotivo, Data, Hora, DataRegistro, HoraRegistro) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13} )", _
                                iFilialEmpresa, lNumIntNFiscal, objconsRetReciNFe.protNFe(i).versao, objconsRetReciNFe.nRec, objconsRetReciNFe.protNFe(i).infProt.tpAmb, objconsRetReciNFe.protNFe(i).infProt.verAplic, objconsRetReciNFe.protNFe(i).infProt.chNFe, objconsRetReciNFe.protNFe(i).infProt.nProt, objconsRetReciNFe.protNFe(i).infProt.cStat, objconsRetReciNFe.protNFe(i).infProt.xMotivo, UTCParaDate(objconsRetReciNFe.protNFe(i).infProt.dhRecbto), UTCParaHora(objconsRetReciNFe.protNFe(i).infProt.dhRecbto), Now.Date, TimeOfDay.ToOADate)

                            End If


                            'se o arquivo ainda nao existir
                            If Dir(gobjApp.sDirXml & objconsRetReciNFe.protNFe(i).infProt.chNFe & "-procNfe.xml") <> objconsRetReciNFe.protNFe(i).infProt.chNFe & "-procNfe.xml" Then

                                Dim DocDados2 As XmlDocument = New XmlDocument
                                Dim XMLStreamDados1 = New MemoryStream(10000)
                                Dim sArquivo As String
                                Dim XMLString As String
                                Dim XMLString3 As String
                                Dim XMLStreamDados = New MemoryStream(10000)

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
                                ''sArquivo = sDir & objNFe.infNFe.Id & ".xml"
                                '                                    DocDados3.Save(sArquivo)
                                sArquivo = gobjApp.sDirXml & objconsRetReciNFe.protNFe(i).infProt.chNFe & "-procNfe.xml"

                                Dim writer As New XmlTextWriter(sArquivo, Nothing)

                                writer.Formatting = Formatting.None
                                DocDados3.WriteTo(writer)

                                writer.Close()

                            End If

                            Call gobjApp.GravarLog("A Nota Fiscal " & objNFiscal.NumNotaFiscal & " foi consultada com sucesso. Chave = " & objconsRetReciNFe.protNFe(i).infProt.chNFe, lLote, 0)

                            'se for duplicidade pesquisa a chave e faz consulta
                        ElseIf objconsRetReciNFe.protNFe(i).infProt.cStat = "204" Or objconsRetReciNFe.protNFe(i).infProt.cStat = "539" Then

                            Dim colArq As New Collection

                            sChaveNFe = Left(objconsRetReciNFe.protNFe(i).infProt.chNFe, 34)

                            sArq = Dir(gobjApp.sDirXml & sChaveNFe & "*-pre.xml")

                            Do While sArq <> ""

                                colArq.Add(sArq)

                                sArq = Dir()

                            Loop


                            For Each sArq In colArq

                                sChaveNFe = Left(sArq, 44)
                                Sleep(10)
                                lErro = objConsultaNFe.Consulta_NFe(sEmpresa, sChaveNFe, iFilialEmpresa, 0)

                                If lErro = SUCESSO Then Exit For

                            Next

                            'se for uma nota denegada vai guardar a informacao em NFeFedDenegada
                        ElseIf (objconsRetReciNFe.protNFe(i).infProt.cStat = "205" Or objconsRetReciNFe.protNFe(i).infProt.cStat = "110" Or objconsRetReciNFe.protNFe(i).infProt.cStat = "301" Or objconsRetReciNFe.protNFe(i).infProt.cStat = "302") Then


                            resNfeFedDenegada = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedDenegada) _
                            ("SELECT * FROM NFeFedDenegada WHERE NumIntNF = {0} ", objNFiscal.NumIntDoc)

                            If resNfeFedDenegada.Count = 0 Then

                                iResult = gobjApp.dbDadosNfe.ExecuteCommand("INSERT INTO NFeFedDenegada ( NumIntNF, Processada) VALUES ( {0}, {1})", _
                                objNFiscal.NumIntDoc, 0)

                            End If

                            '******** pega o ultimo retorno de envio do lote em questao *************
                            resNFeFedProtNFe = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedProtNFe) _
                            ("SELECT * FROM NFeFedProtNFe WHERE (cStat = '205' Or cStat = '110' Or cStat = '301' Or  Or cStat = '302') AND  chNFe = {0} AND FilialEmpresa = {1} AND NumIntNF = {2}", objconsRetReciNFe.protNFe(i).infProt.chNFe, iFilialEmpresa, lNumIntNFiscal)

                            If resNFeFedProtNFe.Count = 0 Then

                                If objconsRetReciNFe.protNFe(i).versao <> "1.10" And objconsRetReciNFe.protNFe(i).versao <> "2.00" And objconsRetReciNFe.protNFe(i).versao <> "3.10" Then
                                    Throw New System.Exception("Versao não tratada. Versao = " & objconsRetReciNFe.protNFe(i).versao)
                                End If

                                iResult = gobjApp.dbDadosNfe.ExecuteCommand("INSERT INTO NFeFedProtNFe ( FilialEmpresa, NumIntNF, versao, nRec, tpAmb, verAplic, chNFe, nProt, cStat, xMotivo, Data, Hora, DataRegistro, HoraRegistro) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13} )", _
                                iFilialEmpresa, lNumIntNFiscal, objconsRetReciNFe.protNFe(i).versao, objconsRetReciNFe.nRec, objconsRetReciNFe.protNFe(i).infProt.tpAmb, objconsRetReciNFe.protNFe(i).infProt.verAplic, objconsRetReciNFe.protNFe(i).infProt.chNFe, objconsRetReciNFe.protNFe(i).infProt.nProt, objconsRetReciNFe.protNFe(i).infProt.cStat, objconsRetReciNFe.protNFe(i).infProt.xMotivo, UTCParaDate(objconsRetReciNFe.protNFe(i).infProt.dhRecbto), UTCParaHora(objconsRetReciNFe.protNFe(i).infProt.dhRecbto), Now.Date, TimeOfDay.ToOADate)

                            End If

                            Call gobjApp.GravarLog(objconsRetReciNFe.protNFe(i).infProt.xMotivo, lLote, 0)

                        Else

                            Call gobjApp.GravarLog(objconsRetReciNFe.protNFe(i).infProt.xMotivo, lLote, 0)

                        End If

                    Next

                Else

                    'If gobjApp.iDebug = 1 Then gobjApp.sErro ="18"
                End If


                Call gobjApp.GravarLog("Consulta do lote " & CStr(lLote) & " realizada com sucesso.", lLote, 0)

            Else

                Call gobjApp.GravarLog(Replace("o Lote " & CStr(lLote) & " não tinha sido processado ainda - motivo = " & objNFeFedRetEnvi.xMotivo, "'", "*"), 0, 0)

            End If

            gobjApp.DadosCommit()

            Consulta_Lote_NFe = SUCESSO

        Catch ex As Exception

            Consulta_Lote_NFe = 1

            Call gobjApp.GravarLog("ERRO - a consulta do lote " & CStr(lLote) & " foi encerrado por erro. Erro = " & ex.Message, lLote, lNumIntNFiscal)

            Call gobjApp.GravarLog(gobjApp.sErro & " " & gobjApp.sMsg1 & " NumIntNF = " & lNumIntNFiscal, lLote, lNumIntNFiscal)

        Finally

            If gobjApp.iDebug = 1 Then Call MsgBox(gobjApp.sErro)

        End Try

    End Function

End Class
