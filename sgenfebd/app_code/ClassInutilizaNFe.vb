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

Public Class ClassInutilizaNFe

    Public Function Inutiliza_NFe(ByVal sEmpresa As String, ByVal sSerie As String, ByVal sNumInicial As String, ByVal sNumFinal As String, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sMotivo As String, Optional ByVal iScan As Integer = -1) As Long


        Dim objInut As TInutNFe = New TInutNFe

        Dim sArquivo As String

        Dim XMLStream As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamRet As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamDados As MemoryStream = New MemoryStream(10000)

        Dim AD As AssinaturaDigital = New AssinaturaDigital
        Dim XMLString As String
        Dim XMLStringRetInutNFE As String

        Dim iResult As Integer

        Dim iPos As Integer
        Dim lErro As Long

        Dim objValidaXML As ClassValidaXML = New ClassValidaXML
        Dim xmlNode1 As XmlNode

        Dim resSerie As IEnumerable(Of Serie), objSerie As Serie, bNFCe As Boolean = False

        Try

            Call gobjApp.GravarLog("Iniciando a inutiliacao das notas da serie " & sSerie & " notas fiscais de: " & sNumInicial & " a " & sNumFinal, 0, 0)

            Dim NFeInutilizacao As New nfeinutilizacao2.NFeInutilizacao4
            'Dim NFeCabec As Object

            resSerie = gobjApp.dbDadosNfe.ExecuteQuery(Of Serie) _
            ("SELECT * FROM Serie WHERE FilialEmpresa = {0} AND Serie = {1}", iFilialEmpresa, sSerie & "-e")

            For Each objSerie In resSerie

                If objSerie.ModDocFis = 35 Then bNFCe = True
                Exit For

            Next

            objInut.versao = NFE_VERSAO_XML

            objInut.infInut = New TInutNFeInfInut


            If gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then
                objInut.infInut.tpAmb = TAmb.Item2
            Else
                objInut.infInut.tpAmb = TAmb.Item1
            End If
            objInut.infInut.xServ = TInutNFeInfInutXServ.INUTILIZAR

            objInut.infInut.cUF = GetCode(Of TCodUfIBGE)(CStr(gobjApp.objEstado.CodIBGE))

            objInut.infInut.Id = "ID" & gobjApp.objEstado.CodIBGE & Right(iAno, 2) & gobjApp.sCGC & IIf(bNFCe, "65", "55") & Format(CInt(sSerie), "000") & Format(CLng(sNumInicial), "000000000") & Format(CLng(sNumFinal), "000000000")

            If gobjApp.objEstado.Sigla = "PR" Then
                For iIndice = 1 To 41 - Len(objInut.infInut.Id)
                    objInut.infInut.Id = objInut.infInut.Id & "0"
                Next
            End If


            Application.DoEvents()

            objInut.infInut.ano = Right(CStr(iAno), 2)
            objInut.infInut.CNPJ = gobjApp.sCGC
            objInut.infInut.mod = IIf(bNFCe, TMod.Item65, TMod.Item55)
            objInut.infInut.serie = sSerie
            objInut.infInut.nNFIni = sNumInicial
            objInut.infInut.nNFFin = sNumFinal

            sMotivo = Replace(sMotivo, "_", " ")
            If sMotivo = "*" Then sMotivo = "motivo não fornecido"

            If Len(RTrim(sMotivo)) < 15 Then sMotivo = sMotivo & StrDup(15 - Len(RTrim(sMotivo)), "*")

            objInut.infInut.xJust = sMotivo

            Dim mySerializer As New XmlSerializer(GetType(TInutNFe))

            XMLStream = New MemoryStream(10000)

            mySerializer.Serialize(XMLStream, objInut)

            Dim xm As Byte()
            xm = XMLStream.ToArray

            XMLString = System.Text.Encoding.UTF8.GetString(xm)

            iPos = InStr(XMLString, "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""")

            If iPos <> 0 Then

                XMLString = Mid(XMLString, 1, iPos - 1) & Mid(XMLString, iPos + 99)

            End If

            lErro = AD.Assinar(XMLString, "infInut", gobjApp.cert, gobjApp.iDebug)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            Dim xMlD As XmlDocument

            xMlD = AD.XMLDocAssinado()

            Dim xString As String
            xString = AD.XMLStringAssinado

            xString = Mid(xString, 1, 19) & " encoding=""utf-8"" " & Mid(xString, 20)




            '************* valida dados antes do envio **********************
            Dim xDados As Byte()

            xDados = System.Text.Encoding.UTF8.GetBytes(xString)

            XMLStreamDados = New MemoryStream(10000)

            XMLStreamDados.Write(xDados, 0, xDados.Length)


            Dim DocDados As XmlDocument = New XmlDocument
            XMLStreamDados.Position = 0
            DocDados.Load(XMLStreamDados)
            sArquivo = gobjApp.sDirXml & Mid(objInut.infInut.Id, 3) & "-ped-inu.xml"
            DocDados.Save(sArquivo)

            lErro = objValidaXML.validaXML(sArquivo, gobjApp.sDirXsd & "InutNFe_v4.00.xsd", 0, 0, iFilialEmpresa)
            If lErro = 1 Then

                Call gobjApp.GravarLog("ERRO - a inutilizacao da serie " & sSerie & " notas fiscais de: " & sNumInicial & " a " & sNumFinal & " foi encerrado por erro.", 0, 0)

                Exit Try
            End If

            Dim DocDados1 As New XmlDocument

            Call Salva_Arquivo(DocDados1, xString)

            'NFeCabec.cUF = CStr(gobjApp.objEstado.CodIBGE)
            'NFeCabec.versaoDados = NFE_VERSAO_XML

            'NFeInutilizacao.nfeCabecMsgValue = NFeCabec

            Dim sURL As String
            sURL = ""
            Call WS_Obter_URL(sURL, gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO, gobjApp.objEstado.Sigla, "NfeInutilizacao", IIf(bNFCe, "NFCe", "NFe"))

            NFeInutilizacao.Url = sURL

            NFeInutilizacao.ClientCertificates.Add(gobjApp.cert)

            Select Case WS_Obter_Autorizador(gobjApp.objEstado.Sigla)

                'Case "PR"
                '    xmlNode1 = NFeInutilizacao.nfeInutilizacaoNF(DocDados1)

                Case Else
                    xmlNode1 = NFeInutilizacao.nfeInutilizacaoNF(DocDados1)

            End Select

            XMLStringRetInutNFE = xmlNode1.OuterXml

            Dim xRet As Byte()

            xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringRetInutNFE)

            XMLStreamRet.Write(xRet, 0, xRet.Length)

            Dim mySerializerRetEnvNFe As New XmlSerializer(GetType(TRetInutNFe))

            Dim objRetInutNFe As TRetInutNFe = New TRetInutNFe

            XMLStreamRet.Position = 0

            objRetInutNFe = mySerializerRetEnvNFe.Deserialize(XMLStreamRet)

            If objRetInutNFe.infInut.cStat = "102" Then

                iResult = gobjApp.dbDadosNfe.ExecuteCommand("INSERT INTO NFeFedRetInutNFe ( FilialEmpresa, versao, tpAmb, verAplic, cStat, xMotivo, cUF, ano, CNPJ, mod, serie, nNFIni, nNFFim, nProt, Data, Hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, (9), {10}, {11}, {12}, {13}, {14}, {15})", _
                iFilialEmpresa, objRetInutNFe.versao, objRetInutNFe.infInut.tpAmb, objRetInutNFe.infInut.verAplic, objRetInutNFe.infInut.cStat, objRetInutNFe.infInut.xMotivo, objRetInutNFe.infInut.cUF, objRetInutNFe.infInut.ano, objRetInutNFe.infInut.CNPJ, objRetInutNFe.infInut.mod, objRetInutNFe.infInut.serie, objRetInutNFe.infInut.nNFIni, objRetInutNFe.infInut.nNFFin, IIf(objRetInutNFe.infInut.nProt = "", "0", objRetInutNFe.infInut.nProt), UTCParaDate(objRetInutNFe.infInut.dhRecbto), UTCParaHora(objRetInutNFe.infInut.dhRecbto))

            Else

                iResult = gobjApp.dbDadosNfe.ExecuteCommand("INSERT INTO NFeFedRetInutNFe ( FilialEmpresa, versao, tpAmb, verAplic, cStat, xMotivo, cUF, serie, nNFIni, nNFFim, Data, Hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, (9), {10}, {11} )", _
                  iFilialEmpresa, objRetInutNFe.versao, objRetInutNFe.infInut.tpAmb, objRetInutNFe.infInut.verAplic, objRetInutNFe.infInut.cStat, objRetInutNFe.infInut.xMotivo, objRetInutNFe.infInut.cUF, sSerie, sNumInicial, sNumFinal, Now.Date, Now.TimeOfDay.TotalDays)

            End If

            gobjApp.DadosCommit()

            Call gobjApp.GravarLog(objRetInutNFe.infInut.xMotivo & "  Serie = " & sSerie & " NFIni = " & sNumInicial & " NFFim = " & sNumFinal, 0, 0)

            Inutiliza_NFe = SUCESSO

        Catch ex As Exception

            Inutiliza_NFe = 1

            iResult = gobjApp.dbDadosNfe.ExecuteCommand("INSERT INTO NFeFedRetInutNFe ( FilialEmpresa, versao, tpAmb, verAplic, cStat, xMotivo, cUF, Serie, NNFIni, nNFFim, Data, Hora) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, (9), {10}, {11} )", _
            iFilialEmpresa, "", 0, "", "", "as notas nao estao inutilizadas", "", sSerie, sNumInicial, sNumFinal, Now.Date, TimeOfDay.ToOADate)

            Call gobjApp.GravarLog("ERRO - a inutilizacao da serie " & sSerie & " notas fiscais de: " & sNumInicial & " a " & sNumFinal & " foi encerrado por erro. Erro = as notas NÂO estao inutilizadas", 0, 0)

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        Finally

        End Try

    End Function

End Class

