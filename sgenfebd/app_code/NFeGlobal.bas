Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Math
Imports System.Text
Imports System.Xml
Imports System.Xml.Schema
Imports Microsoft.Win32
Imports System.Security.Cryptography.X509Certificates
Imports sgenfebd4.NFeXsd

Module NFeGlobal

    Public gobjApp As New ClassGlobalApp

    Public Const QRCODE_PROVISORIO = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx qrcode provisorio xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" 'depois de assinar será substituido pelo valor certo

    Public Const NFE_VERSAO_XML = "4.00"

    Public Const DATA_NULA As Date = #9/7/1822#

    Public Const NFE_AMBIENTE_HOMOLOGACAO As Integer = 2
    Public Const NFE_AMBIENTE_PRODUCAO As Integer = 1
    Public Const OPERACAO_CANCELAMENTO As Integer = 1

    Public Const TRIB_TIPO_CALCULO_VALOR = 0
    Public Const TRIB_TIPO_CALCULO_PERCENTUAL = 1

    Public Const SUCESSO As Integer = 0
    Public Const TIPODOC_TRIB_NF = 0

    Public Const TIPO_ORIGEM_ITEMNF = 1
    Public Const MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE = 7
    Public Const TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE = 0
    Public Const PRODUTO_RASTRO_NENHUM = 0

    Public Const IMPORTCOMPL_ORIGEM_NF = 1
    Public Const IMPORTCOMPL_TIPO_AFRMM = 6

    Public Const DELTA_HORA = 0.00001

    Public Const DOCINFO_NFIEDV = 24

    Public Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

    Public Declare Unicode Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Int32, _
        ByVal lpFileName As String) As Int32

    Public Function App_Iniciar() As Long

    End Function

    Public Sub App_Terminar()

    End Sub

    Public Sub CalculaDV_Modulo11(ByVal sString As String, ByRef iDigito As Integer)
        Dim iIndice As Integer
        Dim iMult As Integer
        Dim iTotal As Integer

        iMult = 2

        For iIndice = Len(sString) To 1 Step -1

            iTotal = iTotal + (Mid(sString, iIndice, 1) * iMult)

            If iMult = 9 Then
                iMult = 2
            Else
                iMult = iMult + 1
            End If

        Next

        iDigito = iTotal Mod 11

        iDigito = 11 - iDigito

        If iDigito > 9 Then iDigito = 0

    End Sub

    Public Sub Formata_String_Numero(ByVal sStringRecebe As String, ByRef sStringRetorna As String)

        Dim iTamanho As Integer
        Dim sCaracter As String
        Dim iIndice As Integer

        iTamanho = Len(Trim(sStringRecebe))

        sStringRetorna = ""

        For iIndice = 1 To iTamanho

            sCaracter = Mid(sStringRecebe, iIndice, 1)

            If IsNumeric(sCaracter) Then
                sStringRetorna = sStringRetorna & sCaracter
            End If

        Next

    End Sub

    Public Sub Formata_Sem_Espaco(ByVal sStringRecebe As String, ByRef sStringRetorna As String)

        Dim iTamanho As Integer
        Dim sCaracter As String
        Dim iIndice As Integer

        iTamanho = Len(Trim(sStringRecebe))

        sStringRetorna = ""

        For iIndice = 1 To iTamanho

            sCaracter = Mid(sStringRecebe, iIndice, 1)

            If sCaracter = " " Then
                sStringRetorna = sStringRetorna & "_"
            Else
                sStringRetorna = sStringRetorna & sCaracter
            End If

        Next

    End Sub

    Public Function Fomata_ZerosEsquerda(ByVal sTexto As String, ByVal iTam As Integer) As String
        'completa com zeros nao significativos

        Dim sAux As String

        sAux = Right(StrDup(iTam, "0") & Trim(sTexto), iTam)

        Fomata_ZerosEsquerda = sAux

    End Function

    Public Function DesacentuaTexto(ByVal sTexto As String) As String

        'retorna uma copia do texto com a troca dos caracteres acentuados por nao acentuados

        Dim iIndice As Integer
        Dim sCaracter As String
        Dim sGuardaTexto As String
        Dim iCodigo As Integer

        sGuardaTexto = ""

        sTexto = Trim(sTexto)

        'Para cada Caracter do Texto
        For iIndice = 1 To Len(sTexto)

            'Seleciona caracter da posição iIndice
            sCaracter = Mid(sTexto, iIndice, 1)

            'Pega codigo ASC do caracter da selecionado acima
            iCodigo = Asc(sCaracter)

            'Verifica se caracter é acentuado
            Select Case iCodigo

                Case 10, 13
                    If iIndice <> 1 And iIndice <> Len(sTexto) Then
                        sCaracter = " "
                    Else
                        sCaracter = ""
                    End If

                Case 1 To 31
                    sCaracter = ""

                Case 32
                    If iIndice = 1 Or iIndice = Len(sTexto) Then
                        sCaracter = ""
                    End If

                Case 186
                    sCaracter = "."

                Case 192 To 197
                    sCaracter = Chr(65)

                Case 199
                    sCaracter = Chr(67)

                Case 200 To 203
                    sCaracter = Chr(69)

                Case 204 To 207
                    sCaracter = Chr(73)

                Case 210 To 214
                    sCaracter = Chr(79)

                Case 217 To 220
                    sCaracter = Chr(85)

                Case 224 To 229
                    sCaracter = Chr(97)

                Case 231
                    sCaracter = Chr(99)

                Case 232 To 235
                    sCaracter = Chr(101)

                Case 236 To 239
                    sCaracter = Chr(105)

                Case 242 To 246
                    sCaracter = Chr(111)

                Case 249 To 252
                    sCaracter = Chr(117)



            End Select

            If sCaracter <> "." And sCaracter <> "/" And sCaracter <> "-" Then
                sGuardaTexto = sGuardaTexto & sCaracter
            End If

        Next

        DesacentuaTexto = Trim(sGuardaTexto)


    End Function



    Public Function Serie_Sem_E(ByVal sSerie As String) As String
        'retira -E da serie

        Dim sSerieNova As String
        Dim iPos As Integer

        iPos = InStr(sSerie, "-e")

        If iPos <> 0 Then
            sSerieNova = Mid(sSerie, 1, iPos - 1)
        Else
            sSerieNova = sSerie
        End If

        Serie_Sem_E = sSerieNova


    End Function

    Public Function Arredonda_Moeda(ByVal dValor As Double, Optional ByVal iNumDigitos As Integer = 2) As Double

        If dValor >= 0 Then
            Arredonda_Moeda = Round(dValor + 0.0000000001, iNumDigitos)
        Else
            Arredonda_Moeda = Round(dValor - 0.0000000001, iNumDigitos)
        End If

    End Function

    Public Sub Salva_Arquivo(ByVal DocDados1 As XmlDocument, ByVal XMLString4 As String)

        '****************  salva o arquivo 

        Dim XMLStreamDados As MemoryStream = New MemoryStream(10000)
        Dim XMLStreamDados1 As MemoryStream = New MemoryStream(10000)

        Dim xDados10 As Byte()

        xDados10 = System.Text.Encoding.UTF8.GetBytes(XMLString4)

        XMLStreamDados.Write(xDados10, 0, xDados10.Length)

        XMLStreamDados.Position = 0
        DocDados1.Load(XMLStreamDados)

        Dim writer1 As New XmlTextWriter(XMLStreamDados1, Nothing)

        writer1.Formatting = Formatting.None
        DocDados1.WriteTo(writer1)
        writer1.Flush()
        XMLStreamDados1.Position = 0
        DocDados1.Load(XMLStreamDados1)

    End Sub

    Public Function NFeXml_Conv_Valor(ByVal objValor As Object) As Double
        If Not (objValor Is Nothing) Then
            NFeXml_Conv_Valor = CDbl(Replace(objValor, ".", ","))
        Else
            NFeXml_Conv_Valor = 0
        End If
    End Function

    Public Function NFeXml_Conv_Perc(ByVal objValor As Object) As Double
        If Not (objValor Is Nothing) Then
            NFeXml_Conv_Perc = CDbl(Replace(objValor, ".", ",")) / 100
        Else
            NFeXml_Conv_Perc = 0
        End If
    End Function

    Public Function NFeXml_Conv_Texto(ByVal objValor As Object) As String
        If Not (objValor Is Nothing) Then
            NFeXml_Conv_Texto = objValor.ToString
        Else
            NFeXml_Conv_Texto = ""
        End If
    End Function

    Public Function NFeXml_Conv_Data(ByVal objValor As Object) As Date
        If Not (objValor Is Nothing) Then
            NFeXml_Conv_Data = objValor
        Else
            NFeXml_Conv_Data = DATA_NULA
        End If
    End Function

    Public Function NFeXml_Conv_Long(ByVal objValor As Object) As Long
        If Not (objValor Is Nothing) Then
            If IsNumeric(objValor) Then
                NFeXml_Conv_Long = objValor
            Else
                NFeXml_Conv_Long = 0
            End If
        Else
            NFeXml_Conv_Long = 0
        End If
    End Function

    Public Function GetXmlAttrNameFromEnumValue(Of T)(ByVal pEnumVal As T) As String

        Dim type As Type = pEnumVal.GetType
        Dim info As System.Reflection.FieldInfo = type.GetField([Enum].GetName(GetType(T), pEnumVal))
        If info.GetCustomAttributes(GetType(XmlEnumAttribute), False).Length = 0 Then

            Return pEnumVal.ToString

        Else

            Dim att As XmlEnumAttribute = CType(info.GetCustomAttributes(GetType(XmlEnumAttribute), False)(0), XmlEnumAttribute) 'If there is an xmlattribute defined, return the name
            Return att.Name

        End If

    End Function

    Public Function GetCode(Of T)(ByVal value As String) As T
        For Each o As Object In System.Enum.GetValues(GetType(T))
            Dim enumValue As T = CType(o, T)
            If GetXmlAttrNameFromEnumValue(Of T)(enumValue).Equals(value, StringComparison.OrdinalIgnoreCase) Then
                Return CType(o, T)
            End If
        Next

        Throw New ArgumentException("No code exists for type " + GetType(T).ToString() + " corresponding to value of " + value)
    End Function

    Private Sub WS_Obter_URL_Producao(ByRef sURL As String, ByVal sMetodo As String, ByVal sAutorizador As String, ByVal sModelo As String)

        sURL = ""

        Select Case sAutorizador

            Case "AM"

                Select Case sMetodo

                    Case "RecepcaoEvento" : sURL = "https://nfe.sefaz.am.gov.br/services2/services/RecepcaoEvento4"
                    Case "NfeInutilizacao" : sURL = "https://nfe.sefaz.am.gov.br/services2/services/NfeInutilizacao4"
                    Case "NfeConsultaProtocolo" : sURL = "https://nfe.sefaz.am.gov.br/services2/services/NfeConsulta4"
                    Case "NfeStatusServico" : sURL = "https://nfe.sefaz.am.gov.br/services2/services/NfeStatusServico4"
                    Case "NfeConsultaCadastro" : sURL = "https://nfe.sefaz.am.gov.br/services2/services/cadconsultacadastro4"
                    Case "NFeAutorizacao" : sURL = "https://nfe.sefaz.am.gov.br/services2/services/NfeAutorizacao4"
                    Case "NFeRetAutorizacao" : sURL = "https://nfe.sefaz.am.gov.br/services2/services/NfeRetAutorizacao4"

                End Select

            Case "BA"

                Select Case sMetodo

                    Case "NfeInutilizacao" : sURL = "https://nfe.sefaz.ba.gov.br/webservices/NFeInutilizacao4/NFeInutilizacao4.asmx"
                    Case "NfeConsultaCadastro" : sURL = "https://nfe.sefaz.ba.gov.br/webservices/CadConsultaCadastro4/CadConsultaCadastro4.asmx"
                    Case "RecepcaoEvento" : sURL = "	https://nfe.sefaz.ba.gov.br/webservices/NFeRecepcaoEvento4/NFeRecepcaoEvento4.asmx"
                    Case "NfeConsultaProtocolo" : sURL = "https://nfe.sefaz.ba.gov.br/webservices/NFeConsultaProtocolo4/NFeConsultaProtocolo4.asmx"
                    Case "NfeStatusServico" : sURL = "	https://nfe.sefaz.ba.gov.br/webservices/NFeStatusServico4/NFeStatusServico4.asmx"
                    Case "NFeAutorizacao" : sURL = "	https://nfe.sefaz.ba.gov.br/webservices/NFeAutorizacao4/NFeAutorizacao4.asmx"
                    Case "NFeRetAutorizacao" : sURL = "https://nfe.sefaz.ba.gov.br/webservices/NFeRetAutorizacao4/NFeRetAutorizacao4.asmx"

                End Select

            Case "CE"

                Select Case sMetodo

                    Case "NfeInutilizacao" : sURL = "https://nfe.sefaz.ce.gov.br/nfe4/services/NFeInutilizacao4?wsdl"
                    Case "NfeConsultaProtocolo" : sURL = "https://nfe.sefaz.ce.gov.br/nfe4/services/NFeConsultaProtocolo4?wsdl"
                    Case "NfeStatusServico" : sURL = "https://nfe.sefaz.ce.gov.br/nfe4/services/NFeStatusServico4?wsdl"
                    Case "NfeConsultaCadastro" : sURL = "https://nfe.sefaz.ce.gov.br/nfe4/services/CadConsultaCadastro4?wsdl"
                    Case "RecepcaoEvento" : sURL = "https://nfe.sefaz.ce.gov.br/nfe4/services/NFeRecepcaoEvento4?wsdl"
                    Case "NFeAutorizacao" : sURL = "https://nfe.sefaz.ce.gov.br/nfe4/services/NFeAutorizacao4?wsdl"
                    Case "NFeRetAutorizacao" : sURL = "https://nfe.sefaz.ce.gov.br/nfe4/services/NFeRetAutorizacao4?wsdl"

                End Select

            Case "GO"

                Select Case sMetodo

                    Case "RecepcaoEvento" : sURL = "https://nfe.sefaz.go.gov.br/nfe/services/NFeRecepcaoEvento4?wsdl"
                    Case "NfeInutilizacao" : sURL = "https://nfe.sefaz.go.gov.br/nfe/services/NFeInutilizacao4?wsdl"
                    Case "NfeConsultaProtocolo" : sURL = "https://nfe.sefaz.go.gov.br/nfe/services/NFeConsultaProtocolo4?wsdl"
                    Case "NfeStatusServico" : sURL = "https://nfe.sefaz.go.gov.br/nfe/services/NFeStatusServico4?wsdl"
                    Case "NfeConsultaCadastro" : sURL = "https://nfe.sefaz.go.gov.br/nfe/services/CadConsultaCadastro4?wsdl"
                    Case "NFeAutorizacao" : sURL = "https://nfe.sefaz.go.gov.br/nfe/services/NFeAutorizacao4?wsdl"
                    Case "NFeRetAutorizacao" : sURL = "https://nfe.sefaz.go.gov.br/nfe/services/NFeRetAutorizacao4?wsdl"

                End Select

            Case "MG"

                Select Case sMetodo

                    Case "NfeInutilizacao" : sURL = "https://nfe.fazenda.mg.gov.br/nfe2/services/NFeInutilizacao4"
                    Case "NfeConsultaProtocolo" : sURL = "https://nfe.fazenda.mg.gov.br/nfe2/services/NFeConsultaProtocolo4"
                    Case "NfeStatusServico" : sURL = "https://nfe.fazenda.mg.gov.br/nfe2/services/NFeStatusServico4"
                        'Case "NfeConsultaCadastro" : sURL = "https://nfe.fazenda.mg.gov.br/nfe2/services/cadconsultacadastro2"
                    Case "RecepcaoEvento" : sURL = "https://nfe.fazenda.mg.gov.br/nfe2/services/NFeRecepcaoEvento4"
                    Case "NFeAutorizacao" : sURL = "https://nfe.fazenda.mg.gov.br/nfe2/services/NFeAutorizacao4"
                    Case "NFeRetAutorizacao" : sURL = "https://nfe.fazenda.mg.gov.br/nfe2/services/NFeRetAutorizacao4"

                End Select

            Case "MS"

                Select Case sMetodo

                    Case "RecepcaoEvento" : sURL = "https://nfe.fazenda.ms.gov.br/ws/NFeRecepcaoEvento4"
                    Case "NfeConsultaCadastro" : sURL = "https://nfe.fazenda.ms.gov.br/ws/CadConsultaCadastro4"
                    Case "NfeInutilizacao" : sURL = "https://nfe.fazenda.ms.gov.br/ws/NFeInutilizacao4"
                    Case "NfeConsultaProtocolo" : sURL = "https://nfe.fazenda.ms.gov.br/ws/NFeConsultaProtocolo4"
                    Case "NfeStatusServico" : sURL = "https://nfe.fazenda.ms.gov.br/ws/NFeStatusServico4"
                    Case "NFeAutorizacao" : sURL = "https://nfe.fazenda.ms.gov.br/ws/NFeAutorizacao4"
                    Case "NFeRetAutorizacao" : sURL = "https://nfe.fazenda.ms.gov.br/ws/NFeRetAutorizacao4"

                End Select

            Case "MT"

                Select Case sMetodo

                    Case "NfeInutilizacao" : sURL = "https://nfe.sefaz.mt.gov.br/nfews/v2/services/NfeInutilizacao4?wsdl"
                    Case "NfeConsultaProtocolo" : sURL = "https://nfe.sefaz.mt.gov.br/nfews/v2/services/NfeConsulta4?wsdl"
                    Case "NfeStatusServico" : sURL = "https://nfe.sefaz.mt.gov.br/nfews/v2/services/NfeStatusServico4?wsdl"
                    Case "NfeConsultaCadastro" : sURL = "https://nfe.sefaz.mt.gov.br/nfews/v2/services/CadConsultaCadastro4?wsdl"
                    Case "RecepcaoEvento" : sURL = "https://nfe.sefaz.mt.gov.br/nfews/v2/services/RecepcaoEvento4?wsdl"
                    Case "NFeAutorizacao" : sURL = "https://nfe.sefaz.mt.gov.br/nfews/v2/services/NfeAutorizacao4?wsdl"
                    Case "NFeRetAutorizacao" : sURL = "https://nfe.sefaz.mt.gov.br/nfews/v2/services/NfeRetAutorizacao4?wsdl"

                End Select

            Case "PE"

                Select Case sMetodo

                    Case "RecepcaoEvento" : sURL = "https://nfe.sefaz.pe.gov.br/nfe-service/services/RecepcaoEvento4"
                    Case "NfeInutilizacao" : sURL = "https://nfe.sefaz.pe.gov.br/nfe-service/services/NFeInutilizacao4"
                    Case "NfeConsultaProtocolo" : sURL = "https://nfe.sefaz.pe.gov.br/nfe-service/services/NfeConsulta4"
                    Case "NfeStatusServico" : sURL = "https://nfe.sefaz.pe.gov.br/nfe-service/services/NfeStatusServico4"
                    Case "NfeConsultaCadastro" : sURL = "https://nfe.sefaz.pe.gov.br/nfe-service/services/CadConsultaCadastro4"
                    Case "NFeAutorizacao" : sURL = "https://nfe.sefaz.pe.gov.br/nfe-service/services/NFeAutorizacao4"
                    Case "NFeRetAutorizacao" : sURL = "https://nfe.sefaz.pe.gov.br/nfe-service/services/NFeRetAutorizacao4"

                End Select

            Case "PR"

                Select Case sMetodo

                    Case "NfeInutilizacao" : sURL = "https://nfe.sefa.pr.gov.br/nfe/NFeInutilizacao4?wsdl"
                    Case "NfeConsultaProtocolo" : sURL = "https://nfe.sefa.pr.gov.br/nfe/NFeConsultaProtocolo4?wsdl"
                    Case "NfeStatusServico" : sURL = "https://nfe.sefa.pr.gov.br/nfe/NFeStatusServico4?wsdl"
                    Case "NfeConsultaCadastro" : sURL = "https://nfe.sefa.pr.gov.br/nfe/CadConsultaCadastro4?wsdl"
                    Case "RecepcaoEvento" : sURL = "https://nfe.sefa.pr.gov.br/nfe/NFeRecepcaoEvento4?wsdl"
                    Case "NFeAutorizacao" : sURL = "https://nfe.sefa.pr.gov.br/nfe/NFeAutorizacao4?wsdl"
                    Case "NFeRetAutorizacao" : sURL = "https://nfe.sefa.pr.gov.br/nfe/NFeRetAutorizacao4?wsdl"

                End Select

            Case "RS"

                If sModelo = "NFCe" Then

                    Select Case sMetodo

                        Case "RecepcaoEvento" : sURL = "https://nfce.sefazrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx"
                        Case "NfeInutilizacao" : sURL = "https://nfce.sefazrs.rs.gov.br/ws/nfeinutilizacao/nfeinutilizacao4.asmx"
                        Case "NfeConsultaProtocolo" : sURL = "https://nfce.sefazrs.rs.gov.br/ws/NfeConsulta/NfeConsulta4.asmx"
                        Case "NfeStatusServico" : sURL = "https://nfce.sefazrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx"
                        Case "NFeAutorizacao" : sURL = "https://nfce.sefazrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        Case "NFeRetAutorizacao" : sURL = "https://nfce.sefazrs.rs.gov.br/ws/NfeRetAutorizacao/NFeRetAutorizacao4.asmx"

                    End Select

                Else

                    Select Case sMetodo

                        Case "RecepcaoEvento" : sURL = "https://nfe.sefazrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx"
                        Case "NfeConsultaCadastro" : sURL = "https://cad.sefazrs.rs.gov.br/ws/cadconsultacadastro/cadconsultacadastro4.asmx"
                        Case "NfeConsultaDest" : sURL = "https://nfe.sefazrs.rs.gov.br/ws/nfeConsultaDest/nfeConsultaDest.asmx"
                        Case "NfeDownloadNF" : sURL = "https://nfe.sefazrs.rs.gov.br/ws/nfeDownloadNF/nfeDownloadNF.asmx"
                        Case "NfeInutilizacao" : sURL = "https://nfe.sefazrs.rs.gov.br/ws/nfeinutilizacao/nfeinutilizacao4.asmx"
                        Case "NfeConsultaProtocolo" : sURL = "https://nfe.sefazrs.rs.gov.br/ws/NfeConsulta/NfeConsulta4.asmx"
                        Case "NfeStatusServico" : sURL = "https://nfe.sefazrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx"
                        Case "NFeAutorizacao" : sURL = "https://nfe.sefazrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        Case "NFeRetAutorizacao" : sURL = "https://nfe.sefazrs.rs.gov.br/ws/NfeRetAutorizacao/NFeRetAutorizacao4.asmx"

                    End Select

                End If

            Case "SP"

                If sModelo = "NFCe" Then

                    Select Case sMetodo

                        Case "NfeInutilizacao" : sURL = "https://nfce.fazenda.sp.gov.br/ws/nfeinutilizacao4.asmx"
                        Case "NfeConsultaProtocolo" : sURL = "https://nfce.fazenda.sp.gov.br/ws/nfeconsultaprotocolo4.asmx"
                        Case "NfeStatusServico" : sURL = "https://nfce.fazenda.sp.gov.br/ws/nfestatusservico4.asmx"
                        Case "NfeConsultaCadastro" : sURL = "https://nfce.fazenda.sp.gov.br/ws/cadconsultacadastro4.asmx"
                        Case "RecepcaoEvento" : sURL = "https://nfce.fazenda.sp.gov.br/ws/nferecepcaoevento4.asmx"
                        Case "NFeAutorizacao" : sURL = "https://nfce.fazenda.sp.gov.br/ws/NFeAutorizacao4.asmx"
                        Case "NFeRetAutorizacao" : sURL = "https://nfce.fazenda.sp.gov.br/ws/nferetautorizacao4.asmx"

                    End Select

                Else

                    Select Case sMetodo

                        Case "NfeInutilizacao" : sURL = "https://nfe.fazenda.sp.gov.br/ws/nfeinutilizacao4.asmx"
                        Case "NfeConsultaProtocolo" : sURL = "https://nfe.fazenda.sp.gov.br/ws/nfeconsultaprotocolo4.asmx"
                        Case "NfeStatusServico" : sURL = "https://nfe.fazenda.sp.gov.br/ws/nfestatusservico4.asmx"
                        Case "NfeConsultaCadastro" : sURL = "https://nfe.fazenda.sp.gov.br/ws/cadconsultacadastro4.asmx"
                        Case "RecepcaoEvento" : sURL = "https://nfe.fazenda.sp.gov.br/ws/nferecepcaoevento4.asmx"
                        Case "NFeAutorizacao" : sURL = "https://nfe.fazenda.sp.gov.br/ws/nfeautorizacao4.asmx"
                        Case "NFeRetAutorizacao" : sURL = "https://nfe.fazenda.sp.gov.br/ws/nferetautorizacao4.asmx"

                    End Select

                End If

            Case "SVAN" 'Sefaz Virtual Ambiente Nacional

                Select Case sMetodo

                    Case "RecepcaoEvento" : sURL = "https://www.sefazvirtual.fazenda.gov.br/NFeRecepcaoEvento4/NFeRecepcaoEvento4.asmx"
                    Case "NfeInutilizacao" : sURL = "https://www.sefazvirtual.fazenda.gov.br/NFeInutilizacao4/NFeInutilizacao4.asmx"
                    Case "NfeConsultaProtocolo" : sURL = "https://www.sefazvirtual.fazenda.gov.br/NFeConsultaProtocolo4/NFeConsultaProtocolo4.asmx"
                    Case "NfeStatusServico" : sURL = "https://www.sefazvirtual.fazenda.gov.br/NFeStatusServico4/NFeStatusServico4.asmx"
                    Case "NfeDownloadNF" : sURL = "https://www.sefazvirtual.fazenda.gov.br/NfeDownloadNF/NfeDownloadNF.asmx"
                    Case "NFeAutorizacao" : sURL = "https://www.sefazvirtual.fazenda.gov.br/NFeAutorizacao4/NFeAutorizacao4.asmx"
                    Case "NFeRetAutorizacao" : sURL = "https://www.sefazvirtual.fazenda.gov.br/NFeRetAutorizacao4/NFeRetAutorizacao4.asmx"

                End Select

            Case "SVRS" 'Sefaz Virtual Rio Grande do Sul

                If sModelo = "NFCe" Then

                    Select Case sMetodo

                        Case "RecepcaoEvento" : sURL = "https://nfce.svrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx"
                        Case "NfeInutilizacao" : sURL = "https://nfce.svrs.rs.gov.br/ws/nfeinutilizacao/nfeinutilizacao4.asmx"
                        Case "NfeConsultaProtocolo" : sURL = "https://nfce.svrs.rs.gov.br/ws/NfeConsulta/NfeConsulta4.asmx"
                        Case "NfeStatusServico" : sURL = "https://nfce.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx"
                        Case "NFeAutorizacao" : sURL = "https://nfce.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        Case "NFeRetAutorizacao" : sURL = "https://nfce.svrs.rs.gov.br/ws/NfeRetAutorizacao/NFeRetAutorizacao4.asmx"

                    End Select

                Else

                    Select Case sMetodo

                        Case "RecepcaoEvento" : sURL = "https://nfe.svrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx"
                        Case "NfeConsultaCadastro" : sURL = "https://cad.svrs.rs.gov.br/ws/cadconsultacadastro/cadconsultacadastro4.asmx"
                        Case "NfeInutilizacao" : sURL = "https://nfe.svrs.rs.gov.br/ws/nfeinutilizacao/nfeinutilizacao4.asmx"
                        Case "NfeConsultaProtocolo" : sURL = "https://nfe.svrs.rs.gov.br/ws/NfeConsulta/NfeConsulta4.asmx"
                        Case "NfeStatusServico" : sURL = "https://nfe.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx"
                        Case "NFeAutorizacao" : sURL = "https://nfe.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        Case "NFeRetAutorizacao" : sURL = "https://nfe.svrs.rs.gov.br/ws/NfeRetAutorizacao/NFeRetAutorizacao4.asmx"

                    End Select

                End If

                'Case "SCAN"

                '    Select Case sMetodo

                '        Case "RecepcaoEvento" : sURL = "https://www.scan.fazenda.gov.br/RecepcaoEvento/RecepcaoEvento.asmx"
                '        Case "NfeInutilizacao" : sURL = "https://www.scan.fazenda.gov.br/NfeInutilizacao2/NfeInutilizacao2.asmx"
                '        Case "NfeConsultaProtocolo" : sURL = "https://www.scan.fazenda.gov.br/NfeConsulta2/NfeConsulta2.asmx"
                '        Case "NfeStatusServico" : sURL = "https://www.scan.fazenda.gov.br/NfeStatusServico2/NfeStatusServico2.asmx"
                '        Case "NFeAutorizacao" : sURL = "https://www.scan.fazenda.gov.br/NfeAutorizacao/NfeAutorizacao.asmx"
                '        Case "NFeRetAutorizacao" : sURL = "https://www.scan.fazenda.gov.br/NfeRetAutorizacao/NfeRetAutorizacao.asmx"

                '    End Select

            Case "SVC-AN" 'Sefaz Virtual de Contingência Ambiente Nacional

                Select Case sMetodo

                    Case "RecepcaoEvento" : sURL = "https://www.svc.fazenda.gov.br/NFeRecepcaoEvento4/NFeRecepcaoEvento4.asmx"
                    Case "NfeConsultaProtocolo" : sURL = "https://www.svc.fazenda.gov.br/NFeConsultaProtocolo4/NFeConsultaProtocolo4.asmx"
                    Case "NfeStatusServico" : sURL = "https://www.svc.fazenda.gov.br/NFeStatusServico4/NFeStatusServico4.asmx"
                    Case "NFeAutorizacao" : sURL = "	https://www.svc.fazenda.gov.br/NFeAutorizacao4/NFeAutorizacao4.asmx"
                    Case "NFeRetAutorizacao" : sURL = "https://www.svc.fazenda.gov.br/NFeRetAutorizacao4/NFeRetAutorizacao4.asmx"

                End Select

            Case "SVC-RS" 'Sefaz Virtual de Contingência Rio Grande do Sul

                Select Case sMetodo

                    Case "RecepcaoEvento" : sURL = "https://nfe.svrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx"
                    Case "NfeConsultaProtocolo" : sURL = "https://nfe.svrs.rs.gov.br/ws/NfeConsulta/NfeConsulta4.asmx"
                    Case "NfeStatusServico" : sURL = "https://nfe.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx"
                    Case "NFeAutorizacao" : sURL = "https://nfe.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                    Case "NFeRetAutorizacao" : sURL = "https://nfe.svrs.rs.gov.br/ws/NfeRetAutorizacao/NFeRetAutorizacao4.asmx"

                End Select

            Case "AN" 'Ambiente Nacional

                Select Case sMetodo

                    Case "RecepcaoEvento" : sURL = "https://www.nfe.fazenda.gov.br/NFeRecepcaoEvento4/NFeRecepcaoEvento4.asmx"
                    Case "NfeConsultaDest" : sURL = "https://www.nfe.fazenda.gov.br/NFeConsultaDest/NFeConsultaDest.asmx"
                    Case "NfeDownloadNF" : sURL = "https://www.nfe.fazenda.gov.br/NfeDownloadNF/NfeDownloadNF.asmx"

                End Select

        End Select

    End Sub

    Private Sub WS_Obter_URL_Homologacao(ByRef sURL As String, ByVal sMetodo As String, ByVal sAutorizador As String, ByVal sModelo As String)

        sURL = ""

        Select Case sAutorizador

            Case "AM"

                Select Case sMetodo

                    Case "RecepcaoEvento" : sURL = "https://homnfe.sefaz.am.gov.br/services2/services/RecepcaoEvento4"
                    Case "NfeInutilizacao" : sURL = "https://homnfe.sefaz.am.gov.br/services2/services/NfeInutilizacao4"
                    Case "NfeConsultaProtocolo" : sURL = "https://homnfe.sefaz.am.gov.br/services2/services/NfeConsulta4"
                    Case "NfeStatusServico" : sURL = "https://homnfe.sefaz.am.gov.br/services2/services/NfeStatusServico4"
                    Case "NfeConsultaCadastro" : sURL = "https://homnfe.sefaz.am.gov.br/services2/services/cadconsultacadastro4"
                    Case "NFeAutorizacao" : sURL = "https://homnfe.sefaz.am.gov.br/services2/services/NfeAutorizacao4"
                    Case "NFeRetAutorizacao" : sURL = "https://homnfe.sefaz.am.gov.br/services2/services/NfeRetAutorizacao4"

                End Select

            Case "BA"

                Select Case sMetodo

                    Case "NfeInutilizacao" : sURL = "https://hnfe.sefaz.ba.gov.br/webservices/NFeInutilizacao4/NFeInutilizacao4.asmx"
                    Case "NfeConsultaCadastro" : sURL = "https://hnfe.sefaz.ba.gov.br/webservices/CadConsultaCadastro4/CadConsultaCadastro4.asmx"
                    Case "RecepcaoEvento" : sURL = "https://hnfe.sefaz.ba.gov.br/webservices/NFeRecepcaoEvento4/NFeRecepcaoEvento4.asmx"
                    Case "NfeConsultaProtocolo" : sURL = "https://hnfe.sefaz.ba.gov.br/webservices/NFeConsultaProtocolo4/NFeConsultaProtocolo4.asmx"
                    Case "NfeStatusServico" : sURL = "https://hnfe.sefaz.ba.gov.br/webservices/NFeStatusServico4/NFeStatusServico4.asmx"
                    Case "NFeAutorizacao" : sURL = "https://hnfe.sefaz.ba.gov.br/webservices/NFeAutorizacao4/NFeAutorizacao4.asmx"
                    Case "NFeRetAutorizacao" : sURL = "https://hnfe.sefaz.ba.gov.br/webservices/NFeRetAutorizacao4/NFeRetAutorizacao4.asmx"

                End Select


            Case "CE"

                Select Case sMetodo
                    Case "NfeInutilizacao" : sURL = "	https://nfeh.sefaz.ce.gov.br/nfe4/services/NFeInutilizacao4?WSDL"
                    Case "NfeConsultaProtocolo" : sURL = "	https://nfeh.sefaz.ce.gov.br/nfe4/services/NFeConsultaProtocolo4?WSDL"
                    Case "NfeStatusServico" : sURL = "https://nfeh.sefaz.ce.gov.br/nfe4/services/NFeStatusServico4?WSDL"
                    Case "NfeConsultaCadastro" : sURL = "https://nfeh.sefaz.ce.gov.br/nfe4/services/CadConsultaCadastro4?WSDL"
                    Case "RecepcaoEvento" : sURL = "https://nfeh.sefaz.ce.gov.br/nfe4/services/NFeRecepcaoEvento4?WSDL"
                    Case "NFeAutorizacao" : sURL = "https://nfeh.sefaz.ce.gov.br/nfe4/services/NFeAutorizacao4?WSDL"
                    Case "NFeRetAutorizacao" : sURL = "https://nfeh.sefaz.ce.gov.br/nfe4/services/NFeRetAutorizacao4?WSDL"

                End Select

            Case "GO"

                Select Case sMetodo

                    Case "RecepcaoEvento" : sURL = "https://homolog.sefaz.go.gov.br/nfe/services/RecepcaoEvento4?wsdl"
                    Case "NfeInutilizacao" : sURL = "https://homolog.sefaz.go.gov.br/nfe/services/NFeInutilizacao4?wsdl"
                    Case "NfeConsultaProtocolo" : sURL = "https://homolog.sefaz.go.gov.br/nfe/services/NFeConsultaProtocolo4?wsdl"
                    Case "NfeStatusServico" : sURL = "https://homolog.sefaz.go.gov.br/nfe/services/NFeStatusServico4?wsdl"
                    Case "NfeConsultaCadastro" : sURL = "https://homolog.sefaz.go.gov.br/nfe/services/CadConsultaCadastro4?wsdl"
                    Case "NFeAutorizacao" : sURL = "https://homolog.sefaz.go.gov.br/nfe/services/NFeAutorizacao4?wsdl"
                    Case "NFeRetAutorizacao" : sURL = "https://homolog.sefaz.go.gov.br/nfe/services/NFeRetAutorizacao4?wsdl"

                End Select

            Case "MG"

                Select Case sMetodo

                    Case "NfeInutilizacao" : sURL = "https://hnfe.fazenda.mg.gov.br/nfe2/services/NfeInutilizacao4"
                    Case "NfeConsultaProtocolo" : sURL = "https://hnfe.fazenda.mg.gov.br/nfe2/services/NfeConsulta4"
                    Case "NfeStatusServico" : sURL = "https://hnfe.fazenda.mg.gov.br/nfe2/services/NfeStatusServico4"
                    Case "NFeAutorizacao" : sURL = "https://hnfe.fazenda.mg.gov.br/nfe2/services/NfeAutorizacao4"
                    Case "NFeRetAutorizacao" : sURL = "https://hnfe.fazenda.mg.gov.br/nfe2/services/NfeRetAutorizacao4"

                    Case "RecepcaoEvento" : sURL = "https://hnfe.fazenda.mg.gov.br/nfe2/services/RecepcaoEvento"
                    Case "NfeConsultaCadastro" : sURL = "https://hnfe.fazenda.mg.gov.br/nfe2/services/cadconsultacadastro2"

                End Select

            Case "MS"

                Select Case sMetodo

                    Case "RecepcaoEvento" : sURL = "https://homologacao.nfe.ms.gov.br/ws/NFeRecepcaoEvento4"
                    Case "NfeInutilizacao" : sURL = "https://homologacao.nfe.ms.gov.br/ws/NFeInutilizacao4"
                    Case "NfeConsultaProtocolo" : sURL = "https://homologacao.nfe.ms.gov.br/ws/NFeConsultaProtocolo4"
                    Case "NfeStatusServico" : sURL = "https://homologacao.nfe.ms.gov.br/ws/NFeStatusServico4"
                    Case "NFeAutorizacao" : sURL = "https://homologacao.nfe.ms.gov.br/ws/NFeAutorizacao4"
                    Case "NFeRetAutorizacao" : sURL = "https://homologacao.nfe.ms.gov.br/ws/NFeRetAutorizacao4"

                    Case "NfeConsultaCadastro" : sURL = "https://homologacao.nfe.ms.gov.br/homologacao/services2/CadConsultaCadastro2"

                End Select

            Case "MT"

                Select Case sMetodo

                    Case "NfeInutilizacao" : sURL = "https://homologacao.sefaz.mt.gov.br/nfews/v2/services/NfeInutilizacao4?wsdl"
                    Case "NfeConsultaProtocolo" : sURL = "https://homologacao.sefaz.mt.gov.br/nfews/v2/services/NfeConsulta4?wsdl"
                    Case "NfeStatusServico" : sURL = "https://homologacao.sefaz.mt.gov.br/nfews/v2/services/NfeStatusServico4?wsdl"
                    Case "RecepcaoEvento" : sURL = "https://homologacao.sefaz.mt.gov.br/nfews/v2/services/RecepcaoEvento4?wsdl"
                    Case "NfeConsultaCadastro" : sURL = "https://homologacao.sefaz.mt.gov.br/nfews/v2/services/CadConsultaCadastro4?wsdl"
                    Case "NFeAutorizacao" : sURL = "https://homologacao.sefaz.mt.gov.br/nfews/v2/services/NfeAutorizacao4?wsdl"
                    Case "NFeRetAutorizacao" : sURL = "https://homologacao.sefaz.mt.gov.br/nfews/v2/services/NfeRetAutorizacao4?wsdl"

                End Select

            Case "PE"

                Select Case sMetodo

                    Case "RecepcaoEvento" : sURL = "https://nfehomolog.sefaz.pe.gov.br/nfe-service/services/NFeRecepcaoEvento4"
                    Case "NfeInutilizacao" : sURL = "https://nfehomolog.sefaz.pe.gov.br/nfe-service/services/NfeInutilizacao4"
                    Case "NfeConsultaProtocolo" : sURL = "https://nfehomolog.sefaz.pe.gov.br/nfe-service/services/NfeConsulta4"
                    Case "NfeStatusServico" : sURL = "https://nfehomolog.sefaz.pe.gov.br/nfe-service/services/NfeStatusServico4"
                    Case "NFeAutorizacao" : sURL = "https://nfehomolog.sefaz.pe.gov.br/nfe-service/services/NFeAutorizacao4"
                    Case "NFeRetAutorizacao" : sURL = "https://nfehomolog.sefaz.pe.gov.br/nfe-service/services/NFeRetAutorizacao4"

                End Select

            Case "PR"

                Select Case sMetodo

                    Case "NfeInutilizacao" : sURL = "https://homologacao.nfe.sefa.pr.gov.br/nfe/NFeInutilizacao4"
                    Case "NfeConsultaProtocolo" : sURL = "https://homologacao.nfe.sefa.pr.gov.br/nfe/NFeConsultaProtocolo4"
                    Case "NfeStatusServico" : sURL = "https://homologacao.nfe.sefa.pr.gov.br/nfe/NFeStatusServico4"
                    Case "NfeConsultaCadastro" : sURL = "https://homologacao.nfe.sefa.pr.gov.br/nfe/CadConsultaCadastro4"
                    Case "RecepcaoEvento" : sURL = "https://homologacao.nfe.sefa.pr.gov.br/nfe/NFeRecepcaoEvento4"
                    Case "NFeAutorizacao" : sURL = "https://homologacao.nfe.sefa.pr.gov.br/nfe/NFeAutorizacao4"
                    Case "NFeRetAutorizacao" : sURL = "https://homologacao.nfe.sefa.pr.gov.br/nfe/NFeRetAutorizacao4"

                End Select

            Case "RS"

                If sModelo = "NFCe" Then

                    Select Case sMetodo

                        Case "RecepcaoEvento" : sURL = "https://nfce-homologacao.sefazrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx"
                        Case "NfeInutilizacao" : sURL = "https://nfce-homologacao.sefazrs.rs.gov.br/ws/nfeinutilizacao/nfeinutilizacao4.asmx"
                        Case "NfeConsultaProtocolo" : sURL = "https://nfce-homologacao.sefazrs.rs.gov.br/ws/NfeConsulta/NfeConsulta4.asmx"
                        Case "NfeStatusServico" : sURL = "https://nfce-homologacao.sefazrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx"
                        Case "NFeAutorizacao" : sURL = "https://nfce-homologacao.sefazrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        Case "NFeRetAutorizacao" : sURL = "https://nfce-homologacao.sefazrs.rs.gov.br/ws/NfeRetAutorizacao/NFeRetAutorizacao4.asmx"

                    End Select

                Else

                    Select Case sMetodo

                        Case "RecepcaoEvento" : sURL = "https://nfe-homologacao.sefazrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx"
                        Case "NfeInutilizacao" : sURL = "https://nfe-homologacao.sefazrs.rs.gov.br/ws/nfeinutilizacao/nfeinutilizacao4.asmx"
                        Case "NfeConsultaProtocolo" : sURL = "	https://nfe-homologacao.sefazrs.rs.gov.br/ws/NfeConsulta/NfeConsulta4.asmx"
                        Case "NfeStatusServico" : sURL = "https://nfe-homologacao.sefazrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx"
                        Case "NFeAutorizacao" : sURL = "https://nfe-homologacao.sefazrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        Case "NFeRetAutorizacao" : sURL = "https://nfe-homologacao.sefazrs.rs.gov.br/ws/NfeRetAutorizacao/NFeRetAutorizacao4.asmx"

                        Case "NfeConsultaCadastro" : sURL = "https://cad.sefazrs.rs.gov.br/ws/cadconsultacadastro/cadconsultacadastro2.asmx"
                        Case "NfeConsultaDest" : sURL = "https://nfe-homologacao.sefazrs.rs.gov.br/ws/nfeConsultaDest/nfeConsultaDest.asmx"
                        Case "NfeDownloadNF" : sURL = "https://nfe-homologacao.sefazrs.rs.gov.br/ws/nfeDownloadNF/nfeDownloadNF.asmx"
                    End Select

                End If

            Case "SP"

                If sModelo = "NFCe" Then

                    Select Case sMetodo

                        Case "NfeInutilizacao" : sURL = "https://homologacao.nfce.fazenda.sp.gov.br/ws/nfeinutilizacao4.asmx"
                        Case "NfeConsultaProtocolo" : sURL = "https://homologacao.nfce.fazenda.sp.gov.br/ws/nfeconsultaprotocolo4.asmx"
                        Case "NfeStatusServico" : sURL = "https://homologacao.nfce.fazenda.sp.gov.br/ws/nfestatusservico4.asmx"
                        Case "NfeConsultaCadastro" : sURL = "https://homologacao.nfce.fazenda.sp.gov.br/ws/cadconsultacadastro4.asmx"
                        Case "RecepcaoEvento" : sURL = "https://homologacao.nfce.fazenda.sp.gov.br/ws/nferecepcaoevento4.asmx"
                        Case "NFeAutorizacao" : sURL = "https://homologacao.nfce.fazenda.sp.gov.br/ws/nfeautorizacao4.asmx"
                        Case "NFeRetAutorizacao" : sURL = "https://homologacao.nfce.fazenda.sp.gov.br/ws/nferetautorizacao4.asmx"

                        Case "RecepcaoEPEC" : sURL = "https://homologacao.nfce.epec.fazenda.sp.gov.br/EPECws/RecepcaoEPEC.asmx"

                    End Select

                Else

                    Select Case sMetodo

                        Case "NfeInutilizacao" : sURL = "https://homologacao.nfe.fazenda.sp.gov.br/ws/nfeinutilizacao4.asmx"
                        Case "NfeConsultaProtocolo" : sURL = "https://homologacao.nfe.fazenda.sp.gov.br/ws/nfeconsultaprotocolo4.asmx"
                        Case "NfeStatusServico" : sURL = "https://homologacao.nfe.fazenda.sp.gov.br/ws/nfestatusservico4.asmx"
                        Case "NfeConsultaCadastro" : sURL = "https://homologacao.nfe.fazenda.sp.gov.br/ws/cadconsultacadastro4.asmx"
                        Case "RecepcaoEvento" : sURL = "https://homologacao.nfe.fazenda.sp.gov.br/ws/nferecepcaoevento4.asmx"
                        Case "NFeAutorizacao" : sURL = "https://homologacao.nfe.fazenda.sp.gov.br/ws/nfeautorizacao4.asmx"
                        Case "NFeRetAutorizacao" : sURL = "https://homologacao.nfe.fazenda.sp.gov.br/ws/nferetautorizacao4.asmx"

                    End Select

                End If

            Case "SVAN"

                Select Case sMetodo

                    Case "RecepcaoEvento" : sURL = "https://hom.sefazvirtual.fazenda.gov.br/NFeRecepcaoEvento4/NFeRecepcaoEvento4.asmx"
                    Case "NfeConsultaCadastro" : sURL = ""
                    Case "NfeInutilizacao" : sURL = "https://hom.sefazvirtual.fazenda.gov.br/NFeInutilizacao4/NFeInutilizacao4.asmx"
                    Case "NfeConsultaProtocolo" : sURL = "https://hom.sefazvirtual.fazenda.gov.br/NFeConsultaProtocolo4/NFeConsultaProtocolo4.asmx"
                    Case "NfeStatusServico" : sURL = "https://hom.sefazvirtual.fazenda.gov.br/NFeStatusServico4/NFeStatusServico4.asmx"
                    Case "NFeAutorizacao" : sURL = "https://hom.sefazvirtual.fazenda.gov.br/NFeAutorizacao4/NFeAutorizacao4.asmx"
                    Case "NFeRetAutorizacao" : sURL = "https://hom.sefazvirtual.fazenda.gov.br/NFeRetAutorizacao4/NFeRetAutorizacao4.asmx"

                End Select

            Case "SVRS"

                If sModelo = "NFCe" Then

                    Select Case sMetodo

                        Case "RecepcaoEvento" : sURL = "https://nfce-homologacao.svrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx"
                        Case "NfeInutilizacao" : sURL = "https://nfce-homologacao.svrs.rs.gov.br/ws/nfeinutilizacao/nfeinutilizacao4.asmx"
                        Case "NfeConsultaProtocolo" : sURL = "	https://nfce-homologacao.svrs.rs.gov.br/ws/NfeConsulta/NfeConsulta4.asmx"
                        Case "NfeStatusServico" : sURL = "	https://nfce-homologacao.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx"
                        Case "NFeAutorizacao" : sURL = "	https://nfce-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        Case "NFeRetAutorizacao" : sURL = "https://nfce-homologacao.svrs.rs.gov.br/ws/NfeRetAutorizacao/NFeRetAutorizacao4.asmx"

                    End Select

                Else

                    Select Case sMetodo

                        Case "RecepcaoEvento" : sURL = "https://nfe-homologacao.svrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx"
                        Case "NfeInutilizacao" : sURL = "https://nfe-homologacao.svrs.rs.gov.br/ws/nfeinutilizacao/nfeinutilizacao4.asmx"
                        Case "NfeConsultaProtocolo" : sURL = "	https://nfe-homologacao.svrs.rs.gov.br/ws/NfeConsulta/NfeConsulta4.asmx"
                        Case "NfeStatusServico" : sURL = "	https://nfe-homologacao.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx"
                        Case "NFeAutorizacao" : sURL = "	https://nfe-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                        Case "NFeRetAutorizacao" : sURL = "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeRetAutorizacao/NFeRetAutorizacao4.asmx"

                        Case "NfeConsultaCadastro" : sURL = "https://cad-homologacao.svrs.rs.gov.br/ws/cadconsultacadastro/cadconsultacadastro2.asmx"

                    End Select

                End If

                'Case "SCAN"

                '    Select Case sMetodo

                '        Case "RecepcaoEvento" : sURL = "https://hom.nfe.fazenda.gov.br/SCAN/RecepcaoEvento/RecepcaoEvento.asmx"
                '        Case "NfeInutilizacao" : sURL = "https://hom.nfe.fazenda.gov.br/SCAN/NfeInutilizacao2/NfeInutilizacao2.asmx"
                '        Case "NfeConsultaProtocolo" : sURL = "https://hom.nfe.fazenda.gov.br/SCAN/NfeConsulta2/NfeConsulta2.asmx"
                '        Case "NfeStatusServico" : sURL = "https://hom.nfe.fazenda.gov.br/SCAN/NfeStatusServico2/NfeStatusServico2.asmx"
                '        Case "NFeAutorizacao" : sURL = "https://hom.nfe.fazenda.gov.br/SCAN/NfeAutorizacao/NfeAutorizacao.asmx"
                '        Case "NFeRetAutorizacao" : sURL = "https://hom.nfe.fazenda.gov.br/SCAN/NfeRetAutorizacao/NfeRetAutorizacao.asmx"

                '    End Select

            Case "SVC-AN" 'Sefaz Virtual de Contingência Ambiente Nacional

                Select Case sMetodo

                    Case "RecepcaoEvento" : sURL = "https://hom.svc.fazenda.gov.br/NFeRecepcaoEvento4/NFeRecepcaoEvento4.asmx"
                    Case "NfeConsultaProtocolo" : sURL = "https://hom.svc.fazenda.gov.br/NFeConsultaProtocolo4/NFeConsultaProtocolo4.asmx"
                    Case "NfeStatusServico" : sURL = "https://hom.svc.fazenda.gov.br/NFeStatusServico4/NFeStatusServico4.asmx"
                    Case "NFeAutorizacao" : sURL = "https://hom.svc.fazenda.gov.br/NFeAutorizacao4/NFeAutorizacao4.asmx"
                    Case "NFeRetAutorizacao" : sURL = "https://hom.svc.fazenda.gov.br/NFeRetAutorizacao4/NFeRetAutorizacao4.asmx"

                    Case "NfeRecepcao" : sURL = "https://hom.svc.fazenda.gov.br/NfeRecepcao2/NfeRecepcao2.asmx"
                    Case "NfeRetRecepcao" : sURL = "https://hom.svc.fazenda.gov.br/NfeRetRecepcao2/NfeRetRecepcao2.asmx"

                End Select

            Case "SVC-RS" 'Sefaz Virtual de Contingência Rio Grande do Sul

                Select Case sMetodo

                    Case "RecepcaoEvento" : sURL = "	https://nfe-homologacao.svrs.rs.gov.br/ws/recepcaoevento/recepcaoevento4.asmx"
                    Case "NfeConsultaProtocolo" : sURL = "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeConsulta/NfeConsulta4.asmx"
                    Case "NfeStatusServico" : sURL = "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeStatusServico/NfeStatusServico4.asmx"
                        'Case "NfeInutilizacao" : sURL = "https://nfe-homologacao.svrs.rs.gov.br/ws/nfeinutilizacao/nfeinutilizacao2.asmx"
                    Case "NFeAutorizacao" : sURL = "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeAutorizacao/NFeAutorizacao4.asmx"
                    Case "NFeRetAutorizacao" : sURL = "https://nfe-homologacao.svrs.rs.gov.br/ws/NfeRetAutorizacao/NFeRetAutorizacao4.asmx"

                End Select

            Case "AN" 'Ambiente(Nacional

                Select Case sMetodo

                    Case "RecepcaoEvento" : sURL = "https://hom.nfe.fazenda.gov.br/NFeRecepcaoEvento4/NFeRecepcaoEvento4.asmx"
                    Case "NfeConsultaDest" : sURL = "https://hom.nfe.fazenda.gov.br/NFeConsultaDest/NFeConsultaDest.asmx"
                    Case "NfeDownloadNF" : sURL = "https://hom.nfe.fazenda.gov.br/NfeDownloadNF/NfeDownloadNF.asmx"

                End Select

        End Select

    End Sub

    Public Function DataHoraParaUTC(ByVal dtData As Date, ByVal dHora As Double) As String
        Dim date1 As Date

        Try
            If dtData <> DATA_NULA Then

                date1 = New Date(dtData.Year, dtData.Month, dtData.Day, Date.FromOADate(dHora).Hour, Date.FromOADate(dHora).Minute, Date.FromOADate(dHora).Second)
                DataHoraParaUTC = date1.ToString("yyyy-MM-ddTHH:mm:sszzz")
            Else
                DataHoraParaUTC = ""
            End If
        Catch
            DataHoraParaUTC = ""
        End Try

    End Function

    Public Function DataParaUTC(ByVal dtData As Date) As String
        Try
            If dtData = DATA_NULA Then
                DataParaUTC = ""
            Else
                DataParaUTC = dtData.ToString("yyyy-MM-ddTHH:mm:sszzz")
            End If
        Catch
            DataParaUTC = ""
        End Try
    End Function

    Public Function UTCParaHora(ByVal sUTC As String) As Double

        Dim localDateTime As System.DateTime

        Try
            localDateTime = System.DateTime.Parse(sUTC).ToLocalTime
            UTCParaHora = localDateTime.TimeOfDay.TotalDays
        Catch ex As Exception
            UTCParaHora = 0
        End Try

    End Function

    Public Function UTCParaDate(ByVal sUTC As String) As Date
        Dim localDateTime As System.DateTime

        Try
            localDateTime = System.DateTime.Parse(sUTC).ToLocalTime
            UTCParaDate = localDateTime.Date
        Catch ex As Exception
            UTCParaDate = DATA_NULA
        End Try

    End Function

    Public Function WS_Obter_Autorizador(ByVal sUF As String, Optional ByVal sMetodo As String = "") As String

        'UF que utilizam a SVAN - Sefaz Virtual do Ambiente Nacional: MA, PA 
        'UF que utilizam a SVRS - Sefaz Virtual do RS: 
        '- Para serviço de Consulta Cadastro: AC, RN, PB, SC 
        '- Para demais serviços relacionados com o sistema da NF-e: AC, AL, AP, DF, ES, PB, PI, RJ, RN, RO, RR, SC, SE, TO 
        'Autorizadores em contingência: 
        '- UF que utilizam a SVC-AN - Sefaz Virtual de Contingência Ambiente Nacional: AC, AL, AP, DF, ES, MG, PB, RJ, RN, RO, RR, RS, SC, SE, SP, TO 
        '- UF que utilizam a SVC-RS - Sefaz Virtual de Contingência Rio Grande do Sul: AM, BA, CE, GO, MA, MS, MT, PA, PE, PI, PR
        'Autorizadores: AMBACEGOMGMSMTPEPRRSSPSVANSVRSSVC(-ANSVC - RSAN)

        Dim sAutorizador As String = ""

        If gobjApp.sSistemaContingencia = "SCAN" Then

            sAutorizador = "SCAN"

        Else

            If gobjApp.sSistemaContingencia = "" Then

                Select Case sUF

                    Case "MA", "PA" ', "PI"
                        sAutorizador = "SVAN"

                    Case "AC", "AL", "AP", "DF", "ES", "PB", "PI", "RJ", "RN", "RO", "RR", "SC", "SE", "TO"
                        If sMetodo = "NfeConsultaCadastro" And sUF <> "AC" And sUF <> "RN" And sUF <> "PB" And sUF <> "SC" Then
                            sAutorizador = ""
                        Else
                            sAutorizador = "SVRS"
                        End If

                    Case Else
                        sAutorizador = sUF

                End Select

            Else


                Select Case sUF

                    Case "AC", "AL", "AP", "DF", "ES", "MG", "PB", "RJ", "RN", "RO", "RR", "RS", "SC", "SE", "SP", "TO"
                        sAutorizador = "SVC-AN"

                    Case "AM", "BA", "CE", "GO", "MA", "MS", "MT", "PA", "PE", "PI", "PR"
                        sAutorizador = "SVC-RS"

                End Select

                'sAutorizador = gobjApp.sSistemaContingencia

            End If

        End If

        WS_Obter_Autorizador = sAutorizador

    End Function

    Public Sub WS_Obter_URL(ByRef sURL As String, ByVal bHomologacao As Boolean, ByVal sUF As String, ByVal sMetodo As String, ByVal sModelo As String)
        'sModelo: "NFe" ou "NFCe"

        Dim sAutorizador As String

        sAutorizador = WS_Obter_Autorizador(sUF, sMetodo)

        If bHomologacao Then

            Call WS_Obter_URL_Homologacao(sURL, sMetodo, sAutorizador, sModelo)

        Else

            Call WS_Obter_URL_Producao(sURL, sMetodo, sAutorizador, sModelo)

        End If

    End Sub

    Public Function ClonarEstruturasSerializaveis(ByRef objDestino As Object, ByVal objOrigem As Object) As Long

        Dim XMLStream = New MemoryStream(10000)

        Try

            Dim mySerializerOrig As New XmlSerializer(objOrigem.GetType)

            mySerializerOrig.Serialize(XMLStream, objOrigem)

            Dim mySerializerDest As New XmlSerializer(objDestino.GetType)

            XMLStream.Position = 0
            objDestino = mySerializerDest.Deserialize(XMLStream)

            ClonarEstruturasSerializaveis = 0

        Catch ex As Exception

            ClonarEstruturasSerializaveis = 1

        End Try

    End Function

    Public Function Texto_Para_Hexa(ByVal sTexto As String, Optional ByVal bLower As Boolean = True) As String

        Dim sTextoHexa As String = ""

        Try

            Dim sVal As String

            While sTexto.Length > 0
                sVal = Conversion.Hex(Strings.Asc(sTexto.Substring(0, 1).ToString()))
                If bLower Then sVal = sVal.ToLower
                sTexto = sTexto.Substring(1, sTexto.Length - 1)
                sTextoHexa = sTextoHexa & sVal
            End While

        Catch ex As Exception

        Finally

            Texto_Para_Hexa = sTextoHexa

        End Try

    End Function

    Public Function Texto_Para_SHA1(ByVal sTexto As String, Optional ByVal bLower As Boolean = True) As String

        Dim sTextoSHA1 As String = ""

        Try

            Dim sha1Obj As New Security.Cryptography.SHA1CryptoServiceProvider
            Dim bytesToHash() As Byte = System.Text.Encoding.ASCII.GetBytes(sTexto)

            bytesToHash = sha1Obj.ComputeHash(bytesToHash)

            For Each b As Byte In bytesToHash
                sTextoSHA1 += b.ToString("x2")
            Next

        Catch ex As Exception

        Finally

            Texto_Para_SHA1 = sTextoSHA1

        End Try

    End Function

    Public Function NFCE_Gera_QRCode(ByVal schNFe As String, ByVal sVersaoQRCode As String, ByVal stpAmb As String, ByVal scDest As String, ByVal sdhEmi As String, ByVal svNF As String, ByVal svICMS As String, ByVal sDigVal As String, ByVal sIdCSC As String, ByVal sCSC As String) As String

        Dim sQRCode As String = ""

        Try

            Dim sdhEmiHexa As String
            Dim sDigValHexa As String
            Dim sAux As String
            Dim sURL As String = ""

            sdhEmiHexa = Texto_Para_Hexa(sdhEmi)
            sDigValHexa = Texto_Para_Hexa(sDigVal)

            sQRCode = "chNFe=" & schNFe & "&nVersao=" & sVersaoQRCode & "&tpAmb=" & stpAmb

            If scDest <> "" Then sQRCode = sQRCode & "&cDest=" & scDest

            sQRCode = sQRCode & "&dhEmi=" & sdhEmiHexa & "&vNF=" & svNF & "&vICMS=" & svICMS & "&digVal=" & sDigValHexa & "&cIdToken=" & sIdCSC

            sAux = Texto_Para_SHA1(sQRCode & sCSC)

            sQRCode = sQRCode & "&cHashQRCode=" & sAux

            Select Case Left(schNFe, 2)

                Case "33" 'RJ

                    sURL = "http://www4.fazenda.rj.gov.br/consultaNFCe/QRCode?"

                Case "28"

                    sURL = "http://www.nfe.se.gov.br/portal/consultarNFCe.jsp?"

            End Select

            sQRCode = sURL & sQRCode

        Catch ex As Exception

        Finally

            NFCE_Gera_QRCode = sQRCode

        End Try

    End Function

    Public Function chNFe_Retorna_Modelo(ByVal chNFe As String) As String

        Dim sModelo As String

        sModelo = "NFe"

        If Len(Trim(chNFe)) >= 23 Then

            If Mid(chNFe, 21, 2) = "65" Then sModelo = "NFCe"

        End If

        chNFe_Retorna_Modelo = sModelo

    End Function

End Module