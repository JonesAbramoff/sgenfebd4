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

Public Class ClassMontaXmlNFe

    Private objNFiscal As NFeNFiscal
    Private lLote As Long
    Private lNumIntNF As Long
    Private sChaveNFe As String
    Private objTributacaoDoc As TributacaoDoc
    Private dTotalDescontoItem As Double
    Private dvProdICMS As Double
    Private dValorServPIS As Double
    Private dValorServCOFINS As Double
    Private dServNTribICMS As Double
    Private dValorPIS As Double
    Private dValorCOFINS As Double
    Private IIValorTotal As Double
    Private cMunFG As String
    Private bComISSQN As Boolean
    Private sModelo As String ' "NFe" ou "NFCe"
    Private dICMSBaseTotal As Double
    Private bICMSUFDest As Boolean 'se precisa incluir a parte nova interestadual para consumidor final criada com a EC 87/2015

    'para qrcode
    Private scDest As String
    Private sdhEmi As String
    Private svNF As String
    Private svICMS As String
    Private sidNFCECSC As String
    Private sNFCECSC As String

    Private Function Monta_NFiscal_Xml25(ByVal infNFe As TNFeInfNFe) As Long
        'preenche identificacao de pessoas autorizadas a acessar o xml

        Dim resInfoAdicDocAutoXml As IEnumerable(Of InfoAdicDocAutoXml)
        Dim resNFePadraoAutXml As IEnumerable(Of NFePadraoAutXml)
        Dim objInfoAdicDocAutoXml As InfoAdicDocAutoXml
        Dim objNFePadraoAutXml As NFePadraoAutXml

        Try

            resInfoAdicDocAutoXml = gobjApp.dbDadosNfe.ExecuteQuery(Of InfoAdicDocAutoXml) _
            ("SELECT * FROM InfoAdicDocAutoXml WHERE TipoDoc = {0} AND NumIntDoc = {1} ORDER BY Seq", 0, objNFiscal.NumIntDoc)

            Dim iIndice As Integer = 0

            For Each objInfoAdicDocAutoXml In resInfoAdicDocAutoXml

                If infNFe.autXML Is Nothing Then

                    Dim aAutXml(10) As TNFeInfNFeAutXML
                    infNFe.autXML = aAutXml

                End If

                Dim objAutXml As New TNFeInfNFeAutXML
                infNFe.autXML(iIndice) = objAutXml
                iIndice = iIndice + 1

                Select Case objInfoAdicDocAutoXml.Tipo

                    Case 1
                        objAutXml.ItemElementName = ItemChoiceType5.CPF
                        objAutXml.Item = Fomata_ZerosEsquerda(objInfoAdicDocAutoXml.Doc, 11)

                    Case 2
                        objAutXml.ItemElementName = ItemChoiceType5.CNPJ
                        objAutXml.Item = Fomata_ZerosEsquerda(objInfoAdicDocAutoXml.Doc, 14)

                End Select

            Next

            resNFePadraoAutXml = gobjApp.dbDadosNfe.ExecuteQuery(Of NFePadraoAutXml) _
            ("SELECT * FROM NFePadraoAutXml WHERE FilialEmpresa = {0}", objNFiscal.FilialEmpresa)

            For Each objNFePadraoAutXml In resNFePadraoAutXml

                If infNFe.autXML Is Nothing Then

                    Dim aAutXml(10) As TNFeInfNFeAutXML
                    infNFe.autXML = aAutXml

                End If

                Dim objAutXml As New TNFeInfNFeAutXML
                infNFe.autXML(iIndice) = objAutXml
                iIndice = iIndice + 1

                Select Case Len(objNFePadraoAutXml.CGC)

                    Case Is <= 11
                        objAutXml.ItemElementName = ItemChoiceType5.CPF
                        objAutXml.Item = Fomata_ZerosEsquerda(objNFePadraoAutXml.CGC, 11)

                    Case Else
                        objAutXml.ItemElementName = ItemChoiceType5.CNPJ
                        objAutXml.Item = Fomata_ZerosEsquerda(objNFePadraoAutXml.CGC, 14)

                End Select

            Next

            'colocar a sefaz da BA como autorizada
            If iIndice = 0 And gobjApp.objEstado.Sigla = "BA" Then

                Dim aAutXml(10) As TNFeInfNFeAutXML
                infNFe.autXML = aAutXml

                Dim objAutXml As New TNFeInfNFeAutXML
                infNFe.autXML(0) = objAutXml

                objAutXml.ItemElementName = ItemChoiceType5.CNPJ
                objAutXml.Item = "13937073000156"

            End If

            Monta_NFiscal_Xml25 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml25 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml23(ByVal a5 As TNFe) As Long
        'preenche informacoes adicionais

        Dim sMsg As String
        Dim resMensagensRegra As IEnumerable(Of MensagensRegra)
        Dim resInfoAdicExp As IEnumerable(Of InfoAdicExportacao)
        Dim resInfoAdicCompra As IEnumerable(Of InfoAdicCompra)
        Dim objInfoAdicCompra As InfoAdicCompra
        Dim objInfoAdicExp As InfoAdicExportacao

        Try

            sMsg = ""

            If gobjApp.sVersaoMsg <> "2" Then

                Throw New System.Exception("gobjApp.sVersaoMsg")

            Else

                resMensagensRegra = gobjApp.dbDadosNfe.ExecuteQuery(Of MensagensRegra) _
                    ("SELECT * FROM MensagensRegra WHERE TipoDoc = 0 And NumIntDoc = {0}", lNumIntNF)


                For Each objMensagensRegra In resMensagensRegra
                    sMsg = sMsg & objMensagensRegra.Mensagem
                Next

                Replace(sMsg, "|", " ")

            End If

            If Len(sMsg) > 0 Then
                a5.infNFe.infAdic = New TNFeInfNFeInfAdic
                a5.infNFe.infAdic.infCpl = DesacentuaTexto(sMsg)
            End If

            resInfoAdicExp = gobjApp.dbDadosNfe.ExecuteQuery(Of InfoAdicExportacao) _
                ("SELECT * FROM InfoAdicExportacao WHERE TipoDoc = {0} And NumIntDoc = {1}", 0, lNumIntNF)

            For Each objInfoAdicExp In resInfoAdicExp
                a5.infNFe.exporta = New TNFeInfNFeExporta
                a5.infNFe.exporta.xLocExporta = objInfoAdicExp.LocalEmbarque
                'a5.infNFe.exporta.xLocDespacho = ???? faltando
                a5.infNFe.exporta.UFSaidaPais = GetCode(Of TUf)(objInfoAdicExp.UFEmbarque)

                Exit For

            Next

            If sModelo <> "NFCe" Then

                resInfoAdicCompra = gobjApp.dbDadosNfe.ExecuteQuery(Of InfoAdicCompra) _
                ("SELECT * FROM InfoAdicCompra WHERE TipoDoc = {0} And NumIntDoc = {1}", 0, lNumIntNF)

                For Each objInfoAdicCompra In resInfoAdicCompra
                    a5.infNFe.compra = New TNFeInfNFeCompra
                    If Len(Trim(objInfoAdicCompra.Contrato)) > 0 Then
                        a5.infNFe.compra.xCont = Trim(objInfoAdicCompra.Contrato)
                    End If

                    If Len(Trim(objInfoAdicCompra.NotaEmpenho)) > 0 Then
                        a5.infNFe.compra.xNEmp = Trim(objInfoAdicCompra.NotaEmpenho)
                    End If

                    If Len(Trim(objInfoAdicCompra.Pedido)) > 0 Then
                        a5.infNFe.compra.xPed = Right(Trim(DesacentuaTexto(objInfoAdicCompra.Pedido)), 60)

                    End If

                    Exit For

                Next

                If a5.infNFe.compra Is Nothing Then
                    a5.infNFe.compra = New TNFeInfNFeCompra
                    If Len(Trim(objNFiscal.NumPedidoTerc)) > 0 Then
                        a5.infNFe.compra.xPed = Right(Trim(DesacentuaTexto(objNFiscal.NumPedidoTerc)), 60)
                    End If

                Else
                    If Len(Trim(a5.infNFe.compra.xPed)) = 0 Then
                        If Len(Trim(objNFiscal.NumPedidoTerc)) > 0 Then
                            a5.infNFe.compra.xPed = Right(Trim(DesacentuaTexto(objNFiscal.NumPedidoTerc)), 60)
                        End If
                    End If
                End If

            End If

            Monta_NFiscal_Xml23 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml23 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml22(ByVal infNFeCobr As TNFeInfNFeCobr, ByVal a5 As TNFe) As Long
        'preenche dados de cobrança/pagamento

        Dim resTitPag As IEnumerable(Of TitulosPagTodo)
        Dim resParcPag As IEnumerable(Of ParcelasPagToda)
        Dim resTitRec As IEnumerable(Of TitulosRecTodo)
        Dim resParcRec As IEnumerable(Of ParcelasRecToda)
        Dim resCondPagto As IEnumerable(Of CondicoesPagto)
        Dim iIndice As Integer
        Dim objCondPagto As CondicoesPagto

        Try

            Dim dValorTotalTitulo As Double
            Dim dValorDesconto As Double

            '865 Rejeição: Total dos pagamentos menor que o total da nota
            'Somatório do valor dos pagamentos (id:YA03, tag: vPag) menor que o total da nota (id:W16, tag: vNF) 
            'Exceção 1: Esta regra não se aplica para nota fiscal de Ajuste, campo finNFe=3 (id: B25) e para nota fiscal de Devolução finNFe=4 (id:B25) 
            'Exceção 2: Esta regra não se aplica quando o campo Meio de Pagamento (id: YA02, tag: tPag) for igual a 90 (sem pagamento)

            '902 Rejeição: Valor Liquido da Fatura difere do Valor Original menos o Valor do Desconto

            '851 Rejeição: Soma do valor das parcelas difere do Valor Líquido da Fatura
            'Se informado o grupo de Parcelas de cobrança (tag: dup, Id: Y07) e a soma do valor das parcelas (vDup, id: Y10) 
            'difere do Valor Líquido da Fatura (vLiq, id: Y06). Obs.: Implementação futura em ambiente de produção a partir de 02-jul-2018.

            If gobjApp.iDebug = 1 Then MsgBox("33")
            gobjApp.sErro = "33"
            gobjApp.sMsg1 = "vai tratar os dados de cobrança"

            'se fir um titulo a pagar
            If objNFiscal.ClasseDocCPR = 1 Then

                resTitPag = gobjApp.dbDadosNfe.ExecuteQuery(Of TitulosPagTodo) _
                ("SELECT ValorTotal, CondicaoPagto FROM TitulosPagTodos WHERE NumIntDoc = {0}", objNFiscal.NumIntDocCPR)

                dValorDesconto = 0
                For Each objTitPag In resTitPag

                    a5.infNFe.cobr = infNFeCobr
                    Dim infNFeCobrFat As TNFeInfNFeCobrFat = New TNFeInfNFeCobrFat
                    infNFeCobr.fat = infNFeCobrFat

                    '***********  cobranca ****************************
                    Dim apag(0) As TNFeInfNFePagDetPag
                    a5.infNFe.pag = New TNFeInfNFePag
                    a5.infNFe.pag.detPag = apag
                    apag(0) = New TNFeInfNFePagDetPag
                    apag(0).tPag = TNFeInfNFePagDetPagTPag.Item90 'Sem pagamento (alterado depois se tiver o pagamento)
                    apag(0).vPag = "0.00"

                    infNFeCobrFat.nFat = objNFiscal.NumNotaFiscal

                    dValorDesconto = CDbl(Format(objNFiscal.ValorDescontoTit, "fixed"))
                    'dValorTotalTitulo = CDbl(Format(objTitPag.ValorTotal, "fixed"))

                    'infNFeCobrFat.vOrig = Replace(Format(dValorTotalTitulo + dValorDesconto, "fixed"), ",", ".")
                    'If dValorDesconto <> 0 Then
                    'infNFeCobrFat.vDesc = Replace(Format(dValorDesconto, "fixed"), ",", ".")
                    'End If
                    'infNFeCobrFat.vLiq = Replace(Format(dValorTotalTitulo, "fixed"), ",", ".")

                    resCondPagto = gobjApp.dbDadosNfe.ExecuteQuery(Of CondicoesPagto) _
                    ("SELECT * FROM CondicoesPagto WHERE Codigo = {0}", objTitPag.CondicaoPagto)

                    'CORPORATOR
                    '1 Dinheiro
                    '2 Boleto
                    '3 Cheque
                    '4 Depósito em Conta
                    '5 Cartão de Crédito
                    '6 Cartão de Débito
                    '7 Troca
                    '8 Bonificação
                    'WEB SERVICE
                    '01=Dinheiro    
                    '02=Cheque 
                    '03=Cartão de Crédito
                    '04=Cartão de Débito 
                    '05=Crédito Loja 
                    '10=Vale Alimentação 
                    '11=Vale Refeição '
                    '12=Vale Presente 
                    '13=Vale Combustível 
                    '14=Duplicata Mercantil 
                    '15=Boleto Bancário 
                    '90= Sem pagamento 
                    '99=Outros
                    For Each objCondPagto In resCondPagto
                        Select Case objCondPagto.FormaPagamento
                            Case 0, 1
                                apag(0).tPag = TNFeInfNFePagDetPagTPag.Item01
                            Case 2
                                apag(0).tPag = TNFeInfNFePagDetPagTPag.Item15
                            Case 3
                                apag(0).tPag = TNFeInfNFePagDetPagTPag.Item02
                            Case 5
                                apag(0).tPag = TNFeInfNFePagDetPagTPag.Item03
                            Case 6
                                apag(0).tPag = TNFeInfNFePagDetPagTPag.Item04
                            Case 7, 8, 4
                                apag(0).tPag = TNFeInfNFePagDetPagTPag.Item99
                        End Select

                        Exit For
                    Next
                    'Se não tem condição de pagamento ou outro que não identificou a forma, coloca como Outros
                    If apag(0).tPag = TNFeInfNFePagDetPagTPag.Item90 Then apag(0).tPag = TNFeInfNFePagDetPagTPag.Item99
                    apag(0).vPag = Replace(Format(objNFiscal.ValorTotal, "fixed"), ",", ".") 'Regra 865

                    resParcPag = gobjApp.dbDadosNfe.ExecuteQuery(Of ParcelasPagToda) _
                    ("SELECT * FROM ParcelasPagTodas WHERE NumIntTitulo = {0}", objNFiscal.NumIntDocCPR)

                    iIndice = -1

                    '--REMOVERAM ESSA VALIDAÇÃO
                    'YA02-30 55 Se não informado Duplicata Mercantil como uma das Formas de Pagamento (tag:tPag, id:YA02, = 14) o Grupo Duplicata (id:Y07) não deve ser preenchido
                    'If apag(0).tPag = TNFeInfNFePagDetPagTPag.Item14 Then

                    Dim Dup(50) As TNFeInfNFeCobrDup

                    infNFeCobr.dup = Dup

                    dValorTotalTitulo = 0
                    For Each objParcPag In resParcPag
                        iIndice = iIndice + 1


                        Dim infNFeCobrDup As TNFeInfNFeCobrDup = New TNFeInfNFeCobrDup
                        infNFeCobr.dup(iIndice) = infNFeCobrDup

                        'V1.60
                        'Obrigatória informação do número de parcelas com 3
                        'algarismos, sequenciais e consecutivos.
                        'Ex.:  “001”,”002”,”003”,...
                        'Observação:             este padrão de preenchimento será
                        'obrigatório somente a partir de 03/09/2018
                        'infNFeCobr.dup(iIndice).nDup = objNFiscal.NumNotaFiscal & "/" & objParcPag.NumParcela
                        infNFeCobr.dup(iIndice).nDup = objParcPag.NumParcela.ToString.PadLeft(3, "0")
                        infNFeCobr.dup(iIndice).dVenc = Format(objParcPag.DataVencimento, "yyyy-MM-dd")
                        infNFeCobr.dup(iIndice).vDup = Replace(Format(objParcPag.Valor, "fixed"), ",", ".")

                        dValorTotalTitulo = dValorTotalTitulo + objParcPag.Valor
                    Next
                    'End If

                    infNFeCobrFat.vOrig = Replace(Format(dValorTotalTitulo + dValorDesconto, "fixed"), ",", ".") 'Regra 902
                    infNFeCobrFat.vDesc = Replace(Format(dValorDesconto, "fixed"), ",", ".")
                    infNFeCobrFat.vLiq = Replace(Format(dValorTotalTitulo, "fixed"), ",", ".") 'Regra 851

                    Exit For
                Next


            ElseIf objNFiscal.ClasseDocCPR = 2 Then

                resTitRec = gobjApp.dbDadosNfe.ExecuteQuery(Of TitulosRecTodo) _
                ("SELECT * FROM TitulosRecTodos WHERE NumIntDoc = {0}", objNFiscal.NumIntDocCPR)

                dValorDesconto = 0
                For Each objTitRec In resTitRec

                    a5.infNFe.cobr = infNFeCobr
                    Dim infNFeCobrFat As TNFeInfNFeCobrFat = New TNFeInfNFeCobrFat
                    infNFeCobr.fat = infNFeCobrFat

                    '***********  cobranca ****************************
                    Dim apag(0) As TNFeInfNFePagDetPag
                    a5.infNFe.pag = New TNFeInfNFePag
                    a5.infNFe.pag.detPag = apag
                    apag(0) = New TNFeInfNFePagDetPag
                    apag(0).tPag = TNFeInfNFePagDetPagTPag.Item90 'Sem pagamento (alterado depois se tiver o pagamento)
                    apag(0).vPag = "0.00"

                    'If objTitRec.CondicaoPagto = 1 Then
                    '    a5.infNFe.ide.indPag = TNFeInfNFeIdeIndPag.Item0
                    'Else
                    '    a5.infNFe.ide.indPag = TNFeInfNFeIdeIndPag.Item1
                    'End If

                    dValorDesconto = CDbl(Format(objNFiscal.ValorDescontoTit, "fixed"))
                    'dValorTotalTitulo = CDbl(Format(objTitRec.Valor, "fixed"))

                    infNFeCobrFat.nFat = objNFiscal.NumNotaFiscal
                    'infNFeCobrFat.vOrig = Replace(Format(dValorTotalTitulo + dValorDesconto, "fixed"), ",", ".")
                    'If dValorDesconto <> 0 Then
                    'infNFeCobrFat.vDesc = Replace(Format(dValorDesconto, "fixed"), ",", ".")
                    'End If
                    'infNFeCobrFat.vLiq = Replace(Format(dValorTotalTitulo, "fixed"), ",", ".")

                    resCondPagto = gobjApp.dbDadosNfe.ExecuteQuery(Of CondicoesPagto) _
                                    ("SELECT * FROM CondicoesPagto WHERE Codigo = {0}", objTitRec.CondicaoPagto)

                    'CORPORATOR
                    '1 Dinheiro
                    '2 Boleto
                    '3 Cheque
                    '4 Depósito em Conta
                    '5 Cartão de Crédito
                    '6 Cartão de Débito
                    '7 Troca
                    '8 Bonificação
                    'WEB SERVICE
                    '01=Dinheiro    
                    '02=Cheque 
                    '03=Cartão de Crédito
                    '04=Cartão de Débito 
                    '05=Crédito Loja 
                    '10=Vale Alimentação 
                    '11=Vale Refeição '
                    '12=Vale Presente 
                    '13=Vale Combustível 
                    '14=Duplicata Mercantil 
                    '15=Boleto Bancário 
                    '90= Sem pagamento 
                    '99=Outros
                    For Each objCondPagto In resCondPagto
                        Select Case objCondPagto.FormaPagamento
                            Case 0, 1
                                apag(0).tPag = TNFeInfNFePagDetPagTPag.Item01
                            Case 2
                                apag(0).tPag = TNFeInfNFePagDetPagTPag.Item15
                            Case 3
                                apag(0).tPag = TNFeInfNFePagDetPagTPag.Item02
                            Case 5
                                apag(0).tPag = TNFeInfNFePagDetPagTPag.Item03
                            Case 6
                                apag(0).tPag = TNFeInfNFePagDetPagTPag.Item04
                            Case 7, 8, 4
                                apag(0).tPag = TNFeInfNFePagDetPagTPag.Item99
                        End Select

                        Exit For
                    Next
                    'Se não tem condição de pagamento ou outro que não identificou a forma, coloca como Outros
                    If apag(0).tPag = TNFeInfNFePagDetPagTPag.Item90 Then apag(0).tPag = TNFeInfNFePagDetPagTPag.Item99
                    apag(0).vPag = Replace(Format(objNFiscal.ValorTotal, "fixed"), ",", ".") 'Regra 865

                    resParcRec = gobjApp.dbDadosNfe.ExecuteQuery(Of ParcelasRecToda) _
                    ("SELECT * FROM ParcelasRecTodas WHERE NumIntTitulo = {0}", objNFiscal.NumIntDocCPR)

                    iIndice = -1

                    '--REMOVERAM ESSA VALIDAÇÃO
                    'YA02-30 55 Se não informado Duplicata Mercantil como uma das Formas de Pagamento (tag:tPag, id:YA02, = 14) o Grupo Duplicata (id:Y07) não deve ser preenchido
                    'If apag(0).tPag = TNFeInfNFePagDetPagTPag.Item14 Then

                    Dim Dup(50) As TNFeInfNFeCobrDup

                    infNFeCobr.dup = Dup

                    dValorTotalTitulo = 0
                    For Each objParcRec In resParcRec
                        iIndice = iIndice + 1

                        Dim infNFeCobrDup As TNFeInfNFeCobrDup = New TNFeInfNFeCobrDup
                        infNFeCobr.dup(iIndice) = infNFeCobrDup

                        'infNFeCobr.dup(iIndice).nDup = objNFiscal.NumNotaFiscal & "/" & objParcRec.NumParcela
                        'V1.60
                        'Obrigatória informação do número de parcelas com 3
                        'algarismos, sequenciais e consecutivos.
                        'Ex.:  “001”,”002”,”003”,...
                        'Observação:             este padrão de preenchimento será
                        'obrigatório somente a partir de 03/09/2018
                        infNFeCobr.dup(iIndice).nDup = objParcRec.NumParcela.ToString.PadLeft(3, "0")
                        infNFeCobr.dup(iIndice).dVenc = Format(objParcRec.DataVencimento, "yyyy-MM-dd")
                        infNFeCobr.dup(iIndice).vDup = Replace(Format(objParcRec.Valor, "fixed"), ",", ".")

                        dValorTotalTitulo = dValorTotalTitulo + objParcRec.Valor
                    Next

                    infNFeCobrFat.vOrig = Replace(Format(dValorTotalTitulo + dValorDesconto, "fixed"), ",", ".") 'Regra 902
                    infNFeCobrFat.vDesc = Replace(Format(dValorDesconto, "fixed"), ",", ".")
                    infNFeCobrFat.vLiq = Replace(Format(dValorTotalTitulo, "fixed"), ",", ".") 'Regra 851

                    Exit For
                    'End If
                Next
            End If

            If (a5.infNFe.cobr Is Nothing) Then

                a5.infNFe.cobr = infNFeCobr
                Dim infNFeCobrFat As TNFeInfNFeCobrFat = New TNFeInfNFeCobrFat
                infNFeCobr.fat = infNFeCobrFat
                infNFeCobrFat.nFat = 0
                infNFeCobrFat.vDesc = 0
                infNFeCobrFat.vLiq = 0
                infNFeCobrFat.vOrig = 0

                Dim apag(0) As TNFeInfNFePagDetPag
                a5.infNFe.pag = New TNFeInfNFePag
                a5.infNFe.pag.detPag = apag
                apag(0) = New TNFeInfNFePagDetPag
                apag(0).tPag = TNFeInfNFePagDetPagTPag.Item90 'Sem pagamento (alterado depois se tiver o pagamento)
                apag(0).vPag = "0.00"

            End If

            Monta_NFiscal_Xml22 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml22 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml21(ByVal infNFeTransp As TNFeInfNFeTransp, ByVal a5 As TNFe) As Long
        'preenche dados da transportadora

        Dim resTransp As IEnumerable(Of Transportadora)
        Dim resEndereco As IEnumerable(Of Endereco)
        Dim resPaisTransp As IEnumerable(Of Paise)
        Dim objPaisTransp As New Paise
        Dim sIE As String
        Dim objEnderecoTransp As New Endereco
        Dim resCamposGenericosValores As IEnumerable(Of CamposGenericosValore)

        Try

            If sModelo = "NFCe" And objNFiscal.FreteRespons = 4 Then

                infNFeTransp.modFrete = TNFeInfNFeTranspModFrete.Item9

            Else

                Dim infNFeTranspTransporta As TNFeInfNFeTranspTransporta = New TNFeInfNFeTranspTransporta
                infNFeTransp.transporta = infNFeTranspTransporta

                If gobjApp.iDebug = 1 Then MsgBox("30")
                gobjApp.sErro = "30"
                gobjApp.sMsg1 = "vai tratar os dados da transportadora"
                '***********  transportadora ****************************

                'v2.00
                If objNFiscal.FreteRespons = 0 Or objNFiscal.FreteRespons = 1 Then
                    infNFeTransp.modFrete = TNFeInfNFeTranspModFrete.Item0
                ElseIf objNFiscal.FreteRespons = 2 Then
                    infNFeTransp.modFrete = TNFeInfNFeTranspModFrete.Item1
                ElseIf objNFiscal.FreteRespons = 3 Then
                    infNFeTransp.modFrete = TNFeInfNFeTranspModFrete.Item2
                ElseIf objNFiscal.FreteRespons = 5 Then
                    infNFeTransp.modFrete = TNFeInfNFeTranspModFrete.Item3
                ElseIf objNFiscal.FreteRespons = 6 Then
                    infNFeTransp.modFrete = TNFeInfNFeTranspModFrete.Item4
                ElseIf objNFiscal.FreteRespons = 4 Then
                    infNFeTransp.modFrete = TNFeInfNFeTranspModFrete.Item9
                End If

                resTransp = gobjApp.dbDadosNfe.ExecuteQuery(Of Transportadora) _
                    ("SELECT * FROM Transportadoras WHERE  Codigo = {0}", objNFiscal.CodTransportadora)

                For Each objTransp In resTransp

                    resEndereco = gobjApp.dbDadosNfe.ExecuteQuery(Of Endereco) _
                        ("SELECT * FROM Enderecos WHERE  Codigo = {0}", objTransp.Endereco)

                    For Each objEnderecoTransp In resEndereco
                        Exit For
                    Next


                    resPaisTransp = gobjApp.dbDadosNfe.ExecuteQuery(Of Paise) _
                    ("SELECT * FROM Paises WHERE Codigo = {0}", objEnderecoTransp.CodigoPais)

                    For Each objPaisTransp In resPaisTransp
                        Exit For
                    Next

                    'se for o Brasil
                    If objPaisTransp.CodBacen = 1058 Then

                        If objTransp.Nome = "REMETENTE" And objNFiscal.Tipo = 1 And Len(Trim(objTransp.CGC)) = 0 Then

                            If Len(a5.infNFe.dest.Item) = 14 Then
                                infNFeTransp.transporta.ItemElementName = ItemChoiceType7.CNPJ
                            Else
                                infNFeTransp.transporta.ItemElementName = ItemChoiceType7.CPF
                            End If

                            infNFeTransp.transporta.Item = a5.infNFe.dest.Item


                        Else

                            If Len(objTransp.CGC) = 14 Then
                                infNFeTransp.transporta.ItemElementName = ItemChoiceType7.CNPJ
                            Else
                                infNFeTransp.transporta.ItemElementName = ItemChoiceType7.CPF
                            End If
                            infNFeTransp.transporta.Item = objTransp.CGC

                        End If

                        infNFeTransp.transporta.xNome = DesacentuaTexto(Trim(objTransp.Nome))

                        'v2.00
                        If objTransp.IEIsento <> 1 And Len(Trim(objTransp.InscricaoEstadual)) > 0 Then
                            sIE = ""
                            Call Formata_String_Numero(objTransp.InscricaoEstadual, sIE)
                            infNFeTransp.transporta.IE = sIE
                        End If

                        resEndereco = gobjApp.dbDadosNfe.ExecuteQuery(Of Endereco) _
                            ("SELECT * FROM Enderecos WHERE  Codigo = {0}", objTransp.Endereco)

                        For Each objEnderecoTransp In resEndereco
                            If Len(objEnderecoTransp.Logradouro) > 0 Then
                                infNFeTransp.transporta.xEnder = DesacentuaTexto(Left(IIf(Len(objEnderecoTransp.TipoLogradouro) > 0, objEnderecoTransp.TipoLogradouro & " ", "") & objEnderecoTransp.Logradouro & IIf(objEnderecoTransp.Numero <> 0, " " & objEnderecoTransp.Numero, "") & IIf(Len(objEnderecoTransp.Complemento) > 0, " " & objEnderecoTransp.Complemento, "") & IIf(Len(objEnderecoTransp.Bairro) > 0, " " & objEnderecoTransp.Bairro, ""), 60))
                            Else
                                infNFeTransp.transporta.xEnder = DesacentuaTexto(Left(objEnderecoTransp.Endereco & IIf(Len(objEnderecoTransp.Bairro) > 0, " " & objEnderecoTransp.Bairro, ""), 60))
                            End If

                            infNFeTransp.transporta.xMun = DesacentuaTexto(objEnderecoTransp.Cidade)

                            infNFeTransp.transporta.UF = GetCode(Of TUf)(objEnderecoTransp.SiglaEstado)

                            If Len(objEnderecoTransp.SiglaEstado) > 0 Then
                                infNFeTransp.transporta.UFSpecified = True
                            End If

                            Exit For

                        Next
                        Exit For
                    End If

                Next

                If sModelo <> "NFCe" Then

                    If gobjApp.iDebug = 1 Then MsgBox("31")
                    gobjApp.sErro = "31"
                    gobjApp.sMsg1 = "vai tratar os dados de veiculo"

                    '***********  veiculo ****************************

                    If Len(Trim(objNFiscal.Placa)) > 0 And Len(Trim(objNFiscal.PlacaUF)) > 0 Then

                        Dim veiculo(0 To 0) As TVeiculo

                        Dim ItemsElementName0(0 To 0) As ItemsChoiceType5

                        infNFeTransp.ItemsElementName = ItemsElementName0

                        ItemsElementName0(0) = ItemsChoiceType5.veicTransp

                        infNFeTransp.Items = veiculo
                        veiculo(0) = New TVeiculo
                        veiculo(0).placa = Trim(objNFiscal.Placa)

                        veiculo(0).UF = GetCode(Of TUf)(objNFiscal.PlacaUF)


                    End If

                End If

                If gobjApp.iDebug = 1 Then MsgBox("32")
                gobjApp.sErro = "32"
                gobjApp.sMsg1 = "vai tratar os dados de volume"

                '***********  volume ****************************

                If objNFiscal.VolumeQuant > 0 Or objNFiscal.VolumeEspecie > 0 Or objNFiscal.VolumeMarca > 0 Or Len(objNFiscal.VolumeNumero) > 0 Or objNFiscal.PesoLiq <> 0 Or objNFiscal.PesoBruto <> 0 Then

                    Dim infNFeTranspVol(0 To 10) As TNFeInfNFeTranspVol

                    infNFeTransp.vol = infNFeTranspVol

                    Dim infNFeTranspVol1 As TNFeInfNFeTranspVol = New TNFeInfNFeTranspVol

                    infNFeTransp.vol(0) = infNFeTranspVol1

                    If objNFiscal.VolumeQuant = 0 Then Throw New System.Exception("O campo volume da nota fiscal deve estar preenchido pois há outros elementos do grupo vol como especie, marca, nVol, pesoL ou pesoB que estão preenchidos.")

                    infNFeTransp.vol(0).qVol = objNFiscal.VolumeQuant

                    resCamposGenericosValores = gobjApp.dbDadosNfe.ExecuteQuery(Of CamposGenericosValore) _
                        ("SELECT * FROM CamposGenericosValores WHERE  CodCampo = 1 AND CodValor = {0}", objNFiscal.VolumeEspecie)

                    For Each objCamposGenericosValores In resCamposGenericosValores
                        infNFeTransp.vol(0).esp = DesacentuaTexto(objCamposGenericosValores.Valor)
                        Exit For
                    Next

                    resCamposGenericosValores = gobjApp.dbDadosNfe.ExecuteQuery(Of CamposGenericosValore) _
                            ("SELECT * FROM CamposGenericosValores WHERE  CodCampo = 2 AND CodValor = {0}", objNFiscal.VolumeMarca)

                    For Each objCamposGenericosValores In resCamposGenericosValores
                        infNFeTransp.vol(0).marca = DesacentuaTexto(objCamposGenericosValores.Valor)
                        Exit For
                    Next

                    If Len(Trim(objNFiscal.VolumeNumero)) > 0 Then infNFeTransp.vol(0).nVol = objNFiscal.VolumeNumero
                    If objNFiscal.PesoLiq <> 0 Then infNFeTransp.vol(0).pesoL = Replace(Format(objNFiscal.PesoLiq, "##########0.000"), ",", ".")
                    If objNFiscal.PesoBruto <> 0 Then infNFeTransp.vol(0).pesoB = Replace(Format(objNFiscal.PesoBruto, "##########0.000"), ",", ".")

                End If

            End If

            Monta_NFiscal_Xml21 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml21 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml20(ByVal infNFeTotal As TNFeInfNFeTotal) As Long
        'preenche totais da nf

        Try

            If gobjApp.iDebug = 1 Then MsgBox("28")
            gobjApp.sErro = "28"
            gobjApp.sMsg1 = "vai iniciar o tratamento dos totais da nota"

            '***********  icms total ****************************

            Dim infNFeTotalICMSTot As TNFeInfNFeTotalICMSTot = New TNFeInfNFeTotalICMSTot
            infNFeTotal.ICMSTot = infNFeTotalICMSTot

            'infNFeTotal.ICMSTot.vBC = Replace(Format(objTributacaoDoc.ICMSBase, "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vBC = Replace(Format(dICMSBaseTotal, "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vICMS = Replace(Format((objTributacaoDoc.ICMSValor - objTributacaoDoc.ICMSVlrFCP), "fixed"), ",", ".")
            svICMS = infNFeTotal.ICMSTot.vICMS
            infNFeTotal.ICMSTot.vICMSDeson = Replace(Format(objTributacaoDoc.ICMSValorIsento, "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vBCST = Replace(Format(objTributacaoDoc.ICMSSubstBase, "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vST = Replace(Format((objTributacaoDoc.ICMSSubstValor - objTributacaoDoc.ICMSVlrFCPST), "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vProd = Replace(Format(dvProdICMS + objNFiscal.ValorDesconto, "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vFrete = Replace(Format(objNFiscal.ValorFrete, "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vSeg = Replace(Format(objNFiscal.ValorSeguro, "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vDesc = Replace(Format(dTotalDescontoItem, "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vIPI = Replace(Format(objTributacaoDoc.IPIValor, "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vPIS = Replace(Format(dValorPIS - dValorServPIS, "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vCOFINS = Replace(Format(dValorCOFINS - dValorServCOFINS, "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vOutro = Replace(Format(objNFiscal.ValorOutrasDespesas, "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vNF = Replace(Format(objNFiscal.ValorTotal + objNFiscal.ValorDesconto + objTributacaoDoc.IPIVlrDevolvido, "fixed"), ",", ".")

            infNFeTotal.ICMSTot.vFCP = Replace(Format(objTributacaoDoc.ICMSVlrFCP, "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vFCPST = Replace(Format(objTributacaoDoc.ICMSVlrFCPST, "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vFCPSTRet = Replace(Format(objTributacaoDoc.ICMSVlrFCPSTRet, "fixed"), ",", ".")
            infNFeTotal.ICMSTot.vIPIDevol = Replace(Format(objTributacaoDoc.IPIVlrDevolvido, "fixed"), ",", ".")


            svNF = infNFeTotal.ICMSTot.vNF

            infNFeTotal.ICMSTot.vII = Replace(Format(IIValorTotal, "fixed"), ",", ".")



            If objNFiscal.DataEmissao >= #6/1/2013# And objTributacaoDoc.TotTrib <> 0 Then
                infNFeTotal.ICMSTot.vTotTrib = Replace(Format(objTributacaoDoc.TotTrib, "fixed"), ",", ".")
            End If

            'EC 87/2015
            If bICMSUFDest And (objTributacaoDoc.ICMSInterestVlrFCPUFDest <> 0 Or objTributacaoDoc.ICMSInterestVlrUFDest <> 0 Or objTributacaoDoc.ICMSInterestVlrUFRemet <> 0) Then
                infNFeTotal.ICMSTot.vFCPUFDest = Replace(Format(objTributacaoDoc.ICMSInterestVlrFCPUFDest, "fixed"), ",", ".")
                infNFeTotal.ICMSTot.vICMSUFDest = Replace(Format(objTributacaoDoc.ICMSInterestVlrUFDest, "fixed"), ",", ".")
                infNFeTotal.ICMSTot.vICMSUFRemet = Replace(Format(objTributacaoDoc.ICMSInterestVlrUFRemet, "fixed"), ",", ".")
            End If

            ' ************ ISSQN total ***********************
            If bComISSQN Then

                Dim infNFeTotalISSQNtot As TNFeInfNFeTotalISSQNtot = New TNFeInfNFeTotalISSQNtot
                infNFeTotal.ISSQNtot = infNFeTotalISSQNtot

                infNFeTotalISSQNtot.dCompet = Format(objTributacaoDoc.DataPrestServico, "yyyy-MM-dd")

                If objTributacaoDoc.ISSBase > 0 Then
                    infNFeTotalISSQNtot.vBC = Replace(Format(objTributacaoDoc.ISSBase, "fixed"), ",", ".")
                End If

                If dServNTribICMS > 0 Then
                    infNFeTotalISSQNtot.vServ = Replace(Format(dServNTribICMS, "fixed"), ",", ".")
                End If

                If objTributacaoDoc.ISSValor > 0 Then
                    infNFeTotalISSQNtot.vISS = Replace(Format(objTributacaoDoc.ISSValor, "fixed"), ",", ".")
                End If

                If dValorServPIS > 0 Then
                    infNFeTotalISSQNtot.vPIS = Replace(Format(dValorServPIS, "fixed"), ",", ".")
                End If

                If dValorServCOFINS > 0 Then
                    infNFeTotalISSQNtot.vCOFINS = Replace(Format(dValorServCOFINS, "fixed"), ",", ".")
                End If

                If gobjApp.iDebug = 1 Then MsgBox("29")
                gobjApp.sErro = "29"
                gobjApp.sMsg1 = "vai iniciar o tratamento das retencoes de impostos da nota"

            End If

            '***********  retencao total ****************************

            Dim infNFeTotalRetTrib As TNFeInfNFeTotalRetTrib = New TNFeInfNFeTotalRetTrib
            infNFeTotal.retTrib = infNFeTotalRetTrib

            If objTributacaoDoc.PISRetido > 0.0 Or objTributacaoDoc.COFINSRetido > 0 Or objTributacaoDoc.CSLLRetido > 0 Or objTributacaoDoc.IRRFBase > 0 Or objTributacaoDoc.IRRFValor > 0 Or objTributacaoDoc.INSSValorBase > 0 Or objTributacaoDoc.INSSRetido > 0 Then

                If objTributacaoDoc.PISRetido > 0.0 Then
                    infNFeTotal.retTrib.vRetPIS = Replace(Format(objTributacaoDoc.PISRetido, "fixed"), ",", ".")
                End If
                If objTributacaoDoc.COFINSRetido > 0 Then
                    infNFeTotal.retTrib.vRetCOFINS = Replace(Format(objTributacaoDoc.COFINSRetido, "fixed"), ",", ".")
                End If
                If objTributacaoDoc.CSLLRetido > 0 Then
                    infNFeTotal.retTrib.vRetCSLL = Replace(Format(objTributacaoDoc.CSLLRetido, "fixed"), ",", ".")
                End If
                If objTributacaoDoc.IRRFBase > 0 Then
                    infNFeTotal.retTrib.vBCIRRF = Replace(Format(objTributacaoDoc.IRRFBase, "fixed"), ",", ".")
                End If
                If objTributacaoDoc.IRRFValor > 0 Then
                    infNFeTotal.retTrib.vIRRF = Replace(Format(objTributacaoDoc.IRRFValor, "fixed"), ",", ".")
                End If
                If objTributacaoDoc.INSSValorBase > 0 Then
                    infNFeTotal.retTrib.vBCRetPrev = Replace(Format(objTributacaoDoc.INSSValorBase, "fixed"), ",", ".")
                End If
                If objTributacaoDoc.INSSRetido > 0 Then
                    infNFeTotal.retTrib.vRetPrev = Replace(Format(objTributacaoDoc.INSSRetido, "fixed"), ",", ".")
                End If

            End If

            Monta_NFiscal_Xml20 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml20 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml19(ByVal objNFeInfNFeDet As TNFeInfNFeDet, ByVal objItemNF As ItensNFiscal, ByVal objTribDocItem As TributacaoDocItem) As Long
        'preenche a parte de ISS do item da nf

        Dim resISSQN As IEnumerable(Of ISSQN)
        Dim objISSQN As ISSQN

        Try

            ''***************** ISS **************************************
            If Len(Trim(objTribDocItem.ISSQN)) <> 0 And objTribDocItem.ICMSTipo = 0 Then

                bComISSQN = True

                dServNTribICMS = dServNTribICMS + (objItemNF.PrecoUnitario * IIf(objItemNF.Quantidade = 0, 1, objItemNF.Quantidade))

                dValorServPIS = dValorServPIS + objTribDocItem.PISValor
                dValorServCOFINS = dValorServCOFINS + objTribDocItem.COFINSValor

                Dim infNFeDetImpostoISSQN As TNFeInfNFeDetImpostoISSQN = New TNFeInfNFeDetImpostoISSQN
                objNFeInfNFeDet.imposto.Items(0) = infNFeDetImpostoISSQN

                infNFeDetImpostoISSQN.vBC = Replace(Format(objTribDocItem.ISSBase, "fixed"), ",", ".")
                infNFeDetImpostoISSQN.vAliq = Replace(Format(objTribDocItem.ISSAliquota * 100, "##0.00"), ",", ".")
                infNFeDetImpostoISSQN.vISSQN = Replace(Format(objTribDocItem.ISSValor, "fixed"), ",", ".")
                infNFeDetImpostoISSQN.cMunFG = cMunFG '????? objTribDocItem.

                resISSQN = gobjApp.dbDadosNfe.ExecuteQuery(Of ISSQN) _
                ("SELECT * FROM ISSQN WHERE  Codigo = {0} ", objTribDocItem.ISSQN)

                objISSQN = resISSQN(0)

                If objISSQN Is Nothing Then
                    Throw New System.Exception("O campo ISSQN deste produto não estava preenchido no momento da gravação da nota. Produto = " & objItemNF.Produto)
                End If

                ''classificacao do servico conforme tabela da lei complementar 116 de 2003 (LC 116/03)
                infNFeDetImpostoISSQN.cListServ = objISSQN.CListServNFe

            End If

            Monta_NFiscal_Xml19 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml19 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml18(ByVal objNFeInfNFeDet As TNFeInfNFeDet, ByVal objItemNF As ItensNFiscal, ByVal objTribDocItem As TributacaoDocItem) As Long
        'preenche a parte de cofins st do item da nf

        Try

            If gobjApp.iDebug = 1 Then MsgBox("27")
            gobjApp.sErro = "27"
            gobjApp.sMsg1 = "vai iniciar a tributacao de COFINS ST"

            '***********  COFINS ST****************************

            If objTribDocItem.COFINSSTValor <> 0 Then
                Dim COFINSST As TNFeInfNFeDetImpostoCOFINSST = New TNFeInfNFeDetImpostoCOFINSST
                objNFeInfNFeDet.imposto.COFINSST = COFINSST

                Dim ItemsElementName4(1) As ItemsChoiceType4
                Dim ItemsString4(1) As String

                COFINSST.ItemsElementName = ItemsElementName4
                COFINSST.Items = ItemsString4


                If objTribDocItem.COFINSSTTipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL Then

                    COFINSST.ItemsElementName(0) = ItemsChoiceType4.vBC
                    COFINSST.Items(0) = Replace(Format(objTribDocItem.COFINSSTBase, "fixed"), ",", ".")
                    COFINSST.ItemsElementName(1) = ItemsChoiceType4.pCOFINS
                    COFINSST.Items(1) = Replace(Format(objTribDocItem.COFINSSTAliquota * 100, "##0.00"), ",", ".")

                Else


                    COFINSST.ItemsElementName(0) = ItemsChoiceType4.qBCProd
                    COFINSST.Items(0) = Replace(Format(objTribDocItem.COFINSSTQtde, "#########0.0000"), ",", ".")
                    COFINSST.ItemsElementName(1) = ItemsChoiceType4.vAliqProd
                    COFINSST.Items(1) = Replace(Format(objTribDocItem.COFINSSTAliquotaValor, "#########0.0000"), ",", ".")

                End If

                COFINSST.vCOFINS = Replace(Format(objTribDocItem.COFINSSTValor, "fixed"), ",", ".")

            End If

            Monta_NFiscal_Xml18 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml18 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml17(ByVal objNFeInfNFeDet As TNFeInfNFeDet, ByVal objItemNF As ItensNFiscal, ByVal objTribDocItem As TributacaoDocItem) As Long
        'preenche a parte de cofins do item da nf

        Try

            '***********  COFINS ****************************

            If gobjApp.iDebug = 1 Then MsgBox("26")
            gobjApp.sErro = "26"
            gobjApp.sMsg1 = "vai iniciar a tributacao de COFINS"


            Dim infNFeDetImpostoCOFINS As TNFeInfNFeDetImpostoCOFINS = New TNFeInfNFeDetImpostoCOFINS
            objNFeInfNFeDet.imposto.COFINS = infNFeDetImpostoCOFINS

            dValorCOFINS = dValorCOFINS + objTribDocItem.COFINSValor

            Select Case objTribDocItem.COFINSTipo

                Case 1, 2
                    Dim COFINSAliq As New TNFeInfNFeDetImpostoCOFINSCOFINSAliq

                    objNFeInfNFeDet.imposto.COFINS.Item = COFINSAliq


                    If objTribDocItem.COFINSTipo = 1 Then
                        COFINSAliq.CST = TNFeInfNFeDetImpostoCOFINSCOFINSAliqCST.Item01
                    Else
                        COFINSAliq.CST = TNFeInfNFeDetImpostoCOFINSCOFINSAliqCST.Item02
                    End If


                    COFINSAliq.pCOFINS = Replace(Format(objTribDocItem.COFINSAliquota * 100, "##0.00"), ",", ".")
                    COFINSAliq.vCOFINS = Replace(Format(objTribDocItem.COFINSValor, "fixed"), ",", ".")
                    COFINSAliq.vBC = Replace(Format(objTribDocItem.COFINSBase, "fixed"), ",", ".")

                Case 3
                    Dim COFINSQtde As New TNFeInfNFeDetImpostoCOFINSCOFINSQtde
                    objNFeInfNFeDet.imposto.COFINS.Item = COFINSQtde

                    COFINSQtde.CST = TNFeInfNFeDetImpostoCOFINSCOFINSQtdeCST.Item03
                    COFINSQtde.qBCProd = Replace(Format(objTribDocItem.COFINSQtde, "#########0.0000"), ",", ".")
                    COFINSQtde.vAliqProd = Replace(Format(objTribDocItem.COFINSAliquotaValor, "#########0.0000"), ",", ".")
                    COFINSQtde.vCOFINS = Replace(Format(objTribDocItem.COFINSValor, "fixed"), ",", ".")

                Case 4
                    Dim COFINSNT As New TNFeInfNFeDetImpostoCOFINSCOFINSNT
                    objNFeInfNFeDet.imposto.COFINS.Item = COFINSNT

                    COFINSNT.CST = TNFeInfNFeDetImpostoCOFINSCOFINSNTCST.Item04

                Case 6
                    Dim COFINSNT As New TNFeInfNFeDetImpostoCOFINSCOFINSNT
                    objNFeInfNFeDet.imposto.COFINS.Item = COFINSNT

                    COFINSNT.CST = TNFeInfNFeDetImpostoCOFINSCOFINSNTCST.Item06

                Case 7
                    Dim COFINSNT As New TNFeInfNFeDetImpostoCOFINSCOFINSNT
                    objNFeInfNFeDet.imposto.COFINS.Item = COFINSNT

                    COFINSNT.CST = TNFeInfNFeDetImpostoCOFINSCOFINSNTCST.Item07

                Case 8
                    Dim COFINSNT As New TNFeInfNFeDetImpostoCOFINSCOFINSNT
                    objNFeInfNFeDet.imposto.COFINS.Item = COFINSNT

                    COFINSNT.CST = TNFeInfNFeDetImpostoCOFINSCOFINSNTCST.Item08

                Case 9
                    Dim COFINSNT As New TNFeInfNFeDetImpostoCOFINSCOFINSNT
                    objNFeInfNFeDet.imposto.COFINS.Item = COFINSNT

                    COFINSNT.CST = TNFeInfNFeDetImpostoCOFINSCOFINSNTCST.Item09

                Case 49 To 56, 60 To 67, 70 To 75, 98, 99
                    Dim COFINSOutr As New TNFeInfNFeDetImpostoCOFINSCOFINSOutr
                    objNFeInfNFeDet.imposto.COFINS.Item = COFINSOutr

                    Dim ItemsElementName3(1) As ItemsChoiceType3
                    Dim ItemsString3(1) As String

                    COFINSOutr.ItemsElementName = ItemsElementName3
                    COFINSOutr.Items = ItemsString3

                    If objTribDocItem.COFINSTipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL Then

                        COFINSOutr.ItemsElementName(0) = ItemsChoiceType3.vBC
                        COFINSOutr.Items(0) = Replace(Format(objTribDocItem.COFINSBase, "fixed"), ",", ".")
                        COFINSOutr.ItemsElementName(1) = ItemsChoiceType3.pCOFINS
                        COFINSOutr.Items(1) = Replace(Format(objTribDocItem.COFINSAliquota * 100, "##0.00"), ",", ".")

                    Else

                        COFINSOutr.ItemsElementName(0) = ItemsChoiceType3.qBCProd
                        COFINSOutr.Items(0) = Replace(Format(objTribDocItem.COFINSQtde, "#########0.0000"), ",", ".")
                        COFINSOutr.ItemsElementName(1) = ItemsChoiceType3.vAliqProd
                        COFINSOutr.Items(1) = Replace(Format(objTribDocItem.COFINSAliquotaValor, "#########0.0000"), ",", ".")

                    End If

                    COFINSOutr.vCOFINS = Replace(Format(objTribDocItem.COFINSValor, "fixed"), ",", ".")
                    COFINSOutr.CST = GetCode(Of TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST)(CStr(objTribDocItem.COFINSTipo))

            End Select

            Monta_NFiscal_Xml17 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml17 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml16(ByVal objNFeInfNFeDet As TNFeInfNFeDet, ByVal objItemNF As ItensNFiscal, ByVal objTribDocItem As TributacaoDocItem) As Long
        'preenche a parte de pis ST do item da nf

        Try

            If gobjApp.iDebug = 1 Then MsgBox("25")
            gobjApp.sErro = "25"
            gobjApp.sMsg1 = "vai iniciar a tributacao de PIS ST"


            '***********  PIS ST ****************************

            If objTribDocItem.PISSTValor <> 0 Then

                Dim PISST As TNFeInfNFeDetImpostoPISST = New TNFeInfNFeDetImpostoPISST
                objNFeInfNFeDet.imposto.PISST = PISST

                Dim ItemsElementName2(1) As ItemsChoiceType2
                Dim ItemsString2(1) As String

                PISST.ItemsElementName = ItemsElementName2
                PISST.Items = ItemsString2

                If objTribDocItem.PISSTTipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL Then

                    PISST.ItemsElementName(0) = ItemsChoiceType2.vBC
                    PISST.Items(0) = Replace(Format(objTribDocItem.PISSTBase, "fixed"), ",", ".")
                    PISST.ItemsElementName(1) = ItemsChoiceType2.pPIS
                    PISST.Items(1) = Replace(Format(objTribDocItem.PISSTAliquota * 100, "##0.00"), ",", ".")

                Else

                    PISST.ItemsElementName(0) = ItemsChoiceType1.qBCProd
                    PISST.Items(0) = Replace(Format(objTribDocItem.PISSTQtde, "#########0.0000"), ",", ".")
                    PISST.ItemsElementName(1) = ItemsChoiceType1.vAliqProd
                    PISST.Items(1) = Replace(Format(objTribDocItem.PISSTAliquotaValor, "#########0.0000"), ",", ".")

                End If

                PISST.vPIS = Replace(Format(objTribDocItem.PISSTValor, "fixed"), ",", ".")

            End If

            Monta_NFiscal_Xml16 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml16 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml15(ByVal objNFeInfNFeDet As TNFeInfNFeDet, ByVal objItemNF As ItensNFiscal, ByVal objTribDocItem As TributacaoDocItem) As Long
        'preenche a parte de pis do item da nf

        Try

            '***********  PIS ****************************

            If gobjApp.iDebug = 1 Then MsgBox("24")
            gobjApp.sErro = "24"
            gobjApp.sMsg1 = "vai iniciar a tributacao de PIS"


            Dim infNFeDetImpostoPIS As TNFeInfNFeDetImpostoPIS = New TNFeInfNFeDetImpostoPIS
            objNFeInfNFeDet.imposto.PIS = infNFeDetImpostoPIS

            dValorPIS = dValorPIS + objTribDocItem.PISValor

            Select Case objTribDocItem.PISTipo

                Case 1, 2
                    Dim PISAliq As New TNFeInfNFeDetImpostoPISPISAliq


                    objNFeInfNFeDet.imposto.PIS.Item = PISAliq

                    If objTribDocItem.PISTipo = 1 Then
                        PISAliq.CST = TNFeInfNFeDetImpostoPISPISAliqCST.Item01
                    Else
                        PISAliq.CST = TNFeInfNFeDetImpostoPISPISAliqCST.Item02
                    End If

                    PISAliq.pPIS = Replace(Format(objTribDocItem.PISAliquota * 100, "##0.00"), ",", ".")
                    PISAliq.vPIS = Replace(Format(objTribDocItem.PISValor, "fixed"), ",", ".")
                    PISAliq.vBC = Replace(Format(objTribDocItem.PISBase, "fixed"), ",", ".")

                Case 3
                    Dim PISQtde As New TNFeInfNFeDetImpostoPISPISQtde
                    objNFeInfNFeDet.imposto.PIS.Item = PISQtde

                    PISQtde.CST = TNFeInfNFeDetImpostoPISPISQtdeCST.Item03
                    PISQtde.qBCProd = Replace(Format(objTribDocItem.PISQtde, "#########0.0000"), ",", ".")
                    PISQtde.vAliqProd = Replace(Format(objTribDocItem.PISAliquotaValor, "#########0.0000"), ",", ".")
                    PISQtde.vPIS = Replace(Format(objTribDocItem.PISValor, "fixed"), ",", ".")

                Case 4
                    Dim PISNT As New TNFeInfNFeDetImpostoPISPISNT
                    objNFeInfNFeDet.imposto.PIS.Item = PISNT

                    PISNT.CST = TNFeInfNFeDetImpostoPISPISNTCST.Item04

                Case 6
                    Dim PISNT As New TNFeInfNFeDetImpostoPISPISNT
                    objNFeInfNFeDet.imposto.PIS.Item = PISNT

                    PISNT.CST = TNFeInfNFeDetImpostoPISPISNTCST.Item06

                Case 7
                    Dim PISNT As New TNFeInfNFeDetImpostoPISPISNT
                    objNFeInfNFeDet.imposto.PIS.Item = PISNT

                    PISNT.CST = TNFeInfNFeDetImpostoPISPISNTCST.Item07

                Case 8
                    Dim PISNT As New TNFeInfNFeDetImpostoPISPISNT
                    objNFeInfNFeDet.imposto.PIS.Item = PISNT

                    PISNT.CST = TNFeInfNFeDetImpostoPISPISNTCST.Item08

                Case 9
                    Dim PISNT As New TNFeInfNFeDetImpostoPISPISNT
                    objNFeInfNFeDet.imposto.PIS.Item = PISNT

                    PISNT.CST = TNFeInfNFeDetImpostoPISPISNTCST.Item09

                Case 49 To 56, 60 To 67, 70 To 75, 98, 99
                    Dim PISOutr As New TNFeInfNFeDetImpostoPISPISOutr
                    objNFeInfNFeDet.imposto.PIS.Item = PISOutr

                    Dim ItemsElementName1(1) As ItemsChoiceType1
                    Dim ItemsString1(1) As String

                    PISOutr.ItemsElementName = ItemsElementName1
                    PISOutr.Items = ItemsString1

                    If objTribDocItem.PISTipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL Then

                        PISOutr.ItemsElementName(0) = ItemsChoiceType1.vBC
                        PISOutr.Items(0) = Replace(Format(objTribDocItem.PISBase, "fixed"), ",", ".")
                        PISOutr.ItemsElementName(1) = ItemsChoiceType1.pPIS
                        PISOutr.Items(1) = Replace(Format(objTribDocItem.PISAliquota * 100, "##0.00"), ",", ".")

                    Else

                        PISOutr.ItemsElementName(0) = ItemsChoiceType1.qBCProd
                        PISOutr.Items(0) = Replace(Format(objTribDocItem.PISQtde, "#########0.0000"), ",", ".")
                        PISOutr.ItemsElementName(1) = ItemsChoiceType1.vAliqProd
                        PISOutr.Items(1) = Replace(Format(objTribDocItem.PISAliquotaValor, "#########0.0000"), ",", ".")
                    End If

                    PISOutr.vPIS = Replace(Format(objTribDocItem.PISValor, "fixed"), ",", ".")

                    PISOutr.CST = GetCode(Of TNFeInfNFeDetImpostoPISPISOutrCST)(CStr(objTribDocItem.PISTipo))

            End Select

            Monta_NFiscal_Xml15 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml15 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml14(ByVal objNFeInfNFeDet As TNFeInfNFeDet, ByVal objItemNF As ItensNFiscal, ByVal objTribDocItem As TributacaoDocItem) As Long
        'preenche a parte de ipi do item da nf

        Dim resTipoTribIPI As IEnumerable(Of TiposTribIPI)
        Dim objTipoTribIPI As New TiposTribIPI
        Dim iCSTIPI As Integer
        Dim IPITrib As TIpiIPITrib = New TIpiIPITrib

        Try

            If gobjApp.iDebug = 1 Then MsgBox("21")
            gobjApp.sErro = "21"
            gobjApp.sMsg1 = "vai iniciar a tributacao de IPI"

            '********************  IPI *******************************************

            If gobjApp.objFilialEmpresa.ContribuinteIPI = 0 And gobjApp.objFilialEmpresa.CGC = "04970473000172" Then
                'nao gerar grupo de IPI para a FANART
            Else

                Dim infNFeDetImpostoIPI As TIpi = New TIpi
                objNFeInfNFeDet.imposto.Items(1) = infNFeDetImpostoIPI

                'If Len(objTribDocItem.IPIEnquadramentoClasse) > 0 Then
                '    infNFeDetImpostoIPI.clEnq = objTribDocItem.IPIEnquadramentoClasse
                'End If

                If Len(objTribDocItem.IPIEnquadramentoCodigo) > 0 Then
                    infNFeDetImpostoIPI.cEnq = objTribDocItem.IPIEnquadramentoCodigo
                Else
                    infNFeDetImpostoIPI.cEnq = "999"
                End If

                If Len(objTribDocItem.IPISeloCodigo) > 0 Then
                    infNFeDetImpostoIPI.cSelo = objTribDocItem.IPISeloCodigo
                End If

                If Len(objTribDocItem.IPICNPJProdutor) > 0 Then
                    infNFeDetImpostoIPI.CNPJProd = objTribDocItem.IPICNPJProdutor
                End If

                If objTribDocItem.IPISeloQtde > 0 Then
                    infNFeDetImpostoIPI.qSelo = objTribDocItem.IPISeloQtde
                End If

                If gobjApp.iDebug = 1 Then MsgBox("22")
                gobjApp.sErro = "22"
                gobjApp.sMsg1 = "vai continuar a tributacao de IPI"

                resTipoTribIPI = gobjApp.dbDadosNfe.ExecuteQuery(Of TiposTribIPI) _
                ("SELECT * FROM TiposTribIPI WHERE  Tipo = {0}", objTribDocItem.IPITipo)

                For Each objTipoTribIPI In resTipoTribIPI
                    Exit For
                Next

                'se for uma nota de entrada pega os CST de Entrada, senao pega os CSTs de Saida
                If objNFiscal.Tipo = 1 Then
                    iCSTIPI = objTipoTribIPI.CSTEntrada
                Else
                    iCSTIPI = objTipoTribIPI.CSTSaida
                End If

                If iCSTIPI = 0 Or iCSTIPI = 49 Or iCSTIPI = 50 Or iCSTIPI = 99 Then
                    IPITrib = New TIpiIPITrib
                    infNFeDetImpostoIPI.Item = IPITrib


                    Select Case iCSTIPI
                        Case 0
                            IPITrib.CST = TIpiIPITribCST.Item00
                        Case 49
                            IPITrib.CST = TIpiIPITribCST.Item49
                        Case 50
                            IPITrib.CST = TIpiIPITribCST.Item50
                        Case 99
                            IPITrib.CST = TIpiIPITribCST.Item99
                    End Select

                    Dim ItemsElementName0(2) As ItemsChoiceType
                    Dim ItemsString0(2) As String

                    IPITrib.ItemsElementName = ItemsElementName0
                    IPITrib.Items = ItemsString0

                    If objTribDocItem.IPITipoCalculo = TRIB_TIPO_CALCULO_PERCENTUAL Then

                        IPITrib.ItemsElementName(0) = ItemsChoiceType.vBC
                        IPITrib.Items(0) = Replace(Format(objTribDocItem.IPIBaseCalculo * (1 - objTribDocItem.IPIPercRedBase), "fixed"), ",", ".")
                        IPITrib.ItemsElementName(1) = ItemsChoiceType.pIPI
                        IPITrib.Items(1) = Replace(Format(objTribDocItem.IPIAliquota * 100, "##0.00"), ",", ".")
                    Else

                        IPITrib.ItemsElementName(0) = ItemsChoiceType.qUnid
                        IPITrib.Items(0) = Replace(Format(objTribDocItem.IPIUnidadePadraoQtde, "#########0.0000"), ",", ".")
                        IPITrib.ItemsElementName(1) = ItemsChoiceType.vUnid
                        IPITrib.Items(1) = Replace(Format(objTribDocItem.IPIUnidadePadraoValor, "#########0.0000"), ",", ".")
                    End If

                    IPITrib.vIPI = Replace(Format(objTribDocItem.IPIValor, "fixed"), ",", ".")

                ElseIf iCSTIPI = 1 Then
                    Dim IPINT As TIpiIPINT = New TIpiIPINT
                    infNFeDetImpostoIPI.Item = IPINT
                    IPINT.CST = TIpiIPINTCST.Item01

                ElseIf iCSTIPI = 2 Then
                    Dim IPINT As TIpiIPINT = New TIpiIPINT
                    infNFeDetImpostoIPI.Item = IPINT
                    IPINT.CST = TIpiIPINTCST.Item02

                ElseIf iCSTIPI = 3 Then
                    Dim IPINT As TIpiIPINT = New TIpiIPINT
                    infNFeDetImpostoIPI.Item = IPINT
                    IPINT.CST = TIpiIPINTCST.Item03

                ElseIf iCSTIPI = 4 Then
                    Dim IPINT As TIpiIPINT = New TIpiIPINT
                    infNFeDetImpostoIPI.Item = IPINT
                    IPINT.CST = TIpiIPINTCST.Item04

                ElseIf iCSTIPI = 5 Then
                    Dim IPINT As TIpiIPINT = New TIpiIPINT
                    infNFeDetImpostoIPI.Item = IPINT
                    IPINT.CST = TIpiIPINTCST.Item04

                ElseIf iCSTIPI = 51 Then
                    Dim IPINT As TIpiIPINT = New TIpiIPINT
                    infNFeDetImpostoIPI.Item = IPINT
                    IPINT.CST = TIpiIPINTCST.Item51
                ElseIf iCSTIPI = 52 Then
                    Dim IPINT As TIpiIPINT = New TIpiIPINT
                    infNFeDetImpostoIPI.Item = IPINT
                    IPINT.CST = TIpiIPINTCST.Item52
                ElseIf iCSTIPI = 53 Then
                    Dim IPINT As TIpiIPINT = New TIpiIPINT
                    infNFeDetImpostoIPI.Item = IPINT
                    IPINT.CST = TIpiIPINTCST.Item53
                ElseIf iCSTIPI = 54 Then
                    Dim IPINT As TIpiIPINT = New TIpiIPINT
                    infNFeDetImpostoIPI.Item = IPINT
                    IPINT.CST = TIpiIPINTCST.Item54
                ElseIf iCSTIPI = 55 Then
                    Dim IPINT As TIpiIPINT = New TIpiIPINT
                    infNFeDetImpostoIPI.Item = IPINT
                    IPINT.CST = TIpiIPINTCST.Item55
                End If

                If objTribDocItem.IPIVlrDevolvido <> 0 Then

                    objNFeInfNFeDet.impostoDevol = New TNFeInfNFeDetImpostoDevol
                    objNFeInfNFeDet.impostoDevol.IPI = New TNFeInfNFeDetImpostoDevolIPI
                    objNFeInfNFeDet.impostoDevol.pDevol = Replace(Format(objTribDocItem.pDevol * 100, "##0.00"), ",", ".")
                    objNFeInfNFeDet.impostoDevol.IPI.vIPIDevol = Replace(Format(objTribDocItem.IPIVlrDevolvido, "fixed"), ",", ".")

                End If

            End If

            Monta_NFiscal_Xml14 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml14 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml13(ByVal objNFeInfNFeDet As TNFeInfNFeDet, ByVal objItemNF As ItensNFiscal, ByVal objTribDocItem As TributacaoDocItem) As Long
        'preenche a parte de icms do item da nf

        Try

            If gobjApp.iDebug = 1 Then MsgBox("20")
            gobjApp.sErro = "20"
            gobjApp.sMsg1 = "vai iniciar a tributacao de ICMS"

            If objNFeInfNFeDet.prod.indTot = TNFeInfNFeDetProdIndTot.Item1 Then
                dvProdICMS = dvProdICMS + Arredonda_Moeda(objItemNF.PrecoUnitario * IIf(objItemNF.Quantidade = 0, 1, objItemNF.Quantidade))
            End If

            '*************** ICMS ***************************************

            Dim infNFeDetImpostoICMS As TNFeInfNFeDetImpostoICMS = New TNFeInfNFeDetImpostoICMS
            objNFeInfNFeDet.imposto.Items(0) = infNFeDetImpostoICMS

            'v2.0 - se o regime tributario for normal
            If objTribDocItem.RegimeTributario = 3 Then

                Select Case objTribDocItem.ICMSTipo

                    'tributacao integral
                    Case 1
                        Dim ICMS00 As New TNFeInfNFeDetImpostoICMSICMS00
                        infNFeDetImpostoICMS.Item = ICMS00
                        ICMS00.orig = objTribDocItem.OrigemMercadoria
                        ICMS00.CST = TNFeInfNFeDetImpostoICMSICMS00CST.Item00
                        ICMS00.modBC = TNFeInfNFeDetImpostoICMSICMS00ModBC.Item3
                        ICMS00.vBC = Replace(Format(objTribDocItem.ICMSBase, "fixed"), ",", ".")
                        ICMS00.pICMS = Replace(Format((objTribDocItem.ICMSAliquota - objTribDocItem.ICMSpFCP) * 100, "##0.00"), ",", ".")
                        ICMS00.vICMS = Replace(Format((objTribDocItem.ICMSValor - objTribDocItem.ICMSvFCP), "fixed"), ",", ".")

                        If objTribDocItem.ICMSvFCP <> 0 Then
                            ICMS00.pFCP = Replace(Format(objTribDocItem.ICMSpFCP * 100, "##0.00"), ",", ".")
                            ICMS00.vFCP = Replace(Format(objTribDocItem.ICMSvFCP, "fixed"), ",", ".")
                        End If

                        dICMSBaseTotal = Arredonda_Moeda(dICMSBaseTotal + objTribDocItem.ICMSBase)

                        'Tributado com substituição
                    Case 6
                        Dim ICMS10 As New TNFeInfNFeDetImpostoICMSICMS10
                        infNFeDetImpostoICMS.Item = ICMS10
                        ICMS10.orig = objTribDocItem.OrigemMercadoria
                        ICMS10.CST = TNFeInfNFeDetImpostoICMSICMS10CST.Item10
                        ICMS10.modBC = TNFeInfNFeDetImpostoICMSICMS00ModBC.Item3
                        ICMS10.vBC = Replace(Format(objTribDocItem.ICMSBase, "fixed"), ",", ".")
                        ICMS10.pICMS = Replace(Format((objTribDocItem.ICMSAliquota - objTribDocItem.ICMSpFCP) * 100, "##0.00"), ",", ".")
                        ICMS10.vICMS = Replace(Format((objTribDocItem.ICMSValor - objTribDocItem.ICMSvFCP), "fixed"), ",", ".")
                        ICMS10.modBCST = TNFeInfNFeDetImpostoICMSICMS10ModBCST.Item4
                        If objTribDocItem.ICMSSubstPercMVA > 0 Then
                            ICMS10.pMVAST = Replace(Format(objTribDocItem.ICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                        End If
                        ICMS10.vBCST = Replace(Format(objTribDocItem.ICMSSubstBase, "fixed"), ",", ".")
                        ICMS10.pICMSST = Replace(Format((objTribDocItem.ICMSSubstAliquota - objTribDocItem.ICMSpFCPST) * 100, "##0.00"), ",", ".")
                        ICMS10.vICMSST = Replace(Format((objTribDocItem.ICMSSubstValor - objTribDocItem.ICMSvFCPST), "fixed"), ",", ".")
                        If objTribDocItem.ICMSSubstPercRedBase > 0 Then
                            ICMS10.pRedBCST = Replace(Format(objTribDocItem.ICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                        End If

                        If objTribDocItem.ICMSvFCP <> 0 Then
                            ICMS10.vBCFCP = Replace(Format(objTribDocItem.ICMSvBCFCP, "fixed"), ",", ".")
                            ICMS10.pFCP = Replace(Format(objTribDocItem.ICMSpFCP * 100, "##0.00"), ",", ".")
                            ICMS10.vFCP = Replace(Format(objTribDocItem.ICMSvFCP, "fixed"), ",", ".")
                        End If

                        If objTribDocItem.ICMSvFCPST <> 0 Then
                            ICMS10.vBCFCPST = Replace(Format(objTribDocItem.ICMSvBCFCPST, "fixed"), ",", ".")
                            ICMS10.pFCPST = Replace(Format(objTribDocItem.ICMSpFCPST * 100, "##0.00"), ",", ".")
                            ICMS10.vFCPST = Replace(Format(objTribDocItem.ICMSvFCPST, "fixed"), ",", ".")
                        End If

                        dICMSBaseTotal = Arredonda_Moeda(dICMSBaseTotal + objTribDocItem.ICMSBase)

                        'Com redução da base de calc.
                    Case 7
                        Dim ICMS20 As New TNFeInfNFeDetImpostoICMSICMS20
                        infNFeDetImpostoICMS.Item = ICMS20
                        ICMS20.orig = objTribDocItem.OrigemMercadoria
                        ICMS20.CST = TNFeInfNFeDetImpostoICMSICMS20CST.Item20
                        ICMS20.modBC = TNFeInfNFeDetImpostoICMSICMS00ModBC.Item3
                        If objTribDocItem.ICMSPercRedBase > 0 Then
                            ICMS20.pRedBC = Replace(Format(objTribDocItem.ICMSPercRedBase * 100, "##0.00"), ",", ".")
                        Else
                            ICMS20.pRedBC = 0
                        End If
                        ICMS20.vBC = Replace(Format(objTribDocItem.ICMSBase * (1 - objTribDocItem.ICMSPercRedBase), "fixed"), ",", ".")
                        ICMS20.pICMS = Replace(Format((objTribDocItem.ICMSAliquota - objTribDocItem.ICMSpFCP) * 100, "##0.00"), ",", ".")
                        ICMS20.vICMS = Replace(Format((objTribDocItem.ICMSValor - objTribDocItem.ICMSvFCP), "fixed"), ",", ".")
                        ICMS20.motDesICMSSpecified = False
                        If objTribDocItem.ICMSMotivo <> 0 Then
                            ICMS20.vICMSDeson = Replace(Format(objTribDocItem.ICMSValorIsento, "fixed"), ",", ".")
                            ICMS20.motDesICMS = GetCode(Of TNFeInfNFeDetImpostoICMSICMS20MotDesICMS)(objTribDocItem.ICMSMotivo)
                        End If

                        If objTribDocItem.ICMSvFCP <> 0 Then
                            ICMS20.vBCFCP = Replace(Format(objTribDocItem.ICMSvBCFCP, "fixed"), ",", ".")
                            ICMS20.pFCP = Replace(Format(objTribDocItem.ICMSpFCP * 100, "##0.00"), ",", ".")
                            ICMS20.vFCP = Replace(Format(objTribDocItem.ICMSvFCP, "fixed"), ",", ".")
                        End If

                        dICMSBaseTotal = Arredonda_Moeda(dICMSBaseTotal + Arredonda_Moeda(objTribDocItem.ICMSBase * (1 - objTribDocItem.ICMSPercRedBase)))

                        'Isento com cobrança por subst.
                        'Não trib com cobrança por subst.
                    Case 9, 10
                        Dim ICMS30 As New TNFeInfNFeDetImpostoICMSICMS30
                        infNFeDetImpostoICMS.Item = ICMS30
                        ICMS30.orig = objTribDocItem.OrigemMercadoria
                        ICMS30.CST = TNFeInfNFeDetImpostoICMSICMS30CST.Item30
                        ICMS30.modBCST = TNFeInfNFeDetImpostoICMSICMS10ModBCST.Item4
                        If objTribDocItem.ICMSSubstPercMVA > 0 Then
                            ICMS30.pMVAST = Replace(Format(objTribDocItem.ICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                        End If
                        ICMS30.vBCST = Replace(Format(objTribDocItem.ICMSSubstBase, "fixed"), ",", ".")
                        ICMS30.pICMSST = Replace(Format((objTribDocItem.ICMSSubstAliquota - objTribDocItem.ICMSpFCPST) * 100, "##0.00"), ",", ".")
                        ICMS30.vICMSST = Replace(Format((objTribDocItem.ICMSSubstValor - objTribDocItem.ICMSpFCPST), "fixed"), ",", ".")
                        If objTribDocItem.ICMSSubstPercRedBase > 0 Then
                            ICMS30.pRedBCST = Replace(Format(objTribDocItem.ICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                        End If
                        ICMS30.motDesICMSSpecified = False
                        If objTribDocItem.ICMSMotivo <> 0 Then
                            ICMS30.vICMSDeson = Replace(Format(objTribDocItem.ICMSValorIsento, "fixed"), ",", ".")
                            ICMS30.motDesICMS = GetCode(Of TNFeInfNFeDetImpostoICMSICMS30MotDesICMS)(objTribDocItem.ICMSMotivo)
                            ICMS30.motDesICMSSpecified = True
                        End If

                        If objTribDocItem.ICMSvFCPST <> 0 Then
                            ICMS30.vBCFCPST = Replace(Format(objTribDocItem.ICMSvBCFCPST, "fixed"), ",", ".")
                            ICMS30.pFCPST = Replace(Format(objTribDocItem.ICMSpFCPST * 100, "##0.00"), ",", ".")
                            ICMS30.vFCPST = Replace(Format(objTribDocItem.ICMSvFCPST, "fixed"), ",", ".")
                        End If

                        'Isenta
                    Case 2
                        Dim ICMS40 As New TNFeInfNFeDetImpostoICMSICMS40
                        infNFeDetImpostoICMS.Item = ICMS40
                        ICMS40.orig = objTribDocItem.OrigemMercadoria
                        ICMS40.CST = TNFeInfNFeDetImpostoICMSICMS40CST.Item40
                        ICMS40.motDesICMSSpecified = False
                        If objTribDocItem.ICMSMotivo <> 0 Then
                            ICMS40.vICMSDeson = Replace(Format(objTribDocItem.ICMSValorIsento, "fixed"), ",", ".")
                            ICMS40.motDesICMS = GetCode(Of TNFeInfNFeDetImpostoICMSICMS40MotDesICMS)(objTribDocItem.ICMSMotivo)
                            ICMS40.motDesICMSSpecified = True
                        End If

                        'Não Tributado
                    Case 0

                        Dim ICMS40 As New TNFeInfNFeDetImpostoICMSICMS40
                        infNFeDetImpostoICMS.Item = ICMS40
                        ICMS40.orig = objTribDocItem.OrigemMercadoria
                        ICMS40.CST = TNFeInfNFeDetImpostoICMSICMS40CST.Item41
                        ICMS40.motDesICMSSpecified = False
                        ICMS40.motDesICMS = Nothing
                        If objTribDocItem.ICMSMotivo <> 0 Then
                            ICMS40.vICMSDeson = Replace(Format(objTribDocItem.ICMSValorIsento, "fixed"), ",", ".")
                            ICMS40.motDesICMS = GetCode(Of TNFeInfNFeDetImpostoICMSICMS40MotDesICMS)(objTribDocItem.ICMSMotivo)
                            ICMS40.motDesICMSSpecified = True
                        End If

                        'Com suspensão
                    Case 3

                        Dim ICMS40 As New TNFeInfNFeDetImpostoICMSICMS40
                        infNFeDetImpostoICMS.Item = ICMS40
                        ICMS40.orig = objTribDocItem.OrigemMercadoria
                        ICMS40.CST = TNFeInfNFeDetImpostoICMSICMS40CST.Item50
                        ICMS40.motDesICMSSpecified = False
                        If objTribDocItem.ICMSMotivo <> 0 Then
                            ICMS40.vICMSDeson = Replace(Format(objTribDocItem.ICMSValorIsento, "fixed"), ",", ".")
                            ICMS40.motDesICMS = GetCode(Of TNFeInfNFeDetImpostoICMSICMS40MotDesICMS)(objTribDocItem.ICMSMotivo)
                            ICMS40.motDesICMSSpecified = True
                        End If

                        'Com diferimento
                    Case 5
                        Dim ICMS51 As New TNFeInfNFeDetImpostoICMSICMS51
                        infNFeDetImpostoICMS.Item = ICMS51
                        ICMS51.orig = objTribDocItem.OrigemMercadoria
                        ICMS51.CST = TNFeInfNFeDetImpostoICMSICMS51CST.Item51
                        ICMS51.modBC = TNFeInfNFeDetImpostoICMSICMS51ModBC.Item3
                        ICMS51.modBCSpecified = True
                        ICMS51.pRedBC = Replace(Format(objTribDocItem.ICMSPercRedBase * 100, "##0.00"), ",", ".")
                        ICMS51.vBC = Replace(Format(objTribDocItem.ICMSBase * (1 - objTribDocItem.ICMSPercRedBase), "fixed"), ",", ".")
                        ICMS51.pICMS = Replace(Format((objTribDocItem.ICMSAliquota - objTribDocItem.ICMSpFCP) * 100, "##0.00"), ",", ".")

                        'nfe 3.10
                        'diferimento parcial
                        If objTribDocItem.ICMSPercDifer <> 0 Then
                            ICMS51.vICMSOp = Replace(Format(objTribDocItem.ICMS51ValorOp, "fixed"), ",", ".")
                            ICMS51.pDif = Replace(Format(objTribDocItem.ICMSPercDifer * 100, "##0.00"), ",", ".")
                            ICMS51.vICMSDif = Replace(Format(objTribDocItem.ICMSValorDif, "fixed"), ",", ".")
                        End If
                        'fim nfe 3.10

                        ICMS51.vICMS = Replace(Format((objTribDocItem.ICMSValor - objTribDocItem.ICMSvFCP), "fixed"), ",", ".")

                        If objTribDocItem.ICMSvFCP <> 0 Then
                            ICMS51.vBCFCP = Replace(Format(objTribDocItem.ICMSvBCFCP, "fixed"), ",", ".")
                            ICMS51.pFCP = Replace(Format(objTribDocItem.ICMSpFCP * 100, "##0.00"), ",", ".")
                            ICMS51.vFCP = Replace(Format(objTribDocItem.ICMSvFCP, "fixed"), ",", ".")
                        End If

                        dICMSBaseTotal = Arredonda_Moeda(dICMSBaseTotal + Arredonda_Moeda(objTribDocItem.ICMSBase * (1 - objTribDocItem.ICMSPercRedBase)))

                        'Cobrado anteriormente por subst.
                    Case 8
                        Dim ICMS60 As New TNFeInfNFeDetImpostoICMSICMS60
                        infNFeDetImpostoICMS.Item = ICMS60
                        ICMS60.orig = objTribDocItem.OrigemMercadoria
                        ICMS60.CST = TNFeInfNFeDetImpostoICMSICMS60CST.Item60
                        ICMS60.vBCSTRet = Replace(Format(objTribDocItem.ICMSSTCobrAntBase, "fixed"), ",", ".")
                        ICMS60.vICMSSTRet = Replace(Format((objTribDocItem.ICMSSTCobrAntValor - objTribDocItem.ICMSvFCPSTRet), "fixed"), ",", ".")

                        If objTribDocItem.ICMSvFCPSTRet <> 0 Then
                            ICMS60.vBCFCPSTRet = Replace(Format(objTribDocItem.ICMSvBCFCPSTRet, "fixed"), ",", ".")
                            ICMS60.pFCPSTRet = Replace(Format(objTribDocItem.ICMSpFCPSTRet * 100, "##0.00"), ",", ".")
                            ICMS60.vFCPSTRet = Replace(Format(objTribDocItem.ICMSvFCPSTRet, "fixed"), ",", ".")
                        End If

                        ICMS60.pST = Replace(Format((objTribDocItem.ICMSSTCobrAntAliquota - objTribDocItem.ICMSpFCPSTRet) * 100, "##0.00"), ",", ".")
                        ICMS60.vICMSSubstituto = "0.00"

                        'Com redução da base e cobr. Por subst.
                    Case 4
                        Dim ICMS70 As New TNFeInfNFeDetImpostoICMSICMS70
                        infNFeDetImpostoICMS.Item = ICMS70
                        ICMS70.orig = objTribDocItem.OrigemMercadoria
                        ICMS70.CST = TNFeInfNFeDetImpostoICMSICMS70CST.Item70
                        ICMS70.modBC = TNFeInfNFeDetImpostoICMSICMS00ModBC.Item3
                        ICMS70.pRedBC = Replace(Format(objTribDocItem.ICMSPercRedBase * 100, "##0.00"), ",", ".")
                        ICMS70.vBC = Replace(Format(objTribDocItem.ICMSBase * (1 - objTribDocItem.ICMSPercRedBase), "fixed"), ",", ".")
                        ICMS70.pICMS = Replace(Format((objTribDocItem.ICMSAliquota - objTribDocItem.ICMSpFCP) * 100, "##0.00"), ",", ".")
                        ICMS70.vICMS = Replace(Format((objTribDocItem.ICMSValor - objTribDocItem.ICMSvFCP), "fixed"), ",", ".")
                        ICMS70.modBCST = TNFeInfNFeDetImpostoICMSICMS10ModBCST.Item4
                        If objTribDocItem.ICMSSubstPercMVA > 0 Then
                            ICMS70.pMVAST = Replace(Format(objTribDocItem.ICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                        End If
                        ICMS70.vBCST = Replace(Format(objTribDocItem.ICMSSubstBase, "fixed"), ",", ".")
                        ICMS70.pICMSST = Replace(Format((objTribDocItem.ICMSSubstAliquota - objTribDocItem.ICMSpFCPST) * 100, "##0.00"), ",", ".")
                        ICMS70.vICMSST = Replace(Format((objTribDocItem.ICMSSubstValor - objTribDocItem.ICMSvFCPST), "fixed"), ",", ".")
                        If objTribDocItem.ICMSSubstPercRedBase > 0 Then
                            ICMS70.pRedBCST = Replace(Format(objTribDocItem.ICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                        End If
                        ICMS70.motDesICMSSpecified = False
                        If objTribDocItem.ICMSMotivo <> 0 Then
                            ICMS70.vICMSDeson = Replace(Format(objTribDocItem.ICMSValorIsento, "fixed"), ",", ".")
                            ICMS70.motDesICMS = GetCode(Of TNFeInfNFeDetImpostoICMSICMS70MotDesICMS)(objTribDocItem.ICMSMotivo)
                            ICMS70.motDesICMSSpecified = True
                        End If

                        If objTribDocItem.ICMSvFCP <> 0 Then
                            ICMS70.vBCFCP = Replace(Format(objTribDocItem.ICMSvBCFCP, "fixed"), ",", ".")
                            ICMS70.pFCP = Replace(Format(objTribDocItem.ICMSpFCP * 100, "##0.00"), ",", ".")
                            ICMS70.vFCP = Replace(Format(objTribDocItem.ICMSvFCP, "fixed"), ",", ".")
                        End If

                        If objTribDocItem.ICMSvFCPST <> 0 Then
                            ICMS70.vBCFCPST = Replace(Format(objTribDocItem.ICMSvBCFCPST, "fixed"), ",", ".")
                            ICMS70.pFCPST = Replace(Format(objTribDocItem.ICMSpFCPST * 100, "##0.00"), ",", ".")
                            ICMS70.vFCPST = Replace(Format(objTribDocItem.ICMSvFCPST, "fixed"), ",", ".")
                        End If

                        dICMSBaseTotal = Arredonda_Moeda(dICMSBaseTotal + Arredonda_Moeda(objTribDocItem.ICMSBase * (1 - objTribDocItem.ICMSPercRedBase)))

                        'Outras
                    Case 99
                        Dim ICMS90 As New TNFeInfNFeDetImpostoICMSICMS90
                        infNFeDetImpostoICMS.Item = ICMS90
                        ICMS90.orig = objTribDocItem.OrigemMercadoria
                        ICMS90.CST = TNFeInfNFeDetImpostoICMSICMS90CST.Item90
                        ICMS90.modBC = TNFeInfNFeDetImpostoICMSICMS00ModBC.Item3
                        If objTribDocItem.ICMSPercRedBase > 0 Then
                            ICMS90.pRedBC = Replace(Format(objTribDocItem.ICMSPercRedBase * 100, "##0.00"), ",", ".")
                        End If
                        ICMS90.vBC = Replace(Format(objTribDocItem.ICMSBase * (1 - objTribDocItem.ICMSPercRedBase), "fixed"), ",", ".")
                        ICMS90.pICMS = Replace(Format((objTribDocItem.ICMSAliquota - objTribDocItem.ICMSpFCP) * 100, "##0.00"), ",", ".")
                        ICMS90.vICMS = Replace(Format((objTribDocItem.ICMSValor - objTribDocItem.ICMSvFCP), "fixed"), ",", ".")
                        ICMS90.modBCST = TNFeInfNFeDetImpostoICMSICMS10ModBCST.Item4
                        If objTribDocItem.ICMSSubstPercMVA > 0 Then
                            ICMS90.pMVAST = Replace(Format(objTribDocItem.ICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                        End If
                        ICMS90.vBCST = Replace(Format(objTribDocItem.ICMSSubstBase, "fixed"), ",", ".")
                        ICMS90.pICMSST = Replace(Format((objTribDocItem.ICMSSubstAliquota - objTribDocItem.ICMSpFCPST) * 100, "##0.00"), ",", ".")
                        ICMS90.vICMSST = Replace(Format((objTribDocItem.ICMSSubstValor - objTribDocItem.ICMSpFCPST), "fixed"), ",", ".")
                        If objTribDocItem.ICMSSubstPercRedBase > 0 Then
                            ICMS90.pRedBCST = Replace(Format(objTribDocItem.ICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                        End If
                        ICMS90.motDesICMSSpecified = False
                        If objTribDocItem.ICMSMotivo <> 0 Then
                            ICMS90.vICMSDeson = Replace(Format(objTribDocItem.ICMSValorIsento, "fixed"), ",", ".")
                            ICMS90.motDesICMS = GetCode(Of TNFeInfNFeDetImpostoICMSICMS90MotDesICMS)(objTribDocItem.ICMSMotivo)
                            ICMS90.motDesICMSSpecified = True
                        End If

                        If objTribDocItem.ICMSvFCP <> 0 Then
                            ICMS90.vBCFCP = Replace(Format(objTribDocItem.ICMSvBCFCP, "fixed"), ",", ".")
                            ICMS90.pFCP = Replace(Format(objTribDocItem.ICMSpFCP * 100, "##0.00"), ",", ".")
                            ICMS90.vFCP = Replace(Format(objTribDocItem.ICMSvFCP, "fixed"), ",", ".")
                        End If

                        If objTribDocItem.ICMSvFCPST <> 0 Then
                            ICMS90.vBCFCPST = Replace(Format(objTribDocItem.ICMSvBCFCPST, "fixed"), ",", ".")
                            ICMS90.pFCPST = Replace(Format(objTribDocItem.ICMSpFCPST * 100, "##0.00"), ",", ".")
                            ICMS90.vFCPST = Replace(Format(objTribDocItem.ICMSvFCPST, "fixed"), ",", ".")
                        End If

                        dICMSBaseTotal = Arredonda_Moeda(dICMSBaseTotal + Arredonda_Moeda(objTribDocItem.ICMSBase * (1 - objTribDocItem.ICMSPercRedBase)))

                        'Partilha do ICMS - Não trib com cobrança por subst
                    Case 11
                        Dim ICMSPart As New TNFeInfNFeDetImpostoICMSICMSPart
                        infNFeDetImpostoICMS.Item = ICMSPart
                        ICMSPart.orig = objTribDocItem.OrigemMercadoria
                        ICMSPart.CST = TNFeInfNFeDetImpostoICMSICMSPartCST.Item10
                        ICMSPart.modBC = TNFeInfNFeDetImpostoICMSICMSPartModBC.Item3
                        ICMSPart.vBC = Replace(Format(objTribDocItem.ICMSBase, "fixed"), ",", ".")
                        ICMSPart.pICMS = Replace(Format(objTribDocItem.ICMSAliquota * 100, "##0.00"), ",", ".")
                        ICMSPart.vICMS = Replace(Format(objTribDocItem.ICMSValor, "fixed"), ",", ".")
                        ICMSPart.modBCST = TNFeInfNFeDetImpostoICMSICMSPartModBCST.Item4
                        If objTribDocItem.ICMSSubstPercMVA > 0 Then
                            ICMSPart.pMVAST = Replace(Format(objTribDocItem.ICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                        End If
                        ICMSPart.vBCST = Replace(Format(objTribDocItem.ICMSSubstBase, "fixed"), ",", ".")
                        ICMSPart.pICMSST = Replace(Format(objTribDocItem.ICMSSubstAliquota * 100, "##0.00"), ",", ".")
                        ICMSPart.vICMSST = Replace(Format(objTribDocItem.ICMSSubstValor, "fixed"), ",", ".")
                        If objTribDocItem.ICMSSubstPercRedBase > 0 Then
                            ICMSPart.pRedBCST = Replace(Format(objTribDocItem.ICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                        End If
                        ICMSPart.pBCOp = Replace(Format(objTribDocItem.ICMSpercBaseOperacaoPropria * 100, "##0.00"), ",", ".")

                        ICMSPart.UFST = GetCode(Of TUf)(objTribDocItem.ICMSUFDevidoST)
                        dICMSBaseTotal = Arredonda_Moeda(dICMSBaseTotal + objTribDocItem.ICMSBase)


                        'Partilha do ICMS - Outras
                    Case 13
                        Dim ICMSPart As New TNFeInfNFeDetImpostoICMSICMSPart
                        infNFeDetImpostoICMS.Item = ICMSPart
                        ICMSPart.orig = objTribDocItem.OrigemMercadoria
                        ICMSPart.CST = TNFeInfNFeDetImpostoICMSICMSPartCST.Item90
                        ICMSPart.modBC = TNFeInfNFeDetImpostoICMSICMSPartModBC.Item3
                        ICMSPart.vBC = Replace(Format(objTribDocItem.ICMSBase, "fixed"), ",", ".")
                        If objTribDocItem.ICMSPercRedBase > 0 Then
                            ICMSPart.pRedBC = Replace(Format(objTribDocItem.ICMSPercRedBase * 100, "##0.00"), ",", ".")
                        End If
                        ICMSPart.pICMS = Replace(Format(objTribDocItem.ICMSAliquota * 100, "##0.00"), ",", ".")
                        ICMSPart.vICMS = Replace(Format(objTribDocItem.ICMSValor, "fixed"), ",", ".")
                        ICMSPart.modBCST = TNFeInfNFeDetImpostoICMSICMSPartModBCST.Item4
                        If objTribDocItem.ICMSSubstPercMVA > 0 Then
                            ICMSPart.pMVAST = Replace(Format(objTribDocItem.ICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                        End If
                        ICMSPart.vBCST = Replace(Format(objTribDocItem.ICMSSubstBase, "fixed"), ",", ".")
                        ICMSPart.pICMSST = Replace(Format(objTribDocItem.ICMSSubstAliquota * 100, "##0.00"), ",", ".")
                        ICMSPart.vICMSST = Replace(Format(objTribDocItem.ICMSSubstValor, "fixed"), ",", ".")
                        If objTribDocItem.ICMSSubstPercRedBase > 0 Then
                            ICMSPart.pRedBCST = Replace(Format(objTribDocItem.ICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                        End If

                        ICMSPart.pBCOp = Replace(Format(objTribDocItem.ICMSpercBaseOperacaoPropria * 100, "##0.00"), ",", ".")

                        ICMSPart.UFST = GetCode(Of TUf)(objTribDocItem.ICMSUFDevidoST)
                        dICMSBaseTotal = Arredonda_Moeda(dICMSBaseTotal + objTribDocItem.ICMSBase)

                        'repasse de ICMSST retido ant. - Não tributado
                    Case 12
                        Dim ICMSST As New TNFeInfNFeDetImpostoICMSICMSST
                        infNFeDetImpostoICMS.Item = ICMSST
                        ICMSST.orig = objTribDocItem.OrigemMercadoria
                        ICMSST.CST = TNFeInfNFeDetImpostoICMSICMSSTCST.Item41
                        ICMSST.vBCSTRet = Replace(Format(objTribDocItem.ICMSvBCSTRet, "fixed"), ",", ".")
                        ICMSST.vICMSSTRet = Replace(Format(objTribDocItem.ICMSvICMSSTRet, "fixed"), ",", ".")
                        ICMSST.vBCSTDest = Replace(Format(objTribDocItem.ICMSvBCSTDest, "fixed"), ",", ".")
                        ICMSST.vICMSSTDest = Replace(Format(objTribDocItem.ICMSvICMSSTDest, "fixed"), ",", ".")

                End Select

            End If

            'v2.0 - se for regime tributario simples
            If objTribDocItem.RegimeTributario = 1 Or objTribDocItem.RegimeTributario = 2 Then

                Select Case objTribDocItem.ICMSSimplesTipo

                    'Trib. pelo Simples permissão de crédito
                    Case 1
                        Dim ICMSSN101 As New TNFeInfNFeDetImpostoICMSICMSSN101
                        infNFeDetImpostoICMS.Item = ICMSSN101
                        ICMSSN101.orig = objTribDocItem.OrigemMercadoria
                        ICMSSN101.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN101CSOSN.Item101
                        ICMSSN101.pCredSN = Replace(Format(objTribDocItem.ICMSpCredSN * 100, "##0.00"), ",", ".")
                        ICMSSN101.vCredICMSSN = Replace(Format(objTribDocItem.ICMSvCredSN, "fixed"), ",", ".")

                        'Trib. pelo Simples s/permissão de crédito
                    Case 2
                        Dim ICMSSN102 As New TNFeInfNFeDetImpostoICMSICMSSN102
                        infNFeDetImpostoICMS.Item = ICMSSN102
                        ICMSSN102.orig = objTribDocItem.OrigemMercadoria
                        ICMSSN102.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN102CSOSN.Item102

                        'Isenção do ICMS no Simples Nacional
                    Case 3
                        Dim ICMSSN102 As New TNFeInfNFeDetImpostoICMSICMSSN102
                        infNFeDetImpostoICMS.Item = ICMSSN102
                        ICMSSN102.orig = objTribDocItem.OrigemMercadoria
                        ICMSSN102.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN102CSOSN.Item103

                        'Simples Nacional - Imune
                    Case 7
                        Dim ICMSSN102 As New TNFeInfNFeDetImpostoICMSICMSSN102
                        infNFeDetImpostoICMS.Item = ICMSSN102
                        ICMSSN102.orig = objTribDocItem.OrigemMercadoria
                        ICMSSN102.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN102CSOSN.Item300

                        'Não tributada pelo Simples Nacional
                    Case 8
                        Dim ICMSSN102 As New TNFeInfNFeDetImpostoICMSICMSSN102
                        infNFeDetImpostoICMS.Item = ICMSSN102
                        ICMSSN102.orig = objTribDocItem.OrigemMercadoria
                        ICMSSN102.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN102CSOSN.Item400

                        'Simples c/permissão cred. e cobr. ICMS ST
                    Case 4
                        Dim ICMSSN201 As New TNFeInfNFeDetImpostoICMSICMSSN201
                        infNFeDetImpostoICMS.Item = ICMSSN201
                        ICMSSN201.orig = objTribDocItem.OrigemMercadoria
                        ICMSSN201.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN201CSOSN.Item201
                        ICMSSN201.modBCST = TNFeInfNFeDetImpostoICMSICMSSN201ModBCST.Item4
                        If objTribDocItem.ICMSSubstPercMVA > 0 Then
                            ICMSSN201.pMVAST = Replace(Format(objTribDocItem.ICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                        End If
                        If objTribDocItem.ICMSSubstPercRedBase > 0 Then
                            ICMSSN201.pRedBCST = Replace(Format(objTribDocItem.ICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                        End If
                        ICMSSN201.vBCST = Replace(Format(objTribDocItem.ICMSSubstBase, "fixed"), ",", ".")
                        ICMSSN201.pICMSST = Replace(Format((objTribDocItem.ICMSSubstAliquota - objTribDocItem.ICMSpFCPST) * 100, "##0.00"), ",", ".")
                        ICMSSN201.vICMSST = Replace(Format((objTribDocItem.ICMSSubstValor - objTribDocItem.ICMSvFCPST), "fixed"), ",", ".")
                        ICMSSN201.pCredSN = Replace(Format(objTribDocItem.ICMSpCredSN * 100, "##0.00"), ",", ".")
                        ICMSSN201.vCredICMSSN = Replace(Format(objTribDocItem.ICMSvCredSN, "fixed"), ",", ".")

                        If objTribDocItem.ICMSvFCPST <> 0 Then
                            ICMSSN201.vBCFCPST = Replace(Format(objTribDocItem.ICMSvBCFCPST, "fixed"), ",", ".")
                            ICMSSN201.pFCPST = Replace(Format(objTribDocItem.ICMSpFCPST * 100, "##0.00"), ",", ".")
                            ICMSSN201.vFCPST = Replace(Format(objTribDocItem.ICMSvFCPST, "fixed"), ",", ".")
                        End If

                        'Simples s/permissão cred. e cobr. ICMS ST
                    Case 5
                        Dim ICMSSN202 As New TNFeInfNFeDetImpostoICMSICMSSN202
                        infNFeDetImpostoICMS.Item = ICMSSN202
                        ICMSSN202.orig = objTribDocItem.OrigemMercadoria
                        ICMSSN202.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN202CSOSN.Item202
                        ICMSSN202.modBCST = TNFeInfNFeDetImpostoICMSICMSSN202ModBCST.Item4
                        If objTribDocItem.ICMSSubstPercMVA > 0 Then
                            ICMSSN202.pMVAST = Replace(Format(objTribDocItem.ICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                        End If
                        If objTribDocItem.ICMSSubstPercRedBase > 0 Then
                            ICMSSN202.pRedBCST = Replace(Format(objTribDocItem.ICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                        End If
                        ICMSSN202.vBCST = Replace(Format(objTribDocItem.ICMSSubstBase, "fixed"), ",", ".")
                        ICMSSN202.pICMSST = Replace(Format((objTribDocItem.ICMSSubstAliquota - objTribDocItem.ICMSpFCPST) * 100, "##0.00"), ",", ".")
                        ICMSSN202.vICMSST = Replace(Format((objTribDocItem.ICMSSubstValor - objTribDocItem.ICMSvFCPST), "fixed"), ",", ".")

                        If objTribDocItem.ICMSvFCPST <> 0 Then
                            ICMSSN202.vBCFCPST = Replace(Format(objTribDocItem.ICMSvBCFCPST, "fixed"), ",", ".")
                            ICMSSN202.pFCPST = Replace(Format(objTribDocItem.ICMSpFCPST * 100, "##0.00"), ",", ".")
                            ICMSSN202.vFCPST = Replace(Format(objTribDocItem.ICMSvFCPST, "fixed"), ",", ".")
                        End If

                        'Simples - Isenção ICMS e cobr. ICMS ST
                    Case 6
                        Dim ICMSSN202 As New TNFeInfNFeDetImpostoICMSICMSSN202
                        infNFeDetImpostoICMS.Item = ICMSSN202
                        ICMSSN202.orig = objTribDocItem.OrigemMercadoria
                        ICMSSN202.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN202CSOSN.Item203
                        ICMSSN202.modBCST = TNFeInfNFeDetImpostoICMSICMSSN202ModBCST.Item4
                        If objTribDocItem.ICMSSubstPercMVA > 0 Then
                            ICMSSN202.pMVAST = Replace(Format(objTribDocItem.ICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                        End If
                        If objTribDocItem.ICMSSubstPercRedBase > 0 Then
                            ICMSSN202.pRedBCST = Replace(Format(objTribDocItem.ICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                        End If
                        ICMSSN202.vBCST = Replace(Format(objTribDocItem.ICMSSubstBase, "fixed"), ",", ".")
                        ICMSSN202.pICMSST = Replace(Format((objTribDocItem.ICMSSubstAliquota - objTribDocItem.ICMSpFCPST) * 100, "##0.00"), ",", ".")
                        ICMSSN202.vICMSST = Replace(Format((objTribDocItem.ICMSSubstValor - objTribDocItem.ICMSvFCPST), "fixed"), ",", ".")

                        If objTribDocItem.ICMSvFCPST <> 0 Then
                            ICMSSN202.vBCFCPST = Replace(Format(objTribDocItem.ICMSvBCFCPST, "fixed"), ",", ".")
                            ICMSSN202.pFCPST = Replace(Format(objTribDocItem.ICMSpFCPST * 100, "##0.00"), ",", ".")
                            ICMSSN202.vFCPST = Replace(Format(objTribDocItem.ICMSvFCPST, "fixed"), ",", ".")
                        End If

                    Case 9
                        Dim ICMSSN500 As New TNFeInfNFeDetImpostoICMSICMSSN500
                        infNFeDetImpostoICMS.Item = ICMSSN500
                        ICMSSN500.orig = objTribDocItem.OrigemMercadoria
                        ICMSSN500.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN500CSOSN.Item500
                        ICMSSN500.vBCSTRet = Replace(Format(objTribDocItem.ICMSSTCobrAntBase, "fixed"), ",", ".")
                        ICMSSN500.vICMSSTRet = Replace(Format((objTribDocItem.ICMSSTCobrAntValor - objTribDocItem.ICMSvFCPSTRet), "fixed"), ",", ".")

                        ICMSSN500.pST = Replace(Format((objTribDocItem.ICMSSTCobrAntAliquota - objTribDocItem.ICMSpFCPSTRet) * 100, "##0.00"), ",", ".")

                        If objTribDocItem.ICMSvFCPSTRet <> 0 Then
                            ICMSSN500.vBCFCPSTRet = Replace(Format(objTribDocItem.ICMSvBCFCPSTRet, "fixed"), ",", ".")
                            ICMSSN500.pFCPSTRet = Replace(Format(objTribDocItem.ICMSpFCPSTRet * 100, "##0.00"), ",", ".")
                            ICMSSN500.vFCPSTRet = Replace(Format(objTribDocItem.ICMSvFCPSTRet, "fixed"), ",", ".")
                        End If
                        ICMSSN500.vICMSSubstituto = 0
                    Case 10
                        Dim ICMSSN900 As New TNFeInfNFeDetImpostoICMSICMSSN900
                        infNFeDetImpostoICMS.Item = ICMSSN900
                        ICMSSN900.orig = objTribDocItem.OrigemMercadoria
                        ICMSSN900.CSOSN = TNFeInfNFeDetImpostoICMSICMSSN900CSOSN.Item900
                        ICMSSN900.modBC = TNFeInfNFeDetImpostoICMSICMSSN900ModBC.Item3
                        ICMSSN900.vBC = Replace(Format(objTribDocItem.ICMSBase, "fixed"), ",", ".")
                        If objTribDocItem.ICMSPercRedBase > 0 Then
                            ICMSSN900.pRedBC = Replace(Format(objTribDocItem.ICMSPercRedBase * 100, "##0.00"), ",", ".")
                        End If
                        ICMSSN900.pICMS = Replace(Format(objTribDocItem.ICMSAliquota * 100, "##0.00"), ",", ".")
                        ICMSSN900.vICMS = Replace(Format(objTribDocItem.ICMSValor, "fixed"), ",", ".")
                        If objTribDocItem.ICMSSubstPercMVA > 0 Then
                            ICMSSN900.modBCST = TNFeInfNFeDetImpostoICMSICMSSN900ModBCST.Item4
                            ICMSSN900.pMVAST = Replace(Format(objTribDocItem.ICMSSubstPercMVA * 100, "##0.00"), ",", ".")
                        Else
                            ICMSSN900.modBCST = TNFeInfNFeDetImpostoICMSICMSSN900ModBCST.Item0
                        End If
                        If objTribDocItem.ICMSSubstPercRedBase > 0 Then
                            ICMSSN900.pRedBCST = Replace(Format(objTribDocItem.ICMSSubstPercRedBase * 100, "##0.00"), ",", ".")
                        End If
                        ICMSSN900.vBCST = Replace(Format(objTribDocItem.ICMSSubstBase, "fixed"), ",", ".")
                        ICMSSN900.pICMSST = Replace(Format((objTribDocItem.ICMSSubstAliquota - objTribDocItem.ICMSpFCPST) * 100, "##0.00"), ",", ".")
                        ICMSSN900.vICMSST = Replace(Format((objTribDocItem.ICMSSubstValor - objTribDocItem.ICMSvFCPST), "fixed"), ",", ".")
                        ICMSSN900.pCredSN = Replace(Format(objTribDocItem.ICMSpCredSN * 100, "##0.00"), ",", ".")
                        ICMSSN900.vCredICMSSN = Replace(Format(objTribDocItem.ICMSvCredSN, "fixed"), ",", ".")

                        If objTribDocItem.ICMSvFCPST <> 0 Then
                            ICMSSN900.vBCFCPST = Replace(Format(objTribDocItem.ICMSvBCFCPST, "fixed"), ",", ".")
                            ICMSSN900.pFCPST = Replace(Format(objTribDocItem.ICMSpFCPST * 100, "##0.00"), ",", ".")
                            ICMSSN900.vFCPST = Replace(Format(objTribDocItem.ICMSvFCPST, "fixed"), ",", ".")
                        End If

                        dICMSBaseTotal = Arredonda_Moeda(dICMSBaseTotal + objTribDocItem.ICMSBase)

                End Select

            End If

            'EC 87/2015
            If bICMSUFDest And (objTribDocItem.ICMSInterestPercPartilha <> 0 Or objTribDocItem.ICMSInterestVlrFCPUFDest <> 0 Or objTribDocItem.ICMSInterestVlrUFDest <> 0 Or objTribDocItem.ICMSInterestVlrUFRemet <> 0) Then

                Dim InfNFeDetImpostoICMSUFDest As New TNFeInfNFeDetImpostoICMSUFDest
                objNFeInfNFeDet.imposto.ICMSUFDest = InfNFeDetImpostoICMSUFDest

                With InfNFeDetImpostoICMSUFDest

                    .vBCUFDest = Replace(Format(objTribDocItem.ICMSInterestBCUFDest, "fixed"), ",", ".")
                    .pFCPUFDest = Replace(Format(objTribDocItem.ICMSInterestPercFCPUFDest * 100, "##0.00"), ",", ".")
                    .pICMSUFDest = Replace(Format(objTribDocItem.ICMSInterestAliqUFDest * 100, "##0.00"), ",", ".")
                    .pICMSInter = GetCode(Of TNFeInfNFeDetImpostoICMSUFDestPICMSInter)(Replace(Format(objTribDocItem.ICMSInterestAliq * 100, "##0.00"), ",", "."))
                    .pICMSInterPart = Replace(Format(objTribDocItem.ICMSInterestPercPartilha * 100, "##0.00"), ",", ".")
                    .vFCPUFDest = Replace(Format(objTribDocItem.ICMSInterestVlrFCPUFDest, "fixed"), ",", ".")
                    .vICMSUFDest = Replace(Format(objTribDocItem.ICMSInterestVlrUFDest, "fixed"), ",", ".")
                    .vICMSUFRemet = Replace(Format(objTribDocItem.ICMSInterestVlrUFRemet, "fixed"), ",", ".")

                    .vBCFCPUFDest = Replace(Format(objTribDocItem.ICMSInterestBCFCPUFDest, "fixed"), ",", ".")

                End With

            End If

            Monta_NFiscal_Xml13 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml13 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml12a(ByVal objNFeInfNFeDet As TNFeInfNFeDet, ByVal lNumIntDI As Long, ByRef dValorAFRMM As Double, ByRef objDI As DIInfo) As Long
        'preenche dados referentes à DI para o item de nf

        Dim resDIINFO As IEnumerable(Of DIInfo)
        Dim resImportCompl As IEnumerable(Of ImportCompl), objImportCompl As ImportCompl

        Try

            dValorAFRMM = 0

            resDIINFO = gobjApp.dbDadosNfe.ExecuteQuery(Of DIInfo) _
         ("SELECT * FROM DIINFO WHERE  NumIntDoc = {0}", lNumIntDI)

            If gobjApp.iDebug = 1 Then MsgBox("19.8")

            For Each objDIINFO In resDIINFO

                objDI = objDIINFO

                If gobjApp.iDebug = 1 Then MsgBox("19.9")

                Dim aNFeInfNFeDetProdDI(1) As TNFeInfNFeDetProdDI
                objNFeInfNFeDet.prod.DI = aNFeInfNFeDetProdDI

                Dim objNFeInfNFeDetProdDI As New TNFeInfNFeDetProdDI

                objNFeInfNFeDet.prod.DI(0) = objNFeInfNFeDetProdDI

                If objDIINFO.Numero.Substring(0, 6) <> "999999" Then
                    objNFeInfNFeDetProdDI.nDI = DesacentuaTexto(objDIINFO.Numero)
                Else
                    objNFeInfNFeDetProdDI.nDI = "NIHIL"
                End If
                objNFeInfNFeDetProdDI.dDI = Format(objDIINFO.Data, "yyyy-MM-dd")
                objNFeInfNFeDetProdDI.xLocDesemb = DesacentuaTexto(objDIINFO.LocalDesembaraco)

                objNFeInfNFeDetProdDI.UFDesemb = GetCode(Of TUf)(objDIINFO.UFDesembaraco)

                objNFeInfNFeDetProdDI.dDesemb = Format(objDIINFO.DataDesembaraco, "yyyy-MM-dd")
                objNFeInfNFeDetProdDI.cExportador = DesacentuaTexto(objDIINFO.CodExportador)

                objNFeInfNFeDetProdDI.tpViaTransp = GetCode(Of TNFeInfNFeDetProdDITpViaTransp)(CStr(objDIINFO.ViaTransp))

                'se for via maritima
                If objNFeInfNFeDetProdDI.tpViaTransp = TNFeInfNFeDetProdDITpViaTransp.Item1 Then

                    resImportCompl = gobjApp.dbDadosNfe.ExecuteQuery(Of ImportCompl) _
                    ("SELECT * FROM ImportCompl WHERE TipoDocOrigem = {0} AND NumIntDocOrigem = {1} AND Tipo = {2}", IMPORTCOMPL_ORIGEM_NF, objNFiscal.NumIntDoc, IMPORTCOMPL_TIPO_AFRMM)

                    For Each objImportCompl In resImportCompl

                        dValorAFRMM = objImportCompl.Valor
                        Exit For

                    Next

                End If

                objNFeInfNFeDetProdDI.tpIntermedio = GetCode(Of TNFeInfNFeDetProdDITpIntermedio)(CStr(objDIINFO.Intermedio))

                If Len(Trim(objDIINFO.CNPJAdquir)) <> 0 Then objNFeInfNFeDetProdDI.CNPJ = Trim(objDIINFO.CNPJAdquir)

                If Len(Trim(objDIINFO.UFAdquir)) <> 0 Then

                    objNFeInfNFeDetProdDI.UFTerceiro = Trim(objDIINFO.UFAdquir)
                    objNFeInfNFeDetProdDI.UFTerceiroSpecified = True

                End If

                If gobjApp.iDebug = 1 Then MsgBox("19.91")

                Dim aNFeInfNFeDetProdDIAdi(50) As TNFeInfNFeDetProdDIAdi
                objNFeInfNFeDetProdDI.adi = aNFeInfNFeDetProdDIAdi

                Exit For

            Next

            Monta_NFiscal_Xml12a = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml12a = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml12(ByVal objNFeInfNFeDet As TNFeInfNFeDet, ByVal objItemNF As ItensNFiscal, ByRef iExisteDI As Integer) As Long
        'preenche dados referentes à importação do item de nf

        Dim resItemAdicaoDIItemNF As IEnumerable(Of ItemAdicaoDIItemNF)
        Dim resItensAdicaoDI As IEnumerable(Of ItensAdicaoDI)
        Dim resAdicaoDI As IEnumerable(Of AdicaoDI)
        Dim dValorAFRMM As Double, lErro As Long, objDI As New DIInfo

        Try

            If gobjApp.iDebug = 1 Then MsgBox("19.1")
            gobjApp.sErro = "19.1"
            gobjApp.sMsg1 = "vai iniciar a declaracao de importacao"

            ''***********  Declaracao de IMPORTACAO ****************************
            Dim iExisteDIINFO As Integer = 0
            Dim dPesoItem As Double = 0

            resItemAdicaoDIItemNF = gobjApp.dbDadosNfe.ExecuteQuery(Of ItemAdicaoDIItemNF) _
            ("SELECT * FROM ItemAdicaoDIItemNF WHERE  NumIntItemNF = {0}", objItemNF.NumIntDoc)

            If gobjApp.iDebug = 1 Then MsgBox("19.2")

            For Each objItemAdicaoDIItemNF In resItemAdicaoDIItemNF

                If gobjApp.iDebug = 1 Then MsgBox("19.3")

                resItensAdicaoDI = gobjApp.dbDadosNfe.ExecuteQuery(Of ItensAdicaoDI) _
                ("SELECT * FROM ItensAdicaoDI WHERE  NumIntDoc = {0}", objItemAdicaoDIItemNF.NumIntItemAdicaoDI)

                If gobjApp.iDebug = 1 Then MsgBox("19.4")

                For Each objItensAdicaoDI In resItensAdicaoDI

                    dPesoItem = dPesoItem + objItensAdicaoDI.PesoBruto

                    If gobjApp.iDebug = 1 Then MsgBox("19.5")

                    resAdicaoDI = gobjApp.dbDadosNfe.ExecuteQuery(Of AdicaoDI) _
                    ("SELECT * FROM AdicaoDI WHERE  NumIntDoc = {0}", objItensAdicaoDI.NumIntAdicaoDI)

                    If gobjApp.iDebug = 1 Then MsgBox("19.6")

                    Dim iIndiceAdicaoDI As Integer = -1

                    For Each objAdicaoDI In resAdicaoDI

                        If gobjApp.iDebug = 1 Then MsgBox("19.7")

                        If iExisteDIINFO = 0 Then

                            iExisteDIINFO = 1

                            lErro = Monta_NFiscal_Xml12a(objNFeInfNFeDet, objAdicaoDI.NumIntDI, dValorAFRMM, objDI)
                            If lErro <> SUCESSO Then Throw New System.Exception("")

                        End If

                        iIndiceAdicaoDI = iIndiceAdicaoDI + 1

                        Dim objNFeInfNFeDetProdDIAdi As New TNFeInfNFeDetProdDIAdi
                        objNFeInfNFeDet.prod.DI(0).adi(iIndiceAdicaoDI) = objNFeInfNFeDetProdDIAdi

                        With objNFeInfNFeDetProdDIAdi
                            .nSeqAdic = objAdicaoDI.Seq
                            If objNFeInfNFeDet.prod.DI(0).nDI = "NIHIL" Then
                                .nAdicao = 999
                            Else
                                .nAdicao = objAdicaoDI.Seq
                            End If
                            .cFabricante = DesacentuaTexto(objAdicaoDI.CodFabricante)
                            '?????? .vDescDI = Replace(Format(????, "fixed"), ",", ".")
                            If Len(objAdicaoDI.NumDrawback) <> 0 Then .nDraw = DesacentuaTexto(objAdicaoDI.NumDrawback)
                        End With

                        If gobjApp.iDebug = 1 Then MsgBox("19.92")

                    Next

                Next

            Next

            If iExisteDIINFO <> 0 Then

                'se for via maritima
                If objNFeInfNFeDet.prod.DI(0).tpViaTransp = TNFeInfNFeDetProdDITpViaTransp.Item1 Then

                    Dim dRateioItem As Double

                    If objDI.PesoBrutoKG <> 0 Then
                        dRateioItem = dPesoItem / objDI.PesoBrutoKG
                    Else
                        If objNFiscal.ValorProdutos <> 0 Then
                            dRateioItem = objNFeInfNFeDet.prod.vProd / objNFiscal.ValorProdutos
                        Else
                            dRateioItem = 1
                        End If
                    End If

                    objNFeInfNFeDet.prod.DI(0).vAFRMM = Replace(Format(dValorAFRMM * dRateioItem, "fixed"), ",", ".")

                End If

            End If

            iExisteDI = iExisteDIINFO

            Monta_NFiscal_Xml12 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml12 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml11(ByVal objNFeInfNFeDet As TNFeInfNFeDet, ByVal objItemNF As ItensNFiscal, ByVal objProduto As Produto) As Long
        'preenche informacoes especificas de medicamentos

        Dim colMovEst As New Collection
        Dim resMovimentoEstoque As IEnumerable(Of MovimentoEstoque)
        Dim resItensNFiscalGrade As IEnumerable(Of ItensNFiscalGrade)
        Dim resRastreamentoMovto As IEnumerable(Of RastreamentoMovto)
        Dim resRastreamentoLote As IEnumerable(Of RastreamentoLote)
        Dim resRastroEstIni As IEnumerable(Of RastroEstIni)
        Dim resRastreamentoMovto1 As IEnumerable(Of RastreamentoMovto)
        Dim colRastro As New Collection, iIndice As Integer
        Dim objRastro As TNFeInfNFeDetProdRastro

        Try

            'seleciona os movimentos de estoque relacionados ao item da nota fiscal
            resMovimentoEstoque = gobjApp.dbDadosNfe.ExecuteQuery(Of MovimentoEstoque) _
            ("SELECT * FROM MovimentoEstoque WHERE NumIntDocOrigem = {0} AND TipoNumIntDocOrigem = {1} AND FilialEmpresa = {2} ORDER BY NumIntDoc", objItemNF.NumIntDoc, TIPO_ORIGEM_ITEMNF, gobjApp.iFilialEmpresa)

            For Each objMovEst In resMovimentoEstoque

                colMovEst.Add(objMovEst)

            Next

            'seleciona os itens de grade relacionados ao item da nota fiscal, se for o caso
            resItensNFiscalGrade = gobjApp.dbDadosNfe.ExecuteQuery(Of ItensNFiscalGrade) _
            ("SELECT * FROM ItensNFiscalGrade WHERE NumIntItemNF = {0} ", objItemNF.NumIntDoc)

            For Each objItensNFiscalGrade In resItensNFiscalGrade

                resMovimentoEstoque = gobjApp.dbDadosNfe.ExecuteQuery(Of MovimentoEstoque) _
                ("SELECT * FROM MovimentoEstoque WHERE NumIntDocOrigem = {0} AND TipoNumIntDocOrigem = {1} AND FilialEmpresa = {2} ORDER BY NumIntDoc", objItensNFiscalGrade.NumIntDoc, MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE, gobjApp.iFilialEmpresa)

                For Each objMovEst In resMovimentoEstoque

                    colMovEst.Add(objMovEst)

                Next


            Next

            For Each objMovEst In colMovEst

                resRastreamentoMovto = gobjApp.dbDadosNfe.ExecuteQuery(Of RastreamentoMovto) _
                ("SELECT * FROM RastreamentoMovto WHERE TipoDocOrigem = {0} AND NumIntDocOrigem = {1} ORDER BY RastreamentoMovto.NumIntDocLoteSerieIni", TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, objMovEst.NumIntDoc)

                For Each objRastreamentoMovto In resRastreamentoMovto

                    resRastreamentoLote = gobjApp.dbDadosNfe.ExecuteQuery(Of RastreamentoLote) _
                    ("SELECT * FROM RastreamentoLote WHERE NumIntDoc = {0}", objRastreamentoMovto.NumIntDocLote)

                    For Each objRastreamentoLote In resRastreamentoLote

                        'Dim objMed As TNFeInfNFeDetProdMed = New TNFeInfNFeDetProdMed
                        objRastro = New TNFeInfNFeDetProdRastro

                        'Dim iAchou As Integer = 0

                        'If Not objNFeInfNFeDet.prod.Items Is Nothing Then

                        '    For iIndice1 = 0 To objNFeInfNFeDet.prod.Items.Count - 1
                        '        If Not objNFeInfNFeDet.prod.Items(iIndice1) Is Nothing Then
                        '            If objNFeInfNFeDet.prod.Items(iIndice1).GetType Is GetType(TNFeInfNFeDetProdMed) Then
                        '                objMed = objNFeInfNFeDet.prod.Items(iIndice1)
                        '                If objMed.nLote = objRastreamentoLote.Lote Then
                        '                    iAchou = 1
                        '                    Exit For
                        '                End If

                        '            End If
                        '        Else
                        '            Exit For
                        '        End If
                        '    Next

                        'End If

                        'If iAchou = 0 Then

                        '    If objNFeInfNFeDet.prod.Items Is Nothing Then

                        '        Dim objObjeto1(9) As Object
                        '        objNFeInfNFeDet.prod.Items = objObjeto1

                        '    End If

                        '                        Dim iIndice1 As Integer
                        'Dim dQuant As Double

                        'dQuant = 0

                        'For iIndice1 = 0 To objNFeInfNFeDet.prod.Items.GetUpperBound(0)
                        '    If objNFeInfNFeDet.prod.Items(iIndice1) Is Nothing Then
                        '        Exit For
                        '    End If
                        'Next

                        'objMed = New TNFeInfNFeDetProdMed

                        'objNFeInfNFeDet.prod.Items(iIndice1) = objMed

                        'objMed.dFab = Format(objRastreamentoLote.DataFabricacao, "yyyy-MM-dd")
                        'objMed.dVal = Format(objRastreamentoLote.DataValidade, "yyyy-MM-dd")
                        'objMed.nLote = objRastreamentoLote.Lote


                        objRastro.dFab = Format(objRastreamentoLote.DataFabricacao, "yyyy-MM-dd")
                        objRastro.dVal = Format(objRastreamentoLote.DataValidade, "yyyy-MM-dd")
                        objRastro.nLote = objRastreamentoLote.Lote

                        'resRastroEstIni = gobjApp.dbDadosNfe.ExecuteQuery(Of RastroEstIni) _
                        '("SELECT * FROM RastroEstIni WHERE NumIntDocLote = {0}", objRastreamentoMovto.NumIntDocLote)

                        'For Each objRastroEstIni In resRastroEstIni

                        '    dQuant = dQuant + objRastroEstIni.Quantidade

                        'Next

                        'resRastreamentoMovto1 = gobjApp.dbDadosNfe.ExecuteQuery(Of RastreamentoMovto) _
                        '("SELECT R.* FROM RastreamentoMovto AS R,MovimentoEstoque AS M, TiposMovimentoEstoque AS T WHERE R.NumIntDocLote = {0} AND TipoDocOrigem = 0 AND R.NumIntDocOrigem = M.NumIntDoc AND T.EntradaOuSaida = 'E' AND M.FilialEmpresa = {1} AND M.TipoMov = T.Codigo", objRastreamentoMovto.NumIntDocLote, gobjApp.iFilialEmpresa)

                        'For Each objRastreamentoMovto1 In resRastreamentoMovto1
                        '    dQuant = dQuant + objRastreamentoMovto1.Quantidade
                        'Next

                        'objMed.qLote = dQuant
                        objRastro.qLote = objRastreamentoMovto.Quantidade
                        'objMed.vPMC = Replace(Format(objProduto.PrecoMaxConsumidor, "fixed"), ",", ".")

                        colRastro.Add(objRastro)
                        'End If

                        Exit For

                    Next

                Next

            Next

            If colRastro.Count > 0 Then

                ReDim objNFeInfNFeDet.prod.rastro(colRastro.Count - 1)
                iIndice = -1
                For Each objRastro In colRastro
                    iIndice = iIndice + 1
                    objNFeInfNFeDet.prod.rastro(iIndice) = objRastro
                Next

            End If

            If objProduto.PrecoMaxConsumidor > 0 Then
                Dim objMed(0) As TNFeInfNFeDetProdMed
                objMed(0) = New TNFeInfNFeDetProdMed
                objMed(0).cProdANVISA = objProduto.cProdANVISA
                objMed(0).vPMC = Replace(Format(objProduto.PrecoMaxConsumidor, "fixed"), ",", ".")
                objNFeInfNFeDet.prod.Items = objMed
            End If

            Monta_NFiscal_Xml11 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml11 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml10(ByVal objNFeInfNFeDet As TNFeInfNFeDet, ByVal objItemNF As ItensNFiscal, ByVal objProduto As Produto, ByVal iExisteDI As Integer) As Long
        'preenche dados da tributacao de um item da nf

        Dim objTribDocItem As New TributacaoDocItem
        Dim resTribDocItem As IEnumerable(Of TributacaoDocItem)
        Dim lErro As Long

        Try

            If gobjApp.iDebug = 1 Then MsgBox("18")
            gobjApp.sErro = "18"
            gobjApp.sMsg1 = "vai acessar a tabela TributacaoDocItem"

            resTribDocItem = gobjApp.dbDadosNfe.ExecuteQuery(Of TributacaoDocItem) _
            ("SELECT * FROM TributacaoDocItem WHERE TipoDoc = {0} AND NumIntDoc = {1} AND Item = {2}", TIPODOC_TRIB_NF, lNumIntNF, objItemNF.Item)

            For Each objTribDocItem In resTribDocItem
                Exit For
            Next

            gobjApp.sErro = "19"
            gobjApp.sMsg1 = "carrega os dados do Produto"

            objNFeInfNFeDet.prod.uTrib = objTribDocItem.UMTrib
            objNFeInfNFeDet.prod.qTrib = Replace(Format(objTribDocItem.QtdTrib, "######0.0000"), ",", ".")
            objNFeInfNFeDet.prod.vUnTrib = Replace(Format(objTribDocItem.ValorUnitTrib, "#########0.0000######"), ",", ".")

            If objTribDocItem.DescontoGrid > 0 Then
                objNFeInfNFeDet.prod.vDesc = Replace(Format(objTribDocItem.DescontoGrid, "fixed"), ",", ".")
                dTotalDescontoItem = dTotalDescontoItem + CDbl(Format(objTribDocItem.DescontoGrid, "fixed"))
            End If


            If objTribDocItem.ValorFreteItem > 0 Then
                objNFeInfNFeDet.prod.vFrete = Replace(Format(objTribDocItem.ValorFreteItem, "fixed"), ",", ".")
            End If

            If objTribDocItem.ValorSeguroItem > 0 Then
                objNFeInfNFeDet.prod.vSeg = Replace(Format(objTribDocItem.ValorSeguroItem, "fixed"), ",", ".")
            End If

            'v2.00
            If objTribDocItem.ValorOutrasDespesasItem > 0 Then
                objNFeInfNFeDet.prod.vOutro = Replace(Format(objTribDocItem.ValorOutrasDespesasItem, "fixed"), ",", ".")
            End If

            objNFeInfNFeDet.prod.CFOP = objTribDocItem.NaturezaOp

            If Len(objTribDocItem.FCI) = 36 Then
                objNFeInfNFeDet.prod.nFCI = objTribDocItem.FCI
            End If

            'v2.00
            If Len(Trim(objTribDocItem.IPICodProduto)) <> 0 Then

                objNFeInfNFeDet.prod.NCM = Right(Trim(objTribDocItem.IPICodProduto), 8)

            Else

                If objProduto.Natureza <> 8 Or Len(Trim(objProduto.IPICodigo)) <> 0 Then

                    objNFeInfNFeDet.prod.NCM = Right(Trim(objProduto.IPICodigo), 8)

                Else

                    If Len(Trim(objTribDocItem.ISSQN)) = 0 Then

                        objNFeInfNFeDet.prod.NCM = "00000000"

                    Else

                        'v2.00
                        objNFeInfNFeDet.prod.NCM = "00" '99

                    End If

                End If

            End If

            If Len(Trim(objTribDocItem.CEST)) <> 0 And objNFiscal.DataEmissao >= #4/1/2016# Then

                objNFeInfNFeDet.prod.CEST = Trim(objTribDocItem.CEST)

            End If

            If objTribDocItem.CNPJFab <> "" Then
                objNFeInfNFeDet.prod.CNPJFab = objTribDocItem.CNPJFab
            End If

            If objTribDocItem.indEscala = "" Then
                objNFeInfNFeDet.prod.indEscalaSpecified = False
            Else
                If objTribDocItem.indEscala = "S" Then
                    objNFeInfNFeDet.prod.indEscala = TNFeInfNFeDetProdIndEscala.S
                Else
                    objNFeInfNFeDet.prod.indEscala = TNFeInfNFeDetProdIndEscala.N
                End If
                objNFeInfNFeDet.prod.indEscalaSpecified = True
            End If

            If objTribDocItem.cBenef <> "" Then
                objNFeInfNFeDet.prod.cBenef = objTribDocItem.cBenef
            End If

            If gobjApp.iDebug = 1 Then MsgBox("19")

            'preenche dados referentes à tributos do item de nf
            Dim infNFeDetImposto As TNFeInfNFeDetImposto = New TNFeInfNFeDetImposto
            objNFeInfNFeDet.imposto = infNFeDetImposto

            If objNFiscal.DataEmissao >= #6/1/2013# And objTribDocItem.TotTrib <> 0 Then
                objNFeInfNFeDet.imposto.vTotTrib = Replace(Format(objTribDocItem.TotTrib, "fixed"), ",", ".")
            End If

            Dim objObjeto(2) As Object
            objNFeInfNFeDet.imposto.Items = objObjeto

            If Len(Trim(objTribDocItem.ISSQN)) = 0 Or objTribDocItem.ICMSTipo <> 0 Then

                'preenche a parte de icms do item da nf
                lErro = Monta_NFiscal_Xml13(objNFeInfNFeDet, objItemNF, objTribDocItem)
                If lErro <> SUCESSO Then Throw New System.Exception("")

            End If

            If sModelo <> "NFCe" And (objProduto.IPIIncide = 1 Or Len(Trim(objProduto.ISSQN)) = 0) Then 'deveria vir de tributacaodocitem

                'preenche a parte de ipi do item da nf
                lErro = Monta_NFiscal_Xml14(objNFeInfNFeDet, objItemNF, objTribDocItem)
                If lErro <> SUCESSO Then Throw New System.Exception("")

            End If

            If gobjApp.iDebug = 1 Then MsgBox("23")
            gobjApp.sErro = "23"
            gobjApp.sMsg1 = "vai iniciar o imposto de IMPORTACAO"


            '??? só deveria incluir para nfs com importacao
            '***********  Imposto de IMPORTACAO ****************************
            If objTribDocItem.IIValor <> 0 Or objTribDocItem.IIBase <> 0 Or objTribDocItem.IIDespAduaneira <> 0 Or objTribDocItem.IIIOF <> 0 Or iExisteDI = 1 Then

                Dim infNFeDetImpostoII As TNFeInfNFeDetImpostoII = New TNFeInfNFeDetImpostoII
                objNFeInfNFeDet.imposto.Items(2) = infNFeDetImpostoII


                infNFeDetImpostoII.vII = Replace(Format(objTribDocItem.IIValor, "fixed"), ",", ".")
                IIValorTotal = IIValorTotal + objTribDocItem.IIValor


                infNFeDetImpostoII.vBC = Replace(Format(objTribDocItem.IIBase, "fixed"), ",", ".")

                infNFeDetImpostoII.vDespAdu = Replace(Format(objTribDocItem.IIDespAduaneira, "fixed"), ",", ".")
                infNFeDetImpostoII.vIOF = objTribDocItem.IIIOF
                '*************************

            End If

            If sModelo <> "NFCe" Then

                'preenche a parte de pis do item da nf
                lErro = Monta_NFiscal_Xml15(objNFeInfNFeDet, objItemNF, objTribDocItem)
                If lErro <> SUCESSO Then Throw New System.Exception("")

                'preenche a parte de pis ST do item da nf
                lErro = Monta_NFiscal_Xml16(objNFeInfNFeDet, objItemNF, objTribDocItem)
                If lErro <> SUCESSO Then Throw New System.Exception("")

                'preenche a parte de COFINS do item da nf
                lErro = Monta_NFiscal_Xml17(objNFeInfNFeDet, objItemNF, objTribDocItem)
                If lErro <> SUCESSO Then Throw New System.Exception("")

                'preenche a parte de COFINS ST do item da nf
                lErro = Monta_NFiscal_Xml18(objNFeInfNFeDet, objItemNF, objTribDocItem)
                If lErro <> SUCESSO Then Throw New System.Exception("")

            End If

            'preenche a parte de ISS do item da nf
            lErro = Monta_NFiscal_Xml19(objNFeInfNFeDet, objItemNF, objTribDocItem)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            Monta_NFiscal_Xml10 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml10 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml24(ByVal objNFeInfNFeDet As TNFeInfNFeDet, ByVal objItemNF As ItensNFiscal) As Long
        'preenche dados de exportacao de um item da nf

        Dim resInfoAdicDocItemDetExport As IEnumerable(Of InfoAdicDocItemDetExport)
        Dim objInfoAdicDocItemDetExport As InfoAdicDocItemDetExport

        Try

            resInfoAdicDocItemDetExport = gobjApp.dbDadosNfe.ExecuteQuery(Of InfoAdicDocItemDetExport) _
            ("SELECT * FROM InfoAdicDocItemDetExport WHERE TipoDoc = {0} AND NumIntDocItem = {1} ORDER BY Seq", 0, objItemNF.NumIntDoc)

            Dim iIndice As Integer = 0

            For Each objInfoAdicDocItemDetExport In resInfoAdicDocItemDetExport

                If objNFeInfNFeDet.prod.detExport Is Nothing Then

                    Dim aNFeInfNFeDetProdDetExport(500) As TNFeInfNFeDetProdDetExport
                    objNFeInfNFeDet.prod.detExport = aNFeInfNFeDetProdDetExport

                    Dim objDetExport As TNFeInfNFeDetProdDetExport
                    objDetExport = objNFeInfNFeDet.prod.detExport(iIndice)
                    iIndice = iIndice + 1

                    With objDetExport

                        .nDraw = objInfoAdicDocItemDetExport.NumDrawback

                        .exportInd = New TNFeInfNFeDetProdDetExportExportInd
                        .exportInd.nRE = objInfoAdicDocItemDetExport.NumRegistExport
                        .exportInd.chNFe = objInfoAdicDocItemDetExport.ChvNFe
                        .exportInd.qExport = objInfoAdicDocItemDetExport.QuantExport

                    End With

                End If

            Next

            Monta_NFiscal_Xml24 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml24 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml9(ByVal objNFeInfNFeDet As TNFeInfNFeDet, ByVal objItemNF As ItensNFiscal, ByVal objNFiscal As NFeNFiscal) As Long
        'preenche dados de um item da nf

        Dim resProduto As IEnumerable(Of Produto)
        Dim sMsg As String
        Dim resMensagensRegra As IEnumerable(Of MensagensRegra)
        Dim sProduto As String
        Dim lErro As Long
        Dim resInfoAdicDocItem As IEnumerable(Of InfoAdicDocItem)
        Dim objInfoAdicDocItem As InfoAdicDocItem
        Dim objProduto As New Produto, iExisteDI As Integer = 0
        Dim resItemNF As IEnumerable(Of ItensNFiscal)
        Dim objItemNFOrig As ItensNFiscal

        Try

            Dim infNFeDetProd As TNFeInfNFeDetProd = New TNFeInfNFeDetProd

            objNFeInfNFeDet.prod = infNFeDetProd

            objNFeInfNFeDet.nItem = objItemNF.Item

            sProduto = ""
            Call Formata_Sem_Espaco(Trim(objItemNF.Produto), sProduto)
            objNFeInfNFeDet.prod.cProd = sProduto

            'solicitado em 28/5/12
            If gobjApp.objFilialEmpresa.CGC = "322101000189" Or gobjApp.objFilialEmpresa.CGC = "07341121000146" Then
                If objNFiscal.TabelaPreco = 3 Or objNFiscal.TabelaPreco = 4 Then
                    objNFeInfNFeDet.prod.cProd = "C" & sProduto
                End If

            End If

            objNFeInfNFeDet.prod.uCom = objItemNF.UnidadeMed
            objNFeInfNFeDet.prod.qCom = Replace(Format(objItemNF.Quantidade, "######0.0000"), ",", ".")
            objNFeInfNFeDet.prod.vUnCom = Replace(Format(objItemNF.PrecoUnitario, "#########0.0000######"), ",", ".")
            objNFeInfNFeDet.prod.vProd = Replace(Format((objItemNF.PrecoUnitario * IIf(objItemNF.Quantidade = 0, 1, objItemNF.Quantidade)), "fixed"), ",", ".")
            objNFeInfNFeDet.prod.xProd = Mid(DesacentuaTexto(Trim(objItemNF.DescricaoItem)), 1, 120)

            If gobjApp.iDebug = 1 Then MsgBox("17")
            gobjApp.sErro = "17"
            gobjApp.sMsg1 = "vai acessar a tabela Produtos"

            resProduto = gobjApp.dbDadosNfe.ExecuteQuery(Of Produto) _
            ("SELECT * FROM Produtos WHERE  Codigo = {0}", objItemNF.Produto)

            If gobjApp.iDebug = 1 Then MsgBox("17.1")
            gobjApp.sErro = "17.1"
            gobjApp.sMsg1 = "vai tratar os dados do produto"

            For Each objProduto In resProduto

                'se nao for servico
                If objProduto.Natureza <> 8 Then

                    'se estiver preenchido com valor diferente de zero
                    If objProduto.ExTIPI <> 0 Then
                        objNFeInfNFeDet.prod.EXTIPI = objProduto.ExTIPI
                    End If

                End If

                lErro = Produto_Trata_EAN(objProduto)
                If lErro <> SUCESSO Then Throw New System.Exception("")

                objNFeInfNFeDet.prod.cEAN = Trim(objProduto.CodigoBarras)
                If Len(Trim(objProduto.CodigoBarrasTrib)) > 0 Then
                    objNFeInfNFeDet.prod.cEANTrib = Trim(objProduto.CodigoBarrasTrib)
                Else
                    objNFeInfNFeDet.prod.cEANTrib = objNFeInfNFeDet.prod.cEAN
                End If

                'se for medicamento e tiver rastreamento
                If objProduto.Rastro <> PRODUTO_RASTRO_NENHUM And objProduto.ProdutoEspecifico = 3 Then

                    'preenche informacoes especificas de medicamentos
                    lErro = Monta_NFiscal_Xml11(objNFeInfNFeDet, objItemNF, objProduto)
                    If lErro <> SUCESSO Then Throw New System.Exception("")

                End If

                Exit For

            Next

            resInfoAdicDocItem = gobjApp.dbDadosNfe.ExecuteQuery(Of InfoAdicDocItem) _
            ("SELECT * FROM InfoAdicDocItem WHERE TipoDoc = 9 AND NumIntDocItem = {0}", objItemNF.NumIntDoc)

            objInfoAdicDocItem = resInfoAdicDocItem(0)

            If objInfoAdicDocItem Is Nothing Then
                Throw New System.Exception("O registro InfoAdicDocItem do produto não esta cadastrado. Produto = " & objItemNF.Produto)
            End If

            If gobjApp.iDebug = 1 Then MsgBox("19.95")

            If Len(Trim(objNFiscal.NumPedidoTerc)) > 0 Then objNFeInfNFeDet.prod.xPed = Right(Trim(DesacentuaTexto(objNFiscal.NumPedidoTerc)), 15)

            If Len(objInfoAdicDocItem.NumPedidoCompra) > 0 Then
                objNFeInfNFeDet.prod.xPed = Right(Trim(DesacentuaTexto(objInfoAdicDocItem.NumPedidoCompra)), 60)

            End If

            objNFeInfNFeDet.prod.nItemPed = objInfoAdicDocItem.ItemPedCompra
            If objInfoAdicDocItem.IncluiValorTotal = 0 Then
                objNFeInfNFeDet.prod.indTot = TNFeInfNFeDetProdIndTot.Item0
            Else
                objNFeInfNFeDet.prod.indTot = TNFeInfNFeDetProdIndTot.Item1
            End If

            'preenche dados da mensagem de um item da nf
            sMsg = ""
            resMensagensRegra = gobjApp.dbDadosNfe.ExecuteQuery(Of MensagensRegra) _
                ("SELECT * FROM MensagensRegra WHERE TipoDoc = 1 And NumIntDoc = {0}", objItemNF.NumIntDoc)

            For Each objMensagensRegra In resMensagensRegra
                sMsg = sMsg & objMensagensRegra.Mensagem
            Next

            Replace(sMsg, "|", " ")

            If Len(sMsg) > 0 Then
                objNFeInfNFeDet.infAdProd = DesacentuaTexto(sMsg)
            End If

            'preenche dados referentes à importação do item de nf
            lErro = Monta_NFiscal_Xml12(objNFeInfNFeDet, objItemNF, iExisteDI)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            'Se não tiver DI e for CFOP de importação verifica na nota original
            If iExisteDI = 0 And Left(objNFiscal.NaturezaOp, 1) = "3" And objNFiscal.NumIntNotaOriginal <> 0 Then

                resItemNF = gobjApp.dbDadosNfe.ExecuteQuery(Of ItensNFiscal) _
                ("SELECT * FROM ItensNFiscal AS IO WHERE IO.NumIntNF = {0} AND IO.Produto = {1} ORDER BY Item", objNFiscal.NumIntNotaOriginal, objItemNF.Produto)

                For Each objItemNFOrig In resItemNF
                    lErro = Monta_NFiscal_Xml12(objNFeInfNFeDet, objItemNFOrig, iExisteDI)
                    If lErro <> SUCESSO Then Throw New System.Exception("")
                    Exit For
                Next

            End If

            'preenche dados da tributacao de um item da nf
            lErro = Monta_NFiscal_Xml10(objNFeInfNFeDet, objItemNF, objProduto, iExisteDI)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            'preenche dados referentes à exportação do item de nf
            lErro = Monta_NFiscal_Xml24(objNFeInfNFeDet, objItemNF)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            Monta_NFiscal_Xml9 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml9 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml8(ByVal infNFe As TNFeInfNFe, ByVal objNFiscal As NFeNFiscal) As Long
        'preenche dados dos itens da nf

        Dim lErro As Long
        Dim resItemNF As IEnumerable(Of ItensNFiscal)
        Dim objItemNF As ItensNFiscal
        Dim iIndice As Integer, iNumItensNF As Integer

        Try

            'lNumIntNFiscalParam
            resItemNF = gobjApp.dbDadosNfe.ExecuteQuery(Of ItensNFiscal) _
            ("SELECT * FROM ItensNFiscal WHERE  NumIntNF = {0} ORDER BY Item", lNumIntNF)

            iNumItensNF = resItemNF.Count


            Dim NFDet(iNumItensNF) As TNFeInfNFeDet
            infNFe.det() = NFDet

            resItemNF = gobjApp.dbDadosNfe.ExecuteQuery(Of ItensNFiscal) _
            ("SELECT * FROM ItensNFiscal WHERE  NumIntNF = {0} ORDER BY Item", lNumIntNF)

            iIndice = -1

            dValorServPIS = 0
            dValorServCOFINS = 0
            dServNTribICMS = 0
            dValorPIS = 0
            dValorCOFINS = 0
            IIValorTotal = 0
            dICMSBaseTotal = 0

            dTotalDescontoItem = 0
            dvProdICMS = 0

            For Each objItemNF In resItemNF

                iIndice = iIndice + 1

                Dim infNFeDet As TNFeInfNFeDet = New TNFeInfNFeDet
                NFDet(iIndice) = infNFeDet

                lErro = Monta_NFiscal_Xml9(NFDet(iIndice), objItemNF, objNFiscal)
                If lErro <> SUCESSO Then Throw New System.Exception("")

            Next

            Monta_NFiscal_Xml8 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml8 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml7(ByVal infNFe As TNFeInfNFe) As Long
        'preenche local de entrega

        Dim resRetEnt As IEnumerable(Of RetiradaEntrega)
        Dim objRetEnt As RetiradaEntrega
        Dim objEndereco As Endereco
        Dim resEndEnt As IEnumerable(Of Endereco)
        Dim lErro As Long

        Try

            resRetEnt = gobjApp.dbDadosNfe.ExecuteQuery(Of RetiradaEntrega) _
            ("SELECT * FROM RetiradaEntrega WHERE TipoDoc = {0} AND NumIntDoc = {1}", 0, objNFiscal.NumIntDoc)

            For Each objRetEnt In resRetEnt

                resEndEnt = gobjApp.dbDadosNfe.ExecuteQuery(Of Endereco) _
                ("SELECT * FROM Enderecos WHERE Codigo = {0}", objRetEnt.EnderecoEnt)

                For Each objEndereco In resEndEnt

                    If Len(objEndereco.Logradouro) > 0 Or Len(objEndereco.Endereco) > 0 Then

                        Dim objLocal As TLocal = New TLocal
                        infNFe.entrega = objLocal

                        If Len(objRetEnt.CNPJCPFEnt) = 14 Then
                            infNFe.entrega.ItemElementName = ItemChoiceType4.CNPJ
                        Else
                            infNFe.entrega.ItemElementName = ItemChoiceType4.CPF
                        End If

                        infNFe.entrega.Item = objRetEnt.CNPJCPFEnt

                        lErro = gobjApp.Local_ObterBD(objLocal, objEndereco)
                        If lErro <> SUCESSO Then Throw New System.Exception("")

                    End If

                    Exit For

                Next

                Exit For
            Next

            Monta_NFiscal_Xml7 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml7 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml6(ByVal infNFe As TNFeInfNFe) As Long
        'preenche local de retirada

        Dim resRetEnt As IEnumerable(Of RetiradaEntrega)
        Dim objRetEnt As RetiradaEntrega
        Dim objEndereco As Endereco
        Dim resEndRet As IEnumerable(Of Endereco)
        Dim lErro As Long

        Try

            'v2.00
            resRetEnt = gobjApp.dbDadosNfe.ExecuteQuery(Of RetiradaEntrega) _
            ("SELECT * FROM RetiradaEntrega WHERE TipoDoc = {0} AND NumIntDoc = {1}", 0, objNFiscal.NumIntDoc)

            For Each objRetEnt In resRetEnt

                resEndRet = gobjApp.dbDadosNfe.ExecuteQuery(Of Endereco) _
                ("SELECT * FROM Enderecos WHERE Codigo = {0}", objRetEnt.EnderecoRet)

                For Each objEndereco In resEndRet

                    If Len(objEndereco.Logradouro) > 0 Or Len(objEndereco.Endereco) > 0 Then

                        Dim objLocal As TLocal = New TLocal
                        infNFe.retirada = objLocal

                        If Len(objRetEnt.CNPJCPFRet) = 14 Then
                            infNFe.retirada.ItemElementName = ItemChoiceType4.CNPJ
                        Else
                            infNFe.retirada.ItemElementName = ItemChoiceType4.CPF
                        End If

                        infNFe.retirada.Item = objRetEnt.CNPJCPFRet

                        lErro = gobjApp.Local_ObterBD(objLocal, objEndereco)
                        If lErro <> SUCESSO Then Throw New System.Exception("")

                    End If

                    Exit For

                Next

                Exit For

            Next

            Monta_NFiscal_Xml6 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml6 = SUCESSO

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml5(ByVal infNFeDest As TNFeInfNFeDest) As Long
        'preenche dados do destinatario

        Dim resFiliaisClientes As IEnumerable(Of FiliaisCliente)
        Dim resCliente As IEnumerable(Of Cliente)
        Dim resFiliaisFornecedores As IEnumerable(Of FiliaisFornecedore)
        Dim resFornecedor As IEnumerable(Of Fornecedore)
        Dim sIE As String, sEmail As String
        Dim lEndDest As Long
        Dim lEndEntrega As Long
        Dim lErro As Long, iIENaoContrib As Integer, iIEISento As Integer
        Dim sIdEstrangeiro As String, bEstrangeiro As Boolean = False

        Try
            sIE = ""
            sEmail = ""
            iIENaoContrib = 0
            sIdEstrangeiro = ""

            'se o destinatrio for o cliente
            If objNFiscal.Destinatario = 1 Or (objNFiscal.Destinatario = 0 And objNFiscal.Cliente <> 0) Then

                If gobjApp.iDebug = 1 Then MsgBox("12")
                gobjApp.sErro = "12"
                gobjApp.sMsg1 = "vai acessar a tabela Clientes"

                resCliente = gobjApp.dbDadosNfe.ExecuteQuery(Of Cliente) _
                ("SELECT * FROM Clientes WHERE Codigo = {0}", objNFiscal.Cliente)

                For Each objCliente In resCliente
                    infNFeDest.xNome = DesacentuaTexto(Trim(objCliente.RazaoSocial))
                    Exit For
                Next

                resFiliaisClientes = gobjApp.dbDadosNfe.ExecuteQuery(Of FiliaisCliente) _
                ("SELECT * FROM FiliaisClientes WHERE CodCliente = {0} AND CodFilial = {1}", objNFiscal.Cliente, objNFiscal.FilialCli)

                For Each objFiliaisClientes In resFiliaisClientes

                    lEndDest = objFiliaisClientes.Endereco
                    lEndEntrega = objFiliaisClientes.EnderecoEntrega

                    If Len(Trim(objFiliaisClientes.IdEstrangeiro)) <> 0 Then
                        sIdEstrangeiro = Trim(objFiliaisClientes.IdEstrangeiro)
                        infNFeDest.ItemElementName = ItemChoiceType5.idEstrangeiro
                        infNFeDest.Item = sIdEstrangeiro
                    Else
                        infNFeDest.ItemElementName = ItemChoiceType5.CPF
                        If Len(objFiliaisClientes.CGC) = 11 Then
                            infNFeDest.ItemElementName = ItemChoiceType5.CPF
                        Else
                            infNFeDest.ItemElementName = ItemChoiceType5.CNPJ
                        End If

                        infNFeDest.Item = objFiliaisClientes.CGC
                        scDest = objFiliaisClientes.CGC

                    End If

                    Call Formata_String_Numero(objFiliaisClientes.InscricaoEstadual, sIE)

                    iIENaoContrib = objFiliaisClientes.IENaoContrib
                    iIEISento = objFiliaisClientes.IEIsento

                    If Len(Trim(objFiliaisClientes.InscricaoSuframa)) > 0 Then
                        infNFeDest.ISUF = objFiliaisClientes.InscricaoSuframa
                    End If

                    Exit For

                Next

                'se o desti
            ElseIf objNFiscal.Destinatario = 2 Or (objNFiscal.Destinatario = 0 And objNFiscal.Fornecedor <> 0) Then

                If gobjApp.iDebug = 1 Then MsgBox("13")
                gobjApp.sErro = "13"
                gobjApp.sMsg1 = "vai acessar a tabela Fornecedores"


                resFornecedor = gobjApp.dbDadosNfe.ExecuteQuery(Of Fornecedore) _
                ("SELECT * FROM Fornecedores WHERE Codigo = {0}", objNFiscal.Fornecedor)

                For Each objFornecedor In resFornecedor
                    infNFeDest.xNome = DesacentuaTexto(Trim(objFornecedor.RazaoSocial))
                    Exit For
                Next


                If gobjApp.iDebug = 1 Then MsgBox("14")
                gobjApp.sErro = "14"
                gobjApp.sMsg1 = "vai acessar a tabela FiliaisFornecedores"

                resFiliaisFornecedores = gobjApp.dbDadosNfe.ExecuteQuery(Of FiliaisFornecedore) _
                ("SELECT * FROM FiliaisFornecedores WHERE CodFornecedor = {0} AND CodFilial = {1}", objNFiscal.Fornecedor, objNFiscal.FilialForn)

                For Each objFiliaisFornecedores In resFiliaisFornecedores

                    lEndDest = objFiliaisFornecedores.Endereco

                    If Len(Trim(objFiliaisFornecedores.IdEstrangeiro)) <> 0 Then
                        sIdEstrangeiro = Trim(objFiliaisFornecedores.IdEstrangeiro)
                        infNFeDest.ItemElementName = ItemChoiceType5.idEstrangeiro
                        infNFeDest.Item = sIdEstrangeiro
                    Else

                        If Len(Trim(objFiliaisFornecedores.CGC)) <> 0 Then

                            If Len(objFiliaisFornecedores.CGC) = 11 Then
                                infNFeDest.ItemElementName = ItemChoiceType5.CPF
                            Else
                                infNFeDest.ItemElementName = ItemChoiceType5.CNPJ
                            End If
                            infNFeDest.Item = objFiliaisFornecedores.CGC

                        End If

                    End If

                    Call Formata_String_Numero(objFiliaisFornecedores.InscricaoEstadual, sIE)

                    iIENaoContrib = objFiliaisFornecedores.IENaoContrib
                    iIEISento = objFiliaisFornecedores.IEIsento

                    Exit For

                Next

            Else

                lEndDest = gobjApp.lEndereco
                infNFeDest.xNome = DesacentuaTexto(Trim(gobjApp.sRazaoSocial))
                Call Formata_String_Numero(gobjApp.objFilialEmpresa.InscricaoEstadual, sIE)

            End If

            Dim enderDest As TEndereco = New TEndereco
            infNFeDest.enderDest = enderDest

            lErro = gobjApp.Endereco_ObterBD(enderDest, lEndDest, sEmail)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            'If enderDest.cPaisSpecified Then

            If enderDest.cPais <> 1058 Then

                infNFeDest.ItemElementName = ItemChoiceType5.idEstrangeiro
                infNFeDest.Item = sIdEstrangeiro
                bEstrangeiro = True

            End If

            'End If

            'v2.00
            If Len(sEmail) <> 0 Then
                infNFeDest.email = Trim(DesacentuaTexto(sEmail))
            End If

            If gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then infNFeDest.xNome = "NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"

            If sModelo = "NFCe" Then

                infNFeDest.indIEDest = TNFeInfNFeDestIndIEDest.Item9

            Else

                'If gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then

                '    infNFeDest.indIEDest = TNFeInfNFeDestIndIEDest.Item9

                'Else

                If Len(sIE) <> 0 Then infNFeDest.IE = sIE

                If bEstrangeiro Or iIENaoContrib = 1 Then
                    infNFeDest.indIEDest = TNFeInfNFeDestIndIEDest.Item9
                Else
                    If Len(sIE) <> 0 Then
                        infNFeDest.indIEDest = TNFeInfNFeDestIndIEDest.Item1
                    Else
                        infNFeDest.indIEDest = TNFeInfNFeDestIndIEDest.Item2
                    End If
                End If

            End If
            'End If

            'EC 87/2015
            Dim dtDataEC872015 As Date = CDate("01/01/2016")
            Dim resCRFatConfig As IEnumerable(Of CRFatConfig)
            Dim objCRFatConfig As CRFatConfig

            resCRFatConfig = gobjApp.dbDadosNfe.ExecuteQuery(Of CRFatConfig) _
            ("SELECT * FROM CRFatConfig WHERE Codigo = 'DATA_EC_87_2015'")

            For Each objCRFatConfig In resCRFatConfig
                dtDataEC872015 = CDate(objCRFatConfig.Conteudo)
                Exit For
            Next

            bICMSUFDest = False
            If (gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Or (gobjApp.iNFeAmbiente = NFE_AMBIENTE_PRODUCAO And objNFiscal.DataEmissao >= dtDataEC872015)) And infNFeDest.indIEDest = TNFeInfNFeDestIndIEDest.Item9 And objTributacaoDoc.IndConsumidorFinal = 1 And objTributacaoDoc.Destino = 2 And (objNFiscal.Tipo = 2 Or objNFiscal.TipoNFiscal = DOCINFO_NFIEDV) And objNFiscal.Cliente <> 0 Then

                bICMSUFDest = True

            End If
            'EC 87/2015

            Monta_NFiscal_Xml5 = SUCESSO

        Catch ex As Exception
            Monta_NFiscal_Xml5 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try
    End Function

    Private Function Monta_NFiscal_Xml4(ByVal infNFeEmit As TNFeInfNFeEmit) As Long
        'preenche dados do emitente

        Dim resCRFatConfig As IEnumerable(Of CRFatConfig)
        Dim resContSubst As IEnumerable(Of ContribuinteSubstituto)
        Dim sIE As String, sIEST As String, sIM As String, sCNAE As String

        Try

            sIE = ""
            sIEST = ""
            sIM = ""
            sCNAE = ""

            infNFeEmit.ItemElementName = ItemChoiceType.CNPJ
            infNFeEmit.Item = gobjApp.objFilialEmpresa.CGC

            sidNFCECSC = gobjApp.objFilialEmpresa.idNFCECSC
            sNFCECSC = gobjApp.objFilialEmpresa.NFCECSC

            If gobjApp.iDebug = 1 Then MsgBox("10")

            gobjApp.sErro = "10"
            gobjApp.sMsg1 = "vai acessar as tabelas Empresas e AdmConfig"

            infNFeEmit.xNome = gobjApp.sRazaoSocial
            infNFeEmit.xFant = gobjApp.sNomeFantasia

            Dim enderEmit As TEnderEmi = New TEnderEmi
            infNFeEmit.enderEmit = enderEmit

            If Len(gobjApp.objEndereco.Logradouro) > 0 Then
                infNFeEmit.enderEmit.xLgr = DesacentuaTexto(Left(IIf(Len(gobjApp.objEndereco.TipoLogradouro) > 0, gobjApp.objEndereco.TipoLogradouro & " ", "") & gobjApp.objEndereco.Logradouro, 60))
                infNFeEmit.enderEmit.nro = gobjApp.objEndereco.Numero
                If Len(DesacentuaTexto(gobjApp.objEndereco.Complemento)) > 0 Then infNFeEmit.enderEmit.xCpl = DesacentuaTexto(gobjApp.objEndereco.Complemento)
            Else
                infNFeEmit.enderEmit.xLgr = DesacentuaTexto(gobjApp.objEndereco.Endereco)
                infNFeEmit.enderEmit.nro = "0"
            End If
            If Len(gobjApp.objEndereco.Bairro) = 0 Then
                infNFeEmit.enderEmit.xBairro = "a"
            Else
                infNFeEmit.enderEmit.xBairro = Trim(DesacentuaTexto(gobjApp.objEndereco.Bairro))
            End If

            'se for Brasil 
            If gobjApp.objPais.CodBacen = 1058 Then

                infNFeEmit.enderEmit.cMun = gobjApp.objCidade.CodIBGE
                infNFeEmit.enderEmit.xMun = DesacentuaTexto(gobjApp.objCidade.Descricao)
                '                infNFeEmit.enderEmit.UF = gobjapp.objEndereco.SiglaEstado

                infNFeEmit.enderEmit.UF = GetCode(Of TUf)(gobjApp.objEndereco.SiglaEstado)

                If Len(gobjApp.objEndereco.CEP) > 0 Then
                    infNFeEmit.enderEmit.CEP = gobjApp.objEndereco.CEP
                End If
                If Len(gobjApp.objEndereco.TelNumero1) > 0 Then
                    Call Formata_String_Numero(IIf(Len(CStr(gobjApp.objEndereco.TelDDD1)) > 0, CStr(gobjApp.objEndereco.TelDDD1), "") + gobjApp.objEndereco.TelNumero1, infNFeEmit.enderEmit.fone)
                ElseIf Len(gobjApp.objEndereco.Telefone1) > 0 Then
                    Call Formata_String_Numero(gobjApp.objEndereco.Telefone1, infNFeEmit.enderEmit.fone)
                    'Else
                    '    infNFeEmit.enderEmit.fone = "99999999"
                End If

            Else
                infNFeEmit.enderEmit.cMun = "9999999"
                infNFeEmit.enderEmit.xMun = "EXTERIOR"
                infNFeEmit.enderEmit.UF = GetCode(Of TUf)("EX")
            End If

            If gobjApp.iDebug = 1 Then MsgBox("11")
            gobjApp.sErro = "11"
            gobjApp.sMsg1 = "vai acessar os dados do destinatario"

            infNFeEmit.enderEmit.cPais = TEnderEmiCPais.Item1058
            infNFeEmit.enderEmit.cPaisSpecified = True
            infNFeEmit.enderEmit.xPais = TEnderEmiXPais.Brasil
            infNFeEmit.enderEmit.xPaisSpecified = True
            Call Formata_String_Numero(gobjApp.objFilialEmpresa.InscricaoEstadual, sIE)
            infNFeEmit.IE = sIE

            resContSubst = gobjApp.dbDadosNfe.ExecuteQuery(Of ContribuinteSubstituto) _
            ("SELECT * FROM ContribuinteSubstituto WHERE FilialEmpresa = {0} AND UF = {1}", gobjApp.iFilialEmpresa, gobjApp.objEndereco.SiglaEstado)

            For Each objContSubst In resContSubst
                Call Formata_String_Numero(objContSubst.InscricaoEstadual, sIEST)
                infNFeEmit.IEST = sIEST
                Exit For
            Next

            Call Formata_String_Numero(gobjApp.objFilialEmpresa.InscricaoMunicipal, sIM)
            If Len(sIM) > 0 Then
                infNFeEmit.IM = sIM
                Call Formata_String_Numero(gobjApp.objFilialEmpresa.CNAE, sCNAE)
                infNFeEmit.CNAE = sCNAE
            End If

            'v2.00
            If gobjApp.objFilialEmpresa.SuperSimples = 1 Then
                infNFeEmit.CRT = TNFeInfNFeEmitCRT.Item1
            Else
                infNFeEmit.CRT = TNFeInfNFeEmitCRT.Item3
            End If

            'se for uma NF interna de entrada de material importado - codigo 119
            If objNFiscal.TipoNFiscal = 119 Then

                resCRFatConfig = gobjApp.dbDadosNfe.ExecuteQuery(Of CRFatConfig) _
                ("SELECT * FROM CRFatConfig WHERE Codigo = 'NF_IMPORTACAO_TRB_FLAG02'")

                For Each objCRFatConfig In resCRFatConfig
                    If objCRFatConfig.Conteudo = "1" Then
                        infNFeEmit.CRT = TNFeInfNFeEmitCRT.Item3
                    End If
                    Exit For
                Next

            End If

            Monta_NFiscal_Xml4 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml4 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml3(ByVal infNFeIde As TNFeInfNFeIde) As Long
        'preenche dados de nfe referenciada (original)

        Dim resNFeFedProtNFe As IEnumerable(Of NFeFedProtNFe)
        Dim objNFeFedProtNFe As NFeFedProtNFe
        Dim resultsNFOrig As IEnumerable(Of NFiscal)
        Dim resFilialEmpresaOrig As IEnumerable(Of FiliaisEmpresa)
        Dim resEnderecoOrig As IEnumerable(Of Endereco)
        Dim resEstadoOrig As IEnumerable(Of Estado)
        Dim objEstadoOrig As Estado
        Dim objEnderecoOrig As Endereco
        Dim objFilialEmpresaOrig As New FiliaisEmpresa
        Dim objNFiscalOrig As NFiscal
        Dim lEnderecoOrig As Long
        Dim colNumIntNF As New Collection, iIndice As Integer
        Dim resItemNF As IEnumerable(Of ItensNFiscal), bAchou As Boolean
        Dim objItemNF As ItensNFiscal, lNumIntOrig As Long
        Dim resItemNFOrig As IEnumerable(Of ItensNFiscal)
        Dim objItemNFOrig As ItensNFiscal

        Try

            gobjApp.sErro = "7"
            gobjApp.sMsg1 = "vai acessar nota original"

            colNumIntNF = New Collection

            If objNFiscal.NumNFPOrig <> 0 Then

                Dim objNFRef(1) As TNFeInfNFeIdeNFref

                infNFeIde.NFref() = objNFRef

                Dim objInstNFRef As TNFeInfNFeIdeNFref = New TNFeInfNFeIdeNFref

                infNFeIde.NFref(0) = objInstNFRef

                Dim objrefNFP As TNFeInfNFeIdeNFrefRefNFP = New TNFeInfNFeIdeNFrefRefNFP

                With objrefNFP
                    .nNF = CStr(objNFiscal.NumNFPOrig)
                    .serie = objNFiscal.SerieNFPOrig
                    .mod = TNFeInfNFeIdeNFrefRefNFPMod.Item04
                End With

                infNFeIde.NFref(0).Item = objrefNFP
                infNFeIde.NFref(0).ItemElementName = ItemChoiceType3.refNFP

            Else

                If objNFiscal.NumIntNotaOriginal <> 0 Then colNumIntNF.Add(objNFiscal.NumIntNotaOriginal)

                'lNumIntNFiscalParam
                resItemNF = gobjApp.dbDadosNfe.ExecuteQuery(Of ItensNFiscal) _
                ("SELECT * FROM ItensNFiscal WHERE  NumIntNF = {0} AND NumIntDocOrig <> 0 ORDER BY Item", objNFiscal.NumIntDoc)

                For Each objItemNF In resItemNF

                    resItemNFOrig = gobjApp.dbDadosNfe.ExecuteQuery(Of ItensNFiscal) _
                    ("SELECT * FROM ItensNFiscal WHERE  NumIntDoc = {0} ORDER BY Item", objItemNF.NumIntDocOrig)

                    For Each objItemNFOrig In resItemNFOrig

                        bAchou = False
                        For Each lNumIntOrig In colNumIntNF
                            If lNumIntOrig = objItemNFOrig.NumIntNF Then
                                bAchou = True
                                Exit For
                            End If
                        Next
                        If Not bAchou Then
                            colNumIntNF.Add(objItemNFOrig.NumIntNF)
                        End If

                        Exit For
                    Next
                Next

            End If

            'If objNFiscal.NumIntNotaOriginal <> 0 Or objNFiscal.NumNFPOrig <> 0 Then
            If colNumIntNF.Count <> 0 Then

                Dim objNFRef(colNumIntNF.Count) As TNFeInfNFeIdeNFref

                infNFeIde.NFref() = objNFRef

                iIndice = 0
                For Each lNumIntOrig In colNumIntNF

                    iIndice = iIndice + 1

                    Dim objInstNFRef As TNFeInfNFeIdeNFref = New TNFeInfNFeIdeNFref

                    infNFeIde.NFref(iIndice) = objInstNFRef

                    resNFeFedProtNFe = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedProtNFe) _
                        ("SELECT * FROM NFeFedProtNFE WHERE NumIntNF = {0} AND (cStat = '100' Or cStat = '150') ORDER BY Data DESC, Hora Desc ", lNumIntOrig)

                    objNFeFedProtNFe = resNFeFedProtNFe(0)

                    If Not objNFeFedProtNFe Is Nothing Then

                        infNFeIde.NFref(iIndice).Item = objNFeFedProtNFe.chNFe
                        infNFeIde.NFref(iIndice).ItemElementName = ItemChoiceType3.refNFe

                    Else
                        resultsNFOrig = gobjApp.dbDadosNfe.ExecuteQuery(Of NFiscal) _
                        ("SELECT * FROM NFiscal WHERE NumIntDoc = {0} ", lNumIntOrig)

                        If resultsNFOrig.Count = 0 Then Throw New System.Exception("Nota Fiscal Original não encontrada.")

                        resultsNFOrig = gobjApp.dbDadosNfe.ExecuteQuery(Of NFiscal) _
                        ("SELECT * FROM NFiscal WHERE NumIntDoc = {0} ", lNumIntOrig)

                        objNFiscalOrig = resultsNFOrig(0)

                        Dim objNRefRefNF As TNFeInfNFeIdeNFrefRefNF = New TNFeInfNFeIdeNFrefRefNF
                        Dim objNRefRefNFe As TNFeInfNFeIdeNFref = New TNFeInfNFeIdeNFref

                        If Len(objNFiscalOrig.ChvNFe) > 0 Then

                            infNFeIde.NFref(iIndice).Item = objNFiscalOrig.ChvNFe
                            infNFeIde.NFref(iIndice).ItemElementName = ItemChoiceType3.refNFe

                        Else

                            infNFeIde.NFref(iIndice).Item = objNRefRefNF
                            infNFeIde.NFref(iIndice).ItemElementName = ItemChoiceType3.refNF

                            resFilialEmpresaOrig = gobjApp.dbDadosNfe.ExecuteQuery(Of FiliaisEmpresa) _
                            ("SELECT * FROM FiliaisEmpresa WHERE FilialEmpresa = {0} ", objNFiscalOrig.FilialEmpresa)

                            For Each objFilialEmpresaOrig In resFilialEmpresaOrig
                                lEnderecoOrig = objFilialEmpresaOrig.Endereco
                                Exit For
                            Next

                            resEnderecoOrig = gobjApp.dbDadosNfe.ExecuteQuery(Of Endereco) _
                            ("SELECT * FROM Enderecos WHERE Codigo = {0}", lEnderecoOrig)

                            objEnderecoOrig = resEnderecoOrig(0)

                            resEstadoOrig = gobjApp.dbDadosNfe.ExecuteQuery(Of Estado) _
                                ("SELECT * FROM Estados WHERE Sigla = {0}", objEnderecoOrig.SiglaEstado)

                            objEstadoOrig = resEstadoOrig(0)
                            objNRefRefNF.cUF = GetCode(Of TCodUfIBGE)(CStr(objEstadoOrig.CodIBGE))

                            objNRefRefNF.AAMM = Format(objNFiscalOrig.DataEmissao, "yyMM")
                            objNRefRefNF.CNPJ = objFilialEmpresaOrig.CGC
                            objNRefRefNF.mod = TNFeInfNFeIdeNFrefRefNFMod.Item01
                            objNRefRefNF.serie = Serie_Sem_E(objNFiscalOrig.Serie)
                            objNRefRefNF.nNF = objNFiscalOrig.NumNotaFiscal

                        End If

                    End If

                Next

            Else

                If objNFiscal.Complementar = 1 Then
                    Throw New System.Exception("Trata-se de uma nota fiscal de complemento e a nota fiscal original não foi informada.")
                End If

            End If

            Monta_NFiscal_Xml3 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml3 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml2(ByVal infNFeIde As TNFeInfNFeIde) As Long
        'preenche identificacao da nf

        Dim resNFeFedScan As IEnumerable(Of NFeFedScan)
        Dim iAchou As Integer
        Dim resTipoDocInfo As IEnumerable(Of TiposDocInfo)

        Try
            If gobjApp.iFilialEmpresa <> objNFiscal.FilialEmpresa Then Throw New System.Exception("A filialempresa da chamada do envio de nfe é diferente da filialempresa da nfe")

            infNFeIde.cUF = GetCode(Of TCodUfIBGE)(CStr(gobjApp.objEstado.CodIBGE))

            If gobjApp.iDebug = 1 Then MsgBox("6")

            gobjApp.sErro = "6"
            gobjApp.sMsg1 = "vai acessar Cidades, Paises"

            If gobjApp.iNFeAmbiente = NFE_AMBIENTE_HOMOLOGACAO Then
                infNFeIde.tpAmb = TAmb.Item2
            Else
                infNFeIde.tpAmb = TAmb.Item1
            End If

            'isto foi feito pois a consulta de chave de NF usa esta rotina e preenche o lcNF que deve ser mantido nestes casos
            If Len(sChaveNFe) > 0 Then
                infNFeIde.cNF = Mid(sChaveNFe, 36, 8)
            Else
                Randomize()
                infNFeIde.cNF = Format(Rnd() * 100000000, "00000000")
            End If

            infNFeIde.natOp = DesacentuaTexto(objNFiscal.DescrNF)
            'infNFeIde.indPag = TNFeInfNFeIdeIndPag.Item2

            If sModelo <> "NFCe" Then
                infNFeIde.mod = TMod.Item55
            Else
                infNFeIde.mod = TMod.Item65
            End If

            infNFeIde.serie = Serie_Sem_E(objNFiscal.Serie)
            infNFeIde.nNF = objNFiscal.NumNotaFiscal

            If gobjApp.iScan = -1 Then
                infNFeIde.tpEmis = TNFeInfNFeIdeTpEmis.Item1
            Else
                Select Case gobjApp.sSistemaContingencia
                    Case "SCAN"
                        infNFeIde.tpEmis = TNFeInfNFeIdeTpEmis.Item3
                    Case "SVC-AN"
                        infNFeIde.tpEmis = TNFeInfNFeIdeTpEmis.Item6
                    Case "SVC-RS"
                        infNFeIde.tpEmis = TNFeInfNFeIdeTpEmis.Item7
                End Select
            End If

            If sModelo <> "NFCe" Then
                infNFeIde.tpImp = TNFeInfNFeIdeTpImp.Item1
            Else
                infNFeIde.tpImp = TNFeInfNFeIdeTpImp.Item4
            End If
            infNFeIde.dhEmi = DataHoraParaUTC(objNFiscal.DataEmissao, objNFiscal.HoraEmissao)
            sdhEmi = infNFeIde.dhEmi

            If gobjApp.iDebug = 1 Then MsgBox("7")

            gobjApp.sErro = "9"
            gobjApp.sMsg1 = "preenche a data/hora de entrada/saida"

            'se for nota de entrada
            If objNFiscal.Tipo = 1 Then
                If objNFiscal.DataEntrada <> #9/7/1822# And objNFiscal.SemDataSaida <> 1 Then infNFeIde.dhSaiEnt = DataHoraParaUTC(objNFiscal.DataEntrada, IIf(objNFiscal.DataEntrada = objNFiscal.DataEmissao And objNFiscal.HoraEntrada >= objNFiscal.HoraEmissao, objNFiscal.HoraEntrada, objNFiscal.HoraEmissao))
                infNFeIde.tpNF = TNFeInfNFeIdeTpNF.Item0

            Else
                If sModelo <> "NFCe" And objNFiscal.DataSaida <> #9/7/1822# And objNFiscal.SemDataSaida <> 1 Then infNFeIde.dhSaiEnt = DataHoraParaUTC(objNFiscal.DataSaida, IIf(objNFiscal.DataSaida = objNFiscal.DataEmissao And objNFiscal.HoraSaida >= objNFiscal.HoraEmissao, objNFiscal.HoraSaida, objNFiscal.HoraEmissao))
                infNFeIde.tpNF = TNFeInfNFeIdeTpNF.Item1

            End If

            infNFeIde.idDest = GetCode(Of TNFeInfNFeIdeIdDest)(CStr(objTributacaoDoc.Destino))

            infNFeIde.finNFe = GetCode(Of TFinNFe)(CStr(objTributacaoDoc.FinalidadeNFe))

            infNFeIde.indFinal = GetCode(Of TNFeInfNFeIdeIndFinal)(CStr(objTributacaoDoc.IndConsumidorFinal))

            If sModelo = "NFCe" And objNFiscal.FreteRespons <> 4 Then
                infNFeIde.indPres = TNFeInfNFeIdeIndPres.Item4
            Else
                infNFeIde.indPres = GetCode(Of TNFeInfNFeIdeIndPres)(CStr(objTributacaoDoc.IndPresenca))
            End If

            infNFeIde.procEmi = TProcEmi.Item0
            infNFeIde.verProc = "Corporator"

            'se foi emitido em regime de contingencia (SCAN)
            If gobjApp.iScan = 1 Then

                'procura o registro que corresponde ao periodo de SCAN
                resNFeFedScan = gobjApp.dbDadosNfe.ExecuteQuery(Of NFeFedScan) _
                ("SELECT * FROM NFeFedScan WHERE FilialEmpresa = {0} AND DataEntrada <= {1} AND (DataSaida >= {2} OR DataSaida = {3})", gobjApp.iFilialEmpresa, objNFiscal.DataEmissao, objNFiscal.DataEmissao, DATA_NULA)

                iAchou = 0

                For Each objNFeFedScan In resNFeFedScan
                    iAchou = 1
                    infNFeIde.dhCont = DataHoraParaUTC(objNFeFedScan.DataEntrada, objNFeFedScan.HoraEntrada) 'Format(objNFeFedScan.DataEntrada, "yyyy-MM-dd") & "T" & Format(Date.FromOADate(objNFeFedScan.HoraEntrada), "hh:mm:ss")
                    infNFeIde.xJust = DesacentuaTexto(Trim(objNFeFedScan.Justificativa))
                    Exit For
                Next

                If iAchou = 0 Then
                    Throw New System.Exception("Trata-se de uma nota emitida em contingÃªncia a ser enviada ao " & gobjApp.sSistemaContingencia & " e nÃ£o foi encontrado registro em NFeFedScan.")
                End If

            End If

            Monta_NFiscal_Xml2 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml2 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Private Function Monta_NFiscal_Xml1() As Long
        'obtem e valida serie vs scan

        Try

            Dim sSerie As String

            If gobjApp.iDebug = 1 Then MsgBox("3.01 Serie = " & objNFiscal.Serie)

            gobjApp.sErro = "3.01"
            gobjApp.sMsg1 = "vai fazer Serie_Sem_E"

            sSerie = Serie_Sem_E(objNFiscal.Serie)

            If gobjApp.iDebug = 1 Then MsgBox("3.02 Serie = " & sSerie)

            gobjApp.sErro = "3.02"
            gobjApp.sMsg1 = "vai testar a serie para ver se é scan"

            'Em contigência vai seguir a série normal
            'Se a NFe já fi enviado em ambiente normal e não teve retorno não pode tentar enviar em contigência, tem que gerar uma nova NFe
            'Após o ambiente normal voltar o retorno dele tem que ser tratado e se a NFe foi autorizada será necessário cancelá-la
            'Se não foi a numeração tem que ser inutilizada

            If CInt(sSerie) >= 900 And CInt(sSerie) <= 999 Then
                If gobjApp.iScan = -1 Then
                    Throw New System.Exception("Há uma nota fiscal com serie entre 900 e 999 sendo processada e está misturada com outras de série abaixo.")
                Else
                    gobjApp.iScan = 1
                End If

            Else

                If gobjApp.sSistemaContingencia = "" Then
                    If gobjApp.iScan = 1 Then
                        Throw New System.Exception("Há uma nota fiscal com serie abaixo de 900 sendo processada e está misturada com outras de série entre 900 e 999.")
                    Else
                        gobjApp.iScan = -1
                    End If
                Else
                    gobjApp.iScan = 1 'Está em contigência, não necessariamente pelo SCAN
                End If

            End If

            If CInt(sSerie) >= 890 And CInt(sSerie) <= 899 Then
                Throw New System.Exception("Série 890-899 de uso exclusivo para emissão de NF-e avulsa pelo contribuinte com seu certificado digital, através do site do Fisco.")
            End If

            If gobjApp.iDebug = 1 Then MsgBox("3.1")

            gobjApp.sErro = "3.1"
            gobjApp.sMsg1 = "vai fazer INSERT NFeFedLoteLog"

            Call gobjApp.GravarLog("Iniciando o processamento da Nota Fiscal = " & objNFiscal.NumNotaFiscal & " Série = " & Serie_Sem_E(objNFiscal.Serie), lLote, lNumIntNF)

            If gobjApp.iDebug = 1 Then MsgBox("4")

            Monta_NFiscal_Xml1 = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml1 = 1

            Call gobjApp.GravarLog(ex.Message, 0, 0)

        End Try

    End Function

    Public Function Monta_NFiscal_Xml(ByRef a5 As TNFe, ByVal objNF As NFeNFiscal, ByRef sSerie As String,
                                   ByVal lLote1 As Long,
                                   ByVal colNFiscal As Collection, ByRef sChaveNFe1 As String,
                                   ByRef XMLStringNFes As String) As Long

        Dim resTributacaoDoc As IEnumerable(Of TributacaoDoc)

        Dim XMLStream As MemoryStream = New MemoryStream(10000)
        Dim XMLString As String
        Dim iPos As Integer
        Dim iPos2 As Integer
        Dim iPos3 As String
        Dim XMLStreamDados As MemoryStream = New MemoryStream(10000)
        Dim sArquivo As String
        Dim lErro As Long

        Try

            scDest = ""
            sdhEmi = ""
            svNF = ""
            svICMS = ""
            sidNFCECSC = ""
            sNFCECSC = ""

            'armazena variaveis globais à classe
            objNFiscal = objNF
            lLote = lLote1
            lNumIntNF = objNFiscal.NumIntDoc
            sChaveNFe = sChaveNFe1

            If objNF.ModDocFisE = 35 Then
                sModelo = "NFCe"
            Else
                sModelo = "NFe"
            End If

            gobjApp.gsModelo = sModelo

            colNFiscal.Add(objNFiscal) 'colecao de nfes enviadas

            'obtem e valida serie vs scan
            lErro = Monta_NFiscal_Xml1()
            If lErro <> SUCESSO Then Throw New System.Exception("")

            a5 = New TNFe

            Dim infNFe As TNFeInfNFe = New TNFeInfNFe
            a5.infNFe = infNFe

            a5.infNFe.versao = NFE_VERSAO_XML

            Dim infNFeIde As TNFeInfNFeIde = New TNFeInfNFeIde
            a5.infNFe.ide = infNFeIde

            resTributacaoDoc = gobjApp.dbDadosNfe.ExecuteQuery(Of TributacaoDoc) _
            ("SELECT *  FROM TributacaoDoc WHERE TipoDoc = {0} AND NumIntDoc = {1}", TIPODOC_TRIB_NF, objNFiscal.NumIntDoc)

            objTributacaoDoc = resTributacaoDoc(0)

            'preenche identificacao da nf
            lErro = Monta_NFiscal_Xml2(infNFeIde)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            'preenche dados de nfe referenciada (original)
            lErro = Monta_NFiscal_Xml3(infNFeIde)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            If gobjApp.iDebug = 1 Then MsgBox("8")

            Dim infNFeEmit As TNFeInfNFeEmit = New TNFeInfNFeEmit
            a5.infNFe.emit = infNFeEmit

            'preenche dados do emitente
            lErro = Monta_NFiscal_Xml4(infNFeEmit)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            a5.infNFe.ide.cMunFG = infNFeEmit.enderEmit.cMun
            cMunFG = a5.infNFe.ide.cMunFG

            Dim infNFeDest As TNFeInfNFeDest = New TNFeInfNFeDest
            a5.infNFe.dest = infNFeDest

            'preenche dados do destinatrio
            lErro = Monta_NFiscal_Xml5(infNFeDest)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            'se referencia uma nota de produtor rural
            If objNFiscal.NumNFPOrig <> 0 Then

                Dim objrefNFP As TNFeInfNFeIdeNFrefRefNFP = infNFeIde.NFref(0).Item

                With objrefNFP
                    .AAMM = Format(objNFiscal.DataEmissao, "yyMM")
                    .cUF = GetCode(Of TCodUfIBGE)(Left(infNFeDest.enderDest.cMun, 2))
                    If Len(Trim(infNFeDest.IE)) <> 0 Then .IE = infNFeDest.IE
                    .Item = infNFeDest.Item
                    .ItemElementName = infNFeDest.ItemElementName
                End With

            End If

            'preenche local de retirada
            lErro = Monta_NFiscal_Xml6(infNFe)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            'preenche local de entrega
            lErro = Monta_NFiscal_Xml7(infNFe)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            If gobjApp.iDebug = 1 Then MsgBox("16")
            gobjApp.sErro = "16"
            gobjApp.sMsg1 = "vai acessar a tabela TributacaoDoc"

            'preenche dados dos itens da nf
            lErro = Monta_NFiscal_Xml8(infNFe, objNFiscal)
            If lErro <> SUCESSO Then Throw New System.Exception("")


            Dim infNFeTotal As TNFeInfNFeTotal = New TNFeInfNFeTotal
            a5.infNFe.total = infNFeTotal

            'preenche totais da nf
            lErro = Monta_NFiscal_Xml20(infNFeTotal)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            'preenche dados da transportadora
            Dim infNFeTransp As TNFeInfNFeTransp = New TNFeInfNFeTransp
            a5.infNFe.transp = infNFeTransp

            lErro = Monta_NFiscal_Xml21(infNFeTransp, a5)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            Dim infRespTec = New TInfRespTec
            With infRespTec
                .CNPJ = "73841488000153"
                .xContato = "Jones Abramoff"
                .email = "jones@forprint.com.br"
                .fone = "2122247445"
            End With
            a5.infNFe.infRespTec = infRespTec

            If sModelo <> "NFCe" Then

                'preenche dados de cobrança/pagamento
                Dim infNFeCobr As New TNFeInfNFeCobr
                '                a5.infNFe.cobr = infNFeCobr

                lErro = Monta_NFiscal_Xml22(infNFeCobr, a5)
                If lErro <> SUCESSO Then Throw New System.Exception("")

            Else

                Dim apag(0) As TNFeInfNFePagDetPag
                a5.infNFe.pag = New TNFeInfNFePag
                a5.infNFe.pag.detPag = apag
                apag(0) = New TNFeInfNFePagDetPag

                apag(0).tPag = TNFeInfNFePagDetPagTPag.Item01
                apag(0).vPag = Replace(Format(objNFiscal.ValorTotal, "fixed"), ",", ".")
                '??? cAdmC

            End If

            'preenche informacoes adicionais
            lErro = Monta_NFiscal_Xml23(a5)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            'preenche identificacao de pessoas autorizadas a acessar o xml
            lErro = Monta_NFiscal_Xml25(infNFe)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            If gobjApp.iDebug = 1 Then MsgBox("34")
            gobjApp.sErro = "34"
            gobjApp.sMsg1 = "vai montar a chave da nota"

            If Len(sChaveNFe) = 0 Then
                a5.infNFe.Id = GetXmlAttrNameFromEnumValue(Of TCodUfIBGE)(a5.infNFe.ide.cUF) & Format(objNFiscal.DataEmissao, "yyMM") & a5.infNFe.emit.Item
                a5.infNFe.Id = a5.infNFe.Id & GetXmlAttrNameFromEnumValue(Of TMod)(a5.infNFe.ide.mod) & Format(CInt(a5.infNFe.ide.serie), "000")
                a5.infNFe.Id = a5.infNFe.Id & Format(CLng(a5.infNFe.ide.nNF), "000000000") & GetXmlAttrNameFromEnumValue(Of TNFeInfNFeIdeTpEmis)(a5.infNFe.ide.tpEmis)
                a5.infNFe.Id = a5.infNFe.Id & Format(CLng(a5.infNFe.ide.cNF), "00000000")
            Else
                a5.infNFe.Id = sChaveNFe
            End If

            If gobjApp.iDebug = 1 Then MsgBox("35")
            gobjApp.sErro = "35"
            gobjApp.sMsg1 = "vai calcular o DV da chave"

            Dim iDigito As Integer

            If Len(sChaveNFe) = 0 Then
                CalculaDV_Modulo11(a5.infNFe.Id, iDigito)
            End If

            If gobjApp.iDebug = 1 Then MsgBox("36")

            gobjApp.sErro = "36"
            gobjApp.sMsg1 = "vai serializar os dados da nota"


            If Len(sChaveNFe) = 0 Then
                a5.infNFe.Id = "NFe" & a5.infNFe.Id & iDigito
                a5.infNFe.ide.cDV = iDigito
            Else
                a5.infNFe.Id = "NFe" & a5.infNFe.Id
                a5.infNFe.ide.cDV = Left(sChaveNFe, 1)
            End If

            If sModelo = "NFCe" Then

                a5.infNFeSupl = New TNFeInfNFeSupl
                a5.infNFeSupl.qrCode = QRCODE_PROVISORIO

            End If

            Dim AD As AssinaturaDigital = New AssinaturaDigital

            Dim mySerializer As New XmlSerializer(GetType(TNFe))

            XMLStream = New MemoryStream(10000)

            mySerializer.Serialize(XMLStream, a5)

            Dim xm As Byte()
            xm = XMLStream.ToArray

            XMLString = System.Text.Encoding.UTF8.GetString(xm)

            XMLString = Replace(XMLString, "<ICMS100>", "<ICMS40>")
            XMLString = Replace(XMLString, "</ICMS100>", "</ICMS40>")

            '***********************************

            'Remove o motivo da desoneração direto do xml quando ele não foi informado no Corporator
            'XMLString = Replace(XMLString, "<motDesICMS>1</motDesICMS>", "")


            iPos = InStr(XMLString, "xmlns:xsi")
            iPos2 = InStr(XMLString, """>")

            XMLString = Mid(XMLString, 1, iPos - 1) & Mid(XMLString, iPos2 + 1)



            'retirado em 31/03/2010 pois estava dando erro no xml
            iPos3 = InStr(XMLString, "<NFe >")

            XMLString = Mid(XMLString, 1, iPos3 + 4) & "xmlns = ""http://www.portalfiscal.inf.br/nfe""" & Mid(XMLString, iPos3 + 5)

            '****************************************


            iPos = InStr(XMLString, "<infNFe")

            If iPos <> 0 Then

                Dim iPos1 As Integer

                iPos1 = InStr(Mid(XMLString, iPos), "xmlns=""http://www.portalfiscal.inf.br/nfe""")

                If iPos1 <> 0 Then

                    XMLString = Mid(XMLString, 1, iPos + iPos1 - 2) & Mid(XMLString, iPos + iPos1 + 41)

                End If

            End If

            If gobjApp.iDebug = 1 Then MsgBox("37")
            gobjApp.sErro = "37"
            gobjApp.sMsg1 = "vai assinar a nota"


            lErro = AD.Assinar(XMLString, "infNFe", gobjApp.cert, gobjApp.iDebug)
            If lErro <> SUCESSO Then Throw New System.Exception("")

            Dim xMlD As XmlDocument

            xMlD = AD.XMLDocAssinado()

            Dim xString As String
            xString = AD.XMLStringAssinado

            If sModelo = "NFCe" And InStr(xString, QRCODE_PROVISORIO) <> 0 Then

                'calcular o qrcode e substituir
                Dim sDigVal As String = "", sNFCeQRCode As String = ""
                Dim iPos1Aux As Integer, iPos2Aux As Integer
                iPos1Aux = InStr(xString, "<DigestValue>") + Len("<DigestValue>")
                iPos2Aux = InStr(xString, "</DigestValue>")
                If iPos2Aux > iPos1Aux And iPos1Aux <> 0 Then

                    sDigVal = Mid(xString, iPos1Aux, iPos2Aux - iPos1Aux)
                    sNFCeQRCode = NFCE_Gera_QRCode(Mid(a5.infNFe.Id, 4), "100", GetXmlAttrNameFromEnumValue(Of TAmb)(a5.infNFe.ide.tpAmb), scDest, sdhEmi, svNF, svICMS, sDigVal, sidNFCECSC, sNFCECSC)
                    sNFCeQRCode = "<![CDATA[" & sNFCeQRCode & "]]>"
                    xString = Replace(xString, QRCODE_PROVISORIO, sNFCeQRCode)

                End If

            End If

            XMLStringNFes = XMLStringNFes & Mid(xString, 22) & " "

            '****************  salva o arquivo 

            XMLStreamDados = New MemoryStream(10000)

            Dim xDados1 As Byte()

            xDados1 = System.Text.Encoding.UTF8.GetBytes(Mid(xString, 22))

            XMLStreamDados.Write(xDados1, 0, xDados1.Length)

            Dim DocDados1 As XmlDocument = New XmlDocument

            XMLStreamDados.Position = 0
            DocDados1.Load(XMLStreamDados)
            sArquivo = gobjApp.sDirXml & Mid(a5.infNFe.Id, 4) & "-pre.xml"

            Dim writer As New XmlTextWriter(sArquivo, Nothing)

            writer.Formatting = Formatting.None
            DocDados1.WriteTo(writer)

            writer.Close()

            Monta_NFiscal_Xml = SUCESSO

        Catch ex As Exception

            Monta_NFiscal_Xml = 1

            Dim sMsg2 As String

            If ex.InnerException Is Nothing Then
                sMsg2 = ""
            Else
                sMsg2 = " - " & ex.InnerException.Message
            End If

            Call gobjApp.GravarLog("ERRO - " & ex.Message & sMsg2 & IIf(objNFiscal.NumNotaFiscal <> 0, "Serie = " & sSerie & " Nota Fiscal = " & objNFiscal.NumNotaFiscal, ""), lLote, lNumIntNF)

        End Try

    End Function

    Function PIS_CST(ByRef iCST As Integer, ByVal objTributacaoDocItem As TributacaoDocItem) As Long
        If objTributacaoDocItem.PISCredito > 0 Then
            iCST = 1
        Else
            iCST = 4
        End If
        PIS_CST = SUCESSO
    End Function

    Function COFINS_CST(ByRef iCST As Integer, ByVal objTributacaoDocItem As TributacaoDocItem) As Long
        If objTributacaoDocItem.COFINSCredito > 0 Then
            iCST = 1
        Else
            iCST = 4
        End If
        COFINS_CST = SUCESSO
    End Function

    Public Shared Sub Salva_Arquivo(ByVal DocDados1 As XmlDocument, ByVal XMLString4 As String)

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

    Public Sub New()
        bComISSQN = False
        scDest = ""
        sdhEmi = ""
        svNF = ""
        svICMS = ""
        sidNFCECSC = ""
        sNFCECSC = ""
    End Sub

    Private Function Produto_Trata_EAN(ByVal objProd As Produto) As Long

        Dim lErro As Long
        Dim iValidacaoEAN As Integer
        Dim resCRFatConfig As IEnumerable(Of CRFatConfig)
        Dim objCRFatConfig As CRFatConfig
        Dim dtValidacaoEAN As Date

        Try

            resCRFatConfig = gobjApp.dbDadosNfe.ExecuteQuery(Of CRFatConfig) _
            ("SELECT * FROM CRFatConfig WHERE Codigo = 'NFE_VALIDACAO_EAN'")

            For Each objCRFatConfig In resCRFatConfig
                iValidacaoEAN = CInt(objCRFatConfig.Conteudo)
                Exit For
            Next

            resCRFatConfig = gobjApp.dbDadosNfe.ExecuteQuery(Of CRFatConfig) _
            ("SELECT * FROM CRFatConfig WHERE Codigo = 'NFE_VALIDACAO_EAN_A_PARTIR_DE'")

            For Each objCRFatConfig In resCRFatConfig
                dtValidacaoEAN = CDate(objCRFatConfig.Conteudo)
                Exit For
            Next

            If dtValidacaoEAN <= Now.Date Then
                If Left(objProd.CodigoBarras, 1) = "2" Then objProd.CodigoBarras = "" 'Não enviar código de Barras de balança
                If Left(objProd.CodigoBarrasTrib, 1) = "2" Then objProd.CodigoBarrasTrib = "" 'Não enviar código de Barras de balança
            End If

            Select Case iValidacaoEAN

                Case 1 'Valida 

                    If dtValidacaoEAN <= Now.Date Then

                        lErro = Valida_EAN(objProd.CodigoBarras)
                        If lErro <> SUCESSO Then objProd.CodigoBarras = ""

                        lErro = Valida_EAN(objProd.CodigoBarrasTrib)
                        If lErro <> SUCESSO Then objProd.CodigoBarrasTrib = ""

                        'Para produtos que não possuem código de barras com GTIN , deve ser informado o literal “SEM GTIN”;
                        If objProd.CodigoBarras = "" Then objProd.CodigoBarras = "SEM GTIN"
                        If objProd.CodigoBarrasTrib = "" Then objProd.CodigoBarrasTrib = objProd.CodigoBarras

                    End If

                Case 2 'Não Envia nada
                    objProd.CodigoBarras = ""
                    objProd.CodigoBarrasTrib = ""

                Case Else
                    'Vai o que tiver

            End Select

            Produto_Trata_EAN = SUCESSO

        Catch ex As Exception
            Produto_Trata_EAN = 1
        End Try
    End Function

    Private Function Valida_EAN(ByVal sEAN As String) As Long
        Dim lErro As Long = 0
        Try
            Dim intTotalSoma As Integer
            Dim intDv As Integer
            Dim I As Integer
            Dim iNumChar As Integer
            Dim iMult As Integer

            iNumChar = Len(Trim(sEAN))
            sEAN = Trim(sEAN)
            iMult = 3

            If iNumChar <> 0 Then

                If Not (iNumChar = 8 Or iNumChar = 12 Or iNumChar = 13 Or iNumChar = 14) Or Not IsNumeric(sEAN) Then Error 6015

                '0,8,12,13, 14
                'Preencher com o código GTIN-8, GTIN-12, GTIN-13
                'ou GTIN-14 (antigos códigos EAN, UPC e DUN-14).
                'Para produtos que não possuem código de
                'barras com GTIN, deve ser informado o literal
                '“SEM GTIN”;
                'Nos demais casos, preencher com GTIN contido na
                'embalagem com código de barras

                intTotalSoma = 0
                intDv = 0

                For I = iNumChar - 1 To 1 Step -1
                    intTotalSoma = intTotalSoma + CInt(Mid(sEAN, I, 1)) * iMult
                    If iMult = 3 Then
                        iMult = 1
                    Else
                        iMult = 3
                    End If
                Next

                Do While intTotalSoma Mod 10 <> 0
                    intDv = intDv + 1
                    intTotalSoma = intTotalSoma + 1
                Loop

                If Right(sEAN, 1) <> CStr(intDv) Then Error 6016

            End If

            Valida_EAN = SUCESSO

        Catch ex As Exception
            Valida_EAN = Err.Number
        End Try

    End Function
End Class