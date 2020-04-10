Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Xml.Schema
Imports System.Data.Odbc

Public Class Form1

#Const DEPURACAO = 0

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Timer1.Interval = 1000
        Timer1.Start()

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick


        Dim sEmpresa As String
        Dim lLote As Long
        Dim sOperacao As String

        Dim lNumIntNF As Long
        Dim sMotivo As String
        Dim iFilialEmpresa As Integer

        Dim objEnvioNFe As ClassEnvioNFe = New ClassEnvioNFe
        Dim objCancelaNFe As ClassCancelaNFe = New ClassCancelaNFe
        Dim objConsultaLoteNFe As ClassConsultaLoteNFe = New ClassConsultaLoteNFe
        Dim objInutilizaNFe As ClassInutilizaNFe = New ClassInutilizaNFe
        Dim objConsultaNFe As ClassConsultaNFe = New ClassConsultaNFe
        Dim objCCeNFe As ClassCCeNFe = New ClassCCeNFe

        Dim sSerie As String
        Dim sNumInicial As String
        Dim sNumFinal As String
        Dim iAnoParam As Integer
        Dim iScan As Integer
        Dim iDebug As Integer
        Dim sChaveNF As String

        Dim arguments As [String]() = Environment.GetCommandLineArgs()
        Dim dic As DataClasses2DataContext = New DataClasses2DataContext
        Dim odbc As OdbcConnection = New OdbcConnection
        Dim resdicGeraXML As IEnumerable(Of GeraXML)
        Dim objGeraXML As GeraXML
        Dim sChNFe As String
        Dim sSistemaContingencia As String = ""
        Dim lErro As Long

        Try

            Timer1.Stop()
            sChNFe = ""
            iDebug = 0

#If DEPURACAO = 0 Then

            sOperacao = arguments(1)
            If gobjApp.iDebug = 1 Then MsgBox("Operacao =" & sOperacao)

            sEmpresa = arguments(2)
            If gobjApp.iDebug = 1 Then MsgBox("Empresa =" & sEmpresa)

            iFilialEmpresa = CInt(arguments(3))
            If gobjApp.iDebug = 1 Then MsgBox("FilialEmpresa =" & CStr(iFilialEmpresa))

#Else

            MsgBox("VERSAO DE TESTE")

            'os valores abaixo sao setados para depuracao  
            'simulando a chamada pela aplicacao vb6.

            sOperacao = "Envio"
            sEmpresa = 1
            iFilialEmpresa = 1

#End If

            Me.UseWaitCursor = True
            Application.DoEvents()

            lErro = gobjApp.Iniciar(sEmpresa, iFilialEmpresa)
            If lErro <> SUCESSO Then Throw New System.Exception("Erro na inicialização do programa.")

            Select Case sOperacao

                Case "Envio"

#If DEPURACAO = 0 Then
                    lLote = CLng(arguments(4))
#Else
                    lLote = 666
#End If

                    If arguments.Count >= 6 Then gobjApp.sSistemaContingencia = arguments(5)

                    Contingencia.Text = gobjApp.sSistemaContingencia

                    If gobjApp.iDebug = 1 Then MsgBox("Lote =" & CStr(lLote))

                    Lote.Text = lLote

                    objEnvioNFe.Envia_Lote_NFe(sEmpresa, lLote, iFilialEmpresa)

                Case "Cancela"

#If DEPURACAO = 0 Then
                    lNumIntNF = CLng(arguments(4))
                    sMotivo = arguments(5)
#Else
                    lNumIntNF = 291996
                    sMotivo = "erro de preenchimento dos dados"
#End If

                    If gobjApp.iDebug = 1 Then MsgBox("FilialEmpresa = " & CStr(iFilialEmpresa))
                    If gobjApp.iDebug = 1 Then MsgBox("Empresa = " & sEmpresa)
                    If gobjApp.iDebug = 1 Then MsgBox("NumIntNF = " & lNumIntNF)
                    If gobjApp.iDebug = 1 Then MsgBox("Motivo = " & sMotivo)

                    If arguments.Count >= 7 Then
                        If iDebug = 1 Then MsgBox("-5")
                        iScan = CInt(arguments(6))
                        If arguments.Count >= 8 Then gobjApp.sSistemaContingencia = arguments(7)

                        Contingencia.Text = gobjApp.sSistemaContingencia

                        objCancelaNFe.Evento_Cancela_NFe(sEmpresa, lNumIntNF, sMotivo, iFilialEmpresa, sChNFe, iScan)

                        If Len(Trim(sChNFe)) > 0 Then
                            objConsultaNFe.Consulta_NFe(sEmpresa, sChNFe, iFilialEmpresa, OPERACAO_CANCELAMENTO, iScan)
                        End If

                    Else
                        If gobjApp.iDebug = 1 Then MsgBox("-6")
                        objCancelaNFe.Evento_Cancela_NFe(sEmpresa, lNumIntNF, sMotivo, iFilialEmpresa, sChNFe)

                        If Len(Trim(sChNFe)) > 0 Then
                            objConsultaNFe.Consulta_NFe(sEmpresa, sChNFe, iFilialEmpresa, OPERACAO_CANCELAMENTO)
                        End If

                    End If

                Case "Consulta"

                    If gobjApp.iDebug = 1 Then MsgBox("99")

#If DEPURACAO = 0 Then
                    lLote = CLng(arguments(4))
#Else
                    lLote = 23908
#End If

                    Lote.Text = lLote

                    If gobjApp.iDebug = 1 Then MsgBox("99.1")

                    If arguments.Count >= 6 Then
                        If iDebug = 1 Then MsgBox(arguments.Count)
                        iScan = CInt(arguments(5))
                        If arguments.Count >= 7 Then gobjApp.sSistemaContingencia = arguments(6)
                        Contingencia.Text = gobjApp.sSistemaContingencia
                        If iDebug = 1 Then MsgBox(arguments(5))
                        objConsultaLoteNFe.Consulta_Lote_NFe(sEmpresa, lLote, iFilialEmpresa, iScan)
                    Else
                        If gobjApp.iDebug = 1 Then MsgBox("99.2")
                        objConsultaLoteNFe.Consulta_Lote_NFe(sEmpresa, lLote, iFilialEmpresa)
                    End If


                Case "Inutiliza"

#If DEPURACAO = 0 Then
                    sSerie = arguments(4)
                    sNumInicial = arguments(5)
                    sNumFinal = arguments(6)
                    iAnoParam = CInt(arguments(7))
                    sMotivo = arguments(8)
#Else
                    sSerie = "1"
                    sNumInicial = 20083
                    sNumFinal = 20083
                    iAnoParam = 2015
                    sMotivo = "Falha no Sistema de emissão de notas fiscais"
#End If

                    If arguments.Count = 10 Then
                        iScan = CInt(arguments(9))

                        Contingencia.Text = gobjApp.sSistemaContingencia

                        objInutilizaNFe.Inutiliza_NFe(sEmpresa, sSerie, sNumInicial, sNumFinal, iFilialEmpresa, iAnoParam, sMotivo, iScan)
                    Else
                        objInutilizaNFe.Inutiliza_NFe(sEmpresa, sSerie, sNumInicial, sNumFinal, iFilialEmpresa, iAnoParam, sMotivo)
                    End If

                Case "ConsultaNF"

#If DEPURACAO = 0 Then
                    sChaveNF = arguments(4)
#Else
                    sChaveNF = "41150333413527001179550010000203931296810309"
#End If

                    If arguments.Count >= 6 Then
                        If iDebug = 1 Then MsgBox(arguments.Count)
                        iScan = CInt(arguments(5))
                        If arguments.Count >= 7 Then gobjApp.sSistemaContingencia = arguments(6)
                        Contingencia.Text = gobjApp.sSistemaContingencia

                        If iDebug = 1 Then MsgBox(arguments(5))
                        objConsultaNFe.Consulta_NFe(sEmpresa, sChaveNF, iFilialEmpresa, 0, iScan)
                    Else
                        If gobjApp.iDebug = 1 Then MsgBox("100.2")
                        objConsultaNFe.Consulta_NFe(sEmpresa, sChaveNF, iFilialEmpresa, 0)
                    End If

                Case "GeraXML"

                    resdicGeraXML = dic.ExecuteQuery(Of GeraXML) _
                    ("SELECT * FROM GeraXML ORDER BY chNFe")

                    For Each objGeraXML In resdicGeraXML

                        sChaveNF = objGeraXML.chNFe

                        'quando regerar o xml se os arquivos antigos estiverem presentes, vao ser renomeados para _old.xml com a indicacao do ultimo parametro = 1
                        objConsultaNFe.Consulta_NFe(sEmpresa, sChaveNF, iFilialEmpresa, 0, -1, 1)

                    Next

                Case "CartaCorrecao"

#If DEPURACAO = 0 Then
                    lLote = CLng(arguments(4))
#Else
                    lLote = 15
#End If
                    Lote.Text = lLote

                    If arguments.Count = 6 Then
                        If gobjApp.iDebug = 1 Then MsgBox("-11")
                        iScan = CInt(arguments(5))
                        Contingencia.Text = gobjApp.sSistemaContingencia

                        objCCeNFe.CCe_NFe(sEmpresa, iFilialEmpresa, lLote, iScan)

                    Else
                        If gobjApp.iDebug = 1 Then MsgBox("-12")
                        objCCeNFe.CCe_NFe(sEmpresa, iFilialEmpresa, lLote)
                    End If


                Case Else
                    MsgBox("Operação Inválida")

            End Select

        Catch ex As Exception
            Msg.Items.Add("Erro na execucao")
            Msg.Items.Add(ex.Message)

            If Not ex.InnerException Is Nothing Then
                Msg.Items.Add(ex.InnerException.Message)
                If Not ex.InnerException.InnerException Is Nothing Then
                    Msg.Items.Add(ex.InnerException.InnerException.Message)
                    If Not ex.InnerException.InnerException.InnerException Is Nothing Then
                        Msg.Items.Add(ex.InnerException.InnerException.InnerException.Message)
                    End If
                End If
            End If

        Finally

            Call gobjApp.Terminar()

            Me.UseWaitCursor = False

            If gobjApp.iDebug = 1 Then Call MsgBox(gobjApp.sErro & " " & gobjApp.sMsg1)

        End Try

    End Sub

End Class

