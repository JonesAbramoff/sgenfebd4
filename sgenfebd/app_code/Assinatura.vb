Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml
Imports System.Xml.Schema

Public Class AssinaturaDigital

    Public Function Assinar(ByVal XMLString As String, ByVal RefUri As String, ByVal X509Cert As X509Certificate2, ByVal iDebug As Integer) As Integer
        '     Entradas:
        '         XMLString: string XML a ser assinada
        '         RefUri   : Referência da URI a ser assinada (Ex. infNFe
        '         X509Cert : certificado digital a ser utilizado na assinatura digital
        ' 
        '     Retornos:
        '         Assinar : 0 - Assinatura realizada com sucesso
        '                   1 - Erro: Problema ao acessar o certificado digital - %exceção%
        '                   2 - Problemas no certificado digital
        '                  3 - XML mal formado + exceção
        '                   4 - A tag de assinatura %RefUri% inexiste
        '                   5 - A tag de assinatura %RefUri% não é unica
        '                   6 - Erro Ao assinar o documento - ID deve ser string %RefUri(Atributo)%
        '                   7 - Erro: Ao assinar o documento - %exceção%
        ' 
        '        XMLStringAssinado : string XML assinada
        ' 
        '         XMLDocAssinado    : XMLDocument do XML assinado
        '

        Dim resultado As Integer = 0
        msgResultado = "Assinatura realizada com sucesso"
        Try
            '   certificado para ser utilizado na assinatura
            '
            Dim _xnome As String = ""
            If (Not X509Cert Is Nothing) Then
                _xnome = X509Cert.Subject.ToString()
            End If

            Dim _X509Cert As X509Certificate2 = New X509Certificate2()
            Dim store As X509Store = New X509Store("MY", StoreLocation.CurrentUser)

            store.Open(OpenFlags.ReadOnly Or OpenFlags.OpenExistingOnly)
            Dim collection As X509Certificate2Collection = store.Certificates
            '(X509Certificate2Collection)
            Dim collection1 As X509Certificate2Collection = collection.Find(X509FindType.FindBySubjectDistinguishedName, _xnome, False)

            Dim iIndice As Integer, bAchou As Boolean
            bAchou = False
            For iIndice = 0 To collection1.Count - 1
                If collection1(iIndice).NotAfter > DateTime.Now And collection1(iIndice).NotBefore < DateTime.Now Then
                    bAchou = True
                    Exit For
                End If
            Next
            '(X509Certificate2Collection)
            If bAchou = False Or (collection1.Count = 0) Then
                resultado = 2
                msgResultado = "Problemas no certificado digital"
                MsgBox(msgResultado)
            Else
                ' certificado ok
                _X509Cert = collection1(iIndice)
                Dim x As String
                x = _X509Cert.GetKeyAlgorithm().ToString()
                ' Create a new XML document.
                Dim doc As XmlDocument = New XmlDocument()

                ' Format the document to ignore white spaces.
                doc.PreserveWhitespace = False

                ' Load the passed XML file using it's name.
                Try
                    doc.LoadXml(XMLString)

                    ' Verifica se a tag a ser assinada existe é única
                    Dim qtdeRefUri As Integer = doc.GetElementsByTagName(RefUri).Count

                    If (qtdeRefUri = 0) Then
                        '  a URI indicada não existe
                        resultado = 4
                        msgResultado = "A tag de assinatura " + RefUri.Trim() + " inexiste"
                        MsgBox(msgResultado)
                        ' Exsiste mais de uma tag a ser assinada
                    Else

                        If (qtdeRefUri > 1) Then
                            ' existe mais de uma URI indicada
                            resultado = 5
                            msgResultado = "A tag de assinatura " + RefUri.Trim() + " não é unica"
                            MsgBox(msgResultado)

                            '//else if (_listaNum.IndexOf(doc.GetElementsByTagName(RefUri).Item(0).Attributes.ToString().Substring(1,1))>0)
                            '//{
                            '//    resultado = 6;
                            '//    msgResultado = "Erro: Ao assinar o documento - ID deve ser string (" + doc.GetElementsByTagName(RefUri).Item(0).Attributes + ")";
                            '//}
                        Else
                            Try

                                ' Create a SignedXml object.
                                Dim SignedXml As SignedXml = New SignedXml(doc)

                                ' Add the key to the SignedXml document 



                                SignedXml.SigningKey = _X509Cert.PrivateKey

                                ' Create a reference to be signed
                                Dim reference As Reference = New Reference()
                                ' pega o uri que deve ser assinada
                                Dim _Uri As XmlAttributeCollection = doc.GetElementsByTagName(RefUri).Item(0).Attributes
                                Dim _atributo As XmlAttribute
                                For Each _atributo In _Uri
                                    If (_atributo.Name = "Id") Then
                                        reference.Uri = "#" + _atributo.InnerText
                                    End If
                                Next

                                ' Add an enveloped transformation to the reference.
                                Dim env As XmlDsigEnvelopedSignatureTransform = New XmlDsigEnvelopedSignatureTransform()
                                reference.AddTransform(env)

                                Dim c14 As XmlDsigC14NTransform = New XmlDsigC14NTransform()
                                reference.AddTransform(c14)

                                ' Add the reference to the SignedXml object.
                                SignedXml.AddReference(reference)

                                '// Create a new KeyInfo object
                                Dim keyInfo As KeyInfo = New KeyInfo()

                                '// Load the certificate into a KeyInfoX509Data object
                                '// and add it to the KeyInfo object.
                                keyInfo.AddClause(New KeyInfoX509Data(_X509Cert))

                                '// Add the KeyInfo object to the SignedXml object.
                                SignedXml.KeyInfo = keyInfo

                                SignedXml.ComputeSignature()

                                '// Get the XML representation of the signature and save
                                '// it to an XmlElement object.
                                Dim xmlDigitalSignature As XmlElement = SignedXml.GetXml()

                                '// Append the element to the XML document.
                                doc.DocumentElement.AppendChild(doc.ImportNode(xmlDigitalSignature, True))
                                XMLDoc = New XmlDocument()
                                XMLDoc.PreserveWhitespace = False
                                XMLDoc = doc

                            Catch caught As Exception
                                resultado = 7
                                msgResultado = "Erro: Ao assinar o documento - " + caught.Message
                                MsgBox(msgResultado)
                            End Try
                        End If
                    End If
                Catch caught As Exception
                    resultado = 3
                    msgResultado = "Erro: XML mal formado - " + caught.Message
                    MsgBox(msgResultado)
                End Try
            End If
        Catch caught As Exception

            resultado = 1
            msgResultado = "Erro: Problema ao acessar o certificado digital" + caught.Message
            MsgBox(msgResultado)
        End Try
        Assinar = resultado

    End Function
    '//
    '// mensagem de Retorno
    '//
    Private msgResultado As String
    Private XMLDoc As XmlDocument

    Public Function XMLDocAssinado() As XmlDocument
        XMLDocAssinado = XMLDoc
    End Function

    Public Function XMLStringAssinado() As String
        XMLStringAssinado = XMLDoc.OuterXml
    End Function

    Public Function mensagemResultado() As String
        mensagemResultado = msgResultado
    End Function

End Class

Public Class Certificado

    Public Function BuscaNome(ByVal Nome As String, ByRef certAux As X509Certificate2) As Long

        Dim _X509Cert As X509Certificate2 = New X509Certificate2()
        Dim bSelecionou As Boolean = False

        Try

            Dim store As X509Store = New X509Store("MY", StoreLocation.CurrentUser)
            store.Open(OpenFlags.OpenExistingOnly Or OpenFlags.IncludeArchived Or OpenFlags.ReadWrite)
            Dim collection As X509Certificate2Collection = store.Certificates
            Dim collection1 As X509Certificate2Collection = collection.Find(X509FindType.FindByTimeValid, DateTime.Now, False)
            Dim collection2 As X509Certificate2Collection = collection.Find(X509FindType.FindByKeyUsage, X509KeyUsageFlags.DigitalSignature, False)
            If Nome = "" Then
                Dim scollection As X509Certificate2Collection = X509Certificate2UI.SelectFromCollection(collection1, "Certificado(s) Digital(is) disponível(is)", "Selecione o Certificado Digital para uso no aplicativo", X509SelectionFlag.SingleSelection)
                If (scollection.Count = 0) Then
                    _X509Cert.Reset()
                    MsgBox("Nenhum certificado escolhido")
                Else
                    _X509Cert = scollection(0)
                    bSelecionou = True
                End If
            Else
                Dim scollection As X509Certificate2Collection = collection2.Find(X509FindType.FindBySubjectName, Nome, False)
                Dim iIndice As Integer, iPos As Integer, iQtdeValidos As Integer
                iPos = -1
                iQtdeValidos = 0
                For iIndice = 0 To scollection.Count - 1
                    If scollection(iIndice).NotAfter > DateTime.Now And scollection(iIndice).NotBefore < DateTime.Now Then
                        If iPos = -1 Then iPos = iIndice
                        iQtdeValidos = iQtdeValidos + 1
                    End If
                Next
                '(X509Certificate2Collection)
                If iQtdeValidos = 0 Then
                    MsgBox("Nenhum certificado válido foi encontrado com o nome informado: " + Nome)
                    _X509Cert.Reset()
                Else
                    If iQtdeValidos = 1 Then
                        _X509Cert = scollection(iPos)
                        bSelecionou = True
                    Else
                        Dim scollection2 As X509Certificate2Collection = X509Certificate2UI.SelectFromCollection(scollection, "Certificado(s) Digital(is) disponível(is)", "Selecione o Certificado Digital para uso no aplicativo", X509SelectionFlag.SingleSelection)
                        If (scollection2.Count = 0) Then
                            _X509Cert.Reset()
                            MsgBox("Nenhum certificado escolhido")
                        Else
                            _X509Cert = scollection2(0)
                            bSelecionou = True
                        End If
                    End If
                End If
            End If
            store.Close()

            If bSelecionou Then
                certAux = _X509Cert
                BuscaNome = SUCESSO
            Else
                BuscaNome = 1
            End If

        Catch ex As SystemException
            MsgBox(ex.Message)
            BuscaNome = 1
        End Try
    End Function

    Public Function BuscaNroSerie(ByVal NroSerie As String) As X509Certificate2
        Dim _X509Cert As X509Certificate2 = New X509Certificate2()
        Try

            Dim store As X509Store = New X509Store("My", StoreLocation.CurrentUser)
            store.Open(OpenFlags.ReadOnly Or OpenFlags.OpenExistingOnly)
            Dim collection As X509Certificate2Collection = store.Certificates
            Dim collection1 As X509Certificate2Collection = collection.Find(X509FindType.FindByTimeValid, DateTime.Now, True)
            Dim collection2 As X509Certificate2Collection = collection1.Find(X509FindType.FindByKeyUsage, X509KeyUsageFlags.DigitalSignature, True)
            If (NroSerie = "") Then
                Dim scollection As X509Certificate2Collection = X509Certificate2UI.SelectFromCollection(collection2, "Certificados Digitais", "Selecione o Certificado Digital para uso no aplicativo", X509SelectionFlag.SingleSelection)
                If (scollection.Count = 0) Then
                    _X509Cert.Reset()
                    MsgBox("Nenhum certificado válido foi encontrado com o número de série informado: " + NroSerie, "Atenção")
                Else
                    _X509Cert = scollection(0)
                End If
            Else
                Dim scollection As X509Certificate2Collection = collection2.Find(X509FindType.FindBySerialNumber, NroSerie, True)
                If (scollection.Count = 0) Then
                    _X509Cert.Reset()
                    MsgBox("Nenhum certificado válido foi encontrado com o número de série informado: " + NroSerie, "Atenção")
                Else
                    _X509Cert = scollection(0)
                End If
            End If
            store.Close()
            Return _X509Cert
        Catch ex As System.Exception
            MsgBox(ex.Message)
            Return _X509Cert
        End Try

    End Function


End Class
