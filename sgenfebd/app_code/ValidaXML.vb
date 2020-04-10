Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml
Imports System.Xml.Schema


Public Class ClassValidaXML

    Dim gsMsg As String
    Dim glErro As Long

    Public Function validaXML(ByVal _arquivo As String, ByVal _schema As String, ByVal lLote As Long, ByVal lNumIntNF As Long, ByVal iFilialEmpresa As Integer) As Long

        ' Create a new validating reader

        Dim reader As XmlValidatingReader = New XmlValidatingReader(New XmlTextReader(New StreamReader(_arquivo)))
        'Dim reader As XmlValidatingReader = New XmlValidatingReader()
        '        Dim reader1 As XmlWriter
        '       reader1.
        Dim schema(1) As System.Xml.Schema.XmlSchema

        '// Create a schema collection, add the xsd to it

        Dim schemaCollection As XmlSchemaSet = New XmlSchemaSet()
        Dim iLinha As Integer

        Try

            glErro = 0

            schemaCollection.Add("http://www.portalfiscal.inf.br/nfe", _schema)

            schemaCollection.CopyTo(schema, 0)

            '// Add the schema collection to the XmlValidatingReader

            reader.Schemas.Add(schema(0))

            '       Console.Write("Início da validação...\n")

            '    // Wire up the call back.  The ValidationEvent is fired when the
            '    // XmlValidatingReader hits an issue validating a section of the xml

            '            reader. += new ValidationEventHandler(reader_ValidationEventHandler);
            AddHandler reader.ValidationEventHandler, AddressOf reader_ValidationEventHandler

            '            // Iterate through the xml document



            '            while (reader.Read()) {}

            iLinha = 0

            While reader.Read()

                iLinha = iLinha + 1
                If Len(Trim(gsMsg)) > 0 Then

                    gobjApp.GravarLog(" Linha = " & iLinha & gsMsg, lLote, lNumIntNF)

                    gsMsg = ""

                End If

            End While


        Catch ex As Exception

            Dim sMsg As String

            If ex.InnerException Is Nothing Then
                sMsg = ""
            Else
                sMsg = " - " & ex.InnerException.Message
            End If

            Call gobjApp.GravarLog(ex.Message & sMsg, lLote, lNumIntNF)

            Call gobjApp.GravarLog("ERRO - Validação do schema.", lLote, 0)

            glErro = 1

        Finally
            validaXML = glErro
        End Try
        '          Console.WriteLine("\rFim de validação\n");
        'Console.ReadLine();
    End Function

    Sub reader_ValidationEventHandler(ByVal sender As Object, ByVal e As ValidationEventArgs)

        '            // Report back error information to the console...
        '        MessageBox.Show(e.Exception.Message)
        '        Console.WriteLine("\rLinha:{0} Coluna:{1} Erro:{2} Name:[3} Valor:{4}\r", e.Exception.LinePosition, e.Exception.LineNumber, e.Exception.Message, sender.Name, sender.Value)

        gsMsg = e.Exception.Message
        glErro = 1

    End Sub

End Class
