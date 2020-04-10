<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.Msg = New System.Windows.Forms.ListBox
        Me.Lote = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.Contingencia = New System.Windows.Forms.Label
        Me.EmHomologacao = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Msg
        '
        Me.Msg.FormattingEnabled = True
        Me.Msg.HorizontalScrollbar = True
        Me.Msg.Location = New System.Drawing.Point(19, 110)
        Me.Msg.Name = "Msg"
        Me.Msg.Size = New System.Drawing.Size(932, 277)
        Me.Msg.TabIndex = 11
        '
        'Lote
        '
        Me.Lote.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Lote.Location = New System.Drawing.Point(65, 23)
        Me.Lote.Name = "Lote"
        Me.Lote.Size = New System.Drawing.Size(58, 18)
        Me.Lote.TabIndex = 10
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(28, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(31, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Lote:"
        '
        'Timer1
        '
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(16, 68)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(935, 26)
        Me.ProgressBar1.TabIndex = 8
        '
        'Contingencia
        '
        Me.Contingencia.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Contingencia.ForeColor = System.Drawing.Color.Red
        Me.Contingencia.Location = New System.Drawing.Point(885, 18)
        Me.Contingencia.Name = "Contingencia"
        Me.Contingencia.Size = New System.Drawing.Size(66, 23)
        Me.Contingencia.TabIndex = 12
        Me.Contingencia.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'EmHomologacao
        '
        Me.EmHomologacao.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EmHomologacao.ForeColor = System.Drawing.Color.Red
        Me.EmHomologacao.Location = New System.Drawing.Point(564, 23)
        Me.EmHomologacao.Name = "EmHomologacao"
        Me.EmHomologacao.Size = New System.Drawing.Size(257, 23)
        Me.EmHomologacao.TabIndex = 13
        Me.EmHomologacao.Text = "AMBIENTE DE HOMOLOGAÇÂO (TESTE)"
        Me.EmHomologacao.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.EmHomologacao.Visible = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(973, 406)
        Me.Controls.Add(Me.EmHomologacao)
        Me.Controls.Add(Me.Contingencia)
        Me.Controls.Add(Me.Msg)
        Me.Controls.Add(Me.Lote)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Name = "Form1"
        Me.Text = "CORPORATOR - Log de Nota Fiscal Eletrônica - 4.00"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Msg As System.Windows.Forms.ListBox
    Friend WithEvents Lote As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Contingencia As System.Windows.Forms.Label
    Friend WithEvents EmHomologacao As System.Windows.Forms.Label

End Class
