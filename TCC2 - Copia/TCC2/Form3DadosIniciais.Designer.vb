<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form3DadosIniciais
    Inherits System.Windows.Forms.Form

    'Descartar substituições de formulário para limpar a lista de componentes.
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

    'Exigido pelo Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'OBSERVAÇÃO: o procedimento a seguir é exigido pelo Windows Form Designer
    'Pode ser modificado usando o Windows Form Designer.  
    'Não o modifique usando o editor de códigos.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form3DadosIniciais))
        Me.imgMomento = New System.Windows.Forms.PictureBox()
        Me.TableLayoutPanel7 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label383 = New System.Windows.Forms.Label()
        Me.Label91 = New System.Windows.Forms.Label()
        Me.Label136 = New System.Windows.Forms.Label()
        CType(Me.imgMomento, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel7.SuspendLayout()
        Me.SuspendLayout()
        '
        'imgMomento
        '
        Me.imgMomento.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.imgMomento.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.imgMomento.Image = CType(resources.GetObject("imgMomento.Image"), System.Drawing.Image)
        Me.imgMomento.InitialImage = Nothing
        Me.imgMomento.Location = New System.Drawing.Point(197, 149)
        Me.imgMomento.Name = "imgMomento"
        Me.imgMomento.Size = New System.Drawing.Size(366, 268)
        Me.imgMomento.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgMomento.TabIndex = 47
        Me.imgMomento.TabStop = False
        '
        'TableLayoutPanel7
        '
        Me.TableLayoutPanel7.ColumnCount = 1
        Me.TableLayoutPanel7.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel7.Controls.Add(Me.Label383, 0, 2)
        Me.TableLayoutPanel7.Controls.Add(Me.Label91, 0, 1)
        Me.TableLayoutPanel7.Controls.Add(Me.Label136, 0, 0)
        Me.TableLayoutPanel7.Location = New System.Drawing.Point(197, 34)
        Me.TableLayoutPanel7.Name = "TableLayoutPanel7"
        Me.TableLayoutPanel7.RowCount = 3
        Me.TableLayoutPanel7.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 28.91566!))
        Me.TableLayoutPanel7.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 71.08434!))
        Me.TableLayoutPanel7.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35.0!))
        Me.TableLayoutPanel7.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel7.Size = New System.Drawing.Size(407, 90)
        Me.TableLayoutPanel7.TabIndex = 46
        '
        'Label383
        '
        Me.Label383.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label383.AutoSize = True
        Me.Label383.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Label383.Location = New System.Drawing.Point(3, 54)
        Me.Label383.Name = "Label383"
        Me.Label383.Size = New System.Drawing.Size(399, 36)
        Me.Label383.TabIndex = 35
        Me.Label383.Text = "Quando o Momento Fletor é aplicado no eixo y ele causa rotação em x." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Quando apli" &
    "cado no eixo x, ele causa rotação em y."
        Me.Label383.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label91
        '
        Me.Label91.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label91.AutoSize = True
        Me.Label91.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Label91.Location = New System.Drawing.Point(3, 15)
        Me.Label91.Name = "Label91"
        Me.Label91.Size = New System.Drawing.Size(401, 39)
        Me.Label91.TabIndex = 34
        Me.Label91.Text = "As solicitações externas devem ser inseridas com as devidas combinações normativa" &
    "s."
        Me.Label91.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label136
        '
        Me.Label136.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label136.AutoSize = True
        Me.Label136.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Label136.Location = New System.Drawing.Point(174, 0)
        Me.Label136.Name = "Label136"
        Me.Label136.Size = New System.Drawing.Size(59, 15)
        Me.Label136.TabIndex = 32
        Me.Label136.Text = "Notas:"
        Me.Label136.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Form3DadosIniciais
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.imgMomento)
        Me.Controls.Add(Me.TableLayoutPanel7)
        Me.Name = "Form3DadosIniciais"
        Me.Text = "Form3DadosIniciais"
        CType(Me.imgMomento, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel7.ResumeLayout(False)
        Me.TableLayoutPanel7.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents imgMomento As PictureBox
    Friend WithEvents TableLayoutPanel7 As TableLayoutPanel
    Friend WithEvents Label383 As Label
    Friend WithEvents Label91 As Label
    Friend WithEvents Label136 As Label
End Class
