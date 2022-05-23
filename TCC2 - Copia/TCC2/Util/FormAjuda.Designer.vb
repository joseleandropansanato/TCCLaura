<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormAjuda
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormAjuda))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Button11 = New System.Windows.Forms.Button()
        Me.btnInicio = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.gbxInicio = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtCargaPerm = New System.Windows.Forms.TextBox()
        Me.GroupBox1.SuspendLayout()
        Me.gbxInicio.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.GroupBox1.Controls.Add(Me.Button11)
        Me.GroupBox1.Controls.Add(Me.btnInicio)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Button3)
        Me.GroupBox1.Controls.Add(Me.Button4)
        Me.GroupBox1.Controls.Add(Me.Button5)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(195, 383)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Button11
        '
        Me.Button11.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button11.Location = New System.Drawing.Point(6, 312)
        Me.Button11.Name = "Button11"
        Me.Button11.Size = New System.Drawing.Size(171, 53)
        Me.Button11.TabIndex = 11
        Me.Button11.Text = "DIMENSIONAMENTO"
        Me.Button11.UseVisualStyleBackColor = True
        '
        'btnInicio
        '
        Me.btnInicio.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnInicio.Location = New System.Drawing.Point(6, 22)
        Me.btnInicio.Name = "btnInicio"
        Me.btnInicio.Size = New System.Drawing.Size(171, 53)
        Me.btnInicio.TabIndex = 1
        Me.btnInicio.Text = "INICIO"
        Me.btnInicio.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button2.Location = New System.Drawing.Point(6, 80)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(171, 53)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "PROPRIEDADES DA MADEIRA"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button3.Location = New System.Drawing.Point(6, 138)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(171, 53)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "RESISTÊNCIA DE CÁLCULO"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button4.Location = New System.Drawing.Point(6, 196)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(171, 53)
        Me.Button4.TabIndex = 4
        Me.Button4.Text = "DADOS INICIAIS"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button5.Location = New System.Drawing.Point(6, 254)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(171, 53)
        Me.Button5.TabIndex = 5
        Me.Button5.Text = "PROPRIEDADES GEOMÉTRICAS"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 19.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Label1.Location = New System.Drawing.Point(141, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(198, 45)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "BEM VINDO"
        '
        'gbxInicio
        '
        Me.gbxInicio.Controls.Add(Me.txtCargaPerm)
        Me.gbxInicio.Controls.Add(Me.Label4)
        Me.gbxInicio.Controls.Add(Me.Label3)
        Me.gbxInicio.Controls.Add(Me.Label2)
        Me.gbxInicio.Controls.Add(Me.Label1)
        Me.gbxInicio.Location = New System.Drawing.Point(222, 12)
        Me.gbxInicio.Name = "gbxInicio"
        Me.gbxInicio.Size = New System.Drawing.Size(458, 383)
        Me.gbxInicio.TabIndex = 2
        Me.gbxInicio.TabStop = False
        Me.gbxInicio.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Label4.Location = New System.Drawing.Point(218, 52)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(29, 23)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "ao"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Label3.Location = New System.Drawing.Point(193, 75)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(81, 31)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "NOME"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 125)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(472, 100)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = resources.GetString("Label2.Text")
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtCargaPerm
        '
        Me.txtCargaPerm.Location = New System.Drawing.Point(218, 312)
        Me.txtCargaPerm.Name = "txtCargaPerm"
        Me.txtCargaPerm.Size = New System.Drawing.Size(125, 27)
        Me.txtCargaPerm.TabIndex = 4
        '
        'FormAjuda
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(725, 450)
        Me.Controls.Add(Me.gbxInicio)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "FormAjuda"
        Me.Text = "Ajuda"
        Me.GroupBox1.ResumeLayout(False)
        Me.gbxInicio.ResumeLayout(False)
        Me.gbxInicio.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents btnInicio As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button11 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents gbxInicio As GroupBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents txtCargaPerm As TextBox
End Class
