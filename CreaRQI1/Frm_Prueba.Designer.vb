<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_Prueba
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Txt_Entrada = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Btn_Consultar = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Txt_Entrada
        '
        Me.Txt_Entrada.Location = New System.Drawing.Point(23, 58)
        Me.Txt_Entrada.Name = "Txt_Entrada"
        Me.Txt_Entrada.Size = New System.Drawing.Size(294, 22)
        Me.Txt_Entrada.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(134, 151)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Label1"
        '
        'Btn_Consultar
        '
        Me.Btn_Consultar.Location = New System.Drawing.Point(112, 100)
        Me.Btn_Consultar.Name = "Btn_Consultar"
        Me.Btn_Consultar.Size = New System.Drawing.Size(97, 23)
        Me.Btn_Consultar.TabIndex = 2
        Me.Btn_Consultar.Text = "Consultar"
        Me.Btn_Consultar.UseVisualStyleBackColor = True
        '
        'Frm_Prueba
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(357, 206)
        Me.Controls.Add(Me.Btn_Consultar)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Txt_Entrada)
        Me.Name = "Frm_Prueba"
        Me.Text = "Frm_Prueba"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Txt_Entrada As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Btn_Consultar As Windows.Forms.Button
End Class
