<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_ListFormaletas
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_ListFormaletas))
        Me.btnDescargarListado = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cboNumOrden = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboTipoOrden = New System.Windows.Forms.ComboBox()
        Me.ChekRayas = New System.Windows.Forms.CheckedListBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.CheckSoloFM = New System.Windows.Forms.CheckBox()
        Me.CheckSoloAcc = New System.Windows.Forms.CheckBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.checkStockERP = New System.Windows.Forms.CheckBox()
        Me.CheckBox4 = New System.Windows.Forms.CheckBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.checkCliente = New System.Windows.Forms.CheckBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnDescargarListado
        '
        Me.btnDescargarListado.Location = New System.Drawing.Point(26, 40)
        Me.btnDescargarListado.Name = "btnDescargarListado"
        Me.btnDescargarListado.Size = New System.Drawing.Size(140, 38)
        Me.btnDescargarListado.TabIndex = 19
        Me.btnDescargarListado.Text = "Descargar Listado"
        Me.btnDescargarListado.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(379, 9)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(43, 16)
        Me.Label5.TabIndex = 25
        Me.Label5.Text = "Raya:"
        '
        'cboNumOrden
        '
        Me.cboNumOrden.FormattingEnabled = True
        Me.cboNumOrden.Location = New System.Drawing.Point(290, 6)
        Me.cboNumOrden.Name = "cboNumOrden"
        Me.cboNumOrden.Size = New System.Drawing.Size(74, 24)
        Me.cboNumOrden.TabIndex = 23
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(228, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(58, 16)
        Me.Label4.TabIndex = 24
        Me.Label4.Text = "Sol_No.:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(89, 9)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 16)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Tipo_Sol:"
        '
        'cboTipoOrden
        '
        Me.cboTipoOrden.FormattingEnabled = True
        Me.cboTipoOrden.Items.AddRange(New Object() {"PR", "OF", "OK", "OG", "SR", "ID", "OM", "FP", "FA", "RC", "CT", "F4"})
        Me.cboTipoOrden.Location = New System.Drawing.Point(157, 6)
        Me.cboTipoOrden.Name = "cboTipoOrden"
        Me.cboTipoOrden.Size = New System.Drawing.Size(53, 24)
        Me.cboTipoOrden.TabIndex = 22
        '
        'ChekRayas
        '
        Me.ChekRayas.FormattingEnabled = True
        Me.ChekRayas.Location = New System.Drawing.Point(425, 40)
        Me.ChekRayas.Name = "ChekRayas"
        Me.ChekRayas.Size = New System.Drawing.Size(63, 55)
        Me.ChekRayas.TabIndex = 28
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(12, 285)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(404, 50)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 29
        Me.PictureBox1.TabStop = False
        '
        'CheckSoloFM
        '
        Me.CheckSoloFM.AutoSize = True
        Me.CheckSoloFM.Location = New System.Drawing.Point(425, 0)
        Me.CheckSoloFM.Name = "CheckSoloFM"
        Me.CheckSoloFM.Size = New System.Drawing.Size(60, 20)
        Me.CheckSoloFM.TabIndex = 30
        Me.CheckSoloFM.Text = "Form"
        Me.CheckSoloFM.UseVisualStyleBackColor = True
        '
        'CheckSoloAcc
        '
        Me.CheckSoloAcc.AutoSize = True
        Me.CheckSoloAcc.Location = New System.Drawing.Point(425, 19)
        Me.CheckSoloAcc.Name = "CheckSoloAcc"
        Me.CheckSoloAcc.Size = New System.Drawing.Size(52, 20)
        Me.CheckSoloAcc.TabIndex = 31
        Me.CheckSoloAcc.Text = "Acc"
        Me.CheckSoloAcc.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(183, 40)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(140, 38)
        Me.Button2.TabIndex = 32
        Me.Button2.Text = "Descarga y Adapta"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'checkStockERP
        '
        Me.checkStockERP.AutoSize = True
        Me.checkStockERP.Location = New System.Drawing.Point(9, 7)
        Me.checkStockERP.Name = "checkStockERP"
        Me.checkStockERP.Size = New System.Drawing.Size(63, 20)
        Me.checkStockERP.TabIndex = 33
        Me.checkStockERP.Text = "Stock"
        Me.checkStockERP.UseVisualStyleBackColor = True
        '
        'CheckBox4
        '
        Me.CheckBox4.AutoSize = True
        Me.CheckBox4.Checked = True
        Me.CheckBox4.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox4.Location = New System.Drawing.Point(334, 40)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(83, 20)
        Me.CheckBox4.TabIndex = 34
        Me.CheckBox4.Text = "Escalera"
        Me.CheckBox4.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.ColumnHeadersVisible = False
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2})
        Me.DataGridView1.Location = New System.Drawing.Point(54, 84)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowHeadersWidth = 51
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(253, 195)
        Me.DataGridView1.TabIndex = 35
        '
        'Column1
        '
        Me.Column1.HeaderText = "Column1"
        Me.Column1.MinimumWidth = 6
        Me.Column1.Name = "Column1"
        Me.Column1.Width = 125
        '
        'Column2
        '
        Me.Column2.HeaderText = "Column2"
        Me.Column2.MinimumWidth = 6
        Me.Column2.Name = "Column2"
        Me.Column2.Width = 125
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(334, 100)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(85, 38)
        Me.Button3.TabIndex = 36
        Me.Button3.Text = "Combinar"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'checkCliente
        '
        Me.checkCliente.AutoSize = True
        Me.checkCliente.Location = New System.Drawing.Point(334, 74)
        Me.checkCliente.Name = "checkCliente"
        Me.checkCliente.Size = New System.Drawing.Size(70, 20)
        Me.checkCliente.TabIndex = 37
        Me.checkCliente.Text = "Cliente"
        Me.checkCliente.UseVisualStyleBackColor = True
        '
        'Frm_ListFormaletas
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(494, 347)
        Me.Controls.Add(Me.checkCliente)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.CheckBox4)
        Me.Controls.Add(Me.checkStockERP)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.CheckSoloAcc)
        Me.Controls.Add(Me.CheckSoloFM)
        Me.Controls.Add(Me.ChekRayas)
        Me.Controls.Add(Me.btnDescargarListado)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.cboNumOrden)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboTipoOrden)
        Me.Controls.Add(Me.PictureBox1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Frm_ListFormaletas"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Listado de Formaletas"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnDescargarListado As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboNumOrden As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboTipoOrden As System.Windows.Forms.ComboBox
    Friend WithEvents ChekRayas As System.Windows.Forms.CheckedListBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents CheckSoloFM As System.Windows.Forms.CheckBox
    Friend WithEvents CheckSoloAcc As System.Windows.Forms.CheckBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents checkStockERP As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox4 As Windows.Forms.CheckBox
    Friend WithEvents DataGridView1 As Windows.Forms.DataGridView
    Friend WithEvents Column1 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Button3 As Windows.Forms.Button
    Friend WithEvents checkCliente As Windows.Forms.CheckBox
End Class
