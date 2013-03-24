<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Ordini
    Inherits System.Windows.Forms.Form

    'Form esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Richiesto da Progettazione Windows Form
    Private components As System.ComponentModel.IContainer

    'NOTA: la procedura che segue è richiesta da Progettazione Windows Form
    'Può essere modificata in Progettazione Windows Form.  
    'Non modificarla nell'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtPosizione = New System.Windows.Forms.TextBox
        Me.btnFine = New System.Windows.Forms.Button
        Me.btnPrimo = New System.Windows.Forms.Button
        Me.btnUltimo = New System.Windows.Forms.Button
        Me.btnDopo = New System.Windows.Forms.Button
        Me.btnPrima = New System.Windows.Forms.Button
        Me.grpConferma = New System.Windows.Forms.GroupBox
        Me.btnSalva = New System.Windows.Forms.Button
        Me.btnAnnulla = New System.Windows.Forms.Button
        Me.grpOper = New System.Windows.Forms.GroupBox
        Me.btnElimina = New System.Windows.Forms.Button
        Me.btnModifica = New System.Windows.Forms.Button
        Me.btnAdd = New System.Windows.Forms.Button
        Me.grdOrdini = New System.Windows.Forms.DataGridView
        Me.grpText = New System.Windows.Forms.GroupBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.cmbPag = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmbCliente = New System.Windows.Forms.ComboBox
        Me.txtDataConsegna = New System.Windows.Forms.TextBox
        Me.txtDataOrdine = New System.Windows.Forms.TextBox
        Me.txtNumOrdine = New System.Windows.Forms.TextBox
        Me.btnScegli = New System.Windows.Forms.Button
        Me.BntVai = New System.Windows.Forms.Button
        Me.grpConferma.SuspendLayout()
        Me.grpOper.SuspendLayout()
        CType(Me.grdOrdini, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpText.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(42, 508)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(709, 13)
        Me.Label11.TabIndex = 43
        Me.Label11.Text = "_________________________________________________________________________________" & _
            "____________________________________"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(42, 495)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(709, 13)
        Me.Label12.TabIndex = 42
        Me.Label12.Text = "_________________________________________________________________________________" & _
            "____________________________________"
        '
        'txtPosizione
        '
        Me.txtPosizione.Enabled = False
        Me.txtPosizione.Location = New System.Drawing.Point(352, 471)
        Me.txtPosizione.Name = "txtPosizione"
        Me.txtPosizione.Size = New System.Drawing.Size(100, 20)
        Me.txtPosizione.TabIndex = 41
        '
        'btnFine
        '
        Me.btnFine.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.btnFine.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnFine.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnFine.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnFine.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnFine.Location = New System.Drawing.Point(387, 597)
        Me.btnFine.Name = "btnFine"
        Me.btnFine.Size = New System.Drawing.Size(75, 23)
        Me.btnFine.TabIndex = 36
        Me.btnFine.Text = "Fine"
        Me.btnFine.UseVisualStyleBackColor = False
        '
        'btnPrimo
        '
        Me.btnPrimo.Enabled = False
        Me.btnPrimo.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnPrimo.Location = New System.Drawing.Point(275, 471)
        Me.btnPrimo.Name = "btnPrimo"
        Me.btnPrimo.Size = New System.Drawing.Size(33, 23)
        Me.btnPrimo.TabIndex = 39
        Me.btnPrimo.Text = "<<"
        Me.btnPrimo.UseVisualStyleBackColor = True
        '
        'btnUltimo
        '
        Me.btnUltimo.Enabled = False
        Me.btnUltimo.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnUltimo.Location = New System.Drawing.Point(497, 471)
        Me.btnUltimo.Name = "btnUltimo"
        Me.btnUltimo.Size = New System.Drawing.Size(33, 23)
        Me.btnUltimo.TabIndex = 40
        Me.btnUltimo.Text = ">>"
        Me.btnUltimo.UseVisualStyleBackColor = True
        '
        'btnDopo
        '
        Me.btnDopo.Enabled = False
        Me.btnDopo.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnDopo.Location = New System.Drawing.Point(458, 471)
        Me.btnDopo.Name = "btnDopo"
        Me.btnDopo.Size = New System.Drawing.Size(33, 23)
        Me.btnDopo.TabIndex = 38
        Me.btnDopo.Text = ">"
        Me.btnDopo.UseVisualStyleBackColor = True
        '
        'btnPrima
        '
        Me.btnPrima.Enabled = False
        Me.btnPrima.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnPrima.Location = New System.Drawing.Point(313, 471)
        Me.btnPrima.Name = "btnPrima"
        Me.btnPrima.Size = New System.Drawing.Size(33, 23)
        Me.btnPrima.TabIndex = 37
        Me.btnPrima.Text = "<"
        Me.btnPrima.UseVisualStyleBackColor = True
        '
        'grpConferma
        '
        Me.grpConferma.Controls.Add(Me.btnSalva)
        Me.grpConferma.Controls.Add(Me.btnAnnulla)
        Me.grpConferma.Enabled = False
        Me.grpConferma.Location = New System.Drawing.Point(163, 535)
        Me.grpConferma.Name = "grpConferma"
        Me.grpConferma.Size = New System.Drawing.Size(197, 56)
        Me.grpConferma.TabIndex = 34
        Me.grpConferma.TabStop = False
        Me.grpConferma.Text = "Conferma"
        '
        'btnSalva
        '
        Me.btnSalva.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.btnSalva.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSalva.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnSalva.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnSalva.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSalva.Location = New System.Drawing.Point(103, 20)
        Me.btnSalva.Name = "btnSalva"
        Me.btnSalva.Size = New System.Drawing.Size(75, 23)
        Me.btnSalva.TabIndex = 1
        Me.btnSalva.Text = "Salva"
        Me.btnSalva.UseVisualStyleBackColor = False
        '
        'btnAnnulla
        '
        Me.btnAnnulla.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.btnAnnulla.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnAnnulla.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnAnnulla.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnAnnulla.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAnnulla.Location = New System.Drawing.Point(22, 20)
        Me.btnAnnulla.Name = "btnAnnulla"
        Me.btnAnnulla.Size = New System.Drawing.Size(75, 23)
        Me.btnAnnulla.TabIndex = 0
        Me.btnAnnulla.Text = "Annulla"
        Me.btnAnnulla.UseVisualStyleBackColor = False
        '
        'grpOper
        '
        Me.grpOper.Controls.Add(Me.btnElimina)
        Me.grpOper.Controls.Add(Me.btnModifica)
        Me.grpOper.Controls.Add(Me.btnAdd)
        Me.grpOper.Location = New System.Drawing.Point(366, 535)
        Me.grpOper.Name = "grpOper"
        Me.grpOper.Size = New System.Drawing.Size(275, 56)
        Me.grpOper.TabIndex = 35
        Me.grpOper.TabStop = False
        Me.grpOper.Text = "Operazioni"
        '
        'btnElimina
        '
        Me.btnElimina.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.btnElimina.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnElimina.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnElimina.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnElimina.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnElimina.Location = New System.Drawing.Point(21, 20)
        Me.btnElimina.Name = "btnElimina"
        Me.btnElimina.Size = New System.Drawing.Size(75, 23)
        Me.btnElimina.TabIndex = 11
        Me.btnElimina.Text = "Elimina"
        Me.btnElimina.UseVisualStyleBackColor = False
        '
        'btnModifica
        '
        Me.btnModifica.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.btnModifica.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnModifica.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnModifica.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnModifica.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnModifica.Location = New System.Drawing.Point(183, 20)
        Me.btnModifica.Name = "btnModifica"
        Me.btnModifica.Size = New System.Drawing.Size(75, 23)
        Me.btnModifica.TabIndex = 10
        Me.btnModifica.Text = "Modifica"
        Me.btnModifica.UseVisualStyleBackColor = False
        '
        'btnAdd
        '
        Me.btnAdd.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.btnAdd.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnAdd.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnAdd.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnAdd.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAdd.Location = New System.Drawing.Point(102, 20)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(75, 23)
        Me.btnAdd.TabIndex = 9
        Me.btnAdd.Text = "Aggiungi"
        Me.btnAdd.UseVisualStyleBackColor = False
        '
        'grdOrdini
        '
        Me.grdOrdini.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdOrdini.Location = New System.Drawing.Point(25, 142)
        Me.grdOrdini.Name = "grdOrdini"
        Me.grdOrdini.ReadOnly = True
        Me.grdOrdini.Size = New System.Drawing.Size(770, 323)
        Me.grdOrdini.TabIndex = 44
        '
        'grpText
        '
        Me.grpText.Controls.Add(Me.Label5)
        Me.grpText.Controls.Add(Me.cmbPag)
        Me.grpText.Controls.Add(Me.Label4)
        Me.grpText.Controls.Add(Me.Label3)
        Me.grpText.Controls.Add(Me.Label2)
        Me.grpText.Controls.Add(Me.Label1)
        Me.grpText.Controls.Add(Me.cmbCliente)
        Me.grpText.Controls.Add(Me.txtDataConsegna)
        Me.grpText.Controls.Add(Me.txtDataOrdine)
        Me.grpText.Controls.Add(Me.txtNumOrdine)
        Me.grpText.Location = New System.Drawing.Point(26, 13)
        Me.grpText.Name = "grpText"
        Me.grpText.Size = New System.Drawing.Size(771, 99)
        Me.grpText.TabIndex = 45
        Me.grpText.TabStop = False
        Me.grpText.Text = "Inserimento"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(308, 69)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(64, 13)
        Me.Label5.TabIndex = 67
        Me.Label5.Text = " Pagamento"
        '
        'cmbPag
        '
        Me.cmbPag.FormattingEnabled = True
        Me.cmbPag.Location = New System.Drawing.Point(377, 66)
        Me.cmbPag.Name = "cmbPag"
        Me.cmbPag.Size = New System.Drawing.Size(121, 21)
        Me.cmbPag.TabIndex = 66
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(112, 69)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(42, 13)
        Me.Label4.TabIndex = 64
        Me.Label4.Text = " Cliente"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(502, 20)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 13)
        Me.Label3.TabIndex = 63
        Me.Label3.Text = "Data consegna"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(308, 20)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 13)
        Me.Label2.TabIndex = 62
        Me.Label2.Text = "Data Ordine"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(82, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 13)
        Me.Label1.TabIndex = 61
        Me.Label1.Text = "Numero Ordine"
        '
        'cmbCliente
        '
        Me.cmbCliente.FormattingEnabled = True
        Me.cmbCliente.Location = New System.Drawing.Point(160, 66)
        Me.cmbCliente.Name = "cmbCliente"
        Me.cmbCliente.Size = New System.Drawing.Size(121, 21)
        Me.cmbCliente.TabIndex = 60
        '
        'txtDataConsegna
        '
        Me.txtDataConsegna.Location = New System.Drawing.Point(588, 17)
        Me.txtDataConsegna.Name = "txtDataConsegna"
        Me.txtDataConsegna.Size = New System.Drawing.Size(100, 20)
        Me.txtDataConsegna.TabIndex = 59
        '
        'txtDataOrdine
        '
        Me.txtDataOrdine.Location = New System.Drawing.Point(377, 17)
        Me.txtDataOrdine.Name = "txtDataOrdine"
        Me.txtDataOrdine.Size = New System.Drawing.Size(100, 20)
        Me.txtDataOrdine.TabIndex = 58
        '
        'txtNumOrdine
        '
        Me.txtNumOrdine.Location = New System.Drawing.Point(160, 17)
        Me.txtNumOrdine.Name = "txtNumOrdine"
        Me.txtNumOrdine.Size = New System.Drawing.Size(100, 20)
        Me.txtNumOrdine.TabIndex = 57
        '
        'btnScegli
        '
        Me.btnScegli.Location = New System.Drawing.Point(624, 118)
        Me.btnScegli.Name = "btnScegli"
        Me.btnScegli.Size = New System.Drawing.Size(87, 23)
        Me.btnScegli.TabIndex = 65
        Me.btnScegli.Text = "..."
        Me.btnScegli.UseVisualStyleBackColor = True
        '
        'BntVai
        '
        Me.BntVai.Location = New System.Drawing.Point(566, 471)
        Me.BntVai.Name = "BntVai"
        Me.BntVai.Size = New System.Drawing.Size(75, 23)
        Me.BntVai.TabIndex = 66
        Me.BntVai.Text = "Vai a..."
        Me.BntVai.UseVisualStyleBackColor = True
        '
        'Ordini
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(809, 632)
        Me.Controls.Add(Me.BntVai)
        Me.Controls.Add(Me.btnScegli)
        Me.Controls.Add(Me.grpText)
        Me.Controls.Add(Me.grdOrdini)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.txtPosizione)
        Me.Controls.Add(Me.btnFine)
        Me.Controls.Add(Me.btnPrimo)
        Me.Controls.Add(Me.btnUltimo)
        Me.Controls.Add(Me.btnDopo)
        Me.Controls.Add(Me.btnPrima)
        Me.Controls.Add(Me.grpConferma)
        Me.Controls.Add(Me.grpOper)
        Me.Name = "Ordini"
        Me.Text = "Ordini"
        Me.grpConferma.ResumeLayout(False)
        Me.grpOper.ResumeLayout(False)
        CType(Me.grdOrdini, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpText.ResumeLayout(False)
        Me.grpText.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtPosizione As System.Windows.Forms.TextBox
    Friend WithEvents btnFine As System.Windows.Forms.Button
    Friend WithEvents btnPrimo As System.Windows.Forms.Button
    Friend WithEvents btnUltimo As System.Windows.Forms.Button
    Friend WithEvents btnDopo As System.Windows.Forms.Button
    Friend WithEvents btnPrima As System.Windows.Forms.Button
    Friend WithEvents grpConferma As System.Windows.Forms.GroupBox
    Friend WithEvents btnSalva As System.Windows.Forms.Button
    Friend WithEvents btnAnnulla As System.Windows.Forms.Button
    Friend WithEvents grpOper As System.Windows.Forms.GroupBox
    Friend WithEvents btnElimina As System.Windows.Forms.Button
    Friend WithEvents btnModifica As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents grdOrdini As System.Windows.Forms.DataGridView
    Friend WithEvents grpText As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cmbPag As System.Windows.Forms.ComboBox
    Friend WithEvents btnScegli As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbCliente As System.Windows.Forms.ComboBox
    Friend WithEvents txtDataConsegna As System.Windows.Forms.TextBox
    Friend WithEvents txtDataOrdine As System.Windows.Forms.TextBox
    Friend WithEvents txtNumOrdine As System.Windows.Forms.TextBox
    Friend WithEvents BntVai As System.Windows.Forms.Button
End Class
