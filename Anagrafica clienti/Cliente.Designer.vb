<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Cliente
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
        Me.grpInserisci = New System.Windows.Forms.GroupBox
        Me.txtCitta = New System.Windows.Forms.TextBox
        Me.txtCAP = New System.Windows.Forms.TextBox
        Me.txtVia = New System.Windows.Forms.TextBox
        Me.txtCodFisc = New System.Windows.Forms.TextBox
        Me.txtPartIVA = New System.Windows.Forms.TextBox
        Me.txtCognome = New System.Windows.Forms.TextBox
        Me.txtNome = New System.Windows.Forms.TextBox
        Me.txtRagSoc = New System.Windows.Forms.TextBox
        Me.cmbNazione = New System.Windows.Forms.ComboBox
        Me.cmbPag = New System.Windows.Forms.ComboBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
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
        Me.BtnScegli = New System.Windows.Forms.Button
        Me.txtid = New System.Windows.Forms.TextBox
        Me.grpInserisci.SuspendLayout()
        Me.grpConferma.SuspendLayout()
        Me.grpOper.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpInserisci
        '
        Me.grpInserisci.Controls.Add(Me.txtCitta)
        Me.grpInserisci.Controls.Add(Me.txtCAP)
        Me.grpInserisci.Controls.Add(Me.txtVia)
        Me.grpInserisci.Controls.Add(Me.txtCodFisc)
        Me.grpInserisci.Controls.Add(Me.txtPartIVA)
        Me.grpInserisci.Controls.Add(Me.txtCognome)
        Me.grpInserisci.Controls.Add(Me.txtNome)
        Me.grpInserisci.Controls.Add(Me.txtRagSoc)
        Me.grpInserisci.Controls.Add(Me.cmbNazione)
        Me.grpInserisci.Controls.Add(Me.cmbPag)
        Me.grpInserisci.Controls.Add(Me.Label10)
        Me.grpInserisci.Controls.Add(Me.Label9)
        Me.grpInserisci.Controls.Add(Me.Label8)
        Me.grpInserisci.Controls.Add(Me.Label7)
        Me.grpInserisci.Controls.Add(Me.Label6)
        Me.grpInserisci.Controls.Add(Me.Label5)
        Me.grpInserisci.Controls.Add(Me.Label4)
        Me.grpInserisci.Controls.Add(Me.Label3)
        Me.grpInserisci.Controls.Add(Me.Label2)
        Me.grpInserisci.Controls.Add(Me.Label1)
        Me.grpInserisci.Location = New System.Drawing.Point(103, 68)
        Me.grpInserisci.Name = "grpInserisci"
        Me.grpInserisci.Size = New System.Drawing.Size(393, 275)
        Me.grpInserisci.TabIndex = 0
        Me.grpInserisci.TabStop = False
        Me.grpInserisci.Text = "Inserimento"
        '
        'txtCitta
        '
        Me.txtCitta.Location = New System.Drawing.Point(165, 189)
        Me.txtCitta.MaxLength = 20
        Me.txtCitta.Name = "txtCitta"
        Me.txtCitta.Size = New System.Drawing.Size(79, 20)
        Me.txtCitta.TabIndex = 29
        '
        'txtCAP
        '
        Me.txtCAP.Location = New System.Drawing.Point(165, 163)
        Me.txtCAP.MaxLength = 5
        Me.txtCAP.Name = "txtCAP"
        Me.txtCAP.Size = New System.Drawing.Size(48, 20)
        Me.txtCAP.TabIndex = 28
        '
        'txtVia
        '
        Me.txtVia.Location = New System.Drawing.Point(165, 140)
        Me.txtVia.MaxLength = 30
        Me.txtVia.Name = "txtVia"
        Me.txtVia.Size = New System.Drawing.Size(170, 20)
        Me.txtVia.TabIndex = 27
        '
        'txtCodFisc
        '
        Me.txtCodFisc.Location = New System.Drawing.Point(165, 117)
        Me.txtCodFisc.MaxLength = 16
        Me.txtCodFisc.Name = "txtCodFisc"
        Me.txtCodFisc.Size = New System.Drawing.Size(121, 20)
        Me.txtCodFisc.TabIndex = 26
        '
        'txtPartIVA
        '
        Me.txtPartIVA.Location = New System.Drawing.Point(165, 91)
        Me.txtPartIVA.MaxLength = 11
        Me.txtPartIVA.Name = "txtPartIVA"
        Me.txtPartIVA.Size = New System.Drawing.Size(79, 20)
        Me.txtPartIVA.TabIndex = 25
        '
        'txtCognome
        '
        Me.txtCognome.Location = New System.Drawing.Point(165, 67)
        Me.txtCognome.MaxLength = 30
        Me.txtCognome.Name = "txtCognome"
        Me.txtCognome.Size = New System.Drawing.Size(121, 20)
        Me.txtCognome.TabIndex = 24
        '
        'txtNome
        '
        Me.txtNome.Location = New System.Drawing.Point(165, 39)
        Me.txtNome.MaxLength = 20
        Me.txtNome.Name = "txtNome"
        Me.txtNome.Size = New System.Drawing.Size(121, 20)
        Me.txtNome.TabIndex = 23
        '
        'txtRagSoc
        '
        Me.txtRagSoc.Location = New System.Drawing.Point(165, 16)
        Me.txtRagSoc.MaxLength = 30
        Me.txtRagSoc.Name = "txtRagSoc"
        Me.txtRagSoc.Size = New System.Drawing.Size(195, 20)
        Me.txtRagSoc.TabIndex = 22
        '
        'cmbNazione
        '
        Me.cmbNazione.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbNazione.FormattingEnabled = True
        Me.cmbNazione.Location = New System.Drawing.Point(165, 214)
        Me.cmbNazione.Name = "cmbNazione"
        Me.cmbNazione.Size = New System.Drawing.Size(121, 21)
        Me.cmbNazione.TabIndex = 21
        '
        'cmbPag
        '
        Me.cmbPag.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPag.FormattingEnabled = True
        Me.cmbPag.Location = New System.Drawing.Point(165, 240)
        Me.cmbPag.Name = "cmbPag"
        Me.cmbPag.Size = New System.Drawing.Size(121, 21)
        Me.cmbPag.TabIndex = 20
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(20, 240)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(61, 13)
        Me.Label10.TabIndex = 19
        Me.Label10.Text = "Pagamento"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(20, 214)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(46, 13)
        Me.Label9.TabIndex = 18
        Me.Label9.Text = "Nazione"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(17, 189)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(31, 13)
        Me.Label8.TabIndex = 17
        Me.Label8.Text = " Città"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(17, 163)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(31, 13)
        Me.Label7.TabIndex = 16
        Me.Label7.Text = " CAP"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(17, 140)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(25, 13)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = " Via"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(17, 98)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(60, 13)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = " Partita IVA"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(17, 120)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(76, 13)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Codice Fiscale"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(17, 67)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 13)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = " Cognome"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(17, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = " Nome"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(20, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = " Ragione Sociale"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(-63, 430)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(709, 13)
        Me.Label11.TabIndex = 33
        Me.Label11.Text = "_________________________________________________________________________________" & _
            "____________________________________"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(-63, 417)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(709, 13)
        Me.Label12.TabIndex = 32
        Me.Label12.Text = "_________________________________________________________________________________" & _
            "____________________________________"
        '
        'txtPosizione
        '
        Me.txtPosizione.Enabled = False
        Me.txtPosizione.Location = New System.Drawing.Point(247, 393)
        Me.txtPosizione.Name = "txtPosizione"
        Me.txtPosizione.Size = New System.Drawing.Size(100, 20)
        Me.txtPosizione.TabIndex = 31
        '
        'btnFine
        '
        Me.btnFine.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.btnFine.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnFine.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnFine.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnFine.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnFine.Location = New System.Drawing.Point(282, 519)
        Me.btnFine.Name = "btnFine"
        Me.btnFine.Size = New System.Drawing.Size(75, 23)
        Me.btnFine.TabIndex = 26
        Me.btnFine.Text = "Fine"
        Me.btnFine.UseVisualStyleBackColor = False
        '
        'btnPrimo
        '
        Me.btnPrimo.Enabled = False
        Me.btnPrimo.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnPrimo.Location = New System.Drawing.Point(170, 393)
        Me.btnPrimo.Name = "btnPrimo"
        Me.btnPrimo.Size = New System.Drawing.Size(33, 23)
        Me.btnPrimo.TabIndex = 29
        Me.btnPrimo.Text = "<<"
        Me.btnPrimo.UseVisualStyleBackColor = True
        '
        'btnUltimo
        '
        Me.btnUltimo.Enabled = False
        Me.btnUltimo.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnUltimo.Location = New System.Drawing.Point(392, 393)
        Me.btnUltimo.Name = "btnUltimo"
        Me.btnUltimo.Size = New System.Drawing.Size(33, 23)
        Me.btnUltimo.TabIndex = 30
        Me.btnUltimo.Text = ">>"
        Me.btnUltimo.UseVisualStyleBackColor = True
        '
        'btnDopo
        '
        Me.btnDopo.Enabled = False
        Me.btnDopo.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnDopo.Location = New System.Drawing.Point(353, 393)
        Me.btnDopo.Name = "btnDopo"
        Me.btnDopo.Size = New System.Drawing.Size(33, 23)
        Me.btnDopo.TabIndex = 28
        Me.btnDopo.Text = ">"
        Me.btnDopo.UseVisualStyleBackColor = True
        '
        'btnPrima
        '
        Me.btnPrima.Enabled = False
        Me.btnPrima.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnPrima.Location = New System.Drawing.Point(208, 393)
        Me.btnPrima.Name = "btnPrima"
        Me.btnPrima.Size = New System.Drawing.Size(33, 23)
        Me.btnPrima.TabIndex = 27
        Me.btnPrima.Text = "<"
        Me.btnPrima.UseVisualStyleBackColor = True
        '
        'grpConferma
        '
        Me.grpConferma.Controls.Add(Me.btnSalva)
        Me.grpConferma.Controls.Add(Me.btnAnnulla)
        Me.grpConferma.Enabled = False
        Me.grpConferma.Location = New System.Drawing.Point(58, 457)
        Me.grpConferma.Name = "grpConferma"
        Me.grpConferma.Size = New System.Drawing.Size(197, 56)
        Me.grpConferma.TabIndex = 24
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
        Me.grpOper.Location = New System.Drawing.Point(261, 457)
        Me.grpOper.Name = "grpOper"
        Me.grpOper.Size = New System.Drawing.Size(275, 56)
        Me.grpOper.TabIndex = 25
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
        'BtnScegli
        '
        Me.BtnScegli.Location = New System.Drawing.Point(515, 84)
        Me.BtnScegli.Name = "BtnScegli"
        Me.BtnScegli.Size = New System.Drawing.Size(42, 23)
        Me.BtnScegli.TabIndex = 34
        Me.BtnScegli.Text = "..."
        Me.BtnScegli.UseVisualStyleBackColor = True
        '
        'txtid
        '
        Me.txtid.Location = New System.Drawing.Point(148, 21)
        Me.txtid.MaxLength = 456456456
        Me.txtid.Name = "txtid"
        Me.txtid.Size = New System.Drawing.Size(48, 20)
        Me.txtid.TabIndex = 35
        Me.txtid.Visible = False
        '
        'Cliente
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(614, 583)
        Me.Controls.Add(Me.txtid)
        Me.Controls.Add(Me.BtnScegli)
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
        Me.Controls.Add(Me.grpInserisci)
        Me.Name = "Cliente"
        Me.ShowIcon = False
        Me.Text = "Gestione Clienti"
        Me.grpInserisci.ResumeLayout(False)
        Me.grpInserisci.PerformLayout()
        Me.grpConferma.ResumeLayout(False)
        Me.grpOper.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents grpInserisci As System.Windows.Forms.GroupBox
    Friend WithEvents txtCitta As System.Windows.Forms.TextBox
    Friend WithEvents txtCAP As System.Windows.Forms.TextBox
    Friend WithEvents txtVia As System.Windows.Forms.TextBox
    Friend WithEvents txtCodFisc As System.Windows.Forms.TextBox
    Friend WithEvents txtPartIVA As System.Windows.Forms.TextBox
    Friend WithEvents txtCognome As System.Windows.Forms.TextBox
    Friend WithEvents txtNome As System.Windows.Forms.TextBox
    Friend WithEvents txtRagSoc As System.Windows.Forms.TextBox
    Friend WithEvents cmbNazione As System.Windows.Forms.ComboBox
    Friend WithEvents cmbPag As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
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
    Friend WithEvents BtnScegli As System.Windows.Forms.Button
    Friend WithEvents txtid As System.Windows.Forms.TextBox
End Class
