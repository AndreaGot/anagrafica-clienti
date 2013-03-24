<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Articoli
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
        Me.grpDati = New System.Windows.Forms.GroupBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.txtFK = New System.Windows.Forms.TextBox
        Me.txtdescrizione = New System.Windows.Forms.TextBox
        Me.txtultpracq = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.txtcostomedio = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.txtlottoeco = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.txtscortamin = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.txtscortamax = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.cmbIVA = New System.Windows.Forms.ComboBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.txtscontoB = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtscontoA = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtsconto2 = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtqtasconto2 = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtsconto = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtqtasconto1 = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtprlistino = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.cmbcatart = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cmbUM = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.BtnScegli = New System.Windows.Forms.Button
        Me.grpConferma.SuspendLayout()
        Me.grpOper.SuspendLayout()
        Me.grpDati.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(-120, 342)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(1021, 13)
        Me.Label11.TabIndex = 43
        Me.Label11.Text = "_________________________________________________________________________________" & _
            "________________________________________________________________________________" & _
            "________" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(-128, 329)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(1063, 13)
        Me.Label12.TabIndex = 42
        Me.Label12.Text = "_________________________________________________________________________________" & _
            "________________________________________________________________________________" & _
            "_______________"
        '
        'txtPosizione
        '
        Me.txtPosizione.Enabled = False
        Me.txtPosizione.Location = New System.Drawing.Point(353, 305)
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
        Me.btnFine.Location = New System.Drawing.Point(353, 431)
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
        Me.btnPrimo.Location = New System.Drawing.Point(276, 305)
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
        Me.btnUltimo.Location = New System.Drawing.Point(498, 305)
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
        Me.btnDopo.Location = New System.Drawing.Point(459, 305)
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
        Me.btnPrima.Location = New System.Drawing.Point(314, 305)
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
        Me.grpConferma.Location = New System.Drawing.Point(164, 369)
        Me.grpConferma.Name = "grpConferma"
        Me.grpConferma.Size = New System.Drawing.Size(197, 56)
        Me.grpConferma.TabIndex = 34
        Me.grpConferma.TabStop = False
        Me.grpConferma.Text = "Conferma"
        '
        'btnSalva
        '
        Me.btnSalva.Location = New System.Drawing.Point(113, 20)
        Me.btnSalva.Name = "btnSalva"
        Me.btnSalva.Size = New System.Drawing.Size(75, 23)
        Me.btnSalva.TabIndex = 1
        Me.btnSalva.Text = "Salva"
        Me.btnSalva.UseVisualStyleBackColor = True
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
        Me.grpOper.Location = New System.Drawing.Point(367, 369)
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
        'grpDati
        '
        Me.grpDati.Controls.Add(Me.Label19)
        Me.grpDati.Controls.Add(Me.txtFK)
        Me.grpDati.Controls.Add(Me.txtdescrizione)
        Me.grpDati.Controls.Add(Me.txtultpracq)
        Me.grpDati.Controls.Add(Me.Label18)
        Me.grpDati.Controls.Add(Me.txtcostomedio)
        Me.grpDati.Controls.Add(Me.Label17)
        Me.grpDati.Controls.Add(Me.txtlottoeco)
        Me.grpDati.Controls.Add(Me.Label16)
        Me.grpDati.Controls.Add(Me.txtscortamin)
        Me.grpDati.Controls.Add(Me.Label15)
        Me.grpDati.Controls.Add(Me.txtscortamax)
        Me.grpDati.Controls.Add(Me.Label14)
        Me.grpDati.Controls.Add(Me.cmbIVA)
        Me.grpDati.Controls.Add(Me.Label13)
        Me.grpDati.Controls.Add(Me.txtscontoB)
        Me.grpDati.Controls.Add(Me.Label10)
        Me.grpDati.Controls.Add(Me.txtscontoA)
        Me.grpDati.Controls.Add(Me.Label9)
        Me.grpDati.Controls.Add(Me.txtsconto2)
        Me.grpDati.Controls.Add(Me.Label8)
        Me.grpDati.Controls.Add(Me.txtqtasconto2)
        Me.grpDati.Controls.Add(Me.Label7)
        Me.grpDati.Controls.Add(Me.txtsconto)
        Me.grpDati.Controls.Add(Me.Label6)
        Me.grpDati.Controls.Add(Me.txtqtasconto1)
        Me.grpDati.Controls.Add(Me.Label5)
        Me.grpDati.Controls.Add(Me.txtprlistino)
        Me.grpDati.Controls.Add(Me.Label4)
        Me.grpDati.Controls.Add(Me.cmbcatart)
        Me.grpDati.Controls.Add(Me.Label3)
        Me.grpDati.Controls.Add(Me.cmbUM)
        Me.grpDati.Controls.Add(Me.Label2)
        Me.grpDati.Controls.Add(Me.Label1)
        Me.grpDati.Location = New System.Drawing.Point(12, 12)
        Me.grpDati.Name = "grpDati"
        Me.grpDati.Size = New System.Drawing.Size(743, 277)
        Me.grpDati.TabIndex = 44
        Me.grpDati.TabStop = False
        Me.grpDati.Text = "Dati"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(15, 115)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(40, 13)
        Me.Label19.TabIndex = 109
        Me.Label19.Text = "Codice"
        '
        'txtFK
        '
        Me.txtFK.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.txtFK.Location = New System.Drawing.Point(89, 108)
        Me.txtFK.Name = "txtFK"
        Me.txtFK.ReadOnly = True
        Me.txtFK.Size = New System.Drawing.Size(146, 20)
        Me.txtFK.TabIndex = 108
        '
        'txtdescrizione
        '
        Me.txtdescrizione.Location = New System.Drawing.Point(114, 15)
        Me.txtdescrizione.Multiline = True
        Me.txtdescrizione.Name = "txtdescrizione"
        Me.txtdescrizione.Size = New System.Drawing.Size(149, 55)
        Me.txtdescrizione.TabIndex = 107
        '
        'txtultpracq
        '
        Me.txtultpracq.Location = New System.Drawing.Point(522, 212)
        Me.txtultpracq.Name = "txtultpracq"
        Me.txtultpracq.Size = New System.Drawing.Size(97, 20)
        Me.txtultpracq.TabIndex = 106
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(433, 219)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(65, 13)
        Me.Label18.TabIndex = 105
        Me.Label18.Text = "UltimoPrAcq"
        '
        'txtcostomedio
        '
        Me.txtcostomedio.Location = New System.Drawing.Point(330, 212)
        Me.txtcostomedio.Name = "txtcostomedio"
        Me.txtcostomedio.Size = New System.Drawing.Size(61, 20)
        Me.txtcostomedio.TabIndex = 104
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(252, 215)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(66, 13)
        Me.Label17.TabIndex = 103
        Me.Label17.Text = "Costo Medio"
        '
        'txtlottoeco
        '
        Me.txtlottoeco.Location = New System.Drawing.Point(89, 187)
        Me.txtlottoeco.Name = "txtlottoeco"
        Me.txtlottoeco.Size = New System.Drawing.Size(149, 20)
        Me.txtlottoeco.TabIndex = 102
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(15, 194)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(53, 13)
        Me.Label16.TabIndex = 101
        Me.Label16.Text = "LottoECO"
        '
        'txtscortamin
        '
        Me.txtscortamin.Location = New System.Drawing.Point(522, 186)
        Me.txtscortamin.Name = "txtscortamin"
        Me.txtscortamin.Size = New System.Drawing.Size(58, 20)
        Me.txtscortamin.TabIndex = 100
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(433, 193)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(55, 13)
        Me.Label15.TabIndex = 99
        Me.Label15.Text = "ScortaMin"
        '
        'txtscortamax
        '
        Me.txtscortamax.Location = New System.Drawing.Point(330, 186)
        Me.txtscortamax.Name = "txtscortamax"
        Me.txtscortamax.Size = New System.Drawing.Size(61, 20)
        Me.txtscortamax.TabIndex = 98
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(252, 193)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(58, 13)
        Me.Label14.TabIndex = 97
        Me.Label14.Text = "ScortaMax" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'cmbIVA
        '
        Me.cmbIVA.FormattingEnabled = True
        Me.cmbIVA.Location = New System.Drawing.Point(89, 160)
        Me.cmbIVA.Name = "cmbIVA"
        Me.cmbIVA.Size = New System.Drawing.Size(138, 21)
        Me.cmbIVA.TabIndex = 96
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(15, 167)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(65, 13)
        Me.Label13.TabIndex = 95
        Me.Label13.Text = "Aliquota IVA"
        '
        'txtscontoB
        '
        Me.txtscontoB.Location = New System.Drawing.Point(522, 160)
        Me.txtscontoB.Name = "txtscontoB"
        Me.txtscontoB.Size = New System.Drawing.Size(58, 20)
        Me.txtscontoB.TabIndex = 94
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(433, 167)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(51, 13)
        Me.Label10.TabIndex = 93
        Me.Label10.Text = "Sconto B"
        '
        'txtscontoA
        '
        Me.txtscontoA.Location = New System.Drawing.Point(330, 160)
        Me.txtscontoA.Name = "txtscontoA"
        Me.txtscontoA.Size = New System.Drawing.Size(61, 20)
        Me.txtscontoA.TabIndex = 92
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(252, 167)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(51, 13)
        Me.Label9.TabIndex = 91
        Me.Label9.Text = "Sconto A" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'txtsconto2
        '
        Me.txtsconto2.Location = New System.Drawing.Point(522, 134)
        Me.txtsconto2.Name = "txtsconto2"
        Me.txtsconto2.Size = New System.Drawing.Size(58, 20)
        Me.txtsconto2.TabIndex = 90
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(433, 141)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(50, 13)
        Me.Label8.TabIndex = 89
        Me.Label8.Text = "Sconto 2"
        '
        'txtqtasconto2
        '
        Me.txtqtasconto2.Location = New System.Drawing.Point(330, 134)
        Me.txtqtasconto2.Name = "txtqtasconto2"
        Me.txtqtasconto2.Size = New System.Drawing.Size(61, 20)
        Me.txtqtasconto2.TabIndex = 88
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(252, 141)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(72, 13)
        Me.Label7.TabIndex = 87
        Me.Label7.Text = "QTASconto 2"
        '
        'txtsconto
        '
        Me.txtsconto.Location = New System.Drawing.Point(522, 108)
        Me.txtsconto.Name = "txtsconto"
        Me.txtsconto.Size = New System.Drawing.Size(58, 20)
        Me.txtsconto.TabIndex = 86
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(433, 115)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(50, 13)
        Me.Label6.TabIndex = 85
        Me.Label6.Text = "Sconto 1"
        '
        'txtqtasconto1
        '
        Me.txtqtasconto1.Location = New System.Drawing.Point(330, 108)
        Me.txtqtasconto1.Name = "txtqtasconto1"
        Me.txtqtasconto1.Size = New System.Drawing.Size(61, 20)
        Me.txtqtasconto1.TabIndex = 84
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(252, 115)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 13)
        Me.Label5.TabIndex = 83
        Me.Label5.Text = "QTASconto 1"
        '
        'txtprlistino
        '
        Me.txtprlistino.Location = New System.Drawing.Point(89, 134)
        Me.txtprlistino.Name = "txtprlistino"
        Me.txtprlistino.Size = New System.Drawing.Size(149, 20)
        Me.txtprlistino.TabIndex = 82
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(15, 141)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(68, 13)
        Me.Label4.TabIndex = 81
        Me.Label4.Text = "Prezzo listino"
        '
        'cmbcatart
        '
        Me.cmbcatart.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbcatart.FormattingEnabled = True
        Me.cmbcatart.Location = New System.Drawing.Point(586, 15)
        Me.cmbcatart.Name = "cmbcatart"
        Me.cmbcatart.Size = New System.Drawing.Size(138, 21)
        Me.cmbcatart.TabIndex = 80
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(516, 23)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 13)
        Me.Label3.TabIndex = 79
        Me.Label3.Text = "Cat. Articolo"
        '
        'cmbUM
        '
        Me.cmbUM.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbUM.FormattingEnabled = True
        Me.cmbUM.Location = New System.Drawing.Point(360, 15)
        Me.cmbUM.Name = "cmbUM"
        Me.cmbUM.Size = New System.Drawing.Size(138, 21)
        Me.cmbUM.TabIndex = 78
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(273, 23)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 13)
        Me.Label2.TabIndex = 77
        Me.Label2.Text = "Unità di misura"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(99, 13)
        Me.Label1.TabIndex = 76
        Me.Label1.Text = "Descrizione articolo"
        '
        'BtnScegli
        '
        Me.BtnScegli.Location = New System.Drawing.Point(689, 514)
        Me.BtnScegli.Name = "BtnScegli"
        Me.BtnScegli.Size = New System.Drawing.Size(49, 42)
        Me.BtnScegli.TabIndex = 110
        Me.BtnScegli.Text = "..."
        Me.BtnScegli.UseVisualStyleBackColor = True
        '
        'Articoli
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(767, 472)
        Me.Controls.Add(Me.BtnScegli)
        Me.Controls.Add(Me.grpDati)
        Me.Controls.Add(Me.btnFine)
        Me.Controls.Add(Me.btnDopo)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.grpOper)
        Me.Controls.Add(Me.grpConferma)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.btnPrima)
        Me.Controls.Add(Me.btnUltimo)
        Me.Controls.Add(Me.txtPosizione)
        Me.Controls.Add(Me.btnPrimo)
        Me.Name = "Articoli"
        Me.ShowIcon = False
        Me.Text = "Articoli"
        Me.grpConferma.ResumeLayout(False)
        Me.grpOper.ResumeLayout(False)
        Me.grpDati.ResumeLayout(False)
        Me.grpDati.PerformLayout()
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
    Friend WithEvents btnAnnulla As System.Windows.Forms.Button
    Friend WithEvents grpOper As System.Windows.Forms.GroupBox
    Friend WithEvents btnElimina As System.Windows.Forms.Button
    Friend WithEvents btnModifica As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents grpDati As System.Windows.Forms.GroupBox
    Friend WithEvents txtdescrizione As System.Windows.Forms.TextBox
    Friend WithEvents txtultpracq As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtcostomedio As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents txtlottoeco As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtscortamin As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtscortamax As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents cmbIVA As System.Windows.Forms.ComboBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtscontoB As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtscontoA As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtsconto2 As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtqtasconto2 As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtsconto As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtqtasconto1 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtprlistino As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmbcatart As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmbUM As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnSalva As System.Windows.Forms.Button
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents txtFK As System.Windows.Forms.TextBox
    Friend WithEvents BtnScegli As System.Windows.Forms.Button
End Class
