<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Pagamento
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
        Me.txtPag = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.chkFineMese = New System.Windows.Forms.CheckBox
        Me.txtGiorni = New System.Windows.Forms.TextBox
        Me.txtRate = New System.Windows.Forms.TextBox
        Me.txtGiorniRate = New System.Windows.Forms.TextBox
        Me.txtSconto = New System.Windows.Forms.TextBox
        Me.txtSpeseBanca = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.BtnScegli = New System.Windows.Forms.Button
        Me.txtID = New System.Windows.Forms.TextBox
        Me.grpConferma.SuspendLayout()
        Me.grpOper.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(-62, 292)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(709, 13)
        Me.Label11.TabIndex = 56
        Me.Label11.Text = "_________________________________________________________________________________" & _
            "____________________________________"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(-62, 279)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(709, 13)
        Me.Label12.TabIndex = 55
        Me.Label12.Text = "_________________________________________________________________________________" & _
            "____________________________________"
        '
        'txtPosizione
        '
        Me.txtPosizione.Enabled = False
        Me.txtPosizione.Location = New System.Drawing.Point(248, 255)
        Me.txtPosizione.Name = "txtPosizione"
        Me.txtPosizione.Size = New System.Drawing.Size(100, 20)
        Me.txtPosizione.TabIndex = 54
        '
        'btnFine
        '
        Me.btnFine.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.btnFine.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnFine.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnFine.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnFine.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnFine.Location = New System.Drawing.Point(277, 381)
        Me.btnFine.Name = "btnFine"
        Me.btnFine.Size = New System.Drawing.Size(75, 23)
        Me.btnFine.TabIndex = 49
        Me.btnFine.Text = "Fine"
        Me.btnFine.UseVisualStyleBackColor = False
        '
        'btnPrimo
        '
        Me.btnPrimo.Enabled = False
        Me.btnPrimo.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnPrimo.Location = New System.Drawing.Point(171, 255)
        Me.btnPrimo.Name = "btnPrimo"
        Me.btnPrimo.Size = New System.Drawing.Size(33, 23)
        Me.btnPrimo.TabIndex = 52
        Me.btnPrimo.Text = "<<"
        Me.btnPrimo.UseVisualStyleBackColor = True
        '
        'btnUltimo
        '
        Me.btnUltimo.Enabled = False
        Me.btnUltimo.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnUltimo.Location = New System.Drawing.Point(393, 255)
        Me.btnUltimo.Name = "btnUltimo"
        Me.btnUltimo.Size = New System.Drawing.Size(33, 23)
        Me.btnUltimo.TabIndex = 53
        Me.btnUltimo.Text = ">>"
        Me.btnUltimo.UseVisualStyleBackColor = True
        '
        'btnDopo
        '
        Me.btnDopo.Enabled = False
        Me.btnDopo.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnDopo.Location = New System.Drawing.Point(354, 255)
        Me.btnDopo.Name = "btnDopo"
        Me.btnDopo.Size = New System.Drawing.Size(33, 23)
        Me.btnDopo.TabIndex = 51
        Me.btnDopo.Text = ">"
        Me.btnDopo.UseVisualStyleBackColor = True
        '
        'btnPrima
        '
        Me.btnPrima.Enabled = False
        Me.btnPrima.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnPrima.Location = New System.Drawing.Point(209, 255)
        Me.btnPrima.Name = "btnPrima"
        Me.btnPrima.Size = New System.Drawing.Size(33, 23)
        Me.btnPrima.TabIndex = 50
        Me.btnPrima.Text = "<"
        Me.btnPrima.UseVisualStyleBackColor = True
        '
        'grpConferma
        '
        Me.grpConferma.Controls.Add(Me.btnSalva)
        Me.grpConferma.Controls.Add(Me.btnAnnulla)
        Me.grpConferma.Enabled = False
        Me.grpConferma.Location = New System.Drawing.Point(53, 319)
        Me.grpConferma.Name = "grpConferma"
        Me.grpConferma.Size = New System.Drawing.Size(197, 56)
        Me.grpConferma.TabIndex = 47
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
        Me.grpOper.Location = New System.Drawing.Point(256, 319)
        Me.grpOper.Name = "grpOper"
        Me.grpOper.Size = New System.Drawing.Size(275, 56)
        Me.grpOper.TabIndex = 48
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
        'txtPag
        '
        Me.txtPag.Location = New System.Drawing.Point(256, 37)
        Me.txtPag.MaxLength = 30
        Me.txtPag.Name = "txtPag"
        Me.txtPag.Size = New System.Drawing.Size(146, 20)
        Me.txtPag.TabIndex = 45
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(165, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 13)
        Me.Label1.TabIndex = 44
        Me.Label1.Text = "Tipo Pagamento"
        '
        'chkFineMese
        '
        Me.chkFineMese.AutoSize = True
        Me.chkFineMese.Location = New System.Drawing.Point(195, 193)
        Me.chkFineMese.Name = "chkFineMese"
        Me.chkFineMese.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkFineMese.Size = New System.Drawing.Size(74, 17)
        Me.chkFineMese.TabIndex = 57
        Me.chkFineMese.Text = "Fine mese"
        Me.chkFineMese.UseVisualStyleBackColor = True
        '
        'txtGiorni
        '
        Me.txtGiorni.Location = New System.Drawing.Point(256, 63)
        Me.txtGiorni.Name = "txtGiorni"
        Me.txtGiorni.Size = New System.Drawing.Size(100, 20)
        Me.txtGiorni.TabIndex = 58
        '
        'txtRate
        '
        Me.txtRate.Location = New System.Drawing.Point(256, 89)
        Me.txtRate.Name = "txtRate"
        Me.txtRate.Size = New System.Drawing.Size(100, 20)
        Me.txtRate.TabIndex = 59
        '
        'txtGiorniRate
        '
        Me.txtGiorniRate.Location = New System.Drawing.Point(256, 115)
        Me.txtGiorniRate.Name = "txtGiorniRate"
        Me.txtGiorniRate.Size = New System.Drawing.Size(100, 20)
        Me.txtGiorniRate.TabIndex = 60
        '
        'txtSconto
        '
        Me.txtSconto.Location = New System.Drawing.Point(256, 141)
        Me.txtSconto.Name = "txtSconto"
        Me.txtSconto.Size = New System.Drawing.Size(100, 20)
        Me.txtSconto.TabIndex = 61
        '
        'txtSpeseBanca
        '
        Me.txtSpeseBanca.Location = New System.Drawing.Point(256, 167)
        Me.txtSpeseBanca.Name = "txtSpeseBanca"
        Me.txtSpeseBanca.Size = New System.Drawing.Size(100, 20)
        Me.txtSpeseBanca.TabIndex = 62
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(216, 66)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(34, 13)
        Me.Label2.TabIndex = 63
        Me.Label2.Text = "Giorni"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(220, 92)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 64
        Me.Label3.Text = "Rate"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(175, 118)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(75, 13)
        Me.Label4.TabIndex = 65
        Me.Label4.Text = "Giorni tra Rate"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(211, 144)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(41, 13)
        Me.Label5.TabIndex = 66
        Me.Label5.Text = "Sconto"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(179, 170)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(71, 13)
        Me.Label6.TabIndex = 67
        Me.Label6.Text = "Spese Banca"
        '
        'BtnScegli
        '
        Me.BtnScegli.Location = New System.Drawing.Point(439, 34)
        Me.BtnScegli.Name = "BtnScegli"
        Me.BtnScegli.Size = New System.Drawing.Size(42, 23)
        Me.BtnScegli.TabIndex = 68
        Me.BtnScegli.Text = "..."
        Me.BtnScegli.UseVisualStyleBackColor = True
        '
        'txtID
        '
        Me.txtID.Location = New System.Drawing.Point(62, 98)
        Me.txtID.Name = "txtID"
        Me.txtID.Size = New System.Drawing.Size(42, 20)
        Me.txtID.TabIndex = 69
        Me.txtID.Visible = False
        '
        'Pagamento
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(582, 484)
        Me.Controls.Add(Me.txtID)
        Me.Controls.Add(Me.BtnScegli)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtSpeseBanca)
        Me.Controls.Add(Me.txtSconto)
        Me.Controls.Add(Me.txtGiorniRate)
        Me.Controls.Add(Me.txtRate)
        Me.Controls.Add(Me.txtGiorni)
        Me.Controls.Add(Me.chkFineMese)
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
        Me.Controls.Add(Me.txtPag)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Pagamento"
        Me.ShowIcon = False
        Me.Text = "Modalità pagamento"
        Me.grpConferma.ResumeLayout(False)
        Me.grpOper.ResumeLayout(False)
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
    Friend WithEvents txtPag As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents chkFineMese As System.Windows.Forms.CheckBox
    Friend WithEvents txtGiorni As System.Windows.Forms.TextBox
    Friend WithEvents txtRate As System.Windows.Forms.TextBox
    Friend WithEvents txtGiorniRate As System.Windows.Forms.TextBox
    Friend WithEvents txtSconto As System.Windows.Forms.TextBox
    Friend WithEvents txtSpeseBanca As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents BtnScegli As System.Windows.Forms.Button
    Friend WithEvents txtID As System.Windows.Forms.TextBox
End Class
