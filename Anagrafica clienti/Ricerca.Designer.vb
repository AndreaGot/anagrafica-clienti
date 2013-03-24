<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Ricerca
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
        Me.grdRicerca = New System.Windows.Forms.DataGridView
        Me.txtID = New System.Windows.Forms.TextBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.txtQuery = New System.Windows.Forms.TextBox
        Me.txtCampi = New System.Windows.Forms.TextBox
        Me.lblRicerca = New System.Windows.Forms.Label
        Me.txtRicerca = New System.Windows.Forms.TextBox
        Me.btnannulla = New System.Windows.Forms.Button
        CType(Me.grdRicerca, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grdRicerca
        '
        Me.grdRicerca.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdRicerca.Location = New System.Drawing.Point(24, 12)
        Me.grdRicerca.Name = "grdRicerca"
        Me.grdRicerca.ReadOnly = True
        Me.grdRicerca.Size = New System.Drawing.Size(642, 361)
        Me.grdRicerca.TabIndex = 0
        '
        'txtID
        '
        Me.txtID.Location = New System.Drawing.Point(582, 436)
        Me.txtID.Name = "txtID"
        Me.txtID.Size = New System.Drawing.Size(100, 20)
        Me.txtID.TabIndex = 1
        Me.txtID.Visible = False
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(233, 450)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(107, 29)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "Ok"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'txtQuery
        '
        Me.txtQuery.Location = New System.Drawing.Point(476, 406)
        Me.txtQuery.Name = "txtQuery"
        Me.txtQuery.Size = New System.Drawing.Size(100, 20)
        Me.txtQuery.TabIndex = 4
        Me.txtQuery.Visible = False
        '
        'txtCampi
        '
        Me.txtCampi.Location = New System.Drawing.Point(582, 406)
        Me.txtCampi.Name = "txtCampi"
        Me.txtCampi.Size = New System.Drawing.Size(100, 20)
        Me.txtCampi.TabIndex = 5
        Me.txtCampi.Visible = False
        '
        'lblRicerca
        '
        Me.lblRicerca.AutoSize = True
        Me.lblRicerca.Location = New System.Drawing.Point(21, 394)
        Me.lblRicerca.Name = "lblRicerca"
        Me.lblRicerca.Size = New System.Drawing.Size(65, 13)
        Me.lblRicerca.TabIndex = 6
        Me.lblRicerca.Text = "Ricerca per:"
        '
        'txtRicerca
        '
        Me.txtRicerca.Location = New System.Drawing.Point(92, 391)
        Me.txtRicerca.Name = "txtRicerca"
        Me.txtRicerca.Size = New System.Drawing.Size(159, 20)
        Me.txtRicerca.TabIndex = 7
        '
        'btnannulla
        '
        Me.btnannulla.Location = New System.Drawing.Point(355, 450)
        Me.btnannulla.Name = "btnannulla"
        Me.btnannulla.Size = New System.Drawing.Size(107, 29)
        Me.btnannulla.TabIndex = 8
        Me.btnannulla.Text = "Annulla"
        Me.btnannulla.UseVisualStyleBackColor = True
        '
        'Ricerca
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(694, 491)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnannulla)
        Me.Controls.Add(Me.txtRicerca)
        Me.Controls.Add(Me.lblRicerca)
        Me.Controls.Add(Me.txtCampi)
        Me.Controls.Add(Me.txtQuery)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.txtID)
        Me.Controls.Add(Me.grdRicerca)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "Ricerca"
        Me.Text = "Ricerca"
        CType(Me.grdRicerca, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents grdRicerca As System.Windows.Forms.DataGridView
    Friend WithEvents txtID As System.Windows.Forms.TextBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents txtQuery As System.Windows.Forms.TextBox
    Friend WithEvents txtCampi As System.Windows.Forms.TextBox
    Friend WithEvents lblRicerca As System.Windows.Forms.Label
    Friend WithEvents txtRicerca As System.Windows.Forms.TextBox
    Friend WithEvents btnannulla As System.Windows.Forms.Button
End Class
