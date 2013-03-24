<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Categorie
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
        Me.btnAggiorna = New System.Windows.Forms.Button
        Me.grdCatArt = New System.Windows.Forms.DataGridView
        CType(Me.grdCatArt, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnAggiorna
        '
        Me.btnAggiorna.Location = New System.Drawing.Point(333, 241)
        Me.btnAggiorna.Name = "btnAggiorna"
        Me.btnAggiorna.Size = New System.Drawing.Size(122, 35)
        Me.btnAggiorna.TabIndex = 5
        Me.btnAggiorna.Text = "&Aggiorna"
        Me.btnAggiorna.UseVisualStyleBackColor = True
        '
        'grdCatArt
        '
        Me.grdCatArt.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdCatArt.Location = New System.Drawing.Point(15, 16)
        Me.grdCatArt.Name = "grdCatArt"
        Me.grdCatArt.Size = New System.Drawing.Size(455, 219)
        Me.grdCatArt.TabIndex = 4
        '
        'Categorie
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(485, 292)
        Me.Controls.Add(Me.btnAggiorna)
        Me.Controls.Add(Me.grdCatArt)
        Me.Name = "Categorie"
        Me.ShowIcon = False
        Me.Text = "Categorie"
        CType(Me.grdCatArt, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnAggiorna As System.Windows.Forms.Button
    Friend WithEvents grdCatArt As System.Windows.Forms.DataGridView
End Class
