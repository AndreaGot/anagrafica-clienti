<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Nazione
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
        Me.grdNazione = New System.Windows.Forms.DataGridView
        Me.Aggiorna = New System.Windows.Forms.Button
        CType(Me.grdNazione, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grdNazione
        '
        Me.grdNazione.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdNazione.Location = New System.Drawing.Point(29, 16)
        Me.grdNazione.Name = "grdNazione"
        Me.grdNazione.Size = New System.Drawing.Size(428, 215)
        Me.grdNazione.TabIndex = 0
        '
        'Aggiorna
        '
        Me.Aggiorna.Location = New System.Drawing.Point(287, 255)
        Me.Aggiorna.Name = "Aggiorna"
        Me.Aggiorna.Size = New System.Drawing.Size(122, 35)
        Me.Aggiorna.TabIndex = 1
        Me.Aggiorna.Text = "&Aggiorna"
        Me.Aggiorna.UseVisualStyleBackColor = True
        '
        'Nazione
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(519, 343)
        Me.Controls.Add(Me.Aggiorna)
        Me.Controls.Add(Me.grdNazione)
        Me.Name = "Nazione"
        Me.ShowIcon = False
        Me.Text = "Gestione nazioni"
        CType(Me.grdNazione, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grdNazione As System.Windows.Forms.DataGridView
    Friend WithEvents Aggiorna As System.Windows.Forms.Button
End Class
