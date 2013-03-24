<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Aliquote
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
        Me.grdIva = New System.Windows.Forms.DataGridView
        CType(Me.grdIva, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnAggiorna
        '
        Me.btnAggiorna.Location = New System.Drawing.Point(445, 282)
        Me.btnAggiorna.Name = "btnAggiorna"
        Me.btnAggiorna.Size = New System.Drawing.Size(122, 35)
        Me.btnAggiorna.TabIndex = 5
        Me.btnAggiorna.Text = "&Aggiorna"
        Me.btnAggiorna.UseVisualStyleBackColor = True
        '
        'grdIva
        '
        Me.grdIva.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdIva.Location = New System.Drawing.Point(24, 21)
        Me.grdIva.Name = "grdIva"
        Me.grdIva.Size = New System.Drawing.Size(554, 255)
        Me.grdIva.TabIndex = 4
        '
        'Aliquote
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(595, 347)
        Me.Controls.Add(Me.btnAggiorna)
        Me.Controls.Add(Me.grdIva)
        Me.Name = "Aliquote"
        Me.ShowIcon = False
        Me.Text = "Aliquote"
        CType(Me.grdIva, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnAggiorna As System.Windows.Forms.Button
    Friend WithEvents grdIva As System.Windows.Forms.DataGridView
End Class
