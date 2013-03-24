<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UM
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
        Me.grdUnitaMisura = New System.Windows.Forms.DataGridView
        CType(Me.grdUnitaMisura, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnAggiorna
        '
        Me.btnAggiorna.Location = New System.Drawing.Point(330, 237)
        Me.btnAggiorna.Name = "btnAggiorna"
        Me.btnAggiorna.Size = New System.Drawing.Size(122, 35)
        Me.btnAggiorna.TabIndex = 3
        Me.btnAggiorna.Text = "&Aggiorna"
        Me.btnAggiorna.UseVisualStyleBackColor = True
        '
        'grdUnitaMisura
        '
        Me.grdUnitaMisura.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdUnitaMisura.Location = New System.Drawing.Point(12, 12)
        Me.grdUnitaMisura.Name = "grdUnitaMisura"
        Me.grdUnitaMisura.Size = New System.Drawing.Size(455, 219)
        Me.grdUnitaMisura.TabIndex = 2
        '
        'UM
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(488, 310)
        Me.Controls.Add(Me.btnAggiorna)
        Me.Controls.Add(Me.grdUnitaMisura)
        Me.Name = "UM"
        Me.ShowIcon = False
        Me.Text = "UM"
        CType(Me.grdUnitaMisura, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnAggiorna As System.Windows.Forms.Button
    Friend WithEvents grdUnitaMisura As System.Windows.Forms.DataGridView
End Class
