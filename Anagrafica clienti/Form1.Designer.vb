<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.MenuStrip2 = New System.Windows.Forms.MenuStrip
        Me.GestisciToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ClienteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ModalitàPagamentoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.NazioniToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.ArticoliToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CategorieArticoliToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator
        Me.UnitàDiMisuraToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AliquoteIVAToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator
        Me.OrdiniToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.StampaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ArticoliToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.MenuStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip2
        '
        Me.MenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.GestisciToolStripMenuItem, Me.StampaToolStripMenuItem})
        Me.MenuStrip2.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip2.Name = "MenuStrip2"
        Me.MenuStrip2.Size = New System.Drawing.Size(709, 24)
        Me.MenuStrip2.TabIndex = 2
        Me.MenuStrip2.Text = "MenuStrip2"
        '
        'GestisciToolStripMenuItem
        '
        Me.GestisciToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ClienteToolStripMenuItem, Me.ModalitàPagamentoToolStripMenuItem, Me.NazioniToolStripMenuItem, Me.ToolStripSeparator1, Me.ArticoliToolStripMenuItem, Me.CategorieArticoliToolStripMenuItem, Me.ToolStripSeparator2, Me.UnitàDiMisuraToolStripMenuItem, Me.AliquoteIVAToolStripMenuItem, Me.ToolStripSeparator3, Me.OrdiniToolStripMenuItem})
        Me.GestisciToolStripMenuItem.Name = "GestisciToolStripMenuItem"
        Me.GestisciToolStripMenuItem.Size = New System.Drawing.Size(61, 20)
        Me.GestisciToolStripMenuItem.Text = "Gestione"
        '
        'ClienteToolStripMenuItem
        '
        Me.ClienteToolStripMenuItem.Name = "ClienteToolStripMenuItem"
        Me.ClienteToolStripMenuItem.Size = New System.Drawing.Size(193, 22)
        Me.ClienteToolStripMenuItem.Text = "Clienti"
        '
        'ModalitàPagamentoToolStripMenuItem
        '
        Me.ModalitàPagamentoToolStripMenuItem.Name = "ModalitàPagamentoToolStripMenuItem"
        Me.ModalitàPagamentoToolStripMenuItem.Size = New System.Drawing.Size(193, 22)
        Me.ModalitàPagamentoToolStripMenuItem.Text = "Modalità di Pagamento"
        '
        'NazioniToolStripMenuItem
        '
        Me.NazioniToolStripMenuItem.Name = "NazioniToolStripMenuItem"
        Me.NazioniToolStripMenuItem.Size = New System.Drawing.Size(193, 22)
        Me.NazioniToolStripMenuItem.Text = "Nazioni"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(190, 6)
        '
        'ArticoliToolStripMenuItem
        '
        Me.ArticoliToolStripMenuItem.Name = "ArticoliToolStripMenuItem"
        Me.ArticoliToolStripMenuItem.Size = New System.Drawing.Size(193, 22)
        Me.ArticoliToolStripMenuItem.Text = "Articoli"
        '
        'CategorieArticoliToolStripMenuItem
        '
        Me.CategorieArticoliToolStripMenuItem.Name = "CategorieArticoliToolStripMenuItem"
        Me.CategorieArticoliToolStripMenuItem.Size = New System.Drawing.Size(193, 22)
        Me.CategorieArticoliToolStripMenuItem.Text = "Categorie Articoli"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(190, 6)
        '
        'UnitàDiMisuraToolStripMenuItem
        '
        Me.UnitàDiMisuraToolStripMenuItem.Name = "UnitàDiMisuraToolStripMenuItem"
        Me.UnitàDiMisuraToolStripMenuItem.Size = New System.Drawing.Size(193, 22)
        Me.UnitàDiMisuraToolStripMenuItem.Text = "Unità di Misura"
        '
        'AliquoteIVAToolStripMenuItem
        '
        Me.AliquoteIVAToolStripMenuItem.Name = "AliquoteIVAToolStripMenuItem"
        Me.AliquoteIVAToolStripMenuItem.Size = New System.Drawing.Size(193, 22)
        Me.AliquoteIVAToolStripMenuItem.Text = "Aliquote IVA"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(190, 6)
        '
        'OrdiniToolStripMenuItem
        '
        Me.OrdiniToolStripMenuItem.Name = "OrdiniToolStripMenuItem"
        Me.OrdiniToolStripMenuItem.Size = New System.Drawing.Size(193, 22)
        Me.OrdiniToolStripMenuItem.Text = "Ordini"
        '
        'StampaToolStripMenuItem
        '
        Me.StampaToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ArticoliToolStripMenuItem1})
        Me.StampaToolStripMenuItem.Name = "StampaToolStripMenuItem"
        Me.StampaToolStripMenuItem.Size = New System.Drawing.Size(55, 20)
        Me.StampaToolStripMenuItem.Text = "Stampa"
        '
        'ArticoliToolStripMenuItem1
        '
        Me.ArticoliToolStripMenuItem1.Name = "ArticoliToolStripMenuItem1"
        Me.ArticoliToolStripMenuItem1.Size = New System.Drawing.Size(152, 22)
        Me.ArticoliToolStripMenuItem1.Text = "Articoli"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(709, 621)
        Me.Controls.Add(Me.MenuStrip2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Name = "Form1"
        Me.Text = "Gestione aziendale"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip2.ResumeLayout(False)
        Me.MenuStrip2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip2 As System.Windows.Forms.MenuStrip
    Friend WithEvents GestisciToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ClienteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ModalitàPagamentoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NazioniToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ArticoliToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents CategorieArticoliToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents UnitàDiMisuraToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AliquoteIVAToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StampaToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ArticoliToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents OrdiniToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
