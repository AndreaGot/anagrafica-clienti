Public Class Form1

    Private Sub EsciToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        End

    End Sub
    Private Sub ClienteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClienteToolStripMenuItem.Click
        Dim x As New Cliente
        x.MdiParent = Me
        x.Show()

    End Sub
    Private Sub Modalit‡PagamentoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Modalit‡PagamentoToolStripMenuItem.Click
        Dim x As New Pagamento
        x.MdiParent = Me
        x.Show()

    End Sub
    Private Sub NazioniToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NazioniToolStripMenuItem.Click
        Dim x As New Nazione
        x.MdiParent = Me
        x.Show()
    End Sub
    Private Sub Unit‡DiMisuraToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Unit‡DiMisuraToolStripMenuItem.Click
        Dim x As New UM
        x.MdiParent = Me
        x.Show()
    End Sub
    Private Sub AliquoteIVAToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AliquoteIVAToolStripMenuItem.Click
        Dim x As New Aliquote
        x.MdiParent = Me
        x.Show()
    End Sub
    Private Sub CategorieArticoliToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CategorieArticoliToolStripMenuItem.Click
        Dim X As New Categorie
        X.MdiParent = Me
        X.Show()
    End Sub
    Private Sub ArticoliToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ArticoliToolStripMenuItem.Click
        Dim x As New Articoli
        x.MdiParent = Me
        x.Show()
    End Sub
    Private Sub ArticoliToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ArticoliToolStripMenuItem1.Click
        Dim x As New StampArt
        x.MdiParent = Me
        x.Show()
    End Sub
    Private Sub OrdiniToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OrdiniToolStripMenuItem.Click
        Dim x As New Ordini
        x.MdiParent = Me
        x.Show()
    End Sub


End Class
