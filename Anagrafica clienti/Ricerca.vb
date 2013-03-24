Public Class Ricerca

    Dim ColClick As Integer
    Private Sub Ricerca_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        grdRicerca.DataSource = BSRicerca


        If NasColonna Then
            Call NascondiCampi()
            txtCampi.Text = grdRicerca.Columns(1).HeaderText
            lblRicerca.Text = grdRicerca.Columns(1).HeaderText
        Else

            txtCampi.Text = grdRicerca.Columns(0).HeaderText
            lblRicerca.Text = grdRicerca.Columns(0).HeaderText
        End If

    End Sub
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim a As String
        Dim b As Integer

        b = grdRicerca.CurrentCell.RowIndex
        a = grdRicerca.Rows(b).Cells(0).Value
        txtID.Text = a.ToString

        Me.Close()
    End Sub
    Private Sub txtRicerca_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRicerca.TextChanged
        cmdLeggiRicerca.CommandText = txtQuery.Text & " WHERE " & txtCampi.Text & " like '%" & txtRicerca.Text & "%'"

        DSRicerca.Clear()

        DARicerca.Fill(DSRicerca, "Ric")


    End Sub
    Private Sub NascondiCampi()
        grdRicerca.Columns(0).Visible = False
    End Sub
    Private Sub grdRicerca_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles grdRicerca.ColumnHeaderMouseClick

        txtRicerca.Text = ""
        ColClick = e.ColumnIndex

        txtCampi.Text = grdRicerca.Columns(ColClick).HeaderText
        lblRicerca.Text = grdRicerca.Columns(ColClick).HeaderText

    End Sub
    Private Sub btnannulla_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnannulla.Click
        Me.Close()
    End Sub

    Private Sub grdRicerca_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdRicerca.CellContentClick

    End Sub
End Class