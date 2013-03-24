Public Class Nazione

    Private Sub Nazione_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call ApriConn()
        Call ImpostaDS()
        grdNazione.DataSource = bsNaz

        Call NasColonna()

    End Sub
    Private Sub Aggiorna_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Aggiorna.Click
        Try

            bsNaz.EndEdit()
            Dim cb As New OleDb.OleDbCommandBuilder(DANaz)
            DANaz.Update(ds, "YYY")

           

        Catch ex As OleDb.OleDbException
            If InStr(ex.Message, "contiene record correlati") > 0 Then
                MsgBox("Impossibile eliminare il seguente dato perchè è contenuto in un'altra tabella", MsgBoxStyle.Exclamation, "Errore")
                Call ImpostaDS()
                grdNazione.DataSource = bsNaz

                Call NasColonna()
                Exit Sub
            Else
                MsgBox("Errore generico")
            End If

        End Try
    End Sub
    Private Sub NasColonna()
        grdNazione.DataSource = bsNaz
        grdNazione.Columns(0).Visible = False



    End Sub
    Private Sub grdNazione_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdNazione.CellContentClick

    End Sub
End Class