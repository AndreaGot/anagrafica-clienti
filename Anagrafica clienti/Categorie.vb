Public Class Categorie
    Private Sub Aggiorna_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAggiorna.Click
        Try

            bsCATART.EndEdit()
            Dim cb As New OleDb.OleDbCommandBuilder(DACATART)
            DACATART.Update(DSCATART, "YYY")



        Catch ex As OleDb.OleDbException
            If InStr(ex.Message, "contiene record correlati") > 0 Then
                MsgBox("Impossibile eliminare il seguente dato perchè è contenuto in un'altra tabella", MsgBoxStyle.Exclamation, "Errore")
                Call ImpostaDS()
                grdCatArt.DataSource = bsCATART


                Exit Sub
            Else
                MsgBox("Errore generico")
            End If

        End Try
    End Sub
    Private Sub Categorie_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call ApriConn()
        Call ImpostaDSCATART()
        grdCatArt.DataSource = bsCATART
    End Sub
    Private Sub grdCatArt_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdCatArt.CellContentClick

    End Sub
End Class