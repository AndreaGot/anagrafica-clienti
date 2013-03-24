Public Class Aliquote

    Private Sub Aggiorna_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAggiorna.Click
        Try

            bsIVA.EndEdit()
            Dim cb As New OleDb.OleDbCommandBuilder(DAIVA)
            DAIVA.Update(DSIVA, "YYY")



        Catch ex As OleDb.OleDbException
            If InStr(ex.Message, "contiene record correlati") > 0 Then
                MsgBox("Impossibile eliminare il seguente dato perchè è contenuto in un'altra tabella", MsgBoxStyle.Exclamation, "Errore")
                Call ImpostaDS()
                grdIva.DataSource = bsIVA


                Exit Sub
            Else
                MsgBox("Errore generico")
            End If

        End Try
    End Sub


    Private Sub Aliquote_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call ApriConn()
        Call ImpostaDSIVA()
        grdIva.DataSource = bsIVA
    End Sub

End Class