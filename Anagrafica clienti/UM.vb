Public Class UM

    Private Sub Aggiorna_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAggiorna.Click
        Try

            bsUM.EndEdit()
            Dim cb As New OleDb.OleDbCommandBuilder(DAUM)
            DAUM.Update(DSUM, "YYY")



        Catch ex As OleDb.OleDbException
            If InStr(ex.Message, "contiene record correlati") > 0 Then
                MsgBox("Impossibile eliminare il seguente dato perchè è contenuto in un'altra tabella", MsgBoxStyle.Exclamation, "Errore")
                Call ImpostaDS()
                grdUnitaMisura.DataSource = bsUM


                Exit Sub
            Else
                MsgBox("Errore generico")
            End If

        End Try
    End Sub
    Private Sub UM_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call ApriConn()
        Call ImpostaDSUM()
        grdUnitaMisura.DataSource = bsUM


    End Sub

End Class