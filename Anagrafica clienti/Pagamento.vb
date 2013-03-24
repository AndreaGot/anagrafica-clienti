Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data
Imports System.Data.oledb

Public Class Pagamento

    Private BSPagamento As New BindingSource
    Dim Aggiungi As Boolean


    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Call ChiudiConnNaz()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'dt.Clear()
        'DA.Dispose()
        'dt.Dispose()

        dtPag.Clear()
        DAPag.Dispose()
        dtPag.Dispose()

        Call ApriConn()
        If Primavolta = False Then

        End If
        Call PagDaDt()

        Primavolta = True



        Call Disabilitasalva()
        Call DisabilitaText()
        Call Abilitanavigatore()

        Call AssociaCampi()
        Call VisualizzaPosizione()

    End Sub
    
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        BSPagamento.AddNew()

        Aggiungi = True
        Call AbilitaSalva()
        Call AbilitaText()
        Call DisabilitaNavigatore()

        txtPosizione.Text = "Aggiunta"

        txtPag.Focus()
    End Sub
    Private Sub btnSalva_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalva.Click
        Dim FK As Integer
        Dim pos As Integer
        Dim TipoPag As String
        Dim Giorni As Integer
        Dim Rate As Integer
        Dim GiorniRate As Integer
        Dim Sconto As Single
        Dim SpeseBanca As Single
        Dim FineMese As Boolean

        TipoPag = txtPag.Text
        Giorni = Val(txtGiorni.Text)
        Rate = Val(txtRate.Text)
        GiorniRate = Val(txtGiorniRate.Text)
        Sconto = Val(txtSconto.Text)
        SpeseBanca = Val(txtSpeseBanca.Text)
        FineMese = chkFineMese.Checked

        Try
            '----controlli----'
            '---/controlli---'

            If Aggiungi Then

                cmdAggiungiPag.Parameters.Add(New OleDb.OleDbParameter("Modalita", TipoPag))
                cmdAggiungiPag.Parameters.Add(New OleDb.OleDbParameter("Giorni", Giorni))
                cmdAggiungiPag.Parameters.Add(New OleDb.OleDbParameter("Rate", Rate))
                cmdAggiungiPag.Parameters.Add(New OleDb.OleDbParameter("GiorniRate", GiorniRate))
                cmdAggiungiPag.Parameters.Add(New OleDb.OleDbParameter("Sconto", Sconto))
                cmdAggiungiPag.Parameters.Add(New OleDb.OleDbParameter("SpeseBanca", SpeseBanca))
                ' cmdAggiungiPag.Parameters.Add(New OleDb.OleDbParameter("FineMese", FineMese))
                cmdAggiungiPag.ExecuteNonQuery()
                cmdAggiungiPag.Parameters.Clear()
                BSPagamento.Position = BSPagamento.Count - 1
                Call VisualizzaPosizione()
            Else
                pos = BSPagamento.Position
                FK = CInt(dtPag.Rows(pos).Item("PKPagamento"))
                cmdSalvaPag.Parameters.Add(New OleDb.OleDbParameter("Modalita", TipoPag))
                cmdSalvaPag.Parameters.Add(New OleDb.OleDbParameter("Giorni", Giorni))
                cmdSalvaPag.Parameters.Add(New OleDb.OleDbParameter("Rate", Rate))
                cmdSalvaPag.Parameters.Add(New OleDb.OleDbParameter("GiorniRate", GiorniRate))
                cmdSalvaPag.Parameters.Add(New OleDb.OleDbParameter("Sconto", Sconto))
                cmdSalvaPag.Parameters.Add(New OleDb.OleDbParameter("SpeseBanca", SpeseBanca))
                cmdSalvaPag.Parameters.Add(New OleDb.OleDbParameter("FineMese", FineMese))
                cmdSalvaPag.Parameters.Add(New OleDb.OleDbParameter("PKPagamento", FK))
                cmdSalvaPag.ExecuteNonQuery()
                cmdSalvaPag.Parameters.Clear()
            End If
            Aggiungi = False
            Call Disabilitasalva()
            Call DisabilitaText()
            Call Abilitanavigatore()

            dtPag.Clear()
            DAPag.Fill(dtPag)


            Call VisualizzaPosizione()

        Catch x As InvalidCastException
            '    MsgBox("Il record appena inserito non è modificabile")
        End Try
    End Sub
    Private Sub btnElimina_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnElimina.Click


        Dim Pos As Integer
        Dim FK As Integer
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult

        Try

       
            Pos = BSPagamento.Position

            If Pos = -1 Then
                MsgBox("Non ci sono record eliminabili")
                Exit Sub
            End If

            msg = "Sei sicuro di volerlo eliminare?"   ' Define message.
            style = MsgBoxStyle.DefaultButton2 Or _
               MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo
            title = "Eliminazione Record"   ' Define title.
            ' Display message.
            response = MsgBox(msg, style, title)
            If response = MsgBoxResult.Yes Then   ' User chose Yes.


                FK = CInt((dtPag.Rows(Pos).Item("PKPagamento")))

                cmdEliminaPag.Parameters.Add(New OleDb.OleDbParameter("PKPagamento", FK))
                cmdEliminaPag.ExecuteNonQuery()
                cmdEliminaPag.Parameters.Clear()
            End If

            dtPag.Clear()
            DAPag.Fill(dtPag)

            BSPagamento.Position = 0

            Call VisualizzaPosizione()
        Catch ex As OleDb.OleDbException
            If InStr(ex.Message, "contiene record correlati") > 0 Then
                MsgBox("record utilizzato")
            Else
                MsgBox("erro generico")
            End If
        End Try

    End Sub
    Private Sub btnModifica_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModifica.Click
        Aggiungi = False

        If BSPagamento.Position < 0 Then
            MsgBox("Nessun record presente")
            Exit Sub
        End If

        Call AbilitaSalva()
        Call AbilitaText()

        Call DisabilitaNavigatore()
    End Sub
    Private Sub btnFine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFine.Click
        Me.Close()
    End Sub
    Private Sub btnAnnulla_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAnnulla.Click
        BSPagamento.CancelEdit()

        Call Disabilitasalva()
        Call DisabilitaText()
        Call Abilitanavigatore()
        Call VisualizzaPosizione()

    End Sub

    Private Sub AssociaCampi()
        BSPagamento.DataSource = dtPag

        txtID.DataBindings.Add("text", BSPagamento, "PKPagamento")
        txtPag.DataBindings.Add("text", BSPagamento, "Modalita")
        txtGiorni.DataBindings.Add("text", BSPagamento, "Giorni")
        txtRate.DataBindings.Add("text", BSPagamento, "Rate")
        txtGiorniRate.DataBindings.Add("text", BSPagamento, "GiorniRate")
        txtSconto.DataBindings.Add("text", BSPagamento, "Sconto")
        txtSpeseBanca.DataBindings.Add("text", BSPagamento, "SpeseBanca")
        ' chkFineMese.DataBindings.Add("checked", BSPagamento, "FineMese")

    End Sub
    
    Private Sub BtnScegli_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnScegli.Click
        Dim TestoQuery As String
        Dim NomeForm As String
        Dim IDTrovato As String
        Dim Pos As Integer
        Dim IDAnnulla As String





        TestoQuery = "SELECT * FROM tblPagamento"
        NomeForm = "Selezione Tipo di Pagamento"
        NasColonna = True
        IDAnnulla = txtID.Text

        RicercaDati(Me, TestoQuery, NomeForm, IDTrovato, NasColonna)

        If IDTrovato = "" Then
            IDTrovato = IDannulla
        Else
            Pos = BSPagamento.Find("PKPagamento", CInt(IDTrovato))

        End If


        BSPagamento.Position = Pos
        Call VisualizzaPosizione()
    End Sub

    '---------- ABILITAZIONE E DISABILITAZIONE -----------
    
    Private Sub DisabilitaText()
        txtPag.Enabled = False
        txtGiorni.Enabled = False
        txtRate.Enabled = False
        txtGiorniRate.Enabled = False
        txtSconto.Enabled = False
        txtSpeseBanca.Enabled = False
        chkFineMese.Enabled = False
    End Sub
    Private Sub AbilitaText()
        txtPag.Enabled = True
        txtGiorni.Enabled = True
        txtRate.Enabled = True
        txtGiorniRate.Enabled = True
        txtSconto.Enabled = True
        txtSpeseBanca.Enabled = True
        chkFineMese.Enabled = True

    End Sub

    Private Sub Abilitanavigatore()
        btnPrimo.Enabled = True
        btnUltimo.Enabled = True
        btnPrima.Enabled = True
        btnDopo.Enabled = True
    End Sub
    Private Sub DisabilitaNavigatore()
        btnPrimo.Enabled = False
        btnUltimo.Enabled = False
        btnPrima.Enabled = False
        btnDopo.Enabled = False
    End Sub

    Private Sub AbilitaSalva()
        grpConferma.Enabled = True
        grpOper.Enabled = False
        btnFine.Enabled = False
    End Sub
    Private Sub Disabilitasalva()
        grpConferma.Enabled = False
        grpOper.Enabled = True
        btnFine.Enabled = True
    End Sub

    '--------- NAVIGATORE ----------

    Private Sub btnPrimo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrimo.Click
        BSPagamento.Position = 0

        Call VisualizzaPosizione()
    End Sub
    Private Sub btnPrima_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrima.Click
        BSPagamento.Position = BSPagamento.Position - 1

        Call VisualizzaPosizione()
    End Sub
    Private Sub btnDopo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDopo.Click
        BSPagamento.Position = BSPagamento.Position + 1

        Call VisualizzaPosizione()
    End Sub
    Private Sub btnUltimo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUltimo.Click
        BSPagamento.Position = BSPagamento.Count - 1

        Call VisualizzaPosizione()
    End Sub

    Private Sub VisualizzaPosizione()
        Dim totale As Integer
        Dim posizione As Integer
        Dim stringaposizione As String

        totale = BSPagamento.Count
        posizione = BSPagamento.Position + 1
        stringaposizione = CStr(posizione) & " di " & CStr(totale)
        txtPosizione.Text = stringaposizione

    End Sub
End Class

