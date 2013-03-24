
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data
Imports System.Data.oledb

Public Class Articoli

    Private BSArticoli As New BindingSource
    Dim Aggiungi As Boolean
    Dim VecchioArt As String


    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Call ChiudiConn()
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        dtUMART.Clear()
        DAUMART.Dispose()
        dtUMART.Dispose()

        dtIVAART.Clear()
        DAIVAART.Dispose()
        dtIVAART.Dispose()

        dtCATARTART.Clear()
        DACATARTART.Dispose()
        dtCATARTART.Dispose()

        Call ApriConn()
        If Primavolta = False Then

        End If
        Call ArticoliDaDt()

        Primavolta = True



        Call Disabilitasalva()
        Call DisabilitaText()
        Call Abilitanavigatore()

        Call AssociaCampi()
        Call VisualizzaPosizione()

    End Sub
    
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        BSArticoli.AddNew()

        Aggiungi = True
        Call AbilitaSalva()
        Call AbilitaText()
        Call DisabilitaNavigatore()

        txtPosizione.Text = "Aggiunta"

        txtdescrizione.Focus()
    End Sub
    Private Sub btnSalva_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalva.Click
        Dim descrizione As String
        Dim FKUM As String
        Dim FKIVA As String
        Dim FKCATART As String
        Dim prlistino As Single
        Dim qtasconto1 As Single
        Dim qtasconto2 As Single
        Dim sconto1 As Single
        Dim sconto2 As Single
        Dim scontoA As Single
        Dim scontoB As Single
        Dim scortamax As Integer
        Dim scortamin As Integer
        Dim lottoeco As Integer
        Dim costomedio As Single
        Dim ultimopracq As Single
        Dim FK As String



        'Try
        FK = txtFK.text
        descrizione = txtdescrizione.Text
        FKUM = cmbUM.SelectedValue
        FKIVA = cmbIVA.SelectedValue
        FKCATART = cmbcatart.SelectedValue
        prlistino = txtprlistino.Text
        qtasconto1 = txtqtasconto1.Text
        qtasconto2 = txtqtasconto2.Text
        sconto1 = txtsconto.Text
        sconto2 = txtsconto2.Text
        scontoA = txtscontoA.Text
        scontoB = txtscontoB.Text
        scortamax = txtscortamax.Text
        scortamin = txtscortamin.Text
        lottoeco = txtlottoeco.Text
        costomedio = txtcostomedio.Text
        ultimopracq = txtultpracq.Text




        If Aggiungi Then
            cmdAggiungiArt.Parameters.Add(New OleDb.OleDbParameter("PKarticolo", FK))
            cmdAggiungiArt.Parameters.Add(New OleDb.OleDbParameter("DesArticolo", descrizione))
            cmdAggiungiArt.Parameters.Add(New OleDb.OleDbParameter("UM", FKUM))
            cmdAggiungiArt.Parameters.Add(New OleDb.OleDbParameter("FKCatMag", FKCATART))
            cmdAggiungiArt.Parameters.Add(New OleDb.OleDbParameter("Prezzolistino", prlistino))
            cmdAggiungiArt.Parameters.Add(New OleDb.OleDbParameter("QtaSconto1", qtasconto1))
            cmdAggiungiArt.Parameters.Add(New OleDb.OleDbParameter("Sconto1", sconto1))
            cmdAggiungiArt.Parameters.Add(New OleDb.OleDbParameter("Qtasconto2", qtasconto2))
            cmdAggiungiArt.Parameters.Add(New OleDb.OleDbParameter("Sconto2", sconto2))
            cmdAggiungiArt.Parameters.Add(New OleDb.OleDbParameter("ScontoA", scontoA))
            cmdAggiungiArt.Parameters.Add(New OleDb.OleDbParameter("ScontoB", scontoB))
            cmdAggiungiArt.Parameters.Add(New OleDb.OleDbParameter("FKIVA", FKIVA))
            cmdAggiungiArt.Parameters.Add(New OleDb.OleDbParameter("ScortaMax", scortamax))
            cmdAggiungiArt.Parameters.Add(New OleDb.OleDbParameter("ScortaMin", scortamin))
            cmdAggiungiArt.Parameters.Add(New OleDb.OleDbParameter("LottoEco", lottoeco))
            cmdAggiungiArt.Parameters.Add(New OleDb.OleDbParameter("CostoMedio", costomedio))
            cmdAggiungiArt.Parameters.Add(New OleDb.OleDbParameter("UltPrezzoAcq", ultimopracq))
            cmdAggiungiArt.ExecuteNonQuery()
            cmdAggiungiArt.Parameters.Clear()
            BSArticoli.Position = BSArticoli.Count - 1
            Call VisualizzaPosizione()
        Else

            cmdSalvaArt.Parameters.Add(New OleDb.OleDbParameter("PKarticolo", FK))
            cmdSalvaArt.Parameters.Add(New OleDb.OleDbParameter("DesArticolo", descrizione))
            cmdSalvaArt.Parameters.Add(New OleDb.OleDbParameter("UM", FKUM))
            cmdSalvaArt.Parameters.Add(New OleDb.OleDbParameter("FKCatMag", FKCATART))
            cmdSalvaArt.Parameters.Add(New OleDb.OleDbParameter("Prezzolistino", prlistino))
            cmdSalvaArt.Parameters.Add(New OleDb.OleDbParameter("QtaSconto1", qtasconto1))
            cmdSalvaArt.Parameters.Add(New OleDb.OleDbParameter("Sconto1", sconto1))
            cmdSalvaArt.Parameters.Add(New OleDb.OleDbParameter("Qtasconto2", qtasconto2))
            cmdSalvaArt.Parameters.Add(New OleDb.OleDbParameter("sconto2", sconto2))
            cmdSalvaArt.Parameters.Add(New OleDb.OleDbParameter("ScontoA", scontoA))
            cmdSalvaArt.Parameters.Add(New OleDb.OleDbParameter("ScontoB", scontoB))
            cmdSalvaArt.Parameters.Add(New OleDb.OleDbParameter("FKIVA", FKIVA))
            cmdSalvaArt.Parameters.Add(New OleDb.OleDbParameter("ScortaMax", scortamax))
            cmdSalvaArt.Parameters.Add(New OleDb.OleDbParameter("ScortaMin", scortamin))
            cmdSalvaArt.Parameters.Add(New OleDb.OleDbParameter("LottoEco", lottoeco))
            cmdSalvaArt.Parameters.Add(New OleDb.OleDbParameter("CostoMedio", costomedio))
            cmdSalvaArt.Parameters.Add(New OleDb.OleDbParameter("UltPrezzoAcq", ultimopracq))
            cmdSalvaArt.Parameters.Add(New OleDb.OleDbParameter("PKarticolo", VecchioArt))
            cmdSalvaArt.ExecuteNonQuery()
            cmdSalvaArt.Parameters.Clear()
        End If
        Aggiungi = False
        Call Disabilitasalva()
        Call Disabilitatext()
        Call Abilitanavigatore()

        dtART.Clear()
        DAART.Fill(dtART)


        Call VisualizzaPosizione()

        'Catch x As InvalidCastException
        '    '    MsgBox("Il record appena inserito non è modificabile")
        'End Try
    End Sub
    Private Sub btnElimina_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnElimina.Click
        Dim Pos As Integer
        Dim CoDaElim As String
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult


        Pos = BSArticoli.Position

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


            CoDaElim = txtFK.Text

            cmdEliminaArt.Parameters.Add(New OleDb.OleDbParameter("PKCliente", CoDaElim))
            cmdEliminaArt.ExecuteNonQuery()
            cmdEliminaArt.Parameters.Clear()
        End If

        dtART.Clear()
        DAART.Fill(dtART)

        BSArticoli.Position = 0

        Call VisualizzaPosizione()

    End Sub
    Private Sub btnModifica_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModifica.Click
        Aggiungi = False

        If BSArticoli.Position < 0 Then
            MsgBox("Nessun record presente")
            Exit Sub
        End If

        VecchioArt = txtFK.Text

        Call AbilitaSalva()
        Call AbilitaText()

        Call DisabilitaNavigatore()
    End Sub
    Private Sub btnFine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFine.Click
        Me.Close()
    End Sub
    Private Sub btnAnnulla_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAnnulla.Click
        BSArticoli.CancelEdit()

        Call Disabilitasalva()
        Call DisabilitaText()
        Call Abilitanavigatore()
        Call VisualizzaPosizione()

    End Sub
    
    Private Sub AssociaCampi()
        BSArticoli.DataSource = dtART

        txtFK.DataBindings.Add("text", BSArticoli, "PKArticolo")
        txtdescrizione.DataBindings.Add("text", BSArticoli, "DesArticolo")
        txtprlistino.DataBindings.Add("text", BSArticoli, "PrezzoListino")
        txtqtasconto1.DataBindings.Add("text", BSArticoli, "QtaSconto1")
        txtqtasconto2.DataBindings.Add("text", BSArticoli, "QtaSconto2")
        txtsconto.DataBindings.Add("text", BSArticoli, "Sconto1")
        txtsconto2.DataBindings.Add("text", BSArticoli, "Sconto2")
        txtscontoA.DataBindings.Add("text", BSArticoli, "ScontoA")
        txtscontoB.DataBindings.Add("text", BSArticoli, "ScontoB")
        txtscortamax.DataBindings.Add("text", BSArticoli, "ScortaMax")
        txtscortamin.DataBindings.Add("text", BSArticoli, "ScortaMin")
        txtlottoeco.DataBindings.Add("text", BSArticoli, "Lottoeco")
        txtcostomedio.DataBindings.Add("text", BSArticoli, "CostoMedio")
        txtultpracq.DataBindings.Add("text", BSArticoli, "UltPrezzoAcq")

        cmbcatart.DataBindings.Add("SelectedValue", BSArticoli, "FKCatMag")
        cmbcatart.DataSource = dtCATARTART
        cmbcatart.DisplayMember = "Descrizione"
        cmbcatart.ValueMember = "PKCatMag"

        cmbiva.DataBindings.Add("SelectedValue", BSArticoli, "FKIVA")
        cmbiva.DataSource = dtIVAART
        cmbiva.DisplayMember = "Aliquota"
        cmbiva.ValueMember = "PKIVA"

        cmbUM.DataBindings.Add("SelectedValue", BSArticoli, "UM")
        cmbUM.DataSource = dtUMART
        cmbUM.DisplayMember = "Descrizione"
        cmbUM.ValueMember = "PKUM"

    End Sub

    Private Sub BtnScegli_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnScegli.Click
        'Dim TestoQuery As String
        'Dim NomeForm As String
        'Dim IDTrovato As String
        'Dim Pos As Integer
        'Dim CodAnnulla As String


        'CodAnnulla = txtFK.Text

        'TestoQuery = "SELECT * FROM tblArticoli"
        'NomeForm = "Selezione Articolo"
        'NasColonna = False

        'RicercaDati(Me, TestoQuery, NomeForm, IDTrovato, NasColonna)

        'If IDTrovato = "" Then
        '    IDTrovato = CodAnnulla
        'Else
        '    Pos = BSArticoli.Find("PKArticolo", IDTrovato)
        'End If


        'BSArticoli.Position = Pos
        'Call VisualizzaPosizione()
    End Sub

    '------------ ABILITAZIONE E DISABILITAZIONE -------------

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

    Private Sub Disabilitatext()


        grpDati.Enabled = False
    End Sub
    Private Sub AbilitaText()
        grpDati.Enabled = True

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

    '----------- NAVIGATORE ---------

    Private Sub btnPrimo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrimo.Click
        BSArticoli.Position = 0

        Call VisualizzaPosizione()
    End Sub
    Private Sub btnPrima_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrima.Click
        BSArticoli.Position = BSArticoli.Position - 1

        Call VisualizzaPosizione()
    End Sub
    Private Sub btnDopo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDopo.Click
        BSArticoli.Position = BSArticoli.Position + 1

        Call VisualizzaPosizione()
    End Sub
    Private Sub btnUltimo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUltimo.Click
        BSArticoli.Position = BSArticoli.Count - 1

        Call VisualizzaPosizione()
    End Sub

    Private Sub VisualizzaPosizione()
        Dim totale As Integer
        Dim posizione As Integer
        Dim stringaposizione As String

        totale = BSArticoli.Count
        posizione = BSArticoli.Position + 1
        stringaposizione = CStr(posizione) & " di " & CStr(totale)
        txtPosizione.Text = stringaposizione

    End Sub

    Private Sub grpDati_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles grpDati.Enter

    End Sub

    Private Sub txtFK_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFK.TextChanged

    End Sub

    Private Sub txtscontoA_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtscontoA.TextChanged

    End Sub
End Class