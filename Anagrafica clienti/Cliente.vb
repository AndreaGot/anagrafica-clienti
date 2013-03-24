Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data
Imports System.Data.oledb


Public Class Cliente


    Private BSClienti As New BindingSource
    Dim Aggiungi As Boolean


    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Call ChiudiConn()
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        
        dtNazCli.Clear()
        DANazCli.Dispose()
        dtNazCli.Dispose()
        dt.Clear()
        DA.Dispose()
        dt.Dispose()

        dtPagCli.Clear()
        DAPagCli.Dispose()
        dtPagCli.Dispose()

        Call ApriConn()
        If Primavolta = False Then

        End If
        Call ClientiDaDt()

        Primavolta = True



        Call Disabilitasalva()
        Call DisabilitaText()
        Call Abilitanavigatore()

        Call AssociaCampi()
        Call VisualizzaPosizione()

    End Sub
    
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        BSClienti.AddNew()

        Aggiungi = True
        Call AbilitaSalva()
        Call AbilitaText()
        Call DisabilitaNavigatore()

        txtPosizione.Text = "Aggiunta"

        txtNome.Focus()
    End Sub
    Private Sub btnSalva_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalva.Click
        Dim FK As Integer
        Dim FKNazione As Integer
        Dim FKPagamento As Integer
        Dim Pos As Integer
        Dim Nome As String
        Dim Cognome As String
        Dim RagioneSociale As String
        Dim PartitaIVA As String
        Dim CodFisc As String
        Dim Via As String
        Dim CAP As String
        Dim Citta As String

        RagioneSociale = txtRagSoc.Text
        Nome = txtNome.Text
        Cognome = txtCognome.Text
        PartitaIVA = txtPartIVA.Text
        CodFisc = txtCodFisc.Text
        Via = txtVia.Text
        CAP = txtCAP.Text
        Citta = txtCitta.Text
        FKPagamento = cmbPag.SelectedValue
        FKNazione = cmbNazione.SelectedValue


        Try
            '----controlli----'
            '---/controlli---'

            If Aggiungi Then

                cmdAggiungi.Parameters.Add(New OleDb.OleDbParameter("RagioneSociale", RagioneSociale))
                cmdAggiungi.Parameters.Add(New OleDb.OleDbParameter("Nome", Nome))
                cmdAggiungi.Parameters.Add(New OleDb.OleDbParameter("Cognome", Cognome))
                cmdAggiungi.Parameters.Add(New OleDb.OleDbParameter("PartitaIVA", PartitaIVA))
                cmdAggiungi.Parameters.Add(New OleDb.OleDbParameter("CodFiscale", CodFisc))
                cmdAggiungi.Parameters.Add(New OleDb.OleDbParameter("Via", Via))
                cmdAggiungi.Parameters.Add(New OleDb.OleDbParameter("CAP", CAP))
                cmdAggiungi.Parameters.Add(New OleDb.OleDbParameter("Citta", Citta))
                cmdAggiungi.Parameters.Add(New OleDb.OleDbParameter("FKNazione", FKNazione))
                cmdAggiungi.Parameters.Add(New OleDb.OleDbParameter("FKPagamento", FKPagamento))
                cmdAggiungi.ExecuteNonQuery()
                cmdAggiungi.Parameters.Clear()
                BSClienti.Position = BSClienti.Count - 1
                Call VisualizzaPosizione()
            Else
                Pos = BSClienti.Position
                FK = CInt(dt.Rows(Pos).Item("PKCliente"))
                cmdSalva.Parameters.Add(New OleDb.OleDbParameter("RagioneSociale", RagioneSociale))
                cmdSalva.Parameters.Add(New OleDb.OleDbParameter("Nome", Nome))
                cmdSalva.Parameters.Add(New OleDb.OleDbParameter("Cognome", Cognome))
                cmdSalva.Parameters.Add(New OleDb.OleDbParameter("PartitaIVA", PartitaIVA))
                cmdSalva.Parameters.Add(New OleDb.OleDbParameter("CodFiscale", CodFisc))
                cmdSalva.Parameters.Add(New OleDb.OleDbParameter("Via", Via))
                cmdSalva.Parameters.Add(New OleDb.OleDbParameter("CAP", CAP))
                cmdSalva.Parameters.Add(New OleDb.OleDbParameter("Citta", Citta))
                cmdSalva.Parameters.Add(New OleDb.OleDbParameter("FKNazione", FKNazione))
                cmdSalva.Parameters.Add(New OleDb.OleDbParameter("FKPagamento", FKPagamento))
                cmdSalva.Parameters.Add(New OleDb.OleDbParameter("PKCliente", FK))
                cmdSalva.ExecuteNonQuery()
                cmdSalva.Parameters.Clear()
            End If
            Aggiungi = False
            Call Disabilitasalva()
            Call DisabilitaText()
            Call Abilitanavigatore()

            dt.Clear()
            DA.Fill(dt)


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


        Pos = BSClienti.Position

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


            FK = CInt((dt.Rows(Pos).Item("PKCliente")))

            cmdElimina.Parameters.Add(New OleDb.OleDbParameter("PKCliente", FK))
            cmdElimina.ExecuteNonQuery()
            cmdElimina.Parameters.Clear()
        End If

        dt.Clear()
        DA.Fill(dt)

        BSClienti.Position = 0

        Call VisualizzaPosizione()

    End Sub
    Private Sub btnModifica_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModifica.Click
        Aggiungi = False

        If BSClienti.Position < 0 Then
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
        BSClienti.CancelEdit()

        Call Disabilitasalva()
        Call DisabilitaText()
        Call Abilitanavigatore()
        Call VisualizzaPosizione()

    End Sub
    
    Private Sub AssociaCampi()
        BSClienti.DataSource = dt
        txtid.DataBindings.Add("text", BSClienti, "PKCliente")
        txtCognome.DataBindings.Add("text", BSClienti, "Cognome")
        txtNome.DataBindings.Add("text", BSClienti, "Nome")
        txtRagSoc.DataBindings.Add("text", BSClienti, "RagioneSociale")
        txtPartIVA.DataBindings.Add("text", BSClienti, "PartitaIVA")
        txtCodFisc.DataBindings.Add("text", BSClienti, "CodFiscale")
        txtVia.DataBindings.Add("text", BSClienti, "Via")
        txtCAP.DataBindings.Add("text", BSClienti, "CAP")
        txtCitta.DataBindings.Add("text", BSClienti, "Citta")


        cmbNazione.DataBindings.Add("SelectedValue", BSClienti, "FKNazione")
        cmbNazione.DataSource = dtNazCli
        cmbNazione.DisplayMember = "Nome"
        cmbNazione.ValueMember = "PKNazione"

        cmbPag.DataBindings.Add("SelectedValue", BSClienti, "FKPagamento")
        cmbPag.DataSource = dtPagCli
        cmbPag.DisplayMember = "Modalita"
        cmbPag.ValueMember = "PKPagamento"

    End Sub
    
    Private Sub BtnScegli_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnScegli.Click
        Dim TestoQuery As String
        Dim NomeForm As String
        Dim IDTrovato As String
        Dim Pos As Integer
        Dim IDAnnulla As String





        TestoQuery = "SELECT * FROM tblCliente"
        NomeForm = "Selezione Clienti"
        NasColonna = True
        IDAnnulla = txtid.Text

        RicercaDati(Me, TestoQuery, NomeForm, IDTrovato, NasColonna)

        If IDTrovato = "" Then
            IDTrovato = IDAnnulla
        Else
            Pos = BSClienti.Find("PKCliente", CInt(IDTrovato))

        End If
        BSClienti.Position = Pos
        Call VisualizzaPosizione()


    End Sub

    '------------- ABILITAZIONE E DISABILITAZIONE ----------

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

    Private Sub DisabilitaText()
        grpInserisci.Enabled = False

    End Sub
    Private Sub AbilitaText()
        grpInserisci.Enabled = True
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

    '-------------- NAVIGATORE ------------

    Private Sub btnPrimo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrimo.Click
        BSClienti.Position = 0

        Call VisualizzaPosizione()
    End Sub
    Private Sub btnPrima_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrima.Click
        BSClienti.Position = BSClienti.Position - 1

        Call VisualizzaPosizione()
    End Sub
    Private Sub btnDopo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDopo.Click
        BSClienti.Position = BSClienti.Position + 1

        Call VisualizzaPosizione()
    End Sub
    Private Sub btnUltimo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUltimo.Click
        BSClienti.Position = BSClienti.Count - 1

        Call VisualizzaPosizione()
    End Sub

    Private Sub VisualizzaPosizione()
        Dim totale As Integer
        Dim posizione As Integer
        Dim stringaposizione As String

        totale = BSClienti.Count
        posizione = BSClienti.Position + 1
        stringaposizione = CStr(posizione) & " di " & CStr(totale)
        txtPosizione.Text = stringaposizione

    End Sub
End Class