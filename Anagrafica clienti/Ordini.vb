

Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data
Imports System.Data.oledb

Public Class Ordini

    Dim Aggiungi As Boolean
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Call ChiudiConn()
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
       
        dtIvaxO.Clear()
        DAIvaxO.Dispose()
        dtIvaxO.Dispose()
        dtCLIORD.Dispose()
        dtCLIORD.Clear()
        dtPAGORD.Dispose()
        dtPAGORD.Clear()

        DSOrdini.Clear()


        Call ApriConn()
        Primavolta = True
        Call ImpostaDSOrdini()

        grdOrdini.DataSource = BsOrdCliDett
        Call Disabilitasalva()
        Call DisabilitaText()


        Call Abilitanavigatore()
        Call AssociaCampi()
        Call VisualizzaPosizione()

        Call ImpostaComandiOrdini()

        grdOrdini.Columns(0).Visible = False
        grdOrdini.Columns(1).Visible = False
        grdOrdini.Columns(2).HeaderText = "Codice Articolo"
        grdOrdini.Columns(3).HeaderText = "Descrizione"
        grdOrdini.Columns(4).HeaderText = "UM"
        grdOrdini.Columns(5).HeaderText = "Quantità"
        grdOrdini.Columns(6).HeaderText = "Prezzo Unitario"
        grdOrdini.Columns(7).HeaderText = "Sconto 1"
        grdOrdini.Columns(8).HeaderText = "Sconto 2"
        grdOrdini.Columns(9).HeaderText = "Sconto 3"



        Call ImpostaComboGriglia()

    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        BsOrdcliTes.AddNew()
        btnScegli.Enabled = False
        Aggiungi = True
        Call AbilitaGrid()
        Call AbilitaSalva()
        Call AbilitaText()
        Call DisabilitaNavigatore()

        txtPosizione.Text = "Aggiunta"


    End Sub
    Private Sub btnSalva_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalva.Click


        Dim NumOrd As Integer
        Dim DataOrd As Date
        Dim DataCons As Date
        Dim i As Integer
        Dim PKOrdTes As Integer
        Dim NumDettagli As Integer
        Dim FineDettagli As Integer
        Dim InizioDettagli As Integer

        btnScegli.Enabled = True
        If txtNumOrdine.Text = "" Then
            MsgBox("inserire numero ordine")
            Exit Sub
        End If

        NumOrd = CInt(txtNumOrdine.Text)


        If Not IsDate(txtDataOrdine.Text) Or Not IsDate(txtDataConsegna.Text) Then
            MsgBox("Inserire una data valida", MsgBoxStyle.Information, "Data non valida")
            Exit Sub
        End If

        DataOrd = CDate(txtDataOrdine.Text)
        DataCons = CDate(txtDataConsegna.Text)
        ' ----------------------------------
        BsOrdcliTes.EndEdit()
        Dim cb As New OleDb.OleDbCommandBuilder(DAOrdCliTes)
        DAOrdCliTes.Update(DSOrdini, "OrdCliTes")


        If Aggiungi Then
            Aggiungi = False
            Call LeggiPKOrdTes(NumOrd, PKOrdTes)

            NumDettagli = BsOrdCliDett.Count
            FineDettagli = DSOrdini.Tables("OrdCliDett").Rows.Count
            InizioDettagli = FineDettagli - NumDettagli

            DSOrdini.Tables("OrdCliDett").BeginLoadData()

            For i = InizioDettagli To FineDettagli - 1
                DSOrdini.Tables("OrdCliDett").Rows(i).Item("FkOrdCliTes") = PKOrdTes
            Next

            BsOrdCliDett.EndEdit()
            Dim cbDett As New OleDb.OleDbCommandBuilder(DAOrdCliDett)
            DAOrdCliDett.Update(DSOrdini, "OrdCliDett")

            DSOrdini.Clear()
            Call RiempiDSOrdini()
            BsOrdcliTes.MoveLast()
            '-----
        Else

            BsOrdCliDett.EndEdit()
            Dim cb2 As New OleDb.OleDbCommandBuilder(DAOrdCliDett)
            DAOrdCliDett.Update(DSOrdini, "OrdCliDett")
        End If

        Call Disabilitasalva()
        Call DisabilitaText()

        Call Abilitanavigatore()




    End Sub
    Private Sub btnElimina_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnElimina.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        msg = "Vuoi veramente eliminare questo record?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo
        title = "Elimina"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            Dim CodOrdine As Integer
            Dim pos As Integer
            pos = BsOrdcliTes.Position
            CodOrdine = CInt((DSOrdini.Tables("OrdCliTes").Rows(pos).Item("PKOrdCliTes")))
            cmdEliminaOrdCliTes.Parameters.Add(New OleDb.OleDbParameter("PKOrdCliTes", CodOrdine))
            cmdEliminaOrdCliDett.Parameters.Add(New OleDb.OleDbParameter("FKOrdCliTes", CodOrdine))

            cmdEliminaOrdCliDett.ExecuteNonQuery()
            cmdEliminaOrdCliTes.ExecuteNonQuery()

            cmdEliminaOrdCliTes.Parameters.Clear()
            cmdEliminaOrdCliDett.Parameters.Clear()

            DSOrdini.Clear()
            Call ImpostaDSOrdini()

            Call VisualizzaPosizione()

        End If




        

    End Sub
    Private Sub btnFine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFine.Click
        Me.Close()
    End Sub
    Private Sub btnAnnulla_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAnnulla.Click
        BSOrdCliTes.CancelEdit()
        btnScegli.Enabled = True
        Call Disabilitasalva()
        Call disabilitatext()

        Call Abilitanavigatore()
        Call VisualizzaPosizione()

    End Sub
    Private Sub btnModifica_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModifica.Click
        Aggiungi = False
        btnScegli.Enabled = False

        Call AbilitaSalva()
        Call AbilitaText()
        Call AbilitaGrid()
    End Sub

    Private Sub AssociaCampi()


        txtNumOrdine.DataBindings.Add("text", BsOrdcliTes, "NumOrd")
        txtDataOrdine.DataBindings.Add("text", BsOrdcliTes, "DataOrdine")
        txtDataConsegna.DataBindings.Add("text", BsOrdcliTes, "DataCons")



        Call ImpostaComboORD()

        cmbCliente.DataBindings.Add("SelectedValue", BsOrdcliTes, "FKCliente")
        cmbCliente.DataSource = dtCLIORD
        cmbCliente.DisplayMember = "RagioneSociale"
        cmbCliente.ValueMember = "PKCliente"

        cmbPag.DataBindings.Add("SelectedValue", BsOrdcliTes, "FKPagamento")
        cmbPag.DataSource = dtPAGORD
        cmbPag.DisplayMember = "Modalita"
        cmbPag.ValueMember = "PKPagamento"

        grdOrdini.Columns.Remove("CodIva")

    End Sub

    Private Sub ImpostaComboGriglia()

        'rimuove chiave esterna dalla griglia
        'grdOrdini.Columns.Remove("CodIva")
        'dichiara nuovo oggetto di tipo Combo
        Dim cmbIvaDett As New DataGridViewComboBoxColumn
        'imposta intestazione
        cmbIvaDett.HeaderText = "Iva"
        'imposta la sorgente data di quest'oggetto
        cmbIvaDett.DataSource = dtIvaxO
        'imposta il campo da visual nella combo
        cmbIvaDett.DisplayMember = "Descrizione"
        'imposta il campo nella tab da cui recuperare il val
        cmbIvaDett.ValueMember = "PKIva"
        'imposta il campo in cui mem il val
        cmbIvaDett.DataPropertyName = "CodIva"
        'aggiungi la combo alla griglia
        grdOrdini.Columns.Add(cmbIvaDett)

        'Call ComboIvaDaDt()


        'Dim cmbIvaDet As New DataGridViewComboBoxColumn

        'cmbIvaDet.HeaderText = "IVA"
        'cmbIvaDet.DataSource = dtIVA
        'cmbIvaDet.DisplayMember = "Aliquota"
        'cmbIvaDet.ValueMember = "PKIva"
        'cmbIvaDet.DataPropertyName = "FkIva"

        'grdOrdini.Columns.Add(cmbIvaDet)
    End Sub

    '----------- NAVIGATORE ----------

    Private Sub btnPrimo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrimo.Click
        BsOrdcliTes.Position = 0

        Call VisualizzaPosizione()
    End Sub
    Private Sub btnPrima_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrima.Click
        BsOrdcliTes.Position = BsOrdcliTes.Position - 1

        Call VisualizzaPosizione()
    End Sub
    Private Sub btnDopo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDopo.Click
        BsOrdcliTes.Position = BsOrdcliTes.Position + 1

        Call VisualizzaPosizione()
    End Sub
    Private Sub btnUltimo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUltimo.Click
        BsOrdcliTes.Position = BsOrdcliTes.Count - 1

        Call VisualizzaPosizione()
    End Sub

    Private Sub VisualizzaPosizione()
        Dim totale As Integer
        Dim posizione As Integer
        Dim stringaposizione As String

        totale = BsOrdcliTes.Count
        posizione = BsOrdcliTes.Position + 1
        stringaposizione = CStr(posizione) & " di " & CStr(totale)
        txtPosizione.Text = stringaposizione

    End Sub

    '----------- ABILITAZIONE - DISABILITAZIONE ---------

    Private Sub AbilitaSalva()
        grpConferma.Enabled = True
        grpOper.Enabled = False
        btnFine.Enabled = False
        grdOrdini.ReadOnly = False

    End Sub
    Private Sub Disabilitasalva()
        grpConferma.Enabled = False
        grpOper.Enabled = True
        btnFine.Enabled = True
        grdOrdini.ReadOnly = True

    End Sub

    Private Sub AbilitaText()
        grpText.Enabled = True
        grdOrdini.ReadOnly = False

    End Sub
    Private Sub DisabilitaText()
        grpText.Enabled = False
        grdOrdini.ReadOnly = True

    End Sub

    Private Sub AbilitaGrid()
        grdOrdini.Enabled = True
        grdOrdini.ReadOnly = False

    End Sub
   

    Private Sub Abilitanavigatore()
        btnPrimo.Enabled = True
        btnUltimo.Enabled = True
        btnPrima.Enabled = True
        btnDopo.Enabled = True
        grdOrdini.ReadOnly = True

    End Sub
    Private Sub DisabilitaNavigatore()
        btnPrimo.Enabled = False
        btnUltimo.Enabled = False
        btnPrima.Enabled = False
        btnDopo.Enabled = False
        grdOrdini.ReadOnly = False

    End Sub

    Private Sub BtnScegli_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnScegli.Click
        Dim TestoQuery As String
        Dim NomeForm As String
        Dim IDTrovato As String
        Dim Pos As Integer






        TestoQuery = "SELECT * FROM tblOrdCliTes"
        NomeForm = "Selezione Clienti"
        NasColonna = True



        RicercaDati(Me, TestoQuery, NomeForm, IDTrovato, NasColonna)

        If IDTrovato <> "" Then
            Pos = BsOrdcliTes.Find("PKOrdCliTes", CInt(IDTrovato))
        End If
        BsOrdcliTes.Position = Pos
        Call VisualizzaPosizione()
    End Sub


    Private Sub grdOrdini_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdOrdini.CellDoubleClick
        Dim Riga As Integer
        Dim Colonna As Integer

        Dim TestoQuery As String
        Dim NomeForm As String
        Dim IDTrovato As String
        Dim prezzolistino As Single
        Dim codiva As Integer
        Dim Unitamisura As String
        Dim descrizione As String

        Colonna = e.ColumnIndex
        Riga = e.RowIndex
        If grdOrdini.ReadOnly = False Then


            If Colonna = 2 Then


                TestoQuery = "SELECT * FROM tblArticoli"
                NomeForm = "Selezione Articoli"
                NasColonna = False


                Call RicercaDati(Me, TestoQuery, NomeForm, IDTrovato, NasColonna)
                grdOrdini.Rows(Riga).Cells(Colonna).Value = IDTrovato

                Call RicercaDatiArticolo(IDTrovato, Unitamisura, prezzolistino, codiva, descrizione)

                grdOrdini.Rows(Riga).Cells("UM").Value = UnitaMisura
                grdOrdini.Rows(Riga).Cells("PrezzoUni").Value = PrezzoListino
                grdOrdini.Rows(Riga).Cells(10).Value = codiva
                grdOrdini.Rows(Riga).Cells("FKDescrizione").Value = descrizione

                grdOrdini.CurrentCell = grdOrdini.Rows(Riga).Cells("Quantita")



            End If

        End If


    End Sub

    
  

   

    Private Sub BntVai_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BntVai.Click
        Dim message, title, defaultValue As String
        Dim myValue As Object

        message = "Inserisci la posizione del record che vuoi raggiungere!"

        title = "Spostamanto"
        defaultValue = "1"

        myValue = InputBox(message, title, defaultValue)

        If myValue Is "" Or IsNumeric(myValue) = False Then
            myValue = defaultValue
        End If

        BsOrdcliTes.Position = myValue - 1

        Call VisualizzaPosizione()



    End Sub

    Private Sub grdOrdini_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdOrdini.CellContentClick

    End Sub
End Class
