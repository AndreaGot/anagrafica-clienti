Imports system.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Windows.Forms


Module Modulo
    Public Primavolta As Boolean

    Public Conn As OleDb.OleDbConnection

    Public DA As New OleDb.OleDbDataAdapter
    Public dt As New DataTable
    Public ds As DataSet
    Public bsNaz As BindingSource
    Public DANaz As New OleDb.OleDbDataAdapter
    Public dtNaz As New DataTable

    Public DSUM As DataSet
    Public bsUM As BindingSource
    Public DAUM As New OleDb.OleDbDataAdapter
    Public dtUM As New DataTable

    Public DSIVA As DataSet
    Public bsIVA As BindingSource
    Public DAIVA As New OleDb.OleDbDataAdapter


    Public DSCATART As DataSet
    Public bsCATART As BindingSource
    Public DACATART As New OleDb.OleDbDataAdapter

    Public DSART As DataSet
    Public bsART As BindingSource
    Public DAART As New OleDb.OleDbDataAdapter
    Public dtART As New DataTable

    Public DAIvaxO As New OleDb.OleDbDataAdapter
    Public dtIvaxO As New DataTable
    '-------ORDINI-------

    Public DSOrdini As New DataSet
    Public DAOrdCliDett As New OleDb.OleDbDataAdapter
    Public DAOrdCliTes As New OleDb.OleDbDataAdapter
    Public BsOrdCliDett As New BindingSource
    Public BsOrdcliTes As New BindingSource
    Public cmdLeggiOrdCliDett As New OleDb.OleDbCommand
    Public cmdLeggiOrdCliTes As New OleDbCommand
    Public RelOrddett As DataRelation
    Public DAPAGORD As New OleDb.OleDbDataAdapter
    Public dtPAGORD As New DataTable
    Public DACLIORD As New OleDb.OleDbDataAdapter
    Public dtCLIORD As New DataTable
    Public cmdLeggiPAGORD As New OleDbCommand
    Public cmdLeggiCLIORD As New OleDbCommand
    Public cmdIvaxO As New OleDbCommand
    Public cmdEliminaOrdCliTes As New OleDbCommand
    Public cmdEliminaOrdCliDett As New OleDbCommand

    Public cmdLeggiIVAORD As New OleDb.OleDbCommand
    Public DAIVAORD As New OleDb.OleDbDataAdapter
    Public dtIVA As New DataTable

    '---------

    Public DANazCli As New OleDb.OleDbDataAdapter
    Public dtNazCli As New DataTable
    Public DAUMART As New OleDb.OleDbDataAdapter
    Public dtUMART As New DataTable
    Public DAIVAART As New OleDb.OleDbDataAdapter
    Public dtIVAART As New DataTable
    Public DACATARTART As New OleDb.OleDbDataAdapter
    Public dtCATARTART As New DataTable
    Public DAPag As New OleDb.OleDbDataAdapter
    Public dtPag As New DataTable
    Public DAPagCli As New OleDb.OleDbDataAdapter
    Public dtPagCli As New DataTable

    Public DARicerca As New OleDb.OleDbDataAdapter
    Public BSRicerca As BindingSource
    Public DSRicerca As DataSet
    Public cmdLeggiRicerca As New OleDbCommand

    Public cmdLeggiIVA As New OleDb.OleDbCommand
    Public cmdLeggiUM As New OleDb.OleDbCommand
    Public cmdLeggiCATART As New OleDb.OleDbCommand

    Public cmdLeggi As New OleDb.OleDbCommand
    Public cmdAggiungi As New OleDb.OleDbCommand
    Public cmdElimina As New OleDb.OleDbCommand
    Public cmdSalva As New OleDb.OleDbCommand

    Public cmdLeggiNazCli As New OleDb.OleDbCommand
    Public cmdLeggiPagCli As New OleDb.OleDbCommand

    Public cmdLeggiUMART As New OleDb.OleDbCommand
    Public cmdLeggiIVAART As New OleDb.OleDbCommand
    Public cmdLeggiCATARTART As New OleDb.OleDbCommand

    Public cmdLeggiNazione As New OleDb.OleDbCommand
    Public cmdAggiungiNazione As New OleDb.OleDbCommand
    Public cmdEliminaNazione As New OleDb.OleDbCommand
    Public cmdSalvaNazione As New OleDb.OleDbCommand

    Public cmdLeggiPag As New OleDb.OleDbCommand
    Public cmdAggiungiPag As New OleDb.OleDbCommand
    Public cmdEliminaPag As New OleDb.OleDbCommand
    Public cmdSalvaPag As New OleDb.OleDbCommand

    Public cmdLeggiArt As New OleDb.OleDbCommand
    Public cmdAggiungiArt As New OleDb.OleDbCommand
    Public cmdEliminaArt As New OleDb.OleDbCommand
    Public cmdSalvaArt As New OleDb.OleDbCommand

    Public cmdLeggids As New OleDb.OleDbCommand
    Public IDannulla As Integer

    Public NasColonna As Boolean


    '-----------CONNESSIONI-------------

    Public Sub ApriConn()


        Dim NomeDb As String

        Conn = New OleDb.OleDbConnection
        Call LeggiIni(NomeDb)

        Conn.ConnectionString = "Provider = Microsoft.Jet.OleDb.4.0;Data Source =" & NomeDb
        Conn.Open()

    End Sub
    Private Sub LeggiIni(ByRef NomeDatabase As String)
        Dim Percorso As New FileStream("Clienti.txt", FileMode.Open, FileAccess.Read)
        Dim OggettoLeggi As New StreamReader(Percorso)
        NomeDatabase = OggettoLeggi.ReadLine
    End Sub

    Public Sub ChiudiConnNaz()

        dtNaz.Clear()
        DANaz.Dispose()
        dtNaz.Dispose()

        Conn.Close()
        Conn.Dispose()
    End Sub
    Public Sub ChiudiConnPag()

        dtPag.Clear()
        DAPag.Dispose()
        dtPag.Dispose()

        Conn.Close()
        Conn.Dispose()
    End Sub
    Public Sub ChiudiConn()

        dtART.Clear()
        DAART.Dispose()
        dtART.Dispose()

        dtCATARTART.Clear()
        DACATARTART.Dispose()
        dtCATARTART.Dispose()

        dtUMART.Clear()
        DAUMART.Dispose()
        dtUMART.Dispose()

        dtIVAART.Clear()
        DAIVAART.Dispose()
        dtIVAART.Dispose()


        Conn.Close()
        Conn.Dispose()
    End Sub

    '-------- DATA ADAPTER E DATA TABLE ---------

    Public Sub ClientiDaDt()
        Call ImpostaComandi()

        DA = New OleDb.OleDbDataAdapter(cmdLeggi)
        DA.Fill(dt)

        DAPagCli = New OleDbDataAdapter(cmdLeggiPagCli)
        DAPagCli.Fill(dtPagCli)

        DANazCli = New OleDbDataAdapter(cmdLeggiNazCli)
        DANazCli.Fill(dtNazCli)

    End Sub
    Public Sub ArticoliDaDt()
        Call ImpostaComandiArt()

        DAART = New OleDb.OleDbDataAdapter(cmdLeggiArt)
        DAART.Fill(dtART)

        DAUMART = New OleDbDataAdapter(cmdLeggiUMART)
        DAUMART.Fill(dtUMART)

        DAIVAART = New OleDbDataAdapter(cmdLeggiIVAART)
        DAIVAART.Fill(dtIVAART)

        DACATARTART = New OleDbDataAdapter(cmdLeggiCATARTART)
        DACATARTART.Fill(dtCATARTART)

    End Sub
    Public Sub PagDaDt()
        Call ImpostaComandiPag()

        DAPag = New OleDb.OleDbDataAdapter(cmdLeggiPag)
        DAPag.Fill(dtPag)
    End Sub
    Public Sub ComboIvaDaDt()
        Call ImpostaComandiIVAORD()

        DAIVAORD = New OleDb.OleDbDataAdapter(cmdLeggiIVAORD)
        DAIVAORD.Fill(dtIVA)
    End Sub

    '--------- IMPOSTAZIONE COMANDI -------
  
    Public Sub ImpostaComandi()
        cmdLeggi.Connection = Conn
        cmdLeggi.CommandType = CommandType.Text
        cmdLeggi.CommandText = "SELECT * FROM tblCliente"

        cmdAggiungi.Connection = Conn
        cmdAggiungi.CommandType = CommandType.Text
        cmdAggiungi.CommandText = "INSERT INTO tblCliente (RagioneSociale,Nome,Cognome,PartitaIVA,CodFiscale,Via,CAP,Citta,FKNazione, FKPagamento) VALUES (?,?,?,?,?,?,?,?,?,?)"

        cmdElimina.Connection = Conn
        cmdElimina.CommandType = CommandType.Text
        cmdElimina.CommandText = "DELETE FROM tblCliente WHERE PKCliente=?"

        cmdSalva.Connection = Conn
        cmdSalva.CommandType = CommandType.Text
        cmdSalva.CommandText = "UPDATE tblCliente SET RagioneSociale=? ,Nome=?, Cognome=?, PartitaIVA=?, CodFiscale=?, Via=?, CAP=?, Citta=?, FKNazione=?, FKPagamento=? WHERE PKCliente=?"

        cmdLeggiNazCli.Connection = Conn
        cmdLeggiNazCli.CommandType = CommandType.Text
        cmdLeggiNazCli.CommandText = "SELECT * FROM tblNazione"

        cmdLeggiPagCli.Connection = Conn
        cmdLeggiPagCli.CommandType = CommandType.Text
        cmdLeggiPagCli.CommandText = "SELECT * FROM tblPagamento"

    End Sub
    Public Sub ImpostaComandiPag()
        cmdLeggiPag.Connection = Conn
        cmdLeggiPag.CommandType = CommandType.Text
        cmdLeggiPag.CommandText = "SELECT * FROM tblPagamento"

        cmdAggiungiPag.Connection = Conn
        cmdAggiungiPag.CommandType = CommandType.Text
        cmdAggiungiPag.CommandText = "INSERT INTO tblPagamento (Modalita,Giorni,Rate,GiorniRate,Sconto,SpeseBanca) VALUES (?,?,?,?,?,?)"

        cmdEliminaPag.Connection = Conn
        cmdEliminaPag.CommandType = CommandType.Text
        cmdEliminaPag.CommandText = "DELETE FROM tblPagamento WHERE PKPagamento=?"

        cmdSalvaPag.Connection = Conn
        cmdSalvaPag.CommandType = CommandType.Text
        cmdSalvaPag.CommandText = "UPDATE tblPagamento SET Modalita=?, Giorni=?, Rate=?, GiorniRate=?, Sconto=?, SpeseBanca=? WHERE PKPagamento=?"
    End Sub
    Public Sub ImpostaComandiArt()
        cmdLeggiArt.Connection = Conn
        cmdLeggiArt.CommandType = CommandType.Text
        cmdLeggiArt.CommandText = "SELECT * FROM tblArticoli"

        cmdAggiungiArt.Connection = Conn
        cmdAggiungiArt.CommandType = CommandType.Text
        cmdAggiungiArt.CommandText = "INSERT INTO tblArticoli (PKArticolo,DesArticolo,UM,FKCatMag,PrezzoListino,QtaSconto1,Sconto1,QtaSconto2,Sconto2,ScontoA,ScontoB,FKIva,ScortaMax,ScortaMin,LottoEco,CostoMedio,UltPrezzoAcq) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

        cmdEliminaArt.Connection = Conn
        cmdEliminaArt.CommandType = CommandType.Text
        cmdEliminaArt.CommandText = "DELETE FROM tblArticoli WHERE PKArticolo=?"

        cmdSalvaArt.Connection = Conn
        cmdSalvaArt.CommandType = CommandType.Text
        cmdSalvaArt.CommandText = "UPDATE tblArticoli SET PKArticolo=?, DesArticolo=?, UM=?, FKCatMag=?, PrezzoListino=?, QtaSconto1=?, Sconto1=?, QtaSconto2=?,Sconto2=?,ScontoA=?,ScontoB=?,FKIva=?,ScortaMax=?,ScortaMin=?,LottoEco=?,CostoMedio=?,UltPrezzoAcq=? WHERE PKArticolo=?"

        cmdLeggiUMART.Connection = Conn
        cmdLeggiUMART.CommandType = CommandType.Text
        cmdLeggiUMART.CommandText = "SELECT * FROM tblUM"

        cmdLeggiIVAART.Connection = Conn
        cmdLeggiIVAART.CommandType = CommandType.Text
        cmdLeggiIVAART.CommandText = "SELECT * FROM tblIVA"

        cmdLeggiCATARTART.Connection = Conn
        cmdLeggiCATARTART.CommandType = CommandType.Text
        cmdLeggiCATARTART.CommandText = "SELECT * FROM tblCatMag"
    End Sub
    Public Sub ImpostaComandiDS()
        cmdLeggiNazione.Connection = Conn
        cmdLeggiNazione.CommandType = CommandType.TableDirect
        cmdLeggiNazione.CommandText = "tblNazione"

    End Sub
    Public Sub ImpostaComandiDSUM()
        cmdLeggiUM.Connection = Conn
        cmdLeggiUM.CommandType = CommandType.TableDirect
        cmdLeggiUM.CommandText = "tblUM"


    End Sub
    Public Sub ImpostaComandiDSIVA()
        cmdLeggiIVA.Connection = Conn
        cmdLeggiIVA.CommandType = CommandType.TableDirect
        cmdLeggiIVA.CommandText = "tblIVA"

    End Sub
    Public Sub ImpostaComandiDSCATART()
        cmdLeggiCATART.Connection = Conn
        cmdLeggiCATART.CommandType = CommandType.TableDirect
        cmdLeggiCATART.CommandText = "tblCatMag"


    End Sub
    Public Sub ImpostaComandiDSART()
        cmdLeggiArt.Connection = Conn
        cmdLeggiArt.CommandType = CommandType.TableDirect
        cmdLeggiArt.CommandText = "tblArticoli"

    End Sub
    Public Sub ImpostaComandiDSOrdini()
        cmdLeggiOrdCliDett.Connection = Conn
        cmdLeggiOrdCliDett.CommandType = CommandType.TableDirect
        cmdLeggiOrdCliDett.CommandText = "tblOrdCliDett"

        cmdLeggiOrdCliTes.Connection = Conn
        cmdLeggiOrdCliTes.CommandType = CommandType.Text
        cmdLeggiOrdCliTes.CommandText = "SELECT * FROM tblOrdCliTes"

        cmdIvaxO.Connection = Conn
        cmdIvaxO.CommandType = CommandType.TableDirect
        cmdIvaxO.CommandText = "tblIva"

    End Sub
    Public Sub ImpostaComandiCOMBOORD()
        cmdLeggiPAGORD.Connection = Conn
        cmdLeggiPAGORD.CommandType = CommandType.Text
        cmdLeggiPAGORD.CommandText = "SELECT * FROM tblPagamento"

        cmdLeggiCLIORD.Connection = Conn
        cmdLeggiCLIORD.CommandType = CommandType.Text
        cmdLeggiCLIORD.CommandText = "SELECT * FROM tblCliente"
    End Sub
    Public Sub ImpostaComandiOrdini()

        cmdEliminaOrdCliTes.Connection = Conn
        cmdEliminaOrdCliTes.CommandType = CommandType.Text
        cmdEliminaOrdCliTes.CommandText = "DELETE FROM tblOrdCliTes WHERE PKOrdCliTes=?"

        cmdEliminaOrdCliDett.Connection = Conn
        cmdEliminaOrdCliDett.CommandType = CommandType.Text
        cmdEliminaOrdCliDett.CommandText = "DELETE FROM tblOrdCliDett WHERE FKOrdCliTes=?"



    End Sub
    Public Sub ImpostaComandiIVAORD()
        cmdLeggiIVAORD.Connection = Conn
        cmdLeggiIVAORD.CommandType = CommandType.Text
        cmdLeggiIVAORD.CommandText = "SELECT * FROM tblIVA"

    End Sub


    '---------IMPOSTAZIONI VARIE -------------

    Public Sub ImpostaDS()
        ds = New DataSet("XXX")
        Call ImpostaComandiDS()
        DANaz = New OleDbDataAdapter(cmdLeggiNazione)
        DANaz.Fill(ds, "YYY")
        bsNaz = New BindingSource
        bsNaz.DataSource = ds
        bsNaz.DataMember = "YYY"
    End Sub
    Public Sub ImpostaDSUM()
        DSUM = New DataSet("XXX")
        Call ImpostaComandiDSUM()
        DAUM = New OleDbDataAdapter(cmdLeggiUM)
        DAUM.Fill(DSUM, "YYY")
        bsUM = New BindingSource
        bsUM.DataSource = DSUM
        bsUM.DataMember = "YYY"
    End Sub
    Public Sub ImpostaDSIVA()
        DSIVA = New DataSet("XXX")
        Call ImpostaComandiDSIVA()
        DAIVA = New OleDbDataAdapter(cmdLeggiIVA)
        DAIVA.Fill(DSIVA, "YYY")
        bsIVA = New BindingSource
        bsIVA.DataSource = DSIVA
        bsIVA.DataMember = "YYY"
    End Sub
    Public Sub ImpostaDSCATART()
        DSCATART = New DataSet("XXX")
        Call ImpostaComandiDSCATART()
        DACATART = New OleDbDataAdapter(cmdLeggiCATART)
        DACATART.Fill(DSCATART, "YYY")
        bsCATART = New BindingSource
        bsCATART.DataSource = DSCATART
        bsCATART.DataMember = "YYY"
    End Sub
    Public Sub ImpostaDSART()
        DSART = New DataSet("XXX")
        Call ImpostaComandiDSART()
        DAART = New OleDbDataAdapter(cmdLeggiArt)
        DAART.Fill(DSART, "YYY")
        bsART = New BindingSource
        bsART.DataSource = DSART
        bsART.DataMember = "YYY"
    End Sub
    Public Sub ImpostaDSOrdini()

        DSOrdini = New DataSet("XXX")
        Call ImpostaComandiDSOrdini()

        DAOrdCliTes = New OleDbDataAdapter(cmdLeggiOrdCliTes)

        DAOrdCliDett = New OleDbDataAdapter(cmdLeggiOrdCliDett)

        Call RiempiDSOrdini()

        RelOrddett = New DataRelation("RelOrdCliTesOrdCliDett", DSOrdini.Tables("OrdCliTes").Columns("PKOrdCliTes"), DSOrdini.Tables("OrdCliDett").Columns("FKOrdCliTes"))
        DSOrdini.Relations.Add(RelOrddett)

        BsOrdcliTes.DataSource = DSOrdini
        BsOrdcliTes.DataMember = "OrdCliTes"

        BsOrdCliDett.DataSource = BsOrdcliTes
        BsOrdCliDett.DataMember = "RelOrdCliTesOrdCliDett"

        'istanzio un nuovo DataAdapter, passandogli come costruttore il Command creato in precedenza
        DAIvaxO = New OleDbDataAdapter(cmdIvaxO)
        '--- esegue il dataadapter e riempie il datatable
        DAIvaxO.Fill(dtIvaxO)


    End Sub
    Public Sub ImpostaComboORD()
        Call ImpostaComandiCOMBOORD()

        DAPAGORD = New OleDb.OleDbDataAdapter(cmdLeggiPAGORD)
        DAPAGORD.Fill(dtPAGORD)

        DACLIORD = New OleDb.OleDbDataAdapter(cmdLeggiCLIORD)
        DACLIORD.Fill(dtCLIORD)


    End Sub

    '-------------- ALTRO --------------

    Public Sub RicercaDati(ByVal x As Form, ByVal TestoQuery As String, ByVal NomeForm As String, ByRef IdTrovato As String, ByVal NasColonna As Boolean)

        cmdLeggiRicerca = New OleDbCommand
        cmdLeggiRicerca.Connection = Conn
        cmdLeggiRicerca.CommandType = CommandType.Text
        cmdLeggiRicerca.CommandText = TestoQuery

        DSRicerca = New DataSet("Ricerca")
        DARicerca = New OleDbDataAdapter(cmdLeggiRicerca)

        DARicerca.Fill(DSRicerca, "Ric")

        BSRicerca = New BindingSource
        BSRicerca.DataSource = DSRicerca
        BSRicerca.DataMember = "Ric"

        Ricerca.txtCampi.Text = "RagioneSociale"
        Ricerca.Text = NomeForm
        Ricerca.txtQuery.Text = TestoQuery
        Ricerca.lblRicerca.Text = "Ricerca " & Ricerca.txtCampi.Text


        Ricerca.ShowDialog()
        IdTrovato = Ricerca.txtID.Text

    End Sub
    Public Sub LeggiPKOrdTes(ByVal NumOrd As Integer, ByRef PKOrdTes As Integer)
        Dim Pos As Integer

        Dim dtPKOrdCliTes As New DataTable
        Dim BSRicercaPK As New BindingSource

        DAOrdCliTes.Fill(dtPKOrdCliTes)
        BSRicercaPK.DataSource = dtPKOrdCliTes
        Pos = BsOrdcliTes.Find("NumOrd", NumOrd)

        PKOrdTes = dtPKOrdCliTes.Rows(Pos).Item("PKOrdCliTes")

    End Sub
    Public Sub RiempiDSOrdini()
        DAOrdCliTes.Fill(DSOrdini, "OrdCliTes")
        DAOrdCliDett.Fill(DSOrdini, "OrdCliDett")
    End Sub
    Public Sub RicercaDatiArticolo(ByVal CodArticolo As String, ByRef Unitamisura As String, ByRef Prezzolistino As Single, ByRef codIva As Integer, ByRef descrizione As String)

        Dim Dtricerca As New DataTable
        Dim bsricerca As New BindingSource
        Dim daricerca As OleDbDataAdapter
        Dim pos As Integer

        cmdLeggiRicerca = New OleDbCommand
        cmdLeggiRicerca.Connection = Conn
        cmdLeggiRicerca.CommandType = CommandType.Text
        cmdLeggiRicerca.CommandText = "SELECT * FROM TblArticoli"
        daricerca = New OleDbDataAdapter(cmdLeggiRicerca)

        DSRicerca = New DataSet("Ricerca")
        daricerca = New OleDbDataAdapter(cmdLeggiRicerca)

        daricerca.Fill(Dtricerca)

        bsricerca.DataSource = Dtricerca

        pos = bsricerca.Find("PkArticolo", CodArticolo)
        Unitamisura = Dtricerca.Rows(pos).Item("UM")
        Prezzolistino = Dtricerca.Rows(pos).Item("Prezzolistino")
        codIva = Dtricerca.Rows(pos).Item("FkIva")
        descrizione = Dtricerca.Rows(pos).Item("DesArticolo")
    End Sub
End Module
