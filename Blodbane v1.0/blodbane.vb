'Imports System.ComponentModel
Imports MySql.Data.MySqlClient
Imports System.Globalization
Public Class Blodbane
    Dim giversøk As New DataTable
    Dim egenerklaering As New DataTable
    Dim innkalling As New DataTable
    Dim blodlager As New DataTable
    Public ansatt As New DataTable
    Dim Erklæringspørsmål As New DataTable
    Dim personstatusK As New Hashtable
    Dim personstatusB As New Hashtable
    Dim postnummer As New Hashtable
    Dim blodgiverData As New Hashtable
    Dim rommene As New ArrayList
    Dim interntabellRom As New DataTable
    Dim antallRom As Integer
    Dim blodgiveren As Blodgiver
    Dim bytteRomTime As Romtime
    Dim egenerklaeringObjekt As Egenerklaering
    Dim fulltimetabell As New ArrayList()
    Dim dummyDato As Date = New Date(1800, 1, 1, 1, 1, 1)
    Dim dummyFodselsnr, aarstallet As String
    Dim dummyEpost As String = "@@.@...@..@."
    Public påloggetAnsatt, påloggetAepost, påloggetBgiver As String
    Dim egenerklæringID, SPMnr, SPMnrPresentert, erklæringSvar(60) As Integer
    Dim presentertGiver, bgSøkParameter As String
    Dim tilkobling As New MySqlConnection("Server=mysql.stud.iie.ntnu.no;" & "Database=g_ioops_02;" & "Uid=g_ioops_02;" & "Pwd=LntL4Owl;")

    'Kjøres ved oppstart
    Private Sub Blodbane_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Hide()
        velkommen.Show()

        Me.Hide()
        velkommen.ProgressBar1.Value = 0
        'Henter statuskoder og legger i combobox(er)
        Dim statuser As New DataTable
        Dim steder As New DataTable
        Dim da As New MySqlDataAdapter
        Dim rad As DataRow
        Dim sifre As String
        Dim statustekst, statuskode, psted, pnr As String
        giversøk.Clear()
        tilkobling.Open()
        Dim sqlSpørring As New MySqlCommand("SELECT * FROM personstatus", tilkobling)
        da.SelectCommand = sqlSpørring
        da.Fill(statuser)
        cBxSøkStatusbeskrivelse.Items.Clear()
        For Each rad In statuser.Rows
            statustekst = rad("beskrivelse")
            statuskode = rad("kode")
            personstatusK.Add(statustekst, statuskode)
            personstatusB.Add(statuskode, statustekst)
            cBxSøkStatusbeskrivelse.Items.Add(statustekst)
            cBxValgtBlodgiverStatusTekst.Items.Add(statustekst)
        Next

        Me.Hide()
        velkommen.ProgressBar1.Value = 20
        blodgiveren = New Blodgiver("", "", "", "", "", dummyDato, "", "", "", "", "", "", "", "", 0)
        bytteRomTime = New Romtime(dummyDato, "", 0)
        egenerklaeringObjekt = New Egenerklaering(-1, dummyEpost, dummyEpost, "", "", dummyDato, dummyDato)
        If Today.Month < 10 Then
            sifre = $"0{Today.Month}"
        Else
            sifre = CStr(Today.Month)
        End If
        aarstallet = CStr(Today.Year).Substring(2, 2)
        dummyFodselsnr = $"{sifre}{sifre}{aarstallet}11111"

        'Lager liste over rommene
        antallRom = 0
        Dim sqlSporringRom As String = "SELECT * FROM rom"
        Dim sqlRom As New MySqlCommand(sqlSporringRom, tilkobling)
        Dim daRom As New MySqlDataAdapter
        daRom.SelectCommand = sqlRom
        daRom.Fill(interntabellRom)
        antallRom = interntabellRom.Rows.Count()

        Me.Hide()
        velkommen.ProgressBar1.Value = 50
        'Henter postnummer og sted og legger i hashtable
        Dim sqlSpørring2 As New MySqlCommand("SELECT * FROM postnummer", tilkobling)
        da.SelectCommand = sqlSpørring2
        da.Fill(steder)
        For Each rad In steder.Rows
            psted = rad("poststed")
            pnr = rad("postnummer")
            postnummer.Add(pnr, psted)
        Next

        'Henter ansatte og legger i datatable
        Dim sqlSpørring3 As New MySqlCommand("SELECT a.epost, b.passord, b.fornavn FROM ansatt a INNER JOIN bruker b ON a.epost = b.epost", tilkobling)
        da.SelectCommand = sqlSpørring3
        da.Fill(ansatt)
        For Each rad In ansatt.Rows
            cbxAnsattUtførtTapping.Items.Add(rad("epost"))
        Next

        Me.Hide()
        velkommen.ProgressBar1.Value = 80
        'Henter ned spørsmål til egenerklæring
        Dim sqlSpørring4 As New MySqlCommand("SELECT * FROM egenerklaeringsporsmaal", tilkobling)
        da.SelectCommand = sqlSpørring4
        da.Fill(Erklæringspørsmål)
        lblEgenerklSpmTekst.Text = Erklæringspørsmål.Rows(0).Item("spoersmaal")
        lblEgenerklSpmNr.Text = $"Spørsmål {SPMnr + 1}"

        tilkobling.Close()

        Me.Hide()
        velkommen.ProgressBar1.Value = 100
        MsgBox("Nå er eg klar!")
        velkommen.Close()
        Me.Show()
    End Sub

    'Nullstiller objektene blodgiveren og bytteRomTime
    Private Sub BlodgiverInit()

        blodgiveren.Fodselsnummer1 = ""
        blodgiveren.Blodtype1 = ""
        blodgiveren.Kontaktform1 = ""
        blodgiveren.Merknad1 = ""
        blodgiveren.Timepreferanse1 = ""
        blodgiveren.Siste_blodtapping1 = dummyDato
        blodgiveren.Epost1 = ""
        blodgiveren.Passord1 = ""
        blodgiveren.Fornavn1 = ""
        blodgiveren.Etternavn1 = ""
        blodgiveren.Adresse1 = ""
        blodgiveren.Telefon11 = ""
        blodgiveren.Telefon21 = ""
        blodgiveren.Postnr1 = ""
        blodgiveren.Status1 = ""

        bytteRomTime.Datotid1 = dummyDato
        bytteRomTime.Romnummer1 = ""
        bytteRomTime.Timenr1 = 0

    End Sub

    'Avslutt program
    Private Sub AvsluttToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AvsluttToolStripMenuItem.Click
        Me.Close()
    End Sub

    'Logg på ansatt
    Private Sub LoggPåansattToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoggPåansattToolStripMenuItem.Click
        pålogging.Show()
    End Sub

    'Logger på blodgiver og setter opp personinfo i fanen Personinformasjon
    Private Sub ButtonLoggpåGiver_Click(sender As Object, e As EventArgs) Handles BttnLoggpåGiver.Click
        tilkobling.Open()
        Dim sql As New MySqlCommand("SELECT * FROM bruker br JOIN blodgiver bg ON br.epost=bg.epost INNER JOIN personstatus p ON p.kode=br.statuskode WHERE br.epost = @epostInn AND br.passord = @passordInn", tilkobling)
        sql.Parameters.AddWithValue("@epostInn", txtAInn_epost.Text)
        sql.Parameters.AddWithValue("@passordInn", txtAInn_passord.Text)
        Dim da As New MySqlDataAdapter
        Dim interntabell As New DataTable
        'Objektet "da" utfører spørringen og legger resultatet i "interntabell"
        da.SelectCommand = sql
        da.Fill(interntabell)
        Dim rad() As DataRow
        Dim antallRader As Integer = interntabell.Rows.Count()


        If antallRader = 1 Then
            rad = interntabell.Select()
            If IsDBNull(rad(0)("blodtype")) Then
                rad(0)("blodtype") = ""
            End If
            If IsDBNull(rad(0)("merknad")) Then
                rad(0)("merknad") = ""
            End If
            If IsDBNull(rad(0)("timepreferanse")) Then
                rad(0)("timepreferanse") = ""
            End If
            If IsDBNull(rad(0)("adresse")) Then
                rad(0)("adresse") = ""
            End If
            If IsDBNull(rad(0)("telefon2")) Then
                rad(0)("telefon2") = ""
            End If
            If IsDBNull(rad(0)("siste_blodtapping")) Then
                rad(0)("siste_blodtapping") = dummyDato
            End If
            BlodgiverObjOppdat(rad(0)("epost"), rad(0)("passord"), rad(0)("fornavn"),
                               rad(0)("etternavn"), rad(0)("adresse"), rad(0)("postnr"),
                               rad(0)("telefon1"), rad(0)("telefon2"), rad(0)("statuskode"),
                               rad(0)("fodselsnummer"), rad(0)("blodtype"), rad(0)("siste_blodtapping"),
                               rad(0)("kontaktform"), rad(0)("merknad"), rad(0)("timepreferanse"))
            'Henter eventuell ny innkalling
            Dim idag, sistetime As DateTime
            Dim ingenNyTime As Boolean = False
            idag = Today
            Dim sql3 As New MySqlCommand("SELECT * FROM timeavtale WHERE bgepost = @epostInn", tilkobling)
            sql3.Parameters.AddWithValue("@epostInn", txtAInn_epost.Text)
            Dim da3 As New MySqlDataAdapter
            Dim interntabell3 As New DataTable
            'Objektet "da" utfører spørringen og legger resultatet i "interntabell"
            da3.SelectCommand = sql3
            da3.Fill(interntabell3)
            If interntabell3.Rows.Count > 0 Then
                Dim rad3() As DataRow = interntabell3.Select()
                sistetime = rad3(interntabell3.Rows.Count - 1)("datotid")
                If sistetime > idag Then
                    TxtNesteInnkalling.Text = sistetime
                    bytteRomTime.Timenr1 = rad3(interntabell3.Rows.Count - 1)("timeid")
                    bytteRomTime.Datotid1 = sistetime
                    bytteRomTime.Romnummer1 = rad3(interntabell3.Rows.Count - 1)("romnr")
                Else
                    ingenNyTime = True
                End If
            Else
                ingenNyTime = True
            End If
            If ingenNyTime Then
                TxtNesteInnkalling.Text = "Ikke fastsatt"
                BtnEndreInnkalling.Enabled = False
            End If
            'Bytter til panelet for blodgiver
            PanelPåmelding.Hide()
            PanelAnsatt.Hide()
            PanelGiver.Show()
            PanelGiver.BringToFront()
            'TabPage5.Show()
            TbCtrlBlodgiver.SelectTab(0)
            'Setter personinfo i tekstboksene
            txtPersDataNavn.Text = $"{blodgiveren.Fornavn1} {blodgiveren.Etternavn1}"
            txtPersDataGStatus.Text = blodgiveren.Status1
            txtPersDataBlodtype.Text = blodgiveren.Blodtype1
            If blodgiveren.Siste_blodtapping1 = dummyDato Then
                txtPersDataSisteUnders.Text = ""
            Else
                txtPersDataSisteUnders.Text = blodgiveren.Siste_blodtapping1
            End If
            txtPersDataGateAdr.Text = blodgiveren.Adresse1
            txtPersDataPostnr.Text = blodgiveren.Postnr1
            txtPersDataTlf.Text = blodgiveren.Telefon11
            txtPersDataTlf2.Text = blodgiveren.Telefon21
            txtPersDataEpost.Text = blodgiveren.Epost1

            CBxKontaktform.Text = blodgiveren.Kontaktform1
            RTxtPrefInnkalling.Text = blodgiveren.Timepreferanse1
            påloggetBgiver = blodgiveren.Postnr1
        Else
            MsgBox("Epostadressen eller passordet er feil.", MsgBoxStyle.Critical)

        End If
        tilkobling.Close()
    End Sub

    'Oppdaterer både tabellene bruker og blodgiver i tillegg til blodgiveren-objektet
    Private Sub OppdaterBlodgiver(ByVal epost As String, ByVal passord As String,
                                  ByVal fornavn As String, ByVal etternavn As String,
                                  ByVal adresse As String, ByVal postnr As String,
                                  ByVal telefon1 As String, ByVal telefon2 As String,
                                  ByVal statuskode As Integer,
                                  ByVal fodselsnummer As String, ByVal blodtype As String,
                                  ByVal siste_blodtapping As Date, ByVal kontaktform As String,
                                  ByVal merknad As String, ByVal timepreferanse As String)

        Me.Cursor = Cursors.WaitCursor
        Dim sqlSporring2 As String = $"UPDATE bruker SET epost='{epost}', passord='{passord}'"
        sqlSporring2 += $", fornavn='{fornavn}', etternavn='{etternavn}'"
        sqlSporring2 += $", adresse='{adresse}', postnr='{postnr}'"
        sqlSporring2 += $", telefon1='{telefon1}', telefon2='{telefon2}'"
        sqlSporring2 += $", statuskode={statuskode} WHERE epost = '{blodgiveren.Epost1}'"
        Dim sql2 As New MySqlCommand(sqlSporring2, tilkobling)
        Dim da2 As New MySqlDataAdapter
        Dim interntabell2 As New DataTable
        'Objektet "da" utfører spørringen og legger resultatet i "interntabell1"
        da2.SelectCommand = sql2
        da2.Fill(interntabell2)

        Dim sqlSporring1 As String = $"UPDATE blodgiver SET fodselsnummer='{fodselsnummer}', blodtype=@blod"
        sqlSporring1 += $", siste_blodtapping=@datotime, kontaktform='{kontaktform}'"
        sqlSporring1 += $", merknad='{merknad}', timepreferanse='{timepreferanse}' WHERE epost = '{epost}'"
        Dim sql1 As New MySqlCommand(sqlSporring1, tilkobling)
        sql1.Parameters.Add("datotime", MySqlDbType.DateTime).Value = blodgiveren.Siste_blodtapping1
        sql1.Parameters.AddWithValue("@blod", If(String.IsNullOrEmpty(blodtype), DBNull.Value, blodtype))
        Dim da1 As New MySqlDataAdapter
        Dim interntabell1 As New DataTable
        'Objektet "da" utfører spørringen og legger resultatet i "interntabell1"
        da1.SelectCommand = sql1
        da1.Fill(interntabell1)
        Me.Cursor = Cursors.Default

        BlodgiverObjOppdat(epost, passord, fornavn, etternavn, adresse, postnr, telefon1, telefon2,
                           statuskode, fodselsnummer, blodtype, siste_blodtapping, kontaktform, merknad, timepreferanse)

    End Sub

    'Oppdaterer blodgiverobjektet blodgiveren
    Private Sub BlodgiverObjOppdat(ByVal epost As String, ByVal passord As String,
                                  ByVal fornavn As String, ByVal etternavn As String,
                                  ByVal adresse As String, ByVal postnr As String,
                                  ByVal telefon1 As String, ByVal telefon2 As String,
                                  ByVal statuskode As Integer,
                                  ByVal fodselsnummer As String, ByVal blodtype As String,
                                  ByVal siste_blodtapping As Date, ByVal kontaktform As String,
                                  ByVal merknad As String, ByVal timepreferanse As String)
        Dim sqlSporring1 As String = $"SELECT beskrivelse FROM personstatus WHERE kode ={statuskode}"
        Dim sql1 As New MySqlCommand(sqlSporring1, tilkobling)
        Dim da1 As New MySqlDataAdapter
        Dim interntabell1 As New DataTable
        'Objektet "da" utfører spørringen og legger resultatet i "interntabell1"
        da1.SelectCommand = sql1
        da1.Fill(interntabell1)
        Me.Cursor = Cursors.Default
        If interntabell1.Rows.Count = 1 Then
            Dim rad1() As DataRow = interntabell1.Select

            blodgiveren.Epost1 = epost
            blodgiveren.Passord1 = passord
            blodgiveren.Fornavn1 = fornavn
            blodgiveren.Etternavn1 = etternavn
            blodgiveren.Adresse1 = adresse
            blodgiveren.Postnr1 = postnr
            blodgiveren.Telefon11 = telefon1
            blodgiveren.Telefon21 = telefon2
            blodgiveren.Status1 = rad1(0)("beskrivelse")
            blodgiveren.Fodselsnummer1 = fodselsnummer
            blodgiveren.Blodtype1 = blodtype
            blodgiveren.Siste_blodtapping1 = siste_blodtapping
            blodgiveren.Kontaktform1 = kontaktform
            blodgiveren.Merknad1 = merknad
            blodgiveren.Timepreferanse1 = timepreferanse
        End If

    End Sub

    'Registrer ny blodgiver
    Private Sub BtnRegBlodgiver_Click(sender As Object, e As EventArgs) Handles BtnRegBlodgiver.Click
        Try
            tilkobling.Open()
            Dim spoerring As String = ""
            If bgRegSkjemadata_OK(txtBgInn_fornavn.Text, txtBgInn_etternavn.Text,
                                  txtBgInn_personnr.Text, txtBgInn_poststed.Text,
                                  txtBgInn_tlfnr.Text, txtBgInn_tlfnr2.Text,
                                  txtBgInn_epost.Text, txtBgInn_passord1.Text,
                                  txtBgInn_passord2.Text, "hvilkenSomHelstStreng") Then

                spoerring = $"INSERT INTO bruker VALUES ('{txtBgInn_epost.Text}', '{txtBgInn_passord1.Text}'"
                spoerring = spoerring & $", '{txtBgInn_fornavn.Text}', '{txtBgInn_etternavn.Text}', '{txtBgInn_adresse.Text}'"
                spoerring = spoerring & $", '{txtBgInn_postnr.Text}', '{txtBgInn_tlfnr.Text}', '{txtBgInn_tlfnr2.Text}'"
                spoerring = spoerring & $", '11')"
                Dim sql1 As New MySqlCommand(spoerring, tilkobling)
                Dim da1 As New MySqlDataAdapter
                Dim interntabell1 As New DataTable
                'Objektet "da" utfører spørringen og legger resultatet i "interntabell1"
                da1.SelectCommand = sql1
                da1.Fill(interntabell1)

                'Legger inn ny rad i tabellen "blodgiver":

                spoerring = $"INSERT INTO blodgiver (epost, fodselsnummer, kontaktform, siste_blodtapping) VALUES ('{txtBgInn_epost.Text}', 
                    '{txtBgInn_personnr.Text}', 'Epost', @dummyDato)"
                Dim sql2 As New MySqlCommand(spoerring, tilkobling)
                sql2.Parameters.Add("dummyDato", MySqlDbType.DateTime).Value = dummyDato
                Dim da2 As New MySqlDataAdapter
                Dim interntabell2 As New DataTable
                da2.SelectCommand = sql2
                da2.Fill(interntabell2)
                MsgBox("Skjema Ok! Nå kan du logge deg på.")
                NullstillPålogging()
            Else
                MsgBox("Skjema dessverre ikke ok.")
            End If
        Catch ex As MySqlException
            MsgBox(ex.Message)
        Finally
            tilkobling.Close()
        End Try

    End Sub

    'Skjemavalidering
    Private Function bgRegSkjemadata_OK(ByVal fornavnInn As String, ByVal etternavnInn As String,
                                        ByVal personnrInn As String, ByVal poststedInn As String,
                                        ByVal telefon1Inn As String, ByVal telefon2Inn As String,
                                        ByVal epostInn As String, ByVal passord1Inn As String,
                                        ByVal passord2Inn As String, ByVal kontaktformInn As String) As Boolean

        Dim sqlSporring1 As String = "SELECT epost FROM bruker WHERE epost = @eposten"
        Dim sql1 As New MySqlCommand(sqlSporring1, tilkobling)
        sql1.Parameters.AddWithValue("@eposten", epostInn)
        Dim da1 As New MySqlDataAdapter
        Dim interntabell1 As New DataTable
        'Objektet "da" utfører spørringen og legger resultatet i "interntabell1"
        da1.SelectCommand = sql1
        da1.Fill(interntabell1)

        Dim sqlSporring2 As String = "SELECT fodselsnummer FROM blodgiver WHERE fodselsnummer = @fnr"
        Dim sql2 As New MySqlCommand(sqlSporring2, tilkobling)
        sql2.Parameters.AddWithValue("@fnr", personnrInn)
        Dim da2 As New MySqlDataAdapter
        Dim interntabell2 As New DataTable
        'Objektet "da" utfører spørringen og legger resultatet i "interntabell2"
        da2.SelectCommand = sql2
        da2.Fill(interntabell2)

        If fornavnInn = "" Or etternavnInn = "" Or telefon1Inn = "" Or kontaktformInn = "" Then
            MsgBox("Alle felt må være utfylt unntatt gateadresse- og telefon 2-feltet må være utfylt.", MsgBoxStyle.Critical)
            Return False
        End If

        If Not IsNumeric(personnrInn) Or personnrInn.Length <> 11 Then
            MsgBox("Fødselsnummeret inneholder ikke bare tall, eller består ikke av 11 siffer.", MsgBoxStyle.Critical)
            Return False
        End If

        Dim aar As String
        Dim aarstall As Integer = CInt(personnrInn.Substring(4, 2))
        If aarstall < CInt(aarstallet) Then
            aar = $"20{aarstall}"
        Else
            aar = $"19{aarstall}"
        End If
        Dim fnrDato As String = $"#{personnrInn.Substring(0, 2)}/{personnrInn.Substring(2, 2)}/{aar}#"
        If Not IsDate(fnrDato) Then
            MsgBox($"De seks første tallene i fødselsnummeret, {fnrDato}, ble ikke gjenkjent som en dato.", MsgBoxStyle.Critical)
            Return False
        End If

        If interntabell2.Rows.Count = 1 And personnrInn <> dummyFodselsnr Then
            MsgBox("Fødselsnummeret finnes fra før. Er du allerede registrert, så logg deg på i skjemaet til høyre.", MsgBoxStyle.Critical)
            Return False
        End If

        If poststedInn = "" Then
            MsgBox("Du har tastet inn feil postnummer. Sjekk at poststed kommer opp i det grå feltet ved siden av postnummeret.", MsgBoxStyle.Critical)
            Return False
        End If

        If Not IsNumeric(telefon1Inn) Or telefon1Inn.Length <> 8 Then
            MsgBox("Telefonnummeret ble ikke akseptert.", MsgBoxStyle.Critical)
            Return False
        End If

        If telefon2Inn <> "" Then
            If Not IsNumeric(telefon2Inn) Or telefon2Inn.Length <> 8 Then
                MsgBox("Telefonnummer2 ble ikke akseptert.", MsgBoxStyle.Critical)
                Return False
            End If
        End If

        If interntabell1.Rows.Count = 1 And epostInn <> dummyEpost Then
            MsgBox("Epostadressen finnes fra før. Er du allerede registrert, så logg deg på i skjemaet til høyre.", MsgBoxStyle.Critical)
            Return False
        End If

        If epostInn.IndexOf("@") = -1 Or epostInn.IndexOf(".") = -1 Then
            MsgBox("Epostadressen ble ikke gjenkjent som en epostadresse.", MsgBoxStyle.Critical)
            Return False
        End If

        If Not passordSjekk(passord1Inn, passord2Inn) Then
            Return False
        End If

        Return True

    End Function

    Private Function passordSjekk(ByVal p1Inn As String, p2Inn As String) As Boolean
        If p1Inn <> p2Inn Then
            MsgBox("Passordene er ikke like. Prøv igjen!", MsgBoxStyle.Critical)
            Return False
        End If
        If p1Inn.Length < 6 Or p1Inn.IndexOf(" ") <> -1 Then
            MsgBox("Passordet må ha minst 6 tegn og ingen mellomrom. Prøv igjen!", MsgBoxStyle.Critical)
            Return False
        End If
        Return True
    End Function

    'Slår på visning av objektene for å sette nytt passord
    Private Sub btnPersDataSettNyttPassord_Click(sender As Object, e As EventArgs) Handles btnPersDataSettNyttPassord.Click
        btnPersDataLagreEndringer.Visible = False
        btnPersDataSettNyttPassord.Visible = False
        lblGmlPassord.Visible = True
        lblNyttPassord.Visible = True
        lblNyttPassordGjenta.Visible = True
        txtGmlPassord.Visible = True
        txtNyttPassord.Visible = True
        txtNyttPassordGjenta.Visible = True
        btnLagreNyttPassord.Visible = True
        btnAvbrytNyttPassord.Visible = True
    End Sub

    'Avbryter setting av nytt passord og gjør om visningene
    Private Sub btnAvbrytNyttPassord_Click(sender As Object, e As EventArgs) Handles btnAvbrytNyttPassord.Click
        btnPersDataLagreEndringer.Visible = True
        btnPersDataSettNyttPassord.Visible = True
        lblGmlPassord.Visible = False
        lblNyttPassord.Visible = False
        lblNyttPassordGjenta.Visible = False
        txtGmlPassord.Visible = False
        txtNyttPassord.Visible = False
        txtNyttPassordGjenta.Visible = False
        btnLagreNyttPassord.Visible = False
        btnAvbrytNyttPassord.Visible = False
    End Sub

    'Lagrer nytt passord for blodgiver
    Private Sub btnLagreNyttPassord_Click(sender As Object, e As EventArgs) Handles btnLagreNyttPassord.Click
        If txtGmlPassord.Text = blodgiveren.Passord1 Then
            If passordSjekk(txtNyttPassord.Text, txtNyttPassordGjenta.Text) Then
                'blodgiveren.Passord1 = txtNyttPassord.Text
                Me.Cursor = Cursors.WaitCursor
                tilkobling.Open()
                Dim sqlSporring1 As String = $"SELECT kode FROM personstatus WHERE beskrivelse ='{txtPersDataGStatus.Text}'"
                Dim sql1 As New MySqlCommand(sqlSporring1, tilkobling)
                Dim da1 As New MySqlDataAdapter
                Dim interntabell1 As New DataTable
                'Objektet "da" utfører spørringen og legger resultatet i "interntabell1"
                da1.SelectCommand = sql1
                da1.Fill(interntabell1)
                tilkobling.Close()
                Me.Cursor = Cursors.Default
                If interntabell1.Rows.Count = 1 Then
                    Dim rad1() As DataRow = interntabell1.Select
                    Me.Cursor = Cursors.WaitCursor
                    tilkobling.Open()

                    OppdaterBlodgiver(blodgiveren.Epost1, txtNyttPassord.Text, blodgiveren.Fornavn1,
                                  blodgiveren.Etternavn1, blodgiveren.Adresse1, blodgiveren.Postnr1,
                                  blodgiveren.Telefon11, blodgiveren.Telefon21, rad1(0)("kode"),
                                  blodgiveren.Fodselsnummer1, blodgiveren.Blodtype1, blodgiveren.Siste_blodtapping1,
                                  blodgiveren.Kontaktform1, blodgiveren.Merknad1, blodgiveren.Timepreferanse1)
                    tilkobling.Close()
                    Me.Cursor = Cursors.Default

                    btnPersDataLagreEndringer.Visible = True
                    btnPersDataSettNyttPassord.Visible = True
                    lblGmlPassord.Visible = False
                    lblNyttPassord.Visible = False
                    lblNyttPassordGjenta.Visible = False
                    txtGmlPassord.Visible = False
                    txtNyttPassord.Visible = False
                    txtNyttPassordGjenta.Visible = False
                    btnLagreNyttPassord.Visible = False
                    btnAvbrytNyttPassord.Visible = False
                    MsgBox("Nytt passord ble satt.", MsgBoxStyle.Information)
                End If
            End If
        Else
            MsgBox("Du tastet inn feil gammelt passord. Prøv igjen!", MsgBoxStyle.Critical)
        End If


    End Sub

    'Lagrer ny personinformasjon satt av blodgiver
    Private Sub btnPersDataLagreEndringer_Click(sender As Object, e As EventArgs) Handles btnPersDataLagreEndringer.Click
        Dim epost As String
        If txtPersDataEpost.Text <> blodgiveren.Epost1 Or txtPersDataGateAdr.Text <> blodgiveren.Adresse1 Or txtPersDataPostnr.Text <> blodgiveren.Postnr1 Or txtPersDataTlf.Text <> blodgiveren.Telefon11 Or txtPersDataTlf2.Text <> blodgiveren.Telefon21 Or CBxKontaktform.Text <> blodgiveren.Kontaktform1 Then
            If txtPersDataEpost.Text = blodgiveren.Epost1 Then
                epost = dummyEpost
            Else
                epost = txtPersDataEpost.Text
            End If
            If bgRegSkjemadata_OK(blodgiveren.Fornavn1, blodgiveren.Etternavn1, dummyFodselsnr, txtPersDataPoststed.Text,
                                  txtPersDataTlf.Text, txtPersDataTlf2.Text, epost,
                                  blodgiveren.Passord1, blodgiveren.Passord1, blodgiveren.Kontaktform1) Then
                Me.Cursor = Cursors.WaitCursor
                tilkobling.Open()
                Dim sqlSporring1 As String = $"SELECT kode FROM personstatus WHERE beskrivelse ='{txtPersDataGStatus.Text}'"
                Dim sql1 As New MySqlCommand(sqlSporring1, tilkobling)
                Dim da1 As New MySqlDataAdapter
                Dim interntabell1 As New DataTable
                'Objektet "da" utfører spørringen og legger resultatet i "interntabell1"
                da1.SelectCommand = sql1
                da1.Fill(interntabell1)
                tilkobling.Close()
                Me.Cursor = Cursors.Default
                If interntabell1.Rows.Count = 1 Then
                    Dim rad1() As DataRow = interntabell1.Select
                    Me.Cursor = Cursors.WaitCursor
                    tilkobling.Open()

                    OppdaterBlodgiver(txtPersDataEpost.Text, blodgiveren.Passord1, blodgiveren.Fornavn1,
                              blodgiveren.Etternavn1, txtPersDataGateAdr.Text, txtPersDataPostnr.Text,
                              txtPersDataTlf.Text, txtPersDataTlf2.Text, rad1(0)("kode"),
                              blodgiveren.Fodselsnummer1, blodgiveren.Blodtype1, blodgiveren.Siste_blodtapping1,
                              CBxKontaktform.Text, blodgiveren.Merknad1, blodgiveren.Timepreferanse1)
                    tilkobling.Close()
                    Me.Cursor = Cursors.Default
                    MsgBox("Informasjonen ble oppdatert.", MsgBoxStyle.Information)
                Else
                    MsgBox("Noe gikk feil under søk etter statuskode.", MsgBoxStyle.Critical)
                End If
            End If
        Else
            MsgBox("Ingen endringer av personinformasjonen ble funnet.", MsgBoxStyle.Information)
        End If
    End Sub

    'Logg av blodgiver
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles BttnLoggavGiver.Click
        Dim resultat As Object
        Dim spørsmål As String = "Du har gjort endringer i den personlige informasjonen din som ikke er lagret. Ønsker du å lagre disse før du logger av?"
        Dim tittel As String = "Ulagrede endringer oppdaget"
        Dim loggAv As Boolean = True
        Dim lagre As Boolean = False
        If txtPersDataEpost.Text <> blodgiveren.Epost1 Or txtPersDataGateAdr.Text <> blodgiveren.Adresse1 Or
            txtPersDataPostnr.Text <> blodgiveren.Postnr1 Or txtPersDataTlf.Text <> blodgiveren.Telefon11 Or
            txtPersDataTlf2.Text <> blodgiveren.Telefon21 Or CBxKontaktform.Text <> blodgiveren.Kontaktform1 Then
            resultat = MsgBox(spørsmål, 3, tittel)
            Select Case resultat
                Case 6
                    lagre = True
                    loggAv = False
                Case 7
                Case Else
                    loggAv = False
            End Select
        End If
        If lagre Then
            MsgBox("Se over personinformasjonen din og lagre den før du logger av.", MsgBoxStyle.Information)
        End If
        If loggAv Then
            'Bytter til panelet for pålogging
            NullstillPålogging()
            PanelGiver.Hide()
            PanelAnsatt.Hide()
            PanelPåmelding.Show()
            PanelPåmelding.BringToFront()
            LoggPåansattToolStripMenuItem.Visible = True
            LoggAvToolStripMenuItem.Visible = False
        End If

    End Sub

    'Knapp - Logg av ansatt
    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles BttnLoggavAnsatt.Click
        loggAvAnsatt()
    End Sub

    'Filmeny - Logg av ansatt
    Private Sub LoggAvToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoggAvToolStripMenuItem.Click
        loggAvAnsatt()
    End Sub

    'Sub for å loogge av ansatt
    Private Sub loggAvAnsatt()
        lblVelkommen.Text = ""
        PanelGiver.Hide()
        PanelAnsatt.Hide()
        PanelPåmelding.Show()
        PanelPåmelding.BringToFront()
        LoggPåansattToolStripMenuItem.Visible = True
        LoggAvToolStripMenuItem.Visible = False
    End Sub

    'Nullstiller tekstfeltene på påloggingssiden
    Private Sub NullstillPålogging()
        txtBgInn_adresse.Text = ""
        txtBgInn_epost.Text = ""
        txtBgInn_etternavn.Text = ""
        txtBgInn_fornavn.Text = ""
        txtBgInn_passord1.Text = ""
        txtBgInn_passord2.Text = ""
        txtBgInn_personnr.Text = ""
        txtBgInn_postnr.Text = ""
        txtBgInn_tlfnr.Text = ""
        txtBgInn_tlfnr2.Text = ""
        txtAInn_epost.Text = ""
        txtAInn_passord.Text = ""
    End Sub

    'Blodgiversøk knapp
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles BttnSøkGiver.Click
        Me.Cursor = Cursors.WaitCursor
        Dim personnummer As String = txtSøk.Text
        Dim status As String = txtSøkStatuskode.Text
        Dim statuskode As Integer
        Dim blodtype As String = cBxSøkBlodtype.Text
        bgSøkParameter = ""
        If status = "" Then
            statuskode = 0
        Else
            statuskode = CInt(status)
        End If
        If (personnummer = "") And (status = "") And (blodtype = "") Then
            Exit Sub
        ElseIf (personnummer <> "") And (Not IsNumeric(personnummer)) Then
            MsgBox("Feil i søkeparametrene")
            Exit Sub
        End If
        If (personnummer <> "") And (statuskode = 0) And (blodtype = "") Then
            bgSøkParameter = $" bl.fødselsnummer = '{personnummer}'"
        ElseIf (statuskode > 0) And (personnummer = "") And (blodtype = "") Then
            bgSøkParameter = $" br.statuskode = '{statuskode}'"
        ElseIf (blodtype <> "") And (statuskode = 0) And (personnummer = "") Then
            bgSøkParameter = $" bl.blodtype = '{blodtype}'"
        ElseIf (blodtype <> "") And (statuskode > 0) And (personnummer = "") Then
            bgSøkParameter = $" bl.blodtype = '{blodtype}' and br.statuskode = '{statuskode}'"
        ElseIf (personnummer <> "") And (statuskode > 0) And (blodtype <> "") Then
            bgSøkParameter = $" bl.blodtype = '{blodtype}' and br.statuskde = '{statuskode}' and bl.fødselsnummer = '{personnummer}'"
        End If
        bgSøk(bgSøkParameter)
        Me.Cursor = Cursors.Default
        giverSøkTreff()
    End Sub

    'Vis treff av blodgiversøk i listebox
    Private Sub giverSøkTreff()
        Dim resPnr, resFnavn, resEnavn, resStatus, resKode As String
        Dim rad As DataRow

        lBxSøkResultater.Items.Clear()
        For Each rad In giversøk.Rows
            resPnr = rad("fodselsnummer")
            resFnavn = rad("fornavn")
            resEnavn = rad("etternavn")
            resStatus = rad("beskrivelse")
            resKode = rad("statuskode")
            lBxSøkResultater.Items.Add($"{resPnr} {vbTab}{resFnavn} {resEnavn} {vbTab}{resKode} - {resStatus}")
        Next
        If lBxSøkResultater.Items.Count > 0 Then
            lBxSøkResultater.SetSelected(0, True)
        End If
    End Sub

    'SQL - søk frem blodgiver og egenerklæring
    Private Sub bgSøk(ByVal streng As String)
        Dim sqlStreng As String
        Dim da As New MySqlDataAdapter
        giversøk.Clear()
        egenerklaering.Clear()
        Try
            tilkobling.Open()
            sqlStreng = "SELECT * FROM bruker br INNER JOIN blodgiver bl ON br.epost = bl.epost INNER JOIN personstatus ps ON ps.kode = br.statuskode WHERE"
            Dim sqlSpørring As New MySqlCommand($"{sqlStreng}{streng}", tilkobling)
            da.SelectCommand = sqlSpørring
            da.Fill(giversøk)

            '   Dim sqlSpørring2 As New MySqlCommand("SELECT * FROM egenerklaering", tilkobling)
            '  da.SelectCommand = sqlSpørring2
            ' da.Fill(egenerklaering)
        Catch
            MsgBox("Får ikke kontakt med databasen")
            Exit Sub
        End Try
        tilkobling.Close()
    End Sub

    'Sett rett statuskode i textboks
    Private Sub statuskode(ByVal beskrivelse As String, ByVal utput As Object, input As Object)
        utput.text = personstatusK(beskrivelse)
    End Sub

    'Sett rett statusbeskrivelse i combobox
    Private Sub statusbeskrivelse(ByVal kode As String, ByVal utput As Object, input As Object)
        utput.text = personstatusB(kode)
    End Sub

    'Endring av statusbeskrvelse - henter statuskode
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cBxSøkStatusbeskrivelse.SelectedIndexChanged
        Try
            statuskode(cBxSøkStatusbeskrivelse.SelectedItem, txtSøkStatuskode, cBxSøkStatusbeskrivelse)
        Catch
            txtSøkStatuskode.Text = ""
            cBxSøkStatusbeskrivelse.Text = ""
            Exit Sub
        End Try
    End Sub

    'Endring av statuskode - henter statusbeskrivelse
    Private Sub TextBox20_TextChanged(sender As Object, e As EventArgs) Handles txtSøkStatuskode.TextChanged
        Try
            statusbeskrivelse(txtSøkStatuskode.Text, cBxSøkStatusbeskrivelse, txtSøkStatuskode)
        Catch
            txtSøkStatuskode.Text = ""
            cBxSøkStatusbeskrivelse.Text = ""
            Exit Sub
        End Try
    End Sub

    'Presenter valgt person i blodgiversøk
    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lBxSøkResultater.SelectedIndexChanged
        visBG()
    End Sub

    'Vis blodgiverdata til ansatt
    Private Sub visBG()
        Dim index, i As Integer
        Dim rad, rad1() As DataRow
        Dim fNavn, eNavn, fnr, epost, adresse, postnummmer, tlf1, tlf2, intMerknad, preferanse, jasvar, erklaringLege As String
        Dim status As Integer
        Dim sistTapping, sistErklæring, gjennomgåttErklæring As Date
        Dim dager As Long
        jasvar = ""
        sistErklæring = dummyDato
        lbxHKtrlJasvar.Items.Clear()
        index = lBxSøkResultater.SelectedIndex
        rad1 = giversøk.Select

        If IsDBNull(rad1(index)("blodtype")) Then
            rad1(index)("blodtype") = ""
        End If
        If IsDBNull(rad1(index)("merknad")) Then
            rad1(index)("merknad") = ""
        End If
        If IsDBNull(rad1(index)("timepreferanse")) Then
            rad1(index)("timepreferanse") = ""
        End If
        If IsDBNull(rad1(index)("adresse")) Then
            rad1(index)("adresse") = ""
        End If
        If IsDBNull(rad1(index)("telefon2")) Then
            rad1(index)("telefon2") = ""
        End If
        If IsDBNull(rad1(index)("siste_blodtapping")) Then
            rad1(index)("siste_blodtapping") = dummyDato
        End If
        BlodgiverObjOppdat(rad1(index)("epost"), rad1(index)("passord"), rad1(index)("fornavn"),
                           rad1(index)("etternavn"), rad1(index)("adresse"), rad1(index)("postnr"),
                           rad1(index)("telefon1"), rad1(index)("telefon2"), rad1(index)("statuskode"),
                           rad1(index)("fodselsnummer"), rad1(index)("blodtype"), rad1(index)("siste_blodtapping"),
                           rad1(index)("kontaktform"), rad1(index)("merknad"), rad1(index)("timepreferanse"))
        '  For Each rad In giversøk.Rows
        ' fNavn = rad("fornavn")
        'eNavn = rad("etternavn")
        'fnr = rad("fodselsnummer")
        'epost = rad("epost")
        'adresse = rad("adresse")
        'tlf1 = rad("telefon1")
        'tlf2 = rad("telefon2")
        'postnummmer = rad("postnr")
        'status = rad("statuskode")
        'sistTapping = rad("siste_blodtapping")
        'intMerknad = rad("merknad")
        'preferanse = rad("timepreferanse")
        'If i = index Then
        'Exit For
        'End If
        'i = i + 1
        'Next
        'For Each rad In egenerklaering.Rows
        'If rad("bgepost") = blodgiveren.Epost1 Then
        'If rad("datotidbg") > sistErklæring Then
        'sistErklæring = rad("datotidbg")
        'jasvar = rad("skjema")
        'egenerklæringID = rad("id")
        'presentertGiver = rad("bgepost")
        'If Not IsDBNull(rad("ansattepost")) Then
        'erklaringLege = rad("ansattepost")
        'Else
        'erklaringLege = ""
        'End If
        'If Not IsDBNull(rad("datotidansatt")) Then
        'gjennomgåttErklæring = rad("datotidansatt")
        'Else
        'gjennomgåttErklæring = Nothing
        'End If
        'End If
        'End If
        'Next
        tilkobling.Open()
        Dim sqlSpørring = $"SELECT * FROM egenerklaering WHERE bgepost='{blodgiveren.Epost1}' ORDER BY datotidbg DESC"
        Dim sql1 As New MySqlCommand(sqlSpørring, tilkobling)
        Dim da1 As New MySqlDataAdapter
        Dim interntabell1 As New DataTable
        da1.SelectCommand = sql1
        da1.Fill(interntabell1)
        tilkobling.Close()
        If interntabell1.Rows.Count > 0 Then

            Dim rad2() As DataRow = interntabell1.Select()
            If IsDBNull(rad2(0)("ansattepost")) Then
                rad2(0)("ansattepost") = dummyEpost
            End If

            egenerklaeringObjekt.Id1 = rad2(0)("id")
            egenerklaeringObjekt.BgEpost1 = rad2(0)("bgepost")
            egenerklaeringObjekt.AnsattEpost1 = rad2(0)("ansattepost")
            egenerklaeringObjekt.Skjema1 = rad2(0)("skjema")
            egenerklaeringObjekt.Kommentar1 = rad2(0)("kommentar")
            egenerklaeringObjekt.DatotidBG1 = rad2(0)("datotidbg")
            egenerklaeringObjekt.DatotidAnsatt1 = rad2(0)("datotidansatt")

        End If
        'jasvar = rad1(index)("skjema")
        'egenerklæringID = rad1(index)("id")
        'presentertGiver = blodgiveren.Epost1
        'If Not IsDBNull(rad("ansattepost")) Then
        ' erklaringLege = rad("ansattepost")
        'Else
        'erklaringLege = ""
        'End If
        'If Not IsDBNull(rad("datotidansatt")) Then
        'gjennomgåttErklæring = rad("datotidansatt")
        'Else
        'gjennomgåttErklæring = Nothing
        'End If
        GroupBoxIntervju.Visible = False

        dager = DateDiff(DateInterval.DayOfYear, blodgiveren.Siste_blodtapping1, Today)
        If egenerklaeringObjekt.Skjema1 <> "" Then
            utledJAsvar(egenerklaeringObjekt.Skjema1)
        End If

        txtValgtBlodgiverNavn.Text = $"{blodgiveren.Fornavn1} {blodgiveren.Etternavn1}"
        txtValgtBlodgiverPersnr.Text = blodgiveren.Fodselsnummer1
        txtValgtBlodgiverEpost.Text = blodgiveren.Epost1
        txtValgtBlodgiverTelefon1.Text = blodgiveren.Telefon11
        txtValgtBlodgiverTelefon2.Text = blodgiveren.Telefon21
        txtValgtBlodgiverAdresse.Text = blodgiveren.Adresse1
        txtValgtBlodgiverPostnr.Text = blodgiveren.Postnr1
        cBxValgtBlodgiverStatusTekst.Text = blodgiveren.Status1
        rTxtValgBlodgiverTimepref.Text = blodgiveren.Timepreferanse1
        rTxtValgtBlodgiverInternMrknd.Text = blodgiveren.Merknad1

        If blodgiveren.Siste_blodtapping1 <> dummyDato Then
            GroupBoxIntervju.Visible = True
            txtValgtBlodgiverSistTappDato.Text = blodgiveren.Siste_blodtapping1
            txtValgtBlodgiverSistTappDager.Text = $"{dager} dager"
            txtHKtrlSisteEgenerkl.Text = egenerklaeringObjekt.DatotidBG1
            If egenerklaeringObjekt.AnsattEpost1 <> dummyEpost Then
                txtHKtrlGjennomgAv.Text = egenerklaeringObjekt.AnsattEpost1
                txtHKtrlEKDatoGjennomg.Text = egenerklaeringObjekt.DatotidAnsatt1
            Else
                txtHKtrlGjennomgAv.Text = ""
                txtHKtrlEKDatoGjennomg.Text = ""
            End If
        Else
            txtValgtBlodgiverSistTappDato.Text = "Ikke gitt blod enda"
            txtValgtBlodgiverSistTappDager.Text = ""
            txtHKtrlSisteEgenerkl.Text = ""
            txtHKtrlGjennomgAv.Text = ""
            txtHKtrlEKDatoGjennomg.Text = ""
            lbxHKtrlJasvar.Items.Clear()
            MsgBox($"Blodgiver {blodgiveren.Epost1} har ikke fylt ut noen egenerklæring ennå.", MsgBoxStyle.Exclamation)
        End If

    End Sub

    'Utleder Jasvar og presenterer i Listebox i giversøk
    Private Sub utledJAsvar(ByVal spmNr As String)
        Dim svar() As String = spmNr.Split(", ")
        For i = 0 To svar.Length - 1
            lbxHKtrlJasvar.Items.Add(svar(i))
        Next
    End Sub

    'Endring av statuskode - henter statusbeskrivelse
    Private Sub TextBox21_TextChanged(sender As Object, e As EventArgs) Handles txtValgtBlodgiverStatusKode.TextChanged
        Try
            statusbeskrivelse(txtValgtBlodgiverStatusKode.Text, cBxValgtBlodgiverStatusTekst, txtValgtBlodgiverStatusKode)
        Catch
            txtValgtBlodgiverStatusKode.Text = ""
            cBxValgtBlodgiverStatusTekst.Text = ""
            Exit Sub
        End Try
    End Sub

    'Endring av statusbeskrvelse - henter statuskode
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cBxValgtBlodgiverStatusTekst.SelectedIndexChanged
        Try
            statuskode(cBxValgtBlodgiverStatusTekst.SelectedItem, txtValgtBlodgiverStatusKode, cBxValgtBlodgiverStatusTekst)
        Catch
            txtValgtBlodgiverStatusKode.Text = ""
            cBxValgtBlodgiverStatusTekst.Text = ""
            Exit Sub
        End Try
    End Sub

    'Sett rett poststed ved siden av postnummer i personsøk
    Private Sub TextBox31_TextChanged(sender As Object, e As EventArgs) Handles txtValgtBlodgiverPostnr.TextChanged
        txtValgtBlodgiverPoststed.Text = postnummer(txtValgtBlodgiverPostnr.Text)
    End Sub

    'Sett rett poststed ved siden av postnummer i egenregistrering
    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles txtBgInn_postnr.TextChanged
        txtBgInn_poststed.Text = postnummer(txtBgInn_postnr.Text)
    End Sub

    'Tøm giversøk
    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles btnSøkTømSkjema.Click
        txtSøk.Text = ""
        cBxSøkBlodtype.Text = ""
        txtSøkStatuskode.Text = ""
    End Sub

    'Innkallingsoversikt
    Private Sub TabPage3_Enter(sender As Object, e As EventArgs) Handles TabPage3.Enter
        Me.Cursor = Cursors.WaitCursor
        Dim rad As DataRow
        Dim da As New MySqlDataAdapter
        Dim epost As String
        Dim dato As Date
        Dim romnr, etg As Integer
        Dim sqlSpørring As New MySqlCommand("SELECT * FROM timeavtale t INNER JOIN rom r ON t.romnr = r.romnr ORDER BY datotid DESC", tilkobling)
        innkalling.Clear()
        da.SelectCommand = sqlSpørring
        da.Fill(innkalling)
        ListBox4.Items.Clear()
        ListBox5.Items.Clear()
        ListBox6.Items.Clear()
        Me.Cursor = Cursors.Default
        For Each rad In innkalling.Rows
            dato = DateValue(rad("datotid"))
            epost = rad("bgepost")
            romnr = rad("romnr")
            etg = rad("etasje")
            If dato = Today Then
                dato = rad("datotid")
                ListBox4.Items.Add($"{dato} - Rom:  {romnr} - Etg: {etg} - Giver: {epost}")
            ElseIf dato = DateAdd(DateInterval.Day, 1, Today) Then
                dato = rad("datotid")
                ListBox5.Items.Add($"{dato} - Rom: {romnr} - Etg: {etg} - Giver: {epost}")
            ElseIf dato < Today Then
                dato = rad("datotid")
                ListBox6.Items.Add($"{dato} - Rom: {romnr} - Etg: {etg} - Giver: {epost}")
            End If
        Next
        tilkobling.Close()
    End Sub

    'Blodlager
    Private Sub TabPage2_Enter(sender As Object, e As EventArgs) Handles TabPage2.Enter
        Dim B_legemer, B_plater, B_plasma As Integer
        Me.Cursor = Cursors.WaitCursor
        blodlager.Clear()
        Dim rad As DataRow
        Dim da As New MySqlDataAdapter
        Dim sqlSpørring As New MySqlCommand("SELECT * FROM blodprodukt b INNER JOIN timeavtale t ON b.timeid = t.timeid INNER JOIN blodgiver bl on t.bgepost = bl.epost", tilkobling)
        da.SelectCommand = sqlSpørring
        da.Fill(blodlager)
        Me.Cursor = Cursors.Default

        For Each rad In blodlager.Rows
            If (rad("produkttypeid") = 1) And (rad("statusid") = 1) Then
                B_legemer = B_legemer + (rad("antall"))
            ElseIf (rad("produkttypeid") = 2) And (rad("statusid") = 1) Then
                B_plater = B_plater + (rad("antall"))
            ElseIf (rad("produkttypeid") = 3) And (rad("statusid") = 1) Then
                B_plasma = B_plasma + (rad("antall"))
            End If
        Next
        tilkobling.Close()

        Chart1.Series.Clear()
        MsgBox(B_plasma)
        MsgBox(B_plater)
        Chart1.Series(0).Points(0).YValues(0) = B_plasma
        Chart2.Series(0).Points(1).YValues(0) = B_plater

    End Sub

    'Slår av og på visning av gruppeboksen med skjema for å endre avtalt time
    Private Sub BtnEndreInnkalling_Click(sender As Object, e As EventArgs) Handles BtnEndreInnkalling.Click
        LBxLedigeTimer.Items.Clear()
        If GpBxEndreInnkalling.Visible Then
            GpBxEndreInnkalling.Visible = False
        Else
            GpBxEndreInnkalling.Visible = True
        End If
        If TxtNesteInnkalling.Text <> "" Then
            DateTimePickerNyTime.Value = CDate(TxtNesteInnkalling.Text)
        End If
        For i = 0 To fulltimetabell.Count - 1
            LBxLedigeTimer.Items.Add(fulltimetabell(i))
        Next
        BtnBekreftEndretTime.Enabled = False
    End Sub

    'Henter ledige timer for valgt dato
    Private Sub hentLedigeTimer(ByVal aktuelldato As DateTime)

        Dim aktuelldatopluss1 = aktuelldato.AddDays(1)
        tilkobling.Open()

        Dim sqlSporring1 As String = $"SELECT datotid, COUNT(*) AS 'antall' FROM timeavtale WHERE datotid > '{aktuelldato.ToString("yyyy-MM-dd")}' AND datotid < '{aktuelldatopluss1.ToString("yyyy-MM-dd")}' GROUP BY datotid HAVING (antall>{antallRom - 1})"
        Dim sql1 As New MySqlCommand(sqlSporring1, tilkobling)
        Dim da1 As New MySqlDataAdapter
        Dim interntabell1 As New DataTable
        Dim rad1 As DataRow
        Dim antallTimerPåDetteKlokkeslettet As Integer = 0
        Dim tabort As Integer = 0
        Dim opptatt As Boolean = False
        Dim raddato1 As DateTime
        Dim radnr As Integer
        'Objektet "da" utfører spørringen og legger resultatet i "interntabell1"
        da1.SelectCommand = sql1
        da1.Fill(interntabell1)
        tilkobling.Close()
        FulltimeTabellReset()

        For Each rad1 In interntabell1.Rows
            raddato1 = rad1("datotid")
            radnr = raddato1.Hour
            fulltimetabell.RemoveAt(radnr - 8)

        Next

    End Sub

    'Resetter fulltimetabellen
    Private Sub FulltimeTabellReset()
        Dim i, tabellstørrelse As Integer
        tabellstørrelse = fulltimetabell.Count
        For i = 0 To tabellstørrelse - 1
            fulltimetabell.RemoveAt(0)
        Next
        For i = 0 To 7
            fulltimetabell.Add($"{i + 8}:00")
        Next
    End Sub

    'Kaller subrutinen "hentLedigeTimer", som plukker ut ledige timer når dato blir valgt.
    Private Sub DateTimePickerNyTime_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePickerNyTime.ValueChanged
        'Dim sisteUndersøkelse As Date =
        If (DateTimePickerNyTime.Value - CDate(blodgiverData("siste_blodtapping"))).TotalDays < 90 Then
            MsgBox("Det må være minst 90 dager siden siste blodtapping. Velg en ny dato.", MsgBoxStyle.Critical)
        Else
            If Weekday(DateTimePickerNyTime.Value, FirstDayOfWeek.Monday) > 5 Or fridag(DateTimePickerNyTime.Value) Then
                MsgBox($"Ukedagnr: {Weekday(DateTimePickerNyTime.Value, FirstDayOfWeek.Monday)}, Fridag: {fridag(DateTimePickerNyTime.Value)}.")
                MsgBox("Blodbanken er stengt denne dagen. Velg en en ny dag.", MsgBoxStyle.Critical)
            Else
                LblLedigeTimer.Text = $"Ledige timer {DateTimePickerNyTime.Text}"
                hentLedigeTimer(DateTimePickerNyTime.Value)
            End If

        End If
    End Sub

    'Fjerner ValueChanged-eventet når datovelgeren droppes ned
    Private Sub DateTimePickerNyTime_DropDown(ByVal sender As Object, ByVal e As EventArgs) Handles DateTimePickerNyTime.DropDown
        RemoveHandler DateTimePickerNyTime.ValueChanged, AddressOf DateTimePickerNyTime_ValueChanged
    End Sub

    'Slår på ValueChanged-eventet når datovelgeren rulles opp når dato velges, i tillegg til å kalle ValueChange-prosedyren manuelt
    Private Sub DateTimePickerNyTime_CloseUp(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePickerNyTime.CloseUp
        AddHandler DateTimePickerNyTime.ValueChanged, AddressOf DateTimePickerNyTime_ValueChanged
        Call DateTimePickerNyTime_ValueChanged(sender, EventArgs.Empty)
    End Sub

    'Funksjon som regner ut datoen til 1. påskedag
    Public Shared Function GetEasterDate(ByVal Year As Integer) As Date
        Dim a As Integer
        Dim b As Integer
        Dim c As Integer
        Dim d As Integer
        Dim e As Integer
        Dim f As Integer
        Dim g As Integer
        Dim h As Integer
        Dim i As Integer
        Dim k As Integer
        Dim l As Integer
        Dim m As Integer
        Dim n As Integer
        Dim p As Integer

        If Year < 1583 Then

            MsgBox("Årstallet er feil.")
            Return Nothing
        Else

            ' Step 1: Divide the year by 19 and store the
            ' remainder in variable A.  Example: If the year
            ' is 2000, then A is initialized to 5.

            a = Year Mod 19

            ' Step 2: Divide the year by 100.  Store the integer
            ' result in B and the remainder in C.

            b = Year \ 100
            c = Year Mod 100

            ' Step 3: Divide B (calculated above).  Store the
            ' integer result in D and the remainder in E.

            d = b \ 4
            e = b Mod 4

            ' Step 4: Divide (b+8)/25 and store the integer
            ' portion of the result in F.

            f = (b + 8) \ 25

            ' Step 5: Divide (b-f+1)/3 and store the integer
            ' portion of the result in G.

            g = (b - f + 1) \ 3

            ' Step 6: Divide (19a+b-d-g+15)/30 and store the
            ' remainder of the result in H.

            h = (19 * a + b - d - g + 15) Mod 30

            ' Step 7: Divide C by 4.  Store the integer result
            ' in I and the remainder in K.

            i = c \ 4
            k = c Mod 4

            ' Step 8: Divide (32+2e+2i-h-k) by 7.  Store the
            ' remainder of the result in L.

            l = (32 + 2 * e + 2 * i - h - k) Mod 7

            ' Step 9: Divide (a + 11h + 22l) by 451 and
            ' store the integer portion of the result in M.

            m = (a + 11 * h + 22 * l) \ 451

            ' Step 10: Divide (h + l - 7m + 114) by 31.  Store
            ' the integer portion of the result in N and the
            ' remainder in P.

            n = (h + l - 7 * m + 114) \ 31
            p = (h + l - 7 * m + 114) Mod 31

            ' At this point p+1 is the day on which Easter falls.
            ' n is 3 for March or 4 for April.

            Return DateSerial(Year, n, p + 1)
        End If
    End Function

    'Setter rett poststed ved siden av postnummeret i fanen Personinfo for blodgiveren
    Private Sub txtPersDataPostnr_TextChanged(sender As Object, e As EventArgs) Handles txtPersDataPostnr.TextChanged
        txtPersDataPoststed.Text = postnummer(txtPersDataPostnr.Text)
    End Sub

    'Slår på knappen for å bekrefte nytt tidspunkt for neste time.
    Private Sub LBxLedigeTimer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LBxLedigeTimer.SelectedIndexChanged
        BtnBekreftEndretTime.Enabled = True
    End Sub

    'Bekrefter valg av nytt tidspunkt for neste innkalling og legger det inn i timeavtalen i databasen.
    Private Sub BtnBekreftEndretTime_Click(sender As Object, e As EventArgs) Handles BtnBekreftEndretTime.Click

        Dim time_DateTime As DateTime
        Try
            Dim provider As CultureInfo = CultureInfo.InvariantCulture
            time_DateTime = Date.ParseExact(LBxLedigeTimer.SelectedItem, "H:mm", provider)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        bytteRomTime.Datotid1 = New Date(DateTimePickerNyTime.Value.Year, DateTimePickerNyTime.Value.Month, DateTimePickerNyTime.Value.Day,
                               time_DateTime.Hour, 0, 0)

        Me.Cursor = Cursors.WaitCursor
        tilkobling.Open()
        Dim sqlSporring1 As String = $"SELECT * FROM timeavtale WHERE datotid = @nyDatotime"
        Dim sql1 As New MySqlCommand(sqlSporring1, tilkobling)
        sql1.Parameters.Add("nyDatotime", MySqlDbType.DateTime).Value = bytteRomTime.Datotid1
        Dim da1 As New MySqlDataAdapter
        Dim interntabell1 As New DataTable
        Dim rad1, radRom As DataRow

        'Objektet "da" utfører spørringen og legger resultatet i "interntabell1"
        da1.SelectCommand = sql1
        da1.Fill(interntabell1)
        Me.Cursor = Cursors.Default

        Dim antallLedigeRom As Integer = antallRom - interntabell1.Rows.Count
        'MsgBox($"Antall rom totalt: {antallRom}. Antall rader i spørringsresultat: {interntabell1.Rows.Count}")
        If antallLedigeRom = 0 Then
            MsgBox("Dessverre var ikke den valgte timen ledig likevel. Prøv en annen time.")
        Else
            For Each radRom In interntabellRom.Rows
                If antallLedigeRom = antallRom Then
                    bytteRomTime.Romnummer1 = radRom("romnr")
                Else
                    For Each rad1 In interntabell1.Rows
                        If radRom("romnr") <> rad1("romnr") Then
                            bytteRomTime.Romnummer1 = radRom("romnr")
                        End If
                    Next
                End If
            Next
            GpBxEndreInnkalling.Visible = False
            TxtNesteInnkalling.Text = bytteRomTime.Datotid1
            Me.Cursor = Cursors.WaitCursor
            Dim sqlSporring2 As String = $"UPDATE timeavtale SET datotid = @nyDatotime, romnr = {bytteRomTime.Romnummer1} WHERE timeid = {bytteRomTime.Timenr1}"
            Dim sql2 As New MySqlCommand(sqlSporring2, tilkobling)
            sql2.Parameters.Add("nyDatotime", MySqlDbType.DateTime).Value = bytteRomTime.Datotid1
            Dim da2 As New MySqlDataAdapter
            Dim interntabell2 As New DataTable
            MsgBox($"Ny innkallingstime ble satt til {bytteRomTime.Datotid1}")
            'Objektet "da" utfører spørringen og legger resultatet i "interntabell1"
            da2.SelectCommand = sql2
            da2.Fill(interntabell2)

            Me.Cursor = Cursors.Default
        End If
        tilkobling.Close()
    End Sub

    'Registrer gitt blod
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles btnBlodgivingGjført.Click
        Dim ansatt, SettInnSpørring, SøkSpørring, timeID As String
        Dim bPlater, bLegemer, bPlasma, i As Integer
        Dim timedato As Date
        Dim feil As Boolean
        Dim antallProdukt(2) As Integer
        Dim tabell As New DataTable
        Dim da As New MySqlDataAdapter
        timeID = ""
        ansatt = cbxAnsattUtførtTapping.Text
        bPlater = nudResBlodplater.Value
        bPlasma = nudResPlasma.Value
        bLegemer = nudResRødeBlodl.Value
        antallProdukt(1) = bPlater : antallProdukt(2) = bPlasma : antallProdukt(0) = bLegemer
        If ansatt = "" Then
            MsgBox("Registrer hvem som tappet blodet")
            Exit Sub
        End If
        SøkSpørring = $"SELECT * FROM timeavtale t INNER JOIN blodgiver b ON t.bgepost = b.epost WHERE t.bgepost = '{presentertGiver}' ORDER BY datotid DESC"
        Try
            tilkobling.Open()
            Dim sqlSpørring As New MySqlCommand($"{SøkSpørring}", tilkobling)
            da.SelectCommand = sqlSpørring
            da.Fill(tabell)
            i = 0
            For Each rad In tabell.Rows
                timedato = rad("datotid")
                If DateDiff(DateInterval.DayOfYear, timedato, Today) = 0 Then
                    timedato = Today
                    timeID = rad("timeid")
                    feil = False
                    Exit For
                Else
                    feil = True
                End If
                i = i + 1
            Next
            If feil = False Then
                i = 1
                For i = 1 To 3
                    SettInnSpørring = $"INSERT INTO blodprodukt (timeid, produkttypeid, statusid, antall) VALUES ({timeID}, {i}, 1, {antallProdukt(i - 1)} )"
                    Dim sqlSpørring2 As New MySqlCommand($"{SettInnSpørring}", tilkobling)
                    sqlSpørring2.ExecuteNonQuery()
                Next
            Else
                MsgBox("Denne personen har ikke gitt blod i dag")
            End If
            tilkobling.Close()
            nudResBlodplater.Value = 0
            nudResPlasma.Value = 0
            nudResRødeBlodl.Value = 0
            MsgBox("Blodlager oppdatert")
        Catch ex As Exception
            tilkobling.Close()
            MsgBox("Feil")
        End Try
    End Sub

    'Forrige spørsmål i erklæring
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnEgenerklForrigeSpm.Click
        Dim sisteindex, kjønn, spmID, i As Integer
        Dim spmText, pnr As String
        Dim dame As Boolean
        pnr = blodgiveren.Fodselsnummer1
        sisteindex = Erklæringspørsmål.Rows.Count - 1
        spmText = ""

        kjønn = pnr.Substring(8, 1)
        If (kjønn = 0) Or (kjønn = 2) Or (kjønn = 4) Or (kjønn = 6) Or (kjønn = 8) Then
            dame = True
        Else
            dame = False
        End If

        'Forrige spm
        SPMnr = SPMnr - 1
        spmID = Erklæringspørsmål.Rows(SPMnr).Item("Nr")
        For i = sisteindex + 1 To 1 Step -1
            If spmID < 100 Then
                spmText = Erklæringspørsmål.Rows(SPMnr).Item("spoersmaal")
                Exit For
            ElseIf (spmID > 199) And (dame = False) Then
                spmText = Erklæringspørsmål.Rows(SPMnr).Item("spoersmaal")
                Exit For
            ElseIf (spmID > 99) And (dame = True) Then
                spmText = Erklæringspørsmål.Rows(SPMnr).Item("spoersmaal")
                Exit For
            Else
                SPMnr = SPMnr - 1
                spmID = Erklæringspørsmål.Rows(SPMnr).Item("Nr")
                spmText = Nothing
            End If
        Next
        SPMnrPresentert = SPMnrPresentert - 1
        lblEgenerklSpmTekst.Text = spmText
        lblEgenerklSpmNr.Text = $"Spørsmål {SPMnrPresentert + 1}"
        rBtnEgenerklJa.Checked = False
        rBtnEgenerklNei.Checked = False
        btnEgenerklNesteSpm.Enabled = True
        If SPMnr <= 0 Then
            btnEgenerklForrigeSpm.Enabled = False
            Exit Sub
        End If
    End Sub

    'Send inn egenerklæring
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnEgenerklSendInn.Click
        Dim Jasvar, sporring As String
        Dim i, siste As Integer
        siste = erklæringSvar.Length
        Jasvar = ""

        For i = 0 To siste - 1
            If erklæringSvar(i) = 1 Then
                Jasvar = Jasvar & i & ","
            End If
        Next

        Try
            tilkobling.Open()
            sporring = $"INSERT INTO egenerklaering (bgepost, datotidbg, datotidansatt, skjema, kommentar) VALUES ('{blodgiveren.Epost1}', @idag , @dummydato, '{Jasvar}', 'Ingen kommentar')"
            MsgBox(sporring)
            Dim sqlja As New MySqlCommand(sporring, tilkobling)
            sqlja.Parameters.Add("idag", MySqlDbType.DateTime).Value = Now
            sqlja.Parameters.Add("dummydato", MySqlDbType.DateTime).Value = dummyDato
            sqlja.ExecuteNonQuery()
            sporring = $"UPDATE bruker SET statuskode = 32 WHERE epost = '{blodgiveren.Epost1}'"
            Dim sqlNy As New MySqlCommand(sporring, tilkobling)
            sqlNy.ExecuteNonQuery()
        Catch ex As MySqlException
            MsgBox("Feil ved tilkobling til databasen: " & ex.Message())
        Finally
            tilkobling.Close()
        End Try
    End Sub

    'Lagre svar i erklæring og vis neste spørsmål
    Private Sub btnNeste_Click(sender As Object, e As EventArgs) Handles btnEgenerklNesteSpm.Click
        Dim sisteindex, kjønn, spmID, i As Integer
        Dim spmText, pnr As String
        Dim dame As Boolean
        pnr = blodgiveren.Fodselsnummer1
        sisteindex = Erklæringspørsmål.Rows.Count - 1
        spmText = ""

        kjønn = pnr.Substring(8, 1)
        If (kjønn = 0) Or (kjønn = 2) Or (kjønn = 4) Or (kjønn = 6) Or (kjønn = 8) Then
            dame = True
        Else
            dame = False
        End If
        If (rBtnEgenerklJa.Checked = False) And (rBtnEgenerklNei.Checked = False) Then
            MsgBox("Du må svare før du går videre")
            Exit Sub
        End If
        If rBtnEgenerklJa.Checked Then
            erklæringSvar(SPMnr) = 1
        Else
            erklæringSvar(SPMnr) = 0
        End If

        'Neste spm
        If SPMnr >= sisteindex Then
            btnEgenerklNesteSpm.Enabled = False
            MsgBox("Alle spørsmål besvart - send inn!")
            Exit Sub
        End If
        SPMnr = SPMnr + 1
        spmID = Erklæringspørsmål.Rows(SPMnr).Item("Nr")
        For i = 1 To (sisteindex - SPMnr) + 1
            If spmID < 100 Then
                spmText = Erklæringspørsmål.Rows(SPMnr).Item("spoersmaal")
                Exit For
            ElseIf (spmID > 199) And (dame = False) Then
                spmText = Erklæringspørsmål.Rows(SPMnr).Item("spoersmaal")
                Exit For
            ElseIf (spmID > 99) And (dame = True) Then
                spmText = Erklæringspørsmål.Rows(SPMnr).Item("spoersmaal")
                Exit For
            Else
                SPMnr = SPMnr + 1
                spmID = Erklæringspørsmål.Rows(SPMnr).Item("Nr")
                spmText = Nothing
            End If
        Next
        SPMnrPresentert = SPMnrPresentert + 1
        lblEgenerklSpmTekst.Text = spmText
        lblEgenerklSpmNr.Text = $"Spørsmål {SPMnrPresentert + 1}"
        rBtnEgenerklJa.Checked = False
        rBtnEgenerklNei.Checked = False
        btnEgenerklForrigeSpm.Enabled = True
    End Sub

    'Sjekker om valgt dato er fridag
    Private Function fridag(ByVal dato As Date) As Boolean
        Dim fridagtabell As New Hashtable()
        Dim aaret As Integer = dato.Year
        Dim førstepåskedag As Date = GetEasterDate(aaret)
        fridagtabell.Add("1. nyttårsdag", DateSerial(aaret, 1, 1))
        fridagtabell.Add("Skjærtorsdag", førstepåskedag.AddDays(-3))
        fridagtabell.Add("Langfredag", førstepåskedag.AddDays(-2))
        fridagtabell.Add("1. påskedag", førstepåskedag)
        fridagtabell.Add("2. påskedag", førstepåskedag.AddDays(1))
        fridagtabell.Add("Kristi Himmelfartsdag", førstepåskedag.AddDays(39))
        fridagtabell.Add("1. pinsedag", førstepåskedag.AddDays(49))
        fridagtabell.Add("2. pinsedag", førstepåskedag.AddDays(50))
        fridagtabell.Add("1. mai", DateSerial(aaret, 5, 1))
        fridagtabell.Add("17. mai", DateSerial(aaret, 5, 17))
        fridagtabell.Add("Julekvelden", DateSerial(aaret, 12, 24))
        fridagtabell.Add("1. juledag", DateSerial(aaret, 12, 25))
        fridagtabell.Add("2. juledag", DateSerial(aaret, 12, 26))
        fridagtabell.Add("Nyttårskvelden", DateSerial(aaret, 12, 31))

        For Each nokkel In fridagtabell.Keys
            If DateSerial(aaret, dato.Month, dato.Day) = fridagtabell(nokkel) Then
                Return True
            End If
        Next
        Return False
    End Function

    'Lagre intervju og eventuelle endringer i blodgiver
    Private Sub LagreIntervjuInfo_Click(sender As Object, e As EventArgs) Handles btnHKtrlIntProfGjgått.Click
        Dim epost, adresse, preferanse, merknad, kommentar, spørring, spørring2 As String
        Dim tlf1, tlf2, postnr, status, i, svar As Integer
        Dim da As New MySqlDataAdapter

        'Sikker på at du ikke vil godkjenne/ikke godkjenne giver?
        If rBtnHKtrlIkkeGodkjent.Checked = True Then
            svar = MsgBox("Er du sikker på at du vil sette status til ''Ikke godkjent giver''?", MsgBoxStyle.YesNo)
            If svar = 7 Then
                Exit Sub
            End If
        ElseIf rBtnHKtrlGodkjent.Checked = True Then
            svar = MsgBox("Er du sikker på at du vil godkjenne giveren?", MsgBoxStyle.YesNo)
            If svar = 7 Then
                Exit Sub
            End If
        End If

        i = lBxSøkResultater.SelectedIndex
        epost = txtValgtBlodgiverEpost.Text
        tlf1 = txtValgtBlodgiverTelefon1.Text
        tlf2 = txtValgtBlodgiverTelefon2.Text
        adresse = txtValgtBlodgiverAdresse.Text
        postnr = txtValgtBlodgiverPostnr.Text
        status = txtValgtBlodgiverStatusKode.Text
        preferanse = rTxtValgBlodgiverTimepref.Text
        merknad = rTxtValgtBlodgiverInternMrknd.Text
        kommentar = rTxtHKtrlKommentar.Text
        spørring = $"UPDATE egenerklaering SET ansattepost= '{påloggetAepost}', datotidansatt= '{Now.ToString("yyyy.MM.dd HH:mm.ss")}', kommentar= '{kommentar}' WHERE id= '{egenerklæringID}'"
        spørring2 = $"UPDATE bruker SET epost= '{epost}', telefon1= '{tlf1}', telefon2= '{tlf2}', adresse= '{adresse}', postnr= '{postnr}', statuskode= '{status}' WHERE epost= '{presentertGiver}'"
        Try
            tilkobling.Open()
            If GroupBoxIntervju.Visible = True Then
                Dim sqlSpørring As New MySqlCommand($"{spørring}", tilkobling)
                sqlSpørring.ExecuteNonQuery()
            End If

            Dim sqlSpørring2 As New MySqlCommand($"{spørring2}", tilkobling)
            sqlSpørring2.ExecuteNonQuery()
            tilkobling.Close()
        Catch
            MsgBox("Feil")
        End Try
        bgSøk(bgSøkParameter)
        visBG()
    End Sub

    Private Sub TabPage5_Enter(sender As Object, e As EventArgs) Handles TabPage5.Enter
        btnEgenerklForrigeSpm.Enabled = False
    End Sub
End Class
