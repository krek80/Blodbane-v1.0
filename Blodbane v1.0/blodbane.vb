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
    Public påloggetAnsatt, påloggetAepost, påloggetBgiver As String
    Dim egenerklæringID, SPMnr, erklæringSvar(60) As Integer
    Dim presentertGiver, bgSøkParameter As String
    Dim tilkobling As New MySqlConnection("Server=mysql.stud.iie.ntnu.no;" & "Database=g_ioops_02;" & "Uid=g_ioops_02;" & "Pwd=LntL4Owl;")

    'Kjøres ved oppstart
    Private Sub Blodbane_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Hide()
        velkommen.Show()

        'Henter statuskoder og legger i combobox(er)
        Dim statuser As New DataTable
        Dim steder As New DataTable
        Dim da As New MySqlDataAdapter
        Dim rad As DataRow
        Dim statustekst, statuskode, psted, pnr, spmNR, spmTekst As String
        giversøk.Clear()
        tilkobling.Open()
        Dim sqlSpørring As New MySqlCommand("SELECT * FROM personstatus", tilkobling)
        da.SelectCommand = sqlSpørring
        da.Fill(statuser)
        ComboBox2.Items.Clear()
        For Each rad In statuser.Rows
            statustekst = rad("beskrivelse")
            statuskode = rad("kode")
            personstatusK.Add(statustekst, statuskode)
            personstatusB.Add(statuskode, statustekst)
            ComboBox2.Items.Add(statustekst)
            ComboBox4.Items.Add(statustekst)
        Next

        'Lager liste over rommene
        antallRom = 0
        Dim sqlSporringRom As String = "SELECT * FROM rom"
        Dim sqlRom As New MySqlCommand(sqlSporringRom, tilkobling)
        Dim daRom As New MySqlDataAdapter
        daRom.SelectCommand = sqlRom
        daRom.Fill(interntabellRom)
        antallRom = interntabellRom.Rows.Count()


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
            ComboBox3.Items.Add(rad("epost"))
        Next

        'Henter ned spørsmål til egenerklæring
        Dim sqlSpørring4 As New MySqlCommand("SELECT * FROM egenerklaeringsporsmaal", tilkobling)
        da.SelectCommand = sqlSpørring4
        da.Fill(Erklæringspørsmål)

        tilkobling.Close()
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
        Dim sql As New MySqlCommand("SELECT * FROM bruker WHERE epost = @epostInn AND passord = @passordInn", tilkobling)
        sql.Parameters.AddWithValue("@epostInn", txtAInn_epost.Text)
        sql.Parameters.AddWithValue("@passordInn", txtAInn_passord.Text)
        Dim da As New MySqlDataAdapter
        Dim interntabell As New DataTable
        'Objektet "da" utfører spørringen og legger resultatet i "interntabell"
        da.SelectCommand = sql
        da.Fill(interntabell)
        Dim rad As DataRow

        Dim antallRader As Integer = interntabell.Rows.Count()

        If antallRader = 1 Then

            For Each rad In interntabell.Rows
                blodgiverData.Add("fornavn", rad("fornavn"))
                blodgiverData.Add("etternavn", rad("etternavn"))
                blodgiverData.Add("passord", rad("passord"))
                blodgiverData.Add("adresse", rad("adresse"))
                blodgiverData.Add("postnr", rad("postnr"))
                blodgiverData.Add("telefon1", rad("telefon1"))
                blodgiverData.Add("telefon2", rad("telefon2"))
                blodgiverData.Add("epost", rad("epost"))
                blodgiverData.Add("statuskode", rad("statuskode"))
            Next rad

            Dim sql2 As New MySqlCommand("SELECT * FROM blodgiver WHERE epost = @epostInn", tilkobling)
            sql2.Parameters.AddWithValue("@epostInn", txtAInn_epost.Text)
            Dim da2 As New MySqlDataAdapter
            Dim interntabell2 As New DataTable
            'Objektet "da" utfører spørringen og legger resultatet i "interntabell"
            da2.SelectCommand = sql2
            da2.Fill(interntabell2)

            Dim rad2 As DataRow
            For Each rad2 In interntabell2.Rows
                If IsDBNull(rad2("blodtype")) Then
                    blodgiverData.Add("blodtype", "")
                Else
                    blodgiverData.Add("blodtype", rad2("blodtype"))
                End If
                If IsDBNull(rad2("siste_blodtapping")) Then
                    blodgiverData.Add("siste_blodtapping", "")
                Else
                    blodgiverData.Add("siste_blodtapping", rad2("siste_blodtapping"))
                End If
                blodgiverData.Add("kontaktform", rad2("kontaktform"))

            Next rad2

            Dim idag, sistetime As DateTime
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
                Else
                    TxtNesteInnkalling.Text = "Ikke fastsatt"
                    BtnEndreInnkalling.Enabled = False
                End If
            Else
                TxtNesteInnkalling.Text = "Ikke fastsatt"
            End If

            PanelPåmelding.Hide()
            PanelAnsatt.Hide()
            PanelGiver.Show()
            PanelGiver.BringToFront()

            txtPersDataNavn.Text = $"{blodgiverData("fornavn")} {blodgiverData("etternavn")}"
            txtPersDataGStatus.Text = blodgiverData("statuskode")
            txtPersDataBlodtype.Text = blodgiverData("blodtype")
            txtPersDataSisteUnders.Text = blodgiverData("siste_blodtapping")
            txtPersDataGateAdr.Text = blodgiverData("adresse")
            txtPersDataPostnr.Text = blodgiverData("postnr")
            txtPersDataTlf.Text = blodgiverData("telefon1")
            txtPersDataTlf2.Text = blodgiverData("telefon2")
            txtPersDataEpost.Text = blodgiverData("epost")
            If Not IsDBNull(blodgiverData("kontaktform")) Then
                CBxKontaktform.Text = blodgiverData("kontaktform")
            End If
        Else
            MsgBox("Epostadressen eller passordet er feil.", MsgBoxStyle.Critical)

            tilkobling.Close()
            påloggetBgiver = blodgiverData("epost")
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
                                  txtBgInn_passord2.Text) Then

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

                spoerring = $"INSERT INTO blodgiver (epost, fodselsnummer, kontaktform) VALUES ('{txtBgInn_epost.Text}', '{txtBgInn_personnr.Text}', 'Epost')"
                Dim sql2 As New MySqlCommand(spoerring, tilkobling)
                Dim da2 As New MySqlDataAdapter
                Dim interntabell2 As New DataTable
                da2.SelectCommand = sql2
                da2.Fill(interntabell2)

                'PanelPåmelding.Hide()
                'PanelAnsatt.Hide()
                'PanelGiver.Show()
                'PanelGiver.BringToFront()
                MsgBox("Skjema Ok! Nå kan du logge deg på.")

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
                                        ByVal passord2Inn As String) As Boolean

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

        If fornavnInn = "" Or etternavnInn = "" Or telefon1Inn = "" Then
            MsgBox("Alle felt må være utfylt unntatt gateadresse- og telefon 2-feltet må være utfylt.", MsgBoxStyle.Critical)
            Return False
        End If

        If Not IsNumeric(personnrInn) Or personnrInn.Length <> 11 Then
            MsgBox("Fødselsnummeret inneholder ikke bare tall, eller består ikke av 11 siffer.", MsgBoxStyle.Critical)
            Return False
        End If

        Dim aar As String
        Dim aarstall As Integer = CInt(personnrInn.Substring(4, 2))
        If aarstall < 17 Then
            aar = $"20{aarstall}"
        Else
            aar = $"19{aarstall}"
        End If
        Dim fnrDato As String = $"#{personnrInn.Substring(0, 2)}/{personnrInn.Substring(2, 2)}/{aar}#"
        If Not IsDate(fnrDato) Then
            MsgBox($"De seks første tallene i fødselsnummeret, {fnrDato}, ble ikke gjenkjent som en dato.", MsgBoxStyle.Critical)
            Return False
        End If

        If interntabell2.Rows.Count = 1 Then
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

        If interntabell1.Rows.Count = 1 Then
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

    'Logg av blodgiver
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles BttnLoggavGiver.Click
        PanelGiver.Hide()
        PanelAnsatt.Hide()
        PanelPåmelding.Show()
        PanelPåmelding.BringToFront()
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
        Label23.Text = ""
        PanelGiver.Hide()
        PanelAnsatt.Hide()
        PanelPåmelding.Show()
        PanelPåmelding.BringToFront()
        LoggPåansattToolStripMenuItem.Visible = True
        LoggAvToolStripMenuItem.Visible = False
    End Sub

    'Blodgiversøk knapp
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles BttnSøkGiver.Click
        Me.Cursor = Cursors.WaitCursor
        Dim personnummer As String = TextBox19.Text
        Dim status As String = TextBox20.Text
        Dim statuskode As Integer
        Dim blodtype As String = ComboBox5.Text
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

        ListBox2.Items.Clear()
        For Each rad In giversøk.Rows
            resPnr = rad("fodselsnummer")
            resFnavn = rad("fornavn")
            resEnavn = rad("etternavn")
            resStatus = rad("beskrivelse")
            resKode = rad("statuskode")
            ListBox2.Items.Add($"{resPnr} {vbTab}{resFnavn} {resEnavn} {vbTab}{resKode} - {resStatus}")
        Next
        If ListBox2.Items.Count > 0 Then
            ListBox2.SetSelected(0, True)
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

            Dim sqlSpørring2 As New MySqlCommand("SELECT * FROM egenerklaering", tilkobling)
            da.SelectCommand = sqlSpørring2
            da.Fill(egenerklaering)
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
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            statuskode(ComboBox2.SelectedItem, TextBox20, ComboBox2)
        Catch
            TextBox20.Text = ""
            ComboBox2.Text = ""
            Exit Sub
        End Try
    End Sub

    'Endring av statuskode - henter statusbeskrivelse
    Private Sub TextBox20_TextChanged(sender As Object, e As EventArgs) Handles TextBox20.TextChanged
        Try
            statusbeskrivelse(TextBox20.Text, ComboBox2, TextBox20)
        Catch
            TextBox20.Text = ""
            ComboBox2.Text = ""
            Exit Sub
        End Try
    End Sub

    'Presenter valgt person i blodgiversøk
    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        visBG()
    End Sub

    'Vis blodgiverdata til ansatt
    Private Sub visBG()
        Dim index, i As Integer
        Dim rad As DataRow
        Dim fNavn, eNavn, fnr, epost, adresse, postnummmer, tlf1, tlf2, intMerknad, preferanse, jasvar, erklaringLege As String
        Dim status As Integer
        Dim sistTapping, sistErklæring, gjennomgåttErklæring As Date
        Dim dager As Long
        ListBox3.Items.Clear()
        index = ListBox2.SelectedIndex
        For Each rad In giversøk.Rows
            fNavn = rad("fornavn")
            eNavn = rad("etternavn")
            fnr = rad("fodselsnummer")
            epost = rad("epost")
            adresse = rad("adresse")
            tlf1 = rad("telefon1")
            tlf2 = rad("telefon2")
            postnummmer = rad("postnr")
            status = rad("statuskode")
            sistTapping = rad("siste_blodtapping")
            intMerknad = rad("merknad")
            preferanse = rad("timepreferanse")
            If i = index Then
                Exit For
            End If
            i = i + 1
        Next
#Disable Warning BC42104 ' Variable is used before it has been assigned a value
        For Each rad In egenerklaering.Rows
            If rad("bgepost") = epost Then
                If rad("datotidbg") > sistErklæring Then
                    sistErklæring = rad("datotidbg")
                    jasvar = rad("skjema")
                    egenerklæringID = rad("id")
                    presentertGiver = rad("bgepost")
                    If Not IsDBNull(rad("ansattepost")) Then
                        erklaringLege = rad("ansattepost")
                    Else
                        erklaringLege = ""
                    End If
                    If Not IsDBNull(rad("datotidansatt")) Then
                        gjennomgåttErklæring = rad("datotidansatt")
                    Else
                        gjennomgåttErklæring = Nothing
                    End If
                End If
            End If
        Next
        dager = DateDiff(DateInterval.DayOfYear, sistTapping, Today)
        utledJAsvar(jasvar)

        TextBox24.Text = $"{fNavn} {eNavn}"
        TextBox25.Text = fnr
        TextBox27.Text = epost
        TextBox26.Text = tlf1
        TextBox29.Text = tlf2
        TextBox30.Text = adresse
        TextBox31.Text = postnummmer
        TextBox21.Text = status
        TextBox35.Text = $"{dager} dager"
        TextBox28.Text = sistTapping
        RichTextBox4.Text = preferanse
        RichTextBox2.Text = intMerknad
        TextBox22.Text = sistErklæring
        TextBox33.Text = erklaringLege
        If gjennomgåttErklæring = Nothing Then
            TextBox34.Text = ""
            GroupBoxIntervju.Visible = True
        Else
            TextBox34.Text = gjennomgåttErklæring
            GroupBoxIntervju.Visible = False
        End If
#Enable Warning BC42104
    End Sub

    'Utleder Jasvar og presenterer i Listebox i giversøk
    Private Sub utledJAsvar(ByVal spmNr As String)
        Dim svar() As String = spmNr.Split(",")
        For i = 0 To svar.Length - 1
            ListBox3.Items.Add(svar(i))
        Next
    End Sub

    'Endring av statuskode - henter statusbeskrivelse
    Private Sub TextBox21_TextChanged(sender As Object, e As EventArgs) Handles TextBox21.TextChanged
        Try
            statusbeskrivelse(TextBox21.Text, ComboBox4, TextBox21)
        Catch
            TextBox21.Text = ""
            ComboBox4.Text = ""
            Exit Sub
        End Try
    End Sub

    'Endring av statusbeskrvelse - henter statuskode
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Try
            statuskode(ComboBox4.SelectedItem, TextBox21, ComboBox4)
        Catch
            TextBox21.Text = ""
            ComboBox4.Text = ""
            Exit Sub
        End Try
    End Sub

    'Sett rett poststed ved siden av postnummer i personsøk
    Private Sub TextBox31_TextChanged(sender As Object, e As EventArgs) Handles TextBox31.TextChanged
        TextBox32.Text = postnummer(TextBox31.Text)
    End Sub

    'Sett rett poststed ved siden av postnummer i egenregistrering
    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles txtBgInn_postnr.TextChanged
        txtBgInn_poststed.Text = postnummer(txtBgInn_postnr.Text)
    End Sub

    'Tøm giversøk
    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox19.Text = ""
        ComboBox5.Text = ""
        TextBox20.Text = ""
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
                ListBox4.Items.Add($"{dato} - Rom: {romnr} - Etg: {etg} - Giver: {epost}")
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
        If GpBxEndreInnkalling.Visible Then
            GpBxEndreInnkalling.Visible = False
        Else
            GpBxEndreInnkalling.Visible = True
        End If
        If TxtNesteInnkalling.Text <> "" Then
            DateTimePickerNyTime.Value = CDate(TxtNesteInnkalling.Text)
            hentLedigeTimer(DateTimePickerNyTime.Value)
        End If
        BtnBekreftEndretTime.Enabled = False
    End Sub

    'Henter ledige timer for valgt dato
    Private Sub hentLedigeTimer(ByVal aktuelldato As DateTime)

        Dim aktuelldatopluss1 = aktuelldato.AddDays(1)
        tilkobling.Open()

        Dim sqlSporring1 As String = $"SELECT datotid FROM timeavtale WHERE datotid > '{aktuelldato.ToString("yyyy-MM-dd")}' AND datotid < '{aktuelldatopluss1.ToString("yyyy-MM-dd")}'"
        Dim sql1 As New MySqlCommand(sqlSporring1, tilkobling)
        Dim da1 As New MySqlDataAdapter
        Dim interntabell1 As New DataTable
        Dim rad1 As DataRow
        Dim antallTimerPåDetteKlokkeslettet As Integer = 0
        Dim tabort As Integer = 0
        Dim fulltimetabell As New ArrayList()
        Dim opptatt As Boolean = False
        Dim raddato1 As DateTime
        Dim raddato2 As String
        Dim i As Integer
        'Objektet "da" utfører spørringen og legger resultatet i "interntabell1"
        da1.SelectCommand = sql1
        da1.Fill(interntabell1)
        tilkobling.Close()
        LBxLedigeTimer.Items.Clear()
        For i = 0 To 7
            fulltimetabell.Add($"{i + 8}:00")
        Next
        For Each rad1 In interntabell1.Rows
            raddato1 = rad1("datotid")
            raddato2 = $"{raddato1.Hour}:00"
            For i = 0 To fulltimetabell.Count - 1
                If fulltimetabell(i) = raddato2 Then
                    antallTimerPåDetteKlokkeslettet += 1
                    If antallTimerPåDetteKlokkeslettet = antallRom Then
                        tabort = i
                        opptatt = True
                    End If
                End If
            Next
            If opptatt Then
                fulltimetabell.RemoveAt(tabort)
                opptatt = False
            End If
        Next
        For i = 0 To fulltimetabell.Count - 1
            LBxLedigeTimer.Items.Add(fulltimetabell(i))
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

        Dim nyDato, time_DateTime As DateTime
        Try
            Dim provider As CultureInfo = CultureInfo.InvariantCulture
            time_DateTime = Date.ParseExact(LBxLedigeTimer.SelectedItem, "H:mm", provider)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        nyDato = New Date(DateTimePickerNyTime.Value.Year, DateTimePickerNyTime.Value.Month, DateTimePickerNyTime.Value.Day,
                               time_DateTime.Hour, 0, 0)

        GpBxEndreInnkalling.Visible = False
        TxtNesteInnkalling.Text = nyDato

    End Sub

    'Registrer gitt blod
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim ansatt, SettInnSpørring, SøkSpørring, timeID As String
        Dim bPlater, bLegemer, bPlasma, i As Integer
        Dim timedato As Date
        Dim feil As Boolean
        Dim antallProdukt(2) As Integer
        Dim tabell As New DataTable
        Dim da As New MySqlDataAdapter
        ansatt = ComboBox3.Text
        bPlater = NumericUpDown1.Value
        bPlasma = NumericUpDown2.Value
        bLegemer = NumericUpDown3.Value
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
            NumericUpDown1.Value = 0
            NumericUpDown2.Value = 0
            NumericUpDown3.Value = 0
            MsgBox("Blodlager oppdatert")
        Catch ex As Exception
            tilkobling.Close()
            MsgBox("Feil")
        End Try
    End Sub

    Private Sub TabPage5_Enter(sender As Object, e As EventArgs) Handles TabPage5.Enter
        SPMnr = 1
        lblSpml.Text = Erklæringspørsmål.Rows(0).Item("spoersmaal")
        Label26.Text = "Spørsmål 1"
    End Sub

    'Forrige spørsmål i erklæring
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim spmText As String

        If SPMnr > 1 Then
            SPMnr = SPMnr - 1
            spmText = 1
            lblSpml.Text = spmText
            Label26.Text = $"Spørsmål {SPMnr}"
        Else
            SPMnr = SPMnr
            MsgBox("Dette var siste spørsmål")
        End If
        RadioButton3.Checked = False
        RadioButton4.Checked = False
    End Sub

    'Send inn egenerklæring
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Jasvar As String
        Dim i As Integer
        For i = 0 To 60
            If erklæringSvar(i) = 1 Then
                Jasvar = Jasvar & i & ","
            End If
        Next

        'SQL her:

    End Sub

    'Neste spørsmål i erklæring
    Private Sub btnNeste_Click(sender As Object, e As EventArgs) Handles btnNeste.Click
        Dim sisteindex, kjønn As Integer
        Dim spmText, pnr As String
        Dim dame As Boolean
        pnr = "04079147929" 'Testverdi

        kjønn = pnr.Substring(8, 1)
        If (kjønn = 0) Or (kjønn = 2) Or (kjønn = 4) Or (kjønn = 6) Or (kjønn = 8) Then
            dame = True
        Else
            dame = False
        End If
        If (RadioButton3.Checked = False) And (RadioButton4.Checked = False) Then
            MsgBox("Du må svare før du går videre")
            Exit Sub
        End If
        sisteindex = Erklæringspørsmål.Rows.Count
        SPMnr = SPMnr + 1

        If Erklæringspørsmål.Rows(SPMnr - 1).Item("Nr") < 100 Then
            spmText = Erklæringspørsmål.Rows(SPMnr - 1).Item("spoersmaal")
        ElseIf (Erklæringspørsmål.Rows(SPMnr - 1).Item("Nr") > 199) And (dame = False) Then
            spmText = Erklæringspørsmål.Rows(SPMnr - 1).Item("spoersmaal")
        ElseIf (Erklæringspørsmål.Rows(SPMnr - 1).Item("Nr") > 99) And (dame = True) Then
            spmText = Erklæringspørsmål.Rows(SPMnr - 1).Item("spoersmaal")
        End If
        If spmText = Nothing Then
            MsgBox("dette spørsmålet er ikke for deg, gå videre")
        End If

        lblSpml.Text = spmText
        Label26.Text = $"Spørsmål {SPMnr}"

        'Registrere svar
        If RadioButton3.Checked Then
            erklæringSvar(SPMnr - 1) = 1
        Else
            erklæringSvar(SPMnr - 1) = 0
        End If
        RadioButton3.Checked = False
        RadioButton4.Checked = False
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
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim epost, adresse, preferanse, merknad, kommentar, spørring, spørring2 As String
        Dim tlf1, tlf2, postnr, status, i, svar As Integer
        Dim da As New MySqlDataAdapter

        'Sikker på at du ikke vil godkjenne/ikke godkjenne giver?
        If RadioButton2.Checked = True Then
            svar = MsgBox("Er du sikker på at du vil sette status til ''Ikke godkjent giver''?", MsgBoxStyle.YesNo)
            If svar = 7 Then
                Exit Sub
            End If
        ElseIf RadioButton1.Checked = True Then
            svar = MsgBox("Er du sikker på at du vil godkjenne giveren?", MsgBoxStyle.YesNo)
            If svar = 7 Then
                Exit Sub
            End If
        End If

        i = ListBox2.SelectedIndex
        epost = TextBox27.Text
        tlf1 = TextBox26.Text
        tlf2 = TextBox29.Text
        adresse = TextBox30.Text
        postnr = TextBox31.Text
        status = TextBox21.Text
        preferanse = RichTextBox4.Text
        merknad = RichTextBox2.Text
        kommentar = RichTextBox3.Text
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
End Class
