'Imports System.ComponentModel
Imports MySql.Data.MySqlClient
Imports System.Globalization
Public Class Blodbane
    Dim giversøk As New DataTable
    Dim egenerklaering As New DataTable
    Dim innkalling As New DataTable
    Dim blodlager As New DataTable
    Public ansatt As New DataTable
    Dim personstatusK As New Hashtable
    Dim personstatusB As New Hashtable
    Dim postnummer As New Hashtable
    Public påloggetAnsatt, påloggetAepost As String
    Dim egenerklærigID As Integer
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
        Dim statustekst, statuskode, psted, pnr As String
        giversøk.Clear()
        tilkobling.Open()
        Dim sqlSpørring As New MySqlCommand("SELECT * FROM personstatus", tilkobling)
        da.SelectCommand = sqlSpørring
        da.Fill(statuser)
        tilkobling.Close()
        ComboBox2.Items.Clear()
        For Each rad In statuser.Rows
            statustekst = rad("beskrivelse")
            statuskode = rad("kode")
            personstatusK.Add(statustekst, statuskode)
            personstatusB.Add(statuskode, statustekst)
            ComboBox2.Items.Add(statustekst)
            ComboBox4.Items.Add(statustekst)
        Next

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

    'Logg på blodgiver og setter opp personinfo i fanen Personinformasjon
    Private Sub ButtonLoggpåGiver_Click(sender As Object, e As EventArgs) Handles BttnLoggpåGiver.Click
        Dim sql As New MySqlCommand("SELECT * FROM bruker WHERE epost = @epostInn AND passord = @passordInn", tilkobling)
        sql.Parameters.AddWithValue("@epostInn", txtAInn_epost.Text)
        sql.Parameters.AddWithValue("@passordInn", txtAInn_passord.Text)
        Dim da As New MySqlDataAdapter
        Dim interntabell As New DataTable
        'Objektet "da" utfører spørringen og legger resultatet i "interntabell"
        da.SelectCommand = sql
        da.Fill(interntabell)

        Dim antallRader As Integer = interntabell.Rows.Count()
        Dim fornavnet As String = ""
        Dim etternavnet As String = ""
        Dim passordet As String = ""
        Dim adressen As String = ""
        Dim postnummeret As String = ""
        Dim telefonen1 As String = ""
        Dim telefonen2 As String = ""
        Dim statuskoden As String = ""
        Dim eposten As String = ""
        Dim blodtypen As String = ""
        Dim siste_timen As DateTime

        If antallRader = 1 Then

            Dim rad As DataRow
            For Each rad In interntabell.Rows
                fornavnet = rad("fornavn")
                etternavnet = rad("etternavn")
                passordet = rad("passord")
                adressen = rad("adresse")
                postnummeret = rad("postnr")
                telefonen1 = rad("telefon1")
                telefonen2 = rad("telefon2")
                eposten = rad("epost")
                statuskoden = rad("statuskode")
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
                blodtypen = rad2("blodtype")
                If IsDate(rad2("siste_blodtapping")) Then
                    siste_timen = rad2("siste_blodtapping")
                End If
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
                End If
            End If

            PanelPåmelding.Hide()
            PanelAnsatt.Hide()
            PanelGiver.Show()
            PanelGiver.BringToFront()

            txtPersDataNavn.Text = $"{fornavnet} {etternavnet}"
            txtPersDataGStatus.Text = personstatusB(statuskoden)
            txtPersDataBlodtype.Text = blodtypen
            txtPersDataSisteUnders.Text = siste_timen
            txtPersDataGateAdr.Text = adressen
            txtPersDataPostnr.Text = postnummeret
            txtPersDataTlf.Text = telefonen1
            txtPersDataTlf2.Text = telefonen2
            txtPersDataEpost.Text = eposten
        Else
            MsgBox("Epostadressen eller passordet er feil.", MsgBoxStyle.Critical)
        End If
        tilkobling.Close()

    End Sub

    'Registrer ny blodgiver
    Private Sub BtnRegBlodgiver_Click(sender As Object, e As EventArgs) Handles BtnRegBlodgiver.Click
        Try
            tilkobling.Open()
            Dim spoerring As String = ""
            If bgRegSkjemadata_OK(txtBgInn_personnr.Text, txtBgInn_poststed.Text, txtBgInn_tlfnr.Text, txtBgInn_tlfnr2.Text, txtBgInn_epost.Text, txtBgInn_passord1.Text, txtBgInn_passord2.Text) Then

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

                spoerring = $"INSERT INTO blodgiver (epost, fodselsnummer) VALUES ('{txtBgInn_epost.Text}', '{txtBgInn_personnr.Text}')"
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
    Private Function bgRegSkjemadata_OK(ByVal personnrInn As String, ByVal poststedInn As String, ByVal telefon1Inn As String, ByVal telefon2Inn As String, ByVal epostInn As String, ByVal passord1Inn As String, ByVal passord2Inn As String) As Boolean

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

        If txtBgInn_fornavn.Text = "" Or txtBgInn_etternavn.Text = "" Or txtBgInn_postnr.Text = "" Or txtBgInn_tlfnr.Text = "" Then
            MsgBox("Alle felt må være utfylt unntatt gateadresse- og telefon 2-feltet må være utfylt.", MsgBoxStyle.Critical)
            Return False
        End If

        If Not IsNumeric(personnrInn) Or personnrInn.Length <> 11 Then
            MsgBox("Fødselsnummeret ble ikke godtatt.", MsgBoxStyle.Critical)
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

        If passord1Inn <> passord2Inn Then
            MsgBox("Passordene er ikke like. Prøv igjen!", MsgBoxStyle.Critical)
            Return False
        End If
        If passord1Inn.Length < 6 Or passord1Inn.IndexOf(" ") <> -1 Then
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

    'Logg av ansatt
    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles BttnLoggavAnsatt.Click
        PanelGiver.Hide()
        PanelAnsatt.Hide()
        PanelPåmelding.Show()
        PanelPåmelding.BringToFront()
        LoggPåansattToolStripMenuItem.Visible = True
        LoggAvToolStripMenuItem.Visible = False
    End Sub

    'Logg av ansatt
    Private Sub LoggAvToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoggAvToolStripMenuItem.Click
        Label23.Text = ""
        PanelGiver.Hide()
        PanelAnsatt.Hide()
        PanelPåmelding.Show()
        PanelPåmelding.BringToFront()
        LoggPåansattToolStripMenuItem.Visible = True
        LoggAvToolStripMenuItem.Visible = False
    End Sub

    'Blodgiversøk
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles BttnSøkGiver.Click
        Me.Cursor = Cursors.WaitCursor
        Dim personnummer As String = TextBox19.Text
        Dim status As String = TextBox20.Text
        Dim statuskode As Integer
        Dim blodtype As String = ComboBox5.Text
        Dim rad As DataRow
        Dim resPnr, resFnavn, resEnavn, resStatus, resKode As String
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
        søk(personnummer, statuskode, blodtype)
        ListBox2.Items.Clear()
        Me.Cursor = Cursors.Default
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

    'SQL - blodgiversøk basert på 3 parameter
    Private Sub søk(ByVal pnr As String, ByVal status As Integer, ByVal blodtype As String)
        Dim sqlStreng As String
        Dim da As New MySqlDataAdapter
        giversøk.Clear()
        egenerklaering.Clear()
        Try
            tilkobling.Open()
            sqlStreng = "SELECT * FROM bruker br INNER JOIN blodgiver bl ON br.epost = bl.epost INNER JOIN personstatus ps ON ps.kode = br.statuskode WHERE"
            If (pnr <> "") And (status = 0) And (blodtype = "") Then
                sqlStreng = sqlStreng & $" bl.fødselsnummer = '{pnr}'"
            ElseIf (status > 0) And (pnr = "") And (blodtype = "") Then
                sqlStreng = sqlStreng & $" br.statuskode = '{status}'"
            ElseIf (blodtype <> "") And (status = 0) And (pnr = "") Then
                sqlStreng = sqlStreng & $" bl.blodtype = '{blodtype}'"
            ElseIf (blodtype <> "") And (status > 0) And (pnr = "") Then
                sqlStreng = sqlStreng & $" bl.blodtype = '{blodtype}' and br.statuskode = '{status}'"
            ElseIf (pnr <> "") And (status > 0) And (blodtype <> "") Then
                sqlStreng = sqlStreng & $" bl.blodtype = '{blodtype}' and br.statuskde = '{status}' and bl.fødselsnummer = '{pnr}'"
            End If
            Dim sqlSpørring As New MySqlCommand($"{sqlStreng}", tilkobling)
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
                    egenerklærigID = rad("id")
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
        Me.Cursor = Cursors.WaitCursor
        Dim rad As DataRow
        Dim da As New MySqlDataAdapter
        Dim sqlSpørring As New MySqlCommand("SELECT * FROM blodprodukt b INNER JOIN timeavtale t ON b.timeid = t.timeid INNER JOIN blodgiver bl on t.bgepost = bl.epost", tilkobling)
        da.SelectCommand = sqlSpørring
        da.Fill(blodlager)
        Me.Cursor = Cursors.Default

        For Each rad In blodlager.Rows
            MsgBox(rad("produkttypeid"))
        Next

        tilkobling.Close()
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
        End If


        Dim aktuell_dato As DateTime = DateTimePickerNyTime.Value
        Dim aktuell_datopluss1 = aktuell_dato.AddDays(1)
        MsgBox($"Aktuell dato: {aktuell_dato}. Aktuell dato pluss 1: {aktuell_datopluss1}.")


        Dim sqlSporring1 As String = $"SELECT datotid FROM timeavtale WHERE datotid > '{aktuell_dato.ToString("yyyy-MM-dd")}' AND datotid < '{aktuell_datopluss1.ToString("yyyy-MM-dd")}'"
        Dim sql1 As New MySqlCommand(sqlSporring1, tilkobling)
        Dim da1 As New MySqlDataAdapter
        Dim interntabell1 As New DataTable
        Dim rad1 As DataRow
        'Objektet "da" utfører spørringen og legger resultatet i "interntabell1"
        da1.SelectCommand = sql1
        da1.Fill(interntabell1)
        For Each rad1 In interntabell1.Rows
            LBxLedigeTimer.Items.Add(rad1("datotid"))
        Next

    End Sub

    'Plukker ut ledige timer når dato blir valgt.
    Private Sub DateTimePickerNyTime_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePickerNyTime.ValueChanged
        LblLedigeTimer.Text = $"Ledige timer {DateTimePickerNyTime.Text}"
    End Sub

    'Setter rett poststed ved siden av postnummeret i fanen Personinfo for blodgiveren
    Private Sub txtPersDataPostnr_TextChanged(sender As Object, e As EventArgs) Handles txtPersDataPostnr.TextChanged
        txtPersDataPoststed.Text = postnummer(txtPersDataPostnr.Text)
    End Sub

    'Lagre intervju og eventuelle endringer i blodgiver
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim epost, adresse, preferanse, merknad, kommentar, spørring As String
        Dim tlf1, tlf2, postnr, status As Integer
        Dim da As New MySqlDataAdapter
        epost = TextBox27.Text
        tlf1 = TextBox26.Text
        tlf2 = TextBox29.Text
        adresse = TextBox30.Text
        postnr = TextBox31.Text
        status = TextBox21.Text
        preferanse = RichTextBox4.Text
        merknad = RichTextBox2.Text
        kommentar = RichTextBox3.Text

        spørring = $"UPDATE egenerklaering SET ansattepost= '{påloggetAepost}', datotidansatt= '{Now.ToString("yyyy.MM.dd HH:mm.ss")}', kommentar= '{kommentar}' WHERE id= '{egenerklærigID}'"
        Try
            tilkobling.Open()
            MsgBox(spørring)
            Dim sqlSpørring As New MySqlCommand($"{spørring}", tilkobling)
            sqlSpørring.ExecuteNonQuery()
            tilkobling.Close()
        Catch
            MsgBox("Feil")
        End Try

    End Sub

End Class
