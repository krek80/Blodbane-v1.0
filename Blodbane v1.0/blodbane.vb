Imports System.ComponentModel
Imports MySql.Data.MySqlClient
Public Class Blodbane
    Dim giversøk As New DataTable
    Dim egenerklaering As New DataTable
    Dim innkalling As New DataTable
    Dim blodlager As New DataTable
    Public ansatt As New DataTable
    Dim personstatusK As New Hashtable
    Dim personstatusB As New Hashtable
    Dim postnummer As New Hashtable
    Public påloggetAnsatt As String
    Dim tilkobling As New MySqlConnection("Server=mysql.stud.iie.ntnu.no;" & "Database=g_ioops_02;" & "Uid=g_ioops_02;" & "Pwd=LntL4Owl;")
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

        'Henter postnummer og sted og legegr i hastable
        Dim sqlSpørring2 As New MySqlCommand("SELECT * FROM postnummer", tilkobling)
        da.SelectCommand = sqlSpørring2
        da.Fill(steder)
        For Each rad In steder.Rows
            psted = rad("poststed")
            pnr = rad("postnummer")
            postnummer.Add(pnr, psted)
        Next

        'Henter anasatte og legger i datatable
        Dim sqlSpørring3 As New MySqlCommand("SELECT a.epost, b.passord, b.fornavn FROM ansatt a Inner JOIN bruker b ON a.epost = b.epost", tilkobling)
        da.SelectCommand = sqlSpørring3
        da.Fill(ansatt)
        tilkobling.Close()
    End Sub

    Private Sub AvsluttToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AvsluttToolStripMenuItem.Click
        Me.Close()
    End Sub

    'Logg på ansatt
    Private Sub LoggPåansattToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoggPåansattToolStripMenuItem.Click
        pålogging.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles BttnLoggpåGiver.Click
        PanelPåmelding.Hide()
        PanelAnsatt.Hide()
        PanelGiver.Show()
        PanelGiver.BringToFront()
    End Sub

    Private Sub BttnSendSkjema_Click(sender As Object, e As EventArgs) Handles BttnSendSkjema.Click
        PanelPåmelding.Hide()
        PanelAnsatt.Hide()
        PanelGiver.Show()
        PanelGiver.BringToFront()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles BttnLoggavGiver.Click
        PanelGiver.Hide()
        PanelAnsatt.Hide()
        PanelPåmelding.Show()
        PanelPåmelding.BringToFront()
    End Sub

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

    'Knapp for å søke etter blodgivere basert på parametre - legger resultater i listeboks
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

    'Kjører SQL med søk mot database - legger resultat i DataTable
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

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            statuskode(ComboBox2.SelectedItem, TextBox20, ComboBox2)
        Catch
            TextBox20.Text = ""
            ComboBox2.Text = ""
            Exit Sub
        End Try
    End Sub

    Private Sub TextBox20_TextChanged(sender As Object, e As EventArgs) Handles TextBox20.TextChanged
        Try
            statusbeskrivelse(TextBox20.Text, ComboBox2, TextBox20)
        Catch
            TextBox20.Text = ""
            ComboBox2.Text = ""
            Exit Sub
        End Try
    End Sub

    'Presenter valgt person i giversøk
    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        Dim index, i As Integer
        Dim rad As DataRow
        Dim fNavn, eNavn, fnr, epost, adresse, postnummmer, tlf1, tlf2, intMerknad, preferanse, jasvar As String
        Dim status As Integer
        Dim sistTapping, sistErklæring As Date
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
#Enable Warning BC42104
    End Sub

    'Utleder Jasvar og presenterer i Listebox i giversøk
    Private Sub utledJAsvar(ByVal spmNr As String)
        Dim svar() As String = spmNr.Split(",")
        For i = 0 To svar.Length - 1
            ListBox3.Items.Add(svar(i))
        Next
    End Sub

    Private Sub TextBox21_TextChanged(sender As Object, e As EventArgs) Handles TextBox21.TextChanged
        Try
            statusbeskrivelse(TextBox21.Text, ComboBox4, TextBox21)
        Catch
            TextBox21.Text = ""
            ComboBox4.Text = ""
            Exit Sub
        End Try
    End Sub

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
    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        TextBox2.Text = postnummer(TextBox8.Text)
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
        Dim sqlSpørring As New MySqlCommand("SELECT * FROM blodprodukt b INNER JOIN blodstatus s ON b.statusid = s.id", tilkobling)
        da.SelectCommand = sqlSpørring
        da.Fill(blodlager)
        Me.Cursor = Cursors.Default

        For Each rad In blodlager.Rows
        Next

        tilkobling.Close()
    End Sub
End Class
