Imports MySql.Data.MySqlClient
Public Class Blodbane
    Dim giversøk As New DataTable
    Dim personstatusK As New Hashtable
    Dim personstatusB As New Hashtable
    Dim postnummer As New Hashtable
    Dim tilkobling As New MySqlConnection("Server=mysql.stud.iie.ntnu.no;" & "Database=g_ioops_02;" & "Uid=g_ioops_02;" & "Pwd=LntL4Owl;")
    Private Sub Blodbane_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Hide()
        velkommen.Show()

        'Henter statuskoder og legger i combobox(er)
        Dim statuser As New DataTable
        Dim steder As New DataTable
        Dim da As New MySqlDataAdapter
        Dim rad As DataRow
        Dim statustekst, statuskode As String
        Dim psted, pnr As String
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
        tilkobling.Close()
        For Each rad In steder.Rows
            psted = rad("poststed")
            pnr = rad("postnummer")
            postnummer.Add(pnr, psted)
        Next
    End Sub

    Private Sub AvsluttToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AvsluttToolStripMenuItem.Click
        Me.Close()
    End Sub

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

    Private Sub LoggAvToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoggAvToolStripMenuItem.Click
        PanelGiver.Hide()
        PanelAnsatt.Hide()
        PanelPåmelding.Show()
        PanelPåmelding.BringToFront()
        LoggPåansattToolStripMenuItem.Visible = True
        LoggAvToolStripMenuItem.Visible = False
    End Sub

    'Knapp for å søke etter blodgivere basert på parametre - legger resultater i listeboks
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles BttnSøkGiver.Click
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

    'Presenter valgt person
    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        Dim index, i As Integer
        Dim rad As DataRow
        Dim fNavn, eNavn, fnr, epost, adresse, postnummmer, tlf1, tlf2 As String
        Dim status As Integer
        Dim sistTapping As Date
        Dim dager As Long
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
            If i = index Then
                Exit For
            End If
            i = i + 1
        Next
        dager = DateDiff(DateInterval.DayOfYear, sistTapping, Today)

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
End Class
