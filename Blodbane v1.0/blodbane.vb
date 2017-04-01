Imports MySql.Data.MySqlClient
Public Class Blodbane
    Dim giversøk As New DataTable
    Dim tilkobling As New MySqlConnection("Server=mysql.stud.iie.ntnu.no;" & "Database=g_ioops_02;" & "Uid=g_ioops_02;" & "Pwd=LntL4Owl;")
    Private Sub Blodbane_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Hide()
        velkommen.Show()
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

    'Knapp for å søke etter blodgivere basert på parametre
    'Legger resultater i listeboks
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles BttnSøkGiver.Click
        Dim personnummer As String = TextBox19.Text
        Dim status As String = TextBox20.Text
        Dim statuskode As Integer
        Dim blodtype As String = ComboBox5.Text
        Dim rad As DataRow
        Dim resPnr, resFnavn, resEnavn, resStatus As String
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
            resStatus = rad("status")
            ListBox2.Items.Add($"{resPnr},{vbTab}{resFnavn} {resEnavn},{vbTab}{resStatus}")
        Next
    End Sub

    'Kjører SQL med søk mot database
    Private Sub søk(ByVal pnr As String, ByVal status As Integer, ByVal blodtype As String)
        Dim sqlStreng As String
        Dim da As New MySqlDataAdapter
        giversøk.Clear()
        tilkobling.Open()
        sqlStreng = "SELECT *  FROM blodgiver WHERE"
        If (pnr <> "") And (status = 0) And (blodtype = "") Then
            sqlStreng = sqlStreng & $" fodselsnummer = '{pnr}'"
        ElseIf (status > 0) And (pnr = "") And (blodtype = "") Then
            sqlStreng = sqlStreng & $" status = '{status}'"
        ElseIf (blodtype <> "") And (status = 0) And (pnr = "") Then
            sqlStreng = sqlStreng & $" blodtype = '{blodtype}'"
        ElseIf (blodtype <> "") And (status > 0) And (pnr = "") Then
            sqlStreng = sqlStreng & $" blodtype = '{blodtype}' and status = '{status}'"
        ElseIf (pnr <> "") And (status > 0) And (blodtype <> "") Then
            sqlStreng = sqlStreng & $" blodtype = '{blodtype}' and status = '{status}' and fodselsnummer = '{pnr}'"
        End If
        Dim sqlSpørring As New MySqlCommand($"{sqlStreng}", tilkobling)
        da.SelectCommand = sqlSpørring
        da.Fill(giversøk)
        tilkobling.Close()
    End Sub
End Class
