Public Class Blodbane
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
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles BttnSøkGiver.Click
        Dim personnummer As String = TextBox19.Text
        Dim telefon As Integer = TextBox20.Text
        Dim epost As String = TextBox21.Text
        Dim status As String = ComboBox2.Text
        Dim blodtype As String = ComboBox5.Text
        søk(personnummer, telefon, epost, status, blodtype)
    End Sub

    'Kjører SQL med søk mot database
    Private Sub søk(ByVal pnr As String, ByVal telefon As Integer, ByVal epost As String, ByVal status As String, ByVal blodtype As String)
        Dim parameterListe(4) As String
        Dim sql As String
        parameterListe(0) = pnr : parameterListe(1) = telefon : parameterListe(2) = epost : parameterListe(3) = status : parameterListe(4) = blodtype

        sql = "SELECT *  FROM giver WHERE"
        For i = 0 To 4
            If parameterListe(i) = 0 Then
                sql = sql
            Else
                sql = sql & parameterListe(i)
            End If
        Next
        MsgBox("Søkestreng:"sql)

    End Sub

End Class
