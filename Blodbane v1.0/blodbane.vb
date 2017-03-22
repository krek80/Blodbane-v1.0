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
        'PanelPåmelding.Hide()
        'PanelAnsatt.Hide()
        'PanelGiver.Show()
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
End Class
