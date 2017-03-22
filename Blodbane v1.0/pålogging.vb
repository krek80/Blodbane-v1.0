Public Class pålogging
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Blodbane.PanelAnsatt.BringToFront()
        Blodbane.PanelAnsatt.Show()
        Blodbane.PanelGiver.Hide()
        Blodbane.PanelPåmelding.Hide()
        Blodbane.LoggAvToolStripMenuItem.Visible = True
        Blodbane.LoggPåansattToolStripMenuItem.Visible = False
        Me.Close()
    End Sub
End Class