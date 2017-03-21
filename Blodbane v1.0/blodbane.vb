Public Class Blodbane
    Private Sub Blodbane_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Hide()
        velkommen.Show()

    End Sub

    Private Sub LoggPåToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoggPåToolStripMenuItem.Click
        pålogging.Show()
    End Sub

    Private Sub AvsluttToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AvsluttToolStripMenuItem.Click
        Me.Close()
    End Sub
End Class
