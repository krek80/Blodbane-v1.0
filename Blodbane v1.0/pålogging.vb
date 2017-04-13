Public Class pålogging
    Public pålogget As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim brukere As New DataTable
        Dim rad As DataRow
        Dim epost, passord As String
        brukere = Blodbane.ansatt
        epost = TextBox1.Text
        passord = TextBox2.Text

        For Each rad In brukere.Rows
            If (epost = rad("epost")) And (passord = rad("passord")) Then
                pålogget = rad("fornavn")
                Blodbane.PanelAnsatt.BringToFront()
                Blodbane.PanelAnsatt.Show()
                Blodbane.PanelGiver.Hide()
                Blodbane.PanelPåmelding.Hide()
                Blodbane.LoggAvToolStripMenuItem.Visible = True
                Blodbane.LoggPåansattToolStripMenuItem.Visible = False
                Me.Close()
            Else
                MsgBox("Denne kombinasjonen av epost og passord eksisterer ikke", vbInformation)
                Exit Sub
            End If
        Next
        Blodbane.påloggetAnsatt = pålogget
        Blodbane.Label23.Text = $"Velkommen {pålogget}"
    End Sub
End Class