Public Class pålogging
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim brukere As New DataTable
        Dim rad As DataRow
        Dim epost, passord, pålogget, påloggetEpost As String
        Dim feil As Boolean = False
        brukere = Blodbane.ansatt
        epost = TextBox1.Text
        passord = TextBox2.Text
        pålogget = ""
        påloggetEpost = ""

        For Each rad In brukere.Rows
            If (epost = rad("epost")) And (passord = rad("passord")) Then
                feil = False
                pålogget = rad("fornavn")
                påloggetEpost = rad("epost")
                Blodbane.PanelAnsatt.BringToFront()
                Blodbane.PanelAnsatt.Show()
                Blodbane.PanelGiver.Hide()
                Blodbane.PanelPåmelding.Hide()
                Blodbane.LoggAvToolStripMenuItem.Visible = True
                Blodbane.LoggPåansattToolStripMenuItem.Visible = False
                Me.Close()
            Else
                feil = True
            End If
        Next
        If feil = True Then
            MsgBox("Denne kombinasjonen av epost og passord eksisterer ikke", vbInformation)
            Exit Sub
        End If
        Blodbane.påloggetAnsatt = pålogget
        Blodbane.påloggetAepost = påloggetEpost
        Blodbane.Label23.Text = $"Velkommen {pålogget}"
    End Sub
End Class