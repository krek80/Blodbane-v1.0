Public Class pålogging
    Dim brukere As New DataTable
    Dim rad As DataRow
    Dim epost, passord, pålogget, påloggetEpost As String
    Dim riktigPålogging As Boolean = False

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnAnsattPålogg.Click

        brukere = Blodbane.ansatt
        epost = txtAnsattBrNavn.Text
        passord = txtAnsattPassord.Text
        pålogget = ""
        påloggetEpost = ""

        For Each rad In brukere.Rows
            If (epost = rad("epost")) And (passord = rad("passord")) Then
                riktigPålogging = True
                pålogget = rad("fornavn")
                påloggetEpost = rad("epost")

                Blodbane.AnsattObjOppdat(rad("ansattnummer"), rad("epost"), rad("passord"), rad("fornavn"), rad("etternavn"), rad("adresse"), rad("postnr"), rad("telefon1"), rad("telefon2"), rad("statuskode"))
                Blodbane.PanelAnsatt.BringToFront()
                Blodbane.PanelAnsatt.Show()
                Blodbane.PanelGiver.Hide()
                Blodbane.PanelPåmelding.Hide()
                Blodbane.LoggAvToolStripMenuItem.Visible = True
                Blodbane.LoggPåansattToolStripMenuItem.Visible = False
                Me.Close()
            End If
        Next
        If riktigPålogging = False Then
            MsgBox("Denne kombinasjonen av epost og passord eksisterer ikke", vbInformation)
            Exit Sub
        End If
        Blodbane.lblVelkommen.Text = $"Velkommen {Blodbane.ansattObj.Fornavn1}"
    End Sub

    Private Sub pålogging_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtAnsattBrNavn.Select()
    End Sub
End Class