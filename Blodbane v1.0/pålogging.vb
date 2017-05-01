Public Class pålogging
    Dim brukere As New DataTable
    Dim rad As DataRow
    Dim epost, passord, pålogget, påloggetEpost As String
    Dim riktigPålogging As Boolean = False

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        brukere = Blodbane.ansatt
        epost = TextBox1.Text
        passord = TextBox2.Text
        pålogget = ""
        påloggetEpost = ""

        For Each rad In brukere.Rows
            If (epost = rad("epost")) And (passord = rad("passord")) Then
                riktigPålogging = True
                pålogget = rad("fornavn")
                påloggetEpost = rad("epost")
                Blodbane.ansattObjekt.Ansattnummer1 = rad("ansattnummer")
                Blodbane.ansattObjekt.Epost1 = rad("epost")
                Blodbane.ansattObjekt.Passord1 = rad("passord")
                Blodbane.ansattObjekt.Fornavn1 = rad("fornavn")
                Blodbane.ansattObjekt.Etternavn1 = rad("etternavn")
                Blodbane.ansattObjekt.Adresse1 = rad("adresse")
                Blodbane.ansattObjekt.Postnr1 = rad("postnr")
                Blodbane.ansattObjekt.Telefon11 = rad("telefon1")
                Blodbane.ansattObjekt.Telefon21 = rad("telefon2")
                Blodbane.ansattObjekt.Status1 = rad("beskrivelse")

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
        Blodbane.påloggetAnsatt = pålogget
        Blodbane.påloggetAepost = påloggetEpost
        Blodbane.lblVelkommen.Text = $"Velkommen {pålogget}"
    End Sub

    Private Sub pålogging_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Select()
    End Sub
End Class