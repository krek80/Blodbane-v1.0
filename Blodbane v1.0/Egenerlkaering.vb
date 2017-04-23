Imports System.ComponentModel
Imports MySql.Data.MySqlClient

Public Class Egenerlkaering
    'Henter spørsmål fra egenerklæringsskjema og legger i Hashtable
    Dim sqlsporringEgenerkl As New mySqlcommand("SELECT * FROM egenerklaeringssporsmaal", tilkobling)
    Dim spoersmaal As New Hashtable

    da.selectcommand= sqlsporringEgenerkl
 da.Fill(spoersmaal)
    For Each rad In sporsmaal.Rows
    spml = rad("spoersmaal")
    nr = rad("Nr")
    spoersmaal.add(nr, spml)
    Next

    Private i As Integer
    Private Jasvar As String
    Private nr As Integer
    Private sisteindeks As Integer



    Private Sub btnNeste_Click(sender As Object, e As EventArgs) Handles btnNeste.Click
        'Funksjon: lagrer svar og blar til neste spørsmål 
        For i = nr To sisteindeks - 1
            spoersmaal(i) = spoersmaal(i + 1)
            Label1.Text = spoersmaal(i)


            'oppdaterer spørsmåsteller
            Label2.Text = i & " av 60 spørsmål er besvart"
            'Registrere eventuelt jasvar
            If rdbtnJa.Checked Then
                Jasvar = Jasvar & ", " & i
            Else
                Jasvar = Jasvar
            End If


        Next i
        'Lagre jasvar i tabellen egenerklæring i databasen:
        If i = sisteindeks - 1 Then
            tilkobling.open

            Dim sporring As String
            sporring = "INSERT INTO egenerklaering (skjema) VALUES (Jasvar)"

        End If
    End Sub

    Private Sub btnForrige_Click(sender As Object, e As EventArgs) Handles btnForrige.Click
        'Funksjon: blar tilbake i spørsmålene, så lenge man ikke har kommet til første spørsmål
        If i > 0 Then
            i = i - 1
            Label1.Text = spoersmaal(i)
        Else
            MsgBox("Ingen flere spørsmål.")
        End If

    End Sub

End Class
