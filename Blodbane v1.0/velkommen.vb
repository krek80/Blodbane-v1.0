Public Class velkommen
    Dim fremdrift As Integer
    Private Sub velkommen_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        fremdrift = fremdrift + 20
        ProgressBar1.Value = fremdrift
        If ProgressBar1.Value >= 100 Then
            Timer1.Stop()
            Me.Close()
            Blodbane.Show()
        End If
    End Sub
End Class