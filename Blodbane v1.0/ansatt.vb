Public Class Ansatt
    Inherits Bruker
    Private ansattnummer As Integer

    Public Sub New(ansattnummer As Integer, epost As String, passord As String, fornavn As String, etternavn As String, adresse As String, telefon1 As String, telefon2 As String, postnr As String, status As String)
        MyBase.New(epost, passord, fornavn, etternavn, adresse, telefon1, telefon2, postnr, status)
        Me.ansattnummer = ansattnummer
    End Sub

    Public Property Ansattnummer1 As Integer
        Get
            Return ansattnummer
        End Get
        Set(value As Integer)
            ansattnummer = value
        End Set
    End Property
End Class
