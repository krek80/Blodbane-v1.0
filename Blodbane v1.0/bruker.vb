Public Class Bruker
    Private epost, passord, fornavn, etternavn, adresse, telefon1, telefon2, postnr, status As String

    Public Sub New(epost As String, passord As String, fornavn As String, etternavn As String, adresse As String, telefon1 As String, telefon2 As String, postnr As String, status As String)
        Me.epost = epost
        Me.passord = passord
        Me.fornavn = fornavn
        Me.etternavn = etternavn
        If Not IsDBNull(adresse) Then
            Me.adresse = adresse
        Else
            Me.adresse = ""
        End If
        Me.telefon1 = telefon1
        Me.telefon2 = telefon2
        Me.postnr = postnr
        Me.status = status
    End Sub

    Public Property Adresse1 As String
        Get
            Return adresse
        End Get
        Set(value As String)
            adresse = value
        End Set
    End Property

    Public Property Epost1 As String
        Get
            Return epost
        End Get
        Set(value As String)
            epost = value
        End Set
    End Property

    Public Property Etternavn1 As String
        Get
            Return etternavn
        End Get
        Set(value As String)
            etternavn = value
        End Set
    End Property

    Public Property Fornavn1 As String
        Get
            Return fornavn
        End Get
        Set(value As String)
            fornavn = value
        End Set
    End Property

    Public Property Passord1 As String
        Get
            Return passord
        End Get
        Set(value As String)
            passord = value
        End Set
    End Property

    Public Property Postnr1 As String
        Get
            Return postnr
        End Get
        Set(value As String)
            postnr = value
        End Set
    End Property

    Public Property Status1 As String
        Get
            Return status
        End Get
        Set(value As String)
            status = value
        End Set
    End Property

    Public Property Telefon11 As String
        Get
            Return telefon1
        End Get
        Set(value As String)
            telefon1 = value
        End Set
    End Property

    Public Property Telefon21 As String
        Get
            Return telefon2
        End Get
        Set(value As String)
            telefon2 = value
        End Set
    End Property
End Class
