Public Class Blodgiver
    Inherits Bruker
    Private fodselsnummer, blodtype, kontaktform, merknad, timepreferanse As String
    Private siste_blodtapping As Date

    Public Sub New(fodselsnummer As String, blodtype As String,
                   kontaktform As String, merknad As String,
                   timepreferanse As String, siste_blodtapping As Date,
                   epost As String, passord As String,
                   fornavn As String, etternavn As String,
                   adresse As String, telefon1 As String, telefon2 As String,
                   postnr As String, statuskode As Integer)
        MyBase.New(epost, passord, fornavn, etternavn, adresse, telefon1, telefon2, postnr, statuskode)
        Me.fodselsnummer = fodselsnummer
        Me.blodtype = blodtype
        Me.kontaktform = kontaktform
        Me.merknad = merknad
        Me.timepreferanse = timepreferanse
        Me.siste_blodtapping = siste_blodtapping
    End Sub

    Public Property Blodtype1 As String
        Get
            Return blodtype
        End Get
        Set(value As String)
            blodtype = value
        End Set
    End Property

    Public Property Fodselsnummer1 As String
        Get
            Return fodselsnummer
        End Get
        Set(value As String)
            fodselsnummer = value
        End Set
    End Property

    Public Property Kontaktform1 As String
        Get
            Return kontaktform
        End Get
        Set(value As String)
            kontaktform = value
        End Set
    End Property

    Public Property Merknad1 As String
        Get
            Return merknad
        End Get
        Set(value As String)
            merknad = value
        End Set
    End Property

    Public Property Siste_blodtapping1 As Date
        Get
            Return siste_blodtapping
        End Get
        Set(value As Date)
            siste_blodtapping = value
        End Set
    End Property

    Public Property Timepreferanse1 As String
        Get
            Return timepreferanse
        End Get
        Set(value As String)
            timepreferanse = value
        End Set
    End Property
End Class
