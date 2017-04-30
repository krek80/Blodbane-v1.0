Public Class Egenerklaering
    Private id As Integer
    Private bgEpost, ansattEpost, skjema, kommentar As String
    Private datotidBG, datotidAnsatt As Date

    Public Property Id1 As Integer
        Get
            Return id
        End Get
        Set(value As Integer)
            id = value
        End Set
    End Property

    Public Property BgEpost1 As String
        Get
            Return bgEpost
        End Get
        Set(value As String)
            bgEpost = value
        End Set
    End Property

    Public Property AnsattEpost1 As String
        Get
            Return ansattEpost
        End Get
        Set(value As String)
            ansattEpost = value
        End Set
    End Property

    Public Property Skjema1 As String
        Get
            Return skjema
        End Get
        Set(value As String)
            skjema = value
        End Set
    End Property

    Public Property Kommentar1 As String
        Get
            Return kommentar
        End Get
        Set(value As String)
            kommentar = value
        End Set
    End Property

    Public Property DatotidBG1 As Date
        Get
            Return datotidBG
        End Get
        Set(value As Date)
            datotidBG = value
        End Set
    End Property

    Public Property DatotidAnsatt1 As Date
        Get
            Return datotidAnsatt
        End Get
        Set(value As Date)
            datotidAnsatt = value
        End Set
    End Property

    Public Sub New(id As Integer, bgEpost As String, ansattEpost As String, skjema As String, kommentar As String, datotidBG As Date, datotidAnsatt As Date)
        Me.id = id
        Me.bgEpost = bgEpost
        Me.ansattEpost = ansattEpost
        Me.skjema = skjema
        Me.kommentar = kommentar
        Me.datotidBG = datotidBG
        Me.datotidAnsatt = datotidAnsatt
    End Sub
End Class
