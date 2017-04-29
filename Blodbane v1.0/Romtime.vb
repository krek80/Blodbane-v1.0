Public Class Romtime
    Private datotid As Date
    Private romnummer As String
    Private timenr As Integer

    Public Sub New(datotid As Date, romnummer As String, timenr As Integer)
        Me.datotid = datotid
        Me.romnummer = romnummer
        Me.timenr = timenr
    End Sub

    Public Property Datotid1 As Date
        Get
            Return datotid
        End Get
        Set(value As Date)
            datotid = value
        End Set
    End Property

    Public Property Romnummer1 As String
        Get
            Return romnummer
        End Get
        Set(value As String)
            romnummer = value
        End Set
    End Property

    Public Property Timenr1 As Integer
        Get
            Return timenr
        End Get
        Set(value As Integer)
            timenr = value
        End Set
    End Property
End Class
