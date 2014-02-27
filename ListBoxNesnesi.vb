Public Class ListBoxNesnesi

#Region "Yerel Degiskenler"
    Private mValue As String
    Private mNo As Integer
#End Region

#Region "Kurucu Kod"
    Public Sub New(ByVal strAd As String, ByVal intNo As Integer)
        mValue = strAd
        mNo = intNo
    End Sub

    Public Sub New()
        mValue = ""
        mNo = 0
    End Sub
#End Region

#Region "Ozellikler"
    Property value()
        Get
            Return mValue
        End Get
        Set(ByVal Value)
            mValue = Value
        End Set
    End Property

    Property No()
        Get
            Return mNo
        End Get
        Set(ByVal Value)
            mNo = Value
        End Set
    End Property
#End Region
    
    Public Overrides Function ToString() As String
        Return mValue
    End Function
End Class
