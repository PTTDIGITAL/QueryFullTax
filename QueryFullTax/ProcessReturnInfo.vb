Imports Microsoft.VisualBasic

Public Class ProcessReturnInfo
    Dim _IsSuccess As Boolean = False
    Dim _ErrorMessage As String = ""
    Dim _DT As New DataTable

    Public Property IsSuccess As Boolean
        Get
            Return _IsSuccess
        End Get
        Set(value As Boolean)
            _IsSuccess = value
        End Set
    End Property

    Public Property ErrorMessage As String
        Get
            Return _ErrorMessage.Trim
        End Get
        Set(value As String)
            _ErrorMessage = value
        End Set
    End Property

    Public Property DT As DataTable
        Get
            Return _DT
        End Get
        Set(value As DataTable)
            _DT = value
        End Set
    End Property
End Class
