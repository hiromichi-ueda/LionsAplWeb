Public Class clsSQLString

#Region "変数定義"

    Private _str_SQL As StringBuilder

#End Region

#Region "プロパティ"

    Public Property sb() As StringBuilder
        Get
            Return _str_SQL
        End Get
        Set(ByVal Value As StringBuilder)
            _str_SQL = Value
        End Set
    End Property

    Public ReadOnly Property ToStr() As String
        Get
            Return _str_SQL.ToString
        End Get
    End Property
    Public ReadOnly Property ToStrLog() As String
        Get
            Return "SQL:" & System.Environment.NewLine & _str_SQL.ToString
        End Get
    End Property

#End Region

#Region "コンストラクタ"

    ''' <summary>
    ''' 新規作成
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

        _str_SQL = New StringBuilder

    End Sub

#End Region

#Region "メソッド"

    ''' <summary>
    ''' SQL文クリア
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Remove(Optional ByVal intStartIndex As Integer = 0, Optional ByVal intLength As Integer = 0)

        Try

            If intLength = 0 Then intLength = _str_SQL.ToString.Length

            _str_SQL.Remove(intStartIndex, intLength)

        Catch ex As Exception
            Throw ' 上層にExceptionを投げる
        End Try

    End Sub

    ''' <summary>
    ''' SQL文追加
    ''' </summary>
    ''' <remarks>空白,改行あり</remarks>
    Public Sub Append(ByVal strSQL As String)

        Try

            _str_SQL.Append(strSQL & " " & Environment.NewLine)

        Catch ex As Exception
            Throw ' 上層にExceptionを投げる
        End Try

    End Sub

#End Region

End Class
