Option Explicit On

Public Class CmnUtility

#Region "変数定義"

    ''' <summary>
    ''' 値の型
    ''' </summary>
    Public Enum eValueType
        [String]
        Number
    End Enum

#End Region

#Region "プロパティ"


#End Region

#Region "コンストラクタ"

    ''' <summary>
    ''' 新規作成
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

    End Sub

#End Region

#Region "メソッド"
    ''' <summary>
    ''' SQL用に値を変換する ブランクの場合はNullとする
    ''' </summary>
    ''' <param name="value">代入値</param>
    ''' <param name="b_Number">数値ﾌﾗｸﾞ True:数値 False:文字</param>
    ''' <param name="b_BlankZero">数値の場合 ブランクを０にするか True:0 False:Null</param>
    ''' <returns ></returns>
    ''' <remarks></remarks>
    Public Function ToSql(ByVal value As Object, Optional ByVal b_Number As Boolean = False, Optional ByVal b_BlankZero As Boolean = True) As String

        Dim strRet As String = String.Empty

        Try
            'Nothingの場合
            If value Is Nothing Then Return "Null"

            '文字に変換
            strRet = value.ToString

            '型別
            If b_Number Then
                '数値

                '数値チェック
                If Not IsNumeric(strRet) Then
                    strRet = String.Empty
                End If

                '
                Select Case strRet
                    Case String.Empty
                        '空白
                        If b_BlankZero Then
                            ' ゼロを登録
                            strRet = "0"
                        Else
                            ' Nullを登録
                            strRet = "Null"
                        End If

                    Case Else

                End Select

            Else
                '文字

                ' 特定文字を変換
                ' シングルクォートの変換
                strRet = strRet.TrimEnd.Replace("'", "''") 'シングルクォート

                Select Case strRet
                    Case String.Empty
                        '空文字列

                        ' Nullを登録
                        strRet = "Null"

                    Case Else
                        ' シングルクォートで囲む
                        strRet = "'" & strRet & "'"

                End Select

            End If

            'If value.Equals(String.Empty) Then
            '    '空白
            '    If b_Number Then
            '        ' 数値の場合
            '        If b_BlankZero Then
            '            ' ゼロを登録
            '            value = "0"
            '        Else
            '            ' Nullを登録
            '            value = "Null"
            '        End If
            '    Else

            '        value = "Null"
            '    End If

            'ElseIf value.ToString.Trim.Equals("-") Then
            '    ' ハイフンのみ場合はゼロに変換する
            '    strRet = "0"

            'ElseIf b_Number = False Then   '文字列
            '    ' シングルクォート・立法メートル・リットルの変換
            '    strRet = "'" & value.ToString.TrimEnd.Replace("'", "''").Replace(ChrW(13221), "m3").Replace(ChrW(8467), "L") & "'"

            'End If

        Catch ex As Exception
            Throw  ' 上層にExceptionを投げる
        End Try

        Return strRet

    End Function

    ''' <summary>
    ''' SQL用に値を変換する ブランクの場合はNullとする
    ''' </summary>
    ''' <param name="value">代入値</param>
    ''' <returns ></returns>
    ''' <remarks></remarks>
    Public Function NullToEmpty(ByVal value As Object) As String

        Try
            If value Is Nothing Then Return String.Empty
            If value.Equals(DBNull.Value) Then Return String.Empty

            Return value.ToString

        Catch ex As Exception
            Throw  ' 上層にExceptionを投げる
        End Try

    End Function

    ''' <summary>
    ''' SQL用に値を変換する ブランクの場合はNullとする
    ''' </summary>
    ''' <param name="value">代入値</param>
    ''' <param name="ValueType">数値ﾌﾗｸﾞ True:数値 False:文字</param>
    ''' <param name="b_BlankZero">数値の場合 ブランクを０にするか True:0 False:Null</param>
    ''' <returns ></returns>
    ''' <remarks></remarks>
    Public Function ToSql2(ByVal value As Object, ByVal ValueType As eValueType, Optional ByVal b_BlankZero As Boolean = True) As String

        Dim strRet As String = String.Empty

        Try

            If Not value Is Nothing Then
                strRet = value.ToString
            End If

            Select Case ValueType
                Case eValueType.Number
                    ' 数値の場合
                    If strRet.Equals(String.Empty) Then

                        If b_BlankZero Then
                            '数値の場合 ブランクを０にするか
                            ' ゼロを登録
                            strRet = "0"
                        Else
                            ' Nullを登録
                            strRet = "Null"
                        End If
                    End If

                Case eValueType.String
                    If strRet.Equals(String.Empty) Then
                        strRet = "Null"
                    Else
                        ' シングルクォートの変換
                        strRet = strRet.Replace("'", "''")
                        ' シングルクォートで囲む
                        strRet = "'" & strRet & "'"
                    End If

            End Select

        Catch ex As Exception
            Throw  ' 上層にExceptionを投げる
        End Try

        Return strRet

    End Function

    ''' <summary>
    ''' InfoLog メッセージ
    ''' </summary>
    ''' <returns ></returns>
    ''' <remarks></remarks>
    Public Function InfoLogFormat(ByVal DeviceID As String, ByVal TableName As String,
                        ByVal ProcessID As String,
                        ByVal ProcessName As String,
                        ByVal ProcessSubName As String,
                        ByVal Message As String,
                        Optional ByVal addInfo As String = "") As String

        Dim strRet As String = String.Empty

        Try

            '出力メッセージ設定
            strRet &= String.Format("[{0}]", DeviceID)
            strRet &= String.Format("[{0} {1}/{2}-{3}]", TableName, ProcessID, ProcessName, ProcessSubName)
            strRet &= " "
            strRet &= Message
            strRet &= " "
            strRet &= addInfo

        Catch ex As Exception
            Throw  ' 上層にExceptionを投げる
        End Try

        Return strRet

    End Function

    ''' <summary>
    ''' DataLog メッセージ
    ''' </summary>
    ''' <returns ></returns>
    ''' <remarks>2019/11/14 add</remarks>
    Public Function DataLogFormat(ByVal DeviceID As String, ByVal TableName As String,
                        ByVal ProcessID As String,
                        ByVal ProcessName As String,
                        ByVal ProcessSubName As String,
                        ByVal Message As String,
                        Optional ByVal addInfo As String = "") As String

        Dim strRet As String = String.Empty

        Try

            '出力メッセージ設定
            strRet &= String.Format(",[{0}],", DeviceID)
            strRet &= String.Format("{0},", TableName) 'テーブル名
            strRet &= Message
            strRet &= " "
            strRet &= addInfo

        Catch ex As Exception
            Throw  ' 上層にExceptionを投げる
        End Try

        Return strRet

    End Function

#End Region

    ''' <summary>
    ''' 文字列バイト数取得
    ''' </summary>
    ''' <param name="targetString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetStringByte(ByVal targetString As String) As Integer

        ' 変数定義
        Dim encode As Encoding

        ' エンコーディング取得
        encode = Encoding.GetEncoding("Shift_JIS")

        ' 文字列バイト数取得
        Return encode.GetByteCount(targetString)

    End Function

    ''' <summary>
    ''' 文字列カット
    ''' </summary>
    ''' <param name="targetString">対象文字列</param>
    ''' <param name="cutSize">サイズ</param>
    ''' <param name="cutFront">先頭からカットするか(T:先頭 F:後方)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CutString(ByVal targetString As String, ByVal cutSize As Integer, Optional ByVal cutFront As Boolean = False) As String

        ' 変数定義
        Dim encodeString As Byte()
        Dim returnString As String
        Dim encode As Encoding

        ' エンコーディング取得
        encode = Encoding.GetEncoding("Shift_JIS")

        If targetString Is Nothing Then

            Return String.Empty

        End If

        ' 文字サイズがカットサイズ以下の場合は処理を抜ける
        If encode.GetByteCount(targetString) <= cutSize Then

            Return targetString

        End If

        ' 文字列カット
        encodeString = encode.GetBytes(targetString)
        If cutFront.Equals(True) Then

            ' マルチバイト文字の分断を避けるため、一文字余分に取得してからその文字を削除する
            returnString = encode.GetString(encodeString, encode.GetByteCount(targetString) - (cutSize + 1), cutSize + 1)
            returnString = returnString.Remove(0, 1)

        Else

            ' マルチバイト文字の分断を避けるため、一文字余分に取得してからその文字を削除する
            returnString = encode.GetString(encodeString, 0, (cutSize + 1))
            returnString = returnString.Remove(returnString.Length - 1)

        End If

        Return returnString

    End Function
End Class
