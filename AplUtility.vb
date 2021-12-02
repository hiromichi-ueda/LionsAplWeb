Imports System.IO
Imports log4net
Public Class AplUtility

    ''' プロパティ
    Protected _logger As ILog

    Public Region As String                 ' リジョン
    Public Zone As String                   ' ゾーン
    Public ClubCode As String               ' クラブコード
    Public MemberCode As String             ' 会員番号
    Public PeriodStart As String            ' 期開始（年度開始日月）
    Public PeriodEnd As String              ' 期終了（年度終了日月）
    Public LastUpdDate As String            ' 最終更新日時

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    Sub New(ByRef logger As ILog)
        _logger = logger
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLite情報(A_ACCOUNT)から情報を取得する
    ''' </summary>
    ''' <param name="arrSQLite">SQLite情報(A_ACCOUNT)</param>
    Sub SetAccountInfoSQLite(ByRef arrSQLite As ArrayList)

        _logger.Info(" ▽ SetAccountInfoSQLite Start ▽")
        Try
            For Each dicWrk As Dictionary(Of String, String) In arrSQLite
                Region = dicWrk("Region")
                _logger.Info("  ○Read : " & Region)
                Zone = dicWrk("Zone")
                _logger.Info("   ○Zone : " & Zone)
                ClubCode = dicWrk("ClubCode")
                _logger.Info("   ○ClubCode : " & ClubCode)
                MemberCode = dicWrk("MemberCode")
                _logger.Info("   ○MemberCode : " & MemberCode)
                LastUpdDate = dicWrk("LastUpdDate")
                _logger.Info("   ○LastUpdateDate : " & LastUpdDate)
            Next

        Catch ex As Exception
            _logger.Error(" ●SetAccountInfoSQLite 例外エラー:", ex)
            Throw

        End Try
        _logger.Info(" △ SetAccountInfoSQLite End △")

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLite情報(A_SETTING)から年度情報を設定する
    ''' </summary>
    ''' <param name="arrSQLite">SQLite情報(A_SETTING)</param>
    Sub SetFiscalInfoSQLite(ByRef arrSQLite As ArrayList)

        For Each dicWrk As Dictionary(Of String, String) In arrSQLite
            PeriodStart = dicWrk("PeriodStart")
            PeriodEnd = dicWrk("PeriodEnd")
        Next

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServer情報(A_SETTING)から年度情報を設定する
    ''' </summary>
    ''' <param name="arrSQLsvr">SQLServer情報(A_SETTING)</param>
    Sub SetFiscalInfoSQLsvr(ByRef arrSQLsvr As ArrayList)

        For Each dicWrk As Dictionary(Of String, String) In arrSQLsvr
            PeriodStart = dicWrk("PeriodStart")
            PeriodEnd = dicWrk("PeriodEnd")
        Next

    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' 入力年月日の年度を取得する
    ''' </summary>
    ''' <param name="InYmd">年度を取得する年月日</param>
    ''' <returns>入力年月日の年度</returns>
    Function GetFiscalYear(InYmd As String) As String
        Dim ReFiscalYear As String
        Dim InY As String
        Dim InMd As String

        ' 指定年月日の年を取得
        InY = InYmd.Substring(0, 4)
        ' 指定年月日の月日を取得
        InMd = InYmd.Substring(5, 5)

        ' 指定年月日の年度チェック
        ' 戻り値に入力年を設定する。
        ReFiscalYear = InY
        ' 年度開始月日より入力月日が前の場合
        If InMd < PeriodStart Then
            ' 戻り値に昨年を設定する
            Dim lastYear As Integer = Integer.Parse(InY) - 1
            ReFiscalYear = lastYear.ToString
        End If

        Return ReFiscalYear

    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' 指定年度の年度開始年月日を取得する。
    ''' </summary>
    ''' <param name="InY">年度</param>
    ''' <returns></returns>
    Function GetFiscalStartYmdFromY(InY As String) As String
        Dim ReFiscalStartYmd As String = InY & PeriodStart

        Return ReFiscalStartYmd

    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' 指定年度の年度終了年月日を取得する。
    ''' </summary>
    ''' <param name="InY">年度</param>
    ''' <returns></returns>
    Function GetFiscalEndYmdFromY(InY As String) As String
        Dim ReFiscalEndYmd As String = InY & PeriodEnd

        Return ReFiscalEndYmd

    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' 指定フォルダに指定ファイルと同名のファイルがあるか確認する
    ''' 同名のファイルがある場合、ファイル名の後ろに数値を追加して返す
    ''' </summary>
    ''' <param name="filepath"></param>
    ''' <param name="filename"></param>
    ''' <returns></returns>
    Function ChkFileName(filepath As String, filename As String) As String
        ' 戻り値(入力ファイル名を設定)
        Dim RetFileName = filename
        ' ファイルパス＋ファイル名
        Dim FilePN = filepath & filename

        _logger.Info("▽ ChkFileName Start ▽")
        _logger.Info("   ○filepath : " & filepath)
        _logger.Info("   ○filename : " & filename)

        If System.IO.File.Exists(FilePN) Then
            ' 同じファイル名がある場合

            ' ファイル名（拡張子なし）
            Dim FileNameNex = System.IO.Path.GetFileNameWithoutExtension(filename)
            ' ファイル拡張子
            Dim FileNameEx = System.IO.Path.GetExtension(filename)
            ' ファイルパス＋ファイル名（拡張子なし）
            Dim FilePNNex = filepath & FileNameNex
            ' 同名ファイル用カウンタ取得
            Dim counter = GetCountFile(filepath, LionsAplWeb.My.Settings.CountFile)

            ' 番号をつけたファイル名を返す（ファイル名のみ）
            'RetFileName = filepath & FileNameNex & counter & FileNameEx
            RetFileName = FileNameNex & counter & FileNameEx

        End If

        _logger.Info("   ○RetFileName : " & RetFileName)
        _logger.Info("△ ChkFileName End △")

        Return RetFileName

    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' 同名ファイル用カウンタファイルから値を取得＆カウントして更新
    ''' </summary>
    ''' <param name="cfpath">カウンタファイルパス</param>
    ''' <param name="cfname">カウンタファイル名</param>
    ''' <returns></returns>
    Function GetCountFile(cfpath As String, cfname As String) As UInteger
        Dim gcnt As UInteger        ' ファイルから取得した値
        Dim scnt As UInteger        ' ファイルに書き込む値
        Dim ret As UInteger = 0     ' 戻り値

        Dim finfo As New System.IO.FileInfo(cfpath + cfname)
        Dim stream As System.IO.Stream = Nothing
        Dim lp As Integer = 0
        Dim count As Integer = 0

        _logger.Info("▽ GetCountFile Start ▽")
        _logger.Info("   ○cfpath : " & cfpath)
        _logger.Info("   ○cfname : " & cfname)

        While True
            Try
                ' ファイルオープン
                stream = finfo.Open(IO.FileMode.OpenOrCreate, IO.FileAccess.ReadWrite, IO.FileShare.None)
                Dim bs(stream.Length - 1) As Byte

                ' ファイル読込
                stream.Read(bs, 0, bs.Length)
                If bs.Length = 0 Then
                    ' 値なし⇒0を返す
                    gcnt = 0
                Else
                    ' 値あり⇒バイトから数値に変換
                    gcnt = BitConverter.ToUInt32(bs, 0)
                End If
                _logger.Info("   ○gcnt : " & gcnt)
                ' 書込値用にカウント
                scnt = gcnt + 1
                ' カウントが100万に達したら0に戻る
                If scnt > 999999 Then
                    scnt = 0
                End If
                ' 書込値をバイトに変換
                _logger.Info("   ○scnt : " & scnt)
                Dim bss() As Byte = BitConverter.GetBytes(scnt)

                _logger.Info("   ○bss.Length : " & bss.Length)
                ' ファイル書込
                stream.Position = 0
                stream.Write(bss, 0, bss.Length)
                ' ファイルクロース
                stream.Close()
                ret = gcnt
                Exit While

            Catch ex As IO.IOException
                ' ロック解除待ち
                lp += 1
                If lp > 100 Then
                    _logger.Error("●Count File Lock Error : " & ex.Message)
                    Throw ex ' 10秒(0.1秒×100回)で例外送出
                End If
                System.Threading.Thread.Sleep(100) ' 0.1秒待つ

            Finally
                If Not stream Is Nothing Then stream.Close()
            End Try
        End While

        _logger.Info("   ○ret : " & ret)
        _logger.Info("△ GetCountFile End △")

        Return ret

    End Function

End Class
