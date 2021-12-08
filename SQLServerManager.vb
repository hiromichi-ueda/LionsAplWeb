Option Explicit On
Option Strict On

Imports System.IO
Imports System.Web
Imports System.Data.SqlClient
Imports log4net
Public Class SQLServerManager

    ' エラー発生時に返すHTTPコード
    Private Const httpErrorCode As Integer = 500
    ' データ採番用データ種類：地区誌購入者
    Private Const ST_DATACLASS_MAGAZINEBUY As String = "7"

    Protected _logger As ILog
    Protected _utl As CmnUtility = Nothing                                          ' ユーティリティ
    Protected _TblMgr As TableManager = Nothing                                     ' テーブルマネージャー
    Protected _context As HttpContext = Nothing                                     'コンテキストクラス
    'Protected _SqlCommand As SqlCommand

    Protected _strTableName As String = String.Empty                                ' 操作対象テーブル名

    Protected _appPath As String = AppDomain.CurrentDomain.BaseDirectory.TrimEnd    ' 実行パス

    ' HOME用検索条件
    Protected _aplDate As DateTime                                                  ' アプリケーション時刻
    Protected _t_eventret_MonOfs As DateTime                                        ' T_EVENTRET取得対象月のオフセット
    Protected _t_infomation_Cabi_MonOfs As DateTime                                 ' T_INFOMATION(CABINET)取得対象月のオフセット
    Protected _t_director_MonOfs As DateTime                                        ' T_DIRECTOR取得対象月のオフセット
    Protected _t_infomation_Club_MonOfs As DateTime                                 ' T_INFOMATION(CLUB)取得対象月のオフセット
    Protected _strNowFiscal As String = String.Empty                                ' 今年度(年)
    Protected _strPreFiscal As String = String.Empty                                ' 前年度(年)
    'Protected _strNowFiscalStart As String = String.Empty                           ' 今年度開始年月日
    'Protected _strNowFiscalEnd As String = String.Empty                             ' 今年度終了年月日
    Protected _strPreFiscalStart As String = String.Empty                           ' 前年度開始年月日
    'Protected _strPreFiscalEnd As String = String.Empty                             ' 前年度終了年月日
    Protected _strDataClass As String = String.Empty                                ' データ種類

    Public aplUtil As AplUtility                                                    ' アプリ用ユーティリティクラス

    'Public arrSQLsrv As New ArrayList ' SQLsrv情報
    'キャビネット用
    Public arrSQLsrv_A_SETTING As New ArrayList             ' SQLsrv情報(A_SETTING用)
    Public arrSQLsrv_A_DATANO As New ArrayList              ' SQLsrv情報(A_DATANO用)
    Public arrSQLsrv_A_FILEPATH As New ArrayList            ' SQLsrv情報(A_FILEPATH用)
    Public arrSQLsrv_T_SLOGAN As New ArrayList              ' SQLsrv情報(T_SLOGAN用)
    Public arrSQLsrv_T_LETTER As New ArrayList              ' SQLsrv情報(T_LETTER用)
    Public arrSQLsrv_T_EVENTRET As New ArrayList            ' SQLsrv情報(T_EVENTRET用)
    Public arrSQLsrv_T_EVENT As New ArrayList               ' SQLsrv情報(T_EVENT用)
    Public arrSQLsrv_T_INFOMATION_CABI As New ArrayList     ' SQLsrv情報(T_INFOMATION用(キャビネット用))
    Public arrSQLsrv_T_MAGAZINE As New ArrayList            ' SQLsrv情報(T_MAGAZINE用)
    Public arrSQLsrv_T_MAGAZINEBUY As New ArrayList         ' SQLsrv情報(T_MAGAZINEBUY用)
    Public arrSQLsrv_M_DISTRICTOFFICER As New ArrayList     ' SQLsrv情報(M_DISTRICTOFFICER用)
    Public arrSQLsrv_M_CABINET As New ArrayList             ' SQLsrv情報(M_CABINET用)
    Public arrSQLsrv_M_CLUB As New ArrayList                ' SQLsrv情報(M_CLUB用)
    Public arrSQLsrv_M_AREA As New ArrayList                ' SQLsrv情報(M_AREA用)     2021/12 ADD
    Public arrSQLsrv_M_JOB As New ArrayList                 ' SQLsrv情報(M_JOB用)     2021/12 ADD
    Public arrSQLsrv_T_MATCHING As New ArrayList            ' SQLsrv情報(T_MATCHING用)     2021/12 ADD
    'クラブ用
    Public arrSQLsrv_T_CLUBSLOGAN As New ArrayList          ' SQLsrv情報(T_CLUBSLOGAN用)
    Public arrSQLsrv_T_MEETINGSCHEDULE As New ArrayList     ' SQLsrv情報(T_MEETINGSCHEDULE用)
    Public arrSQLsrv_T_DIRECTOR As New ArrayList            ' SQLsrv情報(T_DIRECTOR用)
    Public arrSQLsrv_T_MEETINGPROGRAM As New ArrayList      ' SQLsrv情報(T_MEETINGPROGRAM用)
    Public arrSQLsrv_T_INFOMATION_CLUB As New ArrayList     ' SQLsrv情報(T_INFOMATION用（クラブ用）)
    Public arrSQLsrv_M_MEMBER As New ArrayList              ' SQLsrv情報(M_MEMBER用)

    ' 更新フラグ（False=更新なし/True=更新あり）
    Public UpFlg_A_SETTING As Boolean
    Public UpFlg_A_FILEPATH As Boolean
    Public UpFlg_T_SLOGAN As Boolean
    Public UpFlg_T_LETTER As Boolean
    Public UpFlg_T_EVENTRET As Boolean
    Public UpFlg_T_EVENT As Boolean
    Public UpFlg_T_INFOMATIONCABI As Boolean
    Public UpFlg_T_MAGAZINE As Boolean
    Public UpFlg_T_MAGAZINEBUY As Boolean
    Public UpFlg_M_CABINET As Boolean
    Public UpFlg_M_CLUB As Boolean
    Public UpFlg_M_AREA As Boolean      '2021/12 ADD
    Public UpFlg_M_JOB As Boolean      '2021/12 ADD
    Public UpFlg_T_MATCHING As Boolean      '2021/12 ADD
    Public UpFlg_T_CLUBSLOGAN As Boolean
    Public UpFlg_T_MEETINGSCHEDULE As Boolean
    Public UpFlg_T_DIRECTOR As Boolean
    Public UpFlg_T_MEETINGPROGRAM As Boolean
    Public UpFlg_T_INFOMATION As Boolean
    Public UpFlg_M_MEMBER As Boolean


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="logger">ログクラス</param>
    Public Sub New(ByRef logger As ILog, ByRef context As HttpContext)
        _logger = logger
        aplUtil = New AplUtility(logger)
        _context = context

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' HOME処理の初期化
    ''' </summary>
    ''' <param name="arrSQLite_A_ACCOUNT">SQLiteから取得したA_ACCOUNT情報</param>
    ''' <param name="arrSQLite_A_SETTING">SQLiteから取得したA_SETTING情報</param>
    ''' <param name="aplDate">アプリの日時</param>
    Public Sub Init_HOME(arrSQLite_A_ACCOUNT As ArrayList,
                         arrSQLite_A_SETTING As ArrayList,
                         aplDate As DateTime)
        Dim strDate As String                           ' 時刻文字列

        _logger.Info(" ▽ Start Init_HOME ▽") 'ログ

        ' アプリ時刻を取得
        _aplDate = aplDate

        ' 更新フラグをクリアする
        ClearUpdFlgAll()

        ' T_EVENTRET取得対象月のオフセット
        _t_eventret_MonOfs = _aplDate.AddMonths(My.Settings.T_EVENTRET_MonOfs)

        ' T_INFOMATION(CABINET)取得対象月のオフセット
        _t_infomation_Cabi_MonOfs = _aplDate.AddMonths(My.Settings.T_INFOMATION_CABI_MonOfs)

        ' T_DIRECTOR取得対象月のオフセット
        _t_director_MonOfs = _aplDate.AddMonths(My.Settings.T_DIRECTOR_MonOfs)

        ' T_INFOMATION(CLUB)取得対象月のオフセット
        _t_infomation_Club_MonOfs = _aplDate.AddMonths(My.Settings.T_INFOMATION_CLUB_MonOfs)

        ' 現在時刻の年月日を文字列で取得
        strDate = _aplDate.ToString("yyyy/MM/dd")

        ' SQLite情報(A_ACCOUNT)から情報を取得する
        aplUtil.SetAccountInfoSQLite(arrSQLite_A_ACCOUNT)

        ' SQLite情報(A_SETTING)から年度情報を設定する
        aplUtil.SetFiscalInfoSQLite(arrSQLite_A_SETTING)

        ' 年度設定
        ' 今年度
        _strNowFiscal = aplUtil.GetFiscalYear(strDate)
        ' 前年度
        _strPreFiscal = (Long.Parse(_strNowFiscal) - 1).ToString
        ' 前年度開始年月日
        _strPreFiscalStart = _strPreFiscal & "/" & aplUtil.PeriodStart

        _logger.Info(" △ End Init_HOME △") 'ログ

    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQL Serverよりデータを読み込む（ACCOUNT用）
    ''' </summary>
    ''' <param name="context">エラーステータス出力用</param>
    Public Sub Read_SQLServer_ACCOUNT(ByRef context As HttpContext)

        _logger.Info(" ▽ Read SQLServer ACCOUNT Start ▽") 'ログ

        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        ' SQL Serverより データを取得 (334B)
        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

        ' 2021/10/18 TOPデータ取得機能追加により、設定ファイルの取得処理を移動
        '' SQLServerコネクション生成(334B)
        'Using _connSvDB As SqlConnection = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("LionsAplWeb.My.MySettings.DB334B").ConnectionString)
        '    _connSvDB.Open()

        '    'SQLServerコマンド生成
        '    Using _sqlCommand As SqlCommand = _connSvDB.CreateCommand()
        '        Try
        '            ' SQLServerのA_SETTINGテーブルから情報を取得する
        '            GetA_SETTING(_sqlCommand)

        '        Catch ex As Exception
        '            ' エラーログ出力
        '            context.Response.StatusCode = httpErrorCode
        '            _logger.Error("●Read_SQLServer_ACCOUNT DB Error : ", ex)
        '            Throw

        '        End Try

        '    End Using
        '    _connSvDB.Close()

        'End Using

        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        ' SQL Serverより データを取得 (334BCLUB)
        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

        ' SQLServerコネクション生成(334BCLUB)
        Using _connSvDBCLUB As SqlConnection = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("LionsAplWeb.My.MySettings.DB334BCLUB").ConnectionString)
            _connSvDBCLUB.Open()

            'SQLServerコマンド生成
            Using _sqlCommand As SqlCommand = _connSvDBCLUB.CreateCommand()
                Try
                    ' SQLServerのM_MEMBERテーブルから情報を取得する
                    GetM_MEMBER(_sqlCommand)

                Catch ex As Exception
                    ' エラーログ出力
                    context.Response.StatusCode = httpErrorCode
                    _logger.Error("●Read_SQLServer_ACCOUNT DBCLUB Error : ", ex)
                    Throw

                End Try

            End Using
            _connSvDBCLUB.Close()

        End Using

        _logger.Info(" △ End Read SQLServer △") 'ログ

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQL Serverよりデータを読み込む（HOME用）
    ''' </summary>
    ''' <param name="context">エラーステータス出力用</param>
    Public Sub Read_SQLServer_HOME(ByRef context As HttpContext)

        _logger.Info(" ▽ Read SQLServer HOME Start ▽") 'ログ

        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        ' SQL Serverより データを取得 (334B)
        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

        ' SQLServerコネクション生成(334B)
        Using _connSvDB As SqlConnection = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("LionsAplWeb.My.MySettings.DB334B").ConnectionString)
            _connSvDB.Open()

            'SQLServerコマンド生成
            Using _sqlCommand As SqlCommand = _connSvDB.CreateCommand()
                Try
                    ' 2021/10/18 TOPデータ取得機能追加により、添付パスファイルの取得処理を移動
                    '' SQLServerのA_FILEPATHテーブルから情報を取得する
                    'GetA_FILEPATH(_sqlCommand)

                    ' 2021/10/18 TOPデータ取得機能追加により、地区スローガンの取得処理を移動
                    '' 地区スローガン
                    '_strTableName = "T_SLOGAN"
                    '' 対象テーブル更新チェック
                    'UpFlg_T_SLOGAN = ChkUpdate(_sqlCommand)
                    'If UpFlg_T_SLOGAN Then
                    '    ' SQLServerのT_SLOGANテーブルから情報を取得する
                    '    GetT_SLOGAN(_sqlCommand)
                    'End If

                    ' キャビネットレター
                    _strTableName = "T_LETTER"
                    ' 対象テーブル更新チェック
                    UpFlg_T_LETTER = ChkUpdate(_sqlCommand)
                    If UpFlg_T_LETTER Then
                        ' SQLServerのT_LETTERテーブルから情報を取得する
                        GetT_LETTER(_sqlCommand)
                    End If

                    ' イベント出欠
                    _strTableName = "T_EVENTRET"
                    ' 対象テーブル更新チェック
                    UpFlg_T_EVENTRET = ChkUpdate(_sqlCommand)
                    If UpFlg_T_EVENTRET Then
                        ' SQLServerのT_EVENTRETテーブルから情報を取得する
                        GetT_EVENTRET(_sqlCommand)
                    End If

                    ' キャビネットイベント
                    ' ※イベント出欠の結果を元にキャビネットイベントの情報を取得するため、
                    ' 　必ずイベント出欠の処理後にキャビネットイベントの処理を実行すること
                    _strTableName = "T_EVENT"
                    ' 対象テーブル更新チェック
                    UpFlg_T_EVENT = ChkUpdate(_sqlCommand)
                    If UpFlg_T_EVENT Then
                        ' SQLServerのT_EVENTテーブルから情報を取得する
                        GetT_EVENT(_sqlCommand, arrSQLsrv_T_EVENTRET)
                    End If

                    ' 連絡事項（キャビネット）
                    _strTableName = "T_INFOMATIONCABI"
                    ' 対象テーブル更新チェック
                    UpFlg_T_INFOMATIONCABI = ChkUpdate(_sqlCommand)
                    If UpFlg_T_INFOMATIONCABI Then
                        ' SQLServerのT_INFOMATIONテーブル(キャビネット用)から情報を取得する
                        GetT_INFOMATION_CABI(_sqlCommand)
                    End If

                    ' 2021/10/18 TOPデータ取得機能追加により、地区誌の取得処理を移動
                    '' 地区誌
                    '_strTableName = "T_MAGAZINE"
                    '' 対象テーブル更新チェック
                    'UpFlg_T_MAGAZINE = ChkUpdate(_sqlCommand)
                    'If UpFlg_T_MAGAZINE Then
                    '    ' SQLServerのT_MAGAZINEテーブルから情報を取得する
                    '    GetT_MAGAZINE(_sqlCommand)
                    'End If

                    ' 地区誌購入者
                    _strTableName = "T_MAGAZINEBUY"
                    ' 対象テーブル更新チェック
                    UpFlg_T_MAGAZINEBUY = ChkUpdate(_sqlCommand)
                    If UpFlg_T_MAGAZINEBUY Then
                        ' SQLServerのT_MAGAZINEBUYテーブルから情報を取得する
                        GetT_MAGAZINEBUY(_sqlCommand)
                    End If

                    ' SQLServerのM_DISTRICTOFFICERテーブルから情報を取得する ⇒ 2021/09/02 アプリには不要とし、コメントアウト。
                    'GetM_DISTRICTOFFICER(_sqlCommand)

                    ' 2021/10/18 TOPデータ取得機能追加により、キャビネット構成の取得処理を移動
                    '' キャビネット構成
                    '_strTableName = "M_CABINET"
                    '' 対象テーブル更新チェック
                    'UpFlg_M_CABINET = ChkUpdate(_sqlCommand)
                    'If UpFlg_M_CABINET Then
                    '    ' SQLServerの(M_CABINETテーブルから情報を取得する
                    '    GetM_CABINET(_sqlCommand)
                    'End If

                    ' クラブマスタ
                    _strTableName = "M_CLUB"
                    ' 対象テーブル更新チェック
                    UpFlg_M_CLUB = ChkUpdate(_sqlCommand)
                    If UpFlg_M_CLUB Then
                        ' SQLServerのM_CLUBテーブルから情報を取得する
                        GetM_CLUB(_sqlCommand)
                    End If


                Catch ex As Exception
                    ' エラーログ出力
                    context.Response.StatusCode = httpErrorCode
                    _logger.Error("●Read_SQLServer_HOME DB Error : ", ex)
                    Throw

                End Try

            End Using
            _connSvDB.Close()

        End Using

        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        ' SQL Serverより データを取得 (334BCLUB)
        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

        ' SQLServerコネクション生成(334BCLUB)
        Using _connSvDBCLUB As SqlConnection = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("LionsAplWeb.My.MySettings.DB334BCLUB").ConnectionString)
            _connSvDBCLUB.Open()

            'SQLServerコマンド生成
            Using _sqlCommand As SqlCommand = _connSvDBCLUB.CreateCommand()
                Try

                    ' クラブスローガン
                    _strTableName = "T_CLUBSLOGAN"
                    ' 対象テーブル更新チェック
                    UpFlg_T_CLUBSLOGAN = ChkUpdate(_sqlCommand)
                    If UpFlg_T_CLUBSLOGAN Then
                        ' SQLServerのT_CLUBSLOGANテーブルから情報を取得する
                        GetT_CLUBSLOGAN(_sqlCommand)
                    End If

                    ' 年間例会スケジュール
                    _strTableName = "T_MEETINGSCHEDULE"
                    ' 対象テーブル更新チェック
                    UpFlg_T_MEETINGSCHEDULE = ChkUpdate(_sqlCommand)
                    If UpFlg_T_MEETINGSCHEDULE Then
                        ' SQLServerのT_MEETINGSCHEDULEテーブルから情報を取得する
                        GetT_MEETINGSCHEDULE(_sqlCommand, arrSQLsrv_T_EVENTRET)
                    End If

                    ' 理事・委員会
                    _strTableName = "T_DIRECTOR"
                    ' 対象テーブル更新チェック
                    UpFlg_T_DIRECTOR = ChkUpdate(_sqlCommand)
                    If UpFlg_T_DIRECTOR Then
                        ' SQLServerのT_DIRECTORテーブルから情報を取得する
                        GetT_DIRECTOR(_sqlCommand)
                    End If

                    ' 年間例会プログラム
                    _strTableName = "T_MEETINGPROGRAM"
                    ' 対象テーブル更新チェック
                    UpFlg_T_MEETINGPROGRAM = ChkUpdate(_sqlCommand)
                    If UpFlg_T_MEETINGPROGRAM Then
                        ' SQLServerのT_MEETINGPROGRAMテーブルから情報を取得する
                        GetT_MEETINGPROGRAM(_sqlCommand)
                    End If

                    ' 連絡事項（クラブ）
                    _strTableName = "T_INFOMATION"
                    ' 対象テーブル更新チェック
                    UpFlg_T_INFOMATION = ChkUpdate(_sqlCommand)
                    If UpFlg_T_INFOMATION Then
                        ' SQLServerのT_INFOMATIONテーブル(クラブ用)から情報を取得する
                        GetT_INFOMATION_CLUB(_sqlCommand)
                    End If

                    ' 会員情報
                    _strTableName = "M_MEMBER"
                    ' 対象テーブル更新チェック
                    UpFlg_M_MEMBER = ChkUpdate(_sqlCommand)
                    If UpFlg_M_MEMBER Then
                        ' SQLServerのM_MEMBERテーブルから情報を取得する
                        GetM_MEMBER_HOME(_sqlCommand)
                    End If

                Catch ex As Exception
                    ' エラーログ出力
                    context.Response.StatusCode = httpErrorCode
                    _logger.Error("●Read_SQLServer_HOME DBCLUB Error : ", ex)
                    Throw

                End Try

            End Using
            _connSvDBCLUB.Close()

        End Using

        _logger.Info(" △ Read SQLServer HOME End △") 'ログ

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' 最終データNo取得
    ''' </summary>
    ''' <param name="_sqlCommand"></param>
    ''' <param name="dataClass"></param>
    ''' <returns></returns>
    Public Function Read_SQLServer_A_DATANO(ByRef _sqlCommand As SqlCommand, dataClass As String) As Integer

        Dim dataNo As Integer

        _logger.Info(" ▽ Read SQLServer A_DATANO Start ▽") 'ログ

        dataNo = 0
        Try
            GetA_DATANO(_sqlCommand, dataClass)

            ' A_DATANO検索結果リストから検索条件を作成する。
            For Each dicWrk As Dictionary(Of String, String) In arrSQLsrv_A_DATANO
                ' 指定DataClassの最終データNoを取得する
                If dicWrk("DataClass").Equals(dataClass) Then

                    ' 最終データNo.を取得する
                    dataNo = Integer.Parse(dicWrk("LastDataNo"))
                    _logger.Debug("○(A_DATANO検索)dataNo[" & dataNo & "]")
                    dataNo += 1
                    _logger.Debug("○(A_DATANO検索)dataNo+1[" & dataNo & "]")
                End If
            Next

        Catch ex As Exception
            ' エラーログ出力
            _context.Response.StatusCode = httpErrorCode
            _logger.Error("●(A_DATANO検索)例外エラー[SQL:" & _sqlCommand.CommandText & "]", ex)
            Throw
        End Try

        _logger.Info(" △ Read SQLServer A_DATANO End △") 'ログ

        Return dataNo

    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQL Serverよりデータを読み込む（TOP用）
    ''' </summary>
    ''' <param name="context">エラーステータス出力用</param>
    Public Sub Read_SQLServer_TOP(ByRef context As HttpContext, aplDate As DateTime)

        _logger.Info(" ▽ Read_SQLServer_TOP Start ▽") 'ログ

        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        ' SQL Serverより データを取得 (334B)
        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

        ' SQLServerコネクション生成(334B)
        Using _connSvDB As SqlConnection = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("LionsAplWeb.My.MySettings.DB334B").ConnectionString)

            _logger.Info("    ○SQLServer Open : " & _connSvDB.ConnectionString) 'ログ

            _connSvDB.Open()

            'SQLServerコマンド生成
            Using _sqlCommand As SqlCommand = _connSvDB.CreateCommand()

                Try
                    _logger.Info("    ○SQLServer Command Start") 'ログ

                    ' 設定ファイル
                    ' SQLServerのA_SETTINGテーブルから情報を取得する
                    GetA_SETTING(_sqlCommand)

                    ' 設定ファイルから年度情報を取得する
                    Init_TOP(aplDate)

                    ' 添付ファイルパス
                    ' SQLServerのA_FILEPATHテーブルから情報を取得する
                    GetA_FILEPATH(_sqlCommand)

                    ' 地区スローガン
                    ' SQLServerのT_SLOGANテーブルから情報を取得する
                    GetT_SLOGAN(_sqlCommand)

                    ' 地区誌
                    ' SQLServerのT_MAGAZINEテーブルから情報を取得する
                    GetT_MAGAZINE(_sqlCommand)

                    ' キャビネット構成
                    ' SQLServerの(M_CABINETテーブルから情報を取得する
                    GetM_CABINET(_sqlCommand)

                    ' エリアマスタ    2021/12 ADD
                    ' SQLServerの(M_AREAテーブルから情報を取得する
                    GetM_AREA(_sqlCommand)

                    ' 職種マスタ    2021/12 ADD
                    ' SQLServerの(M_JOBテーブルから情報を取得する
                    GetM_JOB(_sqlCommand)

                    ' マッチング    2021/12 ADD
                    ' SQLServerの(M_CABINETテーブルから情報を取得する
                    GetT_MATCHING(_sqlCommand)

                Catch ex As Exception
                    ' エラーログ出力
                    context.Response.StatusCode = httpErrorCode
                    _logger.Error("●Read_SQLServer_TOP DB Error : ", ex)
                    Throw

                End Try

                _logger.Info("    ○SQLServer Command End") 'ログ

            End Using
            _connSvDB.Close()

        End Using

        '' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        '' SQL Serverより データを取得 (334BCLUB)
        '' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

        '' SQLServerコネクション生成(334BCLUB)
        'Using _connSvDBCLUB As SqlConnection = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("LionsAplWeb.My.MySettings.DB334BCLUB").ConnectionString)
        '    _connSvDBCLUB.Open()

        '    'SQLServerコマンド生成
        '    Using _sqlCommand As SqlCommand = _connSvDBCLUB.CreateCommand()
        '        Try

        '        Catch ex As Exception
        '            ' エラーログ出力
        '            context.Response.StatusCode = httpErrorCode
        '            _logger.Error("●Read_SQLServer_TOP DBCLUB Error : ", ex)
        '            Throw

        '        End Try

        '    End Using
        '    _connSvDBCLUB.Close()

        'End Using

        _logger.Info(" △ Read_SQLServer_TOP End △") 'ログ

    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQL Serverより取得したA_SETTING情報より
    ''' 他のDBを検索するための条件（年度情報）を取得する（TOP用）
    ''' </summary>
    ''' <param name="aplDate">アプリの日時</param>
    Public Sub Init_TOP(aplDate As DateTime)
        Dim strDate As String                           ' 時刻文字列

        _logger.Info(" ▽ Start Init_TOP ▽") 'ログ

        ' アプリ時刻を取得
        _aplDate = aplDate

        ' 現在時刻の年月日を文字列で取得
        strDate = _aplDate.ToString("yyyy/MM/dd")

        ' SQLite情報(A_SETTING)から年度情報を設定する
        aplUtil.SetFiscalInfoSQLsvr(arrSQLsrv_A_SETTING)

        ' 年度設定
        ' 今年度
        _strNowFiscal = aplUtil.GetFiscalYear(strDate)
        ' 前年度
        _strPreFiscal = (Long.Parse(_strNowFiscal) - 1).ToString
        ' 前年度開始年月日
        _strPreFiscalStart = _strPreFiscal & "/" & aplUtil.PeriodStart

        _logger.Info(" △ End Init_TOP △") 'ログ
    End Sub



    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerを検索してデータを取得する
    ''' </summary>
    ''' <param name="Sqlreader"></param>
    ''' <returns>読取情報</returns>
    Public Function GetFldDicData(sqlreader As SqlDataReader) As Dictionary(Of String, String)
        Dim dicData As Dictionary(Of String, String)
        Dim strValue As String = String.Empty

        dicData = New Dictionary(Of String, String)

        'フィールド値格納
        For Each wFld As String In _TblMgr._lstFld
            '_logger.Info("wFld : " & wFld) ' ログ：項目
            ' 検索した情報をDBから取得する
            strValue = sqlreader(wFld).ToString()
            '_logger.Info("strValue : " & strValue) ' ログ：値
            ' 取得した情報を引数のDictionary型に書き込む
            dicData.Add(wFld, strValue)
            _logger.Info(wFld & ":" & strValue) 'ログ：「項目:値」

        Next

        Return dicData

    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' 指定テーブルの項目名を取得してSQLServerを検索する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    ''' <param name="arrSqlSrv"></param>
    Private Sub GetSQLServerData(ByRef sqlCommand As SqlCommand, arrSqlSrv As ArrayList)
        Dim intCnt As Integer = 0

        Try
            ' ユーティリティクラス生成
            _utl = New CmnUtility
            ' テーブルマネージャークラス生成
            _TblMgr = New TableManager(_logger, _utl, _strTableName)

            'SQLServerDBから取得
            Using sqlDataReader = sqlCommand.ExecuteReader()
                ' SQLServer検索結果配列クリア
                arrSqlSrv.Clear()

                ' 指定テーブルの項目名取得
                _TblMgr.Field_Name(_strTableName)

                While sqlDataReader.Read()
                    'SQLServerを検索してデータを取得する
                    arrSqlSrv.Add(GetFldDicData(sqlDataReader))
                    intCnt += 1
                End While

                sqlDataReader.Close()

            End Using

            _logger.Debug("○SQLserver " & _strTableName & " より データを取得 終了 " & intCnt.ToString & "件") ' ログ：データ取得完了

        Catch ex As Exception
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' 指定テーブルの項目名を取得してSQLServerを更新する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    ''' <param name="arrSqlSrv"></param>
    Private Sub UpdSQLServerData(ByRef sqlCommand As SqlCommand, arrSqlSrv As ArrayList)
        Dim intCnt As Integer = 0

        Try
            ' ユーティリティクラス生成
            _utl = New CmnUtility
            ' テーブルマネージャークラス生成
            _TblMgr = New TableManager(_logger, _utl, _strTableName)

            'SQLServerDBから取得
            Using sqlDataReader = sqlCommand.ExecuteReader()
                ' SQLServer検索結果配列クリア
                arrSqlSrv.Clear()

                ' 指定テーブルの項目名取得
                _TblMgr.Field_Name(_strTableName)

                While sqlDataReader.Read()
                    'SQLServerを検索してデータを取得する
                    arrSqlSrv.Add(GetFldDicData(sqlDataReader))
                    intCnt += 1
                End While

                sqlDataReader.Close()

            End Using

            _logger.Debug("○SQLserver " & _strTableName & " より データを取得 終了 " & intCnt.ToString & "件") ' ログ：データ取得完了

        Catch ex As Exception
            Throw
        End Try

    End Sub



    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' キャビネット用読込

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのA_SETTINGテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetA_SETTING(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "A_SETTING"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM A_SETTING")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_A_SETTING)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのA_DATANOテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetA_DATANO(ByRef sqlCommand As SqlCommand, dataClass As String)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "A_DATANO"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM A_DATANO WHERE DataClass = '" & dataClass & "'")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_A_DATANO)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのA_FILEPATHテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetA_FILEPATH(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "A_FILEPATH"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM A_FILEPATH")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_A_FILEPATH)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのT_SLOGANテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetT_SLOGAN(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "T_SLOGAN"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM T_SLOGAN WHERE " &
                             "DelFlg = '0' AND " &
                             "FiscalStart = '" & _strNowFiscal & "'")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_T_SLOGAN)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのT_LETTERテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetT_LETTER(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "T_LETTER"
            ' SQL文字列クリア
            sqlString.Remove()

            ' SQL文字列作成
            sqlString.Append("SELECT * FROM T_LETTER WHERE " &
                             "OpenFlg = '1' AND " &
                             "DelFlg = '0' AND " &
                             "EventDate >= '" & _strPreFiscalStart & "'")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_T_LETTER)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのT_EVENTRETテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetT_EVENTRET(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "T_EVENTRET"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM T_EVENTRET WHERE " &
                             "MemberCode = '" & aplUtil.MemberCode & "' AND " &
                             "DelFlg = '0' AND " &
                             "(TargetFlg <> '1' OR " &
                             "TargetFlg IS NULL) AND " &
                             "EventDate >= '" & _t_eventret_MonOfs & "'")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_T_EVENTRET)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのT_EVENTテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetT_EVENT(ByRef sqlCommand As SqlCommand,
                           ByRef arrSqlSrv As ArrayList)
        Dim sqlString As New clsSQLString

        Dim whereStr As String = String.Empty

        ' T_EVENTRET検索結果リストから検索条件を作成する。
        For Each dicWrk As Dictionary(Of String, String) In arrSqlSrv
            ' EventClass="1":キャビネットのみ対象
            If dicWrk("EventClass").Equals("1") Then
                ' 先頭以外は" OR "を追加する
                If whereStr.Length > 0 Then
                    whereStr += " OR "
                End If
                whereStr += "DataNo = '" & dicWrk("EventDataNo") & "'"
            End If
        Next

        If whereStr <> String.Empty Then
            ' 対象データがある場合のみ検索する
            Try
                ' フィールド名セット
                _strTableName = "T_EVENT"
                ' SQL文字列クリア
                sqlString.Remove()
                ' SQL文字列作成
                sqlString.Append("SELECT * FROM T_EVENT WHERE " &
                             "OpenFlg = '1' AND " &
                             "DelFlg = '0' AND " &
                             "(" & whereStr & ")")
                '"EventDate >= '" & _strPreFiscalStart & "'")
                ' SQLコマンド文字列追加
                sqlCommand.CommandText = sqlString.ToStr
                ' SQLServerからデータを取得
                GetSQLServerData(sqlCommand, arrSQLsrv_T_EVENT)

            Catch ex As Exception
                ' エラーログ出力
                _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
                Throw
            End Try
        End If

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのT_INFOMATIONテーブル(キャビネット用)から情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetT_INFOMATION_CABI(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "T_INFOMATION_CABI"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM T_INFOMATIONCABI WHERE " &
                             "AddDate >= '" & _t_infomation_Cabi_MonOfs & "' AND " &
                             "OpenFlg = '1' AND " &
                             "DelFlg = '0'")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_T_INFOMATION_CABI)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのT_MAGAZINEテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetT_MAGAZINE(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "T_MAGAZINE"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM T_MAGAZINE WHERE " &
                             "OpenFlg = '1' AND " &
                             "DelFlg = '0'")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_T_MAGAZINE)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのT_MAGAZINEBUYテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetT_MAGAZINEBUY(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "T_MAGAZINEBUY"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM T_MAGAZINEBUY WHERE " &
                             "MemberCode = '" & aplUtil.MemberCode & "'")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_T_MAGAZINEBUY)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのM_DISTRICTOFFICERテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetM_DISTRICTOFFICER(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "M_DISTRICTOFFICER"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM M_DISTRICTOFFICER WHERE " &
                             "DelFlg = '0'")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_M_DISTRICTOFFICER)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのM_CABINETテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetM_CABINET(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "M_CABINET"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM M_CABINET WHERE " &
                             "DelFlg = '0' AND " &
                             "FiscalStart = '" & _strNowFiscal & "'")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_M_CABINET)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのM_CLUBテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetM_CLUB(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "M_CLUB"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM M_CLUB WHERE " &
                             "DelFlg = '0'")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_M_CLUB)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのM_AREAテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    ''' <remarks>2021/12 ADD</remarks>
    Private Sub GetM_AREA(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "M_AREA"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM M_AREA")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_M_AREA)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのM_JOBテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    ''' <remarks>2021/12 ADD</remarks>
    Private Sub GetM_JOB(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "M_JOB"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM M_JOB")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_M_JOB)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのT_MATCHINGテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    ''' <remarks>2021/12 ADD</remarks>
    Private Sub GetT_MATCHING(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "T_MATCHING"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM T_MATCHING WHERE " &
                             "DelFlg = '0' AND " &
                             "OpenFlg = '1'")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_T_MATCHING)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' クラブ用読込

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのT_CLUBSLOGANテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetT_CLUBSLOGAN(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "T_CLUBSLOGAN"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM T_CLUBSLOGAN WHERE " &
                             "DelFlg = '0' AND " &
                             "ClubCode = '" & aplUtil.ClubCode & "' AND " &
                             "FiscalStart = '" & _strNowFiscal & "'")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_T_CLUBSLOGAN)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのT_MEETINGSCHEDULEテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetT_MEETINGSCHEDULE(ByRef sqlCommand As SqlCommand,
                                     ByRef arrSqlSrv As ArrayList)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "T_MEETINGSCHEDULE"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM T_MEETINGSCHEDULE WHERE " &
                            "ClubCode = '" & aplUtil.ClubCode & "' AND " &
                            "(Fiscal = '" & _strNowFiscal & "' OR " &
                            "MeetingDate >= '" & _t_director_MonOfs & "') AND " &
                            "DelFlg = '0'")
            '"DelFlg = '0' AND " &
            '"(" & whereStr & ")")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_T_MEETINGSCHEDULE)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try
        'End If

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのT_DIRECTORテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetT_DIRECTOR(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "T_DIRECTOR"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM T_DIRECTOR WHERE " &
                             "ClubCode = '" & aplUtil.ClubCode & "' AND " &
                             "OpenFlg = '1' AND " &
                             "DelFlg = '0' AND " &
                             "EventDate >= '" & _t_director_MonOfs & "'")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_T_DIRECTOR)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのT_MEETINGPROGRAMテーブルから情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetT_MEETINGPROGRAM(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "T_MEETINGPROGRAM"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM T_MEETINGPROGRAM WHERE " &
                             "ClubCode = '" & aplUtil.ClubCode & "' AND " &
                             "Fiscal = '" & _strNowFiscal & "' AND " &
                             "DelFlg = '0'")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_T_MEETINGPROGRAM)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのT_INFOMATIONテーブル(クラブ用)から情報を取得する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetT_INFOMATION_CLUB(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "T_INFOMATION_CLUB"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM T_INFOMATION WHERE " &
                             "AddDate >= '" & _t_infomation_Club_MonOfs & "' AND " &
                             "ClubCode = '" & aplUtil.ClubCode & "' AND " &
                             "OpenFlg = '1' AND " &
                             "DelFlg = '0'")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_T_INFOMATION_CLUB)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのM_MEMBERテーブルから情報を取得する（ACCOUNT用）
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetM_MEMBER(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "M_MEMBER"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM M_MEMBER WHERE LeaveFlg = '0'")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_M_MEMBER)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのM_MEMBERテーブルから情報を取得する(HOME用)
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub GetM_MEMBER_HOME(ByRef sqlCommand As SqlCommand)
        Dim sqlString As New clsSQLString

        Try
            ' フィールド名セット
            _strTableName = "M_MEMBER"
            ' SQL文字列クリア
            sqlString.Remove()
            ' SQL文字列作成
            sqlString.Append("SELECT * FROM M_MEMBER WHERE " &
                             "LeaveFlg = '0' AND " &
                             "ClubCode = '" & aplUtil.ClubCode & "'")
            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            GetSQLServerData(sqlCommand, arrSQLsrv_M_MEMBER)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " より データを取得)例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' 更新日を検索条件に指定テーブル（_strTableName）の件数を取得して
    ''' 件数がある、もしくは最終更新日がない場合は更新あり（True）を返す。
    ''' 最終更新日があり、件数がなければ更新なし（False）を返す。
    ''' </summary>
    ''' <param name="sqlCommand">SQLコマンド</param>
    ''' <returns></returns>
    Private Function ChkUpdate(ByRef sqlCommand As SqlCommand) As Boolean
        Dim ret As Boolean = True
        Dim count As Integer

        If aplUtil.LastUpdDate <> String.Empty Then

            If _strTableName.Equals("T_EVENT") Then
                ' T_EVENT用
                '  T_EVENTRETが更新されたらT_EVENTも更新する
                ret = UpFlg_T_EVENTRET
            Else
                ' 共通
                count = GetUpdCount(sqlCommand)
                If count = 0 Then
                    ret = False
                End If
            End If

        End If

        Return ret

    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' 更新日を検索条件に指定テーブル（_strTableName）の件数を取得する
    ''' </summary>
    ''' <param name="sqlCommand">SQLコマンド</param>
    ''' <returns></returns>
    Private Function GetUpdCount(ByRef sqlCommand As SqlCommand) As Integer
        Dim sqlString As New clsSQLString
        Dim count As Integer = 0

        Try
            ' SQL文字列クリア
            sqlString.Remove()
            If _strTableName.Equals("T_EVENTRET") Or
               _strTableName.Equals("T_MAGAZINEBUY") Then
                ' T_EVENTRET, T_MAGAZINEBUY用
                ' SQL文字列作成
                sqlString.Append("SELECT COUNT(*) FROM " & _strTableName & " WHERE " &
                                 "MemberCode = '" & aplUtil.MemberCode & "' AND " &
                                 "EditDate >= '" & aplUtil.LastUpdDate & "'")
            ElseIf _strTableName.Equals("M_CLUB") Or
                   _strTableName.Equals("T_CLUBSLOGAN") Or
                   _strTableName.Equals("T_MEETINGSCHEDULE") Or
                   _strTableName.Equals("T_DIRECTOR") Or
                   _strTableName.Equals("T_MEETINGPROGRAM") Or
                   _strTableName.Equals("T_INFOMATION") Or
                   _strTableName.Equals("M_MEMBER") Then
                ' M_CLUB, T_CLUBSLOGAN, T_MEETINGSCHEDULE, T_DIRECTOR, T_MEETINGPROGRAM,
                ' T_INFOMATION, M_MEMBER用
                ' SQL文字列作成
                sqlString.Append("SELECT COUNT(*) FROM " & _strTableName & " WHERE " &
                                 "ClubCode = '" & aplUtil.ClubCode & "' AND " &
                                 "EditDate >= '" & aplUtil.LastUpdDate & "'")
            Else
                ' 共通
                ' SQL文字列作成
                sqlString.Append("SELECT COUNT(*) FROM " & _strTableName & " WHERE " &
                                 "EditDate >= '" & aplUtil.LastUpdDate & "'")
            End If

            ' SQLコマンド文字列追加
            sqlCommand.CommandText = sqlString.ToStr
            ' SQLServerからデータを取得
            count = CType(sqlCommand.ExecuteScalar(), Integer)

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver " & _strTableName & " の件数確認 例外エラー[sql:" & sqlString.ToStr & "]", ex)
            Throw
        End Try

        Return count

    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    '''  更新フラグをクリア（更新あり）にする
    ''' </summary>
    Private Sub ClearUpdFlgAll()
        UpFlg_A_SETTING = True
        UpFlg_A_FILEPATH = True
        UpFlg_T_SLOGAN = True
        UpFlg_T_LETTER = True
        UpFlg_T_EVENTRET = True
        UpFlg_T_EVENT = True
        UpFlg_T_INFOMATIONCABI = True
        UpFlg_T_MAGAZINE = True
        UpFlg_T_MAGAZINEBUY = True
        UpFlg_M_CABINET = True
        UpFlg_M_CLUB = True
        UpFlg_M_AREA = True         '2021/12 ADD
        UpFlg_M_JOB = True          '2021/12 ADD
        UpFlg_T_MATCHING = True     '2021/12 ADD
        UpFlg_T_CLUBSLOGAN = True
        UpFlg_T_MEETINGSCHEDULE = True
        UpFlg_T_DIRECTOR = True
        UpFlg_T_MEETINGPROGRAM = True
        UpFlg_T_INFOMATION = True
        UpFlg_M_MEMBER = True
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DB書込み


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLServerのA_DATANOテーブルに情報を登録する
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Private Sub SetA_DATANO(ByRef sqlCommand As SqlCommand, dataNo As Integer, dataClass As String)
        Dim sqlString As New clsSQLString
        Dim intCnt As Integer = 0

        Try
            ' SQL文字列作成
            sqlCommand.CommandText = "UPDATE A_DATANO SET " +
                                     " LastDataNo = @LastDataNo " +
                                     "WHERE " +
                                     " DataClass = @DataClass"
            ' 最終データNo
            sqlCommand.Parameters.Add("@LastDataNo", SqlDbType.Int).Value = dataNo
            ' データ種類
            sqlCommand.Parameters.Add("@DataClass", SqlDbType.VarChar).Value = dataClass

            sqlCommand.ExecuteNonQuery()

        Catch ex As Exception
            ' エラーログ出力
            _logger.Error("●(SQLserver(A_DATANO更新)例外エラー[sql:" & sqlCommand.CommandText & "]", ex)
            Throw
        End Try

    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' 地区誌購入情報登録
    ''' </summary>
    ''' <param name="context"></param>
    Public Sub Write_SQLServer_MAGAZINE(ByRef context As HttpContext)
        'Dim dataNo As String
        Dim magazineDataNo As String
        Dim magazine As String
        Dim buyDate As String
        Dim buyNumber As String
        Dim magazinePrice As String
        Dim moneyTotal As String
        Dim region As String
        Dim zone As String
        Dim clubCode As String
        Dim clubNameShort As String
        Dim memberCode As String
        Dim memberName As String
        'Dim shippingDate As String
        'Dim paymentDate As String
        'Dim payment As String
        'Dim delFlg As String
        Dim editUser As String
        Dim editDate As String

        Dim dataNo As Integer

        magazineDataNo = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value3_01)
        magazine = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value3_02)
        buyDate = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value3_03)
        buyNumber = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value3_04)
        magazinePrice = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value3_05)
        moneyTotal = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value3_06)
        region = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value3_07)
        zone = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value3_08)
        clubCode = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value3_09)
        clubNameShort = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value3_10)
        memberCode = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value3_11)
        memberName = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value3_12)
        editUser = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value3_13)
        editDate = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value3_14)

        _logger.Info("01=" & magazineDataNo & Environment.NewLine &
                     "02=" & magazine & Environment.NewLine &
                     "03=" & buyDate & Environment.NewLine &
                     "04=" & buyNumber & Environment.NewLine &
                     "05=" & magazinePrice & Environment.NewLine &
                     "06=" & moneyTotal & Environment.NewLine &
                     "07=" & region & Environment.NewLine &
                     "08=" & zone & Environment.NewLine &
                     "09=" & clubCode & Environment.NewLine &
                     "10=" & clubNameShort & Environment.NewLine &
                     "11=" & memberCode & Environment.NewLine &
                     "12=" & memberName & Environment.NewLine &
                     "13=" & editUser & Environment.NewLine &
                     "14=" & editDate & Environment.NewLine)

        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        ' SQL Serverへ データを追加 (334B)
        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

        ' SQLServerコネクション生成(334B)
        Using _connSvDB As SqlConnection = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("LionsAplWeb.My.MySettings.DB334B").ConnectionString)
            _connSvDB.Open()

            'SQLServerコマンド生成
            Using _sqlCommand As SqlCommand = _connSvDB.CreateCommand()

                ' 地区誌購入者の最終データNo取得
                dataNo = Read_SQLServer_A_DATANO(_sqlCommand, ST_DATACLASS_MAGAZINEBUY)

                Try
                    _sqlCommand.CommandText = "INSERT INTO T_MAGAZINEBUY VALUES (" +
                                              " @DataNo, " +
                                              " @MagazineDataNo, " +
                                              " @Magazine, " +
                                              " @BuyDate, " +
                                              " @BuyNumber, " +
                                              " @MagazinePrice, " +
                                              " @MoneyTotal, " +
                                              " @Region, " +
                                              " @Zone, " +
                                              " @ClubCode, " +
                                              " @ClubNameShort, " +
                                              " @MemberCode, " +
                                              " @MemberName, " +
                                              " @ShippingDate, " +
                                              " @PaymentDate, " +
                                              " @Payment, " +
                                              " @DelFlg, " +
                                              " @EditUser, " +
                                              " @EditDate" +
                                              ")"
                    ' データNo.
                    _sqlCommand.Parameters.Add("@DataNo", SqlDbType.Int).Value = dataNo
                    ' 地区誌データNo.
                    If String.IsNullOrWhiteSpace(magazineDataNo) Then
                        _sqlCommand.Parameters.Add("@MagazineDataNo", SqlDbType.Int).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@MagazineDataNo", SqlDbType.Int).Value = magazineDataNo
                    End If
                    ' 地区誌名
                    If String.IsNullOrWhiteSpace(magazine) Then
                        _sqlCommand.Parameters.Add("@Magazine", SqlDbType.VarChar).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@Magazine", SqlDbType.VarChar).Value = magazine
                    End If
                    ' 購入日
                    If String.IsNullOrWhiteSpace(buyDate) Then
                        _sqlCommand.Parameters.Add("@BuyDate", SqlDbType.DateTime).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@BuyDate", SqlDbType.DateTime).Value = buyDate
                    End If
                    ' 冊子数
                    If String.IsNullOrWhiteSpace(buyNumber) Then
                        _sqlCommand.Parameters.Add("@BuyNumber", SqlDbType.Int).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@BuyNumber", SqlDbType.Int).Value = buyNumber
                    End If
                    ' 購入価格（税込）
                    If String.IsNullOrWhiteSpace(magazinePrice) Then
                        _sqlCommand.Parameters.Add("@MagazinePrice", SqlDbType.Int).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@MagazinePrice", SqlDbType.Int).Value = magazinePrice
                    End If
                    ' 購入金額
                    If String.IsNullOrWhiteSpace(moneyTotal) Then
                        _sqlCommand.Parameters.Add("@MoneyTotal", SqlDbType.Int).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@MoneyTotal", SqlDbType.Int).Value = moneyTotal
                    End If
                    ' リジョン
                    If String.IsNullOrWhiteSpace(region) Then
                        _sqlCommand.Parameters.Add("@Region", SqlDbType.VarChar).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@Region", SqlDbType.VarChar).Value = region
                    End If
                    ' ゾーン
                    If String.IsNullOrWhiteSpace(zone) Then
                        _sqlCommand.Parameters.Add("@Zone", SqlDbType.VarChar).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@Zone", SqlDbType.VarChar).Value = zone
                    End If
                    ' クラブコード
                    If String.IsNullOrWhiteSpace(clubCode) Then
                        _sqlCommand.Parameters.Add("@ClubCode", SqlDbType.VarChar).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@ClubCode", SqlDbType.VarChar).Value = clubCode
                    End If
                    ' クラブ名（略称）
                    If String.IsNullOrWhiteSpace(clubNameShort) Then
                        _sqlCommand.Parameters.Add("@ClubNameShort", SqlDbType.VarChar).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@ClubNameShort", SqlDbType.VarChar).Value = clubNameShort
                    End If
                    ' 会員番号
                    If String.IsNullOrWhiteSpace(memberCode) Then
                        _sqlCommand.Parameters.Add("@MemberCode", SqlDbType.VarChar).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@MemberCode", SqlDbType.VarChar).Value = memberCode
                    End If
                    ' 会員名
                    If String.IsNullOrWhiteSpace(memberName) Then
                        _sqlCommand.Parameters.Add("@MemberName", SqlDbType.VarChar).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@MemberName", SqlDbType.VarChar).Value = memberName
                    End If
                    ' 発送日
                    _sqlCommand.Parameters.Add("@ShippingDate", SqlDbType.DateTime).Value = DBNull.Value
                    ' 入金日
                    _sqlCommand.Parameters.Add("@PaymentDate", SqlDbType.DateTime).Value = DBNull.Value
                    ' 入金額
                    _sqlCommand.Parameters.Add("@Payment", SqlDbType.Int).Value = DBNull.Value
                    ' 購入キャンセル
                    _sqlCommand.Parameters.Add("@DelFlg", SqlDbType.Char).Value = DBNull.Value
                    ' 更新者
                    If String.IsNullOrWhiteSpace(editUser) Then
                        _sqlCommand.Parameters.Add("@EditUser", SqlDbType.VarChar).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@EditUser", SqlDbType.VarChar).Value = editUser
                    End If
                    ' 更新日
                    If String.IsNullOrWhiteSpace(editDate) Then
                        _sqlCommand.Parameters.Add("@EditDate", SqlDbType.DateTime).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@EditDate", SqlDbType.DateTime).Value = editDate
                    End If

                    _sqlCommand.ExecuteNonQuery()

                Catch ex As Exception

                    ' エラーログ出力
                    context.Response.StatusCode = httpErrorCode
                    _logger.Error("●SQLserver(T_MAGAZINEBUY更新)例外エラー[SQL:" & _sqlCommand.CommandText & "]", ex)
                    Throw

                End Try

                ' 最終データNo登録
                SetA_DATANO(_sqlCommand, dataNo, ST_DATACLASS_MAGAZINEBUY)

            End Using
            _connSvDB.Close()

        End Using

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' イベント出欠用情報登録
    ''' </summary>
    ''' <param name="context"></param>
    Public Sub Write_SQLServer_EVENTRET(ByRef context As HttpContext)
        'Dim sqlCommand As New clsSQLString

        Dim dataNo As String
        Dim answer As String
        Dim answerLate As String
        Dim answerEarly As String
        Dim online As String
        Dim option1 As String
        Dim option2 As String
        Dim option3 As String
        Dim option4 As String
        Dim option5 As String
        Dim otherCount As String

        dataNo = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value4_01)
        answer = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value4_02)
        answerLate = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value4_03)
        answerEarly = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value4_04)
        online = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value4_05)
        option1 = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value4_06)
        option2 = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value4_07)
        option3 = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value4_08)
        option4 = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value4_09)
        option5 = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value4_10)
        otherCount = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value4_11)

        _logger.Info("01=" & dataNo & Environment.NewLine &
                     "02=" & answer & Environment.NewLine &
                     "03=" & answerLate & Environment.NewLine &
                     "04=" & answerEarly & Environment.NewLine &
                     "05=" & online & Environment.NewLine &
                     "06=" & option1 & Environment.NewLine &
                     "07=" & option2 & Environment.NewLine &
                     "08=" & option3 & Environment.NewLine &
                     "09=" & option4 & Environment.NewLine &
                     "10=" & option5 & Environment.NewLine &
                     "11=" & otherCount & Environment.NewLine)

        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        ' SQL Serverの データを更新 (334B)
        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

        ' SQLServerコネクション生成(334B)
        Using _connSvDB As SqlConnection = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("LionsAplWeb.My.MySettings.DB334B").ConnectionString)
            _connSvDB.Open()

            'SQLServerコマンド生成
            Using _sqlCommand As SqlCommand = _connSvDB.CreateCommand()

                Try
                    _sqlCommand.CommandText = "UPDATE T_EVENTRET SET " +
                                              " Answer = @Answer, " +
                                              " AnswerLate = @AnswerLate, " +
                                              " AnswerEarly = @AnswerEarly, " +
                                              " Online = @Online, " +
                                              " Option1 = @Option1, " +
                                              " Option2 = @Option2, " +
                                              " Option3 = @Option3, " +
                                              " Option4 = @Option4, " +
                                              " Option5 = @Option5, " +
                                              " OtherCount = @OtherCount " +
                                              "WHERE " +
                                              " DataNo = @DataNo"
                    ' 出欠
                    If String.IsNullOrWhiteSpace(answer) Then
                        _sqlCommand.Parameters.Add("@Answer", SqlDbType.Char).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@Answer", SqlDbType.Char).Value = answer
                    End If
                    ' 出席（遅刻）
                    If String.IsNullOrWhiteSpace(answerLate) Then
                        _sqlCommand.Parameters.Add("@AnswerLate", SqlDbType.Char).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@AnswerLate", SqlDbType.Char).Value = answerLate
                    End If
                    ' 出席（早退）
                    If String.IsNullOrWhiteSpace(answerEarly) Then
                        _sqlCommand.Parameters.Add("@AnswerEarly", SqlDbType.Char).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@AnswerEarly", SqlDbType.Char).Value = answerEarly
                    End If
                    ' オンライン
                    If String.IsNullOrWhiteSpace(online) Then
                        _sqlCommand.Parameters.Add("@Online", SqlDbType.Char).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@Online", SqlDbType.Char).Value = online
                    End If
                    ' オプション1
                    If String.IsNullOrWhiteSpace(option1) Then
                        _sqlCommand.Parameters.Add("@Option1", SqlDbType.Char).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@Option1", SqlDbType.Char).Value = option1
                    End If
                    ' オプション2
                    If String.IsNullOrWhiteSpace(option2) Then
                        _sqlCommand.Parameters.Add("@Option2", SqlDbType.Char).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@Option2", SqlDbType.Char).Value = option2
                    End If
                    ' オプション3
                    If String.IsNullOrWhiteSpace(option3) Then
                        _sqlCommand.Parameters.Add("@Option3", SqlDbType.Char).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@Option3", SqlDbType.Char).Value = option3
                    End If
                    ' オプション4
                    If String.IsNullOrWhiteSpace(option4) Then
                        _sqlCommand.Parameters.Add("@Option4", SqlDbType.Char).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@Option4", SqlDbType.Char).Value = option4
                    End If
                    ' オプション5
                    If String.IsNullOrWhiteSpace(option5) Then
                        _sqlCommand.Parameters.Add("@Option5", SqlDbType.Char).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@Option5", SqlDbType.Char).Value = option5
                    End If
                    ' 本人以外の参加人数
                    If String.IsNullOrWhiteSpace(otherCount) Then
                        _sqlCommand.Parameters.Add("@OtherCount", SqlDbType.Int).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@OtherCount", SqlDbType.Int).Value = otherCount
                    End If

                    _sqlCommand.Parameters.Add("@DataNo", SqlDbType.Int).Value = dataNo

                    _sqlCommand.ExecuteNonQuery()

                Catch ex As Exception

                    ' エラーログ出力
                    context.Response.StatusCode = httpErrorCode
                    _logger.Error("●SQLserver(T_EVENTRET更新)例外エラー[SQL:" & _sqlCommand.CommandText & "]", ex)
                    Throw

                End Try

            End Using
            _connSvDB.Close()

        End Using

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' ACCOUNTREG：アカウント設定情報の登録
    ''' </summary>
    ''' <param name="context"></param>
    Public Sub Write_SQLServer_ACCOUNTREG(ByRef context As HttpContext)
        'Dim sqlCommand As New clsSQLString

        Dim accountDate As String
        Dim region As String
        Dim zone As String
        Dim clubCode As String
        Dim memberCode As String

        accountDate = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value6_01)
        region = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value6_02)
        zone = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value6_03)
        clubCode = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value6_04)
        memberCode = context.Request.Form(LionsAplWeb.My.Settings.RequestItem1_Value6_05)

        _logger.Info("01=" & accountDate & Environment.NewLine)

        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        ' SQL Serverの データを更新 (334BCLUB)
        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

        ' SQLServerコネクション生成(334BCLUB)
        Using _connSvDB As SqlConnection = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("LionsAplWeb.My.MySettings.DB334BCLUB").ConnectionString)
            _connSvDB.Open()

            'SQLServerコマンド生成
            Using _sqlCommand As SqlCommand = _connSvDB.CreateCommand()

                Try
                    _sqlCommand.CommandText = "UPDATE M_MEMBER SET " +
                                              " AccountDate = @AccountDate " +
                                              "WHERE " +
                                              " Region = @Region AND " +
                                              " Zone = @Zone AND " +
                                              " ClubCode = @ClubCode AND " +
                                              " MemberCode = @MemberCode"
                    ' アカウント設定日
                    If String.IsNullOrWhiteSpace(accountDate) Then
                        _sqlCommand.Parameters.Add("@AccountDate", SqlDbType.Char).Value = DBNull.Value
                    Else
                        _sqlCommand.Parameters.Add("@AccountDate", SqlDbType.Char).Value = accountDate
                    End If

                    ' リジョン
                    _sqlCommand.Parameters.Add("@Region", SqlDbType.Char).Value = region

                    ' ゾーン
                    _sqlCommand.Parameters.Add("@Zone", SqlDbType.Char).Value = zone

                    ' クラブコード
                    _sqlCommand.Parameters.Add("@ClubCode", SqlDbType.Char).Value = clubCode

                    ' 会員コード
                    _sqlCommand.Parameters.Add("@MemberCode", SqlDbType.Char).Value = memberCode

                    _sqlCommand.ExecuteNonQuery()

                Catch ex As Exception

                    ' エラーログ出力
                    context.Response.StatusCode = httpErrorCode
                    _logger.Error("●SQLserver(M_MEMBER更新)例外エラー[SQL:" & _sqlCommand.CommandText & "]", ex)
                    Throw

                End Try

            End Using
            _connSvDB.Close()

        End Using

    End Sub

End Class
