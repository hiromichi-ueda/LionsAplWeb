Option Explicit On
Option Strict On

Imports System.IO
Imports System.Web
Imports System.Data.SQLite
Imports log4net

Public Class SQLiteManager

    ' エラー発生時に返すHTTPコード
    Private Const httpErrorCode As Integer = 500

    Protected _Command_SQLite As SQLiteCommand                                      ' SQLiteコマンドクラス

    Protected _logger As ILog                                                       ' ロガー
    Protected _utl As CmnUtility = Nothing                                          ' ユーティリティ
    Protected _tblMgr As TableManager = Nothing                                     ' テーブルマネージャー

    Protected _appPath As String = AppDomain.CurrentDomain.BaseDirectory.TrimEnd    ' アプリケーションパス
    Protected _strWorkDbPath As String = _appPath + My.Settings.WorkDbPath          ' SQLiteDBファイル保管先パス
    Protected _strWorkDbPost As String = String.Empty                               ' SQLiteDBファイル名（Post時）
    Protected _strWorkDbName As String = String.Empty                               ' SQLiteDBファイル名
    Protected _strTableName As String = String.Empty                                ' 操作対象テーブル名

    Public arrSQLite As New ArrayList                                               ' SQLite情報(取得及び登録データ)
    Public arrSQLite_A_ACCOUNT As New ArrayList                                     ' SQLite情報(A_ACCOUNT)
    Public arrSQLite_A_SETTING As New ArrayList                                     ' SQLite情報(A_SETTING)

    Public aplUtil As AplUtility                                                    ' アプリ用ユーティリティクラス

    Protected _lstKey As New List(Of String)                                        ' Key 情報
    'Protected _lstFld As New List(Of String)                                        'フィールド名
    Protected _lstKeyFld As New List(Of String)                                     'キーフィールド名

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="logger"></param>
    Public Sub New(ByRef logger As ILog)
        _logger = logger
        aplUtil = New AplUtility(logger)

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' HttpContextからSqliteファイルを取得する
    ''' </summary>
    ''' <param name="context"></param>
    ''' <param name="dbPath"></param>
    ''' <param name="dbName"></param>
    Public Sub GetFile_SQLite_fromContent(ByRef context As HttpContext, ByRef dbPath As String, ByRef dbName As String)
        _logger.Info(" ▽ GetFile_SQLite_fromContent ▽") 'ログ
        _strWorkDbPath = dbPath
        _strWorkDbName = dbName

        ' Sqliteファイル取得
        Dim postedFile As HttpPostedFile = context.Request.Files("Sqlite")
        Dim aryData(postedFile.ContentLength) As Byte
        postedFile.InputStream.Read(aryData, 0, postedFile.ContentLength)

        ' WorkDbFileの名称セット
        _strWorkDbPost = postedFile.FileName
        _logger.Info("   ○パス名:" & _strWorkDbPath) 'ログ
        _logger.Info("   ○変更前：SQLiteファイル名:" & _strWorkDbPost) 'ログ
        _strWorkDbName = aplUtil.ChkFileName(_strWorkDbPath, _strWorkDbPost)
        _logger.Info("   ○変更後：SQLiteファイル名:" & _strWorkDbName) 'ログ
        ' ファイル名を引数で返す
        dbName = _strWorkDbName

        ' POSTされてきたSQLiteのDBをローカル環境に書き出す
        Try
            Using fs As New FileStream(_strWorkDbPath & _strWorkDbName, FileMode.Create)
                Dim bw As New BinaryWriter(fs)
                bw.Write(aryData)
                bw.Close()
            End Using
        Catch ex As Exception
            _logger.Error("●Sqlite file error : " & ex.Message) 'ログ
            Throw
        End Try

        _logger.Info(" △ GetFile_SQLite_fromContent End △") 'ログ
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLiteファイル作成
    ''' </summary>
    Public Sub Create_SQLite()

        ' Webアプリケーションパス
        Dim appPath As String = AppDomain.CurrentDomain.BaseDirectory.TrimEnd
        ' Web.config よりWorkDbPath値を取得する
        Dim strWorkDbPath As String = appPath + My.Settings.WorkDbPath
        Dim strWorkDbName As String = My.Settings.SQLiteFileName + My.Settings.SQLiteFileEx



    End Sub



    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' SQLiteファイル読込み処理


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLiteファイルからデータを読み込む
    ''' </summary>
    ''' <param name="dbPath"></param>
    ''' <param name="dbName"></param>
    ''' <param name="context"></param>
    Public Sub Read_SQLite(ByRef dbPath As String, ByRef dbName As String, ByRef context As HttpContext)
        Dim sqlString As New clsSQLString
        'Dim dicData As Dictionary(Of String, String)
        Dim intCnt As Integer = 0
        Dim strLogTitle As String = String.Empty
        Dim strLogDetail As String = String.Empty
        Dim strValue As String = String.Empty

        _logger.Info(" ▽ Read SQLite Start ▽") 'ログ
        _strWorkDbPath = dbPath
        _strWorkDbName = dbName
        _logger.Info("   ○SQList File=" & _strWorkDbPath & _strWorkDbName)  ' ログ SQLiteファイルパス

        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        ' SQLiteより データを取得
        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

        ' SQLiteコネクション生成
        Using _conn As New SQLiteConnection()
            ' ローカルに書き出した SQLiteのDBを開く
            _conn.ConnectionString = "Data Source=" & _strWorkDbPath & _strWorkDbName
            _conn.Open()

            'SQLiteコマンド生成
            Using Command_SQLite As SQLiteCommand = _conn.CreateCommand()

                Try
                    GetA_ACCOUNT(Command_SQLite)

                Catch ex As Exception
                    ' エラーログ出力
                    context.Response.StatusCode = httpErrorCode
                    _logger.Error("●Read_SQLite Error : ", ex)
                    Throw

                End Try

            End Using
            _conn.Close()
        End Using

        _logger.Info(" △ Read SQLite End △") 'ログ

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLiteファイルからデータを読み込む(HOME用)
    ''' </summary>
    ''' <param name="dbPath"></param>
    ''' <param name="dbName"></param>
    ''' <param name="context"></param>
    Public Sub Read_SQLite_HOME(ByRef dbPath As String, ByRef dbName As String, ByRef context As HttpContext)

        _logger.Info(" ▽ Read SQLite HOME Start ▽") 'ログ
        _strWorkDbPath = dbPath
        _strWorkDbName = dbName
        _logger.Debug("   SQList File=" & _strWorkDbPath & _strWorkDbName)  ' ログ SQLiteファイルパス

        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        ' SQLiteより データを取得
        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

        ' SQLiteコネクション生成
        Using _conn As New SQLiteConnection()
            ' ローカルに書き出した SQLiteのDBを開く
            _conn.ConnectionString = "Data Source=" & _strWorkDbPath & _strWorkDbName
            _conn.Open()

            'SQLiteコマンド生成
            Using Command_SQLite As SQLiteCommand = _conn.CreateCommand()

                Try
                    ' SQLiteファイルからA_ACCOUNTテーブル情報を取得する
                    GetA_ACCOUNT(Command_SQLite)
                    ' SQLiteファイルからA_SETTINGテーブル情報を取得する
                    GetA_SETTING(Command_SQLite)

                Catch ex As Exception
                    ' エラーログ出力
                    context.Response.StatusCode = httpErrorCode
                    _logger.Error("●Read_SQLite Error : ", ex)
                    Throw

                End Try

            End Using
            _conn.Close()
        End Using

        _logger.Info(" △ Read SQLite HOME End △") 'ログ
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLiteを検索してデータを取得する
    ''' </summary>
    ''' <param name="InSQLitereader"></param>
    ''' <returns></returns>
    Public Function GetFldDicData(InSQLitereader As SQLiteDataReader) As Dictionary(Of String, String)
        Dim dicData As Dictionary(Of String, String)
        Dim strValue As String = String.Empty

        dicData = New Dictionary(Of String, String)

        'フィールド値格納
        For Each wFld As String In _tblMgr._lstFld
            ' 取得値
            strValue = _utl.NullToEmpty(InSQLitereader(wFld))
            ' 項目別 文字数チェック＆切り捨て
            strValue = _tblMgr.CutKeta(wFld, strValue)
            dicData.Add(wFld, strValue)
            _logger.Debug(wFld & ":" & strValue) 'ログ 項目:値

        Next

        Return dicData

    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLiteファイルから指定テーブルの情報を取得する。
    ''' </summary>
    ''' <param name="InCommand_SQLite"></param>
    ''' <param name="InArrSqlite"></param>
    Private Sub GetSQLiteData(InCommand_SQLite As SQLiteCommand,
                              InArrSqlite As ArrayList)
        Dim intCnt As Integer = 0

        Try
            _logger.Debug(" ▽ SQLiteより データを取得 開始 ▽") ' ログ：データ取得完了

            ' ユーティリティクラス生成
            _utl = New CmnUtility
            ' テーブルマネージャークラス生成
            _tblMgr = New TableManager(_logger, _utl, _strTableName)

            'SQLiteDBから取得
            Using sqlDataReader = InCommand_SQLite.ExecuteReader()
                ' SQLite検索結果配列クリア
                InArrSqlite.Clear()

                ' 指定テーブルの項目名取得
                _tblMgr.Field_Name(_strTableName)

                While sqlDataReader.Read()
                    'SQLiteを検索してデータを取得する
                    InArrSqlite.Add(GetFldDicData(sqlDataReader))
                    intCnt += 1
                End While

                sqlDataReader.Close()

            End Using

            _logger.Debug(" △ SQLiteより データを取得 終了 " & intCnt.ToString & "件 △") ' ログ：データ取得完了

        Catch ex As Exception
            Throw
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLiteファイルからA_ACCOUNTテーブル情報を取得する
    ''' </summary>
    ''' <param name="InCommand_SQLite"></param>
    Private Sub GetA_ACCOUNT(InCommand_SQLite As SQLiteCommand)
        Dim sqlString As New clsSQLString
        Dim intCnt As Integer = 0

        Try
            'フィールド名セット
            _strTableName = "A_ACCOUNT"
            sqlString.Remove()
            sqlString.Append("SELECT * FROM A_ACCOUNT")
            InCommand_SQLite.CommandText = sqlString.ToStr

            ' SQLiteファイルから指定テーブルの情報を取得する。
            GetSQLiteData(InCommand_SQLite, arrSQLite_A_ACCOUNT)

        Catch ex As Exception
            _logger.Error("●(SQLiteより データを取得)例外エラー[SQL:" & sqlString.ToStr & "]", ex)
            Throw

        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLiteファイルからA_SETTINGテーブル情報を取得する
    ''' </summary>
    ''' <param name="InCommand_SQLite"></param>
    Private Sub GetA_SETTING(InCommand_SQLite As SQLiteCommand)
        Dim sqlString As New clsSQLString
        Dim intCnt As Integer = 0

        Try
            'フィールド名セット
            _strTableName = "A_SETTING"
            sqlString.Remove()
            sqlString.Append("SELECT * FROM A_SETTING")
            InCommand_SQLite.CommandText = sqlString.ToStr

            ' SQLiteファイルから指定テーブルの情報を取得する。
            GetSQLiteData(InCommand_SQLite, arrSQLite_A_SETTING)

        Catch ex As Exception
            _logger.Error("●(SQLiteより データを取得)例外エラー[SQL:" & sqlString.ToStr & "]", ex)
            Throw

        End Try

    End Sub



    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' SQLiteファイル書込み処理


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' SQLiteファイルにデータを書き込む
    ''' </summary>
    ''' <param name="dbPath"></param>
    ''' <param name="dbName"></param>
    ''' <param name="arrSqlSrv"></param>
    ''' <param name="context"></param>
    Private Sub WriteSQLiteData(ByRef dbPath As String,
                            ByRef dbName As String,
                            ByRef arrSqlSrv As ArrayList,
                            ByRef context As HttpContext)

        Dim sqlCommand As New clsSQLString
        Dim intCnt As Integer = 0
        Dim intCnt2 As Integer = 0
        Dim strLogTitle As String = String.Empty
        Dim strLogDetail As String = String.Empty
        Dim strValue As String = String.Empty

        _logger.Info(" ▽ WriteSQLiteData Start ▽")   ' ログSQLite書込み開始
        _strWorkDbPath = dbPath
        _strWorkDbName = dbName

        _logger.Info("    ○SQLite:" & dbPath & dbName)   ' ログSQLiteパス&ファイル名
        _logger.Info("    ○Table:" & _strTableName)      ' ログSQLiteテーブル名

        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        ' SQLiteより データを取得
        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

        '＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋
        ' SQLiteコネクション生成
        Using _conn As New SQLiteConnection()

            ' ローカルに書き出した SQLiteのDBを開く
            _conn.ConnectionString = "Data Source=" & _strWorkDbPath & _strWorkDbName
            _conn.Open()

            'SQLiteコマンド生成
            Using Command_SQLite As SQLiteCommand = _conn.CreateCommand()

                ' ユーティリティクラス生成
                _utl = New CmnUtility
                ' テーブルマネージャークラス生成
                _tblMgr = New TableManager(_logger, _utl, _strTableName)

                ' 指定テーブルの項目名取得
                _tblMgr.Field_Name(_strTableName)

                intCnt = 0
                Try

                    ' 対象テーブル削除
                    Try
                        sqlCommand.Remove()
                        sqlCommand.Append("DELETE FROM " & _strTableName & " ")
                        _logger.Debug("   delete：" & sqlCommand.ToStr)  ' ログ：削除
                        Command_SQLite.CommandText = sqlCommand.ToStr
                        Command_SQLite.ExecuteNonQuery()
                    Catch ex As Exception
                        ' エラーログ出力
                        context.Response.StatusCode = httpErrorCode
                        _logger.Error("●(SQLite 削除)例外エラー[SQL:" & sqlCommand.ToStr & "]", ex)
                        Throw
                    End Try

                    ' 対象テーブルカウンタリセット
                    Try
                        sqlCommand.Remove()
                        sqlCommand.Append("DELETE FROM sqlite_sequence WHERE name='" & _strTableName & "'")
                        _logger.Debug("   table counter reset sqlCommand：" & sqlCommand.ToStr)  ' ログ：削除
                        Command_SQLite.CommandText = sqlCommand.ToStr
                        Command_SQLite.ExecuteNonQuery()
                    Catch ex As Exception
                        ' エラーログ出力
                        context.Response.StatusCode = httpErrorCode
                        _logger.Error("●(SQLite カウンタクリア)例外エラー[SQL:" & sqlCommand.ToStr & "]", ex)
                        Throw
                    End Try

                    _logger.Debug("   arrSqlSrv : 件数=" & arrSqlSrv.Count.ToString)   ' ログSQLite書込み開始

                    intCnt = 0
                    For Each dicWrk As Dictionary(Of String, String) In arrSqlSrv
                        _logger.Debug("dicWrk : 項目数=" & dicWrk.Count)   ' ログSQLite書込み開始

                        '-------------
                        '↓SQL作成
                        sqlCommand.Remove()
                        'sqlCommand.Append("INSERT OR REPLACE INTO " & _strTableName & " (")
                        sqlCommand.Append("INSERT INTO " & _strTableName & " (")

                        'フィールド名セット
                        sqlCommand.Append(_tblMgr.GetFldString(_tblMgr.FldStringType.NameOnly, True))

                        sqlCommand.Append(") VALUES (")

                        'フィールド値 SQLセット
                        sqlCommand.Append(_tblMgr.GetFldValueString(dicWrk, _tblMgr.FldValueStringType.SQLite1, True))

                        sqlCommand.Append(")")

                        _logger.Debug("Insert Table:" & _strTableName) ' ログ：Insert Table name
                        _logger.Debug(" sqlCommand：" & sqlCommand.ToStr) ' ログ：SQLite登録用SQL

                        '↑SQL作成
                        '-------------

                        Command_SQLite.CommandText = sqlCommand.ToStr
                        Command_SQLite.ExecuteNonQuery()

                        intCnt += 1
                    Next

                Catch ex As Exception

                    ' エラーログ出力
                    context.Response.StatusCode = httpErrorCode
                    _logger.Error("●(SQLite Marge)例外エラー[SQL:" & sqlCommand.ToStr & "]", ex)
                    Throw

                End Try

            End Using
            _conn.Close()
        End Using

        _logger.Info(" △ WriteSQLiteData End：" & intCnt.ToString & "件 △") ' ログ：SQLite書込み終了＆件数

    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' ACCOUNT用SQLiteファイル書込み処理
    ''' </summary>
    ''' <param name="dbPath"></param>
    ''' <param name="dbName"></param>
    ''' <param name="sqlsvrManager"></param>
    ''' <param name="context"></param>
    Public Sub Write_SQLite_ACCOUNT(ByRef dbPath As String,
                            ByRef dbName As String,
                            ByRef sqlsvrManager As SQLServerManager,
                            ByRef context As HttpContext)

        ' 2021/10/18 TOPデータ取得機能追加により、キャビネット構成の取得処理を移動
        '' A_SETTING情報書込み
        '' フィールド名セット
        '_strTableName = "A_SETTING"
        '' SQLiteファイルへの書込み
        'WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_A_SETTING, context)

        ' M_MEMBER情報書込み
        ' フィールド名セット
        _strTableName = "M_MEMBER"
        ' SQLiteファイルへの書込み
        WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_M_MEMBER, context)

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' HOME用SQLiteファイル書込み処理
    ''' </summary>
    ''' <param name="dbPath"></param>
    ''' <param name="dbName"></param>
    ''' <param name="sqlsvrManager"></param>
    ''' <param name="context"></param>
    Public Sub Write_SQLite_HOME(ByRef dbPath As String,
                            ByRef dbName As String,
                            ByRef sqlsvrManager As SQLServerManager,
                            ByRef context As HttpContext)

        ' 2021/10/18 TOPデータ取得機能追加により、キャビネット構成の取得処理を移動
        '' 添付パスファイル
        '' A_FILEPATH情報書込み
        '' フィールド名セット
        '_strTableName = "A_FILEPATH"
        '' SQLiteファイルへの書込み
        'WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_A_FILEPATH, context)

        ' 2021/10/18 TOPデータ取得機能追加により、キャビネット構成の取得処理を移動
        '' 地区スローガン
        'If sqlsvrManager.UpFlg_T_SLOGAN Then
        '    ' T_SLOGAN情報書込み
        '    ' フィールド名セット
        '    _strTableName = "T_SLOGAN"
        '    ' SQLiteファイルへの書込み
        '    WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_T_SLOGAN, context)
        'End If

        ' キャビネットレター
        If sqlsvrManager.UpFlg_T_LETTER Then
            ' T_LETTER情報書込み
            ' フィールド名セット
            _strTableName = "T_LETTER"
            ' SQLiteファイルへの書込み
            WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_T_LETTER, context)
        End If

        ' イベント出欠
        If sqlsvrManager.UpFlg_T_EVENTRET Then
            ' T_EVENTRET情報書込み
            ' フィールド名セット
            _strTableName = "T_EVENTRET"
            ' SQLiteファイルへの書込み
            WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_T_EVENTRET, context)
        End If

        ' キャビネットイベント
        If sqlsvrManager.UpFlg_T_EVENT Then
            ' T_EVENT情報書込み
            ' フィールド名セット
            _strTableName = "T_EVENT"
            ' SQLiteファイルへの書込み
            WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_T_EVENT, context)
        End If

        ' 連絡事項（キャビネット）
        If sqlsvrManager.UpFlg_T_INFOMATIONCABI Then
            ' T_INFOMATION_CABI情報書込み
            ' フィールド名セット
            _strTableName = "T_INFOMATION_CABI"
            ' SQLiteファイルへの書込み
            WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_T_INFOMATION_CABI, context)
        End If

        ' 2021/10/18 TOPデータ取得機能追加により、キャビネット構成の取得処理を移動
        '' 地区誌
        'If sqlsvrManager.UpFlg_T_MAGAZINE Then
        '    ' T_MAGAZINE情報書込み
        '    ' フィールド名セット
        '    _strTableName = "T_MAGAZINE"
        '    ' SQLiteファイルへの書込み
        '    WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_T_MAGAZINE, context)
        'End If

        ' 地区誌購入者
        If sqlsvrManager.UpFlg_T_MAGAZINEBUY Then
            ' T_MAGAZINEBUY情報書込み
            ' フィールド名セット
            _strTableName = "T_MAGAZINEBUY"
            ' SQLiteファイルへの書込み
            WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_T_MAGAZINEBUY, context)
        End If

        ' M_DISTRICTOFFICER情報書込み ⇒ 2021/09/02 アプリには不要とし、コメントアウト。
        ' フィールド名セット
        ' _strTableName = "M_DISTRICTOFFICER"
        ' SQLiteファイルへの書込み
        ' WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_M_DISTRICTOFFICER, context)

        ' 2021/10/18 TOPデータ取得機能追加により、キャビネット構成の取得処理を移動
        '' キャビネット構成
        'If sqlsvrManager.UpFlg_M_CABINET Then
        '    ' M_CABINET情報書込み
        '    ' フィールド名セット
        '    _strTableName = "M_CABINET"
        '    ' SQLiteファイルへの書込み
        '    WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_M_CABINET, context)
        'End If

        ' クラブマスタ
        If sqlsvrManager.UpFlg_M_CLUB Then
            ' M_CLUB情報書込み
            ' フィールド名セット
            _strTableName = "M_CLUB"
            ' SQLiteファイルへの書込み
            WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_M_CLUB, context)
        End If

        ' クラブスローガン
        If sqlsvrManager.UpFlg_T_CLUBSLOGAN Then
            ' T_CLUBSLOGAN情報書込み
            ' フィールド名セット
            _strTableName = "T_CLUBSLOGAN"
            ' SQLiteファイルへの書込み
            WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_T_CLUBSLOGAN, context)
        End If

        ' 年間例会スケジュール
        If sqlsvrManager.UpFlg_T_MEETINGSCHEDULE Then
            ' T_MEETINGSCHEDULE情報書込み
            ' フィールド名セット
            _strTableName = "T_MEETINGSCHEDULE"
            ' SQLiteファイルへの書込み
            WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_T_MEETINGSCHEDULE, context)
        End If

        ' 理事・委員会
        If sqlsvrManager.UpFlg_T_DIRECTOR Then
            ' T_DIRECTOR情報書込み
            ' フィールド名セット
            _strTableName = "T_DIRECTOR"
            ' SQLiteファイルへの書込み
            WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_T_DIRECTOR, context)
        End If

        ' 年間例会プログラム
        If sqlsvrManager.UpFlg_T_MEETINGPROGRAM Then
            ' T_MEETINGPROGRAM情報書込み
            ' フィールド名セット
            _strTableName = "T_MEETINGPROGRAM"
            ' SQLiteファイルへの書込み
            WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_T_MEETINGPROGRAM, context)
        End If

        ' 連絡事項（クラブ）
        If sqlsvrManager.UpFlg_T_INFOMATION Then
            ' M_CLUB情報書込み
            ' フィールド名セット
            _strTableName = "T_INFOMATION_CLUB"
            ' SQLiteファイルへの書込み
            WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_T_INFOMATION_CLUB, context)
        End If

        ' 会員情報
        If sqlsvrManager.UpFlg_M_MEMBER Then
            ' M_MEMBER情報書込み
            ' フィールド名セット
            _strTableName = "M_MEMBER"
            ' SQLiteファイルへの書込み
            WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_M_MEMBER, context)
        End If

        ' A_ACCOUNTテーブル情報の最終更新日項目を更新する。
        Upd_T_ACCOUNT(dbPath, dbName, context)

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' T_ACCOUNTの最終更新日更新処理
    ''' </summary>
    ''' <param name="dbPath"></param>
    ''' <param name="dbName"></param>
    ''' <param name="context"></param>
    Public Sub Upd_T_ACCOUNT(ByRef dbPath As String,
                             ByRef dbName As String,
                             ByRef context As HttpContext)
        ' A_ACCOUNTテーブル情報の最終更新日項目を更新する。
        Dim nowDt As DateTime
        Dim stDt As String
        Dim sqlCommand As New clsSQLString

        nowDt = DateTime.Now
        stDt = nowDt.ToString("yyyy/MM/dd HH:mm:ss")

        _logger.Info(" ▽ Start Update SQLite A_ACCOUNT ▽") 'ログ
        _strWorkDbPath = dbPath
        _strWorkDbName = dbName

        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        ' SQLiteの データを更新
        ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

        '＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋
        ' SQLiteコネクション生成
        Using _conn As New SQLiteConnection()

            ' ローカルに書き出した SQLiteのDBを開く
            _conn.ConnectionString = "Data Source=" & _strWorkDbPath & _strWorkDbName
            _conn.Open()

            'SQLiteコマンド生成
            Using Command_SQLite As SQLiteCommand = _conn.CreateCommand()
                ' ユーティリティクラス生成
                '_utl = New CmnUtility

                Try

                    _logger.Info("○SQLiteへデータを更新 開始 A_ACCOUNT.LastUpdDate:" & stDt) ' ログSQLite書込み開始

                    '-------------
                    '↓SQL作成
                    sqlCommand.Remove()
                    sqlCommand.Append("UPDATE A_ACCOUNT SET LastUpdDate = '" & stDt & "'")

                    _logger.Debug("sqlCommand：" & sqlCommand.ToStr) ' ログ：SQLServer登録様SQL

                    '↑SQL作成
                    '-------------

                    Command_SQLite.CommandText = sqlCommand.ToStr
                    Command_SQLite.ExecuteNonQuery()

                Catch ex As Exception

                    ' エラーログ出力
                    context.Response.StatusCode = httpErrorCode
                    _logger.Error("●(SQLite 更新)例外エラー[SQL:" & sqlCommand.ToStr & "]", ex)
                    Throw

                End Try

            End Using
            _conn.Close()

            _logger.Info(" △ End Update SQLite A_ACCOUNT △") 'ログ

        End Using
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' TOP用SQLiteファイル書込み処理
    ''' </summary>
    ''' <param name="dbPath"></param>
    ''' <param name="dbName"></param>
    ''' <param name="sqlsvrManager"></param>
    ''' <param name="context"></param>
    Public Sub Write_SQLite_TOP(ByRef dbPath As String,
                            ByRef dbName As String,
                            ByRef sqlsvrManager As SQLServerManager,
                            ByRef context As HttpContext)

        ' A_SETTING情報書込み
        ' フィールド名セット
        _strTableName = "A_SETTING"
        ' SQLiteファイルへの書込み
        WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_A_SETTING, context)

        ' 添付パスファイル
        ' A_FILEPATH情報書込み
        _strTableName = "A_FILEPATH"
        ' SQLiteファイルへの書込み
        WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_A_FILEPATH, context)

        ' 地区スローガン
        ' T_SLOGAN情報書込み
        _strTableName = "T_SLOGAN"
        ' SQLiteファイルへの書込み
        WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_T_SLOGAN, context)

        ' 地区誌
        ' T_MAGAZINE情報書込み
        _strTableName = "T_MAGAZINE"
        ' SQLiteファイルへの書込み
        WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_T_MAGAZINE, context)

        ' キャビネット構成
        ' M_CABINET情報書込み
        _strTableName = "M_CABINET"
        ' SQLiteファイルへの書込み
        WriteSQLiteData(dbPath, dbName, sqlsvrManager.arrSQLsrv_M_CABINET, context)

    End Sub

End Class
