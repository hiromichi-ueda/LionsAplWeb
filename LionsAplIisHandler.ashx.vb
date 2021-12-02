Option Explicit On
Option Strict On
Imports System.IO
Imports System.Data.SqlClient
Imports System.Data.SQLite
Imports log4net

Public Class LionsAplIisHandler
    Implements System.Web.IHttpHandler

    ' log4net ロガー
    Private logger As ILog = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    ' エラー発生時に返すHTTPコード
    Private Const httpErrorCode As Integer = 500

    Protected _arrSQLite As New ArrayList ' SQLite情報
    Protected _arrSQLsrv As New ArrayList ' SQLsrv情報
    Protected _lstKey As New List(Of String) ' Key 情報
    Protected _lstFld As New List(Of String) 'フィールド名
    Protected _lstKeyFld As New List(Of String) 'キーフィールド名
    Protected _Utl As CmnUtility = Nothing ' ユーティリティ
    Protected _TableName As String = String.Empty
    Protected _AplDate As DateTime ' アプリケーション側の時間
    Protected _ResultCount As Integer() ' 更新件数
    Protected _ResultCountName As String() '更新処理 

    Protected _sqliteManager As SQLiteManager   ' SQLiteマネージャー
    Protected _sqlsvrManager As SQLServerManager   ' SQLServerマネージャー

    Protected _region As String = String.Empty
    Protected _zone As String = String.Empty
    Protected _clubCode As String = String.Empty
    Protected _memberCode As String = String.Empty

    ''' <summary>
    ''' リクエスト処理
    ''' </summary>
    ''' <param name="context"></param>
    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest

        Dim m_samplexmlString As String
        'リクエストの種類（TOP：ユーザ情報に影響しない情報の取得
        '                  ACCOUNT：アカウント設定用情報の取得
        '                  ACCOUNTREG：アカウント設定情報の登録
        ' 　　　　　　　　 HOME： 設定アカウントのアプリ表示用情報の取得
        ' 　　　　　　　　 MAGAZINE：地区誌購入者用情報の登録
        ' 　　　　　　　　 EVENTRET：イベント出欠用情報の登録）
        Dim strDbType As String = String.Empty

        'Dim lstResultLog As New List(Of String)

        ' Webアプリケーションパス
        Dim appPath As String = AppDomain.CurrentDomain.BaseDirectory.TrimEnd

        ' [Web.config]よりWorkDbPath値を取得する
        Dim strWorkDbPath As String = appPath + My.Settings.WorkDbPath
        ' 
        Dim strWorkDbName As String = String.Empty

        ' ログのパスを追加
        Dim strFileInfo = New System.IO.FileInfo(appPath + My.Settings.LogConfigFileName)
        Dim logFileCollection = Config.XmlConfigurator.Configure(strFileInfo)

        ' SQLiteマネージャー生成
        _sqliteManager = New SQLiteManager(logger)

        Try
            logger.Info("■>>>>>>>>>> LionsAplWeb Handler開始 >>>>>>>>>> " & Now.ToString("yyyy/MM/dd HH:mm:ss"))

            Using reader As StreamReader = New StreamReader(context.Request.InputStream)
                m_samplexmlString = reader.ReadToEnd()
            End Using

            ' HttpContextからリクエストの種類を取得する
            strDbType = context.Request.Form(My.Settings.RequestItem1)
            logger.Info("○リクエスト種類:" & strDbType) 'ログ DB種類

            ' HttpContextからアプリケーション側の時間を取得する
            _AplDate = DateTime.Parse(context.Request.Form(My.Settings.RequestItem2))
            logger.Info("○アプリ時間:" & _AplDate) 'ログ DB種類

            ' DBファイル保管先フォルダがない場合
            If Not System.IO.Directory.Exists(strWorkDbPath) Then
                ' エラーログ出力
                context.Response.StatusCode = httpErrorCode
                InfoErrLog("●(引数)作業用DB格納フォルダが存在しません。")
                Exit Try
            End If

            Dim strErrInfo(5) As String
            strErrInfo(0) = "-postedFile.ContentLength="
            strErrInfo(1) = "-aryData.Count="
            strErrInfo(2) = "-strWorkDbPath="
            strErrInfo(3) = "-strWorkDbName="

            'TOP：ユーザ情報に影響しない情報の取得
            'ACCOUNT：アカウント設定用情報の取得
            'HOME： 設定アカウントのアプリ表示用情報の取得
            'の場合
            If (strDbType.Equals(My.Settings.RequestItem1_Value1) Or
                strDbType.Equals(My.Settings.RequestItem1_Value2) Or
                strDbType.Equals(My.Settings.RequestItem1_Value5)) Then
                ' アプリ表示用情報(SQLiteファイル)を取得する

                Try
                    ' HttpContextからSqliteファイルを取得してローカルに保管する。
                    _sqliteManager.GetFile_SQLite_fromContent(context, strWorkDbPath, strWorkDbName)

                Catch ex As Exception
                    ' エラーログ出力
                    context.Response.StatusCode = httpErrorCode
                    logger.Error("●(ローカルDB作成)SystemError", ex)
                    Throw
                End Try

            End If

            ' SQLServerマネージャー生成
            _sqlsvrManager = New SQLServerManager(logger, context)

            'TOP：ユーザ情報に影響しない情報の取得
            'ACCOUNT：アカウント設定用情報の取得
            'ACCOUNTREG：アカウント設定情報の登録
            'HOME： 設定アカウントのアプリ表示用情報の取得
            'MAGAZINE：地区誌購入者用情報の登録
            'EVENTRET：イベント出欠用情報の登録
            'の場合
            If (strDbType.Equals(My.Settings.RequestItem1_Value1)) Then
                'ACCOUNT：アカウント設定用情報の取得

                ' SQLServerよりデータを読み込む
                Try
                    _sqlsvrManager.Read_SQLServer_ACCOUNT(context)
                Catch ex As Exception
                    logger.Error("●(SQLServer Read ACCOUNT)SystemError", ex)
                    Throw
                End Try

                'SQLiteファイルにデータを書き込む 
                Try
                    _sqliteManager.Write_SQLite_ACCOUNT(strWorkDbPath, strWorkDbName, _sqlsvrManager, context)
                Catch ex As Exception
                    logger.Error("●(SQLite Write ACCOUNT)SystemError", ex)
                    Throw
                End Try

            ElseIf (strDbType.Equals(My.Settings.RequestItem1_Value2)) Then
                'HOME： 設定アカウントのアプリ表示用情報の取得

                Try
                    ' SQLiteファイルからデータを読み込む(HOME用)
                    _sqliteManager.Read_SQLite_HOME(strWorkDbPath, strWorkDbName, context)
                Catch ex As Exception
                    logger.Error("●(SQLite Read HOME)SystemError", ex)
                    Throw
                End Try

                ' SQLServerマネージャーのHOME用初期化
                _sqlsvrManager.Init_HOME(_sqliteManager.arrSQLite_A_ACCOUNT,
                                         _sqliteManager.arrSQLite_A_SETTING,
                                         _AplDate)

                ' SQLServerよりデータを読み込む
                Try
                    _sqlsvrManager.Read_SQLServer_HOME(context)
                Catch ex As Exception
                    logger.Error("●(SQLServer Read HOME)SystemError", ex)
                    Throw
                End Try

                'SQLiteファイルにデータを書き込む 
                Try
                    _sqliteManager.Write_SQLite_HOME(strWorkDbPath, strWorkDbName, _sqlsvrManager, context)
                Catch ex As Exception
                    logger.Error("●(SQLite Write HOME)SystemError", ex)
                    Throw
                End Try

            ElseIf (strDbType.Equals(My.Settings.RequestItem1_Value3)) Then
                'MAGAZINE：地区誌購入者用情報の登録
                ' SQLServerへデータを書き込む
                Try
                    _sqlsvrManager.Write_SQLServer_MAGAZINE(context)
                Catch ex As Exception
                    logger.Error("●(SQLServer Write MAGAZINE)SystemError", ex)
                    Throw
                End Try

            ElseIf (strDbType.Equals(My.Settings.RequestItem1_Value4)) Then
                'EVENTRET：イベント出欠用情報の登録

                ' SQLServerへデータを書き込む
                Try
                    _sqlsvrManager.Write_SQLServer_EVENTRET(context)
                Catch ex As Exception
                    logger.Error("●(SQLServer Write EVENTRET)SystemError", ex)
                    Throw
                End Try

            ElseIf (strDbType.Equals(My.Settings.RequestItem1_Value5)) Then
                'TOP：ユーザ情報に影響しない情報の取得

                ' SQLServerよりデータを読み込む
                Try
                    _sqlsvrManager.Read_SQLServer_TOP(context, _AplDate)
                Catch ex As Exception
                    logger.Error("●(SQLServer Read TOP)SystemError", ex)
                    Throw
                End Try

                'SQLiteファイルにデータを書き込む 
                Try
                    _sqliteManager.Write_SQLite_TOP(strWorkDbPath, strWorkDbName, _sqlsvrManager, context)
                Catch ex As Exception
                    logger.Error("●(SQLite Write TOP)SystemError", ex)
                    Throw
                End Try

            ElseIf (strDbType.Equals(My.Settings.RequestItem1_Value6)) Then
                'ACCOUNTREG：アカウント設定情報の登録

                ' SQLServerへデータを書き込む
                Try
                    _sqlsvrManager.Write_SQLServer_ACCOUNTREG(context)
                Catch ex As Exception
                    logger.Error("●(SQLServer Write ACCOUNTREG)SystemError", ex)
                    Throw
                End Try

            End If

            'TOP：ユーザ情報に影響しない情報の取得
            'ACCOUNT：アカウント設定用情報の取得
            'HOME： 設定アカウントのアプリ表示用情報の取得
            'の場合
            If (strDbType.Equals(My.Settings.RequestItem1_Value1) Or
                strDbType.Equals(My.Settings.RequestItem1_Value2) Or
                strDbType.Equals(My.Settings.RequestItem1_Value5)) Then

                ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
                ' ResponseへWorkDbをセット
                ' *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
                context.Response.WriteFile(strWorkDbPath & strWorkDbName)
                logger.Info("○context.Response.WriteFile:" & strWorkDbPath & strWorkDbName)
                context.Response.Flush()


            End If

            '結果の出力
            'For Each wString As String In lstResultLog
            '   logger.Info(String.Format("{0}", wString))
            'Next

        Catch ex As Exception
            InfoErrLog("●<LionsAplWeb Handler> 例外エラー ", ex)
        Finally
            If System.IO.File.Exists(strWorkDbPath & strWorkDbName) Then
                System.IO.File.Delete(strWorkDbPath & strWorkDbName)
                logger.Info("○ファイル削除処理実行:" & strWorkDbPath & strWorkDbName)
            End If
            ' テスト確認用
            If System.IO.File.Exists(strWorkDbPath & strWorkDbName) Then
                    logger.Info("○ファイル未削除:" & strWorkDbPath & strWorkDbName)
                Else
                    logger.Info("○ファイル削除済:" & strWorkDbPath & strWorkDbName)
            End If
            logger.Info("■<<<<<<<<<< LionsAplWeb Handler終了 <<<<<<<<<< " & Now.ToString("yyyy/MM/dd HH:mm:ss"))
        End Try


    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

    ''' <summary>
    ''' エラーログ出力(Info＆Error)
    ''' </summary>
    ''' <param name="strMsg">出力メッセージ</param>
    Private Sub InfoErrLog(strMsg As String)
        logger.Info(strMsg)
        logger.Error(strMsg)
    End Sub

    ''' <summary>
    ''' エラーログ出力(Info＆Error) exあり
    ''' </summary>
    ''' <param name="strMsg">出力メッセージ</param>
    Private Sub InfoErrLog(strMsg As String, ex As Exception)
        logger.Info(strMsg, ex)
        logger.Error(strMsg, ex)

        'ｽﾀｯｸﾄﾚｰｽ
        'logger.Info(" StackTrace: " & Environment.StackTrace)
        'Dim StackTrace As System.Diagnostics.StackTrace = New System.Diagnostics.StackTrace(ex, True)
        'Dim lineNumber As Integer = StackTrace.GetFrame(Index).GetFileLineNumber()

    End Sub

    ''' <summary>
    ''' フィールド文字列(項目名/値)の作成
    ''' </summary>
    ''' <param name="dicData">値</param>
    ''' <param name="wType">作成タイプ FldValueStringType</param>
    ''' <param name="bolNewLine">改行</param>
    ''' <returns></returns>
    Protected Function GetFldValueString(ByRef dicData As Dictionary(Of String, String),
                                         ByVal wType As FldValueStringType, ByVal bolNewLine As Boolean) As String
        Dim strRet As String = String.Empty
        Dim strNewLine As String = String.Empty
        Dim bolFirst As Boolean

        If bolNewLine Then
            strNewLine = System.Environment.NewLine
        End If

        Select Case wType
            Case FldValueStringType.ValueOnly
                'フィールド値のみ Log用
                bolFirst = True
                For Each wFld As String In _lstFld
                    strRet &= IIf(bolFirst, " ", strNewLine & ",").ToString
                    strRet &= dicData(wFld)
                    bolFirst = False
                Next
            Case FldValueStringType.KeyValueOnly
                'キーフィールド値のみ
                bolFirst = True
                For Each wFld As String In _lstKeyFld
                    strRet &= IIf(bolFirst, " ", strNewLine & ",").ToString
                    strRet &= dicData(wFld)
                    bolFirst = False
                Next
            Case FldValueStringType.SQLsrv1
                '" 'xxx' AS ShainName "
                bolFirst = True
                For Each wFld As String In _lstFld
                    strRet &= IIf(bolFirst, " ", strNewLine & ",").ToString
                    bolFirst = False
                Next

            Case FldValueStringType.SQLsrv2 'T_INPUT系用
                '",'xxx' "
                bolFirst = True
                For Each wFld As String In _lstFld
                    strRet &= IIf(bolFirst, " ", strNewLine & ",").ToString
                    bolFirst = False
                Next

            Case FldValueStringType.SQL_Where
                '" AND KinmuDate ={0}"

                bolFirst = True
                For Each wFld As String In _lstKeyFld
                    strRet &= IIf(bolFirst, " ", strNewLine & " AND ").ToString
                    ''strRet &= String.Format("{0} = {1}", wFld, FieldSQL(wFld, dicData(wFld)))
                    bolFirst = False
                Next

            Case FldValueStringType.SQLite1
                '" 'xxx' "
                bolFirst = True
                For Each wFld As String In _lstFld
                    strRet &= IIf(bolFirst, " ", strNewLine & ",").ToString
                    ''strRet &= FieldSQL_Lite(wFld, dicData(wFld))
                    bolFirst = False
                Next

        End Select

        Return strRet

    End Function


    ''' <summary>
    ''' 項目別 値の桁チェック＆カット
    ''' </summary>
    ''' <param name="strFldName">フィールド名</param>
    ''' <param name="strFldValue">値</param>
    ''' <returns>加工後 値</returns>
    Protected Overridable Function CutKeta(ByVal strFldName As String, ByVal strFldValue As String) As String

        Dim strRet As String = strFldValue
        Dim eType As CmnUtility.eValueType = CmnUtility.eValueType.String '型
        Dim intKeta As Integer = -1 '桁

        'フィールド別に型・桁を取得
        Field_Difinition(strFldName, eType, intKeta)

        '桁オーバーチェック
        If _Utl.GetStringByte(strRet) > intKeta Then
            InfoErrLog(String.Format("{0} {1}='{2}' 桁({3})オーバー", _TableName, strFldName, strRet, intKeta.ToString))

            'カット
            strRet = _Utl.CutString(strRet, intKeta)
        End If

        Return strRet

    End Function

    ''' <summary>
    ''' テーブル項目の定義 (型と桁)
    ''' </summary>
    ''' <param name="strFldName"></param>
    ''' <param name="eType">型(戻り値)</param>
    ''' <param name="intKeta">桁 (戻り値)</param>
    ''' <remarks>項目チェックに使用</remarks>
    Protected Sub Field_Difinition(ByVal strFldName As String,
                                   ByRef eType As CmnUtility.eValueType,
                                   ByRef intKeta As Integer)

        intKeta = -1

        'フィールド別に型を決め、SQL文を作成
        Select Case _TableName
            Case "M_MEMBER"

                Select Case strFldName
                    Case "Region" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "Zone" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "ClubCode" : eType = CmnUtility.eValueType.String : intKeta = 6
                    Case "ClubNameShort" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "MemberCode" : eType = CmnUtility.eValueType.String : intKeta = 7
                    Case "MemberFirstName" : eType = CmnUtility.eValueType.String : intKeta = 30
                    Case "MemberLastName" : eType = CmnUtility.eValueType.String : intKeta = 30
                    Case "MemberNameKana" : eType = CmnUtility.eValueType.String : intKeta = 40
                End Select

        End Select

    End Sub

    ''' <summary>
    ''' フィールド文字列(項目名/値) GetFldValueString 引数用
    ''' </summary>
    Protected Enum FldValueStringType
        ValueOnly '値のみ
        KeyValueOnly 'キー値のみ
        SQLsrv1   'SQL srvマージ用
        SQLsrv2   'SQL srvインサート用
        SQL_Where 'Where句
        SQLite1   'SQLiteマージ用
    End Enum

End Class