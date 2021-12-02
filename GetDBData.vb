'23456789+123456789+123456789+123456789+123456789+123456789+123456789+123456789+123456789+123456789+
'---------------------------------------------------------------------------------------------------
'
' GetDBData
'
' Created by INSAT on 2021/07/06.
' Copyright (c) 2021年 INSAT. All right reserved.
'
'---------------------------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports System.Data.SQLite
Imports log4net
Public Class GetDBData

    Protected Const httpErrorCode As Integer = 500 ' エラー発生時に返すHTTPコード
    Protected _ProcessID As String = "Cls00xxxx"
    Protected _ProcessName As String = "xx処理"
    Protected _ProcessSubName As String = String.Empty
    Protected _ErrMsg As String = String.Empty

    Protected _Command_SQLite As SQLiteCommand
    Protected _SqlCommand As SqlCommand
    Protected _HttpContext As HttpContext
    Protected _log As ILog


    Public Property TimeOutTime As String = "20"
    Public Property ConnectionInfo334B As String = "Data Source=APSERVER;" &
                                                   "Initial Catalog=334B;" &
                                                   "Connect Timeout=" & TimeOutTime & ";" &
                                                   "User ID=sa;" &
                                                   "Password=P@ssw0rd;"
    Public Property ConnectionInfo334BCLUB As String = "Data Source=APSERVER;" &
                                                       "Connect Timeout=" & TimeOutTime & ";" &
                                                       "User ID=sa;" &
                                                       "Password=P@ssw0rd;"

    Public Property Result As String = ""

    'JSONList設定ファイルクラス
    'Public Property DbSetting As DBSetting

    'JSONスローガンクラス（ホームTOP画面）
    'Public Property DbSlogan As DBSlogan

    'JSONListキャビネットレタークラス（ホームTOP画面）
    'Public Property DbTopLetter As DBTopLetter

    'JSONList参加予定一覧クラス（ホームTOP画面）
    'Public Property DbTopEvent As DBTopEvent

    'JSONListキャビネットレタークラス
    'Public Property DbLetter As DBLetter

    'JSONListイベント一覧クラス
    'Public Property DbEventList As DBEventList

    'TEST
    'Public Property DbClub As Club

    ' 会員マスタDBクラス(アカウント画面)
    'Public Property DbMember As DBMember

    ' アカウント設定データ（from 会員マスタDB）クラス（アカウント設定画面）
    'Public Property DbAccountSet As AccountSetting

    'Public Property WUtil As WebUtil

    '-----------------------------------------------------------------------------------------------
    ' コンストラクタ
    '-----------------------------------------------------------------------------------------------
    Public Sub New(ByRef logger As ILog)
        _log = logger
        _log.Info("GetJsonFromDB コンストラクタ")

        'If My.Settings.DEVEOP_ENV Then
        '    ConnectionInfo334B = My.Settings.DE334B
        '    ConnectionInfo334BCLUB = My.Settings.DE334BCLUB
        'Else
        '    ConnectionInfo334B = My.Settings.OE334B
        '    ConnectionInfo334BCLUB = My.Settings.OE334BCLUB
        'End If

        _log.Info("DB 334B : " + ConnectionInfo334B)
        _log.Info("DB 334BCLUB : " + ConnectionInfo334BCLUB)

    End Sub

    '-----------------------------------------------------------------------------------------------
    ' スタート画面
    '-----------------------------------------------------------------------------------------------
    ' 設定ファイルマスタDBクラス
    '-----------------------------------------------------------------------------------------------

    ' 設定ファイルマスタ検索＆結果取得
    'Sub GetDBSetting(ByVal DistrictCode As String)

    '    Dim Sql As String

    '    DbSetting = New DBSetting With {.Data = New List(Of DataSetting)}

    '    Sql = "SELECT " &
    '            "DistrictCode, " &
    '            "DistrictName, " &
    '            "CabinetName, " &
    '            "PeriodStart, " &
    '            "PeriodEnd, " &
    '            "DistrictID, " &
    '            "MagazineMoney, " &
    '            "EventDataDay, " &
    '            "CabinetEmailAddeess, " &
    '            "AdminEmailAddress, " &
    '            "CabinetTelNo, " &
    '            "AdminTelNo " &
    '          "FROM A_SETTING"
    '    Sql = Sql & " WHERE"
    '    If (DistrictCode.Length > 0) Then
    '        Sql = Sql & " DistrictCode='" & DistrictCode & "';"
    '    End If
    '    Try
    '        Using Conn As New SqlConnection
    '            Conn.ConnectionString = (ConnectionInfo334B)

    '            Conn.Open()
    '            Using cmd As New SqlCommand(Sql)
    '                cmd.Connection = Conn
    '                cmd.CommandType = CommandType.Text
    '                Using reader As SqlDataReader = cmd.ExecuteReader()
    '                    While (reader.Read())
    '                        Dim RowData = New DataSetting
    '                        Dim idx As Integer = 0
    '                        RowData.DistrictCode = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.DistrictName = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.CabinetName = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.PeriodStart = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.PeriodEnd = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.DistrictID = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.EventDataDay = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.CabinetEmailAddress = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.AdminEmailAddress = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.CabinetEmailAddress = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.CabinetTelNo = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.AdminTelNo = reader.GetValue(idx).ToString : idx += 1

    '                        DbSetting.Data.Add(RowData)
    '                    End While
    '                    Dim jsonObj As String = JsonConvert.SerializeObject(DbSetting)
    '                    Result += jsonObj
    '                End Using
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        Result = ex.Message
    '    End Try
    'End Sub

    '-----------------------------------------------------------------------------------------------
    ' アカウント設定画面
    '-----------------------------------------------------------------------------------------------
    ' 会員マスタ検索＆結果取得
    '-----------------------------------------------------------------------------------------------
    'Sub GetDBAccountSetting(ByVal Region As String,
    '                ByVal Zone As String,
    '                ByVal ClubCode As String,
    '                ByVal MemberCode As String)
    '    Dim Sql As String

    '    DbAccountSet = New AccountSetting With {.Data = New List(Of DataAccount)}

    '    ' SQL作成
    '    Sql = "SELECT " &
    '            "Region, " &
    '            "Zone, " &
    '            "ClubCode, " &
    '            "ClubNameShort, " &
    '            "MemberCode, " &
    '            "MemberFirstName, " &
    '            "MemberLastName, " &
    '            "MemberNameKana " &
    '          "FROM M_MEMBER"
    '    Sql = Sql & " WHERE"
    '    If (Region.Length > 0) Then
    '        Sql = Sql & " Region='" & Region & "'"
    '    End If
    '    If (Zone.Length > 0) Then
    '        Sql = Sql & " Zone='" & Zone & "'"
    '    End If
    '    If (ClubCode.Length > 0) Then
    '        Sql = Sql & " ClubCode='" & ClubCode & "'"
    '    End If
    '    If (MemberCode.Length > 0) Then
    '        Sql = Sql & " MemberCode='" & MemberCode & "'"
    '    End If
    '    Sql = Sql & " LeaveFlg='0';"

    '    Try
    '        Using Conn As New SqlConnection
    '            Conn.ConnectionString = (ConnectionInfo334BCLUB)

    '            Conn.Open()
    '            Using cmd As New SqlCommand(Sql)
    '                cmd.Connection = Conn
    '                cmd.CommandType = CommandType.Text
    '                Using reader As SqlDataReader = cmd.ExecuteReader()
    '                    While (reader.Read())
    '                        Dim RowData = New DataAccount
    '                        Dim idx As Integer = 0
    '                        RowData.Region = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.Zone = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.ClubCode = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.ClubNameShort = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.MemberCode = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.MemberFirstName = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.MemberLastName = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.MemberNameKana = reader.GetValue(idx).ToString : idx += 1

    '                        DbAccountSet.Data.Add(RowData)
    '                    End While
    '                    Dim jsonObj As String = JsonConvert.SerializeObject(DbAccountSet)
    '                    Result += jsonObj
    '                End Using
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        Result = ex.Message
    '    End Try
    'End Sub

    '-----------------------------------------------------------------------------------------------
    ' ホームTop画面：地区スローガン
    '-----------------------------------------------------------------------------------------------
    ' 地区スローガンDBクラス
    '-----------------------------------------------------------------------------------------------
    ' 地区スローガンDB検索＆結果取得
    '-----------------------------------------------------------------------------------------------
    'Sub GetDBSlogan(ByVal NowDate As String)

    '    Dim Sql As String
    '    Dim NowYear As String

    '    WUtil = New WebUtil
    '    NowYear = WUtil.GetFiscalYear("334-B", NowDate)

    '    DbSlogan = New DBSlogan With {.Data = New List(Of DataSlogan)}

    '    ' SQL作成
    '    Sql = "SELECT " &
    '            "Slogan, " &
    '            "DistrictGovernor " &
    '          "FROM T_SLOGAN WHERE " &
    '            "FiscalStart <= '" & NowYear & "' " &
    '          "AND " &
    '              "FiscalEnd > '" & NowYear & "'" &
    '          ";"

    '    Try
    '        Using Conn As New SqlConnection
    '            Conn.ConnectionString = (ConnectionInfo334B)

    '            Conn.Open()
    '            Using cmd As New SqlCommand(Sql)
    '                cmd.Connection = Conn
    '                cmd.CommandType = CommandType.Text
    '                Using reader As SqlDataReader = cmd.ExecuteReader()
    '                    While (reader.Read())
    '                        Dim RowData = New DataSlogan
    '                        Dim idx As Integer = 0
    '                        RowData.Slogan = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.DistrictGovernor = reader.GetValue(idx).ToString : idx += 1

    '                        DbSlogan.Data.Add(RowData)
    '                    End While
    '                    Dim jsonObj As String = JsonConvert.SerializeObject(DbSlogan)
    '                    Result += jsonObj
    '                End Using
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        Result = ex.Message
    '    End Try
    'End Sub

    '-----------------------------------------------------------------------------------------------
    ' ホームTop画面：キャビネットレター
    '-----------------------------------------------------------------------------------------------
    ' キャビネットレターDB検索＆結果取得
    '-----------------------------------------------------------------------------------------------
    'Sub GetDBTopLetter(ByVal NowDate As String,
    '                   ByVal NowTime As String,
    '                   ByVal DataNum As String)

    '    Dim Sql As String

    '    DbTopLetter = New DBTopLetter With {.Data = New List(Of DataTopLetter)}

    '    ' SQL作成
    '    Sql = "SELECT " &
    '            "Top " & DataNum & " " &
    '            "DataNo, " &
    '            "FORMAT(EventDate, 'MM/dd'), " &
    '            "EventTime, " &
    '            "Title " &
    '          "From T_LETTER " &
    '          "WHERE " &
    '             "OpenFlg = '1' " &
    '             "AND " &
    '             "DelFlg = '0' " &
    '          "ORDER BY EventDate DESC" &
    '          ";"

    '    Try
    '        Using Conn As New SqlConnection
    '            Conn.ConnectionString = (ConnectionInfo334B)

    '            Conn.Open()
    '            Using cmd As New SqlCommand(Sql)
    '                cmd.Connection = Conn
    '                cmd.CommandType = CommandType.Text
    '                Using reader As SqlDataReader = cmd.ExecuteReader()
    '                    While (reader.Read())
    '                        Dim RowData = New DataTopLetter
    '                        Dim idx As Integer = 0
    '                        RowData.DataNo = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.EventDate = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.EventTime = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.Title = reader.GetValue(idx).ToString : idx += 1

    '                        DbTopLetter.Data.Add(RowData)
    '                    End While
    '                    Dim jsonObj As String = JsonConvert.SerializeObject(DbTopLetter)
    '                    Result += jsonObj
    '                End Using
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        Result = ex.Message
    '    End Try
    'End Sub

    '-----------------------------------------------------------------------------------------------
    ' ホームTop画面：参加予定一覧
    '-----------------------------------------------------------------------------------------------
    ' イベント出欠＆キャビネットイベント＆年間例会スケジュールDB検索＆結果取得
    '-----------------------------------------------------------------------------------------------
    'Sub GetDBTopEvent(ByVal MemberCode As String,
    '                  ByVal NowDate As String,
    '                  ByVal NowTime As String,
    '                  ByVal DataNum As String)

    '    Dim Sql As String

    '    DbTopEvent = New DBTopEvent With {.Data = New List(Of DataTopEvent)}

    '    ' SQL作成
    '    Sql = "SELECT " &
    '            "Top " & DataNum & " " &
    '            "t1.DataNo, " &
    '            "t1.EventClass, " &
    '            "t1.EventDataNo, " &
    '            "FORMAT(t1.EventDate, 'MM/dd'), " &
    '            "t1.Answer, " &
    '            "t2.Title, " &
    '            "t3.MeetingName, " &
    '            "datediff(day, '" & NowDate & "', t1.EventDate) " &
    '          "FROM " &
    '            "[334B].dbo.T_EVENTRET t1 " &
    '          "LEFT JOIN " &
    '            "[334B].dbo.T_EVENT t2 " &
    '          "ON " &
    '            "t1.EventClass = '1' and " &
    '            "t1.EventDataNo = t2.DataNo " &
    '          "LEFT JOIN " &
    '            "[334BCLUB].dbo.T_MEETINGSCHEDULE t3 " &
    '          "ON " &
    '            "t1.EventClass = '2' and " &
    '            "t1.EventDataNo = t3.DataNo " &
    '          "WHERE " &
    '            "t1.MemberCode = '" & MemberCode & "' AND " &
    '            "t1.EventDate >= '" & NowDate & "' AND " &
    '            "(t1.Answer = '1' or t1.Answer = '3') AND " &
    '            "t1.CancelFlg = '0' and " &
    '            "t1.DelFlg = '0' " &
    '          ";"

    '    Try
    '        Using Conn As New SqlConnection
    '            Conn.ConnectionString = (ConnectionInfo334B)

    '            Conn.Open()
    '            Using cmd As New SqlCommand(Sql)
    '                cmd.Connection = Conn
    '                cmd.CommandType = CommandType.Text
    '                Using reader As SqlDataReader = cmd.ExecuteReader()
    '                    While (reader.Read())
    '                        Dim RowData = New DataTopEvent
    '                        Dim idx As Integer = 0
    '                        RowData.DataNo = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.EventClass = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.EventDateNo = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.EventDate = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.Answer = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.Title = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.MeetingName = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.CountDate = reader.GetValue(idx).ToString : idx += 1

    '                        DbTopEvent.Data.Add(RowData)
    '                    End While
    '                    Dim jsonObj As String = JsonConvert.SerializeObject(DbTopEvent)
    '                    Result += jsonObj
    '                End Using
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        Result = ex.Message
    '    End Try
    'End Sub

    '-----------------------------------------------------------------------------------------------
    ' キャビネットレター画面
    '-----------------------------------------------------------------------------------------------
    ' キャビネットレターDB検索＆結果取得
    '-----------------------------------------------------------------------------------------------
    'Sub GetDBLetter(ByVal DataNo As String)

    '    Dim Sql As String

    '    DbLetter = New DBLetter With {.Data = New List(Of DataLetter)}

    '    ' SQL作成
    '    Sql = "SELECT " &
    '            "DataNo, " &
    '            "FORMAT(EventDate, 'yyyy/MM/dd'), " &
    '            "EventTime, " &
    '            "Title, " &
    '            "Body, " &
    '            "Image1FileName, " &
    '            "Image1FileName " &
    '          "From T_LETTER " &
    '          "WHERE " &
    '             "DataNo = '" & DataNo & "' AND " &
    '             "OpenFlg = '1' AND " &
    '             "DelFlg = '0' " &
    '          ";"

    '    Try
    '        Using Conn As New SqlConnection
    '            Conn.ConnectionString = (ConnectionInfo334B)

    '            Conn.Open()
    '            Using cmd As New SqlCommand(Sql)
    '                cmd.Connection = Conn
    '                cmd.CommandType = CommandType.Text
    '                Using reader As SqlDataReader = cmd.ExecuteReader()
    '                    While (reader.Read())
    '                        Dim RowData = New DataLetter
    '                        Dim idx As Integer = 0
    '                        RowData.DataNo = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.EventDate = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.EventTime = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.Title = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.Body = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.Image1FileName = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.Image2FileName = reader.GetValue(idx).ToString : idx += 1
    '                        DbLetter.Data.Add(RowData)
    '                    End While
    '                    Dim jsonObj As String = JsonConvert.SerializeObject(DbLetter)
    '                    Result += jsonObj
    '                End Using
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        Result = ex.Message
    '    End Try
    'End Sub

    '-----------------------------------------------------------------------------------------------
    ' キャビネットレター一覧画面
    '-----------------------------------------------------------------------------------------------
    ' キャビネットレターDB検索＆結果取得
    '-----------------------------------------------------------------------------------------------
    'Sub GetDBLetterList(ByVal DistrictCode As String,
    '                    ByVal NowDate As String,
    '                    ByVal NowTime As String)

    '    Dim Sql As String

    '    Dim FiscalYear As String = ""
    '    Dim PeriodStartYmd As String = ""
    '    Dim PeriodEndYmd As String = ""
    '    Dim PeriodStartYmdLy As String = ""
    '    Dim PeriodEndYmdLy As String = ""

    '    ' 設定ファイルDBから「前年度開始」「年度終了」を取得する。
    '    WUtil = New WebUtil
    '    WUtil.GetPeriodYmdLy(DistrictCode, NowDate,
    '                            FiscalYear,
    '                            PeriodStartYmd,
    '                            PeriodEndYmd,
    '                            PeriodStartYmdLy,
    '                            PeriodEndYmdLy)


    '    DbTopLetter = New DBTopLetter With {.Data = New List(Of DataTopLetter)}

    '    ' SQL作成
    '    Sql = "SELECT " &
    '            "DataNo, " &
    '            "FORMAT(EventDate, 'yyyy/MM/dd'), " &
    '            "EventTime, " &
    '            "Title " &
    '          "From T_LETTER " &
    '          "WHERE " &
    '             "OpenFlg = '1' AND " &
    '             "DelFlg = '0' AND " &
    '             "(EventDate >= '" & PeriodStartYmdLy & "' AND " &
    '             "EventDate <= '" & PeriodEndYmd & "') " &
    '          "ORDER BY EventDate DESC" &
    '          ";"

    '    Try
    '        Using Conn As New SqlConnection
    '            Conn.ConnectionString = (ConnectionInfo334B)

    '            Conn.Open()
    '            Using cmd As New SqlCommand(Sql)
    '                cmd.Connection = Conn
    '                cmd.CommandType = CommandType.Text
    '                Using reader As SqlDataReader = cmd.ExecuteReader()
    '                    While (reader.Read())
    '                        Dim RowData = New DataTopLetter
    '                        Dim idx As Integer = 0
    '                        RowData.DataNo = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.EventDate = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.EventTime = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.Title = reader.GetValue(idx).ToString : idx += 1

    '                        DbTopLetter.Data.Add(RowData)
    '                    End While
    '                    Dim jsonObj As String = JsonConvert.SerializeObject(DbTopLetter)
    '                    Result += jsonObj
    '                End Using
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        Result = ex.Message
    '    End Try
    'End Sub

    '-----------------------------------------------------------------------------------------------
    ' イベント一覧
    '-----------------------------------------------------------------------------------------------
    ' イベント出欠＆キャビネットイベント＆年間例会スケジュールDB検索＆結果取得
    '-----------------------------------------------------------------------------------------------
    'Sub GetDBEventList(ByVal MemberCode As String,
    '                  ByVal NowDate As String,
    '                  ByVal NowTime As String)

    '    Dim Sql As String

    '    DbEventList = New DBEventList With {.Data = New List(Of DataEventList)}

    '    ' SQL作成
    '    Sql = "SELECT " &
    '            "t1.DataNo, " &
    '            "t1.EventClass, " &
    '            "t1.EventDataNo, " &
    '            "FORMAT(t1.EventDate, 'yyyy/MM/dd'), " &
    '            "t1.Answer, " &
    '            "t2.EventPlace, " &
    '            "t2.Title, " &
    '            "t3.MeetingPlace, " &
    '            "t3.MeetingName, " &
    '            "datediff(day, " & NowDate & ", t1.EventDate) " &
    '          "FROM " &
    '            "[334B].dbo.T_EVENTRET t1 " &
    '          "LEFT JOIN " &
    '            "[334B].dbo.T_EVENT t2 " &
    '          "ON " &
    '            "t1.EventClass = '1' and " &
    '            "t1.EventDataNo = t2.DataNo " &
    '          "LEFT JOIN " &
    '            "[334BCLUB].dbo.T_MEETINGSCHEDULE t3 " &
    '          "ON " &
    '            "t1.EventClass = '2' and " &
    '            "t1.EventDataNo = t3.DataNo " &
    '          "WHERE " &
    '            "t1.MemberCode = '" & MemberCode & "' AND " &
    '            "t1.EventDate >= '" & NowDate & "' AND " &
    '            "t1.CancelFlg = '0' and " &
    '            "t1.DelFlg = '0' " &
    '          ";"

    '    Try
    '        Using Conn As New SqlConnection
    '            Conn.ConnectionString = (ConnectionInfo334B)

    '            Conn.Open()
    '            Using cmd As New SqlCommand(Sql)
    '                cmd.Connection = Conn
    '                cmd.CommandType = CommandType.Text
    '                Using reader As SqlDataReader = cmd.ExecuteReader()
    '                    While (reader.Read())
    '                        Dim RowData = New DataEventList
    '                        Dim idx As Integer = 0
    '                        RowData.DataNo = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.EventClass = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.EventDateNo = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.EventDate = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.Answer = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.EventPlace = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.Title = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.MeetingPlace = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.MeetingName = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.CountDate = reader.GetValue(idx).ToString : idx += 1

    '                        DbEventList.Data.Add(RowData)
    '                    End While
    '                    Dim jsonObj As String = JsonConvert.SerializeObject(DbEventList)
    '                    Result += jsonObj
    '                End Using
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        Result = ex.Message
    '    End Try
    'End Sub

    '-----------------------------------------------------------------------------------------------
    ' アカウント画面
    '-----------------------------------------------------------------------------------------------
    ' 会員マスタ検索＆結果取得
    '-----------------------------------------------------------------------------------------------
    'Sub GetDBMember(ByVal Region As String,
    '                ByVal Zone As String,
    '                ByVal ClubCode As String,
    '                ByVal MemberCode As String,
    '                ByVal NowDate As String)

    '    Dim Sql As String
    '    Dim NowYear As String

    '    DbMember = New DBMember
    '    DbMember.Data = New List(Of DataMember)

    '    WUtil = New WebUtil
    '    NowYear = WUtil.GetFiscalYear("334-B", NowDate)

    '    ' SQL作成
    '    Sql = "SELECT " &
    '            "t1.Region, " &
    '            "t1.Zone, " &
    '            "t1.ClubCode, " &
    '            "t1.ClubNameShort, " &
    '            "t2.ClubName, " &
    '            "t1.MemberCode, " &
    '            "t1.MemberFirstName, " &
    '            "t1.MemberLastName, " &
    '            "t1.Sex, " &
    '            "t1.TypeName, " &
    '            "CONVERT(DATE, t1.JoinDate), " &
    '            "t1.ExecutiveName, " &
    '            "t1.ExecutiveName1, " &
    '            "t1.CommitteeName, " &
    '            "t1.CommitteeFlg, " &
    '            "t1.CommitteeName1, " &
    '            "t1.CommitteeFlg1, " &
    '            "t1.CommitteeName2, " &
    '            "t1.CommitteeFlg2, " &
    '            "t1.CommitteeName3, " &
    '            "t1.CommitteeFlg3, " &
    '            "t3.DistrictClass, " &
    '            "t3.DistrictName, " &
    '            "t3.DistrictClass1, " &
    '            "t3.DistrictName1, " &
    '            "t3.DistrictClass2, " &
    '            "t3.DistrictName2, " &
    '            "t3.DistrictClass3, " &
    '            "t3.DistrictName3, " &
    '            "t3.DistrictClass4, " &
    '            "t3.DistrictName4, " &
    '            "t3.DistrictClass5, " &
    '            "t3.DistrictName5 " &
    '        "FROM " &
    '            "[334BCLUB].dbo.M_MEMBER t1 " &
    '        "LEFT JOIN " &
    '            "[334B].dbo.M_CLUB t2 " &
    '        "ON " &
    '            "t1.Region = t2.Region and " &
    '            "t1.Zone = t2.Zone and " &
    '            "t1.ClubCode = t2.ClubCode " &
    '        "LEFT JOIN " &
    '            "[334B].dbo.M_CABINET t3 " &
    '        "ON " &
    '            "t1.MemberCode = t3.MemberCode and " &
    '            "t1.ClubCode = t3.ClubCode "
    '    Sql = Sql & " WHERE "
    '    If (Region.Length > 0) Then
    '        Sql = Sql & " t1.Region = '" & Region & "'"
    '    End If
    '    If (Zone.Length > 0) Then
    '        If (Region.Length > 0) Then
    '            Sql = Sql & " and"
    '        End If
    '        Sql = Sql & " t1.Zone = '" & Zone & "'"
    '    End If
    '    If (ClubCode.Length > 0) Then
    '        If (Region.Length > 0) Or (Zone.Length > 0) Then
    '            Sql = Sql & " and"
    '        End If
    '        Sql = Sql & " t1.ClubCode = '" & ClubCode & "'"
    '    End If
    '    If (MemberCode.Length > 0) Then
    '        If (Region.Length > 0) Or (Zone.Length > 0) Or (ClubCode.Length > 0) Then
    '            Sql = Sql & " and"
    '        End If
    '        Sql = Sql & " t1.MemberCode = '" & MemberCode & "'"
    '    End If
    '    If (Region.Length > 0) Or (Zone.Length > 0) Or (ClubCode.Length > 0) Or (MemberCode.Length > 0) Then
    '        Sql = Sql & " and"
    '    End If
    '    Sql = Sql & " t1.LeaveFlg = '0' and " &
    '        "t2.DelFlg = '0' or t2.DelFlg = null and " &
    '        "t3.DelFlg = '0' or t3.DelFlg = null and " &
    '        "t3.FiscalStart <= '" & NowYear & "' or t3.FiscalEnd >= '" & NowYear & "' or t3.FiscalStart = null or t3.FiscalEnd = null;"
    '    Try
    '        Using Conn As New SqlConnection
    '            Conn.ConnectionString = (ConnectionInfo334BCLUB)

    '            Conn.Open()
    '            Using cmd As New SqlCommand(Sql)
    '                cmd.Connection = Conn
    '                cmd.CommandType = CommandType.Text
    '                Using reader As SqlDataReader = cmd.ExecuteReader()
    '                    While (reader.Read())
    '                        Dim RowData = New DataMember
    '                        Dim idx As Integer = 0
    '                        RowData.Region = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.Zone = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.ClubCode = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.ClubNameShort = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.ClubName = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.MemberCode = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.MemberFirstName = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.MemberLastName = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.Sex = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.TypeName = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.JoinDate = reader.GetValue(idx) : idx += 1
    '                        RowData.ExecutiveName = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.ExecutiveName1 = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.CommitteeName = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.CommitteeFlg = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.CommitteeName1 = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.CommitteeFlg1 = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.CommitteeName2 = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.CommitteeFlg2 = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.CommitteeName3 = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.CommitteeFlg3 = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.DistrictClass = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.DistrictName = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.DistrictClass1 = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.DistrictName1 = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.DistrictClass2 = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.DistrictName2 = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.DistrictClass3 = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.DistrictName3 = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.DistrictClass4 = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.DistrictName4 = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.DistrictClass5 = reader.GetValue(idx).ToString : idx += 1
    '                        RowData.DistrictName5 = reader.GetValue(idx).ToString : idx += 1
    '                        DbMember.Data.Add(RowData)
    '                    End While
    '                    Dim jsonObj As String = JsonConvert.SerializeObject(DbMember)
    '                    Result += jsonObj
    '                End Using
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        Result = ex.Message
    '    End Try
    'End Sub

End Class
