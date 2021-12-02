Imports System.Data.SQLite
Imports log4net

Public Class TableManager
    Protected _logger As ILog                                                       ' ロガー
    Protected _utl As CmnUtility = Nothing                                          ' ユーティリティ

    Protected _strTableName As String = String.Empty                                ' 操作対象テーブル名

    Public _lstFld As New List(Of String)                                           'フィールド名
    Protected _lstKey As New List(Of String)                                        ' Key 情報
    Protected _lstKeyFld As New List(Of String)                                     'キーフィールド名


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="logger"></param>
    Public Sub New(ByRef logger As ILog, ByRef utl As CmnUtility, ByRef Tablename As String)
        _logger = logger
        _utl = utl
        _strTableName = Tablename

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' フィールド文字列(項目名) GetFldString 引数用
    ''' </summary>
    Public Enum FldStringType
        NameOnly 'フィールド名のみ
        KeyNameOnly 'キーフィールド名のみ
        SQL_MergeKey
        SQL_MergeFld
        SQL_MergeFld2
    End Enum

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' フィールド文字列(項目名)の作成
    ''' </summary>
    ''' <param name="wType">作成タイプ FldStringType</param>
    ''' <param name="bolNewLine">改行</param>
    ''' <returns></returns>
    Public Function GetFldString(wType As FldStringType, ByVal bolNewLine As Boolean) As String
        Dim strRet As String = String.Empty
        Dim strNewLine As String = String.Empty
        Dim bolFirst As Boolean

        If bolNewLine Then
            strNewLine = System.Environment.NewLine
        End If

        Select Case wType
            Case FldStringType.NameOnly
                'フィールド名をカンマ区切り
                'フィールド名セット
                bolFirst = True
                For Each wFld As String In _lstFld
                    '2行目以降は改行とカンマ
                    strRet &= IIf(bolFirst, " ", strNewLine & ",").ToString
                    strRet &= wFld
                    bolFirst = False
                Next

            Case FldStringType.KeyNameOnly
                'キーフィールド名をカンマ区切り
                'フィールド名セット
                bolFirst = True
                For Each wFld As String In _lstKeyFld
                    '2行目以降は改行とカンマ
                    strRet &= IIf(bolFirst, " ", strNewLine & ",").ToString
                    strRet &= wFld
                    bolFirst = False
                Next

            Case FldStringType.SQL_MergeKey
                '"AND A.KinmuDate = B.KinmuDate "

                'フィールド名セット
                bolFirst = True
                For Each wFld As String In _lstKeyFld
                    '2行目以降は改行とカンマ
                    strRet &= IIf(bolFirst, " ", strNewLine & " AND ").ToString
                    strRet &= String.Format("A.{0} = B.{0} ", wFld)
                    bolFirst = False
                Next

            Case FldStringType.SQL_MergeFld
                '",ShainName = B.ShainName "

                'フィールド名セット
                bolFirst = True
                For Each wFld As String In _lstFld
                    If Not _lstKeyFld.Contains(wFld) Then 'キー以外のフィールド
                        '2行目以降は改行とカンマ
                        strRet &= IIf(bolFirst, " ", strNewLine & ",").ToString
                        strRet &= String.Format("{0} = B.{0} ", wFld)
                        bolFirst = False
                    End If
                Next

            Case FldStringType.SQL_MergeFld2
                '" B.ShainCode "
                'フィールド名セット
                bolFirst = True
                For Each wFld As String In _lstFld
                    '2行目以降は改行とカンマ
                    strRet &= IIf(bolFirst, " ", strNewLine & ",").ToString
                    strRet &= "B." & wFld
                    bolFirst = False
                Next

        End Select

        Return strRet

    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' フィールド文字列(項目名/値)の作成
    ''' </summary>
    ''' <param name="dicData">値</param>
    ''' <param name="wType">作成タイプ FldValueStringType</param>
    ''' <param name="bolNewLine">改行</param>
    ''' <returns></returns>
    Public Function GetFldValueString(ByRef dicData As Dictionary(Of String, String),
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
                    strRet &= FieldSQL_srv(wFld, dicData(wFld))
                    bolFirst = False
                Next

            Case FldValueStringType.SQLsrv2 'T_INPUT系用
                '",'xxx' "
                bolFirst = True
                For Each wFld As String In _lstFld
                    strRet &= IIf(bolFirst, " ", strNewLine & ",").ToString
                    strRet &= FieldSQL(wFld, dicData(wFld))
                    bolFirst = False
                Next

            Case FldValueStringType.SQL_Where
                '" AND KinmuDate ={0}"

                bolFirst = True
                For Each wFld As String In _lstKeyFld
                    strRet &= IIf(bolFirst, " ", strNewLine & " AND ").ToString
                    strRet &= String.Format("{0} = {1}", wFld, FieldSQL(wFld, dicData(wFld)))
                    bolFirst = False
                Next

            Case FldValueStringType.SQLite1
                '" 'xxx' "
                bolFirst = True
                For Each wFld As String In _lstFld
                    strRet &= IIf(bolFirst, " ", strNewLine & ",").ToString
                    strRet &= FieldSQL_Lite(wFld, dicData(wFld))
                    bolFirst = False
                Next

        End Select

        Return strRet

    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' 項目別 値の桁チェック＆カット
    ''' </summary>
    ''' <param name="strFldName">フィールド名</param>
    ''' <param name="strFldValue">値</param>
    ''' <returns>加工後 値</returns>
    Public Overridable Function CutKeta(ByVal strFldName As String, ByVal strFldValue As String) As String

        Dim strRet As String = strFldValue
        Dim eType As CmnUtility.eValueType = CmnUtility.eValueType.String '型
        Dim intKeta As Integer = -1 '桁

        'フィールド別に型・桁を取得
        Field_Difinition(strFldName, eType, intKeta)

        '桁オーバーチェック
        If _utl.GetStringByte(strRet) > intKeta Then
            _logger.Error(String.Format("●{0} {1}='{2}' 桁({3})オーバー", _strTableName, strFldName, strRet, intKeta.ToString))

            'カット
            strRet = _utl.CutString(strRet, intKeta)
        End If

        Return strRet

    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' テーブル項目の定義（項目名）
    ''' </summary>
    ''' <param name="strFldName"></param>
    Public Sub Field_Name(strFldName As String)
        _lstFld.Clear()
        Select Case strFldName
            Case "A_SETTING"
                '------------------
                _lstFld.Add("DistrictCode")
                _lstFld.Add("DistrictName")
                _lstFld.Add("CabinetName")
                _lstFld.Add("PeriodStart")
                _lstFld.Add("PeriodEnd")
                _lstFld.Add("DistrictID")
                _lstFld.Add("MagazineMoney")
                _lstFld.Add("EventDataDay")
                _lstFld.Add("AdminID")
                _lstFld.Add("AdminName")
                _lstFld.Add("AdminPassWord")
                _lstFld.Add("CabinetEmailAddeess")
                _lstFld.Add("AdminEmailAddress")
                _lstFld.Add("CabinetTelNo")
                _lstFld.Add("AdminTelNo")
                _lstFld.Add("VersionNo")
                '------------------
            Case "A_DATANO"
                '------------------
                _lstFld.Add("DataClass")
                _lstFld.Add("DataName")
                _lstFld.Add("LastDataNo")
                '------------------
            Case "A_ACCOUNT"
                '------------------
                _lstFld.Add("Region")
                _lstFld.Add("Zone")
                _lstFld.Add("ClubCode")
                _lstFld.Add("ClubName")
                _lstFld.Add("MemberCode")
                _lstFld.Add("MemberFirstName")
                _lstFld.Add("MemberLastName")
                _lstFld.Add("AccountDate")
                _lstFld.Add("LastUpdDate")
                '------------------
            Case "A_FILEPATH"
                '------------------
                _lstFld.Add("DataClass")
                _lstFld.Add("FilePath")
                '------------------
            Case "T_SLOGAN"
                '------------------
                _lstFld.Add("DataNo")
                _lstFld.Add("FiscalStart")
                _lstFld.Add("FiscalEnd")
                _lstFld.Add("Slogan")
                _lstFld.Add("DistrictGovernor")
                '------------------
            Case "T_LETTER"
                '------------------
                _lstFld.Add("DataNo")
                _lstFld.Add("EventDate")
                _lstFld.Add("EventTime")
                _lstFld.Add("Title")
                _lstFld.Add("Body")
                _lstFld.Add("Image1FileName")
                _lstFld.Add("Image2FileName")
                _lstFld.Add("NoticeFlg")
                '------------------
            Case "T_EVENTRET"
                '------------------
                _lstFld.Add("DataNo")
                _lstFld.Add("EventClass")
                _lstFld.Add("EventDataNo")
                _lstFld.Add("EventDate")
                _lstFld.Add("ClubCode")
                _lstFld.Add("ClubNameShort")
                _lstFld.Add("MemberCode")
                _lstFld.Add("MemberName")
                _lstFld.Add("DirectorName")
                _lstFld.Add("Answer")
                _lstFld.Add("AnswerLate")
                _lstFld.Add("AnswerEarly")
                _lstFld.Add("Online")
                _lstFld.Add("Option1")
                _lstFld.Add("Option2")
                _lstFld.Add("Option3")
                _lstFld.Add("Option4")
                _lstFld.Add("Option5")
                _lstFld.Add("OtherCount")
                _lstFld.Add("TargetFlg")
                _lstFld.Add("CancelFlg")
                '------------------
            Case "T_EVENT"
                '------------------
                _lstFld.Add("DataNo")
                _lstFld.Add("EventDate")
                _lstFld.Add("EventTimeStart")
                _lstFld.Add("EventTimeEnd")
                _lstFld.Add("ReceptionTime")
                _lstFld.Add("EventPlace")
                _lstFld.Add("Title")
                _lstFld.Add("Body")
                _lstFld.Add("OptionName1")
                _lstFld.Add("OptionRadio1")
                _lstFld.Add("OptionName2")
                _lstFld.Add("OptionRadio2")
                _lstFld.Add("OptionName3")
                _lstFld.Add("OptionRadio3")
                _lstFld.Add("OptionName4")
                _lstFld.Add("OptionRadio4")
                _lstFld.Add("OptionName5")
                _lstFld.Add("OptionRadio5")
                _lstFld.Add("Sake")
                _lstFld.Add("Meeting")
                _lstFld.Add("MeetingUrl")
                _lstFld.Add("MeetingID")
                _lstFld.Add("MeetingPW")
                _lstFld.Add("MeetingOther")
                _lstFld.Add("AnswerDate")
                _lstFld.Add("FileName")
                _lstFld.Add("CabinetUser")
                _lstFld.Add("ClubUser")
                _lstFld.Add("NoticeFlg")
                '------------------
            Case "T_INFOMATION_CABI"
                '------------------
                _lstFld.Add("DataNo")
                _lstFld.Add("AddDate")
                _lstFld.Add("Subject")
                _lstFld.Add("Detail")
                _lstFld.Add("FileName")
                _lstFld.Add("InfoFlg")
                _lstFld.Add("InfoUser")
                _lstFld.Add("NoticeFlg")
                '------------------
            Case "T_MAGAZINE"
                '------------------
                _lstFld.Add("DataNo")
                _lstFld.Add("SortNo")
                _lstFld.Add("Magazine")
                _lstFld.Add("FileName")
                _lstFld.Add("MagazineClass")
                _lstFld.Add("MagazinePrice")
                '------------------
            Case "T_MAGAZINEBUY"
                '------------------
                _lstFld.Add("DataNo")
                _lstFld.Add("MagazineDataNo")
                _lstFld.Add("Magazine")
                _lstFld.Add("BuyDate")
                _lstFld.Add("BuyNumber")
                _lstFld.Add("MagazinePrice")
                _lstFld.Add("MoneyTotal")
                _lstFld.Add("Region")
                _lstFld.Add("Zone")
                _lstFld.Add("ClubCode")
                _lstFld.Add("ClubNameShort")
                _lstFld.Add("MemberCode")
                _lstFld.Add("MemberName")
                _lstFld.Add("ShippingDate")
                _lstFld.Add("PaymentDate")
                _lstFld.Add("Payment")
                _lstFld.Add("DelFlg")
                '------------------
            Case "M_DISTRICTOFFICER"
                '------------------
                _lstFld.Add("DistrictClass")
                _lstFld.Add("DistrictCode")
                _lstFld.Add("DistrictName")
                '------------------
            Case "M_CABINET"
                '------------------
                _lstFld.Add("FiscalStart")
                _lstFld.Add("FiscalEnd")
                _lstFld.Add("MemberCode")
                _lstFld.Add("MemberName")
                _lstFld.Add("ClubCode")
                _lstFld.Add("ClubNameShort")
                _lstFld.Add("DistrictClass")
                _lstFld.Add("DistrictCode")
                _lstFld.Add("DistrictName")
                _lstFld.Add("DistrictClass1")
                _lstFld.Add("DistrictCode1")
                _lstFld.Add("DistrictName1")
                _lstFld.Add("DistrictClass2")
                _lstFld.Add("DistrictCode2")
                _lstFld.Add("DistrictName2")
                _lstFld.Add("DistrictClass3")
                _lstFld.Add("DistrictCode3")
                _lstFld.Add("DistrictName3")
                _lstFld.Add("DistrictClass4")
                _lstFld.Add("DistrictCode4")
                _lstFld.Add("DistrictName4")
                _lstFld.Add("DistrictClass5")
                _lstFld.Add("DistrictCode5")
                _lstFld.Add("DistrictName5")
                '------------------
            Case "M_CLUB"
                '------------------
                _lstFld.Add("Region")
                _lstFld.Add("Zone")
                _lstFld.Add("Sort")
                _lstFld.Add("ClubCode")
                _lstFld.Add("ClubName")
                _lstFld.Add("ClubNameShort")
                _lstFld.Add("PassWord")
                _lstFld.Add("FormationDate")
                _lstFld.Add("CharterNight")
                _lstFld.Add("MeetingDate")
                _lstFld.Add("MeetingTime")
                _lstFld.Add("MeetingPlace")
                _lstFld.Add("AnswerDay")
                _lstFld.Add("AnswerTime")
                _lstFld.Add("Zip")
                _lstFld.Add("Address")
                _lstFld.Add("Tel")
                _lstFld.Add("HP")
                _lstFld.Add("MailAddress")
                '------------------
            Case "T_CLUBSLOGAN"
                '------------------
                _lstFld.Add("DataNo")
                _lstFld.Add("ClubCode")
                _lstFld.Add("ClubNameShort")
                _lstFld.Add("FiscalStart")
                _lstFld.Add("FiscalEnd")
                _lstFld.Add("ClubSlogan")
                _lstFld.Add("ExecutiveName")
                '------------------
            Case "T_MEETINGSCHEDULE"
                '------------------
                _lstFld.Add("DataNo")
                _lstFld.Add("ClubCode")
                _lstFld.Add("ClubNameShort")
                _lstFld.Add("Fiscal")
                _lstFld.Add("MeetingDate")
                _lstFld.Add("MeetingTime")
                _lstFld.Add("MeetingPlace")
                _lstFld.Add("MeetingCount")
                _lstFld.Add("MeetingName")
                _lstFld.Add("Online")
                _lstFld.Add("AnswerDate")
                _lstFld.Add("RemarksItems")
                _lstFld.Add("RemarksCheck")
                _lstFld.Add("Remarks")
                _lstFld.Add("OptionName1")
                _lstFld.Add("OptionRadio1")
                _lstFld.Add("OptionName2")
                _lstFld.Add("OptionRadio2")
                _lstFld.Add("OptionName3")
                _lstFld.Add("OptionRadio3")
                _lstFld.Add("OptionName4")
                _lstFld.Add("OptionRadio4")
                _lstFld.Add("OptionName5")
                _lstFld.Add("OptionRadio5")
                _lstFld.Add("Sake")
                _lstFld.Add("OtherUser")
                _lstFld.Add("CancelFlg")
                '------------------
            Case "T_DIRECTOR"
                '------------------
                _lstFld.Add("DataNo")
                _lstFld.Add("ClubCode")
                _lstFld.Add("ClubNameShort")
                _lstFld.Add("EventClass")
                _lstFld.Add("CommitteeCode")
                _lstFld.Add("CommitteeName")
                _lstFld.Add("Season")
                _lstFld.Add("EventDate")
                _lstFld.Add("EventTime")
                _lstFld.Add("EventPlace")
                _lstFld.Add("Subject")
                _lstFld.Add("Agenda")
                _lstFld.Add("Member")
                _lstFld.Add("MemberAdd")
                _lstFld.Add("AnswerDate")
                _lstFld.Add("CancelFlg")
                _lstFld.Add("NoticeFlg")
                '------------------
            Case "T_MEETINGPROGRAM"
                '------------------
                _lstFld.Add("DataNo")
                _lstFld.Add("ScheduleDataNo")
                _lstFld.Add("ClubCode")
                _lstFld.Add("ClubNameShort")
                _lstFld.Add("Fiscal")
                _lstFld.Add("FileName")
                _lstFld.Add("FileName1")
                _lstFld.Add("FileName2")
                _lstFld.Add("FileName3")
                _lstFld.Add("FileName4")
                _lstFld.Add("FileName5")
                _lstFld.Add("Meeting")
                _lstFld.Add("MeetingUrl")
                _lstFld.Add("MeetingID")
                _lstFld.Add("MeetingPW")
                _lstFld.Add("MeetingOther")
                _lstFld.Add("NoticeFlg")
                '------------------
            Case "T_INFOMATION_CLUB"
                '------------------
                _lstFld.Add("DataNo")
                _lstFld.Add("ClubCode")
                _lstFld.Add("ClubNameShort")
                _lstFld.Add("AddDate")
                _lstFld.Add("Subject")
                _lstFld.Add("Detail")
                _lstFld.Add("FileName")
                _lstFld.Add("InfoFlg")
                _lstFld.Add("TypeCode")
                _lstFld.Add("InfoUser")
                _lstFld.Add("NoticeFlg")
                '------------------
            Case "M_MEMBER"
                '------------------
                _lstFld.Add("Region")
                _lstFld.Add("Zone")
                _lstFld.Add("ClubCode")
                _lstFld.Add("ClubNameShort")
                _lstFld.Add("MemberCode")
                _lstFld.Add("MemberFirstName")
                _lstFld.Add("MemberLastName")
                _lstFld.Add("MemberNameKana")
                _lstFld.Add("TypeCode")
                _lstFld.Add("TypeName")
                _lstFld.Add("Sex")
                _lstFld.Add("Tel")
                _lstFld.Add("JoinDate")
                _lstFld.Add("LeaveDate")
                _lstFld.Add("LeaveFlg")
                _lstFld.Add("AccountDate")
                _lstFld.Add("ExecutiveCode")
                _lstFld.Add("ExecutiveName")
                _lstFld.Add("ExecutiveCode1")
                _lstFld.Add("ExecutiveName1")
                _lstFld.Add("CommitteeCode")
                _lstFld.Add("CommitteeName")
                _lstFld.Add("CommitteeFlg")
                _lstFld.Add("CommitteeCode1")
                _lstFld.Add("CommitteeName1")
                _lstFld.Add("CommitteeFlg1")
                _lstFld.Add("CommitteeCode2")
                _lstFld.Add("CommitteeName2")
                _lstFld.Add("CommitteeFlg2")
                _lstFld.Add("CommitteeCode3")
                _lstFld.Add("CommitteeName3")
                _lstFld.Add("CommitteeFlg3")
                _lstFld.Add("J_ExecutiveCode")
                _lstFld.Add("J_ExecutiveName")
                _lstFld.Add("J_ExecutiveCode1")
                _lstFld.Add("J_ExecutiveName1")
                _lstFld.Add("J_CommitteeCode")
                _lstFld.Add("J_CommitteeName")
                _lstFld.Add("J_CommitteeFlg")
                _lstFld.Add("J_CommitteeCode1")
                _lstFld.Add("J_CommitteeName1")
                _lstFld.Add("J_CommitteeFlg1")
                _lstFld.Add("J_CommitteeCode2")
                _lstFld.Add("J_CommitteeName2")
                _lstFld.Add("J_CommitteeFlg2")
                _lstFld.Add("J_CommitteeCode3")
                _lstFld.Add("J_CommitteeName3")
                _lstFld.Add("J_CommitteeFlg3")
                '------------------

        End Select
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' キーフィールド名セット
    ''' </summary>
    Protected Overridable Sub SetKeyField(strFldName As String)
        _lstKeyFld.Clear()

        Select Case strFldName
            'Case "A_SETTING"
                '------------------
            Case "A_ACCOUNT"
                '------------------
                _lstKeyFld.Add("Region")
                _lstKeyFld.Add("Zone")
                _lstKeyFld.Add("ClubCode")
                _lstKeyFld.Add("MemberCode")
                '------------------
            Case "A_FILEPATH"
                '------------------
                _lstKeyFld.Add("DataClass")
                '------------------
            Case "A_FILEPATH"
                '------------------
                _lstKeyFld.Add("DataClass")
                '------------------
            Case "T_SLOGAN"
                '------------------
                _lstKeyFld.Add("DataNo")
                '------------------
            Case "T_LETTER"
                '------------------
                _lstKeyFld.Add("DataNo")
                '------------------
            Case "T_EVENTRET"
                '------------------
                _lstKeyFld.Add("DataNo")
                '------------------
            Case "T_EVENT"
                '------------------
                _lstKeyFld.Add("DataNo")
                '------------------
            Case "T_INFOMATION_CABI"
                '------------------
                _lstKeyFld.Add("DataNo")
                '------------------
            Case "T_MAGAZINE"
                '------------------
                _lstKeyFld.Add("DataNo")
                '------------------
            Case "T_MAGAZINEBUY"
                '------------------
                _lstKeyFld.Add("DataNo")
                '------------------
            Case "M_DISTRICTOFFICER"
                '------------------
                _lstKeyFld.Add("DistrictClass")
                _lstKeyFld.Add("DistrictCode")
                '------------------
            Case "M_CABINET"
                '------------------
                _lstKeyFld.Add("FiscalStart")
                _lstKeyFld.Add("FiscalEnd")
                _lstKeyFld.Add("MemberCode")
                _lstKeyFld.Add("ClubCode")
                '------------------
            Case "M_CLUB"
                '------------------
                _lstKeyFld.Add("Region")
                _lstKeyFld.Add("Zone")
                _lstKeyFld.Add("ClubCode")
                '------------------
            Case "T_CLUBSLOGAN"
                '------------------
                _lstKeyFld.Add("DataNo")
                '------------------
            Case "T_MEETINGSCHEDULE"
                '------------------
                _lstKeyFld.Add("DataNo")
                '------------------
            Case "T_DIRECTOR"
                '------------------
                _lstKeyFld.Add("DataNo")
                '------------------
            Case "T_MEETINGPROGRAM"
                '------------------
                _lstKeyFld.Add("DataNo")
                '------------------
            Case "T_INFOMATION_CLUB"
                '------------------
                _lstKeyFld.Add("DataNo")
                '------------------
            Case "M_MEMBER"
                '------------------
                _lstKeyFld.Add("Region")
                _lstKeyFld.Add("Zone")
                _lstKeyFld.Add("ClubCode")
                _lstKeyFld.Add("MemberCode")
                '------------------
        End Select

    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' テーブル項目の定義 (型と桁)
    ''' </summary>
    ''' <param name="strFldName"></param>
    ''' <param name="eType">型(戻り値)</param>
    ''' <param name="intKeta">桁 (戻り値)</param>
    ''' <remarks>項目チェックに使用</remarks>
    Public Sub Field_Difinition(ByVal strFldName As String,
                                   ByRef eType As CmnUtility.eValueType,
                                   ByRef intKeta As Integer)

        intKeta = -1

        'フィールド別に型を決め、SQL文を作成
        Select Case _strTableName
            Case "A_SETTING"
                Select Case strFldName
                    Case "DistrictCode" : eType = CmnUtility.eValueType.String : intKeta = 5
                    Case "DistrictName" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "CabinetName" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "PeriodStart" : eType = CmnUtility.eValueType.String : intKeta = 5
                    Case "PeriodEnd" : eType = CmnUtility.eValueType.String : intKeta = 5
                    Case "DistrictID" : eType = CmnUtility.eValueType.String : intKeta = 5
                    Case "MagazineMoney" : eType = CmnUtility.eValueType.String : intKeta = 4
                    Case "EventDataDay" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "AdminID" : eType = CmnUtility.eValueType.String : intKeta = 5
                    Case "AdminName" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "AdminPassWord" : eType = CmnUtility.eValueType.String : intKeta = 10
                    Case "CabinetEmailAddeess" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "AdminEmailAddress" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "CabinetTelNo" : eType = CmnUtility.eValueType.String : intKeta = 13
                    Case "AdminTelNo" : eType = CmnUtility.eValueType.String : intKeta = 13
                    Case "VersionNo" : eType = CmnUtility.eValueType.String : intKeta = 5
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
            Case "A_DATANO"
                Select Case strFldName
                    Case "DataClass" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "DataName" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "LastDataNo" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
            Case "A_ACCOUNT"
                Select Case strFldName
                    Case "Region" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "Zone" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "ClubCode" : eType = CmnUtility.eValueType.String : intKeta = 6
                    Case "ClubName" : eType = CmnUtility.eValueType.String : intKeta = 80
                    Case "MemberCode" : eType = CmnUtility.eValueType.String : intKeta = 7
                    Case "MemberFirstName" : eType = CmnUtility.eValueType.String : intKeta = 30
                    Case "MemberLastName" : eType = CmnUtility.eValueType.String : intKeta = 30
                    Case "AccountDate" : eType = CmnUtility.eValueType.String : intKeta = 19
                    Case "LastUpdDate" : eType = CmnUtility.eValueType.String : intKeta = 19
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
            Case "A_FILEPATH"
                Select Case strFldName
                    Case "DataClass" : eType = CmnUtility.eValueType.String : intKeta = 3
                    Case "FilePath" : eType = CmnUtility.eValueType.String : intKeta = 100
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
            Case "T_SLOGAN"
                Select Case strFldName
                    Case "DataNo" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "FiscalStart" : eType = CmnUtility.eValueType.String : intKeta = 4
                    Case "FiscalEnd" : eType = CmnUtility.eValueType.String : intKeta = 4
                    Case "Slogan" : eType = CmnUtility.eValueType.String : intKeta = 120
                    Case "DistrictGovernor" : eType = CmnUtility.eValueType.String : intKeta = 40
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
            Case "T_LETTER"
                Select Case strFldName
                    Case "DataNo" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "EventDate" : eType = CmnUtility.eValueType.String : intKeta = 19
                    Case "EventTime" : eType = CmnUtility.eValueType.String : intKeta = 5
                    Case "Title" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "Body" : eType = CmnUtility.eValueType.String : intKeta = 300
                    Case "Image1FileName" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "Image2FileName" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "NoticeFlg" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
            Case "T_EVENTRET"
                Select Case strFldName
                    Case "DataNo" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "EventClass" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "EventDataNo" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "EventDate" : eType = CmnUtility.eValueType.String : intKeta = 19
                    Case "ClubCode" : eType = CmnUtility.eValueType.String : intKeta = 6
                    Case "ClubNameShort" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "MemberCode" : eType = CmnUtility.eValueType.String : intKeta = 7
                    Case "MemberName" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "DirectorName" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "Answer" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "AnswerLate" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "AnswerEarly" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "Online" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "Option1" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "Option2" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "Option3" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "Option4" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "Option5" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "OtherCount" : eType = CmnUtility.eValueType.Number : intKeta = 20
                    Case "TargetFlg" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "CancelFlg" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
            Case "T_EVENT"
                Select Case strFldName
                    Case "DataNo" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "EventDate" : eType = CmnUtility.eValueType.String : intKeta = 10
                    Case "EventTimeStart" : eType = CmnUtility.eValueType.String : intKeta = 5
                    Case "EventTimeEnd" : eType = CmnUtility.eValueType.String : intKeta = 5
                    Case "ReceptionTime" : eType = CmnUtility.eValueType.String : intKeta = 5
                    Case "EventPlace" : eType = CmnUtility.eValueType.String : intKeta = 80
                    Case "Title" : eType = CmnUtility.eValueType.String : intKeta = 80
                    Case "Body" : eType = CmnUtility.eValueType.String : intKeta = 500
                    Case "OptionName1" : eType = CmnUtility.eValueType.String : intKeta = 20
                    Case "OptionRadio1" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "OptionName2" : eType = CmnUtility.eValueType.String : intKeta = 20
                    Case "OptionRadio2" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "OptionName3" : eType = CmnUtility.eValueType.String : intKeta = 20
                    Case "OptionRadio3" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "OptionName4" : eType = CmnUtility.eValueType.String : intKeta = 20
                    Case "OptionRadio4" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "OptionName5" : eType = CmnUtility.eValueType.String : intKeta = 20
                    Case "OptionRadio5" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "Sake" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "Meeting" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "MeetingUrl" : eType = CmnUtility.eValueType.String : intKeta = 150
                    Case "MeetingID" : eType = CmnUtility.eValueType.String : intKeta = 20
                    Case "MeetingPW" : eType = CmnUtility.eValueType.String : intKeta = 10
                    Case "MeetingOther" : eType = CmnUtility.eValueType.String : intKeta = 100
                    Case "AnswerDate" : eType = CmnUtility.eValueType.String : intKeta = 10
                    Case "FileName" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "CabinetUser" : eType = CmnUtility.eValueType.String : intKeta = 2000
                    Case "ClubUser" : eType = CmnUtility.eValueType.String : intKeta = 2000
                    Case "NoticeFlg" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
            Case "T_INFOMATION_CABI"
                Select Case strFldName
                    Case "DataNo" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "AddDate" : eType = CmnUtility.eValueType.String : intKeta = 19
                    Case "Subject" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "Detail" : eType = CmnUtility.eValueType.String : intKeta = 400
                    Case "FileName" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "InfoFlg" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "InfoUser" : eType = CmnUtility.eValueType.String : intKeta = 2000
                    Case "NoticeFlg" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
            Case "T_MAGAZINE"
                Select Case strFldName
                    Case "DataNo" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "SortNo" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "Magazine" : eType = CmnUtility.eValueType.String : intKeta = 40
                    Case "FileName" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "MagazineClass" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "MagazinePrice" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
            Case "T_MAGAZINEBUY"
                Select Case strFldName
                    Case "DataNo" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "MagazineDataNo" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "Magazine" : eType = CmnUtility.eValueType.String : intKeta = 40
                    Case "BuyDate" : eType = CmnUtility.eValueType.String : intKeta = 10
                    Case "BuyNumber" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "MagazinePrice" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "MoneyTotal" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "Region" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "Zone" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "ClubCode" : eType = CmnUtility.eValueType.String : intKeta = 6
                    Case "ClubNameShort" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "MemberCode" : eType = CmnUtility.eValueType.String : intKeta = 7
                    Case "MemberName" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "ShippingDate" : eType = CmnUtility.eValueType.String : intKeta = 10
                    Case "PaymentDate" : eType = CmnUtility.eValueType.String : intKeta = 10
                    Case "Payment" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "DelFlg" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
            Case "M_DISTRICTOFFICER"
                Select Case strFldName
                    Case "DistrictClass" : eType = CmnUtility.eValueType.Number : intKeta = 1
                    Case "DistrictCode" : eType = CmnUtility.eValueType.Number : intKeta = 2
                    Case "DistrictName" : eType = CmnUtility.eValueType.String : intKeta = 80
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
            Case "M_CABINET"
                Select Case strFldName
                    Case "FiscalStart" : eType = CmnUtility.eValueType.String : intKeta = 4
                    Case "FiscalEnd" : eType = CmnUtility.eValueType.String : intKeta = 4
                    Case "MemberCode" : eType = CmnUtility.eValueType.String : intKeta = 7
                    Case "MemberName" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "ClubCode" : eType = CmnUtility.eValueType.String : intKeta = 6
                    Case "ClubNameShort" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "DistrictClass" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "DistrictCode" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "DistrictName" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "DistrictClass1" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "DistrictCode1" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "DistrictName1" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "DistrictClass2" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "DistrictCode2" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "DistrictName2" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "DistrictClass3" : eType = CmnUtility.eValueType.String : intKeta = 10
                    Case "DistrictCode3" : eType = CmnUtility.eValueType.String : intKeta = 20
                    Case "DistrictName3" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "DistrictClass4" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "DistrictCode4" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "DistrictName4" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "DistrictClass5" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "DistrictCode5" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "DistrictName5" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
            Case "M_CLUB"
                Select Case strFldName
                    Case "Region" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "Zone" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "Sort" : eType = CmnUtility.eValueType.Number : intKeta = 1
                    Case "ClubCode" : eType = CmnUtility.eValueType.String : intKeta = 6
                    Case "ClubName" : eType = CmnUtility.eValueType.String : intKeta = 80
                    Case "ClubNameShort" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "PassWord" : eType = CmnUtility.eValueType.String : intKeta = 8
                    Case "FormationDate" : eType = CmnUtility.eValueType.String : intKeta = 10
                    Case "CharterNight" : eType = CmnUtility.eValueType.String : intKeta = 10
                    Case "MeetingDate" : eType = CmnUtility.eValueType.String : intKeta = 40
                    Case "MeetingTime" : eType = CmnUtility.eValueType.String : intKeta = 5
                    Case "MeetingPlace" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "AnswerDay" : eType = CmnUtility.eValueType.Number : intKeta = 1
                    Case "AnswerTime" : eType = CmnUtility.eValueType.String : intKeta = 5
                    Case "Zip" : eType = CmnUtility.eValueType.String : intKeta = 8
                    Case "Address" : eType = CmnUtility.eValueType.String : intKeta = 150
                    Case "Tel" : eType = CmnUtility.eValueType.String : intKeta = 15
                    Case "HP" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "MailAddress" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
            Case "T_CLUBSLOGAN"
                Select Case strFldName
                    Case "DataNo" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "ClubCode" : eType = CmnUtility.eValueType.String : intKeta = 6
                    Case "ClubNameShort" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "FiscalStart" : eType = CmnUtility.eValueType.String : intKeta = 4
                    Case "FiscalEnd" : eType = CmnUtility.eValueType.String : intKeta = 4
                    Case "ClubSlogan" : eType = CmnUtility.eValueType.String : intKeta = 120
                    Case "ExecutiveName" : eType = CmnUtility.eValueType.String : intKeta = 40
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
            Case "T_MEETINGSCHEDULE"
                Select Case strFldName
                    Case "DataNo" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "ClubCode" : eType = CmnUtility.eValueType.String : intKeta = 6
                    Case "ClubNameShort" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "Fiscal" : eType = CmnUtility.eValueType.String : intKeta = 4
                    Case "MeetingDate" : eType = CmnUtility.eValueType.String : intKeta = 10
                    Case "MeetingTime" : eType = CmnUtility.eValueType.String : intKeta = 5
                    Case "MeetingPlace" : eType = CmnUtility.eValueType.String : intKeta = 80
                    Case "MeetingCount" : eType = CmnUtility.eValueType.Number : intKeta = 3
                    Case "MeetingName" : eType = CmnUtility.eValueType.String : intKeta = 80
                    Case "Online" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "AnswerDate" : eType = CmnUtility.eValueType.String : intKeta = 10
                    Case "RemarksItems" : eType = CmnUtility.eValueType.String : intKeta = 310
                    Case "RemarksCheck" : eType = CmnUtility.eValueType.String : intKeta = 30
                    Case "Remarks" : eType = CmnUtility.eValueType.String : intKeta = 180
                    Case "OptionName1" : eType = CmnUtility.eValueType.String : intKeta = 20
                    Case "OptionRadio1" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "OptionName2" : eType = CmnUtility.eValueType.String : intKeta = 20
                    Case "OptionRadio2" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "OptionName3" : eType = CmnUtility.eValueType.String : intKeta = 20
                    Case "OptionRadio3" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "OptionName4" : eType = CmnUtility.eValueType.String : intKeta = 20
                    Case "OptionRadio4" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "OptionName5" : eType = CmnUtility.eValueType.String : intKeta = 20
                    Case "OptionRadio5" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "Sake" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "OtherUser" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "CancelFlg" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
            Case "T_DIRECTOR"
                Select Case strFldName
                    Case "DataNo" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "ClubCode" : eType = CmnUtility.eValueType.String : intKeta = 6
                    Case "ClubNameShort" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "EventClass" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "CommitteeCode" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "CommitteeName" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "Season" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "EventDate" : eType = CmnUtility.eValueType.String : intKeta = 10
                    Case "EventTime" : eType = CmnUtility.eValueType.String : intKeta = 5
                    Case "EventPlace" : eType = CmnUtility.eValueType.String : intKeta = 80
                    Case "Subject" : eType = CmnUtility.eValueType.String : intKeta = 80
                    Case "Agenda" : eType = CmnUtility.eValueType.String : intKeta = 1600
                    Case "Member" : eType = CmnUtility.eValueType.String : intKeta = 2000
                    Case "MemberAdd" : eType = CmnUtility.eValueType.String : intKeta = 2000
                    Case "AnswerDate" : eType = CmnUtility.eValueType.String : intKeta = 10
                    Case "CancelFlg" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "NoticeFlg" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
            Case "T_MEETINGPROGRAM"
                Select Case strFldName
                    Case "DataNo" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "ScheduleDataNo" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "ClubCode" : eType = CmnUtility.eValueType.String : intKeta = 6
                    Case "ClubNameShort" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "Fiscal" : eType = CmnUtility.eValueType.String : intKeta = 4
                    Case "FileName" : eType = CmnUtility.eValueType.String : intKeta = 80
                    Case "FileName1" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "FileName2" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "FileName3" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "FileName4" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "FileName5" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "Meeting" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "MeetingUrl" : eType = CmnUtility.eValueType.String : intKeta = 150
                    Case "MeetingID" : eType = CmnUtility.eValueType.String : intKeta = 20
                    Case "MeetingPW" : eType = CmnUtility.eValueType.String : intKeta = 10
                    Case "MeetingOther" : eType = CmnUtility.eValueType.String : intKeta = 100
                    Case "NoticeFlg" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
            Case "T_INFOMATION_CLUB"
                Select Case strFldName
                    Case "DataNo" : eType = CmnUtility.eValueType.Number : intKeta = 10
                    Case "ClubCode" : eType = CmnUtility.eValueType.String : intKeta = 6
                    Case "ClubNameShort" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "AddDate" : eType = CmnUtility.eValueType.String : intKeta = 19
                    Case "Subject" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "Detail" : eType = CmnUtility.eValueType.String : intKeta = 1000
                    Case "FileName" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "InfoFlg" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "TypeCode" : eType = CmnUtility.eValueType.String : intKeta = 40
                    Case "InfoUser" : eType = CmnUtility.eValueType.String : intKeta = 2000
                    Case "NoticeFlg" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select
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
                    Case "TypeCode" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "TypeName" : eType = CmnUtility.eValueType.String : intKeta = 20
                    Case "Sex" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "Tel" : eType = CmnUtility.eValueType.String : intKeta = 11
                    Case "JoinDate" : eType = CmnUtility.eValueType.String : intKeta = 10
                    Case "LeaveDate" : eType = CmnUtility.eValueType.String : intKeta = 10
                    Case "LeaveFlg" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "AccountDate" : eType = CmnUtility.eValueType.String : intKeta = 10
                    Case "ExecutiveCode" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "ExecutiveName" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "ExecutiveCode1" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "ExecutiveName1" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "CommitteeCode" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "CommitteeName" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "CommitteeFlg" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "CommitteeCode1" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "CommitteeName1" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "CommitteeFlg1" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "CommitteeCode2" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "CommitteeName2" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "CommitteeFlg2" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "CommitteeCode3" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "CommitteeName3" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "CommitteeFlg3" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "J_ExecutiveCode" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "J_ExecutiveName" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "J_ExecutiveCode1" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "J_ExecutiveName1" : eType = CmnUtility.eValueType.String : intKeta = 50
                    Case "J_CommitteeCode" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "J_CommitteeName" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "J_CommitteeFlg" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "J_CommitteeCode1" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "J_CommitteeName1" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "J_CommitteeFlg1" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "J_CommitteeCode2" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "J_CommitteeName2" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "J_CommitteeFlg2" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case "J_CommitteeCode3" : eType = CmnUtility.eValueType.String : intKeta = 2
                    Case "J_CommitteeName3" : eType = CmnUtility.eValueType.String : intKeta = 60
                    Case "J_CommitteeFlg3" : eType = CmnUtility.eValueType.String : intKeta = 1
                    Case Else
                        Throw New Exception(_strTableName & "/" & strFldName & ":定義がありません(Field_Difinitionメソッド)")
                End Select

        End Select

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' フィールド文字列(項目名/値) GetFldValueString 引数用
    ''' </summary>
    Public Enum FldValueStringType
        ValueOnly '値のみ
        KeyValueOnly 'キー値のみ
        SQLsrv1   'SQL srvマージ用
        SQLsrv2   'SQL srvインサート用
        SQL_Where 'Where句
        SQLite1   'SQLiteマージ用
    End Enum


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' 項目別SQL文の部分作成 SQL ServerのMerge用
    ''' </summary>
    ''' <param name="strFldName"></param>
    ''' <param name="strFldValue"></param>
    ''' <returns></returns>
    Protected Overridable Function FieldSQL_srv(strFldName As String, strFldValue As String) As String

        Dim strRet As String = String.Empty

        ' "'xxx' AS ShainName "の形式

        '通常
        strRet = FieldSQL(strFldName, strFldValue) & " AS " & strFldName & " "

        Return strRet

    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' 項目別SQL文の部分作成 SQLiteのMerge用
    ''' </summary>
    ''' <param name="strFldName"></param>
    ''' <param name="strFldValue"></param>
    ''' <returns></returns>
    Protected Overridable Function FieldSQL_Lite(strFldName As String, strFldValue As String) As String

        Dim strRet As String = String.Empty

        ' "'xxx' "の形式

        '通常
        strRet = FieldSQL(strFldName, strFldValue)

        Return strRet

    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' 項目別SQLの値 (SQLsrv/SQLite共通)
    ''' </summary>
    ''' <param name="strFldName"></param>
    ''' <param name="strFldValue"></param>
    ''' <returns></returns>
    Protected Overridable Function FieldSQL(strFldName As String, strFldValue As String) As String

        Dim strRet As String = String.Empty
        Dim eType As CmnUtility.eValueType = CmnUtility.eValueType.String
        Dim bolBlankZero As Boolean = False
        Dim intKeta As Integer = -1

        'フィールド別に型・桁を取得
        Field_Difinition(strFldName, eType, intKeta)

        'SQL文を作成
        strRet = _utl.ToSql2(strFldValue, eType, bolBlankZero)

        Return strRet

    End Function

End Class
