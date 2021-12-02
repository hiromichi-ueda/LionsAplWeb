'23456789+123456789+123456789+123456789+123456789+123456789+123456789+123456789+123456789+123456789+
'---------------------------------------------------------------------------------------------------
'
' AccountData
'
' Created by INSAT on 2021/03/30.
' Copyright (c) 2021年 INSAT. All right reserved.
'
'---------------------------------------------------------------------------------------------------
Public Class DBAccountSet
    Public Property Data As List(Of DataAccountSet)
End Class

Public Class DataAccountSet
    Public Property Region As String                ' リジョン
    Public Property Zone As String                  ' ゾーン
    Public Property ClubCode As String              ' クラブコード
    Public Property ClubNameShort As String         ' クラブ名（略称）
    Public Property MemberCode As String            ' 会員番号
    Public Property MemberFirstName As String       ' 会員名（性）
    Public Property MemberLastName As String        ' 会員名（名）
    Public Property MemberNameKana As String        ' 会員名（カナ）

End Class
