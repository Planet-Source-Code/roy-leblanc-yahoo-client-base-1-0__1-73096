Attribute VB_Name = "Login"
'Yahoo Client Base Example 1.0
'Written By : EliteRoy(Roy LeBlanc)
'My Website : http://eliteroy.com
'Credits Go To -
'C-4 BuddyList Parse Method
'Expulsion, Adam, and Dubee For Login Method
Option Explicit

Public Got_Cookies As Boolean
Public U_name As String
Public P_word As String
Public YCook As String
Public TCook As String
Public wsData As String
Public MyName As String

Private Function Header(ByVal StrPacketType As String, ByVal StrStat As String, ByVal StrSession As String, ByVal StrComm As Long) As String
    Dim Version As String
    Version = 15
    Header = "YMSG" & Chr(Int(Version / 256)) & Chr(Int(Version Mod 256)) & String(2, Chr(0)) & Chr(Int(Len(StrPacketType) / 256)) & Chr(Int(Len(StrPacketType) Mod 256)) & Chr(Int(StrComm / 256)) & Chr(Int(StrComm Mod 256)) & Mid(StrStat, 1, 4) & Mid(StrSession, 1, 4) & StrPacketType
End Function

Public Function YLogin(YahooID As String, YCookie As String, TCookie As String)
    YLogin = Header("0��" & YahooID & "��2��" & YahooID & "��1��" & YahooID & "��244��1��6��" & YCookie & " " & TCookie & "��98��us��", String(4, Chr(0)), String(4, Chr(0)), 550)
End Function

Public Function ycHeader(MyID As String, MyPass As String) As String
    Dim LoginYahoo As String
    LoginYahoo = "GET http://login.yahoo.com/config/login?login=" & MyID & "&passwd=" & MyPass & " HTTP/1.1" & vbCrLf
    LoginYahoo = LoginYahoo & "Accept-Language: en-us" & vbCrLf
    LoginYahoo = LoginYahoo & "User-Agent: Mozilla/5.0 (compatible; MSIE 8.0; Windows NT 5.1; Expulsion-Creations; Elite-Roy)" & vbCrLf
    LoginYahoo = LoginYahoo & "Accept: */*" & vbCrLf
    LoginYahoo = LoginYahoo & "Host: login.yahoo.com" & vbCrLf
    LoginYahoo = LoginYahoo & "Connection: Keep-Alive" & vbCrLf & vbCrLf
    ycHeader = LoginYahoo
End Function

Public Sub ActivityLogging(Act As String)
Form1.rtb_Activity.Text = Form1.rtb_Activity.Text + Act + vbNewLine
End Sub

Public Sub CookieDataHandle(Data As String, Winsock As Winsock)
If InStr(Data, "302 Found") Then
        YCook = Split(Data, "Y=")(1)
        YCook = Split(YCook, "np=1")(0)
        YCook = "Y=" & YCook & "np=1;"
        TCook = Split(Data, "T=")(1)
        TCook = Split(TCook, ";")(0)
        TCook = "T=" & TCook
        Debug.Print YCook
        Debug.Print TCook
    Got_Cookies = True
    Winsock.Close
    Winsock.Connect "scs.msg.yahoo.com", 5050
ElseIf InStr(Data, "Yahoo! - 400 Bad Request") Then
    MsgBox "Invalid ID!"
    Winsock.Close
Else
    MsgBox "Invalid Username/Password Combo"
    Winsock.Close
End If
Debug.Print Data
End Sub

Public Sub YMSGDataHandle(Data As String, Winsock As Winsock)
On Error Resume Next
Select Case Asc(Mid(Data, 12, 1))
    Case 2
        'Handles Multiple Things Like Informing the Client Of A Closed Connection By The Server,
        'And Alerting The Client To A User Going Offline
        If InStr(Data, "����") Then
            Form1.Label3.Caption = "Connection Closed By Server"
            Winsock.Close
        ElseIf InStr(Data, "7��") Then
            Dim TempRemoteUser As String
            TempRemoteUser = Split(Data, "7��")(1)
            TempRemoteUser = Split(TempRemoteUser, "��244��")(0)
            ActivityLogging TempRemoteUser & " Is Now Offline"
            TempRemoteUser = Empty
        End If
    Case 6
        'PM Recieved
        '4��vb697��5��vb_deadly��14��dfdfdsfdsf��63��;
        Dim tmpUser As String, tmpMsg As String
        tmpUser = Split(Data, "4��")(1)
        tmpUser = Split(tmpUser, "��5")(0)
        tmpMsg = Split(Data, "��14��")(1)
        tmpMsg = Split(tmpMsg, "��63")(0)
        ActivityLogging tmpUser & ": " & tmpMsg
        tmpUser = Empty
        tmpMsg = Empty
    Case 75
        'User Typing
        '4��vb697��5��vb_deadly��13��0��14�� ��49��TYPING��
        Dim TmpTypeStr As String
        TmpTypeStr = Split(Data, "��5��")(0)
        TmpTypeStr = Split(TmpTypeStr, "4��")(1)
        'TmpTypeStr = Replace(TmpTypeStr, "E", Empty)
        ActivityLogging TmpTypeStr & " is Typing"
        
    Case 85
        'Connection Successful
        Form1.Label3.Caption = "Connected"
        Form1.Command1.Caption = "Disconnect"
        MyName = Split(wsData, "��213��2��216��")(1)
        MyName = Split(MyName, "��281��")(0)
        MyName = Replace(MyName, "��254��", " ")
        ActivityLogging "Welcome : " & MyName
    Case 198
        'User Changed Their Status
        Dim tmprUser, tmpStat As String
        tmprUser = Split(Data, "7��")(1)
        tmprUser = Split(tmprUser, "��10��")(0)
        tmpStat = Split(Data, "��19��")(1)
        tmpStat = Split(tmpStat, "��47��")(0)
        ActivityLogging tmpUser & " Changed Their Status To " & tmpStat
        tmpUser = Empty
        tmpStat = Empty
    Case 240
        'User Came Online
        Dim tmpStr As String
            tmpStr = Split(Data, "��315��7��")(1)
            tmpStr = Split(tmpStr, "��10")(0)
            ActivityLogging tmpStr & " is Now Online"
            tmpStr = Empty
    Case 241
        'Buddy List
        ParseBuddyList Data, Form1.TreeView1
End Select
Debug.Print Asc(Mid(Data, 12, 1)) & "---------" & Data
End Sub
