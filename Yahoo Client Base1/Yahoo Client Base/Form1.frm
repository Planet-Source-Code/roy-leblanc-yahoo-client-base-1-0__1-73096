VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yahoo Client Base"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6000
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0552
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wSock 
      Left            =   7800
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton Command1 
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Idle..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Activity Logger"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   2895
      Begin RichTextLib.RichTextBox rtb_Activity 
         Height          =   1935
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   3413
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         DisableNoScroll =   -1  'True
         Appearance      =   0
         TextRTF         =   $"Form1.frx":0664
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Buddy List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   3120
      TabIndex        =   7
      Top             =   120
      Width           =   2655
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   3735
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   6588
         _Version        =   393217
         Indentation     =   3
         LineStyle       =   1
         Style           =   1
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Data Logger"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   5655
      Begin VB.ListBox lst_DataLog 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "http://eliteroy.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5880
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Yahoo Client Base Example 1.0
'Written By : EliteRoy(Roy LeBlanc)
'My Website : http://eliteroy.com
'Credits Go To -
'C-4 BuddyList Parse Method
'Expulsion, Adam, and Dubee For Login Method

Private Sub Command1_Click()
If Command1.Caption = "Connect" Then
If wSock.State <> 0 Then Exit Sub
U_name = Text1.Text
P_word = Text2.Text
Got_Cookies = False
wSock.Connect "login.yahoo.com", 80
ElseIf Command1.Caption = "Disconnect" Then
wSock.Close
Y_Cook = Empty
T_Cook = Empty
U_name = Empty
P_word = Empty
Label3.Caption = "Idle..."
Command1.Caption = "Connect"
End If
End Sub

Private Sub Form_Load()
rtb_Activity.Text = Empty
End Sub

Private Sub Label4_Click()
Form2.Show
Form2.WebBrowser1.Navigate "http://eliteroy.com"
End Sub

Private Sub wSock_Connect()
If Got_Cookies = False Then
    wSock.SendData ycHeader(U_name, P_word)
    Debug.Print ycHeader(U_name, P_word)
    lst_DataLog.AddItem ycHeader(U_name, P_word)
ElseIf Got_Cookies = True Then
    wSock.SendData YLogin(U_name, YCook, TCook)
    Debug.Print "YMSG Send -----"; YLogin(U_name, YCook, TCook)
    lst_DataLog.AddItem YLogin(U_name, YCook, TCook)
End If
End Sub

Private Sub wSock_DataArrival(ByVal bytesTotal As Long)
wSock.GetData wsData, vbTextCompare
If Got_Cookies = False Then
    Call CookieDataHandle(wsData, wSock)
ElseIf Got_Cookies = True Then
    Call YMSGDataHandle(wsData, wSock)
End If
lst_DataLog.AddItem wsData
End Sub

Private Sub wSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'Displays A Message Box With The Error Number And Description In It
'MsgBox Number & ": " & Description
ActivityLogging Number & ": " & Description
wSock.Close
End Sub
