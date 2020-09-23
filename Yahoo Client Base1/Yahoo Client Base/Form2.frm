VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8355
   LinkTopic       =   "Form2"
   ScaleHeight     =   6645
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      DragMode        =   1  'Automatic
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      ExtentX         =   14631
      ExtentY         =   11668
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
WebBrowser1.Width = Me.Width
WebBrowser1.Height = Me.Height
End Sub
