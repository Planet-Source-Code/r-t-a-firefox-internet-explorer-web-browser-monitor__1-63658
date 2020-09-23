VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "URLMon"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstURL 
      Height          =   2580
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   7935
   End
   Begin VB.CheckBox Chk 
      Caption         =   "Internet Explorer"
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   9
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CheckBox Chk 
      Caption         =   "Firefox"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdMon 
      Caption         =   "Start Monitor"
      Height          =   975
      Left            =   6960
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtMonitor 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   5775
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   5775
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   5775
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      Caption         =   "Victims"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      Caption         =   "Monitor"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      Caption         =   "URL"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      Caption         =   "Title"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents URLMon As clsURLMon
Attribute URLMon.VB_VarHelpID = -1

Private Sub Chk_Click(Index As Integer)
    Select Case Index
        Case 0 ' firefox
            URLMon.MonitorFireFox = Chk(0).Value
        Case 1 ' internet explorer
            URLMon.MonitorIE = Chk(1).Value
    End Select
End Sub

Private Sub cmdMon_Click()
    If URLMon.Active Then
        URLMon.StopMon
        cmdMon.Caption = "Start Monitor"
    Else
        URLMon.StartMon txtMonitor
        cmdMon.Caption = "Stop Monitor"
    End If
End Sub

Private Sub Form_Load()
    Set URLMon = New clsURLMon
    Chk(0).Value = 1
    URLMon.MonitorFireFox = True
    Chk(1).Value = 1
    URLMon.MonitorIE = True
End Sub

Private Sub URLMon_TitleChanged(Title As String)
    txtTitle = Title
End Sub

Private Sub URLMon_URLChanged(URL As String)
    txtURL = URL
    lstURL.AddItem URL
    lstURL.ListIndex = lstURL.ListCount - 1
End Sub
