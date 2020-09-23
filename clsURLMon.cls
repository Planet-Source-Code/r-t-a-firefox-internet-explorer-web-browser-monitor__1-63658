VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsURLMon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents myTimer As CTimer
Attribute myTimer.VB_VarHelpID = -1
Private txtMon As Object
Public LastURL As String

Public Active As Boolean
Public MonitorIE As Boolean
Public MonitorFireFox As Boolean

Public Event TitleChanged(Title As String)
Public Event URLChanged(URL As String)

Private WithEvents MTimer As timer
Attribute MTimer.VB_VarHelpID = -1

Private Sub LinkTextBox(ByRef pTextBox As TextBox, plinkTopic As String, plinkMode As LinkModeConstants, plinkItem As Variant)
On Local Error GoTo listen:
   pTextBox.LinkTopic = plinkTopic
   pTextBox.LinkTimeout = 10
   pTextBox.LinkMode = plinkMode
   pTextBox.LinkItem = plinkItem
   pTextBox.LinkRequest
listen:
   On Error GoTo 0
End Sub

Public Sub GetDDEURL(LinkTopic As String)
Dim sTmp As String, sLink() As String
    Call LinkTextBox(txtMon, LinkTopic, vbLinkManual, &HFFFFFFFF)
    
    sTmp = Mid(txtMon.Text, 2)
    sLink = Split(sTmp, """,""")
    
    If UBound(sLink) > 0 Then
        If Left(sLink(0), 1) = """" Then
            sTmp = Mid(sLink(0), 2, Len(sLink(0)) - 1)
        Else
            sTmp = sLink(0)
        End If
        
        If sTmp <> LastURL Then
            RaiseEvent URLChanged(sTmp)
            LastURL = sTmp
            
            If sLink(1) <> "" Then
                If Right(sLink(1), 1) = """" Then
                    sTmp = Left(sLink(1), Len(sLink(1)) - 1)
                Else
                    sTmp = sLink(1)
                End If
                    
                RaiseEvent TitleChanged(sTmp)
            End If
        End If
    End If
End Sub

Private Sub myTimer_ThatTime()
    If MonitorFireFox Then GetDDEURL "firefox|WWW_GetWindowInfo"
    If MonitorIE Then GetDDEURL "iexplore|WWW_GetWindowInfo"
End Sub

Public Sub StartMon(TextBoxCtl As Object)
    Set txtMon = TextBoxCtl
    Set myTimer = New CTimer
    myTimer.Interval = 500
    Active = True
End Sub

Public Sub StopMon()
    Set myTimer = Nothing
    Set myTimer = Nothing
    Active = False
End Sub
