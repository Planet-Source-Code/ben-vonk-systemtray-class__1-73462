VERSION 5.00
Begin VB.Form frmSystemTray 
   Caption         =   "SystemTray Demo"
   ClientHeight    =   2664
   ClientLeft      =   132
   ClientTop       =   516
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2664
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSystemMenu 
      Caption         =   "&Use SystemMenu"
      Height          =   312
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Value           =   1  'Checked
      Width           =   2892
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3360
      Top             =   2280
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   372
      Left            =   1320
      TabIndex        =   1
      Top             =   2160
      Width           =   972
   End
   Begin VB.TextBox txtDemo 
      Height          =   1332
      Left            =   360
      TabIndex        =   0
      Text            =   "Press the minimize button!"
      Top             =   600
      Width           =   2892
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuFile 
         Caption         =   "Open"
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Close"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmSystemTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SystemTray Form (Demo for the SystemTray Class)
'
'Author Ben Vonk
'23-09-2010 First version
'27-09-2010 Second version Add Balloon events
'02-10-2010 Third version Add Balloon timer and fixed some bugs
'06-11-2010 Fourth version fixed some bugs
'09-11-2010 Fifth version Add hWnd function and make some changes
'06-12-2010 Sixth version Add ReceivedData event

Option Explicit

' Private Class with Events
Private WithEvents SystemTray As clsSystemTray
Attribute SystemTray.VB_VarHelpID = -1

' Private Variable
Private BalloonIsShowed       As Boolean

' Private API
Private Declare Function SetForegroundWindow Lib "User32" (ByVal hWnd As Long) As Long

Private Sub cmdStop_Click()

   Unload Me

End Sub

Private Sub Form_Resize()

   If WindowState = vbNormal Then
      If Not SystemTray Is Nothing Then
         Call SystemTray.DeleteIcon
         
         Set SystemTray = Nothing
      End If
      
   ElseIf WindowState = vbMinimized Then
      Set SystemTray = New clsSystemTray
      
      With SystemTray
         Visible = False
         .Icon = Icon.Handle
         .Menu = hWnd And (chkSystemMenu.Value = vbChecked)
         .Parent = hWnd
         .TipText = Caption
         
         Call .AddIcon
         
         If .Enabled Then Call .ShowBalloon("Hello", "I'am on the SystemTray now!", , 5000, True)
      End With
   End If

End Sub

Private Sub mnuFile_Click(Index As Integer)

   If Index Then
      Unload Me
      
   Else
      Call SystemTray_DblClick(vbLeftButton)
   End If

End Sub

Private Sub ShowPopupMenu()

   SetForegroundWindow hWnd
   PopupMenu mnuMenu

End Sub

Private Sub SystemTray_BalloonClick()

   MsgBox "Balloon is clicked!"

End Sub

Private Sub SystemTray_BalloonClose()

   MsgBox "Balloon is closed!"

End Sub

Private Sub SystemTray_BalloonHide()

   MsgBox "Balloon is hide!"

End Sub

Private Sub SystemTray_BalloonShow()

   MsgBox "Balloon is shown!"

End Sub

Private Sub SystemTray_BalloonTimeOut()

   MsgBox "Balloon is timed out!"

End Sub

Private Sub SystemTray_Click(Button As Integer)

   If Button = vbLeftButton Then Call ShowPopupMenu

End Sub

Private Sub SystemTray_DblClick(Button As Integer)

   If Button = vbLeftButton Then
      WindowState = vbNormal
      Visible = True
      SetForegroundWindow hWnd
      
   ElseIf Button = vbRightButton Then
      Call ShowPopupMenu
   End If

End Sub

Private Sub SystemTray_MouseMove()

   Call SystemTray.HideBalloon

End Sub
