VERSION 5.00
Begin VB.Form frmReceiver 
   Caption         =   "Receiver"
   ClientHeight    =   2400
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmText 
      Caption         =   "Incoming text"
      Height          =   1452
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3492
      Begin VB.Label lblText 
         Caption         =   """"""
         Height          =   1092
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3252
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   372
      Left            =   1200
      TabIndex        =   0
      Top             =   1800
      Width           =   1332
   End
End
Attribute VB_Name = "frmReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private WithEvents
Private WithEvents SystemTray As clsSystemTray
Attribute SystemTray.VB_VarHelpID = -1

' Private API
Private Declare Function SetForegroundWindow Lib "User32" (ByVal hWnd As Long) As Long

Public Sub ReceivedData(ByVal Data As String)

   If Data = Chr(1) & "QUIT" Then
      Unload Me
      
   Else
      lblText.Caption = Data & " - Visible = " & Visible
      
      If Not Visible Then
         Call OpenFromSystemTray
         
      Else
         SetForegroundWindow hWnd
      End If
   End If

End Sub

Private Sub OpenFromSystemTray()

   Call HookSystemTray(hWnd)
   
   WindowState = vbNormal
   Visible = True
   SystemTray.DeleteIcon
   Set SystemTray = Nothing
   SetForegroundWindow hWnd

End Sub

Private Sub cmdStop_Click()

   Unload Me

End Sub

Private Sub Form_Load()

   Call HookSystemTray(hWnd)

End Sub

Private Sub Form_Resize()

   If WindowState = vbMinimized Then
      Call HookSystemTray(hWnd)
      
      Set SystemTray = New clsSystemTray
      
      With SystemTray
         .Icon = Icon.Handle
         .Menu = hWnd
         .Parent = hWnd
         .TipText = Caption
         
         Call .AddIcon
         
         If .Enabled Then
            Call .ShowBalloon("Receiver", "Click on me to open", , 4000)
            
            Visible = False
         End If
      End With
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set SystemTray = Nothing
   Unload Me

End Sub

Private Sub SystemTray_Click(Button As Integer)

   If Button = vbLeftButton Then Call OpenFromSystemTray

End Sub

Private Sub SystemTray_ReceivedData(Data As String)

   Call ReceivedData(Data)

End Sub
