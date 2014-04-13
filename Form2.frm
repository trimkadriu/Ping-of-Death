VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "About Us !!!"
   ClientHeight    =   4635
   ClientLeft      =   6015
   ClientTop       =   4185
   ClientWidth     =   5055
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":08CA
   ScaleHeight     =   4635
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image2 
      Height          =   315
      Left            =   4720
      Picture         =   "Form2.frx":24841
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Top             =   4200
      Width           =   5055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, _
 ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Sub ReleaseCapture Lib "User32" ()
Private Declare Function ShellExecute Lib _
"shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Private Const SW_SHOW = 1
Private Sub webpage()
lret = Shell("cmd.exe", vbNormalNoFocus)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const WM_NCLBUTTONDOWN = &HA1
    Const HTCAPTION = 2
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub Image1_Click()
 Navigate ("http://www.microstar-studio.org")
End Sub
Private Sub Image2_Click()
Unload Form2
form1.Enabled = True
form1.Show
End Sub

Public Sub Navigate(ByVal NavTo As String)

Dim hBrowse As Long
hBrowse = ShellExecute(0&, "open", NavTo, "", "", SW_SHOW)
End Sub

