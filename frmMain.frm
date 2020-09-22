VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4575
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0CCA
   ScaleHeight     =   2670
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   2040
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   0
      Picture         =   "frmMain.frx":6B40
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   13
      Top             =   0
      Width           =   230
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   1950
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox txtMain 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1530
      Width           =   2805
   End
   Begin VB.TextBox txtMain 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1140
      Width           =   2805
   End
   Begin VB.TextBox txtMain 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   750
      Width           =   2805
   End
   Begin VB.TextBox txtMain 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   2805
   End
   Begin VB.Timer timMain 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4320
      Top             =   360
   End
   Begin VB.CommandButton cmdS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Start"
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   4365
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "WOH Spy by Wacko"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Stay On Top"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   11
      Top             =   1980
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Stay On Top"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1980
      Width           =   1935
   End
   Begin VB.Label lblMain 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Window Handle  :"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   420
      Width           =   1305
   End
   Begin VB.Label lblMain 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Window Caption :"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   810
      Width           =   1305
   End
   Begin VB.Label lblMain 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Window Parent   :"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1305
   End
   Begin VB.Label lblMain 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Window Class     :"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1590
      Width           =   1305
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
       Dim s As Integer
       Dim dta As String
Const HTCAPTION = 2
Const WM_NCLBUTTONDOWN = &HA1

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
       "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
       String, ByVal lpFile As String, ByVal lpParameters As String, _
       ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public LastState As Integer

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const SC_RESTORE = &HF120&
Private Const SC_SIZE = &HF000&
Public Sub TrayIconCallback(Msg As Long)
   If Msg = WM_LBUTTONDBLCLK Then
      Me.Visible = True
      Me.WindowState = vbNormal
   End If
End Sub
Private Sub mnuTray_Click()
If frmMain.Visible = "false" Then
frmMain.Visible = "true"
Else
frmMain.Visible = "false"
frmMain.Visible = "true"
End If
End Sub
Private Sub cmdMain_Click()

End Sub

Private Sub about_Click()

End Sub

Private Sub Check1_Click()
If Check1.Value = "1" Then
FormOnTop frmMain
Else
FormNotOnTop frmMain
End If
End Sub

Private Sub cmdS_Click()
'Just for enabling and disabling the timer

If cmdS.Caption = "&Start" Then
    timMain.Enabled = True
    cmdS.Caption = "&Stop"
Else
    cmdS.Caption = "&Start"
    Screen.MousePointer = vbDefault
    timMain.Enabled = False
End If

End Sub


Private Sub frMain_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command6_Click()

End Sub

Private Sub Command7_Click()

End Sub


Private Sub file_Click()

End Sub

Private Sub Form_Load()
FormOnTop frmMain
   frmTrayIcon.SetCallback Me

   frmTrayIcon.Update True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
              If Button = 1 Then ' Checking for Left Button only
                     Dim ReturnVal As Long
                     X = ReleaseCapture()
                     ReturnVal = SendMessage2(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
              End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmTrayIcon.Update False
   End
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Label3_Click()
End
End Sub

Private Sub Label4_Click()
frmMain.Hide
End Sub





Private Sub rm_Click()

End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
              If Button = 1 Then ' Checking for Left Button only
                     Dim ReturnVal As Long
                     X = ReleaseCapture()
                     ReturnVal = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
              End If
End Sub

Private Sub Picture1_Click()
    On Error GoTo fileOpenErrr
       CDialog.CancelError = True
       CDialog.FLAGS = &H4& Or &H100&
       CDialog.DefaultExt = ".jpg"
       CDialog.DialogTitle = "Select File To Open"
       CDialog.Filter = "JPEG (*.jpg)|*.jpg|GIF (*.gif)|*.gif|BITMAP (*.bmp)|*.bmp"
       CDialog.ShowOpen
Set frmMain.Picture = LoadPicture(CDialog.filename)
fileOpenErrr:
       Exit Sub
End Sub

Private Sub timMain_Timer()

Dim P As POINTAPI

Dim hWn As Long

Dim WinCap As String * 255
Dim ClName As String * 255

Dim OldParent As Long, Parent As Long

'First, get the cursor position of mouse
GetCursorPos P


'WindowFromPoint returns the handle of the window under the mouse

hWn = WindowFromPoint(P.X, P.Y)
txtMain(0).Text = hWn


'Determine the caption, using the handle we obtained above

GetWindowText hWn, WinCap, 254
txtMain(1).Text = WinCap
If Trim(txtMain(1).Text) = "" Then txtMain(1).Text = "[No Caption Detected]"


'Find the parent using the GetParent function. The loop is for
'detecting the Zero-th level parent of our window


Parent = GetParent(hWn)
Do While Parent
OldParent = Parent
Parent = GetParent(OldParent)
Loop
If Parent Then OldParent = Parent
GetWindowText OldParent, WinCap, 254
txtMain(2).Text = WinCap
If Trim(txtMain(2).Text) = "" Then txtMain(2).Text = "[No Perent Detected]"


'Get the class name of our window

GetClassName hWn, ClName, 254
txtMain(3).Text = ClName
If Trim(txtMain(3).Text) = "" Then txtMain(3).Text = "[No Class Detected]"
   
   If SendMessage(hWn, EM_GETPASSWORDCHAR, 0, 1&) <> 0 Then
   SendMessage hWn, EM_SETPASSWORDCHAR, 0, 1&
   SendMessage hWn, EM_SETMODIFY, True, 1&
   End If

End Sub

