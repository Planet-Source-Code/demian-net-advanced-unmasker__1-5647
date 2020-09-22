VERSION 5.00
Begin VB.Form frmTrayIcon 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   ClientHeight    =   1620
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   2700
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
   cbSize            As Long
   hwnd              As Long
   uId               As Long
   uFlags            As Long
   ucallbackMessage  As Long
   hIcon             As Long
   szTip             As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private TrayIcon        As NOTIFYICONDATA

Private IconIsOn        As Boolean

Public Owner            As Form
Public ShowTray         As Boolean
Public Sub SetCallback(ob As Object)
   ' *** Set's which object to call "TrayIconCallback(msg As Long)" in.
   
   Set Owner = ob
   
End Sub
Public Sub Update(IconOn As Boolean)
   ' *** Turns the Trayicon on and off.
   ' *** Note: the Icon, and ToolTip are taken from the
   ' ***   form's icon and caption respectively.
   If IconOn Then
      TrayIcon.cbSize = Len(TrayIcon)
      TrayIcon.hwnd = Me.hwnd
      TrayIcon.uId = 1&
      TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      If Me.Caption <> "" Then
         TrayIcon.uFlags = TrayIcon.uFlags Or NIF_TIP
      End If
      TrayIcon.ucallbackMessage = WM_MOUSEMOVE
      TrayIcon.hIcon = Me.Icon
      TrayIcon.szTip = Me.Caption & Chr$(0)
      
      If Not IconIsOn Then
         Shell_NotifyIcon NIM_ADD, TrayIcon
      Else
         Shell_NotifyIcon NIM_MODIFY, TrayIcon
      End If
      IconIsOn = True
   Else
      If IconIsOn Then
         TrayIcon.cbSize = Len(TrayIcon)
         TrayIcon.hwnd = Me.hwnd
         TrayIcon.uId = 1&
         Shell_NotifyIcon NIM_DELETE, TrayIcon
      End If
      IconIsOn = False
   End If

End Sub
Private Sub mnuTrayRestore_Click()

End Sub

Private Sub Form_Load()
   
   IconIsOn = False
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Static rec As Boolean, Msg As Long
    
    Msg = X / Screen.TwipsPerPixelX
    If rec = False Then
        rec = True
        
        ' Ignore an error if the call to
        ' Owner.TrayIconCallback fails
        On Error Resume Next
        If Msg >= WM_LBUTTONDOWN And Msg <= WM_RBUTTONDBLCLK Then
           Call Owner.TrayIconCallback(Msg)
        End If
        
        rec = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Me.Update False

End Sub
