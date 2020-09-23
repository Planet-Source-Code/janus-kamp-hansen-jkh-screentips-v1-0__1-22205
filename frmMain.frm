VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   1635
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      Picture         =   "frmMain.frx":0A58
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin MSComctlLib.ImageList imgList16 
      Left            =   4080
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Label lblHeader 
      BackColor       =   &H80000018&
      Caption         =   "Header"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   540
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H80000018&
      Caption         =   "Message"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   430
      Width           =   4380
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Public Sub ImgListSetup()

  With imgList16
    .ImageHeight = 16
    .ImageWidth = 16
  
    .ListImages.Add , "MB_CRITICAL", LoadResPicture("MB_CRITICAL", vbResIcon)
    .ListImages.Add , "MB_QUESTION", LoadResPicture("MB_QUESTION", vbResIcon)
    .ListImages.Add , "MB_EXCLAMATION", LoadResPicture("MB_EXCLAMATION", vbResIcon)
    .ListImages.Add , "MB_INFORMATION", LoadResPicture("MB_INFORMATION", vbResIcon)
    
  
  End With

End Sub

Public Sub ShowMe()
  Dim HotRed As Long
  HotRed& = RGB(255, 0, 255)
  
  SetRegion Me, HotRed&, GetTempDir & "!"
  Me.Show
Dim Res%
  Res% = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
  
End Sub

Private Sub Form_Click()
  Unload Me
End Sub

Private Sub Form_LostFocus()
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Kill GetTempDir & "!" ' Cleaning up
End Sub

Private Sub lblHeader_Click()
  Unload Me
End Sub

Private Sub lblMessage_Click()
  Unload Me
End Sub

Private Sub PicIcon_Click()
  Unload Me
End Sub

