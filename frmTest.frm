VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "JKH ScreenTips"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CommandButton"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Code by Janus Kamp Hansen"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   2100
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Tips As cTips

Private Sub Check1_Click()
  Tips.ScreenTips "This is checkbox. Click it to enable or disable something...", vbExclamation, "Checkbox information.", Check1
End Sub

Private Sub Command1_Click()
  Tips.ScreenTips "This is JKH Screen Tips, like in Win2K" & vbCrLf & "Code by Janus Kamp Hansen", vbInformation, "JKH Screen Tips", Command1
  
End Sub

Private Sub Form_Load()
  Set Tips = New cTips

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set Tips = Nothing
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Tips.ScreenTips "This is a list box. Did you know that?", vbQuestion, "List box question", List1


End Sub

Private Sub Option1_Click()
  Tips.ScreenTips "... this is an option box, if you click it, you can't unclick it!", vbCritical, "Critial information", Option1

End Sub

