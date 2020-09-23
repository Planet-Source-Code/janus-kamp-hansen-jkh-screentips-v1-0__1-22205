VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type POINTAPI
 x As Integer
 y As Integer
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Enum IconStyle
  vbCritical = 16
  vbQuestion = 32
  vbExclamation = 48
  vbInformation = 64
End Enum


Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Sub SetCursorPos Lib "User" (ByVal x As Integer, ByVal y As Integer)
Private Declare Sub ClientToScreen Lib "User" (ByVal hwnd As Integer, IpPoint As POINTAPI)
Private Declare Function GetParent Lib "User" (ByVal hwnd As Integer) As Integer

Private fMain As frmMain

Public Function ScreenTips(ByVal Prompt As String, Optional ByVal vbIconType As IconStyle = vbInformation, Optional ByVal Header As String = "", Optional ByVal PointTo As Object) As Boolean
  On Error Resume Next
  Dim r As RECT
  
  If fMain Is Nothing Then
  
    Set fMain = New frmMain
    
  Else
    Unload fMain
    Set fMain = Nothing
    DoEvents
    Set fMain = New frmMain
  End If
    
    If fMain.imgList16.ListImages.Count = 0 Then fMain.ImgListSetup
    
    fMain.lblHeader.Caption = Header
    fMain.lblMessage.Caption = Prompt
    
    Select Case vbIconType
    Case vbQuestion
      fMain.picIcon.Picture = fMain.imgList16.ListImages("MB_QUESTION").ExtractIcon
    
    Case vbExclamation
      fMain.picIcon.Picture = fMain.imgList16.ListImages("MB_EXCLAMATION").ExtractIcon
    
    Case vbCritical
      fMain.picIcon.Picture = fMain.imgList16.ListImages("MB_CRITICAL").ExtractIcon
    
    Case vbInformation
      fMain.picIcon.Picture = fMain.imgList16.ListImages("MB_INFORMATION").ExtractIcon
    
    Case Else
      
    End Select
    
    fMain.ShowMe
  
  If Not PointTo Is Nothing Then
    GetWindowRect PointTo.hwnd, r
    fMain.Left = (PixelsToTwips_width(r.Left)) - 50
    fMain.Top = (PixelsToTwips_height(r.Top)) - (fMain.Height - 400)
  
    fMain.Refresh
  End If

End Function

Private Sub Class_Terminate()
On Error Resume Next
  Unload fMain
  Set fMain = Nothing
  Kill GetTempDir & "!" ' Cleaning up
End Sub

Private Function PixelsToTwips_height(pxls)
  PixelsToTwips_height = pxls * Screen.TwipsPerPixelY
End Function


Private Function PixelsToTwips_width(pxls)
  PixelsToTwips_width = pxls * Screen.TwipsPerPixelX
End Function


