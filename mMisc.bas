Attribute VB_Name = "mMisc"
Option Explicit

Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private CurRgn As Long, TempRgn As Long
Private objName1 As Object

Private Const RGN_DIFF = 4
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private lngHeight As Long
Private lngWidth As Long

Public Sub SetRegion(Objct As Object, Colr As Long, Optional FileName As String = vbNullString)
On Error Resume Next
Dim fso, Exists As Boolean
Dim s As Long
Dim lngHDC As Long
    
    Exists = FileExists(FileName)
    
    Set objName1 = Objct
    If CurRgn& Then DeleteObject CurRgn&
    If (Exists = True) Then
      CurRgn& = LoadMask(Str(FileName))
    Else
      CurRgn& = GetBitmapRegion(objName1.Picture, Colr&)
      
      If FileName <> vbNullString Then
          Call SaveMask(FileName)
      End If
    End If
    
    s = SetWindowRgn(objName1.hwnd, CurRgn&, True)
    
    ReleaseDC objName1.hwnd, lngHDC
            
End Sub

Private Function GetBitmapRegion(ByRef cPicture As StdPicture, ByRef cTransparent As Long)
  Dim hRgn As Long
  Dim tRgn As Long
  Dim x As Long
  Dim y As Long
  Dim X0 As Long
  Dim hDC As Long
  Dim mBitMap As BITMAP
    
    hDC = CreateCompatibleDC(0)
    If hDC Then
        
      SelectObject hDC, cPicture
      
      GetObject cPicture, Len(mBitMap), mBitMap
      hRgn = CreateRectRgn(0, 0, mBitMap.bmWidth, mBitMap.bmHeight)
              
      For y = 0 To mBitMap.bmHeight
        For x = 0 To mBitMap.bmWidth
              
          While x <= mBitMap.bmWidth And GetPixel(hDC, x, y) <> cTransparent
            x = x + 1
          Wend
          
          X0 = x
          
          While x <= mBitMap.bmWidth And GetPixel(hDC, x, y) = cTransparent
            x = x + 1
          Wend
          
          If X0 < x Then
            tRgn = CreateRectRgn(X0, y, x, y + 1)
            CombineRgn hRgn, hRgn, tRgn, 4
            DeleteObject tRgn
          End If
        Next x
      Next y
      
      GetBitmapRegion = hRgn
    End If
    
    DeleteDC hDC
    
End Function
Public Sub DragForm(hwnd As Long, intButton As Integer)
On Error Resume Next
  If intButton = vbLeftButton Then
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
  End If
End Sub

Private Sub SaveMask(ByVal sFilename As String)
On Error Resume Next
Dim iFile As Long
Dim nBytes As Long
Dim b() As Byte

  nBytes = GetRegionData(CurRgn, 0, ByVal 0&)
  If nBytes > 0 Then
     ReDim b(0 To nBytes - 1) As Byte
     If nBytes = GetRegionData(CurRgn, nBytes, b(0)) Then
        Kill sFilename
        iFile = FreeFile
        Open sFilename For Binary Access Write Lock Read As #iFile
        Put #iFile, , b
        Close #iFile
     End If
  End If
End Sub

Private Function LoadMask(ByVal sFilename As String) As Long
On Error Resume Next
Dim iFile As Long
Dim b() As Byte
Dim dwCount As Long

   iFile = FreeFile
   Open sFilename For Binary Access Read Lock Write As #iFile
   ReDim b(0 To LOF(iFile) - 1) As Byte
   Get #iFile, , b
   Close #iFile
   
   dwCount = UBound(b) - LBound(b) + 1
   LoadMask = ExtCreateRegion(ByVal 0&, dwCount, b(0))

End Function

Public Function GetTempDir() As String
Dim r As Long, nSize As Long
Dim Tmp As String
  Tmp = Space$(256)
  nSize = Len(Tmp)
  r = GetTempPath(nSize, Tmp)
  GetTempDir = TrimNull(Tmp)
End Function

Private Function TrimNull(item As String)
Dim Pos As Integer
' double check that there is a chr$(0) in the string
  Pos = InStr(item, Chr$(0))

  If Pos Then
    TrimNull = Left$(item, Pos - 1)
  Else
    TrimNull = item
  End If
End Function

Public Function FileExists(ByVal FileName As String) As Boolean
Dim FileNo As Integer
On Local Error Resume Next

  FileExists = False
  FileNo = FreeFile

' Open a specified existing file
  Open FileName For Input As #FileNo

' Error handler generates error and exits the routine
  If Err Then Exit Function
  Close #FileNo
  FileExists = True
End Function
