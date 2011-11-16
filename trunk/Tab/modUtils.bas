Attribute VB_Name = "modUtils"
'modUtils:  Contains Utility Procedures

Option Explicit

' Convert the OLE color into equivalent RGB Combination
' i.e. Convert vbButtonFace into ==> Light Grey
Public Function g_pGetRGBFromOLE(lOleColor As Long) As Long
  Dim lRGBColor As Long
  Call TranslateColor(lOleColor, 0, lRGBColor)
  g_pGetRGBFromOLE = lRGBColor
End Function

' Function used to dispaly fileopen dialog. I didn't used
' MS Common Dialog Control bcozSince i didn't wanted to
' use any 3rd party control...
Public Function g_pShowFileOpenDialog(lhWndOwner As Long, Optional ByVal sInitDir As String = "", Optional ByVal sFilter As String = "") As String
  On Error Resume Next
    
  Dim utOFName As OPENFILENAME
    
  With utOFName
    
    .lStructSize = Len(utOFName)
      
    .flags = 0
      
    .hwndOwner = lhWndOwner
      
    .hInstance = App.hInstance
      
    If sFilter <> "" Then
      .lpstrFilter = Replace$(sFilter, "|", vbNullChar)
    Else
      .lpstrFilter = "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar
    End If
    'create a buffer
    .lpstrFile = Space$(254)
    'set the maximum length of a returned file (important)
    .nMaxFile = 255
      
    .lpstrFileTitle = Space$(254)
      
    .nMaxFileTitle = 255
      
    .lpstrInitialDir = sInitDir
    .lpstrTitle = "Open File"

  End With
    
  'Show the dialog
  If GetOpenFileName(utOFName) Then
    g_pShowFileOpenDialog = Trim$(utOFName.lpstrFile)
  Else
    'Cancel Pressed
    g_pShowFileOpenDialog = ""
  End If
End Function

  
Public Sub DrawImage(ByVal lDestHDC As Long, ByVal lhBmp As Long, ByVal lTransColor As Long, ByVal iLeft As Integer, ByVal iTop As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer)
  Dim lhDC As Long
  Dim lhBmpOld As Long
  Dim utBmp As BITMAP

  lhDC = CreateCompatibleDC(lDestHDC)

  lhBmpOld = SelectObject(lhDC, lhBmp)

  Call GetObjectA(lhBmp, Len(utBmp), utBmp)
  
  Call TransparentBlt(lDestHDC, iLeft, iTop, iWidth, iHeight, lhDC, 0, 0, utBmp.bmWidth, utBmp.bmHeight, lTransColor)

  Call SelectObject(lhDC, lhBmpOld)
  DeleteDC (lhDC)
End Sub
  
