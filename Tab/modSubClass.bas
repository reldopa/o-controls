Attribute VB_Name = "modSubClass"
'modSubClass: Contains Sub class relate code

Option Explicit

'Used to call pWindowProc of the appropriate Control
Dim m_oTmpCtl As oTab

'Used to store a Object pointer
Dim m_lObjPtr As Long

Public Function WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
  On Error Resume Next

  m_lObjPtr = GetWindowLong(hwnd, GWL_USERDATA)

  CopyMemory m_oTmpCtl, m_lObjPtr, 4

  'Call the WindowProc function for the appropriate instance of our control
  WndProc = m_oTmpCtl.pWindowProc(hwnd, Msg, wParam, lParam)

    
  'Destroy tmp control's interface copy (we just need the type defs)
  CopyMemory m_oTmpCtl, 0&, 4
    
  Set m_oTmpCtl = Nothing
End Function


