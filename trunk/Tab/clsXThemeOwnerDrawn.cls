VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xThemeOwnerDrawn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'xThemeOwnerDrawn : Calls events to allow client to draw the complete control
'                   But NOT IMPLEMENTED COMPLETELY. Since Contained controls need to be handled
'                   See ITheme for implemented function description

Option Explicit

Implements ITheme

Private m_oITGTab As oTab

'===Implemented Functions=============================================================================================

Private Sub ITheme_DrawBackground()
  m_oITGTab.bAutoRedraw = False
  m_oITGTab.pRaiseDrawBackground
End Sub

Private Sub ITheme_DrawBackgroundPropertyPages()
  'no need in the Owner Drawn Theme
End Sub

Private Sub ITheme_DrawBackgroundTabbedDialog()

End Sub

'these functions were seperated acc to theme so as to allow painting of only the desired parts (and not the whole control)
Private Sub ITheme_DrawOnActiveTabChange()
  m_oITGTab.bAutoRedraw = False
  m_oITGTab.pRaiseDrawOnActiveTabChange
  'm_oITGTab.bAutoRedraw = True
End Sub

Private Sub ITheme_DrawTabs()
  m_oITGTab.bAutoRedraw = False
  m_oITGTab.pRaiseDrawTabs
  'm_oITGTab.bAutoRedraw = True
End Sub

Private Sub Itheme_DrawTabsPropertyPages()

End Sub

Private Sub ITheme_DrawTabsTabbedDialog()

End Sub

Private Sub ITheme_MouseDownHanlder(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)

End Sub

Private Sub ITheme_MouseMoveHanlder(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)

End Sub

Private Sub ITheme_MouseUpHanlder(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)
  'Do Nothing
End Sub

Private Sub ITheme_SetControl(oITGTab As oTab)
  Dim utRect As RECT
  
  Set m_oITGTab = oITGTab
  
  'must set the following properties for the owner drawn control
  m_oITGTab.bAutoRedraw = True
  m_oITGTab.pCls                  'clear any existing things
  m_oITGTab.bAutoRedraw = False
  
  If Not m_oITGTab.bUserMode Then
    m_oITGTab.bAutoRedraw = True
    utRect.Right = m_oITGTab.iScaleWidth
    utRect.Bottom = m_oITGTab.iScaleHeight
    
    m_oITGTab.pLine 0, 0, m_oITGTab.iScaleWidth - 1, m_oITGTab.iScaleHeight - 1, m_oITGTab.OuterBorderColor, True, False
    
    DrawText m_oITGTab.lhDC, "Owner Drawn not available at design time", -1, utRect, DT_VCENTER Or DT_SINGLELINE Or DT_CENTER Or DT_END_ELLIPSIS
  End If
  
  
End Sub

Private Sub ITheme_ShowHideFocus()
  'm_oITGTab.bAutoRedraw = False
  m_oITGTab.pRaiseShowHideFocus
  'm_oITGTab.bAutoRedraw = True
End Sub

Private Sub ITheme_ShowHideFocusPropertyPages()

End Sub

Private Sub ITheme_ShowHideFocusTabbedDialog()

End Sub


Private Sub ITheme_TimerEvent()

End Sub


Private Sub ITheme_ResetColorsToDefault()
  'do nothing
End Sub
'=====================================================================================================================



