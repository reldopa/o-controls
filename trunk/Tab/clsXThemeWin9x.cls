VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xThemeWin9x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'xThemeWin9x : Theme mimics the standard tabs (Windows 9x,2k,ME).
'              But allows customization
'
'              See ITheme for implemented function description
Option Explicit

Implements ITheme

'==== Constants ======================================================================================================
Private Const m_c_iFOCUS_RECT_AND_TEXT_DISTANCE = 2                   'the distance between the text (caption) of the tab and the focus Rect
Private Const m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE As Integer = 7  'the distance between the text and the border in a Property Pages style tab
Private Const m_c_iPROP_PAGE_INACTIVE_TOP As Integer = 2              'the top for the property page (inactive property page)
'=====================================================================================================================

'==== Private Variables ==============================================================================================

Private m_oITGTab As oTab

'following property are cached (locally) for improved performance
Private m_lHDC As Long
Private m_iActiveTab As Integer
Private m_iActiveTabHeight As Integer
Private m_iInActiveTabHeight As Integer
Private m_iTabCount As Integer
Private m_bIsFocused As Boolean
Private m_iScaleWidth As Integer
Private m_iScaleHeight As Integer
Private m_lOuterBorderColor As Long
Private m_lTopLeftInnerBorderColor As Long
Private m_lBottomRightInnerBorderColor As Long
Private m_lActiveTabForeColor As Long
Private m_lActiveTabBackStartColor As Long
Private m_lActiveTabBackEndColor As Long
Private m_lInActiveTabForeColor As Long
Private m_lInActiveTabBackStartColor As Long
Private m_lInActiveTabBackEndColor As Long
Private m_lDisabledTabForeColor As Long
Private m_lDisabledTabBackColor As Long
Private m_iIconSize As Integer




'=== Event Handlers ==================================================================================================

Private Sub ITheme_DrawBackground()
  
  Call pCacheControlProperties     'cache the control proerties for faster access
  Select Case m_oITGTab.TabStyle
    Case xStyleTabbedDialog:
      Call ITheme_DrawBackgroundTabbedDialog
    Case xStylePropertyPages:
      Call ITheme_DrawBackgroundPropertyPages
  End Select
End Sub

Private Sub ITheme_DrawBackgroundTabbedDialog()
  Dim iTmp As Integer
  
  With m_oITGTab
  
    'get the larger of the active tab height and inactive tab height
    iTmp = IIf(m_iActiveTabHeight > m_iInActiveTabHeight, m_iActiveTabHeight, m_iInActiveTabHeight)

    .pCls   'clear the control
    
    'fill background color based on tab's enabled property
    If .aryTabs(m_iActiveTab).Enabled Then
      
      .lBackColor = m_lActiveTabBackEndColor
    Else
      
      .lBackColor = m_lDisabledTabBackColor
     
    End If
    
    'draw inner shadow (left)
    .pLine 0, iTmp + 1, 0, m_iScaleHeight - 1, m_lOuterBorderColor
    .pLine 1, iTmp + 2, 1, m_iScaleHeight - 2, m_lTopLeftInnerBorderColor
    .pLine 2, iTmp + 3, 2, m_iScaleHeight - 3, m_lTopLeftInnerBorderColor
    
    'draw inner shadow (right)
    .pLine m_iScaleWidth - 1, iTmp + 1, m_iScaleWidth - 1, m_iScaleHeight - 1, m_lOuterBorderColor
    .pLine m_iScaleWidth - 2, iTmp + 2, m_iScaleWidth - 2, m_iScaleHeight - 2, m_lBottomRightInnerBorderColor
    .pLine m_iScaleWidth - 3, iTmp + 3, m_iScaleWidth - 3, m_iScaleHeight - 3, m_lBottomRightInnerBorderColor
    
    'draw inner shadow (bottom)
    .pLine 0, m_iScaleHeight - 1, m_iScaleWidth, m_iScaleHeight - 1, m_lOuterBorderColor
    .pLine 1, m_iScaleHeight - 2, m_iScaleWidth - 1, m_iScaleHeight - 2, m_lBottomRightInnerBorderColor
    .pLine 2, m_iScaleHeight - 3, m_iScaleWidth - 2, m_iScaleHeight - 3, m_lBottomRightInnerBorderColor
  End With
End Sub


Private Sub ITheme_DrawBackgroundPropertyPages()
  Dim iTmp As Integer
  
  With m_oITGTab
    
    'get the larger of the active tab height and inactive tab height
    iTmp = IIf(m_iActiveTabHeight > m_iInActiveTabHeight, m_iActiveTabHeight, m_iInActiveTabHeight)
  
    
    .pCls
    
    'fill background color based on tab's enabled property
    If .aryTabs(m_iActiveTab).Enabled Then
      .lBackColor = m_lActiveTabBackEndColor
    Else
      .lBackColor = m_lDisabledTabBackColor
    End If
  
        
    'draw inner shadow (left)
    .pLine 0, iTmp + 1, 0, m_iScaleHeight - 1, m_lTopLeftInnerBorderColor
    
    'draw inner shadow (right)
    .pLine m_iScaleWidth - 1, iTmp + 1, m_iScaleWidth - 1, m_iScaleHeight - 1, m_lOuterBorderColor
    .pLine m_iScaleWidth - 2, iTmp + 1, m_iScaleWidth - 2, m_iScaleHeight - 2, m_lBottomRightInnerBorderColor
    
    'draw inner shadow (bottom)
    .pLine 0, m_iScaleHeight - 1, m_iScaleWidth, m_iScaleHeight - 1, m_lOuterBorderColor
    .pLine 1, m_iScaleHeight - 2, m_iScaleWidth - 1, m_iScaleHeight - 2, m_lBottomRightInnerBorderColor
  End With
End Sub

Private Sub ITheme_DrawOnActiveTabChange()
  ITheme_DrawTabs   'only redraw the top part
End Sub

Private Sub ITheme_DrawTabs()
  
  Call pCacheControlProperties     'cache the control proerties for faster access
  Select Case m_oITGTab.TabStyle
    Case xStyleTabbedDialog:
      Call ITheme_DrawTabsTabbedDialog
    Case xStylePropertyPages:
      Call Itheme_DrawTabsPropertyPages
  End Select
End Sub

Private Sub Itheme_DrawTabsPropertyPages()
  Dim iCnt As Integer
  Dim iTabWidth As Integer
  Dim utFontRect As RECT
  Dim sTmp As String
  Dim iTmpHeight As Integer
  Dim utTabInfo As TabInfo
  Dim iAdjustedIconSize As Integer
  Dim iTmpY As Integer
  
  'Set the active tab's font as current font (since the TextWidth function
  'will use the current font's size)
  Set m_oITGTab.oFont = m_oITGTab.ActiveTabFont
  
 
  'store the larger height in tmp var
  iTmpHeight = IIf(m_iActiveTabHeight > m_iInActiveTabHeight, m_iActiveTabHeight, m_iInActiveTabHeight)
  
  'initialize the clickable items
  For iCnt = 0 To m_iTabCount - 1
    utTabInfo = m_oITGTab.aryTabs(iCnt)
    sTmp = Replace$(utTabInfo.Caption, "&&", "&")
    
    If InStr(1, sTmp, "&") Then
      
      'if still there is one '&' in the string then reduce the width by one more character (since the '&' will be conveted into an underline when painted)
      sTmp = Mid$(sTmp, 1, Len(sTmp) - 1)
    End If
    

    
    If utTabInfo.TabPicture Is Nothing Then
      'get tab width acc to the text size and border
      iTabWidth = m_oITGTab.pTextWidth(sTmp) + m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE * 2
    Else
      If iTmpHeight - 6 < m_iIconSize Then    '-6 for borders
        'here we adjust the size of the icon if it does not fit into current tab
        iAdjustedIconSize = iTmpHeight - 6
      Else
        iAdjustedIconSize = m_iIconSize
      End If

      'get tab width acc to the text size, border and Image
      iTabWidth = m_oITGTab.pTextWidth(sTmp) + (m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE * 2) + iAdjustedIconSize + 1
    End If
    
    
    'following adjustments are used in case of property pages only. We must shift
    'the left (+2) or (-2) to make it look like standard property pages
    With utTabInfo.ClickableRect
      If iCnt = 0 And iCnt <> m_iActiveTab Then
        .Left = m_c_iPROP_PAGE_INACTIVE_TOP
        .Right = .Left + iTabWidth - m_c_iPROP_PAGE_INACTIVE_TOP
      Else
        If iCnt = 0 Then
          .Left = 0
        Else
          .Left = m_oITGTab.aryTabs(iCnt - 1).ClickableRect.Right
        End If
          
        .Right = .Left + iTabWidth
      End If
        
      If iCnt = m_iActiveTab Then
        If m_iActiveTabHeight > m_iInActiveTabHeight Then
          .Top = 0
        Else
          .Top = m_iInActiveTabHeight - m_iActiveTabHeight
        End If
        .Bottom = .Top + m_iActiveTabHeight
      Else
        If m_iInActiveTabHeight > m_iActiveTabHeight Then
          .Top = 0
          .Bottom = .Top + m_iInActiveTabHeight
        Else
          .Top = m_iActiveTabHeight - m_iInActiveTabHeight
          .Bottom = .Top + m_iInActiveTabHeight
        End If
        
      End If

    End With
    m_oITGTab.aryTabs(iCnt) = utTabInfo       'assign the new tab info to the existing one
  Next
  
  
  'fill the tab strip with TabStripBackColor (customizable... so that tab's can easily blend with the background)
  m_oITGTab.pLine 0, 0, m_iScaleWidth, iTmpHeight, m_oITGTab.TabStripBackColor, True, True
  
  'Now Draw Each Tab
  For iCnt = 0 To m_iTabCount - 1
    
    utTabInfo = m_oITGTab.aryTabs(iCnt)   'get local copy
    
    With utTabInfo.ClickableRect
      
      If iCnt = m_iActiveTab Then       'if we are drawing the active tab
        
        
        If utTabInfo.Enabled Then
          Call m_oITGTab.pFillCurvedGradient(.Left, .Top, .Right - 1, .Bottom, m_lActiveTabBackStartColor, m_lActiveTabBackEndColor, 2, True, True)
        Else
          Call m_oITGTab.pFillCurvedGradient(.Left, .Top, .Right - 1, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor, 2, True, True)
        End If
        
        'top-left corner
        m_oITGTab.pLine .Left, .Top + 2, .Left + 3, .Top - 1, m_lTopLeftInnerBorderColor
        
        'top line
        m_oITGTab.pLine .Left + 2, .Top, .Right - 2, .Top, m_lTopLeftInnerBorderColor
        
        'top-right corner
        m_oITGTab.pLine .Right - 2, .Top + 1, .Right, .Top + 3, m_lOuterBorderColor
      
        'right line
        m_oITGTab.pLine .Right - 1, .Top + 2, .Right - 1, .Bottom + 1, m_lOuterBorderColor
        m_oITGTab.pLine .Right - 2, .Top + 2, .Right - 2, .Bottom + 1, m_lBottomRightInnerBorderColor
        
        'left line
        m_oITGTab.pLine .Left, .Top + 2, .Left, .Bottom + 1, m_lTopLeftInnerBorderColor
        
        Set m_oITGTab.oFont = m_oITGTab.ActiveTabFont 'set the active tab font as current font
        
        'set foreground color (for text drawing)
        If utTabInfo.Enabled Then
          m_oITGTab.lForeColor = m_lActiveTabForeColor
        Else
          m_oITGTab.lForeColor = m_lDisabledTabForeColor
        End If
        
      Else    'we are not drawing active tab, it can be any one from the following
          
        
        If iCnt = m_iActiveTab - 1 Then         'if we are drawing tab just b4 active tab, then
        
          If utTabInfo.Enabled Then
            'following adjustments are needed if the inactive tab's height is more than active tab's height (the tab's corner should be properly rounded)
            If m_iInActiveTabHeight + 2 > m_iActiveTabHeight Then
              Call m_oITGTab.pFillCurvedGradient(.Left, .Top, .Right, .Bottom, m_lInActiveTabBackStartColor, m_lInActiveTabBackEndColor, 2, True, True)
            Else
              Call m_oITGTab.pFillCurvedGradient(.Left, .Top, .Right, .Bottom, m_lInActiveTabBackStartColor, m_lInActiveTabBackEndColor, 2, True, False)
            End If
          Else
            'following adjustments are needed if the inactive tab's height is more than active tab's height (the tab's corner should be properly rounded)
            If m_iInActiveTabHeight + 2 > m_iActiveTabHeight Then
              Call m_oITGTab.pFillCurvedGradient(.Left, .Top, .Right, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor, 2, True, True)
            Else
              Call m_oITGTab.pFillCurvedGradient(.Left, .Top, .Right, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor, 2, True, False)
            End If
          End If

          'top-left corner
          m_oITGTab.pLine .Left, .Top + 2, .Left + 3, .Top - 1, m_lTopLeftInnerBorderColor
          
          'following adjustments are needed if the inactive tab's height is more than active tab's height (the tab's corner should be properly rounded)
          If m_iInActiveTabHeight + 2 > m_iActiveTabHeight Then
            'top line
            m_oITGTab.pLine .Left + 2, .Top, .Right - 2, .Top, m_lTopLeftInnerBorderColor
            
            'top-right corner
            m_oITGTab.pLine .Right - 2, .Top + 1, .Right, .Top + 3, m_lOuterBorderColor
          
            'right line
            m_oITGTab.pLine .Right - 1, .Top + 2, .Right - 1, .Bottom + 1, m_lOuterBorderColor
            m_oITGTab.pLine .Right - 2, .Top + 2, .Right - 2, .Bottom + 1, m_lBottomRightInnerBorderColor
          Else
            'top line
            m_oITGTab.pLine .Left + 2, .Top, .Right, .Top, m_lTopLeftInnerBorderColor
          End If
          
    
          'bottom line
          m_oITGTab.pLine .Left, .Bottom, .Right, .Bottom, m_lTopLeftInnerBorderColor
            
          'Left Line
          m_oITGTab.pLine .Left, .Top + 2, .Left, .Bottom + 1, m_lTopLeftInnerBorderColor
          
        
        ElseIf iCnt = m_iActiveTab + 1 Then   'tab just after active tab
          
          If utTabInfo.Enabled Then
            'following adjustments are needed if the inactive tab's height is more than active tab's height (the tab's corner should be properly rounded)
            If m_iInActiveTabHeight + 2 > m_iActiveTabHeight Then
              Call m_oITGTab.pFillCurvedGradient(.Left, .Top, .Right - 1, .Bottom, m_lInActiveTabBackStartColor, m_lInActiveTabBackEndColor, 2, True, True)
            Else
              Call m_oITGTab.pFillCurvedGradient(.Left, .Top, .Right - 1, .Bottom, m_lInActiveTabBackStartColor, m_lInActiveTabBackEndColor, 2, False, True)
            End If
          Else
            'following adjustments are needed if the inactive tab's height is more than active tab's height (the tab's corner should be properly rounded)
            If m_iInActiveTabHeight + 2 > m_iActiveTabHeight Then
              Call m_oITGTab.pFillCurvedGradient(.Left, .Top, .Right - 1, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor, 2, True, True)
            Else
              Call m_oITGTab.pFillCurvedGradient(.Left, .Top, .Right - 1, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor, 2, False, True)
            End If
          End If
          
          'following adjustments are needed if the inactive tab's height is more than active tab's height (the tab's corner should be properly rounded)
          If m_iInActiveTabHeight + 2 > m_iActiveTabHeight Then
            'top-left corner
            m_oITGTab.pLine .Left, .Top + 2, .Left + 3, .Top - 1, m_lTopLeftInnerBorderColor
          
            'top line
            m_oITGTab.pLine .Left + 2, .Top, .Right - 2, .Top, m_lTopLeftInnerBorderColor
            
            'Left Line
            m_oITGTab.pLine .Left, .Top + 2, .Left, .Bottom + 1, m_lTopLeftInnerBorderColor
          Else
            'top line
            m_oITGTab.pLine .Left, .Top, .Right - 2, .Top, m_lTopLeftInnerBorderColor
          End If
  
    
          'top-right corner
          m_oITGTab.pLine .Right - 2, .Top + 1, .Right, .Top + 3, m_lOuterBorderColor
    
          'right line
          m_oITGTab.pLine .Right - 1, .Top + 2, .Right - 1, .Bottom, m_lOuterBorderColor
          m_oITGTab.pLine .Right - 2, .Top + 2, .Right - 2, .Bottom, m_lBottomRightInnerBorderColor
    
          'bottom line
          m_oITGTab.pLine .Left, .Bottom, .Right, .Bottom, m_lTopLeftInnerBorderColor
          
        Else  'other non active tab (must draw full curves on both the sides)
          
          If utTabInfo.Enabled Then
            Call m_oITGTab.pFillCurvedGradient(.Left, .Top, .Right - 1, .Bottom, m_lInActiveTabBackStartColor, m_lInActiveTabBackEndColor, 2, True, True)
          Else
            Call m_oITGTab.pFillCurvedGradient(.Left, .Top, .Right - 1, .Bottom, m_lDisabledTabBackColor, m_lInActiveTabBackEndColor, 2, True, True)
          End If
          
          'top-left corner
          m_oITGTab.pLine .Left, .Top + 2, .Left + 3, .Top - 1, m_lTopLeftInnerBorderColor
            
          'top line
          m_oITGTab.pLine .Left + 2, .Top, .Right - 2, .Top, m_lTopLeftInnerBorderColor
    
          'top-right corner
          m_oITGTab.pLine .Right - 2, .Top + 1, .Right, .Top + 3, m_lOuterBorderColor
    
          'right line
          m_oITGTab.pLine .Right - 1, .Top + 2, .Right - 1, .Bottom, m_lOuterBorderColor
          m_oITGTab.pLine .Right - 2, .Top + 2, .Right - 2, .Bottom, m_lBottomRightInnerBorderColor
    
          'bottom line
          m_oITGTab.pLine .Left, .Bottom, .Right, .Bottom, m_lTopLeftInnerBorderColor
            
          'Left Line
          m_oITGTab.pLine .Left, .Top + 2, .Left, .Bottom + 1, m_lTopLeftInnerBorderColor
  
        End If
        
        Set m_oITGTab.oFont = m_oITGTab.InActiveTabFont   'set the font
        
        'set fore color
        If utTabInfo.Enabled Then
          m_oITGTab.lForeColor = m_lInActiveTabForeColor
        Else
          m_oITGTab.lForeColor = m_lDisabledTabForeColor
        End If
      End If
      
      'if its left most tab then adjust the bottom line
      If iCnt = 0 Then
        'bottom line
        m_oITGTab.pLine .Left - 2, .Bottom, .Left, .Bottom, m_lTopLeftInnerBorderColor
      End If
      
      'do the adjustments for the border
      utFontRect.Left = .Left + 2
      utFontRect.Top = .Top + 2
      utFontRect.Bottom = .Bottom
      utFontRect.Right = .Right - 3
      
      'Get the Tab's caption in a local variable
      sTmp = utTabInfo.Caption
      
      
      If Not utTabInfo.TabPicture Is Nothing Then
        
        If iTmpHeight - 6 < m_iIconSize Then    '-6 for borders
          'here we adjust the size of the icon if it does not fit into current tab
          iAdjustedIconSize = iTmpHeight - 6
        Else
          iAdjustedIconSize = m_iIconSize
        End If
    
        iTmpY = utFontRect.Top + Round((utFontRect.Bottom - utFontRect.Top - iAdjustedIconSize) / 2)
        
        Select Case m_oITGTab.PictureAlign
          Case xAlignLeftEdge, xAlignLeftOfCaption:
            
            If utTabInfo.TabPicture.Type = vbPicTypeBitmap And m_oITGTab.UseMaskColor Then
              
              Call DrawImage(m_lHDC, utTabInfo.TabPicture.Handle, g_pGetRGBFromOLE(m_oITGTab.PictureMaskColor), utFontRect.Left + 2, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            Else
              
              Call m_oITGTab.pPaintPicture(utTabInfo.TabPicture, utFontRect.Left + 2, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            End If
            
            'shift the text to be drawn after the picture
            utFontRect.Left = (utFontRect.Left + iAdjustedIconSize + 1) - m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE
            
            'call the API for the text drawing
            DrawText m_lHDC, sTmp, -1, utFontRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
            
            
            'revert the changes so that the focus rectangle can be drawn for the whole tab's clickable area
            utFontRect.Left = (utFontRect.Left - iAdjustedIconSize - 1) + m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE
            
          Case xAlignRightEdge, xAlignRightOfCaption:
            
            If utTabInfo.TabPicture.Type = vbPicTypeBitmap And m_oITGTab.UseMaskColor Then
              
              Call DrawImage(m_lHDC, utTabInfo.TabPicture.Handle, g_pGetRGBFromOLE(m_oITGTab.PictureMaskColor), utFontRect.Right - iAdjustedIconSize - 2, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            Else
              
              Call m_oITGTab.pPaintPicture(utTabInfo.TabPicture, utFontRect.Right - iAdjustedIconSize, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            End If
            
            'm_oITGTab.pPaintPicture utTabInfo.TabPicture, utFontRect.Right - iAdjustedIconSize, iTmpY, iAdjustedIconSize, iAdjustedIconSize
            
            'shift the text to be drawn after the picture
            utFontRect.Right = (utFontRect.Right + 1 + 2) - iAdjustedIconSize
            
            'call the API for the text drawing
            DrawText m_lHDC, sTmp, -1, utFontRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
            
            'revert the changes so that the focus rectangle can be drawn for the whole tab's clickable area
            utFontRect.Right = (utFontRect.Right - 1 - 2) + iAdjustedIconSize
            
        End Select
        
      Else
        
        'call the API for the text drawing
        DrawText m_lHDC, sTmp, -1, utFontRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
        
      End If
      
            
      If m_oITGTab.bUserMode Then    'only if in the run mode
      
        If iCnt = m_iActiveTab And m_bIsFocused And m_oITGTab.ShowFocusRect Then
         
          'show the focus rectangle
          Call DrawFocusRect(m_lHDC, utFontRect)
        End If
      End If
    End With
  Next
  
  'draw the line in the empty area after all the property pages heads are drawn
  m_oITGTab.pLine m_oITGTab.aryTabs(m_iTabCount - 1).ClickableRect.Right, m_oITGTab.aryTabs(m_iTabCount - 1).ClickableRect.Bottom, m_iScaleWidth, m_oITGTab.aryTabs(m_iTabCount - 1).ClickableRect.Bottom, m_lTopLeftInnerBorderColor
 
End Sub


'function called to draw the tabs in the tabbed dialog style
Private Sub ITheme_DrawTabsTabbedDialog()
  Dim iCnt As Integer
  Dim iTabWidth As Integer
  Dim utFontRect As RECT    'used to perfom some calculations and passed to DrawText() Api
  Dim sTmp As String
  Dim utTabInfo As TabInfo
  Dim iTmpW As Integer
  Dim iTmpH As Integer
  Dim iTmpX As Integer
  Dim iTmpY As Integer
  Dim iOrigLeft As Integer
  Dim iOrigRight As Integer
 
  Dim iAdjustedIconSize As Integer
  
  iTabWidth = m_iScaleWidth / m_iTabCount    'remember iTabWidth is an integer ... so the result is automatically rounded
  
  'initialize the clickable items
  For iCnt = 0 To m_iTabCount - 1
    utTabInfo = m_oITGTab.aryTabs(iCnt)
    
    'no need to calculate the text size(like in property pages).... since this is a tabbed dialog style
    
    With utTabInfo.ClickableRect
      .Left = iCnt * iTabWidth
      .Right = .Left + iTabWidth
      
      If iCnt = m_iActiveTab Then
        If m_iActiveTabHeight > m_iInActiveTabHeight Then
          .Top = 0
        Else
          .Top = m_iInActiveTabHeight - m_iActiveTabHeight
        End If
        .Bottom = .Top + m_iActiveTabHeight
      Else
        If m_iInActiveTabHeight > m_iActiveTabHeight Then
          .Top = 0
        Else
          .Top = m_iActiveTabHeight - m_iInActiveTabHeight
        End If
        .Bottom = .Top + m_iInActiveTabHeight
      End If
      
    End With
    m_oITGTab.aryTabs(iCnt) = utTabInfo
  Next
  
  'if the last tab is shorter or longer than the usual size.. then adjust it to perfect size
  utTabInfo.ClickableRect.Right = m_iScaleWidth - 1
  m_oITGTab.aryTabs(iCnt - 1) = utTabInfo
  

  'added to prevent lines etc (we are filling the tab strip with the tab strip color)
  m_oITGTab.pLine 0, 0, m_iScaleWidth, IIf(m_iActiveTabHeight > m_iInActiveTabHeight, m_iActiveTabHeight, m_iInActiveTabHeight), m_oITGTab.TabStripBackColor, True, True
  
  
  
  'Now Draw Each Tab
  For iCnt = 0 To m_iTabCount - 1
    utTabInfo = m_oITGTab.aryTabs(iCnt)
    
   
    With utTabInfo.ClickableRect
      
      
      If iCnt = m_iActiveTab Then   'if we are drawing the active tab then
        
        If utTabInfo.Enabled Then
          Call m_oITGTab.pFillCurvedGradient(.Left + 1, .Top, .Right, .Bottom, m_lActiveTabBackStartColor, m_lActiveTabBackEndColor, 4, True, True)
        Else
          Call m_oITGTab.pFillCurvedGradient(.Left + 1, .Top, .Right, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor, 4, True, True)
        End If

        'top-left corner
        m_oITGTab.pLine .Left, .Top + 4, .Left + 4, .Top, m_lOuterBorderColor
        m_oITGTab.pLine .Left + 1, .Top + 4, .Left + 4, .Top + 1, m_lTopLeftInnerBorderColor
        m_oITGTab.pLine .Left + 2, .Top + 4, .Left + 4, .Top + 2, m_lTopLeftInnerBorderColor
        
        'top line
        m_oITGTab.pLine .Left + 4, .Top, .Right - 4, .Top, m_lOuterBorderColor
        m_oITGTab.pLine .Left + 4, .Top + 1, .Right - 4, .Top + 1, m_lTopLeftInnerBorderColor
        m_oITGTab.pLine .Left + 4, .Top + 2, .Right - 4, .Top + 2, m_lTopLeftInnerBorderColor
        
        'top-right corner
        m_oITGTab.pLine .Right - 4, .Top, .Right, .Top + 4, m_lOuterBorderColor
        m_oITGTab.pLine .Right - 4, .Top + 1, .Right - 1, .Top + 4, m_lBottomRightInnerBorderColor
        m_oITGTab.pLine .Right - 4, .Top + 2, .Right - 2, .Top + 4, m_lBottomRightInnerBorderColor
      
        'right line
        m_oITGTab.pLine .Right, .Top + 4, .Right, .Bottom + 1, m_lOuterBorderColor
        m_oITGTab.pLine .Right - 1, .Top + 4, .Right - 1, .Bottom + 2, m_lBottomRightInnerBorderColor
        m_oITGTab.pLine .Right - 2, .Top + 4, .Right - 2, .Bottom + 3, m_lBottomRightInnerBorderColor
        
        If utTabInfo.Enabled Then
          'bottom line  (actually we must erase the previously drawn background (since this is the active tab)
          m_oITGTab.pLine .Left + 3, .Bottom + 1, .Right - 2, .Bottom + 1, m_lActiveTabBackEndColor
          m_oITGTab.pLine .Left + 3, .Bottom + 2, .Right - 2, .Bottom + 2, m_lActiveTabBackEndColor
          m_oITGTab.pLine .Left + 3, .Bottom + 3, .Right - 2, .Bottom + 3, m_lActiveTabBackEndColor
        Else
          'bottom line  (actually we must erase the previously drawn background (since this is the active tab)
          m_oITGTab.pLine .Left + 3, .Bottom + 1, .Right - 2, .Bottom + 1, m_lDisabledTabBackColor
          m_oITGTab.pLine .Left + 3, .Bottom + 2, .Right - 2, .Bottom + 2, m_lDisabledTabBackColor
          m_oITGTab.pLine .Left + 3, .Bottom + 3, .Right - 2, .Bottom + 3, m_lDisabledTabBackColor
        End If

        
        'left line
        m_oITGTab.pLine .Left, .Top + 4, .Left, .Bottom + 1, m_lOuterBorderColor
        m_oITGTab.pLine .Left + 1, .Top + 4, .Left + 1, .Bottom + 2, m_lTopLeftInnerBorderColor
        m_oITGTab.pLine .Left + 2, .Top + 4, .Left + 2, .Bottom + 3, m_lTopLeftInnerBorderColor
        
        Set m_oITGTab.oFont = m_oITGTab.ActiveTabFont       'set the font
        
        
        'set foreground color (for text drawing)
        If utTabInfo.Enabled Then
          m_oITGTab.lForeColor = m_lActiveTabForeColor
        Else
          m_oITGTab.lForeColor = m_lDisabledTabForeColor
        End If
    
      Else  'we are drawing inactive tab
        
        If utTabInfo.Enabled Then
          Call m_oITGTab.pFillCurvedGradient(.Left + 1, .Top, .Right, .Bottom, m_lInActiveTabBackStartColor, m_lInActiveTabBackEndColor, 4, True, True)
        Else
          Call m_oITGTab.pFillCurvedGradient(.Left + 1, .Top, .Right, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor, 4, True, True)
        End If
        
        'top-left corner
        m_oITGTab.pLine .Left, .Top + 4, .Left + 4, .Top, m_lOuterBorderColor
        m_oITGTab.pLine .Left + 1, .Top + 4, .Left + 4, .Top + 1, m_lTopLeftInnerBorderColor
        
        'top line
        m_oITGTab.pLine .Left + 4, .Top, .Right - 4, .Top, m_lOuterBorderColor
        m_oITGTab.pLine .Left + 4, .Top + 1, .Right - 4, .Top + 1, m_lTopLeftInnerBorderColor
        
        'top-right corner
        m_oITGTab.pLine .Right - 4, .Top, .Right, .Top + 4, m_lOuterBorderColor
        m_oITGTab.pLine .Right - 4, .Top + 1, .Right - 1, .Top + 4, m_lBottomRightInnerBorderColor
        
        
        'right line
        m_oITGTab.pLine .Right, .Top + 4, .Right, .Bottom + 1, m_lOuterBorderColor
        m_oITGTab.pLine .Right - 1, .Top + 4, .Right - 1, .Bottom + 1, m_lBottomRightInnerBorderColor
        
        'bottom line
        m_oITGTab.pLine .Left, .Bottom + 1, .Right + 1, .Bottom + 1, m_lOuterBorderColor
        m_oITGTab.pLine .Left - 1, .Bottom + 2, .Right + 1, .Bottom + 2, m_lTopLeftInnerBorderColor
        m_oITGTab.pLine .Left - 2, .Bottom + 3, .Right + 2, .Bottom + 3, m_lTopLeftInnerBorderColor
        
        'left line
        m_oITGTab.pLine .Left, .Top + 4, .Left, .Bottom + 1, m_lOuterBorderColor
        m_oITGTab.pLine .Left + 1, .Top + 4, .Left + 1, .Bottom + 1, m_lTopLeftInnerBorderColor
      
        Set m_oITGTab.oFont = m_oITGTab.InActiveTabFont   'set the font
        
        'set forground color
        If utTabInfo.Enabled Then
          m_oITGTab.lForeColor = m_lInActiveTabForeColor
        Else
          m_oITGTab.lForeColor = m_lDisabledTabForeColor
        End If
            
      End If
      
      'do the adjustments for the border
      utFontRect.Left = .Left + 3
      utFontRect.Top = .Top + 3
      utFontRect.Bottom = .Bottom
      utFontRect.Right = .Right - 2
      
      If Not utTabInfo.TabPicture Is Nothing Then
        
        If utFontRect.Top + m_iIconSize > utFontRect.Bottom + 1 Then '+1 for minor adjustments
          'adjust if going out of current tab's bottom
          iAdjustedIconSize = (utFontRect.Bottom + 1) - utFontRect.Top
        Else
          iAdjustedIconSize = m_iIconSize
        End If
        
        iTmpY = utFontRect.Top + Round((utFontRect.Bottom - utFontRect.Top - iAdjustedIconSize) / 2)
                  
        Select Case m_oITGTab.PictureAlign
          Case xAlignLeftEdge:
            iTmpX = utFontRect.Left
            
            If iCnt = m_iActiveTab Then   'if active tab then give a comeup effect
              iTmpX = iTmpX + 1
              iTmpY = iTmpY - 1
              
              'make sure our adjustment dosen't make it out of the font area
              If iTmpY < utFontRect.Top Then iTmpY = utFontRect.Top
            End If
            
            If utTabInfo.TabPicture.Type = vbPicTypeBitmap And m_oITGTab.UseMaskColor Then
              
              Call DrawImage(m_lHDC, utTabInfo.TabPicture.Handle, g_pGetRGBFromOLE(m_oITGTab.PictureMaskColor), iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            Else
              
              Call m_oITGTab.pPaintPicture(utTabInfo.TabPicture, iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            End If
            
          Case xAlignRightEdge:
            iTmpX = utFontRect.Right - iAdjustedIconSize
            
            If iCnt = m_iActiveTab Then 'if active tab then give a comeup effect
              iTmpX = iTmpX - 1
              iTmpY = iTmpY - 1
              
              'make sure our adjustment dosen't make it out of the font area
              If iTmpY < utFontRect.Top Then iTmpY = utFontRect.Top
            End If
            
            If utTabInfo.TabPicture.Type = vbPicTypeBitmap And m_oITGTab.UseMaskColor Then
              
              Call DrawImage(m_lHDC, utTabInfo.TabPicture.Handle, g_pGetRGBFromOLE(m_oITGTab.PictureMaskColor), iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            Else
              
              Call m_oITGTab.pPaintPicture(utTabInfo.TabPicture, iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            End If
            
          Case xAlignLeftOfCaption:
            iOrigLeft = utFontRect.Left
          Case xAlignRightOfCaption:
            iOrigRight = utFontRect.Right
        End Select
        
      End If
      
      
      sTmp = utTabInfo.Caption
      
      'calculate the rect to draw the text, also modify the string to get ellipsis etc
      DrawText m_lHDC, sTmp, -1, utFontRect, DT_CALCRECT Or DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
      
      iTmpW = utFontRect.Right - utFontRect.Left + m_c_iFOCUS_RECT_AND_TEXT_DISTANCE
      iTmpH = utFontRect.Bottom - utFontRect.Top + m_c_iFOCUS_RECT_AND_TEXT_DISTANCE / 2
      
      
      'do the adjustments to center the text (both vertically and horizontally)
      utFontRect.Left = (utFontRect.Left - (m_c_iFOCUS_RECT_AND_TEXT_DISTANCE / 2)) + .Right / 2 - utFontRect.Right / 2
      utFontRect.Right = utFontRect.Left + iTmpW
      
      utFontRect.Top = utFontRect.Top + .Bottom / 2 - utFontRect.Bottom / 2
      utFontRect.Bottom = utFontRect.Top + iTmpH
      
      If Not utTabInfo.TabPicture Is Nothing Then
        
        Select Case m_oITGTab.PictureAlign
          Case xAlignLeftOfCaption:
            iTmpX = utFontRect.Left - iAdjustedIconSize - 1
            
            'make sure our adjustment dosen't make it out of the font area
            If iTmpX < iOrigLeft Then iTmpX = iOrigLeft
            
            If utTabInfo.TabPicture.Type = vbPicTypeBitmap And m_oITGTab.UseMaskColor Then
              
              Call DrawImage(m_lHDC, utTabInfo.TabPicture.Handle, g_pGetRGBFromOLE(m_oITGTab.PictureMaskColor), iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            Else
              
              Call m_oITGTab.pPaintPicture(utTabInfo.TabPicture, iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            End If
            
          Case xAlignRightOfCaption:
            iTmpX = utFontRect.Right + 1
            
            'make sure our adjustment dosen't make it out of the font area
            If iTmpX + iAdjustedIconSize > iOrigRight Then iTmpX = iOrigRight - iAdjustedIconSize
            
            If utTabInfo.TabPicture.Type = vbPicTypeBitmap And m_oITGTab.UseMaskColor Then
              
              Call DrawImage(m_lHDC, utTabInfo.TabPicture.Handle, g_pGetRGBFromOLE(m_oITGTab.PictureMaskColor), iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            Else
              
              Call m_oITGTab.pPaintPicture(utTabInfo.TabPicture, iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            End If

          
        End Select
      End If
      
      'Now draw the text
      DrawText m_lHDC, sTmp, -1, utFontRect, DT_SINGLELINE

      If m_oITGTab.bUserMode Then    'only if in the run mode

        If iCnt = m_iActiveTab And m_bIsFocused And m_oITGTab.ShowFocusRect Then
          'draw focus rectangle
          Call DrawFocusRect(m_lHDC, utFontRect)
        End If
      End If

    End With
    
  Next
  
  
  'store the larger tab height
  iCnt = IIf(m_iActiveTabHeight > m_iInActiveTabHeight, m_iActiveTabHeight, m_iInActiveTabHeight)
  
  'adjust the corners (whole tab control's corners)
  m_oITGTab.pLine 0, iCnt + 1, 0, iCnt + 4, m_lOuterBorderColor
  m_oITGTab.pLine m_iScaleWidth - 1, iCnt + 1, m_iScaleWidth - 1, iCnt + 4, m_lOuterBorderColor
  m_oITGTab.pLine m_iScaleWidth - 2, iCnt + 2, m_iScaleWidth - 2, iCnt + 4, m_lBottomRightInnerBorderColor
  m_oITGTab.pLine m_iScaleWidth - 3, iCnt + 3, m_iScaleWidth - 3, iCnt + 4, m_lBottomRightInnerBorderColor
 
End Sub

Private Sub ITheme_MouseDownHanlder(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)
  Dim iCnt As Integer
  Dim iX As Integer
  Dim iY As Integer
  Dim utTabInfo As TabInfo
  
  iX = CInt(sngX)
  iY = CInt(sngY)
  
  If iY > IIf(m_iActiveTabHeight > m_iInActiveTabHeight, m_iActiveTabHeight, m_iInActiveTabHeight) Then
      
      'if lower than the larger tab height then exit sub since anything lower than
      'active tab's height will not result in a tab switch
      Exit Sub
  End If
                                                                                                            
  
  'now go through each tab's rect to determine if the mouse was clicked within its boundaries
  For iCnt = 0 To m_iTabCount - 1
    utTabInfo = m_oITGTab.aryTabs(iCnt)
    If (iX >= utTabInfo.ClickableRect.Left And iX <= utTabInfo.ClickableRect.Right) And (iY >= utTabInfo.ClickableRect.Top And iY <= utTabInfo.ClickableRect.Bottom) And utTabInfo.Enabled Then
      
      'if its the active tab then no need to switch
      If m_iActiveTab <> iCnt Then
        m_oITGTab.ActiveTab = iCnt
      End If
      Exit Sub  'our work is finished .... no need to itirate further
    End If
  Next
  
End Sub

Private Sub ITheme_MouseMoveHanlder(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)
  'Do Nothing (no need in this theme)
End Sub

Private Sub ITheme_MouseUpHanlder(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)
  'Do Nothing (no need in this theme)
End Sub

'IMPORTANT: must be called before anything else is called... this function is called in the usercontrol
'the time the object is instantiated
Private Sub ITheme_SetControl(oITGTab As oTab)
  Set m_oITGTab = oITGTab
  m_oITGTab.bAutoRedraw = False
End Sub

Private Sub ITheme_ShowHideFocus()
  
  Call pCacheControlProperties    'get local copy for control properties that are used often
  
  Select Case m_oITGTab.TabStyle
    Case xStylePropertyPages
      ITheme_ShowHideFocusPropertyPages
    Case xStyleTabbedDialog
      ITheme_ShowHideFocusTabbedDialog
  End Select
End Sub

Private Sub ITheme_ShowHideFocusPropertyPages()
  Dim utFontRect As RECT
  Dim utTabInfo As TabInfo
  
  Dim iTmpHeight As Integer
  Dim iAdjustedIconSize As Integer
  
  
  If Not m_oITGTab.bUserMode Then       'only if in the run mode
    
    Exit Sub
    
  End If
  
  If Not m_oITGTab.ShowFocusRect Then   'only if Show Focus Rect is true for the control
      
    Exit Sub
    
  End If
  
  utTabInfo = m_oITGTab.aryTabs(m_iActiveTab)
  With utTabInfo.ClickableRect
    'do the adjustments for the border
    utFontRect.Left = .Left + 2
    utFontRect.Top = .Top + 2
    utFontRect.Bottom = .Bottom
    utFontRect.Right = .Right - 3
    
    If utTabInfo.Enabled Then             'done to allow proper drawing of focus rect
      m_oITGTab.lForeColor = m_lActiveTabForeColor
    Else
      m_oITGTab.lForeColor = m_lDisabledTabForeColor
    End If
    
    'show/hide the focus rectangle (drawn in XOR mode, so calling it again with same coords will erase it)
    Call DrawFocusRect(m_lHDC, utFontRect)
  End With
End Sub

Private Sub ITheme_ShowHideFocusTabbedDialog()
  Dim iTmpW As Integer
  Dim iTmpH As Integer
  Dim utFontRect As RECT
  Dim sTmp As String
  Dim utTabInfo As TabInfo
  Dim iAdjustedIconSize As Integer
  
  If Not m_oITGTab.bUserMode Then       'only if in the run mode
    
    Exit Sub
    
  End If
  
  If Not m_oITGTab.ShowFocusRect Then   'only if Show Focus Rect is true for the control
      
    Exit Sub
    
  End If
  
  
  
  utTabInfo = m_oITGTab.aryTabs(m_iActiveTab)
  With utTabInfo.ClickableRect
    'do the adjustments for the border
    utFontRect.Left = .Left + 3
    utFontRect.Top = .Top + 3
    utFontRect.Bottom = .Bottom
    utFontRect.Right = .Right - 2
    
    sTmp = utTabInfo.Caption
    
    'set the active tab font as current font... used to get proper values when calculating text size
    Set m_oITGTab.oFont = m_oITGTab.ActiveTabFont
    
    'calculate the rect to draw the text, and get proper string (including ellipsis etc)
    DrawText m_lHDC, sTmp, -1, utFontRect, DT_CALCRECT Or DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
    
    
    iTmpW = utFontRect.Right - utFontRect.Left + m_c_iFOCUS_RECT_AND_TEXT_DISTANCE
    iTmpH = utFontRect.Bottom - utFontRect.Top + m_c_iFOCUS_RECT_AND_TEXT_DISTANCE / 2
    
    
    'do the adjustments to center the text (both vertically and horizontally)
    utFontRect.Left = (utFontRect.Left - (m_c_iFOCUS_RECT_AND_TEXT_DISTANCE / 2)) + .Right / 2 - utFontRect.Right / 2
    utFontRect.Right = utFontRect.Left + iTmpW
    
    utFontRect.Top = utFontRect.Top + .Bottom / 2 - utFontRect.Bottom / 2
    utFontRect.Bottom = utFontRect.Top + iTmpH
    

      
    
    If utTabInfo.Enabled Then             'done to allow proper drawing of focus rect
      m_oITGTab.lForeColor = m_lActiveTabForeColor
    Else
      m_oITGTab.lForeColor = m_lDisabledTabForeColor
    End If
          
    'show/hide the focus rectangle (drawn in XOR mode, so calling it again with same coords will erase it)
    Call DrawFocusRect(m_lHDC, utFontRect)

  End With
End Sub

Private Sub ITheme_TimerEvent()
  'Do Nothing (will never be called for this theme)
End Sub


'function will reset all the colors back to default colors
Private Sub ITheme_ResetColorsToDefault()
  With m_oITGTab
    On Error Resume Next
    .ActiveTabBackStartColor = vbButtonFace
    .ActiveTabBackEndColor = vbButtonFace
    .ActiveTabForeColor = vbButtonText
    
    .InActiveTabBackStartColor = vbButtonFace
    .InActiveTabBackEndColor = vbButtonFace
    .InActiveTabForeColor = vbButtonText
    
    .TabStripBackColor = m_oITGTab.oAmbient.BackColor
    .DisabledTabBackColor = vbButtonFace
    .DisabledTabForeColor = vb3DDKShadow
    
    .TopLeftInnerBorderColor = vb3DHighlight
    .OuterBorderColor = vb3DDKShadow
    .BottomRightInnerBorderColor = vb3DShadow
    
    'Not Reset for this theme
    ' .HoverColorInverted
  End With
End Sub

'=====================================================================================================================

'==== Private Functions ==============================================================================================

'cache the properties from the control
'(this prevents trips to again and again fetch properties from the user control)
Private Sub pCacheControlProperties()
  m_lHDC = m_oITGTab.lhDC
  m_iActiveTab = m_oITGTab.ActiveTab
  m_iActiveTabHeight = m_oITGTab.ActiveTabHeight
  m_iInActiveTabHeight = m_oITGTab.InActiveTabHeight
  m_iScaleWidth = m_oITGTab.iScaleWidth
  m_iScaleHeight = m_oITGTab.iScaleHeight
  m_iTabCount = m_oITGTab.TabCount
  m_bIsFocused = m_oITGTab.bIsFocused
  m_lOuterBorderColor = m_oITGTab.OuterBorderColor
  m_lTopLeftInnerBorderColor = m_oITGTab.TopLeftInnerBorderColor
  m_lBottomRightInnerBorderColor = m_oITGTab.BottomRightInnerBorderColor
  
  m_lActiveTabForeColor = m_oITGTab.ActiveTabForeColor
  m_lActiveTabBackStartColor = m_oITGTab.ActiveTabBackStartColor
  m_lActiveTabBackEndColor = m_oITGTab.ActiveTabBackEndColor
  m_lInActiveTabForeColor = m_oITGTab.InActiveTabForeColor
  m_lInActiveTabBackStartColor = m_oITGTab.InActiveTabBackStartColor
  m_lInActiveTabBackEndColor = m_oITGTab.InActiveTabBackEndColor
  m_lDisabledTabBackColor = m_oITGTab.DisabledTabBackColor
  m_lDisabledTabForeColor = m_oITGTab.DisabledTabForeColor
  
  
  'Get System's default size for a Icon.
  If m_oITGTab.PictureSize = xSizeSmall Then
    m_iIconSize = GetSystemMetrics(SM_CXSMICON)
  Else
    m_iIconSize = GetSystemMetrics(SM_CXICON)
  End If
  
End Sub

