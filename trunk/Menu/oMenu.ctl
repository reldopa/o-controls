VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl oMenu 
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3660
   LockControls    =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   3660
   Begin MSComCtl2.FlatScrollBar fScroll 
      Height          =   3135
      Left            =   2820
      TabIndex        =   2
      Top             =   0
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   5530
      _Version        =   393216
      LargeChange     =   30
      Orientation     =   1179648
      SmallChange     =   30
   End
   Begin VB.PictureBox bg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   30000
      Left            =   0
      ScaleHeight     =   30000
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   0
      Width           =   2835
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   510
         TabIndex        =   1
         Top             =   150
         Width           =   555
      End
      Begin VB.Image icon 
         Height          =   255
         Index           =   0
         Left            =   180
         Stretch         =   -1  'True
         Top             =   120
         Width           =   255
      End
      Begin VB.Image img 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   0
         Left            =   30
         Picture         =   "oMenu.ctx":0000
         Stretch         =   -1  'True
         Top             =   30
         Width           =   2805
      End
   End
End
Attribute VB_Name = "oMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Default Property Values:
Const m_def_MenuHeight = 495
Const m_def_CustomMenu = False
Const m_def_MenuBackColor = &HFF8080
Const m_def_MenuCaptionAlignment = 0

'Property Variables:
Dim m_MenuHeight As Integer
Dim m_CustomMenu As Boolean
Dim m_MenuBackColor As OLE_COLOR
Dim m_MenuCaptionAlignment As e_MenuCaptionAlignment
Dim m_MouseOverPicture As Picture

'Event Declarations:
Event Click(MenuKey As String) 'MappingInfo=img(0),img,0,Click
Event DblClick(MenuKey As String)  'MappingInfo=img(0),img,0,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
 
Public Enum e_MenuCaptionAlignment
    CaptionCenter = 0
    CaptionLeft = 1
    CaptionRight = 2
End Enum

Private objMenu As MenuItem
Private colMenu As MenuCollection

Property Get MenuItems() As MenuCollection
    Set MenuItems = colMenu
    ApplyMenuChanges
End Property

Public Sub Refresh()
Dim i As Integer
    For i = 0 To img.Count - 1
        Set img(i).Picture = UserControl.Picture
        img(i).Refresh
    Next i
End Sub

Private Sub ApplyMenuChanges()
Dim i As Integer
    For i = 0 To colMenu.Count - 1
        lbl(i).Enabled = colMenu.Item(i + 1).Enabled
        img(i).Enabled = colMenu.Item(i + 1).Enabled
    Next i
End Sub
Private Sub fScroll_Change()
    bg.Top = -fScroll.Value
End Sub

Private Sub fScroll_Scroll()
    bg.Top = -fScroll.Value
End Sub

Private Sub icon_Click(Index As Integer)
    Call img_Click(Index)
End Sub

Private Sub img_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
    For i = 0 To img.Count - 1
        img(i).BorderStyle = 0
        If m_CustomMenu = True Then
            Set img(i).Picture = colMenu(i + 1).Picture
        Else
            If IsNumeric(colMenu(i + 1).Picture) = False Then Set img(i).Picture = Picture Else Set img(i).Picture = colMenu.Item(i + 1).Picture
        End If
    Next i
    img(Index).BorderStyle = 1
    If m_CustomMenu = False Then
        Set img(Index).Picture = MouseOverPicture
    End If

End Sub

Private Sub lbl_Click(Index As Integer)
    Call img_Click(Index)
End Sub

Private Sub UserControl_Initialize()
    Set colMenu = New MenuCollection
End Sub

Private Sub UserControl_Resize()
Dim i As Integer
Dim Y As Integer
Dim yIcon As Integer

On Error Resume Next

    bg.Left = 0
    bg.Top = 0
    bg.Width = UserControl.Width - fScroll.Width
    fScroll.Left = bg.Width
    fScroll.Height = UserControl.Height
    
    'set all menu to invisible
    For i = 0 To img.Count - 1
        lbl(i).Visible = False
        icon(i).Visible = False
        img(i).Visible = False
    Next i
    Y = 15
    'Show Parent
    For i = 0 To colMenu.Count - 1
        If colMenu.Item(i + 1).Relative = "" Then
            img(i).Height = MenuHeight
            img(i).Left = 15
            img(i).Width = bg.Width - 30
            img(i).Top = Y
            icon(i).Left = 90
            icon(i).Top = ((img(i).Height - icon(i).Height) / 2) + Y
            If MenuCaptionAlignment = 0 Then
                lbl(i).Alignment = 2
                lbl(i).Left = (img(i).Width - lbl(i).Width) / 2 'center
            ElseIf MenuCaptionAlignment = 1 Then
                lbl(i).Alignment = 0
                lbl(i).Left = icon(i).Left + icon(i).Width + 15 ' left
            Else
                lbl(i).Alignment = 1
                lbl(i).Left = img(i).Width - (lbl(i).Width + 15) 'right
            End If
            lbl(i).Top = ((img(i).Height - lbl(i).Height) / 2) + Y
            Y = Y + img(i).Height
            lbl(i).Visible = True
            icon(i).Visible = True
            img(i).Visible = True
            colMenu.Item(i + 1).Top = img(i).Top
            colMenu.Item(i + 1).Visible = True
        End If
    Next i
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = bg.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    bg.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = bg.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    bg.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Sub AddMenu(Optional ByVal Relative As String, Optional ByVal nLevel As Integer, _
Optional ByVal sIndex As Variant, Optional ByVal sText As String, Optional ByVal IconImage As StdPicture, _
Optional ByVal MenuImage As StdPicture, Optional ByVal lForeColor As OLE_COLOR, _
Optional ByVal lBold As Boolean, Optional ByVal nFontSize As Integer)

Dim i As Integer
'On Error Resume Next
        
    colMenu.Add sIndex, Relative, nLevel, sText, IconImage, MenuImage, lForeColor
    i = colMenu.Count - 1
        If i > 0 Then
            Load lbl(i)
            Load icon(i)
            Load img(i)
        End If
    If lForeColor = 0 Then lbl(i).ForeColor = &H80000012 Else lbl(i).ForeColor = colMenu.Item(i + 1).ForeColor
    If CustomMenu = True Then
        Set img(i).Picture = MenuImage
    Else
       If IsNumeric(MenuImage) = False Then Set img(i).Picture = Picture Else Set img(i).Picture = colMenu.Item(i + 1).Picture
    End If
        lbl(i).Caption = colMenu(i + 1).Caption
        img(i).Height = MenuHeight
        Set icon(i).Picture = colMenu(i + 1).IconPic
    
        
    
    Call UserControl_Resize
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Sub Clear()
Dim i As Integer
    For i = 0 To img.Count - 1
            If i > 0 Then
                Unload img(i)
                Unload icon(i)
                Unload lbl(i)
            End If
    Next i
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
        
    m_MenuHeight = m_def_MenuHeight
    m_CustomMenu = m_def_CustomMenu
    m_MenuBackColor = m_def_MenuBackColor
    m_MenuCaptionAlignment = m_def_MenuCaptionAlignment
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer
    bg.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    bg.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    m_CustomMenu = PropBag.ReadProperty("CustomMenu", m_def_CustomMenu)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_MenuBackColor = PropBag.ReadProperty("MenuBackColor", m_def_MenuBackColor)
    Set lbl(0).Font = PropBag.ReadProperty("MenuFont", Ambient.Font)
    m_MenuHeight = PropBag.ReadProperty("MenuHeight", m_def_MenuHeight)
    img(Index) = PropBag.ReadProperty("MenuItem", Nothing)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Set MouseOverPicture = PropBag.ReadProperty("MouseOverPicture", Nothing)
    MenuCaptionAlignment = PropBag.ReadProperty("MenuCaptionAlignment", m_def_MenuCaptionAlignment)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer
    Call PropBag.WriteProperty("BackColor", bg.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderStyle", bg.BorderStyle, 0)
    Call PropBag.WriteProperty("CustomMenu", m_CustomMenu, m_def_CustomMenu)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MenuBackColor", m_MenuBackColor, m_def_MenuBackColor)
    Call PropBag.WriteProperty("MenuCaptionAlignment", m_MenuCaptionAlignment, m_def_MenuCaptionAlignment)
    Call PropBag.WriteProperty("MenuFont", lbl(Index).Font, Ambient.Font)
    Call PropBag.WriteProperty("MenuHeight", m_MenuHeight, m_def_MenuHeight)
    Call PropBag.WriteProperty("MenuItem", img(Index), Nothing)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("MouseOverPicture", MouseOverPicture, Nothing)
End Sub

Private Sub ShrinkMenu(sIndex As Integer)
Dim i As Integer, n As Integer
Dim Y As Integer
Dim TotalRemHeight As Integer, ChildHeight As Integer, TotalHeight As Integer
Dim iY As Integer
Dim skey As String, sRelative As String

    colMenu.Item(sIndex + 1).Expanded = False
    Y = colMenu.Item(sIndex + 1).Top + img(sIndex).Height
    skey = colMenu.Item(sIndex + 1).Key
    sRelative = colMenu.Item(sIndex + 1).Relative
    'Collect and compute height for child
    For i = 0 To colMenu.Count - 1
        If InStr(colMenu.Item(i + 1).sDash, skey & "\") > 0 And skey <> colMenu.Item(i + 1).Key And colMenu.Item(i + 1).Visible = True Then
            ChildHeight = ChildHeight + img(i).Height
        End If
    Next i

    TotalHeight = TotalNodeHeight
    TotalRemHeight = TotalNodeHeight - (Y - 15)
    
    iY = TotalNodeHeight
    
    Call HideChild(sIndex, Y, skey)
    Call UpperNode(colMenu.Item(sIndex + 1).Top + img(sIndex).Height, TotalHeight, ChildHeight * -1)
End Sub

Private Sub HideChild(sIndex As Integer, Y As Integer, Key As String)
Dim i  As Integer
    For i = 0 To colMenu.Count - 1
        If InStr(colMenu.Item(i + 1).sDash, Key & "\") > 0 And Key <> colMenu.Item(i + 1).Key And colMenu.Item(i + 1).Visible = True Then
            Call NodeVisible(i, False, Y)
            colMenu.Item(i + 1).Expanded = False
            Y = Y + img(i).Height
        End If
    Next i
End Sub

Private Sub UpperNode(nStart As Integer, nEnd As Integer, nChildHeight As Integer)
Dim nHeight As Integer
Dim i As Integer, n As Integer
Dim nVal(1 To 10000) As Integer
Dim nIndex As Integer
Dim sval As String

    n = 1
    Do Until nStart >= nEnd
        For i = 0 To colMenu.Count - 1
            If colMenu.Item(i + 1).Top = nStart Then
                If colMenu.Item(i + 1).Visible = True Then
                     nVal(n) = colMenu.Item(i + 1).Index
                    sval = sval & colMenu.Item(i + 1).Index & ","
                     n = n + 1
                     Exit For
                End If
            End If
        Next i
        nStart = nStart + img(0).Height
    Loop
        For nIndex = 1 To n - 1
            Call NodeVisible(nVal(nIndex), True, Abs(nChildHeight + colMenu.Item(nVal(nIndex) + 1).Top))
        Next nIndex
End Sub



Private Sub ExpandMenu(sIndex As Integer)
Dim i As Integer, n As Integer
Dim Y As Integer
Dim TotalRemHeight As Integer, ChildHeight As Integer, TotalHeight As Integer
Dim iY As Integer
Dim skey As String
Dim sUpNode As String, sLowNode As String
    colMenu.Item(sIndex + 1).Expanded = True
    Y = colMenu.Item(sIndex + 1).Top + img(sIndex).Height
    skey = colMenu.Item(sIndex + 1).Key
    
    'Collect and compute height for child
    For i = 0 To colMenu.Count - 1
        If skey = colMenu.Item(i + 1).Relative Then
            ChildHeight = ChildHeight + img(i).Height
        End If
    Next i

    TotalHeight = TotalNodeHeight + ChildHeight
    TotalRemHeight = TotalNodeHeight - (Y - 15)
     
    iY = TotalNodeHeight
    Do Until iY >= TotalHeight
       Call LowerNode(Y, TotalHeight, ChildHeight)
       Call ShowChild(sIndex, colMenu.Item(sIndex + 1).Top + img(sIndex).Height, skey)
        iY = iY + ChildHeight
    Loop
    
End Sub

Private Sub LowerNode(nStart As Integer, nEnd As Integer, nChildHeight As Integer)
Dim nHeight As Integer
Dim i As Integer, n As Integer
Dim nVal(1 To 10000) As Integer
Dim nIndex As Integer
    
    n = 1
    Do Until nStart >= nEnd
        For i = 0 To colMenu.Count - 1
            If colMenu.Item(i + 1).Top = nStart Then
                If colMenu.Item(i + 1).Visible = True Then
                     nVal(n) = colMenu.Item(i + 1).Index
                     n = n + 1
                     Exit For
                End If
            End If
        Next i
        nStart = nStart + img(0).Height
    Loop
        For nIndex = 1 To n - 1
            Call NodeVisible(nVal(nIndex), True, nChildHeight + colMenu.Item(nVal(nIndex) + 1).Top)
        Next nIndex
End Sub

Private Sub ShowChild(sIndex As Integer, Y As Integer, skey As String)
Dim i  As Integer
    For i = sIndex To colMenu.Count - 1
        If skey = colMenu.Item(i + 1).Relative Then
            Call NodeVisible(i, True, Y)
            Y = Y + img(i).Height
        End If
    Next i
End Sub

Private Function TotalNodeHeight() As Integer
Dim i As Integer
    For i = 0 To colMenu.Count - 1
        If colMenu(i + 1).Visible = True Then
            TotalNodeHeight = TotalNodeHeight + img(i).Height
        End If
    Next i
End Function

Private Sub NodeVisible(sIndex As Integer, fVisible As Boolean, nTop As Integer)
    lbl(sIndex).Top = ((img(sIndex).Height - lbl(sIndex).Height) / 2) + nTop
    icon(sIndex).Top = ((img(sIndex).Height - icon(sIndex).Height) / 2) + nTop
    img(sIndex).Top = nTop
    lbl(sIndex).Visible = fVisible
    icon(sIndex).Visible = fVisible
    img(sIndex).Visible = fVisible
    colMenu.Item(sIndex + 1).Top = nTop
    colMenu.Item(sIndex + 1).Visible = fVisible
End Sub

Private Function CheckForChild(sIndex As Integer) As Boolean
Dim i As Integer
    For i = 0 To colMenu.Count - 1
        If colMenu.Item(i + 1).Relative = colMenu.Item(sIndex + 1).Key Then
            CheckForChild = True
        End If
    Next i
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get MouseOverPicture() As Picture
    Set MouseOverPicture = m_MouseOverPicture
End Property

Public Property Set MouseOverPicture(ByVal New_MouseOverPicture As Picture)
    Set m_MouseOverPicture = New_MouseOverPicture
    PropertyChanged "MouseOverPicture"
End Property

Private Sub img_Click(Index As Integer)
Dim sMenuKey As String
    If colMenu.Item(Index + 1).Enabled = False Then Exit Sub
    If colMenu.Item(Index + 1).Expanded = False And CheckForChild(Index) = True Then
        Call ExpandMenu(Index)
    ElseIf colMenu.Item(Index + 1).Expanded = True And CheckForChild(Index) = True Then
        Call ShrinkMenu(Index)
    Else
        sMenuKey = colMenu.Item(Index + 1).Key
        RaiseEvent Click(sMenuKey)
    End If
End Sub

Private Sub img_DblClick(Index As Integer)
Dim sMenuKey As String
    sMenuKey = colMenu.Item(Index + 1).Key
    RaiseEvent DblClick(sMenuKey)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get MenuHeight() As Integer
    MenuHeight = m_MenuHeight
End Property

Public Property Let MenuHeight(ByVal New_MenuHeight As Integer)
    m_MenuHeight = New_MenuHeight
    PropertyChanged "MenuHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CustomMenu() As Boolean
    CustomMenu = m_CustomMenu
End Property

Public Property Let CustomMenu(ByVal New_CustomMenu As Boolean)
    m_CustomMenu = New_CustomMenu
    PropertyChanged "CustomMenu"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,Font
Public Property Get MenuCaptionAlignment() As e_MenuCaptionAlignment
    MenuCaptionAlignment = m_MenuCaptionAlignment
End Property

Public Property Let MenuCaptionAlignment(ByVal New_MenuCaptionAlignment As e_MenuCaptionAlignment)
    m_MenuCaptionAlignment = New_MenuCaptionAlignment
    PropertyChanged "MenuCaptionAlignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl(0),lbl,0,Font
Public Property Get MenuFont() As Font
Attribute MenuFont.VB_Description = "Returns a Font object."
    Set MenuFont = lbl(0).Font
End Property

Public Property Set MenuFont(ByVal New_MenuFont As Font)
    Set lbl(0).Font = New_MenuFont
    PropertyChanged "MenuFont"
End Property





