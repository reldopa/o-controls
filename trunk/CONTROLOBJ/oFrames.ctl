VERSION 5.00
Begin VB.UserControl oFrames 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F7F2E6&
   CanGetFocus     =   0   'False
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5085
   ControlContainer=   -1  'True
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HitBehavior     =   0  'None
   LockControls    =   -1  'True
   ScaleHeight     =   284
   ScaleMode       =   0  'User
   ScaleWidth      =   339
End
Attribute VB_Name = "oFrames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'*************************************************************
'   Required Type Definitions
'*************************************************************

Public Enum com_StyleConst
    XPDefault = 0
    GradientFrame = 1
End Enum

'xp theme
Public Enum com_ThemeConst
    Blue = 0
    Silver = 1
    Olive = 2
    Visual2005 = 3
    Norton2004 = 4
    Custom = 5
End Enum

'icon aligment
Public Enum IconAlignConst
    vbLeftAligment = 0
    vbRightAligment = 1
End Enum

Private Const ALTERNATE = 1      ' ALTERNATE and WINDING are
Private Const WINDING = 2        ' constants for FillMode.
Private Const BLACKBRUSH = 4     ' Constant for brush type.
Private Const WHITE_BRUSH = 0    ' Constant for brush type.


'members
Private m_FrameColor            As OLE_COLOR
Private m_BackColor             As OLE_COLOR
Private m_FillColor             As OLE_COLOR
Private m_Caption               As String
Private m_TextBoxHeight         As Long
Private m_TextHeight            As Long
Private m_TextWidth             As Long
Private m_Height                As Long
Private m_TextColor             As Long
Private m_Alignment             As Long
Private m_Font                  As StdFont
Private m_RoundedCorner         As Boolean
Private m_Style                 As com_StyleConst
Private m_Icon                  As StdPicture
Private m_IconSize              As Integer
Private m_IconAlignment         As IconAlignConst
Private m_ThemeColor            As com_ThemeConst
Private m_ColorTo               As OLE_COLOR
Private m_ColorFrom             As OLE_COLOR
Private m_Indentation           As Integer
Private m_Space                 As Integer

Private Const DT_CENTER = &H1
Private Const DT_BOTTOM = &H8
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_NOCLIP = &H100
Private Const DT_LEFT = &H0

Private com_TextBoxCenter       As Long
Private com_TextDrawParams      As Long
Private com_ColorTo             As OLE_COLOR
Private com_ColorFrom           As OLE_COLOR
Private com_ColorBorderPic      As OLE_COLOR

Private com_Lpp As POINT


'==========================================================================
' Init, Read & Write UserControl
'==========================================================================
Private Sub UserControl_InitProperties()
    'Set default properties
    m_Caption = Ambient.DisplayName
    m_BackColor = &HF7F2E6
    m_FillColor = &HF7F2E6
    m_RoundedCorner = False
    m_Style = GradientFrame
    'm_ThemeColor = Blue
    'Call SetDefaultThemeColor(m_ThemeColor)
    m_TextColor = vbBlack
    m_FrameColor = vbWhite
    m_TextBoxHeight = 22
    m_Font = "Tahoma"
    SetTextDrawParams
End Sub

Private Sub UserControl_Initialize()
    Set m_Font = New StdFont
    Set UserControl.Font = m_Font
    UserControl.Font.Bold = True
    m_IconSize = 16
    m_ThemeColor = Blue
    Call SetDefaultThemeColor(m_ThemeColor)
    m_TextBoxHeight = 22
    m_Alignment = vbLeftAligment
    m_IconAlignment = vbLeftAligment
End Sub

Private Sub UserControl_Resize()
    PaintFrame
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_FrameColor = .ReadProperty("FrameColor", vbBlack)
        m_BackColor = .ReadProperty("BackColor", &HF7F2E6)
        m_FillColor = .ReadProperty("FillColor", &HF7F2E6)
        m_Style = .ReadProperty("Style", GradientFrame)
        m_RoundedCorner = .ReadProperty("RoundedCorner", True)
        m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
        m_TextBoxHeight = .ReadProperty("TextBoxHeight", 22)
        m_TextColor = .ReadProperty("TextColor", vbBlack)
        m_Alignment = .ReadProperty("Alignment", vbCenter)
        m_IconAlignment = .ReadProperty("IconAlignment", vbLeftAligment)
        Set m_Font = .ReadProperty("Font", Ambient.Font)
        Set m_Icon = .ReadProperty("Picture", Nothing)
        m_IconSize = .ReadProperty("IconSize", 16)
        m_ThemeColor = .ReadProperty("ThemeColor", Blue)
        m_ColorFrom = .ReadProperty("ColorFrom", RGB(129, 169, 226))
        m_ColorTo = .ReadProperty("ColorTo", RGB(221, 236, 254))
    End With
    'Add properties
    UserControl.BackColor = m_BackColor
    'SetTextBoxRect
    SetTextDrawParams
    SetFont m_Font
    Call SetDefaultThemeColor(m_ThemeColor)
    'Paint control
    PaintFrame
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "FrameColor", m_FrameColor, vbBlack
        .WriteProperty "BackColor", m_BackColor, &HF7F2E6
        .WriteProperty "FillColor", m_FillColor, &HF7F2E6
        .WriteProperty "Style", m_Style, GradientFrame
        .WriteProperty "RoundedCorner", m_RoundedCorner, True
        .WriteProperty "Caption", m_Caption, Ambient.DisplayName
        .WriteProperty "TextBoxHeight", m_TextBoxHeight, 22
        .WriteProperty "TextColor", m_TextColor, vbBlack
        .WriteProperty "Alignment", m_Alignment, vbCenter
        .WriteProperty "IconAlignment", m_IconAlignment, vbLeftAligment
        .WriteProperty "Font", m_Font, Ambient.Font
        .WriteProperty "Picture", m_Icon, Nothing
        .WriteProperty "IconSize", m_IconSize, 16
        .WriteProperty "ThemeColor", m_ThemeColor, Blue
        .WriteProperty "ColorFrom", m_ColorFrom, RGB(129, 169, 226)
        .WriteProperty "ColorTo", m_ColorTo, RGB(221, 236, 254)
    End With
End Sub

'==========================================================================
' Properties
'==========================================================================
Public Property Let FrameColor(ByRef new_FrameColor As OLE_COLOR)
    m_FrameColor = new_FrameColor
    If m_ThemeColor = Custom Then com_ColorBorderPic = m_FrameColor
    PropertyChanged "FrameColor"
    PaintFrame
End Property

Public Property Get FrameColor() As OLE_COLOR
    FrameColor = m_FrameColor
End Property

Public Property Let FillColor(ByRef new_FillColor As OLE_COLOR)
    m_FillColor = new_FillColor
    PropertyChanged "FillColor"
    PaintFrame
End Property

Public Property Get FillColor() As OLE_COLOR
    FillColor = m_FillColor
End Property


Public Property Let RoundedCorner(ByRef new_RoundedCorner As Boolean)
    m_RoundedCorner = new_RoundedCorner
    PropertyChanged "RoundedCorner"
    PaintFrame
End Property

Public Property Get RoundedCorner() As Boolean
    RoundedCorner = m_RoundedCorner
End Property

Public Property Let Caption(ByRef new_caption As String)
    m_Caption = new_caption
    PaintFrame
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Alignment(ByRef new_Alignment As AlignmentConstants)
    m_Alignment = new_Alignment
    SetTextDrawParams
    PropertyChanged "Alignment"
    PaintFrame
End Property

Public Property Get Alignment() As AlignmentConstants
    Alignment = m_Alignment
End Property

Public Property Let Style(ByRef New_Style As com_StyleConst)
    m_Style = New_Style
    PropertyChanged "Style"
    SetDefault
    PaintFrame
End Property

Public Property Get Style() As com_StyleConst
    Style = m_Style
End Property

Public Property Let TextBoxHeight(ByRef new_TextBoxHeight As Long)
    m_TextBoxHeight = new_TextBoxHeight
    PropertyChanged "TextBoxHeight"
    PaintFrame
End Property

Public Property Get TextBoxHeight() As Long
    TextBoxHeight = m_TextBoxHeight
End Property

Public Property Let TextColor(ByRef new_TextColor As OLE_COLOR)
    m_TextColor = new_TextColor
    PropertyChanged "TextColor"
    PaintFrame
End Property

Public Property Get TextColor() As OLE_COLOR
    TextColor = m_TextColor
End Property

Public Property Let BackColor(ByRef New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    UserControl.BackColor = m_BackColor
    PropertyChanged "BackColor"
    PaintFrame
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Set Font(ByRef New_Font As StdFont)
    SetFont New_Font
    PropertyChanged "Font"
    PaintFrame
End Property

Public Property Let Font(ByRef New_Font As StdFont)
    SetFont New_Font
    PropertyChanged "Font"
    PaintFrame
End Property
Public Property Get Font() As StdFont
    Set Font = m_Font
End Property

Public Property Get Picture() As StdPicture
    Set Picture = m_Icon
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
    Set m_Icon = New_Picture
    PropertyChanged "Picture"
    PaintFrame
End Property

Public Property Get IconSize() As Integer
    IconSize = m_IconSize
End Property

Public Property Let IconSize(ByVal New_Value As Integer)
    m_IconSize = New_Value
    PropertyChanged "IconSize"
    PaintFrame
End Property

Public Property Let IconAlignment(ByRef new_IconAlignment As IconAlignConst)
    m_IconAlignment = new_IconAlignment
    PropertyChanged "IconAlignment"
    PaintFrame
End Property

Public Property Get IconAlignment() As IconAlignConst
    IconAlignment = m_IconAlignment
End Property

Public Property Get ThemeColor() As com_ThemeConst
    ThemeColor = m_ThemeColor
End Property

Public Property Let ThemeColor(ByVal vData As com_ThemeConst)
    If m_ThemeColor <> vData Then
        m_ThemeColor = vData
        Call SetDefaultThemeColor(m_ThemeColor)
        PaintFrame
        PropertyChanged "ThemeColor"
    End If
End Property

Public Property Get ColorFrom() As OLE_COLOR
    ColorFrom = m_ColorFrom
End Property

Public Property Let ColorFrom(ByRef new_ColorFrom As OLE_COLOR)
    m_ColorFrom = new_ColorFrom
    PropertyChanged "ColorFrom"
    com_ColorFrom = m_ColorFrom
    PaintFrame
End Property

Public Property Get ColorTo() As OLE_COLOR
    ColorTo = m_ColorTo
End Property

Public Property Let ColorTo(ByRef new_ColorTo As OLE_COLOR)
    m_ColorTo = new_ColorTo
    PropertyChanged "ColorTo"
    com_ColorTo = m_ColorTo
    PaintFrame
End Property

Private Sub SetTextDrawParams()
    'Set text draw params using m_Alignment
    If m_Alignment = vbLeftJustify Then
        com_TextDrawParams = DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
    ElseIf m_Alignment = vbRightJustify Then
        com_TextDrawParams = DT_RIGHT Or DT_SINGLELINE Or DT_VCENTER
    Else
        com_TextDrawParams = DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
    End If
End Sub

Private Sub SetFont(ByRef New_Font As StdFont)
    With m_Font
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Name = New_Font.Name
        .Size = New_Font.Size
    End With
    Set UserControl.Font = m_Font
End Sub

'==========================================================================
' Functions and subroutines
'==========================================================================

Private Sub SetDefaultThemeColor(ThemeType As Long)
    Select Case ThemeType
        Case 0 '"NormalColor"
            com_ColorFrom = RGB(129, 169, 226)
            com_ColorTo = RGB(221, 236, 254)
            com_ColorBorderPic = RGB(0, 0, 128)
        Case 1 '"Metallic"
            com_ColorFrom = RGB(153, 151, 180)
            com_ColorTo = RGB(244, 244, 251)
            com_ColorBorderPic = RGB(75, 75, 111)
        Case 2 '"HomeStead"
            com_ColorFrom = RGB(181, 197, 143)
            com_ColorTo = RGB(247, 249, 225)
            com_ColorBorderPic = RGB(63, 93, 56)
        Case 3 '"Visual2005"
            com_ColorFrom = RGB(194, 194, 171)
            com_ColorTo = RGB(248, 248, 242)
            com_ColorBorderPic = RGB(145, 145, 115)
        Case 4 '"Norton2004"
            com_ColorFrom = RGB(217, 172, 1)
            com_ColorTo = RGB(255, 239, 165)
            com_ColorBorderPic = RGB(117, 91, 30)
        Case 5  'Custom
            com_ColorFrom = m_ColorFrom
            com_ColorTo = m_ColorTo
            com_ColorBorderPic = m_FrameColor
        Case Else
            com_ColorFrom = RGB(153, 151, 180)
            com_ColorTo = RGB(244, 244, 251)
            com_ColorBorderPic = RGB(75, 75, 111)
    End Select
    
    m_ColorTo = com_ColorTo
    m_ColorFrom = com_ColorFrom
End Sub

Private Sub PaintFrame()
    Dim R As RECT, R_Caption As RECT
    Dim p_left As Long, iX As Integer, iY As Integer
    
    m_Height = 3
    m_Indentation = 15
    m_Space = 6
    iX = 0
    iY = 0
    
    'Clear user control
    UserControl.Cls
    
    'Set caption height and width
    '----------------------------
    If Len(m_Caption) <> 0 Then
        m_TextWidth = UserControl.TextWidth(m_Caption)
        m_TextHeight = UserControl.TextHeight(m_Caption)
        com_TextBoxCenter = m_TextHeight / 2
    Else
        com_TextBoxCenter = 0
    End If

    Select Case m_Style
        Case Is = XPDefault
            
            'Draw border rectangle
            UserControl.FillColor = m_FillColor
            UserControl.ForeColor = m_FrameColor
            
            If m_RoundedCorner = False Then
                RoundRect UserControl.hdc, 0&, com_TextBoxCenter, UserControl.ScaleWidth, UserControl.ScaleHeight, 0&, 0&
            Else
                RoundRect UserControl.hdc, 0&, com_TextBoxCenter, UserControl.ScaleWidth, UserControl.ScaleHeight, 10&, 10&
            End If
            
            If Len(m_Caption) <> 0 Then
                
                If m_Alignment = vbLeftJustify Then
                    p_left = m_Indentation
                ElseIf m_Alignment = vbRightJustify Then
                    p_left = UserControl.ScaleWidth - m_TextWidth - m_Indentation - m_Space
                Else
                    p_left = (UserControl.ScaleWidth - m_TextWidth) / 2
                End If
                
                'Draw a line
                UserControl.ForeColor = UserControl.FillColor
                MoveToEx UserControl.hdc, p_left, com_TextBoxCenter, com_Lpp
                LineTo UserControl.hdc, p_left + m_TextWidth + m_Space, com_TextBoxCenter
                
                'set caption rect
                SetRect R_Caption, p_left + m_Space / 2, 0, m_TextWidth + p_left + m_Space / 2, m_TextHeight
            End If
           
        Case Is = GradientFrame
            'Draw border rectangle
            UserControl.FillColor = m_FillColor 'BlendColors(com_ColorFrom, vbWhite)
            UserControl.ForeColor = com_ColorBorderPic
            
            'UserControl.FillColor = m_FillColor
            'UserControl.ForeColor = com_ColorBorderPic
            
            com_TextBoxCenter = m_TextBoxHeight / 2

            If m_RoundedCorner = False Then
                RoundRect UserControl.hdc, 0&, com_TextBoxCenter, UserControl.ScaleWidth, UserControl.ScaleHeight, 0&, 0&
            Else
                RoundRect UserControl.hdc, 0&, com_TextBoxCenter, UserControl.ScaleWidth, UserControl.ScaleHeight, 10&, 10&
            End If

            UserControl.ForeColor = com_ColorBorderPic
            SetRect R, 0, 0, UserControl.ScaleWidth - 1, m_Height
            DrawGradBorderRect UserControl.hdc, com_ColorTo, com_ColorFrom, R, com_ColorBorderPic

            SetRect R, 0, m_Height, UserControl.ScaleWidth - 1, m_TextBoxHeight
            DrawGradBorderRect UserControl.hdc, com_ColorTo, com_ColorFrom, R, com_ColorBorderPic

            SetRect R, 0, m_Height + m_TextBoxHeight, UserControl.ScaleWidth - 1, m_Height
            DrawGradBorderRect UserControl.hdc, com_ColorTo, com_ColorFrom, R, com_ColorBorderPic

            SetRect R, 1, 1 + m_Height * 2 + m_TextBoxHeight, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - (1 + m_Height * 2 + m_TextBoxHeight) - UserControl.ScaleHeight * 0.2
            'DrawVGradientEx UserControl.hdc, BlendColors(com_ColorTo, vbWhite), BlendColors(com_ColorFrom, vbWhite), R.Left, R.Top, R.Right, R.Bottom
            DrawVGradientEx UserControl.hdc, m_FillColor, m_FillColor, R.Left, R.Top, R.Right, R.Bottom

            'set caption rect
            SetRect R_Caption, m_Space, m_Height + 1, UserControl.ScaleWidth - m_Space, m_TextBoxHeight + 2

            'set icon coordinates
            iY = (m_Height * 2 + m_TextBoxHeight - m_IconSize) / 2

    End Select
    
    'caption and icon alignments
    If Not (m_Icon Is Nothing Or m_Style = XPDefault) Then
        If m_IconAlignment = vbLeftAligment Then
            If m_Alignment = vbLeftJustify Then
                R_Caption.Left = R_Caption.Left + m_Space + m_IconSize
            ElseIf m_Alignment = vbRightJustify Then
                R_Caption.Left = R_Caption.Left + m_Space + m_IconSize
            Else
                R_Caption.Left = R_Caption.Left + m_Space + m_IconSize
                R_Caption.Right = R_Caption.Right - (m_Space + m_IconSize)
            End If
            
            iX = m_Space
            
        ElseIf m_IconAlignment = vbRightAligment Then
            If m_Alignment = vbLeftJustify Then
                R_Caption.Right = R_Caption.Right - (m_Space + m_IconSize)
            ElseIf m_Alignment = vbRightJustify Then
                R_Caption.Right = R_Caption.Right - (m_Space + m_IconSize)
            Else
                R_Caption.Left = R_Caption.Left + m_Space + m_IconSize
                R_Caption.Right = R_Caption.Right - (m_Space + m_IconSize)
            End If
            
            iX = UserControl.ScaleWidth - m_Space - m_IconSize
            
        End If
    End If

    'Draw caption
    '------------
    If Len(m_Caption) <> 0 Then
        'Set text color
        UserControl.ForeColor = m_TextColor
        
        'Draw text
        DrawTextEx UserControl.hdc, m_Caption, Len(m_Caption), R_Caption, com_TextDrawParams, ByVal 0&
    End If
    
    'draw picture
    '------------
    If Not (m_Icon Is Nothing Or m_Style = XPDefault) Then
        If m_Style = GradientFrame Then
            If iY < m_Height + 2 Then iY = m_Height + 2
        Else
            If iY < 0 Then iY = m_Space / 2
        End If
        UserControl.PaintPicture m_Icon, iX, iY, m_IconSize, m_IconSize
    End If
End Sub

Private Sub SetDefault()
    Select Case m_Style
        Case Is = XPDefault
            m_TextColor = &HCF3603
            m_FrameColor = RGB(195, 195, 195)
            m_TextBoxHeight = 22
            m_Alignment = vbLeftJustify
            m_FillColor = Ambient.BackColor
            SetTextDrawParams
        Case Is = GradientFrame
            m_TextColor = vbBlack
            m_FrameColor = vbBlack
            m_TextBoxHeight = 22
            m_Alignment = vbLeftJustify
            m_ThemeColor = Blue
            SetTextDrawParams
    End Select
End Sub

'==========================================================================
' API Functions and subroutines
'==========================================================================

' full version of APILine
Private Sub APILineEx(lhdcEx As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lColor As Long)

    'Use the API LineTo for Fast Drawing
    Dim pt As POINT
    Dim hPen As Long, hPenOld As Long
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(lhdcEx, hPen)
    MoveToEx lhdcEx, X1, Y1, pt
    LineTo lhdcEx, X2, Y2
    SelectObject lhdcEx, hPenOld
    DeleteObject hPen
End Sub

Private Function APIRectangle(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal w As Long, ByVal H As Long, Optional lColor As OLE_COLOR = -1) As Long
    
    Dim hPen As Long, hPenOld As Long
    Dim R
    Dim pt As POINT
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(hdc, hPen)
    MoveToEx hdc, X, Y, pt
    LineTo hdc, X + w, Y
    LineTo hdc, X + w, Y + H
    LineTo hdc, X, Y + H
    LineTo hdc, X, Y
    SelectObject hdc, hPenOld
    DeleteObject hPen
End Function

Private Sub DrawVGradientEx(lhdcEx As Long, lEndColor As Long, lStartColor As Long, ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long)
    
    ''Draw a Vertical Gradient in the current HDC
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    sR = (lStartColor And &HFF)
    sG = (lStartColor \ &H100) And &HFF
    sB = (lStartColor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    dR = (sR - eR) / Y2
    dG = (sG - eG) / Y2
    dB = (sB - eB) / Y2
    For ni = 0 To Y2
        APILineEx lhdcEx, X, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
    Next ni
End Sub

Private Sub DrawGradBorderRect(lhdcEx As Long, lEndColor As Long, lStartColor As Long, R As RECT, Optional lColor As OLE_COLOR = -1)
    'draw gradient rectangle with border
    DrawVGradientEx lhdcEx, lEndColor, lStartColor, R.Left, R.Top, R.Right, R.Bottom
    APIRectangle lhdcEx, R.Left, R.Top, R.Right, R.Bottom, lColor
End Sub

'Blend two colors
Private Function BlendColors(ByVal lcolor1 As Long, ByVal lcolor2 As Long)
    BlendColors = RGB(((lcolor1 And &HFF) + (lcolor2 And &HFF)) / 2, (((lcolor1 \ &H100) And &HFF) + ((lcolor2 \ &H100) And &HFF)) / 2, (((lcolor1 \ &H10000) And &HFF) + ((lcolor2 \ &H10000) And &HFF)) / 2)
End Function

'System color code to long rgb
Private Function TranslateColor(ByVal lColor As Long) As Long

    If OleTranslateColor(lColor, 0, TranslateColor) Then
          TranslateColor = -1
    End If
End Function

