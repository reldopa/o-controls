VERSION 5.00
Begin VB.UserControl oOptionButton 
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2025
   DrawMode        =   6  'Mask Pen Not
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   285
   ScaleWidth      =   2025
   ToolboxBitmap   =   "oOptionButton.ctx":0000
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Option"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   330
      TabIndex        =   0
      Top             =   30
      Width           =   480
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   180
      Picture         =   "oOptionButton.ctx":0312
      Top             =   510
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   0
      Picture         =   "oOptionButton.ctx":0862
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "oOptionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_Value = 0
Const m_def_Alignment = 0
'Property Variables:
Private m_Value As OLE_OPTEXCLUSIVE
Private m_Caption               As String
Dim m_Alignment As e_Alignment
'Event Declarations:
Event Click() 'MappingInfo=img(0),img,0,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
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


Public Enum e_Alignment
    A_Left = 0
    A_Right = 1
End Enum

Private Sub lbl_Click()
    Call img_Click(IIf(img(0).Visible = True, 0, 1))
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 285
    img(0).Top = 0
    img(1).Top = 0
    img(0).Left = 0
    img(1).Left = 0
    lbl.Top = 30
    lbl.Left = img(1).Width + 45
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property


Public Property Let Caption(ByRef new_caption As String)
    lbl.Caption = new_caption
    PropertyChanged "Caption"
End Property

Public Property Get Caption() As String
    Caption = lbl.Caption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl,lbl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lbl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lbl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
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
'MappingInfo=lbl,lbl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lbl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lbl.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub img_Click(Index As Integer)
    If Index = 0 Then
        img(0).Visible = False
        img(1).Visible = True
        Value = True
    Else
        img(0).Visible = True
        img(1).Visible = False
        Value = False
    End If
    RaiseEvent Click
End Sub

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
'MemberInfo=14,0,0,0
Public Property Get Value() As OLE_OPTEXCLUSIVE
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As OLE_OPTEXCLUSIVE)
    If New_Value = True Then
        m_Value = New_Value
        img(0).Visible = False
        img(1).Visible = True
        PropertyChanged "Value"
    Else
        m_Value = New_Value
        img(0).Visible = True
        img(1).Visible = False
        PropertyChanged "Value"
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Alignment() As Variant
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal new_Alignment As Variant)
    m_Alignment = new_Alignment
    PropertyChanged "Alignment"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Caption = lbl.Caption
    m_Value = m_def_Value
    m_Alignment = m_def_Alignment
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    lbl.ForeColor = PropBag.ReadProperty("ForeColor", &H800000)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    lbl.Caption = PropBag.ReadProperty("Caption", "Option")
    Set lbl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Value = PropBag.ReadProperty("Value", m_def_Value)
    m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", lbl.ForeColor, &H800000)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Caption", lbl.Caption, "Option")
    Call PropBag.WriteProperty("Font", lbl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
End Sub

