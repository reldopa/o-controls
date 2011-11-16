VERSION 5.00
Begin VB.UserControl oCheckBox 
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   ScaleHeight     =   1590
   ScaleWidth      =   2940
   Begin VB.TextBox txtBound 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   780
      TabIndex        =   0
      Text            =   "0"
      Top             =   540
      Width           =   195
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   165
      Index           =   0
      Left            =   30
      Picture         =   "oCheckBox.ctx":0000
      Top             =   510
      Width           =   165
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   165
      Index           =   1
      Left            =   480
      Picture         =   "oCheckBox.ctx":0228
      Top             =   870
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   330
      TabIndex        =   1
      Top             =   30
      Width           =   435
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      BorderStyle     =   0  'Transparent
      Height          =   210
      Left            =   0
      Top             =   0
      Width           =   210
   End
End
Attribute VB_Name = "oCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_Value = 0
Const m_def_SendKeysTab = False
Const m_def_Mandatory = False
Const m_def_BorderColor = &HFF8080

'Property Variables:
Dim m_Value As Integer
Dim m_DataField As ADODB.Field
Dim m_Mandatory As Boolean
Dim m_SendKeysTab As Boolean
Dim m_Alignment As e_Alignment
Dim m_BorderColor As OLE_COLOR

'Event Declarations:
Event Click() 'MappingInfo=img(0),img,0,Click
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp

Public Enum e_Value
    Unchecked = 0
    Checked = 1
End Enum



Private Sub lbl_Click()
    Dim lChecked As Boolean
    lChecked = IIf(img(1).Visible = True, 1, 0)
    If lChecked = False Then
        img(0).Visible = False
        img(1).Visible = True
        m_Value = 1
    Else
        img(0).Visible = True
        img(1).Visible = False
        m_Value = 0
    End If
    RaiseEvent Click
        
End Sub

Private Sub txtBound_Change()
    If Abs(IIf(txtBound = "", 0, txtBound)) = 0 Or Abs(IIf(txtBound = "", 0, txtBound)) = 1 Then
        Call img_Click(Abs(IIf(txtBound = "", 0, txtBound)))
    Else
        Err.Raise Number:=vbObjectError + 1001, _
                     Description:="Invalid Value (0 or 1 Only)"
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
   If m_Value = 0 Then
            img(0).Visible = True
            img(1).Visible = False
        Else
            img(0).Visible = False
            img(1).Visible = True
    End If
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
If m_Alignment = A_Left Or m_Alignment = A_Right Then
    If m_Alignment = A_Left Then
        UserControl.Height = Shape.Height
        Shape.Top = 0
        Shape.Left = 0
        img(0).Top = 15
        img(1).Top = 15
        img(0).Left = 15
        img(1).Left = 15
        lbl.Top = (UserControl.Height - lbl.Height) / 2
        lbl.Left = Shape.Width + 45
    Else
        UserControl.Height = Shape.Height
        Shape.Top = 0
        Shape.Left = (UserControl.Width - Shape.Width)
        img(0).Top = 15
        img(1).Top = 15
        img(0).Left = (Shape.Left + 15)
        img(1).Left = (Shape.Left + 15)
        lbl.Top = (UserControl.Height - lbl.Height) / 2
        lbl.Left = 15
    End If
Else
    Err.Raise Number:=vbObjectError + 1001, _
                     Description:="Invalid Value (0 or 1 Only)"
End If
End Sub


Public Property Get Alignment() As e_Alignment
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal new_Alignment As e_Alignment)
    m_Alignment = new_Alignment
    Call UserControl_Resize
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = Shape.BackColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    Shape.BackColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txt,txt,-1,DataSource
Public Property Get DataSource() As ADODB.Recordset
    Set DataSource = txtBound.DataSource
End Property

Public Property Set DataSource(ByVal New_DataSource As ADODB.Recordset)
    Set txtBound.DataSource = New_DataSource
    PropertyChanged "DataSource"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=30,0,0,
Public Property Get DataField() As ADODB.Field
    Set DataField = m_DataField
End Property

Public Property Let DataField(ByRef New_DataField As ADODB.Field)
    Set m_DataField = New_DataField
    txtBound.DataField = m_DataField.Name
    PropertyChanged "DataField"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl,lbl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lbl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lbl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    lbl.Enabled = UserControl.Enabled()
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl,lbl,-1,Font
Public Property Get Font() As Font
    Set Font = lbl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lbl.Font = New_Font
    PropertyChanged "Font"
End Property

'Mandatory
Public Property Get Mandatory() As Boolean
    Mandatory = m_Mandatory
End Property

Public Property Let Mandatory(ByVal New_Mandatory As Boolean)
    m_Mandatory = New_Mandatory
    PropertyChanged "Mandatory"
    If m_Mandatory Then
        lbl.ForeColor = &HFF0000
    Else
        lbl.ForeColor = &H800000
    End If
    
End Property

'SendKeys "{Tab}"
Public Property Get SendKeysTab() As Boolean
    SendKeysTab = m_SendKeysTab
End Property

Public Property Let SendKeysTab(ByVal New_SendKeysTab As Boolean)
    m_SendKeysTab = New_SendKeysTab
    PropertyChanged "SendKeysTab"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
    UserControl.Refresh
End Sub

Private Sub img_Click(Index As Integer)
Dim lChecked As Boolean
 '   lChecked = Index
    If img(0).Visible = False Then
        img(0).Visible = True
        img(1).Visible = False
        m_Value = 0
    Else
        img(0).Visible = False
        img(1).Visible = True
        m_Value = 1
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
'MappingInfo=lbl,lbl,-1,Caption
Public Property Get Caption() As String
    Caption = lbl.Caption
End Property

Public Property Let Caption(ByVal new_caption As String)
    lbl.Caption() = new_caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Value() As e_Value
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As e_Value)
    If New_Value = Unchecked Or New_Value = Checked Then
        m_Value = New_Value
        If m_Value = 0 Then
            img(0).Visible = True
            img(1).Visible = False
        Else
            img(0).Visible = False
            img(1).Visible = True
        End If
        PropertyChanged "Value"
    Else
        Err.Raise Number:=vbObjectError + 1001, _
                     Description:="Invalid Value (0 or 1 Only)"
    End If
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
    m_SendKeysTab = m_def_SendKeysTab
    m_Alignment = A_Left
    m_Mandatory = m_def_Mandatory
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    SendKeysTab = PropBag.ReadProperty("SendKeysTab", m_def_SendKeysTab)
    Mandatory = PropBag.ReadProperty("Mandatory", m_def_Mandatory)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Shape.BackColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    lbl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lbl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lbl.Caption = PropBag.ReadProperty("Caption", "Check")
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_Alignment = PropBag.ReadProperty("Alignment", 0)
    Set DataSource = PropBag.ReadProperty("DataSource", Nothing)
    Set m_DataField = PropBag.ReadProperty("DataField", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Alignment", m_Alignment, 0)
    Call PropBag.WriteProperty("SendKeysTab", m_SendKeysTab, m_def_SendKeysTab)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderColor", Shape.BackColor, m_def_BorderColor)
    Call PropBag.WriteProperty("ForeColor", lbl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lbl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Mandatory", m_Mandatory, m_def_Mandatory)
    Call PropBag.WriteProperty("Caption", lbl.Caption, "Check")
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("DataSource", DataSource, Nothing)
    Call PropBag.WriteProperty("DataField", m_DataField, Nothing)
End Sub



