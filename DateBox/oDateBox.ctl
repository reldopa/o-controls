VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl oDateBox 
   Appearance      =   0  'Flat
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2595
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MousePointer    =   1  'Arrow
   PropertyPages   =   "oDateBox.ctx":0000
   ScaleHeight     =   1005
   ScaleWidth      =   2595
   ToolboxBitmap   =   "oDateBox.ctx":0023
   Begin VB.TextBox mTextbox 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM/dd/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   1140
      TabIndex        =   2
      Top             =   780
      Visible         =   0   'False
      Width           =   945
   End
   Begin MSMask.MaskEdBox mDateBox 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM/dd/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   255
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   180
      Left            =   30
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   510
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   318
      _Version        =   393216
      Format          =   42139649
      CurrentDate     =   37444
   End
   Begin VB.Image img 
      Height          =   255
      Left            =   2310
      Picture         =   "oDateBox.ctx":0335
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFC0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   0
      Top             =   0
      Width           =   3075
   End
End
Attribute VB_Name = "oDateBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit


Private Const MCM_FIRST  As Long = &H1000&
Private Const MCM_GETMONTHDELTA   As Long = (MCM_FIRST + 19)
Private Const MCM_SETMONTHDELTA  As Long = (MCM_FIRST + 20)

Private Const DTM_FIRST  As Long = &H1000&
Private Const DTM_GETMONTHCAL   As Long = (DTM_FIRST + 8)

Private Const BM_SETSTATE = &HF3
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202


'Default Property Values:
Const m_def_BackStyle = 0
Const m_def_Text = "__/__/____"
Const m_def_SendKeysTab = False
Const m_def_Mandatory = False
Const m_def_DatePicker = True
Const m_def_BackColor = &HFFFFFF
Const m_def_BorderColor = &HFFC0C0
Const m_def_EmptyDate = True
Const m_def_Modal = False

'Property Variables:
Dim m_DataField As ADODB.Field
Dim m_BackStyle As Integer
Dim m_Text As String
Dim m_SendKeysTab As Boolean
Dim m_Mandatory As Boolean
Dim m_DatePicker As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_BorderColor As OLE_COLOR
Dim m_EmptyDate As Boolean
Dim m_Modal As Boolean
Dim m_CalendarPosition As e_CalendarPosition

Dim lDateboxFocus As Boolean

Public Enum e_CalendarPosition
    Bottom = 0
    Top = 1
End Enum

'Event Declarations:
Event Change()
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=mDateBox,mDateBox,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=mDateBox,mDateBox,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=mDateBox,mDateBox,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim lMV As Boolean
Dim nOldWidth As Long

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mDateBox,mDateBox,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = mDateBox.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    mDateBox.BackColor = m_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mDateBox,mDateBox,-1,BackColor
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = Shape1.BackColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    Shape1.BackColor = m_BorderColor
    PropertyChanged "BorderColor"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mDateBox,mDateBox,-1,Enabled
Public Property Get CalendarPosition() As e_CalendarPosition
    CalendarPosition = m_CalendarPosition
End Property

Public Property Let CalendarPosition(ByVal New_CalendarPosition As e_CalendarPosition)
    m_CalendarPosition = New_CalendarPosition
    PropertyChanged "CalendarPosition"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mDateBox,mDateBox,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mDateBox.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    mDateBox.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mDateBox,mDateBox,-1,Enabled
Public Property Get Modal() As Boolean
    Modal = m_Modal
End Property

Public Property Let Modal(ByVal New_Modal As Boolean)
    m_Modal = New_Modal
    PropertyChanged "Modal"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mDateBox,mDateBox,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = mDateBox.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    mDateBox.Enabled() = New_Enabled
    dtpDate.Enabled = New_Enabled
    img.Enabled = dtpDate.Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mDateBox,mDateBox,-1,Font
Public Property Get Font() As Font
    Set Font = mDateBox.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set mDateBox.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mDateBox,mDateBox,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConstants
    BorderStyle = mDateBox.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    mDateBox.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mDateBox,mDateBox,-1,Refresh
Public Sub Refresh()
    mDateBox.Refresh
End Sub

Private Sub dtpDate_Change()
    'mDateBox = Format(dtpDate.Value, "mm/dd/yyyy")
End Sub

Private Sub dtpDate_GotFocus()
    SendKeys "{Tab}"
End Sub

Private Sub dtpDate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mDateBox = Format(dtpDate, "mm/dd/yyyy")
End Sub

Private Sub img_Click()
If img.Enabled Then Call DisplayForm
End Sub

Private Sub mDateBox_Change()
On Error Resume Next
    TextEqualsDate
    m_Text = mTextbox.Text
    PropertyChanged ("Text")
    RaiseEvent Change
End Sub

Private Sub mDateBox_GotFocus()
    Set ctlTextVal = Nothing
    Set ctlTextVal = mDateBox
    lDateboxFocus = True
    mDateBox.SelStart = 0
    mDateBox.SelLength = Len(mDateBox)
End Sub

Private Sub mDateBox_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub mDateBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If m_SendKeysTab = True Then SendKeys "{Tab}"
    ElseIf KeyAscii = 8 Then
        Exit Sub
'    ElseIf KeyAscii <> 13 Then
'        If IsDate(mDateBox.Text) Then
'            mDateBox.Text = "__/__/____"
'            KeyAscii = KeyAscii
'        End If
    Else
        KeyAscii = ValidKeys(KeyAscii, "1234567890", True)
    End If
   RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub mDateBox_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub


Private Sub mDateBox_LostFocus()
    lDateboxFocus = False
End Sub

Private Sub mTextBox_Change()
    DateEqualsText
End Sub

Private Sub mTextbox_GotFocus()
    SendKeys "{Tab}"
End Sub

Private Sub UserControl_GotFocus()
    Set ctlTextVal = Nothing
    Set ctlTextVal = mDateBox
End Sub

Private Sub UserControl_Initialize()
    dtpDate = Date
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackStyle = m_def_BackStyle
    m_Text = m_def_Text
    m_SendKeysTab = m_def_SendKeysTab
    m_Mandatory = m_def_Mandatory
    m_DatePicker = m_def_DatePicker
    m_EmptyDate = m_def_EmptyDate
    m_BackColor = m_def_BackColor
    m_BorderColor = m_def_BorderColor
    m_Modal = m_def_Modal
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Modal = PropBag.ReadProperty("Modal", m_def_Modal)
    Text = PropBag.ReadProperty("Text", m_def_Text)
    CalendarPosition = PropBag.ReadProperty("CalendarPosition", 0)
    SendKeysTab = PropBag.ReadProperty("SendKeysTab", m_def_SendKeysTab)
    EmptyDate = PropBag.ReadProperty("EmptyDate", m_def_EmptyDate)
    mDateBox.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    Shape1.BackColor = PropBag.ReadProperty("BorderColor", &HFFC0C0)
    mDateBox.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    mDateBox.Enabled = PropBag.ReadProperty("Enabled", True)
    Set mDateBox.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    mDateBox.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Set DataSource = PropBag.ReadProperty("DataSource", Nothing)
    Set m_DataField = PropBag.ReadProperty("DataField", Nothing)
    dtpDate.Enabled = PropBag.ReadProperty("Enabled", Enabled)
    Mandatory = PropBag.ReadProperty("Mandatory", m_def_Mandatory)
    DatePicker = PropBag.ReadProperty("DatePicker", m_def_DatePicker)
    BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    If lMV Then Exit Sub
    
    If m_DatePicker Then
        If UserControl.Width < 1300 Then UserControl.Width = 1300
    Else
        If UserControl.Width < 975 Then UserControl.Width = 975
    End If
    
    UserControl.Height = 285
    
    Shape1.Height = UserControl.Height + 15
    Shape1.Width = UserControl.Width
    Shape1.Top = 0
    Shape1.Left = 0
    img.Width = 285
    img.Height = 285
    img.Left = UserControl.Width - (img.Width)
    img.Top = 15
    
    
    
    With mDateBox
        .Top = 15
        .Left = 15
        .Height = UserControl.Height - 30
        .Width = (UserControl.Width - img.Width) - 30
    End With
    
    With dtpDate
        .Top = 15
        .Left = 15
        .Height = mDateBox.Height - 30
        .Width = mDateBox.Width - 30
        .Visible = False
    End With
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Modal", m_Modal, m_def_Modal)
    Call PropBag.WriteProperty("BackColor", mDateBox.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("BorderColor", Shape1.BackColor, &HFFC0C0)
    Call PropBag.WriteProperty("ForeColor", mDateBox.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", mDateBox.Enabled, True)
    Call PropBag.WriteProperty("CalendarPosition", m_CalendarPosition, 0)
    Call PropBag.WriteProperty("Font", mDateBox.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", mDateBox.BorderStyle, 0)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("SendKeysTab", m_SendKeysTab, m_def_SendKeysTab)
    Call PropBag.WriteProperty("EmptyDate", m_EmptyDate, m_def_EmptyDate)
    Call PropBag.WriteProperty("DataSource", DataSource, Nothing)
    Call PropBag.WriteProperty("DataField", m_DataField, Nothing)
    Call PropBag.WriteProperty("Mandatory", m_Mandatory, m_def_Mandatory)
    Call PropBag.WriteProperty("DatePicker", m_DatePicker, m_def_DatePicker)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
End Sub

'TEXT
Public Property Get Text() As String
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Text"
    m_Text = mDateBox.Text
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
    If IsDate(m_Text) = False Then
        If Trim(m_Text) <> "__/__/____" Then
            New_Text = "__/__/____"
            m_Text = New_Text
        End If
    End If
    mDateBox.Text = m_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mTextBox,mTextBox,-1,DataSource
Public Property Get DataSource() As ADODB.Recordset
Attribute DataSource.VB_Description = "Sets a value that specifies the Data control through which the current control is bound to a database. "
    Set DataSource = mTextbox.DataSource
End Property

Public Property Set DataSource(ByRef New_DataSource As ADODB.Recordset)
    Set mTextbox.DataSource = New_DataSource
    PropertyChanged "DataSource"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=30,0,0,
Public Property Get DataField() As ADODB.Field
    Set DataField = m_DataField
End Property

Public Property Let DataField(ByRef New_DataField As ADODB.Field)
    Set m_DataField = New_DataField
    mTextbox.DataField = m_DataField.Name
    PropertyChanged "DataField"
End Property

Sub TextEqualsDate()
    m_EmptyDate = False
    If Trim(mDateBox) = "__/__/____" Then
        dtpDate = Date
        mTextbox = Empty
        'If mTextbox.DataField <> "" Then m_DataField.Value = Empty
        m_EmptyDate = True
    ElseIf Not IsDate(mDateBox) Then
        dtpDate = Date
        mTextbox = Date
        'If mTextbox.DataField <> "" Then m_DataField.Value = Empty
        m_EmptyDate = True
    ElseIf CDate(mDateBox) < CDate("01/01/1900") Then
        dtpDate = Date
        mTextbox = Empty
        'If mTextbox.DataField <> "" Then m_DataField.Value = Empty
        m_EmptyDate = True
    Else
        mTextbox = mDateBox
        dtpDate = mDateBox
        m_DataField.Value = mDateBox
    End If
End Sub

Sub DateEqualsText()
On Error Resume Next
Dim dVal As String
    If lDateboxFocus Then Exit Sub
    If Trim(mTextbox) = vbNullString Then
        mDateBox = "__/__/____"
    ElseIf Not IsDate(mTextbox) Then
        Beep
    ElseIf IsDate(mTextbox) Then
        mDateBox = Format(mTextbox, "mm/dd/yyyy")
    ElseIf IsEmpty(mTextbox) Then
        
    End If
End Sub

'Don't Remove this function
Public Function DisplayForm()
Dim frmMV As Form
Dim dbgWindows As RECT
On Error GoTo ErrHandler:
    Set frmMV = frmMonthView
    GetWindowRect UserControl.hwnd, dbgWindows
    With frmMV
        .currDate = IIf(mDateBox = "__/__/____", Format(Now, "mm/dd/yyyy"), mDateBox)
        .Left = dbgWindows.Left * Screen.TwipsPerPixelX
        If m_CalendarPosition = Bottom Then
            .Top = (dbgWindows.Top * Screen.TwipsPerPixelY) + UserControl.Height
        Else
            .Top = (dbgWindows.Top * Screen.TwipsPerPixelY) - .Height
        End If
        If m_Modal Then .Show vbModal, Me Else .Show
    End With
    Set frmMV = Nothing
ErrHandler:
    If Err.Number <> 0 Then
        If Err.Number = 401 Then
            MsgBox "Please set the control into modal state."
        ElseIf Err.Number = 91 Then
            Exit Function
        Else
            MsgBox Err.Number & ": " & Err.Description
        End If
    End If
End Function

'SendKeys "{Tab}"
Public Property Get SendKeysTab() As Boolean
    SendKeysTab = m_SendKeysTab
End Property

Public Property Let SendKeysTab(ByVal New_SendKeysTab As Boolean)
    m_SendKeysTab = New_SendKeysTab
    PropertyChanged "SendKeysTab"
End Property

'Mandatory
Public Property Get Mandatory() As Boolean
    Mandatory = m_Mandatory
End Property

Public Property Let Mandatory(ByVal New_Mandatory As Boolean)
    m_Mandatory = New_Mandatory
    PropertyChanged "Mandatory"
    If m_Mandatory Then
        mDateBox.BackColor = &HE0FFFF
        BackColor = &HE0FFFF
    Else
        mDateBox.BackColor = &HFFFFFF
        BackColor = &HFFFFFF
    End If
    
End Property

'DatePicker
Public Property Get DatePicker() As Boolean
    DatePicker = m_DatePicker
End Property

Public Property Let DatePicker(ByVal New_DatePicker As Boolean)
On Error GoTo ErrHandler
    m_DatePicker = New_DatePicker
    PropertyChanged "DatePicker"
    If m_DatePicker Then
        If UserControl.Width < 1260 Then UserControl.Width = 1260
    Else
        If UserControl.Width < 975 Then UserControl.Width = 975
    End If
    If m_DatePicker Then
        mDateBox.Width = (UserControl.Width - img.Width) - 30
    Else
        mDateBox.Width = UserControl.Width
    End If
ErrHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description
    End If
End Property

'Empty Date
Public Property Get EmptyDate() As Boolean
    EmptyDate = m_EmptyDate
End Property

Public Property Let EmptyDate(ByRef New_EmptyDate As Boolean)
    m_EmptyDate = New_EmptyDate
    PropertyChanged "EmptyDate"
End Property

Function ValidKeys(KeyIn As Integer, ValidateString As String, Editable As Boolean) As Integer

'ex:  KeyAscii = ValidKeys(KeyAscii, "01234567890.", True)
Dim ValidateList As String
Dim KeyOut As Integer

    If Editable = True Then
        ValidateList = UCase(ValidateString) & Chr(8)
    Else
        ValidateList = UCase(ValidateString)
    End If
    
    If InStr(1, ValidateList, UCase(Chr(KeyIn)), 1) > 0 Then
        KeyOut = KeyIn
    Else
        KeyOut = 0
        Beep
    End If
    
    ValidKeys = KeyOut
End Function
