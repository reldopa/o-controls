VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.UserControl oTextBox 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackStyle       =   0  'Transparent
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   DataBindingBehavior=   1  'vbSimpleBound
   DataSourceBehavior=   1  'vbDataSource
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   PropertyPages   =   "oTextBox.ctx":0000
   ScaleHeight     =   675
   ScaleWidth      =   4155
   ToolboxBitmap   =   "oTextBox.ctx":0011
   Begin VB.TextBox mTextbox 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   30
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid mDataGrid 
      Height          =   1965
      Left            =   390
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2730
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   3466
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton mButton 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3870
      Picture         =   "oTextBox.ctx":0323
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   45
      Top             =   285
   End
   Begin VB.TextBox mBoundText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   870
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   390
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   345
      Left            =   1530
      Top             =   0
      Width           =   2355
   End
   Begin VB.Label mLabel 
      Caption         =   "ITGtext"
      Height          =   255
      Left            =   30
      TabIndex        =   1
      Top             =   45
      Width           =   1500
   End
End
Attribute VB_Name = "oTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Dim ucX As Single, ucY As Single
Dim nKeyTab As Integer

'Default Property Values:
Const m_def_Hover = 0
Const m_def_Required = 0
Const m_def_LinkForeColor = &HFF&
Const m_def_Locked = 0
Const m_def_Passwordchar = ""
Const m_def_BackColor = &HFFFFFF
Const m_def_BorderColor = &HFFC0C0
Const m_def_LabelBackColor = &H8000000F
Const m_def_ForeColor = 0
Const m_def_BorderStyle = 0

Const m_def_Text = ""
Const m_def_AllCaps = False
Const m_def_TextButton = False
Const m_def_DataType = 0
Const m_def_MaxLength = 0
Const m_def_MaximumValue = 0
Const m_def_Mandatory = False
Const m_def_DecimalPlace = 0
Const m_def_Font = "Tahoma"
Const m_def_LabelWidth = 1500
Const m_def_TextBoxWidth = 2000
Const m_def_LabelAlignment = 0
Const m_def_LabelBackStyle = 1
Const m_def_SendKeysTab = False
Const m_def_TextTrim = False
Const m_def_ConnectionStrings = ""
Const m_def_SQLScript = ""
Const m_def_NoOfColumns = 2
Const m_def_GridHeight = 1905
Const m_def_GridWidth = 4260
'Const m_def_ColumnCaption = ""
Const m_def_Headlines = 1
Const m_def_AllowNegative = True

Public Enum e_BorderStyle
    None = 0
    FixedSingle = 1
End Enum

Public Enum e_DataType
    AlphaNumeric = 0
    Numeric = 1
End Enum

Public Enum e_LabelAlignment
    LeftJustified = 0
    RightJustified = 1
    Center = 2
End Enum

Public Enum e_LabelBackStyle
    Transparent = 0
    Opaque = 1
End Enum

'Property Variables:
Dim m_Hover As Boolean
Dim m_Required As Boolean
Dim m_LinkForeColor As OLE_COLOR
Dim m_SendKeysTab As Boolean
Dim m_DataField As ADODB.Field
Dim m_Locked As Boolean
Dim m_Passwordchar As String
Dim m_BackColor As OLE_COLOR
Dim m_BorderColor As OLE_COLOR
Dim m_LabelBackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_Font As Font
Dim m_BorderStyle As e_BorderStyle
Dim m_Text As String
Dim m_AllCaps As Boolean
Dim m_DataType As e_DataType
Dim m_MaxLength As Long
Dim m_MaximumValue As Double
Dim m_Mandatory As Boolean
Dim m_DecimalPlace As Long
Dim passthrough As Boolean
Dim m_ToWords As Boolean
Dim m_LabelWidth As Long
Dim m_TextBoxWidth As Long
Dim m_LabelAlignment As e_LabelAlignment
Dim m_LabelBackStyle As e_LabelBackStyle
Dim m_TextTrim As Boolean
Dim m_TextButton As Boolean
Dim m_ConnectionStrings As String
Dim m_SQLScript As String
Dim m_NoOfColumns As Integer
Dim m_GridHeight As Integer
Dim m_GridWidth As Integer
Dim m_AllowNegative As Boolean

'Dim m_ColumnCaption As String
Dim m_Headlines As Integer

Dim strFormat As String
Dim lGo As Boolean
Dim lLostFocus As Boolean

'Event Declarations:
Event LabelClick() 'MappingInfo=mLabel,mLabel,-1,Click
Attribute LabelClick.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event Change()
Attribute Change.VB_Description = "Occurs when the user sets a value on the mTextBox"
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event SelectionClick()
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Private com_Highlighted         As Boolean
Private com_FontBold            As Boolean
Private SelectedForeColor       As OLE_COLOR
Private lHandle As Long

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    mTextbox.BackColor = m_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    Shape1.BackColor = m_BorderColor
    PropertyChanged "BorderColor"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property


Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    mTextbox.ForeColor = m_ForeColor
    PropertyChanged "ForeColor"
End Property


Public Property Get ButtonPix() As Picture
    Set ButtonPix = mButton.Picture
End Property

Public Property Set ButtonPix(ByVal New_ButtonPix As Picture)
    Set mButton.Picture = New_ButtonPix
    PropertyChanged "ButtonPix"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    Set mTextbox.Font = m_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As e_BorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As e_BorderStyle)
    m_BorderStyle = New_BorderStyle
    mTextbox.BorderStyle = m_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,
Public Property Get Text() As String
Attribute Text.VB_Description = "Set the content of the mBox."
Attribute Text.VB_ProcData.VB_Invoke_Property = "ITGTextBoxPP"
Attribute Text.VB_UserMemId = 0
Attribute Text.VB_MemberFlags = "200"
    m_Text = mTextbox.Text
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
    mTextbox.Text = m_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,1
Public Property Get AllCaps() As Boolean
Attribute AllCaps.VB_Description = "Translates the character automatically to uppercase letters."
Attribute AllCaps.VB_ProcData.VB_Invoke_Property = "ITGTextBoxPP"
    AllCaps = m_AllCaps
End Property

Public Property Let AllCaps(ByVal New_AllCaps As Boolean)
    m_AllCaps = New_AllCaps
    PropertyChanged "AllCaps"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,1
Public Property Get AllowNegative() As Boolean
    AllowNegative = m_AllowNegative
End Property

Public Property Let AllowNegative(ByVal New_AllowNegative As Boolean)
    m_AllowNegative = New_AllowNegative
    PropertyChanged "AllowNegative"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,1
Public Property Get TextButton() As Boolean
    TextButton = m_TextButton
End Property

Public Property Let TextButton(ByVal New_TextButton As Boolean)
    m_TextButton = New_TextButton
    PropertyChanged "TextButton"
    If m_TextButton Then
        mButton.Visible = True
    Else
        mButton.Visible = False
    End If
    Call UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get DataType() As e_DataType
Attribute DataType.VB_Description = "Sets the datatype of textbox to numberic or alphabet"
    DataType = m_DataType
End Property

Public Property Let DataType(ByVal New_DataType As e_DataType)
    m_DataType = New_DataType
    PropertyChanged "DataType"
    If m_DataType = Numeric Then
        Text = "0"
        passthrough = False
        mTextbox.Alignment = 1
        passthrough = True
        Call LostFocusTextBox(mTextbox)
        Text = mTextbox.Text
        AllCaps = False
    Else
        mTextbox.Alignment = 0
        AllCaps = True
        DecimalPlace = 0
    End If
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Set the maxlength of the mBox if alphanumeric."
Attribute MaxLength.VB_ProcData.VB_Invoke_Property = "ITGTextBoxPP"
    MaxLength = m_MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    m_MaxLength = New_MaxLength
    PropertyChanged "MaxLength"
    mTextbox.MaxLength = m_MaxLength
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=4,0,0,99.999999
Public Property Get MaximumValue() As Double
Attribute MaximumValue.VB_Description = "Maximum value if datatype is numeric."
Attribute MaximumValue.VB_ProcData.VB_Invoke_Property = "ITGTextBoxPP"
    MaximumValue = m_MaximumValue
End Property

Public Property Let MaximumValue(ByVal New_MaximumValue As Double)
    m_MaximumValue = New_MaximumValue
    PropertyChanged "MaximumValue"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Mandatory() As Boolean
Attribute Mandatory.VB_Description = "Sets the mBox color to yellow."
Attribute Mandatory.VB_ProcData.VB_Invoke_Property = "ITGTextBoxPP"
    Mandatory = m_Mandatory
End Property

Public Property Let Mandatory(ByVal New_Mandatory As Boolean)
    m_Mandatory = New_Mandatory
    PropertyChanged "Mandatory"
    If m_Mandatory Then
        mTextbox.BackColor = &HE0FFFF  '&H80000018  '&HC0FFFF
        BackColor = &HE0FFFF  '&H80000018  '&HC0FFFF
    Else
        mTextbox.BackColor = &HFFFFFF
        BackColor = &HFFFFFF
    End If
    
End Property

Private Sub mBoundText_Change()
    BoundToText
End Sub

Private Sub mBoundText_GotFocus()
    If m_SendKeysTab = True Then
            SendKeys "{Tab}"
    End If
End Sub


Private Sub mButton_Click()
Dim frm As Form
Dim dbgWindows As RECT
Set frm = frmGrid
GetWindowRect mTextbox.hwnd, dbgWindows
With frm 'form object
    sConnection = m_ConnectionStrings
    sRecordset = m_SQLScript
    .Left = dbgWindows.Left * Screen.TwipsPerPixelX
    .Top = dbgWindows.Top * Screen.TwipsPerPixelY
   .Show vbModal, Me
   mTextbox.Text = sGridText
End With
End Sub

Private Sub mButton_GotFocus()
    If m_TextButton = False Then Call SendKeys("{Tab}")
End Sub

Private Sub mDataGrid_GotFocus()

    If m_SendKeysTab = True Then
            SendKeys "{Tab}"
    End If
End Sub

'Private Sub mButton_GotFocus()
'    If mButton.Visible = False Then Call SendKeys("{Tab}")
'End Sub

Private Sub mLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim lHandle As Long
'    If m_Hover Then
'        lHandle = LoadCursor(0, HandCursor)
'        If (lHandle > 0) Then SetCursor lHandle
'
'        mLabel.ForeColor = m_LinkForeColor
'    End If

    If m_Hover Then
        lHandle = LoadCursor(0, HandCursor)
        If (lHandle > 0) Then SetCursor lHandle
        
        If com_Highlighted Then Exit Sub
        
        com_Highlighted = True
        com_FontBold = mLabel.FontBold
        SelectedForeColor = mLabel.ForeColor
        mLabel.ForeColor = m_LinkForeColor
        mLabel.FontBold = True
        Timer1.Enabled = True
    End If
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub mTextBox_Change()
Dim sfd As New StdDataFormat
On Error Resume Next
    
    If m_DataType = Numeric Then
        If Not (passthrough) Then Exit Sub
        Call VerifyTextBox(mTextbox)
        sfd.Format = strFormat
        Set mTextbox.DataFormat = sfd
        If lGo Then BoundToText
        PropertyChanged ("Text")
    Else
        PropertyChanged ("Text")
        m_Text = mTextbox.Text
        m_DataField.Value = m_Text
    End If
    If Not lLostFocus Then RaiseEvent Change
End Sub

Private Sub mTextbox_Click()
    RaiseEvent Click
    mTextbox.SelLength = Len(mTextbox)
End Sub

Private Sub mTextbox_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub mTextbox_GotFocus()

    If m_DataType = Numeric Then
        Call GotFocusTextBox(mTextbox)
        mTextbox.SelLength = Len(mTextbox.Text)
    Else
        mTextbox.SelStart = 0
        mTextbox.SelLength = Len(mTextbox.Text)
    End If
    lGo = False

End Sub

Private Sub mTextbox_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub mTextbox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If m_SendKeysTab = True Then
            SendKeys "{Tab}"
        End If
    ElseIf KeyAscii = 9 Or KeyAscii = 0 Then
        SendKeys "{Tab}"

    End If
    
    If m_DataType = Numeric Then
        Call KeypressTextbox(mTextbox, KeyAscii)
    ElseIf m_AllCaps = False Then
        KeyAscii = InvalidKeys(KeyAscii, "'")
    ElseIf m_AllCaps = True Then
        KeyAscii = InvalidKeys(KeyAscii, "'")
        If m_AllCaps Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
    
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub mTextbox_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = 115 Then RaiseEvent DblClick
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub mTextbox_LostFocus()
Dim s As String, i As Integer
Dim nKey As KeyCodeConstants
On Error Resume Next
    
    lLostFocus = True
   
    'TextToBound
    If m_DataType = Numeric Then
        s = ""
        For i = 1 To m_DecimalPlace
            s = s & "0"
        Next i
        
        If m_DecimalPlace <> 0 Then
            s = "###,###,###,###,###,###,###,##0." & s
        Else
            s = "###,###,###,###,###,###,###,##0"
        End If
        mTextbox = Format(mTextbox, s)
        lGo = True
        mTextBox_Change
    
    Else
        If m_AllCaps Then
            mTextbox = UCase(mTextbox)
        End If
    End If
    
    If m_TextTrim Then
        mTextbox = Trim(mTextbox)
    End If
   
    PropertyChanged ("Text")
    m_Text = mTextbox.Text
    m_DataField.Value = m_Text
    
    lLostFocus = False
  
   
End Sub

Private Sub mTextbox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub mTextbox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub mTextbox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Timer1_Timer()
Dim pt As POINTAPI

    ' See where the cursor is.
    GetCursorPos pt
    
    ' Translate into window coordinates.
    If WindowFromPointXY(pt.X, pt.Y) <> UserControl.hwnd _
        Then
        com_Highlighted = False
        mLabel.ForeColor = SelectedForeColor
        mLabel.FontBold = com_FontBold
        Timer1.Enabled = False
    End If

End Sub

Private Sub UserControl_Initialize()
    lGo = True
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_BorderColor = m_def_BorderColor
    m_LabelBackColor = m_def_LabelBackColor
    m_ForeColor = m_def_ForeColor
    Set m_Font = Ambient.Font
    m_BorderStyle = m_def_BorderStyle
    m_Text = m_def_Text
    m_DataType = m_def_DataType
    m_MaxLength = m_def_MaxLength
    m_MaximumValue = m_def_MaximumValue
    m_Mandatory = m_def_Mandatory
    m_DecimalPlace = m_def_DecimalPlace
    m_DataType = m_def_DataType
    m_AllCaps = m_def_AllCaps
    m_TextButton = m_def_TextButton
    m_Passwordchar = m_def_Passwordchar
    m_Locked = m_def_Locked
    m_Font = m_def_Font
    m_LabelWidth = m_def_LabelWidth
    m_TextBoxWidth = m_def_TextBoxWidth
    m_LabelAlignment = m_def_LabelAlignment
    m_LabelBackStyle = m_def_LabelBackStyle
    m_SendKeysTab = m_def_SendKeysTab
    m_TextTrim = m_def_TextTrim
    m_Hover = m_def_Hover
    m_Required = m_def_Required
    m_LinkForeColor = m_def_LinkForeColor
    m_ConnectionStrings = m_def_ConnectionStrings
    m_SQLScript = m_def_SQLScript
    'm_NoOfColumns = m_def_NoOfColumns
    m_GridHeight = m_def_GridHeight
    m_GridWidth = m_def_GridWidth
    'm_ColumnCaption = m_def_ColumnCaption
    m_Headlines = m_def_Headlines
    m_AllowNegative = m_def_AllowNegative
End Sub

'Private Sub UserControl_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 9 Or KeyAscii = 0 Then SendKeys "{Tab}"
'End Sub

'Private Sub UserControl_LostFocus()
'    SendKeys "{Tab}"
'End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    SendKeysTab = PropBag.ReadProperty("SendKeysTab", m_def_SendKeysTab)
    BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    LabelBackColor = PropBag.ReadProperty("LabelBackColor", m_def_LabelBackColor)
    
    ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    Text = PropBag.ReadProperty("Text", m_def_Text)
    Set mButton.Picture = PropBag.ReadProperty("ButtonPix", mButton.Picture)
    DataType = PropBag.ReadProperty("DataType", m_def_DataType)
    AllCaps = PropBag.ReadProperty("AllCaps", m_def_AllCaps)
    AllowNegative = PropBag.ReadProperty("AllowNegative", m_def_AllowNegative)
    TextButton = PropBag.ReadProperty("TextButton", m_def_TextButton)
    MaxLength = PropBag.ReadProperty("MaxLength", m_def_MaxLength)
    MaximumValue = PropBag.ReadProperty("MaximumValue", m_def_MaximumValue)
    Mandatory = PropBag.ReadProperty("Mandatory", m_def_Mandatory)
    BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    DecimalPlace = PropBag.ReadProperty("DecimalPlace", m_def_DecimalPlace)
    Passwordchar = PropBag.ReadProperty("Passwordchar", m_def_Passwordchar)
    Locked = PropBag.ReadProperty("Locked", m_def_Locked)
    mLabel.Caption = PropBag.ReadProperty("Label", "ITG")
    Set mLabel.Font = PropBag.ReadProperty("LabelFont", Ambient.Font)
    mLabel.ForeColor = PropBag.ReadProperty("LabelForeColor", &H80000012)
    LabelWidth = PropBag.ReadProperty("LabelWidth", m_def_LabelWidth)
    TextBoxWidth = PropBag.ReadProperty("TextBoxWidth", m_def_TextBoxWidth)
    LabelAlignment = PropBag.ReadProperty("LabelAlignment", m_def_LabelAlignment)
    LabelBackStyle = PropBag.ReadProperty("LabelBackStyle", m_def_LabelBackStyle)
    mTextbox.Enabled = PropBag.ReadProperty("Enabled", True)
    Set DataSource = PropBag.ReadProperty("DataSource", Nothing)
    Set m_DataField = PropBag.ReadProperty("DataField", Nothing)
    TextTrim = PropBag.ReadProperty("TextTrim", m_def_TextTrim)
    m_Hover = PropBag.ReadProperty("Hover", m_def_Hover)
    m_Required = PropBag.ReadProperty("Required", m_def_Required)
    m_LinkForeColor = PropBag.ReadProperty("LinkForeColor", m_def_LinkForeColor)
    ConnectionStrings = PropBag.ReadProperty("ConnectionStrings", m_def_ConnectionStrings)
    SQLScript = PropBag.ReadProperty("SQLScript", m_def_SQLScript)
'    NoOfColumns = PropBag.ReadProperty("NoOfColumns", m_def_NoOfColumns)
    GridHeight = PropBag.ReadProperty("GridHeight", m_def_GridHeight)
    GridWidth = PropBag.ReadProperty("GridWidth", m_def_GridWidth)
    'ColumnCaption = PropBag.ReadProperty("ColumnCaption", m_def_ColumnCaption)
    Headlines = PropBag.ReadProperty("Headlines", m_def_Headlines)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    If ScaleWidth <= LabelWidth Then Exit Sub
    UserControl.Height = 285
    m_TextBoxWidth = ((ScaleWidth - LabelWidth) - IIf(TextButton = True, mButton.Width + 15, 0)) - 60
    mTextbox.Left = LabelWidth + 45
    mTextbox.Top = 15
    mTextbox.Width = ((ScaleWidth - LabelWidth) - IIf(TextButton = True, mButton.Width + 15, 0)) - 60
    mTextbox.Height = 285 - 30
    mButton.Top = 0
    mButton.Left = UserControl.Width - mButton.Width
    Shape1.Top = 0
    Shape1.Left = mTextbox.Left - 15
    Shape1.Height = UserControl.Height + 15
    Shape1.Width = mTextbox.Width + 45
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SendKeysTab", m_SendKeysTab, m_def_SendKeysTab)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("LabelBackColor", m_LabelBackColor, m_def_LabelBackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("ButtonPix", mButton.Picture, mButton.Picture)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("AllCaps", m_AllCaps, m_def_AllCaps)
    Call PropBag.WriteProperty("AllowNegative", m_AllowNegative, m_def_AllowNegative)
    Call PropBag.WriteProperty("TextButton", m_TextButton, m_def_TextButton)
    Call PropBag.WriteProperty("DataType", m_DataType, m_def_DataType)
    Call PropBag.WriteProperty("MaxLength", m_MaxLength, m_def_MaxLength)
    Call PropBag.WriteProperty("MaximumValue", m_MaximumValue, m_def_MaximumValue)
    Call PropBag.WriteProperty("Mandatory", m_Mandatory, m_def_Mandatory)
    Call PropBag.WriteProperty("DecimalPlace", m_DecimalPlace, m_def_DecimalPlace)
    Call PropBag.WriteProperty("Passwordchar", m_Passwordchar, m_def_Passwordchar)
    Call PropBag.WriteProperty("Locked", m_Locked, m_def_Locked)
    Call PropBag.WriteProperty("Label", mLabel.Caption, "itg")
    Call PropBag.WriteProperty("LabelFont", mLabel.Font, Ambient.Font)
    Call PropBag.WriteProperty("LabelForeColor", mLabel.ForeColor, &H80000012)
    Call PropBag.WriteProperty("LabelWidth", m_LabelWidth, m_def_LabelWidth)
    Call PropBag.WriteProperty("TextBoxWidth", m_TextBoxWidth, m_def_TextBoxWidth)
    Call PropBag.WriteProperty("LabelAlignment", m_LabelAlignment, m_def_LabelAlignment)
    Call PropBag.WriteProperty("LabelBackStyle", m_LabelBackStyle, m_def_LabelBackStyle)
    Call PropBag.WriteProperty("Enabled", mTextbox.Enabled, True)
    Call PropBag.WriteProperty("DataSource", DataSource, Nothing)
    Call PropBag.WriteProperty("DataField", m_DataField, Nothing)
    Call PropBag.WriteProperty("TextTrim", m_TextTrim, m_def_TextTrim)
    Call PropBag.WriteProperty("Hover", m_Hover, m_def_Hover)
    Call PropBag.WriteProperty("Required", m_Required, m_def_Required)
    Call PropBag.WriteProperty("LinkForeColor", m_LinkForeColor, m_def_LinkForeColor)
    Call PropBag.WriteProperty("ConnectionStrings", m_ConnectionStrings, m_def_ConnectionStrings)
    Call PropBag.WriteProperty("SQLScript", m_SQLScript, m_def_SQLScript)
    'Call PropBag.WriteProperty("NoOfColumns", m_NoOfColumns, m_def_NoOfColumns)
    Call PropBag.WriteProperty("GridHeight", m_GridHeight, m_def_GridHeight)
    Call PropBag.WriteProperty("GridWidth", m_GridWidth, m_def_GridWidth)
    'Call PropBag.WriteProperty("ColumnCaption", m_ColumnCaption, m_def_ColumnCaption)
    Call PropBag.WriteProperty("Headlines", m_Headlines, m_def_Headlines)
End Sub

Public Property Get DecimalPlace() As Long
Attribute DecimalPlace.VB_ProcData.VB_Invoke_Property = "ITGTextBoxPP"
    DecimalPlace = m_DecimalPlace
End Property

Public Property Let DecimalPlace(ByVal New_DecimalPlace As Long)
    m_DecimalPlace = New_DecimalPlace
    PropertyChanged "DecimalPlace"
    If m_DataType = Numeric Then
        Text = "0"
        Call LostFocusTextBox(mTextbox)
    End If
End Property

Private Sub VerifyTextBox(mTextbox As TextBox)
Dim s As String, i As Integer
    
    s = ""
    strFormat = ""
    
    For i = 1 To m_DecimalPlace
        s = s & "0"
    Next i
    
    If m_DecimalPlace <> 0 Then
        s = "###0." & s
    Else
        s = "###0"
    End If
    
    strFormat = s
    
    If Trim(mTextbox.Text) = "" Then
        mTextbox.Text = Format(0, s)
        PropertyChanged "Text"
    ElseIf Trim(mTextbox.Text) = "." Then
        mTextbox.Text = "0."
        PropertyChanged "Text"
        mTextbox.SelStart = 3
    ElseIf Not (IsNumeric(mTextbox.Text)) Then
        If Trim(mTextbox.Text) <> "-" Then
            MsgBox "Invalid input.", vbExclamation ', ApplicationTitle
            mTextbox.Text = Format(0, s)
            PropertyChanged "Text"
            mTextbox.SelStart = 0
            mTextbox.SelLength = Len(mTextbox.Text)
        End If
    ElseIf CDec(mTextbox.Text) > m_MaximumValue Then
        If m_MaximumValue = 0 Then Exit Sub
        MsgBox "Value is out of range", vbExclamation ', ApplicationTitle
        mTextbox.Text = Format(0, s)
        PropertyChanged "Text"
        mTextbox.SelStart = 0
        mTextbox.SelLength = Len(mTextbox.Text)
    ElseIf CDec(mTextbox.Text) >= 100000000000000# Then
        MsgBox "Value is out of range", vbExclamation ', ApplicationTitle
        mTextbox.Text = Format(0, s)
        PropertyChanged "Text"
        mTextbox.SelStart = 0
        mTextbox.SelLength = Len(mTextbox.Text)
    ElseIf m_DecimalPlace = 0 And InStr(1, mTextbox.Text, ".", 1) > 0 Then
        mTextbox.Text = Format(mTextbox.Text, "###,###,###,###,###,###,###,##0")
    ElseIf m_DecimalPlace > 0 And InStr(1, mTextbox.Text, ".", 1) > 0 And (Len(mTextbox.Text) - InStr(1, mTextbox.Text, ".", 1)) > m_DecimalPlace Then
        mTextbox.Text = Format(mTextbox.Text, "###,###,###,###,###,###,###," & Right(s, Len(s) - 1))
    ElseIf AllowNegative = False Then
        If IsNumeric(mTextbox) Then
            If CDec(mTextbox) < 0 Then
                mTextbox.Text = Abs(mTextbox.Text)
                mTextbox.SelStart = 0
                mTextbox.SelLength = Len(mTextbox.Text)
            End If
        End If
    End If
    
End Sub

Private Sub GotFocusTextBox(mTextbox As TextBox)
Dim s As String, i As Integer
    
    s = ""
    For i = 1 To m_DecimalPlace
        s = s & "0"
    Next i
    
    If m_DecimalPlace <> 0 Then
        s = "###0." & s
    Else
        s = "###0"
    End If
    
    mTextbox = Format(mTextbox, s)
    'mTextbox.SelStart = 0
    mTextbox.SelLength = Len(mTextbox)
    
End Sub


Private Sub LostFocusTextBox(mTextbox As TextBox)
Dim s As String, i As Integer
    
    s = ""
    For i = 1 To m_DecimalPlace
        s = s & "0"
    Next i
    
    If m_DecimalPlace <> 0 Then
        s = "###,###,###,###,###,###,###,##0." & s
    Else
        s = "###,###,###,###,###,###,###,##0"
    End If
    mTextbox = Format(mTextbox, s)
    
End Sub

Private Sub KeypressTextbox(mTextbox As TextBox, KeyAscii As Integer)

    If m_DecimalPlace <> 0 Then
        KeyAscii = ValidKeys(KeyAscii, "-1234567890.", True)
    Else
        KeyAscii = ValidKeys(KeyAscii, "-1234567890", True)
    End If
    
    If mTextbox.SelLength <> Len(mTextbox) Then
        
        If KeyAscii = Asc("0") And (mTextbox.SelStart = 0 Or (mTextbox.SelStart = 1 And Left(mTextbox, 1) = "0")) Then
            KeyAscii = 0
        End If
        
        If mTextbox.SelStart > InStr(1, mTextbox, ".", 1) And InStr(1, mTextbox, ".", 1) <> 0 And KeyAscii <> 8 Then
            If Len(mTextbox.Text) - InStr(1, mTextbox.Text, ".", 1) = m_DecimalPlace Then
                KeyAscii = 0
                Beep
            End If
        End If
        
        If Chr(KeyAscii) = "." Then
            If InStr(1, mTextbox, Chr(KeyAscii), 1) > 0 Then
                KeyAscii = 0
                Beep
            ElseIf mTextbox.SelStart < (Len(mTextbox) - m_DecimalPlace) Then
                KeyAscii = 0
                Beep
            End If
        End If
        
        If Chr(KeyAscii) = "-" Then
            If InStr(1, mTextbox, Chr(KeyAscii), 1) > 0 Then
                KeyAscii = 0
                Beep
            ElseIf mTextbox.SelStart > 1 Then
                KeyAscii = 0
                Beep
            End If
        End If
        
    End If
    
End Sub

Function InvalidKeys(KeyIn As Integer, strList As String) As Integer
    If InStr(1, UCase(strList), UCase(Chr(KeyIn)), 1) > 0 Then
        InvalidKeys = 0
    Else
        InvalidKeys = KeyIn
    End If
End Function

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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Headlines() As String
    Headlines = m_Headlines
End Property

Public Property Let Headlines(ByVal New_Headlines As String)
    m_Headlines = New_Headlines
    frmGrid.iHeadlines = m_Headlines
    PropertyChanged "Headlines"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get ConnectionStrings() As String
Attribute ConnectionStrings.VB_ProcData.VB_Invoke_Property = "ITGTextBoxPP1"
    ConnectionStrings = m_ConnectionStrings
End Property

Public Property Let ConnectionStrings(ByVal New_ConnectionStrings As String)
    m_ConnectionStrings = New_ConnectionStrings
    'frmGrid.sConnection = m_ConnectionStrings
    PropertyChanged "ConnectionStrings"
End Property

Public Property Get SQLScript() As String
    SQLScript = m_SQLScript
End Property

Public Property Let SQLScript(ByVal New_SQLScript As String)
    m_SQLScript = New_SQLScript
    'frmGrid.sRecordset = m_SQLScript
    PropertyChanged "SQLScript"
End Property

''Public Property Get NoOfColumns() As String
''    NoOfColumns = m_NoOfColumns
''End Property
''
''Public Property Let NoOfColumns(ByVal New_NoOfColumns As String)
''    m_NoOfColumns = New_NoOfColumns
''    frmGrid.iNoOfColumns = m_NoOfColumns
''    PropertyChanged "NoOfColumns"
''End Property

Public Property Get GridHeight() As String
    GridHeight = m_GridHeight
End Property

Public Property Let GridHeight(ByVal New_GridHeight As String)
    m_GridHeight = New_GridHeight
    frmGrid.iHeight = m_GridHeight
    PropertyChanged "GridHeight"
End Property

Public Property Get GridWidth() As String
    GridWidth = m_GridWidth
End Property

Public Property Let GridWidth(ByVal New_GridWidth As String)
    m_GridWidth = New_GridWidth
    frmGrid.iWidth = m_GridWidth
    PropertyChanged "GridWidth"
End Property

'Public Property Get Column() As Columns
'    Column. = m_ColumnCaption
'End Property
'
'Public Property Let ColumnCaption(ByVal New_ColumnCaption As Columns)
'    m_ColumnCaption = New_ColumnCaption(nColIndex).Caption
'    frmGrid.sColumns(nColIndex).Caption = m_ColumnCaption
'    PropertyChanged "ColumnCaption"
'End Property

Public Property Get Columns() As Columns
    Set Columns = mDataGrid.Columns
    Set frmGrid.sColumns = Columns
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Passwordchar() As String
Attribute Passwordchar.VB_ProcData.VB_Invoke_Property = "ITGTextBoxPP"
    Passwordchar = m_Passwordchar
End Property

Public Property Let Passwordchar(ByVal New_Passwordchar As String)
    m_Passwordchar = New_Passwordchar
    mTextbox.Passwordchar = m_Passwordchar
    PropertyChanged "Passwordchar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Locked object"
    Locked = m_Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    m_Locked = New_Locked
    mTextbox.Locked = m_Locked
    mButton.Enabled = Not m_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mLabel,mLabel,-1,Caption
Public Property Get Label() As String
Attribute Label.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Label = mLabel.Caption
End Property

Public Property Let Label(ByVal New_Label As String)
    mLabel.Caption() = New_Label
    PropertyChanged "Label"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mLabel,mLabel,-1,Font
Public Property Get LabelFont() As Font
Attribute LabelFont.VB_Description = "Returns a Font object."
    Set LabelFont = mLabel.Font
End Property

Public Property Set LabelFont(ByVal New_LabelFont As Font)
    Set mLabel.Font = New_LabelFont
    PropertyChanged "LabelFont"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mLabel,mLabel,-1,ForeColor
Public Property Get LabelForeColor() As OLE_COLOR
Attribute LabelForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    LabelForeColor = mLabel.ForeColor
End Property

Public Property Let LabelForeColor(ByVal New_LabelForeColor As OLE_COLOR)
    mLabel.ForeColor() = New_LabelForeColor
    PropertyChanged "LabelForeColor"
End Property

'LABEL WIDTH
Public Property Get LabelWidth() As Long
    LabelWidth = m_LabelWidth
End Property

Public Property Let LabelWidth(ByVal New_LabelWidth As Long)
On Error Resume Next
    m_LabelWidth = New_LabelWidth
    mLabel.Width = m_LabelWidth
    PropertyChanged "LabelWidth"
    UserControl.Width = LabelWidth + TextBoxWidth + IIf(TextButton = True, mButton.Width + 15, 0) - IIf(LabelWidth = 0, 0, 60)
                            
    UserControl_Resize
End Property

'TEXTBOX WIDTH
Public Property Get TextBoxWidth() As Long
    TextBoxWidth = m_TextBoxWidth
End Property

Public Property Let TextBoxWidth(ByVal New_TextBoxWidth As Long)
On Error Resume Next
    m_TextBoxWidth = New_TextBoxWidth
    PropertyChanged "TextBoxWidth"
    UserControl.Width = LabelWidth + TextBoxWidth + IIf(TextButton = True, mButton.Width + 15, 0) - IIf(LabelWidth = 0, 0, 60)
    UserControl_Resize
End Property

'LABEL ALLIGNMENT
Public Property Get LabelAlignment() As e_LabelAlignment
    LabelAlignment = m_LabelAlignment
End Property

Public Property Let LabelAlignment(ByVal New_LabelAlignment As e_LabelAlignment)
    m_LabelAlignment = New_LabelAlignment
    PropertyChanged "LabelAlignment"
    If LabelWidth > 100 Then
        mLabel.Width = LabelWidth '- 150
    End If
    If m_LabelAlignment = LeftJustified Then
        mLabel.Alignment = 0
    ElseIf m_LabelAlignment = RightJustified Then
        mLabel.Alignment = 1
    ElseIf m_LabelAlignment = Center Then
        mLabel.Alignment = 2
    End If
    PropertyChanged "LabelAlignment"
End Property

'BackStyle
Public Property Get LabelBackStyle() As e_LabelBackStyle
Attribute LabelBackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    LabelBackStyle = m_LabelBackStyle
End Property

Public Property Let LabelBackStyle(ByVal New_LabelBackStyle As e_LabelBackStyle)
    m_LabelBackStyle = New_LabelBackStyle
    PropertyChanged "LabelBackStyle"
    If m_LabelBackStyle = Transparent Then
        mLabel.BackStyle = 0
        UserControl.BackStyle = 0
    ElseIf m_LabelBackStyle = Opaque Then
        mLabel.BackStyle = 1
        UserControl.BackStyle = 1
    End If
    PropertyChanged "LabelBackStyle"
End Property

'LabelBackColor
Public Property Get LabelBackColor() As OLE_COLOR
    LabelBackColor = m_LabelBackColor
End Property

Public Property Let LabelBackColor(ByVal New_LabelBackColor As OLE_COLOR)
    m_LabelBackColor = New_LabelBackColor
    mLabel.BackColor = m_LabelBackColor
    UserControl.BackColor = m_LabelBackColor
    PropertyChanged "LabelBackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mTextbox,mTextbox,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = mTextbox.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    mTextbox.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mTextbox,mTextbox,-1,DataSource
Public Property Get DataSource() As ADODB.Recordset
Attribute DataSource.VB_Description = "Sets a value that specifies the Data control through which the current control is bound to a database. "
    Set DataSource = mTextbox.DataSource
    'Set DataSource = mBoundText.DataSource
End Property

Public Property Set DataSource(ByRef New_DataSource As ADODB.Recordset)
    Set mTextbox.DataSource = New_DataSource
    'Set mBoundText.DataSource = New_DataSource
    PropertyChanged "DataSource"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=27,0,0,
Public Property Get DataField() As ADODB.Field
    Set DataField = m_DataField
End Property

Public Property Let DataField(ByRef New_DataField As ADODB.Field)
    Set m_DataField = New_DataField
    mTextbox.DataField = m_DataField.Name
    'mBoundText.DataField = m_DataField.Name
    PropertyChanged "DataField"
End Property

Sub TextToBound()
'    If m_DataType = Numeric Then
'        mBoundText = CDbl(mTextbox)
'    Else
'        mBoundText = Trim(mTextbox)
'    End If
End Sub

Sub BoundToText()
Dim s As String, i As Integer
    
'    If mBoundText = vbNullString Then
'        mTextBox = ""
'    ElseIf Trim(mBoundText) = "" Then
'        mTextBox = ""
'    Else
'        mTextBox = mBoundText
'    End If
    
    If m_DataType = Numeric Then
        s = ""
        For i = 1 To m_DecimalPlace
            s = s & "0"
        Next i

        If m_DecimalPlace <> 0 Then
            s = "#,##0." & s
        Else
            s = "#,##0"
        End If
        mTextbox = Format(mTextbox, s)
        mTextbox.SelLength = Len(mTextbox)
    Else
        If m_AllCaps Then
            mTextbox = UCase(mTextbox)
        End If
    End If
        
End Sub

'SendKeys "{Tab}"
Public Property Get SendKeysTab() As Boolean
    SendKeysTab = m_SendKeysTab
End Property

Public Property Let SendKeysTab(ByVal New_SendKeysTab As Boolean)
    m_SendKeysTab = New_SendKeysTab
    PropertyChanged "SendKeysTab"
End Property

'hWnd
Public Property Get hwnd() As Long
    hwnd = mTextbox.hwnd
End Property

'Trim textbox
Public Property Get TextTrim() As Boolean
    TextTrim = m_TextTrim
End Property

Public Property Let TextTrim(ByVal New_TextTrim As Boolean)
    m_TextTrim = New_TextTrim
    PropertyChanged "TextTrim"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Hover() As Boolean
    Hover = m_Hover
End Property

Public Property Let Hover(ByVal New_Hover As Boolean)
    m_Hover = New_Hover
    PropertyChanged "Hover"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Required() As Boolean
    Required = m_Required
End Property

Public Property Let Required(ByVal New_Required As Boolean)
    m_Required = New_Required
    PropertyChanged "Required"
End Property

Private Sub mLabel_Click()
    If m_Hover Then RaiseEvent LabelClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get LinkForeColor() As OLE_COLOR
    LinkForeColor = m_LinkForeColor
End Property

Public Property Let LinkForeColor(ByVal New_LinkForeColor As OLE_COLOR)
    m_LinkForeColor = New_LinkForeColor
    PropertyChanged "LinkForeColor"
End Property

