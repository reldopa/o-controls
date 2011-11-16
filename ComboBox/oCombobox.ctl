VERSION 5.00
Begin VB.UserControl oCombobox 
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3300
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
   ScaleHeight     =   345
   ScaleWidth      =   3300
   Begin VB.TextBox txt 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1335
      TabIndex        =   0
      Top             =   15
      Width           =   1635
   End
   Begin VB.ListBox lst 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   1050
      TabIndex        =   2
      Top             =   540
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Image img1 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   2985
      Picture         =   "oCombobox.ctx":0000
      Top             =   15
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      FillColor       =   &H00FF8080&
      Height          =   285
      Left            =   2970
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   1200
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFC0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   1320
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "oCombobox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const m_def_AllCaps = False
Const m_def_BackColor = &H80000005
Const m_def_BorderColor = &HFFC0C0
Const m_def_CaptionWidth = 1200
Const m_def_CaptionBackColor = &H8000000F
Const m_def_ConnectionStrings = ""
Const m_def_Enabled = True
Const m_def_ForeColor = 0
Const m_def_Font = "Tahoma"
Const m_def_Mandatory = False
Const m_def_SendKeysTab = False
Const m_def_SQLScript = ""
Const m_def_Modal = False

Dim m_AllCaps As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_BorderColor As OLE_COLOR
Dim m_Caption As String
Dim m_CaptionWidth As Integer
Dim m_Style As e_Style
Dim m_CaptionBackColor As OLE_COLOR
Dim m_DataField As ADODB.Field
Dim m_Enabled As Boolean
Dim m_ForeColor As OLE_COLOR
Dim m_Font As Font
Dim m_Mandatory As Boolean
Dim m_SendKeysTab As Boolean
Dim m_SQLScript As String
Dim m_Text As String
Dim m_DataListField As String
Dim m_Modal As Boolean
Dim m_ListPosition As e_CalendarPosition

Dim lHandle As Long
Private LB As POINTAPI

Dim nIndex As Integer

Public Enum e_Style
    DropDownCombo
    SimpleCombo
End Enum

Event Change()
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,1
Public Property Get AllCaps() As Boolean
    AllCaps = m_AllCaps
End Property

Public Property Let AllCaps(ByVal New_AllCaps As Boolean)
    m_AllCaps = New_AllCaps
    PropertyChanged "AllCaps"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    txt.BackColor = m_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    Shape.BackColor = m_BorderColor
    PropertyChanged "BorderColor"
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

''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,BackStyle
'Public Property Get BackStyle() As e_BackStyle
'    BackStyle = UserControl.BackStyle
'End Property
'
'Public Property Let BackStyle(ByVal New_BackStyle As e_BackStyle)
'    If New_BackStyle = Opaque Or New_BackStyle = Transparent Then
'        UserControl.BackStyle() = New_BackStyle
'        PropertyChanged "BackStyle"
'    Else
'        Err.Raise Number:=vbObjectError + 1001, _
'                     Description:="Invalid BackStyle value (0 or 1 Only)"
'    End If
'End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal new_caption As String)
    m_Caption = new_caption
    lbl.Caption = m_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CaptionBackColor() As OLE_COLOR
    CaptionBackColor = m_CaptionBackColor
End Property

Public Property Let CaptionBackColor(ByVal New_CaptionBackColor As OLE_COLOR)
    m_CaptionBackColor = New_CaptionBackColor
    UserControl.BackColor = m_CaptionBackColor
    PropertyChanged "CaptionBackColor"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CaptionWidth() As Integer
    CaptionWidth = m_CaptionWidth
End Property

Public Property Let CaptionWidth(ByVal New_CaptionWidth As Integer)
On Error Resume Next
    m_CaptionWidth = New_CaptionWidth
    UserControl.Width = UserControl.Width + (m_CaptionWidth - lbl.Width)
    lbl.Width = m_CaptionWidth
    PropertyChanged "CaptionWidth"
    Call UserControl_Resize
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=27,0,0,
Public Property Get DataField() As ADODB.Field
    Set DataField = m_DataField
End Property

Public Property Let DataField(ByRef New_DataField As ADODB.Field)
    Set m_DataField = New_DataField
    txt.DataField = m_DataField.Name
    PropertyChanged "DataField"
End Property

Public Property Get DataListField() As String
    DataListField = m_DataListField
End Property

Public Property Let DataListField(ByRef New_DataListField As String)
    m_DataListField = New_DataListField
    Call LoadDataList(DataSource, SQLScript, DataListField)
    PropertyChanged "DataListField"
End Property

Private Sub LoadDataList(rsSource As ADODB.Recordset, sSQL As String, sField As String)
Dim rs As New ADODB.Recordset
    rs.Open sSQL, rsSource.ActiveConnection, adOpenStatic, adLockReadOnly

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Text = rs.Fields(sField).Value
        Do Until rs.EOF
            Me.AddItem rs.Fields(sField).Value
            rs.MoveNext
        Loop
    End If
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txt,txt,-1,DataSource
Public Property Get DataSource() As ADODB.Recordset
    Set DataSource = txt.DataSource
End Property

Public Property Set DataSource(ByRef New_DataSource As ADODB.Recordset)
    Set txt.DataSource = New_DataSource
    PropertyChanged "DataSource"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txt,txt,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    txt.Locked = Not m_Enabled
    img1.Enabled = m_Enabled
    PropertyChanged "Enabled"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mDateBox,mDateBox,-1,Enabled
Public Property Get ListPosition() As e_CalendarPosition
    ListPosition = m_ListPosition
End Property

Public Property Let ListPosition(ByVal New_ListPosition As e_CalendarPosition)
    m_ListPosition = New_ListPosition
    PropertyChanged "ListPosition"
End Property


Public Property Get Style() As e_Style
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As e_Style)
    m_Style = New_Style
    PropertyChanged "Style"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    Set txt.Font = m_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    lbl.ForeColor = m_ForeColor
    txt.ForeColor = m_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Mandatory() As Boolean
    Mandatory = m_Mandatory
End Property

Public Property Let Mandatory(ByVal New_Mandatory As Boolean)
    m_Mandatory = New_Mandatory
    PropertyChanged "Mandatory"
    If m_Mandatory Then
        'txt.BackColor = &HE0FFFF  '&H80000018  '&HC0FFFF
        'BackColor = &HE0FFFF  '&H80000018  '&HC0FFFF
        'lbl.ForeColor = &HFF0000
    Else
        'txt.BackColor = &HFFFFFF
        'BackColor = &HFFFFFF
        'lbl.ForeColor = &H800000
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

Public Property Get SQLScript() As String
    SQLScript = m_SQLScript
End Property

Public Property Let SQLScript(ByVal New_SQLScript As String)
    m_SQLScript = New_SQLScript
    'frmGrid.sRecordset = m_SQLScript
    PropertyChanged "SQLScript"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,
Public Property Get Text() As String
    m_Text = txt.Text
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
On Error GoTo ErrHandler
    m_Text = New_Text
    txt.Text = m_Text
    PropertyChanged "Text"
ErrHandler:
    If Err.Number <> 0 Then Resume Next
End Property

Private Sub img1_Click()
Dim frm As Form
Dim dbgWindows As RECT
Dim i As Integer
On Error GoTo ErrHandler
If img1.Enabled Then
    Set frm = frmList
    If lPress = False Then
        lPress = True
    Else
        lPress = False
    End If
    If lPress = True Then
        GetWindowRect txt.hwnd, dbgWindows
        
        frm.Left = dbgWindows.Left * Screen.TwipsPerPixelX
        If m_ListPosition = Bottom Then
            frm.Top = (dbgWindows.Top * Screen.TwipsPerPixelY) + UserControl.Height
        Else
            frm.Top = (dbgWindows.Top * Screen.TwipsPerPixelY) - frm.Height
        End If
        
        'frm.Top = (dbgWindows.Top * Screen.TwipsPerPixelY) + (UserControl.Height + 15)
        frm.Font = txt.Font
        frm.Width = txt.Width + img1.Width + 30
        frm.Tag = m_Style
        For i = 1 To lst.ListCount - 1
            frm.lstbox.AddItem lst.List(i)
        Next i
        DoEvents
        If m_Modal Then frm.Show vbModal, Me Else frm.Show , Me
        'RaiseEvent Change
    End If
End If
ErrHandler:
    If Err.Number <> 0 Then
        If Err.Number = 401 Then
            MsgBox "Please set the control into modal state."
        ElseIf Err.Number = 3021 Then
            Exit Sub
        ElseIf Err.Number = 91 Then
            Exit Sub
        Else
            MsgBox Err.Number & ": " & Err.Description
        End If
    End If
End Sub

Private Sub img1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.Visible = True
End Sub

Private Sub img1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lHandle = LoadCursor(0, HandCursor)
    If (lHandle > 0) Then SetCursor lHandle
End Sub

Private Sub img1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.Visible = False
End Sub

Private Sub lst_GotFocus()
    SendKeys "{Tab}"
End Sub

Private Sub txt_Change()
On Error Resume Next
    PropertyChanged ("Text")
    m_Text = txt.Text
    m_DataField.Value = m_Text
    RaiseEvent Change
    If m_Style = 1 Then RaiseEvent Click
    If m_Style = 0 Then RaiseEvent DblClick
End Sub

Private Sub txt_Click()
    RaiseEvent Click
End Sub

Private Sub txt_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub txt_GotFocus()
    Set ctlTextVal = Nothing
    Set ctlTextVal = txt
    txt.SelStart = 0
    txt.SelLength = Len(txt.Text)
    'RaiseEvent Click
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If m_SendKeysTab = True Then
            SendKeys "{Tab}"
        End If
    ElseIf KeyAscii = 9 Or KeyAscii = 0 Then
        SendKeys "{Tab}"
    End If
    
    If IsNumeric(m_Text) Then
        Beep
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

Private Sub txt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 115 Then 'show frmList
        Call img1_Click
    End If
    If KeyCode = 40 Then 'Press Down Arrow
        If lPress = False Then
            If nIndex < 0 Then nIndex = 0
            nIndex = nIndex + 1
            If nIndex >= Me.ListCount Then nIndex = Me.ListCount - 1
            Me.Text = Me.List(nIndex)
            txt.SelStart = 0
            txt.SelLength = Len(txt)
        End If
    ElseIf KeyCode = 38 Then 'Press Up Arrow
        If lPress = False Then
            If nIndex < 0 Then nIndex = 0
            nIndex = nIndex - 1
            If nIndex >= Me.ListCount Then nIndex = Me.ListCount - 1
            Me.Text = Me.List(nIndex)
            txt.SelStart = 0
            txt.SelLength = Len(txt)
        End If
    End If
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_GotFocus()
    Set ctlTextVal = Nothing
    Set ctlTextVal = txt
End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next
    m_AllCaps = m_def_AllCaps
    m_BackColor = m_def_BackColor
    m_BorderColor = m_def_BorderColor
    m_Caption = "Caption"
    m_CaptionBackColor = m_def_CaptionBackColor
    m_CaptionWidth = m_def_CaptionWidth
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_Font = m_def_Font
    m_ForeColor = m_def_ForeColor
    m_Mandatory = m_def_Mandatory
    m_SendKeysTab = m_def_SendKeysTab
    m_SQLScript = m_def_SQLScript
    m_Modal = m_def_Modal
    m_Text = ""
    m_DataListField = ""
    UserControl.Height = 285
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer
    Modal = PropBag.ReadProperty("Modal", m_def_Modal)
    AllCaps = PropBag.ReadProperty("AllCaps", m_def_AllCaps)
    BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    ListPosition = PropBag.ReadProperty("ListPosition", 0)
    Style = PropBag.ReadProperty("Style", 0)
    Caption = PropBag.ReadProperty("Caption", "Caption")
    CaptionWidth = PropBag.ReadProperty("CaptionWidth", m_def_CaptionWidth)
    CaptionBackColor = PropBag.ReadProperty("CaptionBackColor", m_def_CaptionBackColor)
    Set DataSource = PropBag.ReadProperty("DataSource", Nothing)
    Set m_DataField = PropBag.ReadProperty("DataField", Nothing)
    m_DataListField = PropBag.ReadProperty("DataListField", "")
    ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Mandatory = PropBag.ReadProperty("Mandatory", m_def_Mandatory)
    SendKeysTab = PropBag.ReadProperty("SendKeysTab", m_def_SendKeysTab)
    SQLScript = PropBag.ReadProperty("SQLScript", m_def_SQLScript)
    Text = PropBag.ReadProperty("Text", "")
    lst.List(Index) = PropBag.ReadProperty("List" & Index, "")
    lst.ListIndex = PropBag.ReadProperty("ListIndex", 0)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    If UserControl.Height > 285 Then UserControl.Height = 285
    lbl.Top = 30
    lbl.Left = 45
    lbl.Height = UserControl.Height
    lbl.Width = m_CaptionWidth
    img1.Width = 255
    Shape.Top = 0
    Shape.Left = lbl.Width + 15
    Shape.Height = UserControl.Height + 15
    Shape.Width = UserControl.Width - (lbl.Width + 30)
    txt.Top = 15
    txt.Left = Shape.Left + 15
    txt.Height = (UserControl.Height + 15) - 45
    txt.Width = Shape.Width - (30 + (img1.Width + 30))
    img1.Top = 15
    img1.Left = UserControl.Width - (img1.Width + 45)
    img1.Height = (UserControl.Height + 15) - 45
    
    Shape1.Height = UserControl.Height
    Shape1.Width = img1.Width + 30
    Shape1.Top = 0
    Shape1.Left = img1.Left - 15
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer
    Call PropBag.WriteProperty("Modal", m_Modal, m_def_Modal)
    Call PropBag.WriteProperty("AllCaps", m_AllCaps, m_def_AllCaps)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("Caption", m_Caption, "Caption")
    Call PropBag.WriteProperty("ListPosition", m_ListPosition, 0)
    Call PropBag.WriteProperty("Style", m_Style, 0)
    Call PropBag.WriteProperty("CaptionBackColor", m_CaptionBackColor, m_def_CaptionBackColor)
    Call PropBag.WriteProperty("CaptionWidth", m_CaptionWidth, m_def_CaptionWidth)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("DataSource", DataSource, Nothing)
    Call PropBag.WriteProperty("DataField", m_DataField, Nothing)
    Call PropBag.WriteProperty("DataListField", m_DataListField, "")
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Mandatory", m_Mandatory, m_def_Mandatory)
    Call PropBag.WriteProperty("SendKeysTab", m_SendKeysTab, m_def_SendKeysTab)
    Call PropBag.WriteProperty("SQLScript", m_SQLScript, m_def_SQLScript)
    Call PropBag.WriteProperty("Text", m_Text, "")
    Call PropBag.WriteProperty("List" & Index, lst.List(Index), "")
    Call PropBag.WriteProperty("ListIndex", lst.ListIndex, 0)
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
'MappingInfo=lst,lst,-1,AddItem
Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
Attribute AddItem.VB_Description = "Adds an item to a Listbox or ComboBox control or a row to a Grid control."
    lst.AddItem Item, Index
End Sub

Public Sub Clear()
Dim i As Integer
    Do Until (lst.ListCount - 1) = 0
        lst.RemoveItem (lst.ListCount)
    Loop
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lst,lst,-1,List
Public Property Get List(ByVal Index As Integer) As String
Attribute List.VB_Description = "Returns/sets the items contained in a control's list portion."
    List = lst.List(Index)
End Property

Public Property Let List(ByVal Index As Integer, ByVal New_List As String)
    lst.List(Index) = New_List
    PropertyChanged "List"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lst,lst,-1,ListCount
Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
    ListCount = lst.ListCount
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lst,lst,-1,ListIndex
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    ListIndex = lst.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    lst.ListIndex() = New_ListIndex
    txt.Text = lst.Text
    PropertyChanged "ListIndex"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lst,lst,-1,RemoveItem
Public Sub RemoveItem(ByVal Index As Integer)
Attribute RemoveItem.VB_Description = "Removes an item from a ListBox or ComboBox control or a row from a Grid control."
    lst.RemoveItem Index
End Sub


