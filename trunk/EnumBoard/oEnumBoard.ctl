VERSION 5.00
Begin VB.UserControl oEnumBoard 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   LockControls    =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   3240
   Begin VB.ListBox lst 
      Height          =   1230
      Left            =   1080
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label lblDetails 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Details"
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
      Left            =   210
      TabIndex        =   1
      Top             =   390
      Width           =   585
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Header"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   615
   End
   Begin VB.Image imgHeader 
      Height          =   345
      Left            =   0
      Picture         =   "oEnumBoard.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "oEnumBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_HeaderHeight = 345
'Property Variables:
Dim m_HeaderHeight As Variant
'Event Declarations:
Event Click(DetailsIndex As Integer)  'MappingInfo=UserControl,UserControl,-1,Click
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

Private Sub lblDetails_Click(DetailsIndex As Integer)
    RaiseEvent Click(DetailsIndex)
End Sub

Private Sub lblDetails_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
    For i = 0 To lblDetails.Count - 1
        lblDetails(i).FontBold = False
    Next i
    lblDetails(Index).FontBold = True
End Sub

Private Sub UserControl_Resize()
Dim i As Integer, initTop As Integer
    
    imgHeader.Top = 0
    imgHeader.Left = 0
    imgHeader.Width = UserControl.Width
    imgHeader.Height = m_HeaderHeight
    lbl.Top = (imgHeader.Height - lbl.Height) / 2
    lbl.Left = 150
    initTop = imgHeader.Height + 45
    For i = 0 To lblDetails.Count - 1
        lblDetails(i).Top = initTop
        initTop = (initTop + lblDetails(i).Height) + 90
    Next i
    
    UserControl.Height = lblDetails(lblDetails.Count - 1).Top + lblDetails(lblDetails.Count - 1).Height + 105

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
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
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
Dim i As Integer
    For i = 0 To lblDetails.Count - 1
        lblDetails(i).FontBold = False
    Next i
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
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
'MappingInfo=lst,lst,-1,AddItem
Public Sub AddItem(ByVal Item As String, ByVal Index As Variant, Optional ByVal Key As String)
Attribute AddItem.VB_Description = "Adds an item to a Listbox or ComboBox control or a row to a Grid control."
    lst.AddItem Key, Index
    If Index > 0 Then
        Load lblDetails(Index)
        lblDetails(Index).Visible = True
    End If
    lblDetails(Index).Caption = Item
    Call UserControl_Resize
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_HeaderHeight = m_def_HeaderHeight
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    lst.List(Index) = PropBag.ReadProperty("List" & Index, "")
    lbl.Caption = PropBag.ReadProperty("HeaderCaption", "Header")
    m_HeaderHeight = PropBag.ReadProperty("HeaderHeight", m_def_HeaderHeight)
    Set lbl.Font = PropBag.ReadProperty("HeaderFont", lbl.Font)
    Set lblDetails(0).Font = PropBag.ReadProperty("DetailsFont", lblDetails(0).Font)
    lbl.ForeColor = PropBag.ReadProperty("HeaderForeColor", &H80000012)
    lblDetails(0).ForeColor = PropBag.ReadProperty("DetailsForeColor", &H80000012)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("List" & Index, lst.List(Index), "")
    Call PropBag.WriteProperty("HeaderCaption", lbl.Caption, "Header")
    Call PropBag.WriteProperty("HeaderHeight", m_HeaderHeight, m_def_HeaderHeight)
    Call PropBag.WriteProperty("HeaderFont", lbl.Font, lbl.Font)
    Call PropBag.WriteProperty("DetailsFont", lblDetails(0).Font, lblDetails(0).Font)
    Call PropBag.WriteProperty("HeaderForeColor", lbl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("DetailsForeColor", lblDetails(0).ForeColor, &H80000012)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl,lbl,-1,Caption
Public Property Get HeaderCaption() As String
Attribute HeaderCaption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    HeaderCaption = lbl.Caption
End Property

Public Property Let HeaderCaption(ByVal New_HeaderCaption As String)
    lbl.Caption() = New_HeaderCaption
    PropertyChanged "HeaderCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get HeaderHeight() As Variant
    HeaderHeight = m_HeaderHeight
End Property

Public Property Let HeaderHeight(ByVal New_HeaderHeight As Variant)
    m_HeaderHeight = New_HeaderHeight
    imgHeader.Height = New_HeaderHeight
    PropertyChanged "HeaderHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl,lbl,-1,Font
Public Property Get HeaderFont() As Font
Attribute HeaderFont.VB_Description = "Returns a Font object."
    Set HeaderFont = lbl.Font
End Property

Public Property Set HeaderFont(ByVal New_HeaderFont As Font)
    Set lbl.Font = New_HeaderFont
    PropertyChanged "HeaderFont"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblDetails(0),lblDetails,0,Font
Public Property Get DetailsFont() As Font
Attribute DetailsFont.VB_Description = "Returns a Font object."
    Set DetailsFont = lblDetails(0).Font
End Property

Public Property Set DetailsFont(ByVal New_DetailsFont As Font)
Dim i As Integer
    For i = 0 To lblDetails.Count - 1
        Set lblDetails(i).Font = New_DetailsFont
    Next i
    PropertyChanged "DetailsFont"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lbl,lbl,-1,ForeColor
Public Property Get HeaderForeColor() As OLE_COLOR
Attribute HeaderForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    HeaderForeColor = lbl.ForeColor
End Property

Public Property Let HeaderForeColor(ByVal New_HeaderForeColor As OLE_COLOR)
    lbl.ForeColor() = New_HeaderForeColor
    PropertyChanged "HeaderForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblDetails(0),lblDetails,0,ForeColor
Public Property Get DetailsForeColor() As OLE_COLOR
Attribute DetailsForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    DetailsForeColor = lblDetails(0).ForeColor
End Property

Public Property Let DetailsForeColor(ByVal New_DetailsForeColor As OLE_COLOR)
Dim i As Integer
    For i = 0 To lblDetails.Count - 1
        lblDetails(i).ForeColor() = New_DetailsForeColor
    Next i
    PropertyChanged "DetailsForeColor"
End Property

