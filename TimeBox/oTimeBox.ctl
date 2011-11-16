VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl oTimeBox 
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2925
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
   ScaleHeight     =   525
   ScaleWidth      =   2925
   Begin VB.TextBox mTextbox 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   570
      Width           =   135
   End
   Begin MSMask.MaskEdBox mTimeBox 
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
      Left            =   1275
      TabIndex        =   0
      Top             =   15
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:nn:ss AM/PM"
      Mask            =   "##:##:## &&"
      PromptChar      =   "_"
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Caption"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   555
   End
   Begin VB.Image img 
      Height          =   255
      Left            =   2640
      Picture         =   "oTimeBox.ctx":0000
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFC0C0&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   1260
      Top             =   0
      Width           =   1665
   End
End
Attribute VB_Name = "oTimeBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const m_def_CaptionWidth = 1200

Dim m_CaptionWidth  As Integer

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CaptionWidth() As Integer
    CaptionWidth = m_CaptionWidth
End Property

Public Property Let CaptionWidth(ByVal New_CaptionWidth As Integer)
On Error Resume Next
    m_CaptionWidth = New_CaptionWidth
    UserControl.Width = UserControl.Width + (m_CaptionWidth - Shape.Width)
    PropertyChanged "CaptionWidth"
    Call UserControl_Resize
End Property


Private Sub img_Click()
Dim frm As Form
Dim dbgWindows As RECT
Dim i As Integer
If img.Enabled Then
    Set frm = frmTime
    If lPress = False Then
        lPress = True
    Else
        lPress = False
    End If
    If lPress = True Then
        frm.currTime = IIf(mTimeBox = "__:__:__ __", Format(Now, "hh:nn:ss AM/PM"), mTimeBox)
        GetWindowRect mTimeBox.hwnd, dbgWindows
        frm.Left = dbgWindows.Left * Screen.TwipsPerPixelX
        frm.Top = (dbgWindows.Top * Screen.TwipsPerPixelY) + (UserControl.Height + 15)
        frm.Show , Me
    End If
End If
End Sub

Private Sub mTimeBox_GotFocus()
    Set ctlTextVal = Nothing
    Set ctlTextVal = mTimeBox
End Sub

Private Sub UserControl_GotFocus()
    Set ctlTextVal = Nothing
    Set ctlTextVal = mTimeBox
End Sub

Private Sub UserControl_InitProperties()
    m_CaptionWidth = m_def_CaptionWidth
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    CaptionWidth = PropBag.ReadProperty("CaptionWidth", m_def_CaptionWidth)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    If UserControl.Height > 285 Then UserControl.Height = 285
    lbl.Top = 30
    lbl.Left = 30
    Shape.Height = UserControl.Height
    Shape.Width = UserControl.Width - (m_CaptionWidth + 30)
    Shape.Top = 0
    Shape.Left = m_CaptionWidth + 30
    img.Height = UserControl.Height - 30
    img.Top = 15
    img.Left = UserControl.Width - img.Width
    mTimeBox.Height = UserControl.Height - 30
    mTimeBox.Width = Shape.Width - (img.Width + 30)
    mTimeBox.Top = 15
    mTimeBox.Left = Shape.Left + 15
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("CaptionWidth", m_CaptionWidth, m_def_CaptionWidth)
End Sub
