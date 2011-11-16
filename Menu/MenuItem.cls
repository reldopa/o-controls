VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarCaption As String
Private mvarKey As String
Private mvarIndex As Integer
Private mvarEnabled As Boolean
Private mvarVisible As Boolean
Private mvarRelative As String
Private mstrID As String
Private mvarIconPic As StdPicture '
Public Expanded As Boolean
Private mvarTop As Integer
Private mvarLevel As Integer
Private mvarMenuBookmark As Integer
Private mvarsDash As String
Private mvarForeColor As OLE_COLOR  'local copy
Private mvarFontSize As Integer
Private mvarBold As Boolean
Private mvarPicture As StdPicture


Public Property Set Picture(ByVal vData As StdPicture)
    Set mvarPicture = vData
End Property

Public Property Get Picture() As StdPicture
    Set Picture = mvarPicture
End Property

Public Property Let FontBold(ByVal vData As Boolean)
    mvarBold = vData
End Property

Public Property Get FontBold() As Boolean
    FontBold = mvarBold
End Property

Public Property Let FontSize(ByVal vData As Integer)
    mvarFontSize = vData
End Property

Public Property Get FontSize() As Integer
    FontSize = mvarFontSize
End Property

Public Property Let ForeColor(ByVal vData As OLE_COLOR)
    mvarForeColor = vData
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mvarForeColor
End Property

Public Property Let sDash(ByVal vData As String)
    mvarsDash = vData
End Property

Public Property Get sDash() As String
    sDash = mvarsDash
End Property

Public Property Let MenuBookmark(ByVal vData As Integer)
    mvarMenuBookmark = vData
End Property

Public Property Get MenuBookmark() As Integer
    MenuBookmark = mvarMenuBookmark
End Property

Public Property Let Level(ByVal vData As Integer)
    mvarLevel = vData
End Property

Public Property Get Level() As Integer
    Level = mvarLevel
End Property

Public Property Let Top(ByVal vData As Integer)
    mvarTop = vData
End Property

Public Property Get Top() As Integer
    Top = mvarTop
End Property

Public Property Set IconPic(ByVal vData As StdPicture)
    Set mvarIconPic = vData
End Property

Public Property Get IconPic() As StdPicture
    Set IconPic = mvarIconPic
End Property


Public Property Let Relative(ByVal vData As String)
    mvarRelative = vData
End Property

Public Property Get Relative() As String
    Relative = mvarRelative
End Property

Public Property Let Visible(ByVal vData As Boolean)
    mvarVisible = vData
End Property

Public Property Get Visible() As Boolean
    Visible = mvarVisible
End Property

Public Property Let Enabled(ByVal vData As Boolean)
    mvarEnabled = vData
End Property

Public Property Get Enabled() As Boolean
    Enabled = mvarEnabled
End Property

Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property

Public Property Get Key() As String
    Key = mvarKey
End Property

Public Property Let index(ByVal vData As Integer)
    mvarIndex = vData
End Property

Public Property Get index() As Integer
    index = mvarIndex
End Property

Public Property Let Caption(ByVal vData As String)
    mvarCaption = vData
End Property

Public Property Get Caption() As String
    Caption = mvarCaption
End Property

Property Get ID() As String
    ID = mstrID
End Property

Private Sub Class_Initialize()
    mstrID = TypeName(Me) & CLng(Rnd * (2 ^ 31))
End Sub
