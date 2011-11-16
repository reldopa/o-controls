VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mColMenuItem As Collection

Private Sub Class_Initialize()
    Set mColMenuItem = New Collection
End Sub

Public Sub Add(ByVal Index As Variant, Optional ByVal Relative As String, Optional ByVal nLevel As Integer, _
Optional ByVal sText As String, Optional ByVal PIcon As StdPicture, Optional ByVal MenuPic As StdPicture, _
Optional ByVal lForeColor As OLE_COLOR, Optional ByVal lBold As Boolean, Optional ByVal nFontSize As Integer)

    Dim objMenu As New MenuItem
    objMenu.Key = Index
    objMenu.Caption = sText
    objMenu.Relative = Relative
    objMenu.Level = nLevel
    objMenu.Enabled = True
    objMenu.ForeColor = lForeColor
    objMenu.FontBold = lBold
    Set objMenu.Picture = MenuPic
    objMenu.FontSize = nFontSize
    Set objMenu.IconPic = PIcon
    objMenu.Visible = False
    mColMenuItem.Add objMenu, objMenu.ID
    objMenu.Index = mColMenuItem.Count - 1
    objMenu.MenuBookmark = mColMenuItem.Count
    
    objMenu.sDash = SetMenuDash(Relative) & Index & "\"
End Sub

Private Function SetMenuDash(sRelative As String) As String
Dim i As Integer
Dim sDashParent As String
    For i = 0 To mColMenuItem.Count - 1
        If mColMenuItem.Item(i + 1).Key = sRelative Then
            sDashParent = mColMenuItem.Item(i + 1).sDash
            Exit For
        End If
    Next i
    SetMenuDash = "\" & sDashParent
End Function

Property Get Item(ByVal Index As Variant) As MenuItem
Attribute Item.VB_UserMemId = 0
    Set Item = mColMenuItem.Item(Index)
End Property

Public Sub Remove(ByVal varID As Variant)
    mColMenuItem.Remove varID
End Sub

Property Get Count() As Long
    Count = mColMenuItem.Count
End Property

Property Let Expanded(ByVal fExpanded As Boolean)
Dim objMenu As MenuItem
    For Each objMenu In mColMenuItem
        objMenu.Expanded = fExpanded
    Next
End Property

Property Get Expanded() As Boolean
Dim objMenu As MenuItem
    For Each objMenu In mColMenuItem
        If objMenu.Expanded Then
            Expanded = True
            Exit For
        End If
    Next
End Property


