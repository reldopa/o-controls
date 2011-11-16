VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmGrid 
   BorderStyle     =   0  'None
   ClientHeight    =   2070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2070
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   3510
      Top             =   1620
   End
   Begin VB.TextBox txtSQLScript 
      Height          =   345
      Left            =   2460
      TabIndex        =   3
      Top             =   5730
      Width           =   4635
   End
   Begin VB.TextBox txtConnection 
      Height          =   345
      Left            =   2520
      TabIndex        =   2
      Top             =   5100
      Width           =   4635
   End
   Begin VB.TextBox txtLookUP 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   1590
      Width           =   4245
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   2778
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
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
         MarqueeStyle    =   3
         AllowRowSizing  =   -1  'True
         AllowSizing     =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iNoOfColumns As Integer
Public iHeight As Integer
Public iWidth As Integer
Public sColumns As Columns
Public iHeadlines As Integer

Private rsTemp As ADODB.Recordset


Private Sub DataGrid1_DblClick()
On Error GoTo ErrHandler
    If rsTemp.RecordCount <> 0 Then sGridText = rsTemp.Fields(0).Value
    Call DataGrid1_LostFocus
ErrHandler:
    If Err.Number <> 0 Then
        sGridText = ""
        Resume Next
    End If
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
On Error GoTo ErrHandler
Dim strField As String
    strField = DataGrid1.Columns(ColIndex).DataField
    If rsTemp.Sort = strField & " ASC" Then
        rsTemp.Sort = strField & " DESC"
    Else: rsTemp.Sort = strField & " ASC"
    End If
ErrHandler:
    Exit Sub
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        Call DataGrid1_DblClick
    Else
        If KeyAscii = 8 Then
            txtLookUP.Text = Mid(txtLookUP, 1, Len(txtLookUP) - 1)
        ElseIf KeyAscii = 27 Then
            txtLookUP.Text = txtLookUP
        Else
            txtLookUP.Text = txtLookUP & UCase(Chr(KeyAscii))
        End If
        rsTemp.MoveFirst
        rsTemp.Find FieldLookUp(rsTemp, txtLookUP) & " LIKE '%" & txtLookUP & "%'", , adSearchForward
        DataGrid1.Row = rsTemp.Bookmark - 1
    End If
End Sub

Private Function FieldLookUp(rs As ADODB.Recordset, sText As String) As String
On Error Resume Next
Dim i As Integer
Dim rsF As New ADODB.Recordset
    Set rsF.DataSource = rs.Clone
    If rsF.RecordCount > 0 Then
        rsF.MoveFirst
    Else
        Exit Function
    End If

    For i = 0 To rsF.Fields.Count - 1
        rsF.Filter = rsF.Fields(i).Name & " LIKE '%" & sText & "%'"
        If Not (rsF.BOF Or rsF.EOF) Then
            FieldLookUp = rsF.Fields(i).Name
            Exit For
        End If
    Next i
End Function

Private Sub DataGrid1_LostFocus()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Call DataGrid1_LostFocus
End Sub


Private Sub SetupGrid()
On Error GoTo ErrHandler
Dim i As Integer
    
    For i = 0 To sColumns.Count - 1
            If i > 1 Then DataGrid1.Columns.Add i
            DataGrid1.Columns(i).Caption = IIf(sColumns(i).Caption = "", "", sColumns(i).Caption)
            DataGrid1.Columns(i).Width = sColumns(i).Width
    Next i
ErrHandler:
    If Err.Number <> 0 Then Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler

   

    Me.Height = iHeight
    Me.Width = iWidth
    DataGrid1.Headlines = iHeadlines

    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sRecordset, sConnection, adOpenStatic, adLockReadOnly
    Set DataGrid1.DataSource = Nothing
    Set DataGrid1.DataSource = rsTemp
    
    
    SetupGrid
    
ErrHandler:
    Exit Sub
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_Resize()
    DataGrid1.Top = 0
    DataGrid1.Left = 0
    txtLookUP.Top = Me.ScaleHeight - txtLookUP.Height
    txtLookUP.Left = 0
    
    DataGrid1.Width = Me.ScaleWidth
    DataGrid1.Height = Me.ScaleHeight - txtLookUP.Height
    txtLookUP.Width = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rsTemp.State = 1 Then rsTemp.Close
    Set rsTemp = Nothing
End Sub


Private Sub Timer1_Timer()
If txtLookUP.BackColor = &HC0FFFF Then txtLookUP.BackColor = &H80000005 Else txtLookUP.BackColor = &HC0FFFF
End Sub
