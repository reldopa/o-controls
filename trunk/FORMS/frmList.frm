VERSION 5.00
Begin VB.Form frmList 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2400
   ClientLeft      =   15
   ClientTop       =   0
   ClientWidth     =   2640
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2400
   ScaleWidth      =   2640
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   720
      Top             =   960
   End
   Begin VB.ListBox lstbox 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   2595
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetFocus Lib "user32" () As Long


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        lPress = False
        Unload Me
    ElseIf KeyCode = 13 Then
        Call lstbox_DblClick
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    lstbox.Height = Me.Height - 30
    lstbox.Width = Me.Width - 30
    lstbox.Top = 15
    lstbox.Left = 15
End Sub

Private Sub Form_Unload(Cancel As Integer)
    lPress = False
End Sub

Private Sub lstbox_Click()
    If Me.Tag = 1 Then
        ctlTextVal = lstbox.Text
        lPress = False
        Unload Me
    End If
End Sub

Private Sub lstbox_DblClick()
    If Me.Tag = 0 Then
        ctlTextVal = lstbox.Text
        lPress = False
        Unload Me
    End If
End Sub


Private Sub Timer1_Timer()
   Dim hF As Long
    DoEvents
    hF = GetFocus
    
    If hF = Me.hwnd Or hF = lstbox.hwnd Then
       'Focus to form or MV
    Else
       Unload Me
    End If
End Sub
