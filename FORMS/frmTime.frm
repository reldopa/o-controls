VERSION 5.00
Begin VB.Form frmTime 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2235
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   2235
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   0
      Top             =   3240
   End
   Begin VB.TextBox txtTime 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   315
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2715
      Width           =   1035
   End
   Begin VB.ListBox lstAMPM 
      Appearance      =   0  'Flat
      Height          =   2370
      ItemData        =   "frmTime.frx":0000
      Left            =   1710
      List            =   "frmTime.frx":000A
      TabIndex        =   9
      Top             =   300
      Width           =   495
   End
   Begin VB.TextBox txtAMPM 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1725
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   15
      Width           =   465
   End
   Begin VB.ListBox lstSeconds 
      Appearance      =   0  'Flat
      Height          =   2370
      ItemData        =   "frmTime.frx":0016
      Left            =   1140
      List            =   "frmTime.frx":0018
      TabIndex        =   7
      Top             =   300
      Width           =   495
   End
   Begin VB.ListBox lstMinutes 
      Appearance      =   0  'Flat
      Height          =   2370
      ItemData        =   "frmTime.frx":001A
      Left            =   570
      List            =   "frmTime.frx":001C
      TabIndex        =   6
      Top             =   300
      Width           =   495
   End
   Begin VB.ListBox lstHours 
      Appearance      =   0  'Flat
      Height          =   2370
      ItemData        =   "frmTime.frx":001E
      Left            =   0
      List            =   "frmTime.frx":0020
      TabIndex        =   5
      Top             =   300
      Width           =   495
   End
   Begin VB.TextBox txtSeconds 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1155
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   15
      Width           =   465
   End
   Begin VB.TextBox txtMinutes 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   585
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   15
      Width           =   465
   End
   Begin VB.TextBox txtHours 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   15
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   15
      Width           =   465
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFC0C0&
      Height          =   285
      Left            =   300
      Top             =   2700
      Width           =   1065
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   0
      Picture         =   "frmTime.frx":0022
      Top             =   2700
      Width           =   300
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFC0C0&
      Height          =   285
      Left            =   1710
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFC0C0&
      Height          =   285
      Left            =   1140
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   195
      Left            =   1080
      TabIndex        =   3
      Top             =   30
      Width           =   60
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFC0C0&
      Height          =   285
      Left            =   570
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   195
      Left            =   510
      TabIndex        =   1
      Top             =   30
      Width           =   60
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFC0C0&
      Height          =   285
      Left            =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public currTime As String

Private Sub Form_Click()
    ReturnVal
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        lPress = False
        Unload Me
    ElseIf KeyCode = 13 Then
        ReturnVal
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
    lstHours.Clear
    lstMinutes.Clear
    lstSeconds.Clear
    For i = 0 To 12
        lstHours.AddItem Format(i, "00")
    Next i
    For i = 0 To 59
        lstMinutes.AddItem Format(i, "00")
        lstSeconds.AddItem Format(i, "00")
    Next i
    txtTime = currTime
    lstHours.ListIndex = CInt(Mid(currTime, 1, 2))
    lstMinutes.ListIndex = CInt(Mid(currTime, 4, 2))
    lstSeconds.ListIndex = CInt(Mid(currTime, 7, 2))
    lstAMPM.ListIndex = IIf(Right(currTime, 2) = "AM", 0, 1)
End Sub

Private Sub lstAMPM_Click()
    txtAMPM = lstAMPM
    FormatTimeValue
End Sub

Private Sub lstAMPM_DblClick()
    ReturnVal
End Sub

Private Sub lstHours_Click()
    txtHours = lstHours
    FormatTimeValue
End Sub

Private Sub lstHours_DblClick()
    ReturnVal
End Sub
    
Private Sub lstMinutes_Click()
    txtMinutes = lstMinutes
    FormatTimeValue
End Sub

Private Sub lstMinutes_DblClick()
    ReturnVal
End Sub

Private Sub lstSeconds_Click()
    txtSeconds = lstSeconds
    FormatTimeValue
End Sub

Private Sub lstSeconds_DblClick()
    ReturnVal
End Sub

Private Sub Timer1_Timer()
   Dim hF As Long
    DoEvents
    hF = GetFocus
    
    If hF = Me.hwnd Or hF = txtTime.hwnd Or hF = txtHours.hwnd Or hF = txtMinutes.hwnd Or hF = txtSeconds.hwnd Or _
     hF = txtAMPM.hwnd Or hF = lstHours.hwnd Or hF = lstMinutes.hwnd Or hF = lstSeconds.hwnd Or hF = lstAMPM.hwnd Then
       'Focus to form or MV
    Else
        lPress = False
        Unload Me
    End If
End Sub

Private Sub FormatTimeValue()
On Error GoTo ErrHandler
    txtTime = Format(txtHours & ":" & txtMinutes & ":" & txtSeconds & " " & txtAMPM, "hh:nn:ss AM/PM")
ErrHandler:
    Exit Sub
End Sub

Private Sub ReturnVal()
    ctlTextVal = txtTime
    lPress = False
    Unload Me
End Sub

Private Sub txtTime_DblClick()
    ReturnVal
End Sub
