VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMonthView 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2370
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   5475
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMonthView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   690
      Top             =   960
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2310
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   42139649
      TitleBackColor  =   16761024
      TitleForeColor  =   8388608
      CurrentDate     =   37802
   End
End
Attribute VB_Name = "frmMonthView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public currDate As String

Private Declare Function GetFocus Lib "user32" () As Long

Private Sub Form_Load()
On Error GoTo ErrHandler
    MonthView1.Value = currDate
ErrHandler:
    If Err.Number <> 0 Then
        If Err.Number = 13 Then
            MonthView1.Value = Date
        Else
            MsgBox Err.Description
        End If
    End If
End Sub

Private Sub Form_Resize()
    Me.Height = MonthView1.Height
    Me.Width = MonthView1.Width
    MonthView1.Top = 0
    MonthView1.Left = 0
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    dText = Format(Str(Me.MonthView1.Month) & "/" & Str(MonthView1.Day) & "/" & Str(MonthView1.Year), "mm/dd/yyyy")
    ctlTextVal = dText
    Unload Me
End Sub

Private Sub MonthView1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        dText = "__/__/____"
        Unload Me
    End If
End Sub

Private Sub Timer1_Timer()
    Dim hF As Long
    DoEvents
    hF = GetFocus
    
    If hF = Me.hwnd Or hF = MonthView1.hwnd Then
       'Focus to form or MV
    Else
        dText = ""
       Unload Me
    End If
End Sub
