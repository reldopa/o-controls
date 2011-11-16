Attribute VB_Name = "modITGControls"
Option Explicit

Public dText As String
Public sGridText As String
Public nColIndex As Integer
Public sCaption(0 To 100) As String
Public sRecordset As String
Public sConnection As String
Public ctlTextVal As Object
Public lPress As Boolean
'Public Const HandCursor = 32649&
'Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
'Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
'
'Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'
'Public Type POINTAPI
'    X As Long
'    Y As Long
'End Type
