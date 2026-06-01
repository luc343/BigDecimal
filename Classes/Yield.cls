VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Yield"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------
'
'                   Class Yield
'
' Copyright (c) Lucien Cinc 2025-26
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
'---------------------------------------------------

Private Counter As LongPtr  'internal counter
Private MaxCounter As Long  'count before doevents triggered

'########################### Yield Methods ###########################

Private Sub Class_Initialize()
    Counter = 0
    MaxCounter = 500    'good trade-off between performance & not responding
End Sub

'***************************************
'
'              Do Events
'
'***************************************

Public Sub Events()
    Counter = Counter + 1

    If Counter Mod MaxCounter = 0 Then
	DoEvents
    End If
End Sub

'***************************************
'
'           Clear the Counter
'
'***************************************

Public Sub Clear()
    Counter = 0
End Sub

'######################### Yield Properties ##########################

'MaxCount
Public Property Get MaxCount() As Long
    MaxCount = MaxCounter
End Property

Public Property Let MaxCount(Val As Long)
    MaxCounter = Val
End Property
