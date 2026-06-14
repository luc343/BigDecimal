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

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const VK_ESCAPE As Long = &H1B      'escape key

Private Counter As Long     'internal DoEvents counter
Private Counter2 As Long    'internal escape key counter

Private MaxCounter As Long  'count before doevents triggered

'########################### Yield Methods ###########################

Private Sub Class_Initialize()
    Counter = 0
    Counter2 = 0
    MaxCounter = 500    'good trade-off between performance & not responding
End Sub

'***************************************
'
'              Do Events
'
'***************************************

Public Sub Events()
    Counter = Counter + 1
    Counter2 = Counter2 + 1

    If Counter = MaxCount Then
	Counter = 0
	If Counter2 = 250000 Then
	    Counter2 = 0
	    If (GetAsyncKeyState(VK_ESCAPE) And &H8001) <> 0 Then
		Err.Raise vbObjectError + 513, "BigDecimal", "Operation aborted by user."
	    End If
	End If
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

Public Property Let MaxCount(ByVal Val As Long)
    Dim Factors As Variant
    Dim i As Long

    Factors = Array(100, 125, 200, 250, 400, 500, 625, 1000, 1250, 2000, 2500, 3125, 5000, 6250, 10000, 12500, 15625, 25000, 31250, 50000, 62500, 125000, 250000)

    If Val <= 100 Then
	MaxCounter = 100
	Exit Property
    ElseIf Val >= 250000 Then
	MaxCounter = 250000
	Exit Property
    End If

    For i = LBound(Factors) To UBound(Factors) - 1
	If Val >= Factors(i) And Val <= Factors(i + 1) Then
	    If (Val - Factors(i)) < (Factors(i + 1) - Val) Then   'find closest factor
		MaxCounter = Factors(i)
	    Else
		MaxCounter = Factors(i + 1)
	    End If

	    Exit For
	End If
    Next i
End Property
