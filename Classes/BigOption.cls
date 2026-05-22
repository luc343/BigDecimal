VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BigOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'-------------------------------------------------------------------------------
'
'                             Class BigDecimal
'
' Copyright (c) Lucien Cinc 2025
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
'-------------------------------------------------------------------------------

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private List As New Scripting.Dictionary

Private Big_lMaxScale As Long           'max precision for a BigDecimal
Private OptionOutputDebug As Boolean    'switch debug on/off
Private OptionAutoRound As Boolean      'switch rounding on/off

'######################### BigOption Methods #########################

Private Sub Class_Initialize()
    Big_lMaxScale = 100
    OptionOutputDebug = False
    OptionAutoRound = True
End Sub

'***************************************
'
'         Register a BigDecimal
'
'***************************************

Public Sub Register(ByRef Val As BigDecimal)
    Dim Ptr As LongPtr

    Ptr = ObjPtr(Val)
    List(CStr(Ptr)) = 0
End Sub

'***************************************
'
'        Unregister a BigDecimal
'
'***************************************

Public Sub Unregister(ByRef Val As BigDecimal)
    Dim Ptr As LongPtr

    Ptr = ObjPtr(Val)
    List.Remove CStr(Ptr)   'throw exception on error
End Sub

Private Function CastPtr(ByVal Ptr As LongPtr) As BigDecimal
    Dim Temp As BigDecimal
    Dim Zero As LongPtr

    CopyMemory Temp, Ptr, LenB(Ptr)     'copy the ptr to Temp

    Set CastPtr = Temp

    CopyMemory Temp, Zero, LenB(Zero)    'un-point Temp so VBA doesn't try to free the memory
End Function

'####################### BigOption Properties ########################

'MaxScale
Public Property Get MaxScale() As Long
    MaxScale = Big_lMaxScale
End Property

Public Property Let MaxScale(Val As Long)
    Dim Temp As BigDecimal
    Dim Ptr As LongPtr
    Dim Key As Variant

    Big_lMaxScale = Val

    For Each Key In List.Keys
	Ptr = CLngPtr(Key)
	Set Temp = CastPtr(Ptr)             'inc reference count
	Temp.MaxScale = Big_lMaxScale
	Set Temp = Nothing                  'dec reference count
    Next Key
End Property

'AutoRound
Public Property Get AutoRound() As Boolean
    AutoRound = OptionAutoRound
End Property

Public Property Let AutoRound(Flag As Boolean)
    Dim Temp As BigDecimal
    Dim Ptr As LongPtr
    Dim Key As Variant

    OptionAutoRound = Flag

    For Each Key In List.Keys
	Ptr = List(Key)
	Set Temp = CastPtr(Ptr)             'inc reference count
	Temp.AutoRound = OptionAutoRound
	Set Temp = Nothing                  'dec reference count
    Next Key
End Property

'OutputDebug
Public Property Get OutputDebug() As Boolean
    OutputDebug = OptionOutputDebug
End Property

Public Property Let OutputDebug(Flag As Boolean)
    Dim Temp As BigDecimal
    Dim Ptr As LongPtr
    Dim Key As Variant

    OptionOutputDebug = Flag

    For Each Key In List.Keys
	Ptr = List(Key)
	Set Temp = CastPtr(Ptr)             'inc reference count
	Temp.OutputDebug = OptionOutputDebug
	Set Temp = Nothing                  'dec reference count
    Next Key
End Property
