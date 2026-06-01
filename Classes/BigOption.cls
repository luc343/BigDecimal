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

'---------------------------------------------------
'
'                  Class BigOption
'
' Copyright (c) Lucien Cinc 2025-26
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
'---------------------------------------------------

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private List As Scripting.Dictionary

Private Big_lTarScale As Long               'max precision for a BigDecimal

Private Ln2_Cache_Buf As String             'cache for ln(2)
Private Ln2_Cache_Scale As Long             'scale for cache of ln(2)

Private Ln10_Cache_Buf As String            'cache for ln(10)
Private Ln10_Cache_Scale As Long            'scale for cache of ln(10)

Private OptionOutputDebug As Boolean        'switch debug on/off
Private OptionAutoRound As Boolean          'switch rounding on/off when return a string
Private OptionAddTrailingZeroes As Boolean  'switch trailing zeroes on/off when using round()

'######################### BigOption Methods #########################

Private Sub Class_Initialize()
    Set List = New Scripting.Dictionary

    Big_lTarScale = 100
    OptionOutputDebug = False
    OptionAutoRound = True
    OptionAddTrailingZeroes = False
End Sub

Private Sub Class_Terminate()
    Set List = Nothing
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

'Count
Public Property Get Count() As Long
    Count = List.Count
End Property

'MaxScale
Public Property Get TarScale() As Long
    TarScale = Big_lTarScale
End Property

Public Property Let TarScale(Val As Long)
    Dim Temp As BigDecimal
    Dim Ptr As LongPtr
    Dim Key As Variant

    Big_lTarScale = Val

    For Each Key In List.Keys
	Ptr = CLngPtr(Key)
	Set Temp = CastPtr(Ptr)             'inc reference count
	Temp.TarScale = Big_lTarScale
	Set Temp = Nothing                  'dec reference count
    Next Key
End Property

'AddTrailingZeroes
Public Property Get AddTrailingZeroes() As Boolean
    AddTrailingZeroes = OptionAddTrailingZeroes
End Property

Public Property Let AddTrailingZeroes(Flag As Boolean)
    Dim Temp As BigDecimal
    Dim Ptr As LongPtr
    Dim Key As Variant

    OptionAddTrailingZeroes = Flag

    For Each Key In List.Keys
	Ptr = CLngPtr(Key)
	Set Temp = CastPtr(Ptr)             'inc reference count
	Temp.AddTrailingZeroes = OptionAddTrailingZeroes
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
	Ptr = CLngPtr(Key)
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
	Ptr = CLngPtr(Key)
	Set Temp = CastPtr(Ptr)             'inc reference count
	Temp.OutputDebug = OptionOutputDebug
	Set Temp = Nothing                  'dec reference count
    Next Key
End Property

'Ln2_Buf
Friend Property Get Ln2_Buf() As String
    Ln2_Buf = Ln2_Cache_Buf
End Property

Friend Property Let Ln2_Buf(Str As String)
    Ln2_Cache_Buf = Str
End Property

'Ln2_Scale
Friend Property Get Ln2_Scale() As Long
    Ln2_Scale = Ln2_Cache_Scale
End Property

Friend Property Let Ln2_Scale(Val As Long)
    Ln2_Cache_Scale = Val
End Property

'Ln10_Buf
Friend Property Get Ln10_Buf() As String
    Ln10_Buf = Ln10_Cache_Buf
End Property

Friend Property Let Ln10_Buf(Str As String)
    Ln10_Cache_Buf = Str
End Property

'Ln10_Scale
Friend Property Get Ln10_Scale() As Long
    Ln10_Scale = Ln10_Cache_Scale
End Property

Friend Property Let Ln10_Scale(Val As Long)
    Ln10_Cache_Scale = Val
End Property
