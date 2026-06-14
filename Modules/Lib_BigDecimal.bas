Attribute VB_Name = "Lib_BigDecimal"
Option Explicit

'---------------------------------------------------
'
'                 Lib_BigDecimal
'
' Copyright (c) Lucien Cinc 2025-26
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
'---------------------------------------------------

Public BigOption As New BigOption   'set global options
Public Yield As New Yield           'prevent excel from not responding

'***************************************
'
'           New a BigDecimal
'
'***************************************

Public Function New_BigDecimal(Optional Value As Variant) As BigDecimal
    Set New_BigDecimal = New BigDecimal

    If Not IsMissing(Value) Then
	New_BigDecimal = Value
    End If
End Function

'***************************************
'
'       Cast types to BigDecimal
'
'***************************************

Public Function CBgDec(Value As Variant) As BigDecimal
    Set CBgDec = New BigDecimal

    CBgDec = Value
End Function

'############################ Properties #############################

'BigZero
Public Property Get BigZero() As BigDecimal
    Static Num As BigDecimal

    If Num Is Nothing Then
	Set Num = New BigDecimal
	Num.Unregister  'from BigOption
    End If

    Set BigZero = Num
End Property

'BigOne
Public Property Get BigOne() As BigDecimal
    Static Num As BigDecimal

    If Num Is Nothing Then
	Set Num = New BigDecimal
	Num.Unregister  'from BigOption
	Num.sValue = "1"
    End If

    Set BigOne = Num
End Property

'BigTwo
Public Property Get BigTwo() As BigDecimal
    Static Num As BigDecimal

    If Num Is Nothing Then
	Set Num = New BigDecimal
	Num.Unregister  'from BigOption
	Num.sValue = "2"
    End If

    Set BigTwo = Num
End Property

'BigPi
Public Property Get BigPi() As BigDecimal
    Set BigPi = New BigDecimal

    BigPi = BigPi.Pi
End Property
