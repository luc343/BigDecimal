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

Public Function CBgDec(Value As Variant) As BigDecimal
    Set CBgDec = New BigDecimal

    CBgDec = Value
End Function
