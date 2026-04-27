Attribute VB_Name = "Lib_BigDecimal"
Option Explicit

'---------------------------------------------------
'
'                  Lib_BigDecimal
'
' Copyright (c) Lucien Cinc 2025
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
'---------------------------------------------------

'***************************************
'
'           New a BigDecimal
'
'***************************************

Public Function New_BigDecimal(Optional StrValue As String = "") As BigDecimal
	Set New_BigDecimal = New BigDecimal

	On Error GoTo Done
	If StrValue <> "" Then
		New_BigDecimal.StrValue = StrValue
	End If
	On Error GoTo 0

	Exit Function
Done:
	On Error GoTo 0

	Set New_BigDecimal = Nothing
End Function
