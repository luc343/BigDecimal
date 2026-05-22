VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BigDecimalForm
   Caption         =   "BigCalculator"
   ClientHeight    =   5190
   ClientLeft      =   -1065
   ClientTop       =   -4875
   ClientWidth     =   8175
   OleObjectBlob   =   "BigDecimalForm.frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BigDecimalForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------
'
'                BigDecimalForm
'
' Copyright (c) Lucien Cinc 2025
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
'---------------------------------------------------

Private Sub OK_Click()
    Unload BigDecimalForm
End Sub

Private Sub Add_Click()
    Dim Num1 As BigDecimal
    Dim Num2 As BigDecimal

    With BigDecimalForm
	If .Num1.value = "" Then
	    ErrBox "Number 1 is empty!", "Number"
	    Exit Sub
	ElseIf .Num2.value = "" Then
	    ErrBox "Number 2 is empty!", "Number"
	    Exit Sub
	End If

	Set Num1 = New_BigDecimal(.Num1.value)
	Set Num2 = New_BigDecimal(.Num2.value)

	If Num1 Is Nothing Then
	    ErrBox "Number 1 is nothing!", "Number"
	    Exit Sub
	ElseIf Num2 Is Nothing Then
	    ErrBox "Number 2 is nothing!", "Number"
	    Exit Sub
	End If

	.Answer.value = Num1.plus(Num2).valueEx(vbString)
    End With
End Sub

Private Sub Subtract_Click()
    Dim Num1 As BigDecimal
    Dim Num2 As BigDecimal

    With BigDecimalForm
	If .Num1.value = "" Then
	    ErrBox "Number 1 is empty!", "Number"
	    Exit Sub
	ElseIf .Num2.value = "" Then
	    ErrBox "Number 2 is empty!", "Number"
	    Exit Sub
	End If

	Set Num1 = New_BigDecimal(.Num1.value)
	Set Num2 = New_BigDecimal(.Num2.value)

	If Num1 Is Nothing Then
	    ErrBox "Number 1 is nothing!", "Number"
	    Exit Sub
	ElseIf Num2 Is Nothing Then
	    ErrBox "Number 2 is nothing!", "Number"
	    Exit Sub
	End If

	.Answer.value = Num1.minus(Num2).valueEx(vbString)
    End With
End Sub

Private Sub Multiply_Click()
    Dim Num1 As BigDecimal
    Dim Num2 As BigDecimal

    With BigDecimalForm
	If .Num1.value = "" Then
	    ErrBox "Number 1 is empty!", "Number"
	    Exit Sub
	ElseIf .Num2.value = "" Then
	    ErrBox "Number 2 is empty!", "Number"
	    Exit Sub
	End If

	Set Num1 = New_BigDecimal(.Num1.value)
	Set Num2 = New_BigDecimal(.Num2.value)

	If Num1 Is Nothing Then
	    ErrBox "Number 1 is nothing!", "Number"
	    Exit Sub
	ElseIf Num2 Is Nothing Then
	    ErrBox "Number 2 is nothing!", "Number"
	    Exit Sub
	End If

	.Answer.value = Num1.times(Num2).valueEx(vbString)
    End With
End Sub

Private Sub Divide_Click()
    Dim Num1 As BigDecimal
    Dim Num2 As BigDecimal

    With BigDecimalForm
	If .Num1.value = "" Then
	    ErrBox "Number 1 is empty!", "Number"
	    Exit Sub
	ElseIf .Num2.value = "" Then
	    ErrBox "Number 2 is empty!", "Number"
	    Exit Sub
	End If

	Set Num1 = New_BigDecimal(.Num1.value)
	Set Num2 = New_BigDecimal(.Num2.value)

	If Num1 Is Nothing Then
	    ErrBox "Number 1 is nothing!", "Number"
	    Exit Sub
	ElseIf Num2 Is Nothing Then
	    ErrBox "Number 2 is nothing!", "Number"
	    Exit Sub
	End If

	.Answer.value = Num1.Divide(Num2).valueEx(vbString)
    End With
End Sub




Private Sub Power_Click()
    Dim Num1 As BigDecimal
    Dim Num2 As BigDecimal

    With BigDecimalForm
	If .Num1.value = "" Then
	    ErrBox "Number 1 is empty!", "Number"
	    Exit Sub
	ElseIf .Num2.value = "" Then
	    ErrBox "Number 2 is empty!", "Number"
	    Exit Sub
	End If

	Set Num1 = New_BigDecimal(.Num1.value)
	Set Num2 = New_BigDecimal(.Num2.value)

	If Num1 Is Nothing Then
	    ErrBox "Number 1 is nothing!", "Number"
	    Exit Sub
	ElseIf Num2 Is Nothing Then
	    ErrBox "Number 2 is nothing!", "Number"
	    Exit Sub
	End If

	.Answer.value = Num1.Pow(Num2).valueEx(vbString)
    End With
End Sub







Private Sub Modulus_Click()
    Dim Num1 As BigDecimal
    Dim Num2 As BigDecimal

    With BigDecimalForm
	If .Num1.value = "" Then
	    ErrBox "Number 1 is empty!", "Number"
	    Exit Sub
	ElseIf .Num2.value = "" Then
	    ErrBox "Number 2 is empty!", "Number"
	    Exit Sub
	End If

	Set Num1 = New_BigDecimal(.Num1.value)
	Set Num2 = New_BigDecimal(.Num2.value)

	If Num1 Is Nothing Then
	    ErrBox "Number 1 is nothing!", "Number"
	    Exit Sub
	ElseIf Num2 Is Nothing Then
	    ErrBox "Number 2 is nothing!", "Number"
	    Exit Sub
	End If

	.Answer.value = Num1.remdr(Num2).valueEx(vbString)
    End With
End Sub

Private Sub Compare_Click()
    Dim Num1 As BigDecimal
    Dim Num2 As BigDecimal

    With BigDecimalForm
	If .Num1.value = "" Then
	    ErrBox "Number 1 is empty!", "Number"
	    Exit Sub
	ElseIf .Num2.value = "" Then
	    ErrBox "Number 2 is empty!", "Number"
	    Exit Sub
	End If

	Set Num1 = New_BigDecimal(.Num1.value)
	Set Num2 = New_BigDecimal(.Num2.value)

	If Num1 Is Nothing Then
	    ErrBox "Number 1 is nothing!", "Number"
	    Exit Sub
	ElseIf Num2 Is Nothing Then
	    ErrBox "Number 2 is nothing!", "Number"
	    Exit Sub
	End If

	.Answer.value = Num1.Cmp(Num2)
    End With
End Sub

Private Sub Clipboard1_Click()
    CopyText BigDecimalForm.Num1.value
End Sub

Private Sub Clipboard2_Click()
    CopyText BigDecimalForm.Num2.value
End Sub

Private Sub RoundAns_Click()
    Dim Num1 As BigDecimal

    With BigDecimalForm
	If .RoundBy.value = "" Then
	    ErrBox "Round By is empty!", "Number"
	    Exit Sub
	ElseIf .Answer.value = "" Then
	    ErrBox "Answer is empty!", "Number"
	    Exit Sub
	End If

	Set Num1 = New_BigDecimal(.Answer.value)

	If Num1 Is Nothing Then
	    ErrBox "Number 1 is nothing!", "Number"
	    Exit Sub
	End If

	.Answer.value = Num1.Round(.RoundBy.value).valueEx(vbString)
    End With
End Sub

Private Sub WholeAns_Click()
    Dim Num1 As BigDecimal

    With BigDecimalForm
	If .Answer.value = "" Then
	    ErrBox "Answer is empty!", "Number"
	    Exit Sub
	End If

	Set Num1 = New_BigDecimal(.Answer.value)

	If Num1 Is Nothing Then
	    ErrBox "Number 1 is nothing!", "Number"
	    Exit Sub
	End If

	.Answer.value = Num1.trunc().valueEx(vbString)
    End With
End Sub

Private Sub AbsoluteAns_Click()
    Dim Num1 As BigDecimal

    With BigDecimalForm
	If .Answer.value = "" Then
	    ErrBox "Answer is empty!", "Number"
	    Exit Sub
	End If

	Set Num1 = New_BigDecimal(.Answer.value)

	If Num1 Is Nothing Then
	    ErrBox "Number 1 is nothing!", "Number"
	    Exit Sub
	End If

	.Answer.value = Num1.AbsVal().valueEx(vbString)
    End With

End Sub

Private Sub NegativeAns_Click()
    Dim Num1 As BigDecimal

    With BigDecimalForm
	If .Answer.value = "" Then
	    ErrBox "Answer is empty!", "Number"
	    Exit Sub
	End If

	Set Num1 = New_BigDecimal(.Answer.value)

	If Num1 Is Nothing Then
	    ErrBox "Number 1 is nothing!", "Number"
	    Exit Sub
	End If

	.Answer.value = Num1.Neg().valueEx(vbString)
    End With

End Sub

Private Sub PiConst_Click()
    Dim Num1 As BigDecimal

    With BigDecimalForm
	Set Num1 = New_BigDecimal()

	If Num1 Is Nothing Then
	    ErrBox "Number 1 is nothing!", "Number"
	    Exit Sub
	End If

	Num1.value = Num1.Pi
	.Answer.value = Num1.valueEx(vbString)
    End With
End Sub



Private Sub EConst_Click()
    Dim Num1 As BigDecimal

    With BigDecimalForm
	Set Num1 = New_BigDecimal()

	If Num1 Is Nothing Then
	    ErrBox "Number 1 is nothing!", "Number"
	    Exit Sub
	End If

	Num1 = 1
	Num1 = Num1.Exp()
	.Answer.value = Num1.valueEx(vbString)
    End With
End Sub

Private Sub FractionAns_Click()
    Dim Num1 As BigDecimal

    With BigDecimalForm
	If .Answer.value = "" Then
	    ErrBox "Round By is empty!", "Number"
	    Exit Sub
	End If

	Set Num1 = New_BigDecimal(.Answer.value)

	If Num1 Is Nothing Then
	    ErrBox "Number 1 is nothing!", "Number"
	    Exit Sub
	End If

	.Answer.value = Num1.Frac().valueEx(vbString)
    End With

End Sub

Private Sub SquareRoot_Click()
    Dim Num1 As BigDecimal

    With BigDecimalForm
	If .Num1.value = "" Then
	    ErrBox "Number 1 is empty!", "Number"
	    Exit Sub
	End If

	Set Num1 = New_BigDecimal(.Num1.value)

	If Num1 Is Nothing Then
	    ErrBox "Number 1 is nothing!", "Number"
	    Exit Sub
	End If

	.Answer.value = Num1.Sqrt().valueEx(vbString)
    End With

End Sub

Private Sub Logarithm_Click()
    Dim Num1 As BigDecimal

    With BigDecimalForm
	If .Num1.value = "" Then
	    ErrBox "Number 1 is empty!", "Number"
	    Exit Sub
	End If

	Set Num1 = New_BigDecimal(.Num1.value)

	If Num1 Is Nothing Then
	    ErrBox "Number 1 is nothing!", "Number"
	    Exit Sub
	End If

	.Answer.value = Num1.Ln().valueEx(vbString)
    End With

End Sub

Private Sub Exponential_Click()
    Dim Num1 As BigDecimal

    With BigDecimalForm
	If .Num1.value = "" Then
	    ErrBox "Number 1 is empty!", "Number"
	    Exit Sub
	End If

	Set Num1 = New_BigDecimal(.Num1.value)

	If Num1 Is Nothing Then
	    ErrBox "Number 1 is nothing!", "Number"
	    Exit Sub
	End If

	.Answer.value = Num1.Exp().valueEx(vbString)
    End With

End Sub

Private Sub Factorial_Click()
    Dim Num1 As BigDecimal

    With BigDecimalForm
	If .Num1.value = "" Then
	    ErrBox "Number 1 is empty!", "Number"
	    Exit Sub
	End If

	Set Num1 = New_BigDecimal(.Num1.value)

	If Num1 Is Nothing Then
	    ErrBox "Number 1 is nothing!", "Number"
	    Exit Sub
	End If

	.Answer.value = Num1.Fact().valueEx(vbString)
    End With

End Sub

Private Sub ToNum1_Click()
    With BigDecimalForm
	If .Answer.value = "" Then
	    ErrBox "Answer is empty!", "Number"
	    Exit Sub
	End If

	.Num1.value = Answer.value
    End With
End Sub

Private Sub ToNum2_Click()
    With BigDecimalForm
	If .Answer.value = "" Then
	    ErrBox "Answer is empty!", "Number"
	    Exit Sub
	End If

	.Num2.value = Answer.value
    End With
End Sub

Private Sub Clear_Click()
    Dim Num1 As BigDecimal

    With BigDecimalForm
	.Num1.value = ""
	.Num2.value = ""
	Answer.value = ""
    End With
End Sub

Private Sub UserForm_Initialize()
    Dim hwnd As LongPtr, Row As Long

    With BigDecimalForm
	.Height = 246.05    'fix screen driver issue
	.width = 351.05

	.StartUpPosition = 0    'allow manual positioning

	hwnd = FindWindow("ThunderDFrame", .Caption)
	CentreUserForm hwnd
    End With
End Sub
