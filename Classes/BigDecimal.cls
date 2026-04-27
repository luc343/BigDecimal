VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BigDecimal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'-------------------------------------------------------------------------------
'
'                Class BigDecimal
'
' Copyright (c) Lucien Cinc 2025
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
' Based on the work from the site:
'   https://github.com/Sam-bit/BigDecimal-CPP/blob/master/BigDecimal.cpp#L70
'
'-------------------------------------------------------------------------------

Private Big_sValue As String
Private Big_lSign As Long
Private Big_lInt As Long
Private Big_lDot As Long
Private Big_lFrac As Long
Private Big_lScale As Long

Private Big_lMaxScale As Long       'internal decimal places

Private LogOf2 As String
Private LogOf2Scale As Long

Private Const DebugFlag As Boolean = False  'switch debug on/off
Private Const RoundFlag As Boolean = True   'switch rounding on/off

'######################### BigDecimal Methods ########################

Private Sub Class_Initialize()
	Big_sValue = "0"
	Big_lSign = 1
	Big_lInt = 1
	Big_lDot = 2
	Big_lFrac = 2
	Big_lScale = 0

	Big_lMaxScale = 100
End Sub

'***************************************
'
'                  Add
'
'***************************************

Public Function Add(other As BigDecimal) As BigDecimal
	Dim temp As New BigDecimal

	Set Add = New BigDecimal

	If Big_lSign > 0 And other.lSign > 0 Then
		Add.Value = add_positive(other, Big_lMaxScale, 1)
		Exit Function
	End If

	If Big_lSign > 0 And other.lSign < 0 Then
		If compare_positive(other) >= 0 Then
			Add.Value = subtract_positive(other, Big_lMaxScale, 1)
			Exit Function
		Else
			temp.Value = Me

			Me.Value = other
			Add.Value = subtract_positive(temp, Big_lMaxScale, -1)
			Me.Value = temp
			Exit Function
		End If
	End If

	If Big_lSign < 0 And other.lSign > 0 Then
		If compare_positive(other) <= 0 Then
			temp.Value = Me

			Me.Value = other
			Add.Value = subtract_positive(temp, Big_lMaxScale, 1)
			Me.Value = temp
			Exit Function
		Else
			Add.Value = subtract_positive(other, Big_lMaxScale, -1)
			Exit Function
		End If
	End If

	Add.Value = add_positive(other, Big_lMaxScale, -1)     'Big_lSign < 0 And other.lSign < 0
End Function

'***************************************
'
'                Subtract
'
'***************************************

Public Function Subtract(other As BigDecimal) As BigDecimal
	Dim org_lsign As Long

	Set Subtract = New BigDecimal

	org_lsign = other.lSign

	other.lSign = other.lSign * -1
	Subtract.Value = Add(other)

	other.lSign = org_lsign
End Function

'***************************************
'
'                Multiply
'
'***************************************

Public Function Multiply(other As BigDecimal) As BigDecimal
	Set Multiply = New BigDecimal

	Multiply.Value = multiply_positive(other, Big_lMaxScale, Big_lSign * other.lSign)
End Function

'***************************************
'
'                Divide
'
'***************************************

Public Function Divide(other As BigDecimal) As BigDecimal
	Set Divide = New BigDecimal

	Divide.Value = divide_positive(other, Big_lMaxScale, Big_lSign * other.lSign)
End Function

'***************************************
'
'                Modulus
'
'***************************************

Public Function Modulus(other As BigDecimal) As BigDecimal
	Dim q As New BigDecimal
	Dim result As New BigDecimal

	Set Modulus = New BigDecimal

	q.Value = Absolute().Divide(other.Absolute()).Whole()          'Quotient = Int(Abs(x) / Abs(y))
	result.Value = Absolute().Subtract(other.Absolute().Multiply(q))    'Remainder = Abs(dblX) - Abs(dblY) * dblQuotient

	If Big_lSign * other.lSign < 0 Then
		result.Value = other.Absolute().Subtract(result)
	End If

	result.lSign = other.lSign

	Modulus.sValue = result.sValue
	Modulus.lSign = result.lSign
	Modulus.lInt = result.lInt
	Modulus.lDot = result.lDot
	Modulus.lFrac = result.lFrac
	Modulus.lScale = result.lScale

	DebugPrint Modulus, "Modulus"
End Function

'***************************************
'
'              SquareRoot
'
'***************************************

Public Function SquareRoot() As BigDecimal
	Dim result As New BigDecimal
	Dim Y As New BigDecimal
	Dim prec As New BigDecimal
	Dim Str As String

	Set SquareRoot = New BigDecimal

	If Big_lSign < 0 Then
		Err.Raise vbObjectError + 1100, "BigDecimal", "SquareRoot of negative number is not permtted."
		Exit Function
	End If

	Str = "0." & String(Big_lMaxScale - 1, "0") & "1"
	prec.StrValue = Str

	result.Value = Divide(Two)

	While result.Subtract(Y).Absolute().Compare(prec) = 1
		Y.Value = result
		result.Value = result.Add(Divide(result)).Divide(Two)       'x = (x + value / x) / 2

		DoEvents
	Wend

	SquareRoot.sValue = result.sValue
	SquareRoot.lSign = result.lSign
	SquareRoot.lInt = result.lInt
	SquareRoot.lDot = result.lDot
	SquareRoot.lFrac = result.lFrac
	SquareRoot.lScale = result.lScale

	If RoundFlag Then SquareRoot.Value = SquareRoot.Round(Big_lMaxScale)

	DebugPrint SquareRoot, "SquareRoot"
End Function

'***************************************
'
'           Natural Logarithm
'
'***************************************

Public Function Logarithm() As BigDecimal
	Dim result As New BigDecimal
	Dim X As New BigDecimal
	Dim n As New BigDecimal
	Dim Y As New BigDecimal
	Dim term As New BigDecimal
	Dim prec As New BigDecimal
	Dim Str As String

	Set Logarithm = New BigDecimal

	Str = "0." & String(Big_lMaxScale - 1, "0") & "1"
	prec.StrValue = Str

	If Compare(Zero) <= 0 Then   'log is undefined for other <= 0
		Err.Raise vbObjectError + 1101, "BigDecimal", "Logarithm of ZERO or negative number is not permtted."
		Exit Function
	End If

	X.Value = Me
	While X.Compare(Two) >= 0
		n.Value = n.Add(One)
		X.Value = X.Divide(Two)

		DoEvents
	Wend
	While X.Compare(One) < 0
		n.Value = n.Subtract(One)
		X.Value = X.Multiply(Two)

		DoEvents
	Wend
	X.Value = X.Subtract(One)

	term.Value = X
	Y.Value = Two
	While term.Absolute().Compare(prec) = 1
		result.Value = result.Add(term)
		term.Value = term.Negative().Multiply(X).Multiply(Y.Subtract(One)).Divide(Y)       '-term * x * (y - 1) / y
		Y.Value = Y.Add(One)

		DoEvents
	Wend

	result.Value = result.Add(n.Multiply(LogarithmOfTwo()))

	Logarithm.sValue = result.sValue
	Logarithm.lSign = result.lSign
	Logarithm.lInt = result.lInt
	Logarithm.lDot = result.lDot
	Logarithm.lFrac = result.lFrac
	Logarithm.lScale = result.lScale

	If RoundFlag Then Logarithm.Value = Logarithm.Round(Big_lMaxScale)

	DebugPrint Logarithm, "Logarithm"
End Function

'***************************************
'
'            Exponential
'
'***************************************

Public Function Exponential() As BigDecimal
	Dim result As New BigDecimal
	Dim term As New BigDecimal
	Dim prec As New BigDecimal
	Dim n As New BigDecimal
	Dim Str As String

	Set Exponential = New BigDecimal

	Str = "0." & String(Big_lMaxScale - 1, "0") & "1"
	prec.StrValue = Str

	result.Value = One
	term.Value = One
	n.Value = One

	While term.Absolute().Compare(prec) = 1
		term.Value = term.Multiply(Me).Divide(n)
		result.Value = result.Add(term)
		n.Value = n.Add(One)

		DoEvents
	Wend

	Exponential.sValue = result.sValue
	Exponential.lSign = result.lSign
	Exponential.lInt = result.lInt
	Exponential.lDot = result.lDot
	Exponential.lFrac = result.lFrac
	Exponential.lScale = result.lScale

	If RoundFlag Then Exponential.Value = Exponential.Round(Big_lMaxScale)

	DebugPrint Exponential, "Exponential"
End Function

'***************************************
'
'               Factorial
'
'***************************************

Public Function Factorial() As BigDecimal
	Dim result As New BigDecimal
	Dim i As New BigDecimal

	Set Factorial = New BigDecimal

	If Big_lScale <> 0 Then
		Err.Raise vbObjectError + 1102, "BigDecimal", "Factorial of a fraction is not permtted."
		Exit Function
	End If

	result.Value = One
	i.Value = Two

	While i.Compare(Me) < 1
		result.Value = result.Multiply(i)
		i.Value = i.Add(One)

		DoEvents
	Wend

	Factorial.sValue = result.sValue
	Factorial.lSign = result.lSign
	Factorial.lInt = result.lInt
	Factorial.lDot = result.lDot
	Factorial.lFrac = result.lFrac
	Factorial.lScale = result.lScale

	If RoundFlag Then Factorial.Value = Factorial.Round(Big_lMaxScale)

	DebugPrint Factorial, "Factorial"
End Function

'***************************************
'
'               Absolute
'
'***************************************

Public Function Absolute() As BigDecimal
	Dim result As String
	Dim resint As Long, resdot As Long, resfrac As Long

	Set Absolute = New BigDecimal

	result = Big_sValue
	resint = Big_lInt
	resdot = Big_lDot
	resfrac = Big_lFrac

	If Big_lSign < 0 Then
		result = Mid(result, 2)
		resint = resint - 1
		resdot = resdot - 1
		resfrac = resfrac - 1
	End If

	Absolute.sValue = result
	Absolute.lSign = 1
	Absolute.lInt = resint
	Absolute.lDot = resdot
	Absolute.lFrac = resfrac
	Absolute.lScale = Big_lScale

	DebugPrint Absolute, "Absolute"
End Function

'***************************************
'
'              Negative
'
'***************************************

Public Function Negative() As BigDecimal
	Dim result As String
	Dim ressign As Long, resint As Long, resdot As Long, resfrac As Long
	Dim comp As Long

	Set Negative = New BigDecimal

	result = Big_sValue
	ressign = Big_lSign
	resint = Big_lInt
	resdot = Big_lDot
	resfrac = Big_lFrac

	comp = Compare(Zero)
	If comp = -1 Then
		result = Mid(result, 2)
		ressign = 1
		resint = resint - 1
		resdot = resdot - 1
		resfrac = resfrac - 1
	ElseIf comp = 1 Then
		result = "-" & result
		ressign = -1
		resint = resint + 1
		resdot = resdot + 1
		resfrac = resfrac + 1
	End If

	Negative.sValue = result
	Negative.lSign = ressign
	Negative.lInt = resint
	Negative.lDot = resdot
	Negative.lFrac = resfrac
	Negative.lScale = Big_lScale

	DebugPrint Negative, "Negative"
End Function

'***************************************
'
'               Whole
'
'***************************************

Public Function Whole() As BigDecimal
	Dim result As String
	Dim resfrac As Long, resscale As Long

	Set Whole = New BigDecimal

	result = Big_sValue
	resfrac = Big_lFrac
	resscale = Big_lScale

	If Big_lScale > 0 Then
		result = Mid(result, 1, Big_lDot - 1)
		resfrac = Big_lDot
		resscale = 0
	End If

	Whole.sValue = result
	Whole.lSign = Big_lSign
	Whole.lInt = Big_lInt
	Whole.lDot = Big_lDot
	Whole.lFrac = resfrac
	Whole.lScale = resscale

	DebugPrint Whole, "Whole"
End Function

'***************************************
'
'              Fraction
'
'***************************************

Public Function Fraction() As BigDecimal
	Dim result As String
	Dim resint As Long, resdot As Long

	Set Fraction = New BigDecimal

	result = Big_sValue
	resint = Big_lInt
	resdot = Big_lDot

	While resint < resdot
		Mid(result, resint, 1) = "0"
		resint = resint + 1
	Wend
	resint = resdot - 1

	Fraction.sValue = result
	Fraction.lSign = Big_lSign
	Fraction.lInt = resint
	Fraction.lDot = resdot
	Fraction.lFrac = Big_lFrac
	Fraction.lScale = Big_lScale

	DebugPrint Fraction, "Fraction"
End Function

'***************************************
'
'                Round
'
'***************************************

Public Function Round(ByVal result_scale As Long, Optional ByVal add_trailing_zeroes As Boolean = False, Optional ByVal round_last As Boolean = True) As BigDecimal
	Dim i As Long, kI As Long
	Dim ressign As Long, resint As Long, resdot As Long, resfrac As Long, resscale As Long
	Dim result As String

	result = Big_sValue
	ressign = Big_lSign
	resint = Big_lInt
	resdot = Big_lDot
	resfrac = Big_lFrac
	resscale = Big_lScale

	Set Round = New BigDecimal

	While Mid(result, resint, 1) = "0" And resint + 1 < resdot
		resint = resint + 1
	Wend

	Debug.Assert resint > 0 And resscale >= 0 And result_scale >= 0

	If ressign < 0 Then
		result = Mid(result, 2)
		resint = resint - 1
		resdot = resdot - 1
		resfrac = resfrac - 1
	End If

	If round_last Then
		If resscale > result_scale Then
			While result_scale > 0 And Mid(result, resfrac + result_scale - 1, 1) = "9" And Mid(result, resfrac + result_scale, 1) >= "5"
				result_scale = result_scale - 1
			Wend
			resscale = result_scale
			If Mid(result, resfrac + result_scale, 1) >= "5" Then
				If result_scale > 0 Then
					result = Left(result, resfrac + result_scale - 1 - 1) & Chr(asc(Mid(result, resfrac + result_scale - 1, 1)) + 1) & Right(result, Len(result) - resfrac - result_scale + 1)
				Else
					resfrac = resfrac - 1
					Debug.Assert resfrac = resdot

					For i = 0 To Len(result) - resdot - 1
						If resdot - i - 1 = 0 Then
							result = "0" & result
							resdot = resdot + 1
							resfrac = resfrac + 1
						End If

						If Mid(result, resdot - i - 1, 1) = "9" Then
							Mid(result, resdot - i - 1, 1) = "0"
						Else

							Mid(result, resdot - i - 1, 1) = Chr(asc(Mid(result, resdot - i - 1, 1)) + 1)
							Exit For
						End If
					Next i

					If resdot - i - 1 < resint Then
						resint = resdot - i - 1
					End If
				End If
			End If
		End If
	Else
		If resscale > result_scale Then
			resscale = result_scale
		End If
	End If

	If resscale = 0 And resfrac > resdot Then
		resfrac = resdot
	End If

	If resscale = result_scale Or Not add_trailing_zeroes Then
		result = Left(result, resdot) & Mid(result, resfrac, resscale)
	Else
		result = Left(result, resdot) & Mid(result, resfrac, resscale)

		If resscale = 0 Then
			result = result & "."
		End If
		For kI = 1 To result_scale - resscale
			result = result & "0"
		Next kI
		resscale = result_scale
		resfrac = resdot + 1
	End If

	If Right(result, 1) = "." Then
		result = Left(result, Len(result) - 1)
	End If

	If ressign < 0 Then
		result = "-" & result
		resint = resint + 1
		resdot = resdot + 1
		resfrac = resfrac + 1

	End If

	Round.sValue = result
	Round.lSign = ressign
	Round.lInt = resint
	Round.lDot = resdot
	Round.lFrac = resfrac
	Round.lScale = resscale

	DebugPrint Round, "Round"
End Function

'***************************************
'
'                Trim
'
'***************************************

Public Function Trim() As BigDecimal
	Dim resscale As Long, resint As Long, resdot As Long, resfrac As Long, ressign As Long, cur_pos As Long
	Dim result As String

	Set Trim = New BigDecimal

	result = Big_sValue

	ressign = Big_lSign
	resint = Big_lInt
	resdot = Big_lDot
	resfrac = Big_lFrac
	resscale = Big_lScale

	cur_pos = 1
	If ressign < 0 Then
		cur_pos = cur_pos + 1

		While resint > cur_pos
			result = "-" & Mid(result, cur_pos + 1)

			resint = resint - 1
			resdot = resdot - 1
			resfrac = resfrac - 1
		Wend
	Else
		While resint > cur_pos
			result = Mid(result, cur_pos + 1)

			resint = resint - 1
			resdot = resdot - 1
			resfrac = resfrac - 1
		Wend
	End If

	If resscale > 0 Then
		While Mid(result, resfrac + resscale - 1, 1) = "0"
			result = Left(result, Len(result) - 1)
			resscale = resscale - 1
		Wend
		If resscale = 0 Then
			result = Left(result, Len(result) - 1)
			resfrac = resfrac - 1
		End If
	End If

	If result = "-0" Then
		result = "0"
		ressign = 1
		resint = resint - 1
		resdot = resdot - 1
		resfrac = resfrac - 1
	End If

	Trim.sValue = result
	Trim.lSign = ressign
	Trim.lInt = resint
	Trim.lDot = resdot
	Trim.lFrac = resfrac
	Trim.lScale = resscale

	DebugPrint Trim, "Trim"
End Function

'***************************************
'
'               Compare
'
'***************************************

Public Function Compare(other As BigDecimal) As Long
	If Big_lSign <> other.lSign Then
		Compare = IIf(Big_lSign < 0, -1, 1)

		If DebugFlag Then Debug.Print "Compare: " & Compare & vbLf
		Exit Function
	End If

	Compare = compare_positive(other)
	If Big_lSign < 0 Then
		Compare = -Compare
	End If

	If DebugFlag Then Debug.Print "Compare: " & Compare & vbLf
End Function

Public Function DebugPrintLogOf2()
	Debug.Print LogarithmOfTwo().StrValue
End Function

'###################### Private functions ####################################

Private Sub DebugPrint(other As BigDecimal, oName As String)
	If DebugFlag Then
		Debug.Print oName & ": Value=" & other.sValue
		Debug.Print oName & ": lsign=" & other.lSign
		Debug.Print oName & ": lint=" & other.lInt
		Debug.Print oName & ": ldot=" & other.lDot
		Debug.Print oName & ": lfrac=" & other.lFrac
		Debug.Print oName & ": lscale=" & other.lScale
		Debug.Print
	End If
End Sub

Private Function Parse() As Long
	Dim i As Long
	Dim length As Long

	If Left(Big_sValue, 1) = "." Then
		Big_sValue = "0" & Big_sValue
	End If

	If Left(Big_sValue, 2) = "-." Then
		Big_sValue = "-0." & Mid(Big_sValue, 3)
	End If

	i = 1
	Big_lSign = 1
	If Mid(Big_sValue, i, 1) = "-" Or Mid(Big_sValue, i, 1) = "+" Then
		If Mid(Big_sValue, i, 1) = "-" Then
			Big_lSign = -1
		End If
		i = i + 1
	End If

	length = Len(Big_sValue)
	If i > length Then
		Parse = -1
		Exit Function
	End If
	Big_lInt = i

	While i <= length And "0" <= Mid(Big_sValue, i, 1) And Mid(Big_sValue, i, 1) <= "9"
		i = i + 1
	Wend
	Big_lDot = i

	Big_lScale = 0
	If i <= length And Mid(Big_sValue, i, 1) = "." Then
		Big_lScale = Len(Big_sValue) - i
		i = i + 1
	End If
	Big_lFrac = i

	While i <= length And "0" <= Mid(Big_sValue, i, 1) And Mid(Big_sValue, i, 1) <= "9"
		i = i + 1
	Wend

	If i <= length Then
		Parse = -1
		Exit Function
	End If

	While Mid(Big_sValue, Big_lInt, 1) = "0" And Big_lInt + 1 < Big_lDot
		Big_lInt = Big_lInt + 1
	Wend

	'While Big_lScale > 0 And Mid(s, Big_lFrac + Big_lScale - 1, 1) = "0"
	'    Big_lScale = Big_lScale - 1
	'Wend

	If Big_lScale = 0 And Big_lFrac > Big_lDot Then
		Big_lFrac = Big_lFrac - 1
		Debug.Assert Big_lFrac = Big_lDot
	End If

	If Big_lSign < 0 And (Big_lScale = 0 And Mid(Big_sValue, Big_lInt, 1) = "0") Then
		Big_lSign = 1
	End If

	Parse = Big_lScale

	DebugPrint Me, "Parse"
End Function

Private Function add_positive(other As BigDecimal, final_scale As Long, final_sign As Long) As BigDecimal
	Dim llen As Long, olen As Long
	Dim resint As Long, resdot As Long, resfrac As Long, resscale As Long
	Dim result_len As Long, result_scale As Long, result_size As Long
	Dim result As String
	Dim i As Long, um As Long
	Dim cur_pos As Long, was_frac As Long

	Set add_positive = New BigDecimal

	llen = Big_lDot - Big_lInt
	olen = other.lDot - other.lInt

	result_len = Application.WorksheetFunction.Max(llen, olen) + 1
	result_scale = Application.WorksheetFunction.Max(Big_lScale, other.lScale)
	result_size = result_len + result_scale + 3

	result = String(result_size, "0")

	um = 0
	cur_pos = result_size
	was_frac = 0

	For i = result_scale - 1 To 0 Step -1
		If i < Big_lScale Then
			um = um + asc(Mid(Big_sValue, Big_lFrac + i, 1)) - 48
		End If
		If i < other.lScale Then
			um = um + asc(Mid(other.sValue, other.lFrac + i, 1)) - 48
		End If

		If um <> 0 Or was_frac Then
			cur_pos = cur_pos - 1
			Mid(result, cur_pos + 1, 1) = Chr((um Mod 10) + 48)
			um = um \ 10
			was_frac = 1
		End If
	Next i

	resscale = result_size - cur_pos
	resfrac = cur_pos + 1

	If was_frac Then
		cur_pos = cur_pos - 1
		Mid(result, cur_pos + 1, 1) = "."
	End If

	resdot = cur_pos + 1

	For i = 0 To result_len - 1
		If i < llen Then
			um = um + asc(Mid(Big_sValue, Big_lDot - i - 1, 1)) - 48
		End If
		If i < olen Then
			um = um + asc(Mid(other.sValue, other.lDot - i - 1, 1)) - 48
		End If

		cur_pos = cur_pos - 1
		Mid(result, cur_pos + 1, 1) = Chr(um Mod 10 + 48)
		um = um \ 10
	Next i

	resint = cur_pos
	Debug.Assert cur_pos > 0
	While Mid(result, resint, 1) = "0" And resint + 1 < resdot
		resint = resint + 1
	Wend

	If final_sign < 0 Then
		result = "-" & result
		resint = resint + 1
		resdot = resdot + 1
		resfrac = resfrac + 1
	End If

	add_positive.sValue = result
	add_positive.lSign = final_sign
	add_positive.lInt = resint
	add_positive.lDot = resdot
	add_positive.lFrac = resfrac
	add_positive.lScale = resscale

	If RoundFlag Then add_positive.Value = add_positive.Round(final_scale)

	DebugPrint add_positive, "add_positive"
End Function

Private Function subtract_positive(other As BigDecimal, final_scale As Long, final_sign As Long) As BigDecimal
	Dim llen As Long
	Dim olen As Long
	Dim resint As Long, resdot As Long, resfrac As Long, resscale As Long
	Dim result_len As Long, result_scale As Long, result_size As Long
	Dim i As Long, um As Long, next_um As Long
	Dim result As String
	Dim cur_pos As Long
	Dim was_frac As Long

	Set subtract_positive = New BigDecimal

	llen = Big_lDot - Big_lInt
	olen = other.lDot - other.lInt

	result_len = llen
	result_scale = Application.WorksheetFunction.Max(Big_lScale, other.lScale)
	result_size = result_len + result_scale + 3

	result = String(result_size, "0")

	um = 0
	next_um = 0
	cur_pos = result_size
	was_frac = 0

	For i = result_scale - 1 To 0 Step -1
		um = next_um

		If i < Big_lScale Then
			um = um + asc(Mid(Big_sValue, Big_lFrac + i, 1)) - 48
		End If

		If i < other.lScale Then
			um = um - asc(Mid(other.sValue, other.lFrac + i, 1)) + 48
		End If

		If um < 0 Then
			next_um = -1
			um = um + 10
		Else
			next_um = 0
		End If

		If um <> 0 Or was_frac = 1 Then
			Mid(result, cur_pos, 1) = Chr(um + 48)
			cur_pos = cur_pos - 1
			was_frac = 1
		End If
	Next i

	resscale = result_size - cur_pos
	resfrac = cur_pos + 1

	If was_frac = 1 Then
		Mid(result, cur_pos, 1) = "."
		cur_pos = cur_pos - 1
	End If

	resdot = cur_pos + 1

	For i = 0 To result_len - 1
		um = next_um
		um = um + asc(Mid(Big_sValue, Big_lDot - i - 1, 1)) - 48

		If i < olen Then
			um = um - asc(Mid(other.sValue, other.lDot - i - 1, 1)) + 48
		End If

		If um < 0 Then
			next_um = -1
			um = um + 10
		Else
			next_um = 0
		End If

		Mid(result, cur_pos, 1) = Chr(um + 48)
		cur_pos = cur_pos - 1
	Next i

	resint = cur_pos
	Debug.Assert cur_pos > 0
	While Mid(result, resint, 1) = "0" And resint + 1 < resdot
		resint = resint + 1
	Wend

	If final_sign < 0 Then
		result = "-" & result
		resint = resint + 1
		resdot = resdot + 1
		resfrac = resfrac + 1
	End If

	subtract_positive.sValue = result
	subtract_positive.lSign = final_sign
	subtract_positive.lInt = resint
	subtract_positive.lDot = resdot
	subtract_positive.lFrac = resfrac
	subtract_positive.lScale = resscale

	If RoundFlag Then subtract_positive.Value = subtract_positive.Round(final_scale)

	DebugPrint subtract_positive, "subtract_positive"
End Function

Private Function multiply_positive(other As BigDecimal, final_scale As Long, final_sign As Long) As BigDecimal
	Dim llen As Long, olen As Long
	Dim i As Long, j As Long, X As Long, Y As Long
	Dim resint As Long, resdot As Long, resfrac As Long, resscale As Long
	Dim result_len As Long, result_scale As Long, result_size As Long
	Dim result As String
	Dim res() As Long
	Dim cur_pos As Long
	Dim k As Long

	Set multiply_positive = New BigDecimal

	llen = Big_lDot - Big_lInt
	olen = other.lDot - other.lInt

	result_len = llen + olen
	result_scale = Big_lScale + other.lScale
	result_size = result_len + result_scale + 3
	result = String(result_size, "0")

	ReDim res(result_size - 1)

	For i = -Big_lScale To llen - 1
		If i < 0 Then
			X = asc(Mid(Big_sValue, Big_lFrac - i - 1, 1)) - 48
		Else
			X = asc(Mid(Big_sValue, Big_lDot - i - 1, 1)) - 48
		End If

		For j = -other.lScale To olen - 1
			If j < 0 Then
				Y = asc(Mid(other.sValue, other.lFrac - j - 1, 1)) - 48
			Else
				Y = asc(Mid(other.sValue, other.lDot - j - 1, 1)) - 48
			End If
			res(i + j + result_scale) = res(i + j + result_scale) + (X * Y)
		Next j
	Next i

	For i = 0 To result_size - 2
		res(i + 1) = res(i + 1) + (res(i) \ 10)
		res(i) = res(i) Mod 10
	Next i

	cur_pos = result_size + 1
	For k = 0 To result_scale - 1
		cur_pos = cur_pos - 1
		Mid(result, cur_pos, 1) = Chr(res(k) + 48)
	Next k

	resscale = result_size - cur_pos + 1
	resfrac = cur_pos

	If result_scale > 0 Then
		cur_pos = cur_pos - 1
		Mid(result, cur_pos, 1) = "."
	End If
	resdot = cur_pos

	For i = result_scale To result_len + result_scale - 1
		cur_pos = cur_pos - 1
		Mid(result, cur_pos, 1) = Chr(res(i) + 48)
	Next i

	resint = cur_pos
	Debug.Assert cur_pos > 0
	While Mid(result, resint, 1) = "0" And resint + 1 < resdot
		resint = resint + 1
	Wend

	If final_sign < 0 Then
		result = "-" & result
		resint = resint + 1
		resdot = resdot + 1
		resfrac = resfrac + 1
	End If

	multiply_positive.sValue = result
	multiply_positive.lSign = final_sign
	multiply_positive.lInt = resint
	multiply_positive.lDot = resdot
	multiply_positive.lFrac = resfrac
	multiply_positive.lScale = resscale

	If RoundFlag Then multiply_positive.Value = multiply_positive.Round(final_scale)

	DebugPrint multiply_positive, "multiply_positive"
End Function

Private Function divide_positive(other As BigDecimal, final_scale As Long, final_sign As Long) As BigDecimal
	Dim llen As Long
	Dim olen As Long
	Dim resint As Long, resdot As Long, resfrac As Long, resscale As Long
	Dim result_len As Long, result_scale As Long, result_size As Long
	Dim dividend_len As Long
	Dim divider_len As Long, divider_skip As Long
	Dim i As Long, j As Long, X As Long, idx As Long
	Dim cur_pow As Long, cur_pos As Long
	Dim result As String
	Dim beg As Long, real_beg As Long
	Dim dig As String
	Dim less As Boolean

	Set divide_positive = New BigDecimal

	llen = Big_lDot - Big_lInt
	olen = other.lDot - other.lInt

	result_len = Application.WorksheetFunction.Max(llen + other.lScale - olen + 1, 1)
	result_scale = final_scale
	result_size = result_len + result_scale + 3

	If other.lScale = 0 And Mid(other.sValue, other.lInt, 1) = "0" Then
		Err.Raise vbObjectError + 1103, "BigDecimal", "Division By ZERO is not permtted."
		Exit Function
	End If

	dividend_len = llen + Big_lScale
	divider_len = olen + other.lScale

	ReDim dividend(0 To result_size + dividend_len + divider_len - 1) As Long
	ReDim divider(0 To divider_len - 1) As Long

	For i = -Big_lScale To llen - 1
		If i < 0 Then
			X = asc(Mid(Big_sValue, Big_lFrac - i - 1, 1)) - 48
		Else
			X = asc(Mid(Big_sValue, Big_lDot - i - 1, 1)) - 48
		End If
		dividend(llen - i - 1) = X
	Next i

	For i = -other.lScale To olen - 1
		If i < 0 Then
			X = asc(Mid(other.sValue, other.lFrac - i - 1, 1)) - 48
		Else
			X = asc(Mid(other.sValue, other.lDot - i - 1, 1)) - 48
		End If
		divider(olen - i - 1) = X
	Next i

	idx = 0
	divider_skip = 0
	While divider_len > 0 And divider(idx) = 0
		idx = idx + 1
		divider_skip = divider_skip + 1
		divider_len = divider_len - 1
	Wend
	Debug.Assert divider_len > 0

	cur_pow = llen - olen + divider_skip
	cur_pos = 2

	If cur_pow < -final_scale Then
		idx = idx - divider_skip
		divider_len = divider_len + divider_skip
		Exit Function
	End If

	result = String(result_size, "0")

	resint = cur_pos
	If cur_pow < 0 Then
		Mid(result, cur_pos, 1) = "0"
		cur_pos = cur_pos + 1
		resdot = cur_pos
		Mid(result, cur_pos, 1) = "."
		cur_pos = cur_pos + 1
		resfrac = cur_pos
		j = -1
		While j > cur_pow
			Mid(result, cur_pos, 1) = "0"
			cur_pos = cur_pos + 1
			j = j - 1
		Wend
	End If

	beg = 0
	real_beg = 0

	While cur_pow >= -final_scale
		dig = "0"

		Do While True
			If real_beg < beg And dividend(real_beg) = 0 Then
				real_beg = real_beg + 1
			End If

			less = False

			If real_beg = beg Then
				For i = 0 To divider_len - 1
					If dividend(beg + i) <> divider(i + idx) Then
						less = (dividend(beg + i) < divider(i + idx))
						Exit For
					End If
				Next i
			End If

			If less Then
				Exit Do
			End If

			For i = divider_len - 1 To 0 Step -1
				dividend(beg + i) = dividend(beg + i) - divider(i + idx)
				If dividend(beg + i) < 0 Then
					dividend(beg + i) = dividend(beg + i) + 10
					dividend(beg + i - 1) = dividend(beg + i - 1) - 1
				End If
			Next i
			dig = Chr(asc(dig) + 1)
		Loop

		Mid(result, cur_pos, 1) = dig
		cur_pos = cur_pos + 1

		If cur_pow = 0 Then
			resdot = cur_pos
			If final_scale > 0 Then
				Mid(result, cur_pos, 1) = "."
				cur_pos = cur_pos + 1
			End If
			resfrac = cur_pos
		End If

		cur_pow = cur_pow - 1
		beg = beg + 1
	Wend

	resscale = result_size - resfrac + 1

	idx = idx - divider_skip
	divider_len = divider_len + divider_skip

	If final_sign < 0 Then
		result = "-" & result
		resint = resint + 1
		resdot = resdot + 1
		resfrac = resfrac + 1
	End If

	divide_positive.sValue = result
	divide_positive.lSign = final_sign
	divide_positive.lInt = resint
	divide_positive.lDot = resdot
	divide_positive.lFrac = resfrac
	divide_positive.lScale = resscale

	If RoundFlag Then divide_positive.Value = divide_positive.Round(final_scale)

	DebugPrint divide_positive, "divide_positive"
End Function

Private Function LogarithmOfTwo() As BigDecimal
	Dim result As New BigDecimal
	Dim term As New BigDecimal
	Dim i As New BigDecimal
	Dim X As New BigDecimal
	Dim prec As New BigDecimal
	Dim Str As String

	Set LogarithmOfTwo = New BigDecimal

	If LogOf2 = "" Then
		LogOf2 = "0.693147180559945309417232121458176568075500134360255254120680009493393621969694715605863326996418687542001481020570685733685520235758130557032670751637"
		LogOf2Scale = 150
	End If

	If Big_lMaxScale <= LogOf2Scale Then
		LogarithmOfTwo.StrValue = LogOf2

		If Big_lMaxScale < LogOf2Scale Then
			If RoundFlag Then LogarithmOfTwo.Value = LogarithmOfTwo.Round(Big_lMaxScale)
		End If

		Exit Function
	End If

	Str = "0." & String(Big_lMaxScale - 1, "0") & "1"
	prec.StrValue = Str

	X.StrValue = "0.5"
	i.Value = One
	term.Value = One

	While term.Multiply(X).Absolute().Compare(prec) = 1
		result.Value = result.Add(term.Multiply(X))

		i.Value = i.Add(One)
		term.Value = One.Divide(i)
		X.Value = X.Divide(Two)

		DoEvents
	Wend

	LogOf2 = result.sValue           'cache value
	LogOf2Scale = result.lScale

	LogarithmOfTwo.sValue = result.sValue
	LogarithmOfTwo.lSign = 1
	LogarithmOfTwo.lInt = result.lInt
	LogarithmOfTwo.lDot = result.lDot
	LogarithmOfTwo.lFrac = result.lFrac
	LogarithmOfTwo.lScale = result.lScale

	If RoundFlag Then LogarithmOfTwo.Value = LogarithmOfTwo.Round(Big_lMaxScale)

	DebugPrint LogarithmOfTwo, "LogarithmOfTwo"
End Function

Private Function compare_positive(other As BigDecimal) As Long
	Dim llen As Long
	Dim olen As Long
	Dim i As Long
	Dim lchar As Long
	Dim ochar As Long
	Dim resscale As Long

	resscale = Big_lMaxScale

	llen = Big_lDot - Big_lInt
	olen = other.lDot - other.lInt

	If llen <> olen Then
		compare_positive = IIf(llen < olen, -1, 1)
		Exit Function
	End If

	For i = 0 To llen - 1
		If Mid(Big_sValue, Big_lInt + i, 1) <> Mid(other.sValue, other.lInt + i, 1) Then
			compare_positive = IIf(Mid(Big_sValue, Big_lInt + i, 1) < Mid(other.sValue, other.lInt + i, 1), -1, 1)
			Exit Function
		End If
	Next i

	For i = 0 To resscale - 1
		If i < Big_lScale Then
			lchar = asc(Mid(Big_sValue, Big_lFrac + i, 1)) - 48
		Else
			lchar = 0
		End If

		If i < other.lScale Then
			ochar = asc(Mid(other.sValue, other.lFrac + i, 1)) - 48
		Else
			ochar = 0
		End If

		If lchar <> ochar Then
			compare_positive = IIf(lchar < ochar, -1, 1)
			Exit Function
		End If
	Next i

	compare_positive = 0
End Function

'####################### BigDecimal Properties #######################

'Value
Public Property Get Value() As BigDecimal
	Set Value = New BigDecimal

	Value.sValue = Big_sValue
	Value.lSign = Big_lSign
	Value.lInt = Big_lInt
	Value.lDot = Big_lDot
	Value.lFrac = Big_lFrac
	Value.lScale = Big_lScale
End Property

Public Property Let Value(other As BigDecimal)
	Big_sValue = other.sValue
	Big_lSign = other.lSign
	Big_lInt = other.lInt
	Big_lDot = other.lDot
	Big_lFrac = other.lFrac
	Big_lScale = other.lScale
End Property

'StrValue
Public Property Get StrValue() As String
	StrValue = Trim().sValue
End Property

Public Property Let StrValue(v As String)
	Big_sValue = v
	If Parse() = -1 Then
		Err.Raise vbObjectError + 1104, "BigDecimal", "String is a malformed BigDecimal number."
	End If
End Property

'LngValue
Public Property Get LngValue() As Long
	LngValue = CLng(StrValue)
End Property

Public Property Let LngValue(v As Long)
	StrValue = CStr(v)
End Property

'DblValue
Public Property Get DblValue() As Double
	DblValue = CDbl(StrValue)
End Property

Public Property Let DblValue(v As Double)
	StrValue = CStr(v)
End Property

'MaxScale
Public Property Get MaxScale() As Long
	MaxScale = Big_lMaxScale
End Property

Public Property Let MaxScale(newscale As Long)
	If newscale < 0 Then
		Big_lMaxScale = 0
		Exit Property
	End If

	Big_lMaxScale = newscale
End Property

'Pi
Public Property Get Pi() As BigDecimal
	Dim i As New BigDecimal
	Dim result As New BigDecimal
	Dim p As New BigDecimal
	Dim X As New BigDecimal
	Dim term As New BigDecimal
	Dim prec As New BigDecimal
	Dim Str As String

	Set Pi = New BigDecimal

	Str = "0." & String(Big_lMaxScale - 1, "0") & "1"
	prec.StrValue = Str

	term.Value = One
	p.Value = One
	X.StrValue = "16"

	While term.Absolute().Compare(prec) = 1
		term.Value = One.Divide(p).Multiply(Four.Divide(i.Add(One)).Subtract(Two.Divide(i.Add(Four))).Subtract(One.Divide(i.Add(Five))).Subtract(One.Divide(i.Add(Six))))
		'term = (1 / p) * ((4 / (i + 1#)) - (2 / (i + 4#)) - (1 / (i + 5#)) - (1 / (i + 6#)))

		i.Value = i.Add(Eight)
		p.Value = p.Multiply(X)
		result.Value = result.Add(term)

		DoEvents
	Wend

	Pi.sValue = result.sValue
	Pi.lSign = result.lSign
	Pi.lInt = result.lInt
	Pi.lDot = result.lDot
	Pi.lFrac = result.lFrac
	Pi.lScale = result.lScale

	If RoundFlag Then Pi.Value = Pi.Round(Big_lMaxScale)

	DebugPrint Pi, "Pi"
End Property

'Zero
Friend Property Get Zero() As BigDecimal
	Set Zero = New BigDecimal

	Zero.sValue = "0"
	Zero.lSign = 1
	Zero.lInt = 1
	Zero.lDot = 2
	Zero.lFrac = 2
	Zero.lScale = 0

	DebugPrint Zero, "Zero"
End Property

'One
Friend Property Get One() As BigDecimal
	Set One = New BigDecimal

	One.sValue = "1"
	One.lSign = 1
	One.lInt = 1
	One.lDot = 2
	One.lFrac = 2
	One.lScale = 0

	DebugPrint One, "One"
End Property

'Two
Friend Property Get Two() As BigDecimal
	Set Two = New BigDecimal

	Two.sValue = "2"
	Two.lSign = 1
	Two.lInt = 1
	Two.lDot = 2
	Two.lFrac = 2
	Two.lScale = 0

	DebugPrint Two, "Two"
End Property

'Four
Friend Property Get Four() As BigDecimal
	Set Four = New BigDecimal

	Four.sValue = "4"
	Four.lSign = 1
	Four.lInt = 1
	Four.lDot = 2
	Four.lFrac = 2
	Four.lScale = 0

	DebugPrint Four, "Four"
End Property

'Five
Friend Property Get Five() As BigDecimal
	Set Five = New BigDecimal

	Five.sValue = "5"
	Five.lSign = 1
	Five.lInt = 1
	Five.lDot = 2
	Five.lFrac = 2
	Five.lScale = 0

	DebugPrint Five, "Five"
End Property

'Six
Friend Property Get Six() As BigDecimal
	Set Six = New BigDecimal

	Six.sValue = "6"
	Six.lSign = 1
	Six.lInt = 1
	Six.lDot = 2
	Six.lFrac = 2
	Six.lScale = 0

	DebugPrint Six, "Six"
End Property

'Eight
Friend Property Get Eight() As BigDecimal
	Set Eight = New BigDecimal

	Eight.sValue = "8"
	Eight.lSign = 1
	Eight.lInt = 1
	Eight.lDot = 2
	Eight.lFrac = 2
	Eight.lScale = 0

	DebugPrint Eight, "Eight"
End Property

'sValue
Friend Property Get sValue() As String
	sValue = Big_sValue
End Property

Friend Property Let sValue(Val As String)
	Big_sValue = Val
End Property

'lSign
Friend Property Get lSign() As Long
	lSign = Big_lSign
End Property

Friend Property Let lSign(Val As Long)
	Big_lSign = Val
End Property

'lInt
Friend Property Get lInt() As Long
	lInt = Big_lInt
End Property

Friend Property Let lInt(Val As Long)
	Big_lInt = Val
End Property

'lDot
Friend Property Get lDot() As Long
	lDot = Big_lDot
End Property

Friend Property Let lDot(Val As Long)
	Big_lDot = Val
End Property

'lFrac
Friend Property Get lFrac() As Long
	lFrac = Big_lFrac
End Property

Friend Property Let lFrac(Val As Long)
	Big_lFrac = Val
End Property

'lScale
Friend Property Get lScale() As Long
	lScale = Big_lScale
End Property

Friend Property Let lScale(Val As Long)
	Big_lScale = Val
End Property

'lMaxScale
Friend Property Get lMaxScale() As Long
	lMaxScale = Big_lMaxScale
End Property

Friend Property Let lMaxScale(Val As Long)
	Big_lMaxScale = Val
End Property
