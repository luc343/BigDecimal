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
'                             Class BigDecimal
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

' BigDecimal Internal Number Format
'
' -001564.53455666782321312316
'    ^   ^^
'    1   23
'
' Big_sValue = 1564.53455666782321312316
' Big_lSign = -1
' Big_lInt = 4 [1]
' Big_lDot = 8 [2]
' Big_lFrac = 9 [3]
' Big_lScale = 20
'
'note: Len(Big_sValue) = Big_lDot + Big_lScale = 28

Private Big_sValue As String        'number as a string
Private Big_lSign As Long           'sign of number. -1 = -ve, 1 = +ve
Private Big_lInt As Long            'position for start of integer part
Private Big_lDot As Long            'position for decimal point
Private Big_lFrac As Long           'position for start of fraction part
Private Big_lScale As Long          'number of decimal places

Private Big_lMaxScale As Long       'maximum decimal places

Private LogOf2 As String
Private LogOf2Scale As Long

Private OptionOutputDebug As Boolean    'switch debug on/off
Private OptionAutoRound As Boolean      'switch rounding on/off

'######################### BigDecimal Methods ########################

Private Sub Class_Initialize()
    Big_sValue = "0"
    Big_lSign = 1
    Big_lInt = 1
    Big_lDot = 2
    Big_lFrac = 2
    Big_lScale = 0

    Big_lMaxScale = BigOption.MaxScale
    OptionOutputDebug = BigOption.OutputDebug
    OptionAutoRound = BigOption.AutoRound

    BigOption.Register Me
End Sub

Private Sub Class_Terminate()
    BigOption.Unregister Me
End Sub

'***************************************
'
'                  Plus
'
'***************************************

Public Function Plus(Value As Variant) As BigDecimal
    Dim Other As BigDecimal
    Dim Temp As BigDecimal

    Set Other = Cast(Value)
    Set Plus = New BigDecimal

    Set Temp = New BigDecimal

    If Big_lSign > 0 And Other.lSign > 0 Then
	Plus = Plus_Positive(Other, Big_lMaxScale, 1)
	Exit Function
    End If

    If Big_lSign > 0 And Other.lSign < 0 Then
	If compare_positive(Other) >= 0 Then
	    Plus = Minus_Positive(Other, Big_lMaxScale, 1)
	    Exit Function
	Else
	    Temp = Me

	    Me = Other
	    Plus = Minus_Positive(Temp, Big_lMaxScale, -1)
	    Me = Temp
	    Exit Function
	End If
    End If

    If Big_lSign < 0 And Other.lSign > 0 Then
	If compare_positive(Other) <= 0 Then
	    Temp = Me

	    Me = Other
	    Plus = Minus_Positive(Temp, Big_lMaxScale, 1)
	    Me = Temp
	    Exit Function
	Else
	    Plus = Minus_Positive(Other, Big_lMaxScale, -1)
	    Exit Function
	End If
    End If

    Plus = Plus_Positive(Other, Big_lMaxScale, -1)     'Big_lSign < 0 And other.lSign < 0
End Function

'***************************************
'
'                 Minus
'
'***************************************

Public Function Minus(Value As Variant) As BigDecimal
    Dim Other As BigDecimal
    Dim Org_lSign As Long

    Set Other = Cast(Value)
    Set Minus = New BigDecimal

    Org_lSign = Other.lSign

    Other.lSign = Other.lSign * -1
    Minus = Plus(Other)

    Other.lSign = Org_lSign
End Function

'***************************************
'
'                 Times
'
'***************************************

Public Function Times(Value As Variant) As BigDecimal
    Dim Other As BigDecimal

    Set Other = Cast(Value)
    Set Times = New BigDecimal

    Times = Times_Positive(Other, Big_lMaxScale, Big_lSign * Other.lSign)
End Function

'***************************************
'
'                Divide
'
'***************************************

Public Function Divide(Value As Variant) As BigDecimal
    Dim Other As BigDecimal

    Set Other = Cast(Value)
    Set Divide = New BigDecimal

    Divide = Divide_Positive(Other, Big_lMaxScale, Big_lSign * Other.lSign)
End Function

'***************************************
'
'                Remainder
'
'***************************************

Public Function Remdr(Value As Variant) As BigDecimal
    Dim Other As BigDecimal
    Dim Quot As BigDecimal

    Set Other = Cast(Value)
    Set Remdr = New BigDecimal

    Set Quot = New BigDecimal

    Quot = AbsVal().Divide(Other.AbsVal()).Trunc()         'Quotient = Int(Abs(x) / Abs(y))
    Remdr = AbsVal().Minus(Other.AbsVal().Times(Quot))    'Remainder = Abs(dblX) - Abs(dblY) * dblQuotient

    If Big_lSign * Other.lSign < 0 Then
	Remdr = Other.AbsVal().Minus(Remdr)
    End If

    Remdr.lSign = Other.lSign

    DebugPrint Remdr, "Remdr"
End Function

'***************************************
'
'              Square Root
'
'***************************************

Public Function Sqrt() As BigDecimal
    Dim Two As BigDecimal
    Dim X As BigDecimal
    Dim Prec As BigDecimal

    Set Sqrt = New BigDecimal

    Set Two = New BigDecimal
    Set X = New BigDecimal
    Set Prec = New BigDecimal

    If Big_lSign < 0 Then
	Err.Raise vbObjectError + 1101, "BigDecimal", "SquareRoot of negative number is undefined."
	Exit Function
    End If

    Prec = "0." & String$(Big_lMaxScale - 1, "0") & "1"

    Two = 2
    Sqrt = Divide(Two)

    While Sqrt.Minus(X).AbsVal().Cmp(Prec) = 1
	X = Sqrt
	Sqrt = Sqrt.Plus(Divide(Sqrt)).Divide(Two)       'x = (x + value / x) / 2

	DoEvents
    Wend

    If OptionAutoRound Then Sqrt = Sqrt.Round(Big_lMaxScale)

    DebugPrint Sqrt, "Sqrt"
End Function

'***************************************
'
'                Power
'
'***************************************

Public Function Pow(Value As Variant) As BigDecimal
    Dim Other As BigDecimal
    Dim Zero As BigDecimal
    Dim One As BigDecimal

    Set Other = Cast(Value)
    Set Pow = New BigDecimal

    Zero = New BigDecimal
    One = New BigDecimal

    Zero = 0
    One = 1

    If Cmp(Zero) = 0 And Other.Cmp(Zero) <= 0 Then
	Err.Raise vbObjectError + 1102, "BigDecimal", "ZERO to the power of a non-positive number is undefined."
	Exit Function
    End If

    If Big_lSign = -1 And Big_lScale > 0 Then
	Err.Raise vbObjectError + 1102, "BigDecimal", "A negative number to a non-integer power is undefined."
	Exit Function
    End If

    If Other.lScale = 0 Then    'integer power
	Pow = PowInt(Other)
    Else
	Pow = AbsVal().Ln().Times(Other).Exp()

	If Big_lSign = -1 And Other.Trunc().Remdr(2).Cmp(One) = 0 Then
	    Pow = Pow.Neg()
	End If
    End If

    If OptionAutoRound Then Pow = Pow.Round(Big_lMaxScale)

    DebugPrint Pow, "Pow"
End Function

Private Function PowInt(Other As BigDecimal) As BigDecimal
    Dim Base As BigDecimal
    Dim ExpVal As BigDecimal
    Dim Zero As BigDecimal
    Dim One As BigDecimal
    Dim Two As BigDecimal

    Set PowInt = New BigDecimal

    Base = New BigDecimal
    ExpVal = New BigDecimal
    Zero = New BigDecimal
    One = New BigDecimal
    Two = New BigDecimal

    Base = Me
    ExpVal = Other
    Zero = 0
    One = 1
    Two = 2
    PowInt = One

    If ExpVal.Cmp(Zero) = -1 Then   'x^(-y) = (1/x)^y
	Base = One.Divide(Base)
	ExpVal = ExpVal.Neg()
    End If

    Do While ExpVal.Cmp(Zero) = 1
	If ExpVal.Remdr(Two).Cmp(Zero) <> 0 Then    'odd exponent
	    PowInt = PowInt.Times(Base)
	End If

	Base = Base.Times(Base)
	ExpVal = ExpVal.Divide(Two).Trunc()
    Loop
End Function

'***************************************
'
'           Natural Logarithm
'
'***************************************

Public Function Ln() As BigDecimal
    Dim Zero As BigDecimal
    Dim One As BigDecimal
    Dim Two As BigDecimal
    Dim X As BigDecimal
    Dim N As BigDecimal
    Dim Y As BigDecimal
    Dim Term As BigDecimal
    Dim Prec As BigDecimal

    Set Ln = New BigDecimal

    Set Zero = New BigDecimal
    Set One = New BigDecimal
    Set Two = New BigDecimal
    Set X = New BigDecimal
    Set N = New BigDecimal
    Set Y = New BigDecimal
    Set Term = New BigDecimal
    Set Prec = New BigDecimal

    Prec = "0." & String$(Big_lMaxScale - 1, "0") & "1"

    Zero = 0
    One = 1
    Two = 2

    If Cmp(Zero) <= 0 Then   'log is undefined for other <= 0
	Err.Raise vbObjectError + 1102, "BigDecimal", "Logarithm of ZERO or negative number is undefined."
	Exit Function
    End If

    X = Me
    While X.Cmp(Two) >= 0
	N = N.Plus(One)
	X = X.Divide(Two)

	DoEvents
    Wend
    While X.Cmp(One) < 0
	N = N.Minus(One)
	X = X.Times(Two)

	DoEvents
    Wend
    X = X.Minus(One)

    Term = X
    Y = Two
    While Term.AbsVal().Cmp(Prec) = 1
	Ln = Ln.Plus(Term)
	Term = Term.Neg().Times(X).Times(Y.Minus(One)).Divide(Y)       'term = -term * x * (y - 1) / y
	Y = Y.Plus(One)

	DoEvents
    Wend

    Ln = Ln.Plus(N.Times(Ln2()))

    If OptionAutoRound Then Ln = Ln.Round(Big_lMaxScale)

    DebugPrint Ln, "Ln"
End Function

'***************************************
'
'             Exponential
'
'***************************************

Public Function Exp() As BigDecimal
    Dim Term As BigDecimal
    Dim Prec As BigDecimal
    Dim N As BigDecimal
    Dim One As BigDecimal

    Set Exp = New BigDecimal

    Set Term = New BigDecimal
    Set Prec = New BigDecimal
    Set N = New BigDecimal
    Set One = New BigDecimal

    Prec = "0." & String$(Big_lMaxScale - 1, "0") & "1"

    One = 1
    Exp = One
    Term = One
    N = One

    While Term.AbsVal().Cmp(Prec) = 1
	Term = Term.Times(Me).Divide(N)
	Exp = Exp.Plus(Term)
	N = N.Plus(One)

	DoEvents
    Wend

    If OptionAutoRound Then Exp = Exp.Round(Big_lMaxScale)

    DebugPrint Exp, "Exp"
End Function

'***************************************
'
'               Factorial
'
'***************************************

Public Function Fact() As BigDecimal
    Dim One As BigDecimal
    Dim Two As BigDecimal
    Dim i As BigDecimal

    Set Fact = New BigDecimal

    Set One = New BigDecimal
    Set Two = New BigDecimal
    Set i = New BigDecimal

    If Big_lScale <> 0 Then
	Err.Raise vbObjectError + 1103, "BigDecimal", "Factorial of a fraction is not permtted."
	Exit Function
    End If

    One = 1
    Two = 2
    Fact = One
    i = Two

    While i.Cmp(Me) < 1
	Fact = Fact.Times(i)
	i = i.Plus(One)

	DoEvents
    Wend

    If OptionAutoRound Then Fact = Fact.Round(Big_lMaxScale)

    DebugPrint Fact, "Fact"
End Function

'***************************************
'
'            Absolute Value
'
'***************************************

Public Function AbsVal() As BigDecimal
    Dim Res_Int As Long, Res_Dot As Long, Res_Frac As Long
    Dim Res_Value As String

    Set AbsVal = New BigDecimal

    Res_Value = Big_sValue
    Res_Int = Big_lInt
    Res_Dot = Big_lDot
    Res_Frac = Big_lFrac

    If Big_lSign < 0 Then
	Res_Value = Mid$(Res_Value, 2)
	Res_Int = Res_Int - 1
	Res_Dot = Res_Dot - 1
	Res_Frac = Res_Frac - 1
    End If

    AbsVal.sValue = Res_Value
    AbsVal.lSign = 1
    AbsVal.lInt = Res_Int
    AbsVal.lDot = Res_Dot
    AbsVal.lFrac = Res_Frac
    AbsVal.lScale = Big_lScale

    DebugPrint AbsVal, "AbsVal"
End Function

'***************************************
'
'              Negative
'
'***************************************

Public Function Neg() As BigDecimal
    Dim Res_Sign As Long, Res_Int As Long, Res_Dot As Long, Res_Frac As Long
    Dim Res_Value As String
    Dim Comp As Long

    Set Neg = New BigDecimal

    Res_Value = Big_sValue
    Res_Sign = Big_lSign
    Res_Int = Big_lInt
    Res_Dot = Big_lDot
    Res_Frac = Big_lFrac

    Comp = Cmp(0)
    If Comp = -1 Then
	Res_Value = Mid$(Res_Value, 2)
	Res_Sign = 1
	Res_Int = Res_Int - 1
	Res_Dot = Res_Dot - 1
	Res_Frac = Res_Frac - 1
    ElseIf Comp = 1 Then
	Res_Value = "-" & Res_Value
	Res_Sign = -1
	Res_Int = Res_Int + 1
	Res_Dot = Res_Dot + 1
	Res_Frac = Res_Frac + 1
    End If

    Neg.sValue = Res_Value
    Neg.lSign = Res_Sign
    Neg.lInt = Res_Int
    Neg.lDot = Res_Dot
    Neg.lFrac = Res_Frac
    Neg.lScale = Big_lScale

    DebugPrint Neg, "Neg"
End Function

'***************************************
'
'               Truncate
'
'***************************************

Public Function Trunc() As BigDecimal
    Dim Res_Frac As Long, Res_Scale As Long
    Dim Res_Value As String

    Set Trunc = New BigDecimal

    Res_Value = Big_sValue
    Res_Frac = Big_lFrac
    Res_Scale = Big_lScale

    If Big_lScale > 0 Then
	Res_Value = Mid$(Res_Value, 1, Big_lDot - 1)
	Res_Frac = Big_lDot
	Res_Scale = 0
    End If

    Trunc.sValue = Res_Value
    Trunc.lSign = Big_lSign
    Trunc.lInt = Big_lInt
    Trunc.lDot = Big_lDot
    Trunc.lFrac = Res_Frac
    Trunc.lScale = Res_Scale

    DebugPrint Trunc, "Trunc"
End Function

'***************************************
'
'              Fraction
'
'***************************************

Public Function Frac() As BigDecimal
    Dim Res_Int As Long, Res_Dot As Long
    Dim Res_Value As String

    Set Frac = New BigDecimal

    Res_Value = Big_sValue
    Res_Int = Big_lInt
    Res_Dot = Big_lDot

    While Res_Int < Res_Dot
	Mid$(Res_Value, Res_Int, 1) = "0"
	Res_Int = Res_Int + 1
    Wend
    Res_Int = Res_Dot - 1

    Frac.sValue = Res_Value
    Frac.lSign = Big_lSign
    Frac.lInt = Res_Int
    Frac.lDot = Res_Dot
    Frac.lFrac = Big_lFrac
    Frac.lScale = Big_lScale

    DebugPrint Frac, "Frac"
End Function

'***************************************
'
'                Round
'
'***************************************

Public Function Round(ByVal Target_lScale As Long, Optional ByVal Add_Trailing_Zeroes As Boolean = False, Optional ByVal Round_Last As Boolean = True) As BigDecimal
    Dim Res_Sign As Long, Res_Int As Long, Res_Dot As Long, Res_Frac As Long, Res_Scale As Long
    Dim i As Long, kI As Long
    Dim Res_Value As String

    Res_Value = Big_sValue
    Res_Sign = Big_lSign
    Res_Int = Big_lInt
    Res_Dot = Big_lDot
    Res_Frac = Big_lFrac
    Res_Scale = Big_lScale

    Set Round = New BigDecimal

    While Mid$(Res_Value, Res_Int, 1) = "0" And Res_Int + 1 < Res_Dot    'find first non-zero digit except before '.'
	Res_Int = Res_Int + 1
    Wend

    Debug.Assert Res_Int > 0 And Res_Scale >= 0 And Target_lScale >= 0

    If Res_Sign < 0 Then       'temporarily remove the sign
	Res_Value = Mid$(Res_Value, 2)
	Res_Int = Res_Int - 1
	Res_Dot = Res_Dot - 1
	Res_Frac = Res_Frac - 1
    End If

    If Round_Last Then
	If Res_Scale > Target_lScale Then
	    While Target_lScale > 0 And Mid$(Res_Value, Res_Frac + Target_lScale - 1, 1) = "9" And Mid$(Res_Value, Res_Frac + Target_lScale, 1) >= "5"
		Target_lScale = Target_lScale - 1
	    Wend

	    Res_Scale = Target_lScale
	    If Mid$(Res_Value, Res_Frac + Target_lScale, 1) >= "5" Then
		If Target_lScale > 0 Then   'require some decimal places
		    Res_Value = Left$(Res_Value, Res_Frac + Target_lScale - 1 - 1) & Chr$(Asc(Mid$(Res_Value, Res_Frac + Target_lScale - 1, 1)) + 1) & Right$(Res_Value, Len(Res_Value) - Res_Frac - Target_lScale + 1)
		Else                        'require no decimal places
		    Res_Frac = Res_Frac - 1
		    Debug.Assert Res_Frac = Res_Dot

		    For i = 0 To Len(Res_Value) - Res_Dot - 1
			If Res_Dot - i - 1 = 0 Then
			    Res_Value = "0" & Res_Value
			    Res_Dot = Res_Dot + 1
			    Res_Frac = Res_Frac + 1
			End If

			If Mid$(Res_Value, Res_Dot - i - 1, 1) = "9" Then
			    Mid$(Res_Value, Res_Dot - i - 1, 1) = "0"
			Else
			    Mid$(Res_Value, Res_Dot - i - 1, 1) = Chr$(Asc(Mid$(Res_Value, Res_Dot - i - 1, 1)) + 1)
			    Exit For
			End If
		    Next i

		    If Res_Dot - i - 1 < Res_Int Then
			Res_Int = Res_Dot - i - 1
		    End If
		End If
	    End If
	End If
    Else
	If Res_Scale > Target_lScale Then
	    Res_Scale = Target_lScale
	End If
    End If

    If Res_Scale = 0 And Res_Frac > Res_Dot Then    'fix fracton pos
	Res_Frac = Res_Dot
    End If

    If Res_Scale = Target_lScale Or Not Add_Trailing_Zeroes Then
	Res_Value = Left$(Res_Value, Res_Dot) & Mid$(Res_Value, Res_Frac, Res_Scale)
    Else
	Res_Value = Left$(Res_Value, Res_Dot) & Mid$(Res_Value, Res_Frac, Res_Scale)

	If Res_Scale = 0 Then
	    Res_Value = Res_Value & "."
	End If
	For kI = 1 To Target_lScale - Res_Scale
	    Res_Value = Res_Value & "0"
	Next kI
	Res_Scale = Target_lScale
	Res_Frac = Res_Dot + 1
    End If

    If Right$(Res_Value, 1) = "." Then
	Res_Value = Left$(Res_Value, Len(Res_Value) - 1)
    End If

    If Res_Sign < 0 Then       'restore the sign
	Res_Value = "-" & Res_Value
	Res_Int = Res_Int + 1
	Res_Dot = Res_Dot + 1
	Res_Frac = Res_Frac + 1

    End If

    Round.sValue = Res_Value
    Round.lSign = Res_Sign
    Round.lInt = Res_Int
    Round.lDot = Res_Dot
    Round.lFrac = Res_Frac
    Round.lScale = Res_Scale

    DebugPrint Round, "Round"
End Function

'***************************************
'
'              Trim Zero
'
'***************************************

Public Function Trim() As BigDecimal
    Dim Res_Sign As Long, Res_Int As Long, Res_Dot As Long, Res_Frac As Long, Res_Scale As Long
    Dim Res_Value As String
    Dim pos As Long

    Set Trim = New BigDecimal

    Res_Value = Big_sValue
    Res_Sign = Big_lSign
    Res_Int = Big_lInt
    Res_Dot = Big_lDot
    Res_Frac = Big_lFrac
    Res_Scale = Big_lScale

    pos = 1
    If Res_Sign < 0 Then
	pos = pos + 1

	While Res_Int > pos
	    Res_Value = "-" & Mid$(Res_Value, pos + 1)

	    Res_Int = Res_Int - 1
	    Res_Dot = Res_Dot - 1
	    Res_Frac = Res_Frac - 1
	Wend
    Else
	While Res_Int > pos
	    Res_Value = Mid$(Res_Value, pos + 1)

	    Res_Int = Res_Int - 1
	    Res_Dot = Res_Dot - 1
	    Res_Frac = Res_Frac - 1
	Wend
    End If

    If Res_Scale > 0 Then
	While Mid$(Res_Value, Res_Frac + Res_Scale - 1, 1) = "0"
	    Res_Value = Left$(Res_Value, Len(Res_Value) - 1)
	    Res_Scale = Res_Scale - 1
	Wend
	If Res_Scale = 0 Then
	    Res_Value = Left$(Res_Value, Len(Res_Value) - 1)
	    Res_Frac = Res_Frac - 1
	End If
    End If

    If Res_Value = "-0" Then
	Res_Value = "0"
	Res_Sign = 1
	Res_Int = Res_Int - 1
	Res_Dot = Res_Dot - 1
	Res_Frac = Res_Frac - 1
    End If

    Trim.sValue = Res_Value
    Trim.lSign = Res_Sign
    Trim.lInt = Res_Int
    Trim.lDot = Res_Dot
    Trim.lFrac = Res_Frac
    Trim.lScale = Res_Scale

    DebugPrint Trim, "Trim"
End Function

'***************************************
'
'               Compare
'
'***************************************

Public Function Cmp(Value As Variant) As Long
    Dim Other As BigDecimal

    Set Other = Cast(Value)

    If Big_lSign <> Other.lSign Then
	Cmp = IIf(Big_lSign < 0, -1, 1)

	If OptionOutputDebug Then Debug.Print "Cmp: " & Cmp & vbLf
	Exit Function
    End If

    Cmp = compare_positive(Other)
    If Big_lSign < 0 Then
	Cmp = -Cmp
    End If

    If OptionOutputDebug Then Debug.Print "Cmp: " & Cmp & vbLf
End Function

'####################### Private Methods #############################

Private Sub DebugPrint(Other As BigDecimal, oName As String)
    If OptionOutputDebug Then
	Debug.Print oName & ": Value=" & Other.sValue
	Debug.Print oName & ": lSign=" & Other.lSign
	Debug.Print oName & ": lInt=" & Other.lInt
	Debug.Print oName & ": lDot=" & Other.lDot
	Debug.Print oName & ": lFrac=" & Other.lFrac
	Debug.Print oName & ": lScale=" & Other.lScale
	Debug.Print
    End If
End Sub

Friend Function Parse() As Long
    Dim pos As Long
    Dim Length As Long

    If Left$(Big_sValue, 1) = "." Then       'fix missing 0 before .
	Big_sValue = "0" & Big_sValue
    End If

    If Left$(Big_sValue, 2) = "-." Then      'fix missing 0 between - and .
	Big_sValue = "-0." & Mid$(Big_sValue, 3)
    End If

    If Left$(Big_sValue, 2) = "+." Then      'fix missing 0 between + and .
	Big_sValue = "+0." & Mid$(Big_sValue, 3)
    End If

    'set sign
    pos = 1
    Big_lSign = 1
    If Mid$(Big_sValue, pos, 1) = "-" Or Mid$(Big_sValue, pos, 1) = "+" Then
	If Mid$(Big_sValue, pos, 1) = "-" Then
	    Big_lSign = -1
	End If
	pos = pos + 1
    End If

    Length = Len(Big_sValue)
    If pos > Length Then      'catch no digits
	Parse = -1
	Exit Function
    End If
    Big_lInt = pos

    'find dot
    While pos <= Length And "0" <= Mid$(Big_sValue, pos, 1) And Mid$(Big_sValue, pos, 1) <= "9"
	pos = pos + 1
    Wend
    Big_lDot = pos    'set dot position

    Big_lScale = 0
    If pos <= Length And Mid$(Big_sValue, pos, 1) = "." Then
	Big_lScale = Len(Big_sValue) - pos
	pos = pos + 1
	If Big_lScale > Big_lMaxScale Then      'truncate number
	    Big_lScale = Big_lMaxScale
	    Length = Big_lDot + Big_lScale
	    Big_sValue = Mid$(Big_sValue, 1, Length)
	End If
    End If
    Big_lFrac = pos   'set fraction start position

    While pos <= Length And "0" <= Mid$(Big_sValue, pos, 1) And Mid$(Big_sValue, pos, 1) <= "9"
	pos = pos + 1
    Wend

    If pos <= Length Then     'catch non-digits
	Parse = -1
	Exit Function
    End If

    While Mid$(Big_sValue, Big_lInt, 1) = "0" And Big_lInt + 1 < Big_lDot    'find first non-zero digit except before .
	Big_lInt = Big_lInt + 1
    Wend

    While Big_lScale > 0 And Mid$(sValue, Big_lFrac + Big_lScale - 1, 1) = "0"    'trim trailing 0's
	Big_lScale = Big_lScale - 1
    Wend
    Length = Big_lDot + Big_lScale
    Big_sValue = Mid$(Big_sValue, 1, Length)

    If Big_lScale = 0 And Big_lFrac > Big_lDot Then
	Big_lFrac = Big_lFrac - 1
	Debug.Assert Big_lFrac = Big_lDot
    End If

    If Big_lSign < 0 And (Big_lScale = 0 And Mid$(Big_sValue, Big_lInt, 1) = "0") Then   'fix -0 sign
	sValue = Mid$(Big_sValue, 2)
	Big_lSign = 1
	Big_lInt = Big_lInt - 1
	Big_lDot = Big_lDot - 1
	Big_lFrac = Big_lFrac - 1
    End If

    Parse = Big_lScale

    DebugPrint Me, "Parse"
End Function

Private Function Cast(Other As Variant) As BigDecimal
   If IsObject(Other) Then
	Select Case TypeName(Other)
	    Case "BigDecimal"
		Set Cast = Other     'quick, return object pointer

	    Case "Nothing"
		Err.Raise vbObjectError + 1104, "BigDecimal", "Can't cast Nothing to BigDecimal."
		Exit Function

	    Case Else
		Err.Raise vbObjectError + 1105, "BigDecimal", "Can't cast object " & TypeName(Other) & " to BigDecimal."
		Exit Function
	End Select
    Else
	Select Case VarType(Other)
	    Case vbInteger, vbLong, vbSingle, vbDouble, vbString
		Set Cast = New BigDecimal

		Cast.sValue = CStr(Other)
		If Cast.Parse() = -1 Then
		    Err.Raise vbObjectError + 1106, "BigDecimal", TypeName(Other) & " is a malformed BigDecimal number."
		End If

	    Case Else
		Err.Raise vbObjectError + 1107, "BigDecimal", "Can't cast " & TypeName(Other) & " to a BigDecimal."
	End Select
    End If
End Function

Private Function Plus_Positive(Other As BigDecimal, Final_Scale As Long, Final_Sign As Long) As BigDecimal
    Dim lLen As Long, oLen As Long
    Dim Res_Int As Long, Res_Dot As Long, Res_Frac As Long, Res_Scale As Long
    Dim Result_Len As Long, Result_Scale As Long, Result_Size As Long
    Dim Big_sBytes() As Byte, Other_Bytes() As Byte, Res_Bytes() As Byte
    Dim Res_Value As String
    Dim i As Long, um As Long
    Dim pos As Long, Was_Frac As Long

    Set Plus_Positive = New BigDecimal

    lLen = Big_lDot - Big_lInt
    oLen = Other.lDot - Other.lInt

    Result_Len = Application.WorksheetFunction.Max(lLen, oLen) + 1
    Result_Scale = Application.WorksheetFunction.Max(Big_lScale, Other.lScale)
    Result_Size = Result_Len + Result_Scale + 3

    um = 0
    pos = Result_Size
    Was_Frac = 0

    Big_sBytes = Big_sValue
    Other_Bytes = Other.sValue
    Res_Bytes = String$(Result_Size, "0")

    For i = Result_Scale - 1 To 0 Step -1
	If i < Big_lScale Then
	    um = um + Big_sBytes((Big_lFrac + i - 1) * 2) - 48
	End If
	If i < Other.lScale Then
	    um = um + Other_Bytes((Other.lFrac + i - 1) * 2) - 48
	End If

	If um <> 0 Or Was_Frac Then
	    pos = pos - 1
	    Res_Bytes(pos * 2) = (um Mod 10) + 48
	    um = um \ 10
	    Was_Frac = 1
	End If
    Next i

    Res_Scale = Result_Size - pos
    Res_Frac = pos + 1

    If Was_Frac Then
	pos = pos - 1
	Res_Bytes(pos * 2) = 46     'char dot
    End If

    Res_Dot = pos + 1

    For i = 0 To Result_Len - 1
	If i < lLen Then
	    um = um + Big_sBytes((Big_lDot - i - 2) * 2) - 48
	End If
	If i < oLen Then
	    um = um + Other_Bytes((Other.lDot - i - 2) * 2) - 48
	End If

	pos = pos - 1
	Res_Bytes(pos * 2) = (um Mod 10) + 48
	um = um \ 10
    Next i

    Res_Int = pos
    Debug.Assert pos > 0
    While Res_Bytes((Res_Int - 1) * 2) = 48 And Res_Int + 1 < Res_Dot
	Res_Int = Res_Int + 1
    Wend
    Res_Value = Res_Bytes

    If Final_Sign < 0 Then
	Res_Value = "-" & Res_Value
	Res_Int = Res_Int + 1
	Res_Dot = Res_Dot + 1
	Res_Frac = Res_Frac + 1
    End If

    Plus_Positive.sValue = Res_Value
    Plus_Positive.lSign = Final_Sign
    Plus_Positive.lInt = Res_Int
    Plus_Positive.lDot = Res_Dot
    Plus_Positive.lFrac = Res_Frac
    Plus_Positive.lScale = Res_Scale

    If OptionAutoRound Then Plus_Positive = Plus_Positive.Round(Final_Scale)

    DebugPrint Plus_Positive, "Plus_Positive"
End Function

Private Function Minus_Positive(Other As BigDecimal, Final_Scale As Long, Final_Sign As Long) As BigDecimal
    Dim lLen As Long
    Dim oLen As Long
    Dim Res_Int As Long, Res_Dot As Long, Res_Frac As Long, Res_Scale As Long
    Dim Result_Len As Long, Result_Scale As Long, Result_Size As Long
    Dim Big_sBytes() As Byte, Other_Bytes() As Byte, Res_Bytes() As Byte
    Dim i As Long, um As Long, next_um As Long
    Dim Res_Value As String
    Dim pos As Long
    Dim Was_Frac As Long

    Set Minus_Positive = New BigDecimal

    lLen = Big_lDot - Big_lInt
    oLen = Other.lDot - Other.lInt

    Result_Len = lLen
    Result_Scale = Application.WorksheetFunction.Max(Big_lScale, Other.lScale)
    Result_Size = Result_Len + Result_Scale + 3

    um = 0
    next_um = 0
    pos = Result_Size
    Was_Frac = 0

    Big_sBytes = Big_sValue
    Other_Bytes = Other.sValue
    Res_Bytes = String$(Result_Size, "0")

    For i = Result_Scale - 1 To 0 Step -1
	um = next_um

	If i < Big_lScale Then
	    um = um + Big_sBytes((Big_lFrac + i - 1) * 2) - 48
	End If

	If i < Other.lScale Then
	    um = um - Other_Bytes((Other.lFrac + i - 1) * 2) + 48
	End If

	If um < 0 Then
	    next_um = -1
	    um = um + 10
	Else
	    next_um = 0
	End If

	If um <> 0 Or Was_Frac = 1 Then
	    Res_Bytes((pos - 1) * 2) = um + 48
	    pos = pos - 1
	    Was_Frac = 1
	End If
    Next i

    Res_Scale = Result_Size - pos
    Res_Frac = pos + 1

    If Was_Frac = 1 Then
	Res_Bytes((pos - 1) * 2) = 46     'char dot
	pos = pos - 1
    End If

    Res_Dot = pos + 1

    For i = 0 To Result_Len - 1
	um = next_um
	um = um + Big_sBytes((Big_lDot - i - 2) * 2) - 48

	If i < oLen Then
	    um = um - Other_Bytes((Other.lDot - i - 2) * 2) + 48
	End If

	If um < 0 Then
	    next_um = -1
	    um = um + 10
	Else
	    next_um = 0
	End If

	Res_Bytes((pos - 1) * 2) = um + 48
	pos = pos - 1
    Next i

    Res_Int = pos
    Debug.Assert pos > 0
    While Res_Bytes((Res_Int - 1) * 2) = 48 And Res_Int + 1 < Res_Dot
	Res_Int = Res_Int + 1
    Wend

    Res_Value = Res_Bytes
    If Final_Sign < 0 Then
	Res_Value = "-" & Res_Value
	Res_Int = Res_Int + 1
	Res_Dot = Res_Dot + 1
	Res_Frac = Res_Frac + 1
    End If

    Minus_Positive.sValue = Res_Value
    Minus_Positive.lSign = Final_Sign
    Minus_Positive.lInt = Res_Int
    Minus_Positive.lDot = Res_Dot
    Minus_Positive.lFrac = Res_Frac
    Minus_Positive.lScale = Res_Scale

    If OptionAutoRound Then Minus_Positive = Minus_Positive.Round(Final_Scale)

    DebugPrint Minus_Positive, "Minus_Positive"
End Function

Private Function Times_Positive(Other As BigDecimal, Final_Scale As Long, Final_Sign As Long) As BigDecimal
    Dim lLen As Long, oLen As Long, i As Long, j As Long, X As Long, Y As Long
    Dim Res_Int As Long, Res_Dot As Long, Res_Frac As Long, Res_Scale As Long
    Dim Result_Len As Long, Result_Scale As Long, Result_Size As Long
    Dim Big_sBytes() As Byte, Other_Bytes() As Byte, Res_Bytes() As Byte
    Dim Res_Value As String
    Dim Res() As Long
    Dim pos As Long
    Dim k As Long

    Set Times_Positive = New BigDecimal

    lLen = Big_lDot - Big_lInt
    oLen = Other.lDot - Other.lInt

    Result_Len = lLen + oLen
    Result_Scale = Big_lScale + Other.lScale
    Result_Size = Result_Len + Result_Scale + 3

    ReDim Res(Result_Size - 1)

    Big_sBytes = Big_sValue
    Other_Bytes = Other.sValue
    Res_Bytes = String$(Result_Size, "0")

    For i = -Big_lScale To lLen - 1
	If i < 0 Then
	    X = Big_sBytes((Big_lFrac - i - 2) * 2) - 48
	Else
	    X = Big_sBytes((Big_lDot - i - 2) * 2) - 48
	End If

	For j = -Other.lScale To oLen - 1
	    If j < 0 Then
		Y = Other_Bytes((Other.lFrac - j - 2) * 2) - 48
	    Else
		Y = Other_Bytes((Other.lDot - j - 2) * 2) - 48
	    End If
	    Res(i + j + Result_Scale) = Res(i + j + Result_Scale) + (X * Y)
	Next j
    Next i

    For i = 0 To Result_Size - 2
	Res(i + 1) = Res(i + 1) + (Res(i) \ 10)
	Res(i) = Res(i) Mod 10
    Next i

    pos = Result_Size + 1
    For k = 0 To Result_Scale - 1
	pos = pos - 1
	Res_Bytes((pos - 1) * 2) = Res(k) + 48
    Next k

    Res_Scale = Result_Size - pos + 1
    Res_Frac = pos

    If Result_Scale > 0 Then
	pos = pos - 1
	Res_Bytes((pos - 1) * 2) = 46     'char dot
    End If
    Res_Dot = pos

    For i = Result_Scale To Result_Len + Result_Scale - 1
	pos = pos - 1
	Res_Bytes((pos - 1) * 2) = Res(i) + 48
    Next i

    Res_Int = pos
    Debug.Assert pos > 0
    While Res_Bytes((Res_Int - 1) * 2) = 48 And Res_Int + 1 < Res_Dot
	Res_Int = Res_Int + 1
    Wend

    Res_Value = Res_Bytes
    If Final_Sign < 0 Then
	Res_Value = "-" & Res_Value
	Res_Int = Res_Int + 1
	Res_Dot = Res_Dot + 1
	Res_Frac = Res_Frac + 1
    End If

    Times_Positive.sValue = Res_Value
    Times_Positive.lSign = Final_Sign
    Times_Positive.lInt = Res_Int
    Times_Positive.lDot = Res_Dot
    Times_Positive.lFrac = Res_Frac
    Times_Positive.lScale = Res_Scale

    If OptionAutoRound Then Times_Positive = Times_Positive.Round(Final_Scale)

    DebugPrint Times_Positive, "Times_Positive"
End Function

Private Function Divide_Positive(Other As BigDecimal, Final_Scale As Long, Final_Sign As Long) As BigDecimal
    Dim lLen As Long, oLen As Long
    Dim Res_Int As Long, Res_Dot As Long, Res_Frac As Long, Res_Scale As Long
    Dim Result_Len As Long, Result_Scale As Long, Result_Size As Long
    Dim Big_sBytes() As Byte, Other_Bytes() As Byte, Res_Bytes() As Byte
    Dim dividend_len As Long, divider_len As Long, divider_skip As Long
    Dim i As Long, j As Long, X As Long, idx As Long
    Dim cur_pow As Long, pos As Long
    Dim Res_Value As String
    Dim beg As Long, real_beg As Long
    Dim dig As String
    Dim less As Boolean

    Set Divide_Positive = New BigDecimal

    lLen = Big_lDot - Big_lInt
    oLen = Other.lDot - Other.lInt

    Result_Len = Application.WorksheetFunction.Max(lLen + Other.lScale - oLen + 1, 1)
    Result_Scale = Final_Scale
    Result_Size = Result_Len + Result_Scale + 3

    Big_sBytes = Big_sValue
    Other_Bytes = Other.sValue

    If Other.lScale = 0 And Other_Bytes((Other.lInt - 1) * 2) = 48 Then
	Err.Raise vbObjectError + 1108, "BigDecimal", "Division By ZERO is undefined."
	Exit Function
    End If

    dividend_len = lLen + Big_lScale
    divider_len = oLen + Other.lScale

    ReDim dividend(0 To Result_Size + dividend_len + divider_len - 1) As Long
    ReDim divider(0 To divider_len - 1) As Long

    For i = -Big_lScale To lLen - 1
	If i < 0 Then
	    X = Big_sBytes((Big_lFrac - i - 2) * 2) - 48
	Else
	    X = Big_sBytes((Big_lDot - i - 2) * 2) - 48
	End If
	dividend(lLen - i - 1) = X
    Next i

    For i = -Other.lScale To oLen - 1
	If i < 0 Then
	    X = Other_Bytes((Other.lFrac - i - 2) * 2) - 48
	Else
	    X = Other_Bytes((Other.lDot - i - 2) * 2) - 48
	End If
	divider(oLen - i - 1) = X
    Next i

    idx = 0
    divider_skip = 0
    While divider_len > 0 And divider(idx) = 0
	idx = idx + 1
	divider_skip = divider_skip + 1
	divider_len = divider_len - 1
    Wend
    Debug.Assert divider_len > 0

    cur_pow = lLen - oLen + divider_skip
    pos = 2

    If cur_pow < -Final_Scale Then
	idx = idx - divider_skip
	divider_len = divider_len + divider_skip
	Exit Function
    End If

    Res_Bytes = String$(Result_Size, "0")

    Res_Int = pos
    If cur_pow < 0 Then
	Res_Bytes((pos - 1) * 2) = 48   'char zero
	pos = pos + 1
	Res_Dot = pos
	Res_Bytes((pos - 1) * 2) = 46   'char dot
	pos = pos + 1
	Res_Frac = pos
	j = -1
	While j > cur_pow
	    Res_Bytes((pos - 1) * 2) = 48   'char zero
	    pos = pos + 1
	    j = j - 1
	Wend
    End If

    beg = 0
    real_beg = 0

    While cur_pow >= -Final_Scale
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
	    dig = Chr$(Asc(dig) + 1)
	Loop

	Res_Bytes((pos - 1) * 2) = Asc(dig)
	pos = pos + 1

	If cur_pow = 0 Then
	    Res_Dot = pos
	    If Final_Scale > 0 Then
		Res_Bytes((pos - 1) * 2) = 46   'char dot
		pos = pos + 1
	    End If
	    Res_Frac = pos
	End If

	cur_pow = cur_pow - 1
	beg = beg + 1
    Wend

    Res_Scale = Result_Size - Res_Frac + 1

    idx = idx - divider_skip
    divider_len = divider_len + divider_skip

    Res_Value = Res_Bytes
    If Final_Sign < 0 Then
	Res_Value = "-" & Res_Value
	Res_Int = Res_Int + 1
	Res_Dot = Res_Dot + 1
	Res_Frac = Res_Frac + 1
    End If

    Divide_Positive.sValue = Res_Value
    Divide_Positive.lSign = Final_Sign
    Divide_Positive.lInt = Res_Int
    Divide_Positive.lDot = Res_Dot
    Divide_Positive.lFrac = Res_Frac
    Divide_Positive.lScale = Res_Scale

    If OptionAutoRound Then Divide_Positive = Divide_Positive.Round(Final_Scale)

    DebugPrint Divide_Positive, "divide_positive"
End Function

Private Function Ln2() As BigDecimal
    Dim Result As BigDecimal
    Dim Term As BigDecimal
    Dim i As BigDecimal
    Dim X As BigDecimal
    Dim Prec As BigDecimal
    Dim One As BigDecimal
    Dim Two As BigDecimal

    Set Ln2 = New BigDecimal

    Set Result = New BigDecimal
    Set Term = New BigDecimal
    Set i = New BigDecimal
    Set X = New BigDecimal
    Set Prec = New BigDecimal
    Set One = New BigDecimal
    Set Two = New BigDecimal

    If LogOf2 = "" Then
	LogOf2 = "0.693147180559945309417232121458176568075500134360255254120680009493393621969694715605863326996418687542001481020570685733685520235758130557032670751637"
	LogOf2Scale = 150
    End If

    If Big_lMaxScale <= LogOf2Scale Then
	Ln2 = LogOf2

	If Big_lMaxScale < LogOf2Scale Then
	    If OptionAutoRound Then Ln2 = Ln2.Round(Big_lMaxScale)
	End If

	Exit Function
    End If

    Prec = "0." & String$(Big_lMaxScale - 1, "0") & "1"

    One = 1
    Two = 2
    X = 0.5
    i = One
    Term = One

    While Term.Times(X).AbsVal().Cmp(Prec) = 1
	Result = Result.Plus(Term.Times(X))

	i = i.Plus(One)
	Term = One.Divide(i)
	X = X.Divide(Two)

	DoEvents
    Wend

    LogOf2 = Result.sValue       'cache value
    LogOf2Scale = Result.lScale

    Ln2.sValue = Result.sValue
    Ln2.lSign = 1
    Ln2.lInt = Result.lInt
    Ln2.lDot = Result.lDot
    Ln2.lFrac = Result.lFrac
    Ln2.lScale = Result.lScale

    If OptionAutoRound Then Ln2 = Ln2.Round(Big_lMaxScale)

    DebugPrint Ln2, "Ln2"
End Function

Private Function compare_positive(Other As BigDecimal) As Long
    Dim lLen As Long, oLen As Long, i As Long
    Dim lchar As Long, ochar As Long, Res_Scale As Long

    Res_Scale = Big_lMaxScale

    lLen = Big_lDot - Big_lInt
    oLen = Other.lDot - Other.lInt

    If lLen <> oLen Then
	compare_positive = IIf(lLen < oLen, -1, 1)
	Exit Function
    End If

    For i = 0 To lLen - 1
	If Mid$(Big_sValue, Big_lInt + i, 1) <> Mid$(Other.sValue, Other.lInt + i, 1) Then
	    compare_positive = IIf(Mid$(Big_sValue, Big_lInt + i, 1) < Mid$(Other.sValue, Other.lInt + i, 1), -1, 1)
	    Exit Function
	End If
    Next i

    For i = 0 To Res_Scale - 1
	If i < Big_lScale Then
	    lchar = Asc(Mid$(Big_sValue, Big_lFrac + i, 1)) - 48
	Else
	    lchar = 0
	End If

	If i < Other.lScale Then
	    ochar = Asc(Mid$(Other.sValue, Other.lFrac + i, 1)) - 48
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
Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Set Value = New BigDecimal

    Value.sValue = Big_sValue
    Value.lSign = Big_lSign
    Value.lInt = Big_lInt
    Value.lDot = Big_lDot
    Value.lFrac = Big_lFrac
    Value.lScale = Big_lScale
End Property

Public Property Let Value(Other As Variant)
Attribute Value.VB_UserMemId = 0
   If IsObject(Other) Then
	If Other Is Nothing Then
	    Err.Raise vbObjectError + 1109, "BigDecimal", "BigDecimal can't be nothing."
	End If

	If TypeName(Other) = "BigDecimal" Then
	    Dim Temp As BigDecimal

	    Set Temp = Other
	    Big_sValue = Temp.sValue
	    Big_lSign = Temp.lSign
	    Big_lInt = Temp.lInt
	    Big_lDot = Temp.lDot
	    Big_lFrac = Temp.lFrac
	    Big_lScale = Temp.lScale

	    If Big_lScale > Big_lMaxScale Then      'truncate
		Big_lScale = Big_lMaxScale
		Big_sValue = Mid$(Big_sValue, 1, Big_lDot + Big_lScale)
	    End If
	Else
	    Err.Raise vbObjectError + 1110, "BigDecimal", "Can't cast object " & TypeName(Other) & " to BigDecimal."
	End If
    Else
	Select Case VarType(Other)
	    Case vbInteger, vbLong, vbLongLong, vbSingle, vbDouble, vbString
		Big_sValue = CStr(Other)
		If Parse() = -1 Then
		    Err.Raise vbObjectError + 1111, "BigDecimal", TypeName(Other) & " is a malformed BigDecimal number."
		End If
	    Case Else
		Err.Raise vbObjectError + 1112, "BigDecimal", "Can't cast " & TypeName(Other) & " to BigDecimal."
	End Select
    End If
End Property

'ValueEx
Public Property Get ValueEx(RetType As Long) As Variant
    Select Case RetType
    Case vbInteger
	ValueEx = CInt(Big_sValue)
    Case vbLong
	ValueEx = CLng(Big_sValue)
    Case vbLongLong
	ValueEx = CLngLng(Big_sValue)
    Case vbDouble
	Value = CDbl(Big_sValue)
    Case vbString
	ValueEx = CStr(Trim().sValue) 'must trim before returning
    Case Else
	Err.Raise vbObjectError + 1113, "BigDecimal", "Unsupported variable type " & RetType
    End Select
End Property

'MaxScale
Public Property Get MaxScale() As Long
    MaxScale = Big_lMaxScale
End Property

Public Property Let MaxScale(NewScale As Long)
    If NewScale < 0 Then
	Big_lMaxScale = 0
	Exit Property
    End If

    Big_lMaxScale = NewScale

    If Big_lScale > Big_lMaxScale Then      'truncate
	Big_lScale = Big_lMaxScale
	Big_sValue = Mid$(Big_sValue, 1, Big_lDot + Big_lScale)
    End If
End Property

'Pi
Public Property Get Pi() As BigDecimal
    Dim i As New BigDecimal
    Dim Result As New BigDecimal
    Dim p As New BigDecimal
    Dim X As New BigDecimal
    Dim Term As New BigDecimal
    Dim Prec As New BigDecimal
    Dim One As New BigDecimal, Two As New BigDecimal, Four As New BigDecimal, Five As New BigDecimal, Six As New BigDecimal, Eight As New BigDecimal

    Set Pi = New BigDecimal

    Prec = "0." & String$(Big_lMaxScale - 1, "0") & "1"

    One = 1
    Two = 2
    Four = 4
    Five = 5
    Six = 6
    Eight = 8

    Term = One
    p = One
    X = 16

    While Term.AbsVal().Cmp(Prec) = 1
	Term = One.Divide(p).Times(Four.Divide(i.Plus(One)).Minus(Two.Divide(i.Plus(Four))).Minus(One.Divide(i.Plus(Five))).Minus(One.Divide(i.Plus(Six))))
	'term = (1 / p) * ((4 / (i + 1#)) - (2 / (i + 4#)) - (1 / (i + 5#)) - (1 / (i + 6#)))

	i = i.Plus(Eight)
	p = p.Times(X)
	Result = Result.Plus(Term)

	DoEvents
    Wend

    Pi.sValue = Result.sValue
    Pi.lSign = Result.lSign
    Pi.lInt = Result.lInt
    Pi.lDot = Result.lDot
    Pi.lFrac = Result.lFrac
    Pi.lScale = Result.lScale

    If OptionAutoRound Then Pi = Pi.Round(Big_lMaxScale)

    DebugPrint Pi, "Pi"
End Property

'AutoRound
Public Property Get AutoRound() As Boolean
    AutoRound = OptionAutoRound
End Property

Public Property Let AutoRound(Flag As Boolean)
    OptionAutoRound = Flag
End Property

'OutputDebug
Public Property Get OutputDebug() As Boolean
    OutputDebug = OptionOutputDebug
End Property

Public Property Let OutputDebug(Flag As Boolean)
    OptionOutputDebug = Flag
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
