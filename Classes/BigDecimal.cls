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
' Copyright (c) Lucien Cinc 2025-26
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
' Based on the work by Sam-bit from the site:
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

Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)

Private Big_sValue As String        'number as a string
Private Big_lSign As Long           'sign of number. -1 = -ve, 1 = +ve
Private Big_lInt As Long            'position for start of integer part
Private Big_lDot As Long            'position for decimal point
Private Big_lFrac As Long           'position for start of fraction part
Private Big_lScale As Long          'number of decimal places

Private Big_lMaxScale As Long       'maximum decimal places
Private Big_lTarScale As Long       'target decimal places

Private Const Big_lExtra As Long = 5        'guard digits

Private OptionOutputDebug As Boolean        'switch debug on/off
Private OptionAutoRound As Boolean          'switch rounding on/off when return a string
Private OptionAddTrailingZeroes As Boolean  'switch trailing zeroes on/off when using round()

Private Const Asc_Dot As Byte = 46      'asc '.'
Private Const Asc_0 As Byte = 48        'asc '0'
Private Const Asc_5 As Byte = 53        'asc '5'
Private Const Asc_9 As Byte = 57        'asc '9'

'######################### BigDecimal Methods ########################

Private Sub Class_Initialize()
    Big_sValue = "0"
    Big_lSign = 1
    Big_lInt = 1
    Big_lDot = 2
    Big_lFrac = 2
    Big_lScale = 0

    Big_lTarScale = BigOption.TarScale
    Big_lMaxScale = Big_lTarScale + Big_lExtra

    OptionOutputDebug = BigOption.OutputDebug
    OptionAutoRound = BigOption.AutoRound
    OptionAddTrailingZeroes = BigOption.AddTrailingZeroes

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

    Set Other = Cast(Value)

    If Big_lSign > 0 And Other.lSign > 0 Then
	Set Plus = Plus_Positive(Other, 1)
	Exit Function
    End If

    If Big_lSign > 0 And Other.lSign < 0 Then
	If Compare_Positive(Other) >= 0 Then
	    Set Plus = Minus_Positive(Other, 1)
	    Exit Function
	Else
	    Set Plus = Other.Minus_Positive(Me, -1)
	    Exit Function
	End If
    End If

    If Big_lSign < 0 And Other.lSign > 0 Then
	If Compare_Positive(Other) <= 0 Then
	    Set Plus = Other.Minus_Positive(Me, 1)
	    Exit Function
	Else
	    Set Plus = Minus_Positive(Other, -1)
	    Exit Function
	End If
    End If

    Set Plus = Plus_Positive(Other, -1)     'Big_lSign < 0 And other.lSign < 0
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

    Org_lSign = Other.lSign

    Other.lSign = Other.lSign * -1
    Set Minus = Me.Plus(Other)

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

    Set Times = Times_Positive(Other, Big_lSign * Other.lSign)
End Function

'***************************************
'
'                Divide
'
'***************************************

Public Function Divide(Value As Variant) As BigDecimal
    Dim Other As BigDecimal

    Set Other = Cast(Value)

    Set Divide = Divide_Positive(Other, Big_lSign * Other.lSign)
End Function

'***************************************
'
'                Remainder
'
'***************************************

Public Function Remdr(Value As Variant) As BigDecimal
    Dim Other As BigDecimal
    Dim Quot As BigDecimal
    Dim MeAbs As BigDecimal
    Dim OtherAbs As BigDecimal

    Set Other = Cast(Value)

    Set MeAbs = Me.AbsVal()
    Set OtherAbs = Other.AbsVal()

    Set Quot = MeAbs.Divide(OtherAbs).Trunc()       'Quotient = Int(Abs(x / y))
    Set Remdr = MeAbs.Minus(OtherAbs.Times(Quot))   'Remainder = Abs(dblX) - Abs(dblY) * dblQuotient

    If Big_lSign * Other.lSign < 0 Then
	Set Remdr = Other.AbsVal().Minus(Remdr)
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
    Dim X As BigDecimal
    Dim Prec As BigDecimal

    If Big_lSign < 0 Then
	Err.Raise vbObjectError + 1101, "BigDecimal", "SquareRoot of negative number is undefined."
    End If

    Set X = New BigDecimal
    Set Prec = New BigDecimal

    Prec = "0." & String$(Big_lMaxScale - 1, "0") & "1"

    Set Sqrt = Me.Divide(BigTwo)

    While Sqrt.Minus(X).AbsVal().IsGT(Prec)
	X = Sqrt
	Set Sqrt = Sqrt.Plus(Divide(Sqrt)).Divide(BigTwo)       'x = (x + value / x) / 2

	Yield.Events
    Wend

    DebugPrint Sqrt, "Sqrt"
End Function

'***************************************
'
'                Square
'
'***************************************

Public Function Sqr() As BigDecimal
    Set Sqr = Me.Times(Me)

    DebugPrint Sqr, "Sqr"
End Function

'***************************************
'
'                 Cube
'
'***************************************

Public Function Cube() As BigDecimal
    Set Cube = Me.Times(Me).Times(Me)

    DebugPrint Cube, "Cube"
End Function

'***************************************
'
'                Cubert
'
'***************************************

Public Function Cubert() As BigDecimal
    If Me.IsZero() Then
	Set Cubert = New BigDecimal
	Cubert = BigZero
	Exit Function
    End If

    Set Cubert = Me.AbsVal().Ln().Divide(3).Exp()   'exp(ln(|x|) / 3)

    If Big_lSign < 0 Then
	Set Cubert = Cubert.Neg()
    End If

    DebugPrint Cubert, "Cubert"
End Function

'***************************************
'
'                Power
'
'***************************************

Public Function Pow(Value As Variant) As BigDecimal
    Dim Other As BigDecimal

    Set Other = Cast(Value)

    If Me.IsEq(BigZero) And Other.IsLE(BigZero) Then
	Err.Raise vbObjectError + 1102, "BigDecimal", "ZERO to the power of a non-positive number is undefined."
    End If

    If Big_lSign = -1 And Big_lScale > 0 Then
	Err.Raise vbObjectError + 1103, "BigDecimal", "A negative number to a non-integer power is undefined."
    End If

    If Other.lScale = 0 Then    'integer power
	Set Pow = PowInt(Other)
    Else
	Set Pow = Me.AbsVal().Ln().Times(Other).Exp()

	If Big_lSign = -1 And Other.Trunc().Remdr(BigTwo).IsEq(BigOne) Then
	    Set Pow = Pow.Neg()
	End If
    End If

    DebugPrint Pow, "Pow"
End Function

Friend Function PowInt(Other As BigDecimal) As BigDecimal
    Dim X As BigDecimal
    Dim Y As BigDecimal

    Set X = New BigDecimal
    Set Y = New BigDecimal

    X = Me
    Y = Other
    Set PowInt = BigOne

    If Y.IsLT(BigZero) Then   'x^(-y) = (1/x)^y
	Set X = BigOne.Divide(X)
	Set Y = Y.Neg()
    End If

    Do While Y.IsGT(BigZero)
	If Y.Remdr(BigTwo).IsNEq(BigZero) Then    'odd exponent
	    Set PowInt = PowInt.Times(X)
	End If

	Set X = X.Times(X)
	Set Y = Y.Divide(BigTwo).Trunc()

	Yield.Events
    Loop
End Function

'***************************************
'
'           Ten to the Power
'
'***************************************

Public Function Pow10() As BigDecimal
    If Big_lScale = 0 Then
	Dim X As BigDecimal

	Set X = New BigDecimal
	X = 10
	Set Pow10 = X.PowInt(Me)
	Exit Function
    End If

    Set Pow10 = Me.Times(Ln10()).Exp()

    DebugPrint Pow10, "Pow10"
End Function

'***************************************
'
'           Logarithm Base 10
'
'***************************************

Public Function Log() As BigDecimal
    If Me.IsLE(BigZero) Then
	Err.Raise vbObjectError + 1104, "BigDecimal", "Logarithm of ZERO or negative number is undefined."
    End If

    Set Log = Me.Ln().Divide(Ln10())

    DebugPrint Log, "Log"
End Function

'***************************************
'
'           Natural Logarithm
'
'***************************************

'Mercator/Taylor series
Public Function Ln() As BigDecimal
    Dim X As BigDecimal
    Dim n As BigDecimal
    Dim Y As BigDecimal
    Dim Term As BigDecimal
    Dim Prec As BigDecimal

    If Me.IsLE(BigZero) Then   'log is undefined for other <= 0
	Err.Raise vbObjectError + 1105, "BigDecimal", "Logarithm of ZERO or negative number is undefined."
    End If

    Set X = New BigDecimal  'all needed
    Set Prec = New BigDecimal

    Prec = "0." & String$(Big_lMaxScale - 1, "0") & "1"

    Set n = BigZero
    X = Me
    While X.IsGE(BigTwo)
	Set n = n.Plus(BigOne)
	Set X = X.Divide(BigTwo)

	Yield.Events
    Wend
    While X.IsLT(BigOne)
	Set n = n.Minus(BigOne)
	Set X = X.Times(BigTwo)

	Yield.Events
    Wend
    Set X = X.Minus(BigOne)

    Set Ln = BigZero
    Set Term = X
    Set Y = BigTwo
    While Term.AbsVal().IsGT(Prec)
	Set Ln = Ln.Plus(Term)
	Set Term = Term.Neg().Times(X).Times(Y.Minus(BigOne)).Divide(Y)       'term = -term * x * (y - 1) / y
	Set Y = Y.Plus(BigOne)

	Yield.Events
    Wend

    Set Ln = Ln.Plus(n.Times(Ln2()))

    DebugPrint Ln, "Ln"
End Function

Private Function Ln2() As BigDecimal
    Dim X As BigDecimal
    Dim NextX As BigDecimal
    Dim ExpX As BigDecimal
    Dim Fraction As BigDecimal
    Dim Prec As BigDecimal
    Dim Done As Boolean

    If Big_lMaxScale <= BigOption.Ln2_Scale Then
	Set Ln2 = New BigDecimal
	Ln2 = BigOption.Ln2_Buf
	If Big_lMaxScale < BigOption.Ln2_Scale Then
	    Set Ln2 = Ln2.Round(Big_lMaxScale)
	End If
	Exit Function
    End If

    Set X = New BigDecimal
    Set Fraction = New BigDecimal
    Set Prec = New BigDecimal

    Prec = "0." & String$(Big_lMaxScale - 1, "0") & "1"

    X = BigOption.Ln2_Buf   'init with large cache

    Done = False
    Do While Not Done
	Set ExpX = X.Exp()
	Fraction = BigTwo.Minus(ExpX).Divide(BigTwo.Plus(ExpX))
	Set NextX = X.Plus(BigTwo.Times(Fraction))

	If NextX.Minus(X).AbsVal().IsLT(Prec) Then
	    Done = True
	End If

	Set X = NextX
	Yield.Events
    Loop

    BigOption.Ln2_Buf = X.sValue
    BigOption.Ln2_Scale = X.lScale

    Set Ln2 = X

    DebugPrint Ln2, "Ln2"
End Function

Private Function Ln10() As BigDecimal
    Dim X As BigDecimal
    Dim NextX As BigDecimal
    Dim ExpX As BigDecimal
    Dim Ten As BigDecimal
    Dim Fraction As BigDecimal
    Dim Prec As BigDecimal
    Dim Done As Boolean

    If Big_lMaxScale <= BigOption.Ln10_Scale Then
	Set Ln10 = New BigDecimal
	Ln10 = BigOption.Ln10_Buf
	If Big_lMaxScale < BigOption.Ln10_Scale Then
	    Set Ln10 = Ln10.Round(Big_lMaxScale)
	End If
	Exit Function
    End If

    Set X = New BigDecimal
    Set Fraction = New BigDecimal
    Set Ten = New BigDecimal
    Set Prec = New BigDecimal

    Prec = "0." & String$(Big_lMaxScale - 1, "0") & "1"
    Ten = "10"

    X = BigOption.Ln10_Buf   'init with large cache

    Done = False
    Do While Not Done
	Set ExpX = X.Exp()
	Fraction = Ten.Minus(ExpX).Divide(Ten.Plus(ExpX))
	Set NextX = X.Plus(BigTwo.Times(Fraction))

	If NextX.Minus(X).AbsVal().IsLT(Prec) Then
	    Done = True
	End If

	Set X = NextX
	Yield.Events
    Loop

    BigOption.Ln10_Buf = X.sValue
    BigOption.Ln10_Scale = X.lScale

    Set Ln10 = X

    DebugPrint Ln10, "Ln10"
End Function

'***************************************
'
'             Exponential
'
'***************************************

Public Function Exp() As BigDecimal
    Dim Term As BigDecimal
    Dim Prec As BigDecimal
    Dim n As BigDecimal
    Dim ReducedX As BigDecimal
    Dim Half As BigDecimal
    Dim K As Long
    Dim i As Long

    Set Prec = New BigDecimal
    Set ReducedX = New BigDecimal
    Set Half = New BigDecimal

    Prec = "0." & String$(Big_lMaxScale - 1, "0") & "1"

    K = 0
    Half = 0.5
    ReducedX = Me

    While ReducedX.AbsVal().IsGT(Half)
	Set ReducedX = ReducedX.Divide(BigTwo)
	K = K + 1

	Yield.Events
    Wend

    Set Exp = BigOne
    Set Term = BigOne
    Set n = BigOne

    While Term.AbsVal().IsGT(Prec)
	Set Term = Term.Times(ReducedX).Divide(n)
	Set Exp = Exp.Plus(Term)
	Set n = n.Plus(BigOne)

	Yield.Events
    Wend

    For i = 1 To K
	Set Exp = Exp.Times(Exp)

	Yield.Events
    Next i

    DebugPrint Exp, "Exp"
End Function

'***************************************
'
'               Invert
'
'***************************************

Public Function Invert() As BigDecimal
    If Me.IsZero() Then
	Err.Raise vbObjectError + 1105, "BigDecimal", "Division by ZERO is undefined."
    End If

    Set Invert = BigOne.Divide(Me)

    DebugPrint Invert, "Invert"
End Function

'***************************************
'
'       Greatest Common Divisor
'
'***************************************

Public Function GCD(Value As Variant) As BigDecimal
    Dim Other As BigDecimal
    Dim X As BigDecimal
    Dim Y As BigDecimal
    Dim Remainder As BigDecimal
    Dim MaxScale As Long

    Set Other = Cast(Value)

    Set X = Me.AbsVal().Trim()
    Set Y = Other.AbsVal().Trim()

    If X.IsZero() Then  'zero edge cases
	Set GCD = Y
	Exit Function
    End If
    If Y.IsZero() Then
	Set GCD = X
	Exit Function
    End If

    If X.lScale > Y.lScale Then
	MaxScale = X.lScale
    Else
	MaxScale = Y.lScale
    End If

    Set X = X.Shift(MaxScale)
    Set Y = Y.Shift(MaxScale)

    Do While Not Y.IsZero()
	Set Remainder = X.Remdr(Y)

	Set X = Y
	Set Y = Remainder

	Yield.Events
    Loop

    Set GCD = X.Shift(-MaxScale)

    DebugPrint GCD, "GCD"
End Function

'***************************************
'
'       Least Common Multiple
'
'***************************************

Public Function LCM(Value As Variant) As BigDecimal
    Dim Other As BigDecimal

    Set Other = Cast(Value)

    If Me.IsZero() Or Other.IsZero() Then
	Set LCM = New BigDecimal
	LCM = BigZero
	Exit Function
    End If

    Set LCM = Me.Divide(Me.GCD(Other)).Times(Other).AbsVal()

    DebugPrint LCM, "LCM"
End Function

'***************************************
'
'               Factorial
'
'***************************************

Public Function Fact() As BigDecimal
    Dim X As BigDecimal

    If Big_lScale <> 0 Then
	Err.Raise vbObjectError + 1106, "BigDecimal", "Factorial of a fraction is not permtted."
    End If

    Set Fact = BigOne
    Set X = BigTwo

    While Me.IsGE(X)   'me >= X
	Set Fact = Fact.Times(X)
	Set X = X.Plus(BigOne)

	Yield.Events
    Wend

    DebugPrint Fact, "Fact"
End Function

'***************************************
'
'            Absolute Value
'
'***************************************

Public Function AbsVal() As BigDecimal
    Dim Res_Int As Long, Res_Dot As Long, Res_Frac As Long
    Dim Res_Bytes() As Byte

    Set AbsVal = New BigDecimal

    If Big_lSign > 0 Then
	AbsVal.SetRaw Big_sValue, Big_lSign, Big_lInt, Big_lDot, Big_lFrac, Big_lScale
	DebugPrint AbsVal, "AbsVal"
	Exit Function
    End If

    Res_Bytes = Big_sValue

    RemoveCh Res_Bytes
    Res_Int = Big_lInt - 1
    Res_Dot = Big_lDot - 1
    Res_Frac = Big_lFrac - 1

    AbsVal.SetRawB Res_Bytes, 1, Res_Int, Res_Dot, Res_Frac, Big_lScale

    DebugPrint AbsVal, "AbsVal"
End Function

'***************************************
'
'              Negative
'
'***************************************

Public Function Neg() As BigDecimal
    Dim Res_Sign As Long, Res_Int As Long, Res_Dot As Long, Res_Frac As Long
    Dim Res_Bytes() As Byte

    Set Neg = New BigDecimal

    Res_Bytes = Big_sValue
    Res_Sign = Big_lSign
    Res_Int = Big_lInt
    Res_Dot = Big_lDot
    Res_Frac = Big_lFrac

    Select Case CmpTo(0)
    Case -1     'Me < 0
	RemoveCh Res_Bytes
	Res_Sign = 1
	Res_Int = Res_Int - 1
	Res_Dot = Res_Dot - 1
	Res_Frac = Res_Frac - 1
    Case 1      'Me > 0
	PrependCh Res_Bytes, "-"
	Res_Sign = -1
	Res_Int = Res_Int + 1
	Res_Dot = Res_Dot + 1
	Res_Frac = Res_Frac + 1
    End Select

    Neg.SetRawB Res_Bytes, Res_Sign, Res_Int, Res_Dot, Res_Frac, Big_lScale

    DebugPrint Neg, "Neg"
End Function

'***************************************
'
'               Floor
'
'***************************************

Public Function Floor() As BigDecimal
    If Big_lScale = 0 Or Me.IsZero() Then
	Set Floor = New BigDecimal
	Floor = Me
    ElseIf Me.IsPos() Then
	Set Floor = Me.Trunc()
    Else    'is negative
	Set Floor = Me.Trunc().Minus(BigOne)
    End If

    DebugPrint Floor, "Floor"
End Function

'***************************************
'
'               Ceiling
'
'***************************************

Public Function Ceiling() As BigDecimal
    If Big_lScale = 0 Or Me.IsZero() Then
	Set Ceiling = New BigDecimal
	Ceiling = Me
    ElseIf Me.IsPos() Then
	Set Ceiling = Me.Trunc().Plus(BigOne)
    Else    'is negative
	Set Ceiling = Me.Trunc()
    End If

    DebugPrint Ceiling, "Ceiling"
End Function

'***************************************
'
'               Truncate
'
'***************************************

Public Function Trunc() As BigDecimal
    Dim Res_Value As String

    Set Trunc = New BigDecimal

    If Big_lScale = 0 Then
	Trunc.SetRaw Big_sValue, Big_lSign, Big_lInt, Big_lDot, Big_lFrac, Big_lScale
	DebugPrint Trunc, "Trunc"
	Exit Function
    End If

    Res_Value = Mid$(Big_sValue, 1, Big_lDot - 1)

    Trunc.SetRaw Res_Value, Big_lSign, Big_lInt, Big_lDot, Big_lDot, 0

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

	Yield.Events
    Wend
    Res_Int = Res_Dot - 1

    Frac.SetRaw Res_Value, Big_lSign, Res_Int, Res_Dot, Big_lFrac, Big_lScale

    DebugPrint Frac, "Frac"
End Function

'***************************************
'
'          Shift Decimal Point
'
'***************************************

Public Function Shift(ByVal Value As Long) As BigDecimal
    Dim Res_Value As String, Result_Value As String
    Dim Result_Dot As Long, Res_Dot As Long, Pad As Long

    Set Shift = New BigDecimal

    If IsZero() Or Value = 0 Then
	Shift = Me
	Exit Function
    End If

    If Big_lScale > 0 Then
	Res_Value = Replace(Big_sValue, ".", "")
	Res_Dot = Big_lDot - 1
    Else
	Res_Value = Big_sValue
	Res_Dot = Len(Res_Value)
    End If

    If Big_lSign < 0 Then
	Res_Value = Mid$(Res_Value, 2)
	Res_Dot = Res_Dot - 1
    End If

    Result_Dot = Res_Dot + Value

    If Result_Dot <= 0 Then                     'pad left
	Pad = Abs(Result_Dot)
	Result_Value = "0." & String$(Pad, "0") & Res_Value
    ElseIf Result_Dot >= Len(Res_Value) Then    'pad right
	Pad = Result_Dot - Len(Res_Value)
	Result_Value = Res_Value & String$(Pad, "0")
    Else                                        'no pad
	Result_Value = Left$(Res_Value, Result_Dot) & "." & Mid$(Res_Value, Result_Dot + 1)
    End If

    If Big_lSign < 0 Then
	Result_Value = "-" & Result_Value
    End If

    Shift = Result_Value

    DebugPrint Shift, "Shift"
End Function

'***************************************
'
'                Round
'
'***************************************

Public Function Round(ByVal Tar_Scale As Long, Optional ByVal Add_Trailing_Zeroes = False, Optional ByVal Round_Last As Boolean = True) As BigDecimal
    Dim Res_Sign As Long, Res_Int As Long, Res_Dot As Long, Res_Frac As Long, Res_Scale As Long
    Dim Big_sBytes() As Byte, Res_Bytes() As Byte
    Dim i As Long, kI As Long
    Dim Res_Value As String

    Set Round = New BigDecimal

    Res_Bytes = Big_sValue
    Res_Sign = Big_lSign
    Res_Int = Big_lInt
    Res_Dot = Big_lDot
    Res_Frac = Big_lFrac
    Res_Scale = Big_lScale

    While Res_Bytes((Res_Int - 1) * 2) = Asc_0 And Res_Int + 1 < Res_Dot    'find first non-zero digit except before '.'
	Res_Int = Res_Int + 1

	Yield.Events
    Wend

    Debug.Assert Res_Int > 0 And Res_Scale >= 0 And Tar_Scale >= 0

    If Res_Sign < 0 Then       'temporarily remove the sign
	RemoveCh Res_Bytes
	Res_Int = Res_Int - 1
	Res_Dot = Res_Dot - 1
	Res_Frac = Res_Frac - 1
    End If

    If Round_Last Then
	If Res_Scale > Tar_Scale Then   'have something to round
	    While Tar_Scale > 0 And Res_Bytes((Res_Frac + Tar_Scale - 1 - 1) * 2) = Asc_9 And Res_Bytes((Res_Frac + Tar_Scale - 1) * 2) >= Asc_5    'reduce tar_scale when ...9999[5-9]
		Tar_Scale = Tar_Scale - 1

		Yield.Events
	    Wend
	    'will be at a digit < 9

	    Res_Scale = Tar_Scale   'start rounding from tar_scale + 1
	    If Res_Bytes((Res_Frac + Tar_Scale - 1) * 2) >= Asc_5 Then     'round up
		If Tar_Scale > 0 Then       'want some decimal places
		    Res_Bytes((Res_Frac + Tar_Scale - 1 - 1) * 2) = Res_Bytes((Res_Frac + Tar_Scale - 1 - 1) * 2) + 1  'add 1
		Else                        'want no decimal places
		    Res_Frac = Res_Frac - 1
		    Debug.Assert Res_Frac = Res_Dot

		    For i = 0 To (UBound(Res_Bytes) - LBound(Res_Bytes) + 1) / 2 - Res_Dot - 1
			If Res_Dot - i - 1 = 0 Then
			    PrependCh Res_Bytes, "0"
			    Res_Dot = Res_Dot + 1
			    Res_Frac = Res_Frac + 1
			End If

			If Res_Bytes((Res_Dot - i - 1 - 1) * 2) = Asc_9 Then
			    Res_Bytes((Res_Dot - i - 1 - 1) * 2) = Asc_0       '9 rounds to 0 carry 1
			Else
			    Res_Bytes((Res_Dot - i - 1 - 1) * 2) = Res_Bytes((Res_Dot - i - 1 - 1) * 2) + 1    'add carry and finish
			    Exit For
			End If

			Yield.Events
		    Next i

		    If Res_Dot - i - 1 < Res_Int Then
			Res_Int = Res_Dot - i - 1
		    End If
		End If
	    End If
	End If
    Else
	If Res_Scale > Tar_Scale Then
	    Res_Scale = Tar_Scale
	End If
    End If

    If Res_Scale = 0 And Res_Frac > Res_Dot Then    'fix fracton pos
	Res_Frac = Res_Dot
    End If

    Res_Value = Res_Bytes
    Res_Value = Left$(Res_Value, Res_Dot) & Mid$(Res_Value, Res_Frac, Res_Scale)

    If Res_Scale <> Tar_Scale And Add_Trailing_Zeroes Then
	If Res_Scale = 0 Then                   'add trailing '.'
	    Res_Value = Res_Value & "."
	End If
	For kI = 1 To Tar_Scale - Res_Scale     'add trailing '0's
	    Res_Value = Res_Value & "0"
	    Yield.Events
	Next kI
	Res_Scale = Tar_Scale
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

    Round.SetRaw Res_Value, Res_Sign, Res_Int, Res_Dot, Res_Frac, Res_Scale

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
    Dim pos As Long, spos As Long, Num As Long

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

	Num = Res_Int - pos
	If Num > 0 Then
	    Res_Value = "-" & Mid$(Res_Value, pos + Num)

	    Res_Int = Res_Int - Num
	    Res_Dot = Res_Dot - Num
	    Res_Frac = Res_Frac - Num
	End If
    Else
	Num = Res_Int - pos
	If Num > 0 Then
	    Res_Value = Mid$(Res_Value, pos + Num)

	    Res_Int = Res_Int - Num
	    Res_Dot = Res_Dot - Num
	    Res_Frac = Res_Frac - Num
	End If
    End If

    If Res_Scale > 0 Then
	Num = 0
	spos = Res_Frac + Res_Scale - 1
	While Mid$(Res_Value, spos, 1) = "0"
	    Num = Num + 1
	    spos = spos - 1

	    Yield.Events
	Wend
	If Num > 0 Then
	    Res_Value = Left$(Res_Value, Len(Res_Value) - Num)
	    Res_Scale = Res_Scale - Num
	End If

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

    Trim.SetRaw Res_Value, Res_Sign, Res_Int, Res_Dot, Res_Frac, Res_Scale

    DebugPrint Trim, "Trim"
End Function

'***************************************
'
'               Is zero
'
'***************************************

Public Function IsZero() As Boolean
    IsZero = (Me.CmpTo(BigZero) = 0)    'Me = 0
End Function

'***************************************
'
'             Is negative
'
'***************************************

Public Function IsNeg() As Boolean
    IsNeg = (Me.CmpTo(BigZero) < 0)    'Me < 0
End Function

'***************************************
'
'             Is positive
'
'***************************************

Public Function IsPos() As Boolean
    IsPos = (Me.CmpTo(BigZero) > 0)    'Me > 0
End Function

'***************************************
'
'               Is equal
'
'***************************************

Public Function IsEq(Value As Variant) As Boolean
    IsEq = (Me.CmpTo(Value) = 0)     'Me = Value
End Function

'***************************************
'
'             Is not equal
'
'***************************************

Public Function IsNEq(Value As Variant) As Boolean
    IsNEq = (Me.CmpTo(Value) <> 0)   'Me <> Value
End Function

'***************************************
'
'           Is greater than
'
'***************************************

Public Function IsGT(Value As Variant) As Boolean
    IsGT = (Me.CmpTo(Value) > 0)     'Me > Value
End Function

'***************************************
'
'             Is less than
'
'***************************************

Public Function IsLT(Value As Variant) As Boolean
    IsLT = (Me.CmpTo(Value) < 0)     'Me < Value
End Function

'***************************************
'
'        Is greater than or equal
'
'***************************************

Public Function IsGE(Value As Variant) As Boolean
    IsGE = (Me.CmpTo(Value) >= 0)    'Me >= Value
End Function

'***************************************
'
'         Is less than or equal
'
'***************************************

Public Function IsLE(Value As Variant) As Boolean
    IsLE = (Me.CmpTo(Value) <= 0)    'Me <= Value
End Function

'***************************************
'
'              Compare To
'
'***************************************

Public Function CmpTo(Value As Variant) As Long
    Dim Other As BigDecimal

    Set Other = Cast(Value)

    If Big_lSign <> Other.lSign Then
	CmpTo = IIf(Big_lSign < 0, -1, 1)

	If OptionOutputDebug Then Debug.Print Big_sValue & ".CmpTo(" & Other.sValue & ") = " & CmpTo & vbLf
	Exit Function
    End If

    CmpTo = Compare_Positive(Other)
    If Big_lSign < 0 Then
	CmpTo = -CmpTo
    End If

    If OptionOutputDebug Then Debug.Print Big_sValue & ".CmpTo(" & Other.sValue & ") = " & CmpTo & vbLf
End Function

'####################### Private Methods #############################

Private Sub DebugPrint(Other As BigDecimal, oName As String)
    Yield.Events

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

    If Left$(Big_sValue, 1) = "." Then       'fix missing '0' before '.'
	Big_sValue = "0" & Big_sValue
    End If

    If Left$(Big_sValue, 2) = "-." Then      'fix missing '0' between '-' and '.'
	Big_sValue = "-0." & Mid$(Big_sValue, 3)
    End If

    If Left$(Big_sValue, 2) = "+." Then      'fix missing '0' between '+' and '.'
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

    'find '.'
    While pos <= Length And "0" <= Mid$(Big_sValue, pos, 1) And Mid$(Big_sValue, pos, 1) <= "9"
	pos = pos + 1

	Yield.Events
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

	Yield.Events
    Wend

    If pos <= Length Then     'catch non-digits
	Parse = -1
	Exit Function
    End If

    While Mid$(Big_sValue, Big_lInt, 1) = "0" And Big_lInt + 1 < Big_lDot    'find first non-zero digit except before '.'
	Big_lInt = Big_lInt + 1

	Yield.Events
    Wend

    While Big_lScale > 0 And Mid$(sValue, Big_lFrac + Big_lScale - 1, 1) = "0"    'trim trailing '0's
	Big_lScale = Big_lScale - 1

	Yield.Events
    Wend
    Length = Big_lDot + Big_lScale
    Big_sValue = Mid$(Big_sValue, 1, Length)

    If Big_lScale = 0 And Big_lFrac > Big_lDot Then
	Big_lFrac = Big_lFrac - 1
	Debug.Assert Big_lFrac = Big_lDot
    End If

    If Big_lSign < 0 And (Big_lScale = 0 And Mid$(Big_sValue, Big_lInt, 1) = "0") Then   'fix '-0' sign
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
    Set Cast = New BigDecimal

    If IsObject(Other) Then
	Select Case TypeName(Other)
	    Case "BigDecimal"
		Set Cast = Other    'fast pointer copy
	    Case "Nothing"
		Err.Raise vbObjectError + 1107, "BigDecimal", "Can't cast Nothing to BigDecimal."

	    Case Else
		Err.Raise vbObjectError + 1108, "BigDecimal", "Can't cast object " & TypeName(Other) & " to BigDecimal."
	End Select
    Else
	Select Case VarType(Other)
	    Case vbInteger, vbLong, vbSingle, vbDouble, vbString
		Cast.sValue = CStr(Other)
		If Cast.Parse() = -1 Then
		    Err.Raise vbObjectError + 1109, "BigDecimal", TypeName(Other) & " is a malformed BigDecimal number."
		End If

	    Case Else
		Err.Raise vbObjectError + 1110, "BigDecimal", "Can't cast " & TypeName(Other) & " to a BigDecimal."
	End Select
    End If
End Function

Private Function Plus_Positive(Other As BigDecimal, Final_Sign As Long) As BigDecimal
    Dim Oth_Sign As Long, Oth_Int As Long, Oth_Dot As Long, Oth_Frac As Long, Oth_Scale As Long
    Dim Res_Int As Long, Res_Dot As Long, Res_Frac As Long, Res_Scale As Long
    Dim Result_Len As Long, Result_Scale As Long, Result_Size As Long
    Dim Big_sBytes() As Byte, Other_Bytes() As Byte, Res_Bytes() As Byte
    Dim lLen As Long, oLen As Long
    Dim Oth_Value As String
    Dim i As Long, um As Long
    Dim pos As Long, Was_Frac As Long

    Set Plus_Positive = New BigDecimal

    Other.GetRaw Oth_Value, Oth_Sign, Oth_Int, Oth_Dot, Oth_Frac, Oth_Scale

    lLen = Big_lDot - Big_lInt
    oLen = Oth_Dot - Oth_Int

    Result_Len = Application.WorksheetFunction.Max(lLen, oLen) + 1
    Result_Scale = Application.WorksheetFunction.Max(Big_lScale, Oth_Scale)
    Result_Size = Result_Len + Result_Scale + 3

    um = 0
    pos = Result_Size
    Was_Frac = 0

    Big_sBytes = Big_sValue
    Other_Bytes = Oth_Value
    Res_Bytes = String$(Result_Size, "0")

    For i = Result_Scale - 1 To 0 Step -1
	If i < Big_lScale Then
	    um = um + Big_sBytes((Big_lFrac + i - 1) * 2) - Asc_0
	End If
	If i < Oth_Scale Then
	    um = um + Other_Bytes((Oth_Frac + i - 1) * 2) - Asc_0
	End If

	If um <> 0 Or Was_Frac Then
	    pos = pos - 1
	    Res_Bytes(pos * 2) = (um Mod 10) + Asc_0
	    um = um \ 10
	    Was_Frac = 1
	End If

	Yield.Events
    Next i

    Res_Scale = Result_Size - pos
    Res_Frac = pos + 1

    If Was_Frac Then
	pos = pos - 1
	Res_Bytes(pos * 2) = Asc_Dot
    End If

    Res_Dot = pos + 1

    For i = 0 To Result_Len - 1
	If i < lLen Then
	    um = um + Big_sBytes((Big_lDot - i - 2) * 2) - Asc_0
	End If
	If i < oLen Then
	    um = um + Other_Bytes((Oth_Dot - i - 2) * 2) - Asc_0
	End If

	pos = pos - 1
	Res_Bytes(pos * 2) = (um Mod 10) + Asc_0
	um = um \ 10

	Yield.Events
    Next i

    Res_Int = pos
    Debug.Assert pos > 0
    While Res_Bytes((Res_Int - 1) * 2) = Asc_0 And Res_Int + 1 < Res_Dot
	Res_Int = Res_Int + 1

	Yield.Events
    Wend

    If Final_Sign < 0 Then
	PrependCh Res_Bytes, "-"
	Res_Int = Res_Int + 1
	Res_Dot = Res_Dot + 1
	Res_Frac = Res_Frac + 1
    End If

    Plus_Positive.SetRawB Res_Bytes, Final_Sign, Res_Int, Res_Dot, Res_Frac, Res_Scale

    DebugPrint Plus_Positive, "Plus_Positive"
End Function

Friend Function Minus_Positive(Other As BigDecimal, Final_Sign As Long) As BigDecimal
    Dim Oth_Sign As Long, Oth_Int As Long, Oth_Dot As Long, Oth_Frac As Long, Oth_Scale As Long
    Dim lLen As Long, oLen As Long
    Dim Res_Int As Long, Res_Dot As Long, Res_Frac As Long, Res_Scale As Long
    Dim Result_Len As Long, Result_Scale As Long, Result_Size As Long
    Dim Big_sBytes() As Byte, Other_Bytes() As Byte, Res_Bytes() As Byte
    Dim i As Long, um As Long, next_um As Long
    Dim Oth_Value As String
    Dim pos As Long
    Dim Was_Frac As Long

    Set Minus_Positive = New BigDecimal

    Other.GetRaw Oth_Value, Oth_Sign, Oth_Int, Oth_Dot, Oth_Frac, Oth_Scale

    lLen = Big_lDot - Big_lInt
    oLen = Oth_Dot - Oth_Int

    Result_Len = lLen
    Result_Scale = Application.WorksheetFunction.Max(Big_lScale, Oth_Scale)
    Result_Size = Result_Len + Result_Scale + 3

    um = 0
    next_um = 0
    pos = Result_Size
    Was_Frac = 0

    Big_sBytes = Big_sValue
    Other_Bytes = Oth_Value
    Res_Bytes = String$(Result_Size, "0")

    For i = Result_Scale - 1 To 0 Step -1
	um = next_um

	If i < Big_lScale Then
	    um = um + Big_sBytes((Big_lFrac + i - 1) * 2) - Asc_0
	End If

	If i < Oth_Scale Then
	    um = um - Other_Bytes((Oth_Frac + i - 1) * 2) + Asc_0
	End If

	If um < 0 Then
	    next_um = -1
	    um = um + 10
	Else
	    next_um = 0
	End If

	If um <> 0 Or Was_Frac = 1 Then
	    Res_Bytes((pos - 1) * 2) = um + Asc_0
	    pos = pos - 1
	    Was_Frac = 1
	End If

	Yield.Events
    Next i

    Res_Scale = Result_Size - pos
    Res_Frac = pos + 1

    If Was_Frac = 1 Then
	Res_Bytes((pos - 1) * 2) = Asc_Dot
	pos = pos - 1
    End If

    Res_Dot = pos + 1

    For i = 0 To Result_Len - 1
	um = next_um
	um = um + Big_sBytes((Big_lDot - i - 2) * 2) - Asc_0

	If i < oLen Then
	    um = um - Other_Bytes((Oth_Dot - i - 2) * 2) + Asc_0
	End If

	If um < 0 Then
	    next_um = -1
	    um = um + 10
	Else
	    next_um = 0
	End If

	Res_Bytes((pos - 1) * 2) = um + Asc_0
	pos = pos - 1

	Yield.Events
    Next i

    Res_Int = pos
    Debug.Assert pos > 0
    While Res_Bytes((Res_Int - 1) * 2) = Asc_0 And Res_Int + 1 < Res_Dot
	Res_Int = Res_Int + 1

	Yield.Events
    Wend

    If Final_Sign < 0 Then
	PrependCh Res_Bytes, "-"
	Res_Int = Res_Int + 1
	Res_Dot = Res_Dot + 1
	Res_Frac = Res_Frac + 1
    End If

    Minus_Positive.SetRawB Res_Bytes, Final_Sign, Res_Int, Res_Dot, Res_Frac, Res_Scale

    DebugPrint Minus_Positive, "Minus_Positive"
End Function

Private Function Times_Positive(Other As BigDecimal, Final_Sign As Long) As BigDecimal
    Dim Oth_Sign As Long, Oth_Int As Long, Oth_Dot As Long, Oth_Frac As Long, Oth_Scale As Long
    Dim lLen As Long, oLen As Long, i As Long, j As Long, X As Long, Y As Long
    Dim Res_Int As Long, Res_Dot As Long, Res_Frac As Long, Res_Scale As Long
    Dim Result_Len As Long, Result_Scale As Long, Result_Size As Long
    Dim Big_sBytes() As Byte, Other_Bytes() As Byte, Res_Bytes() As Byte
    Dim Res() As Long, pos As Long, K As Long
    Dim Oth_Value As String

    Set Times_Positive = New BigDecimal

    Other.GetRaw Oth_Value, Oth_Sign, Oth_Int, Oth_Dot, Oth_Frac, Oth_Scale

    lLen = Big_lDot - Big_lInt
    oLen = Oth_Dot - Oth_Int

    Result_Len = lLen + oLen
    Result_Scale = Big_lScale + Oth_Scale
    Result_Size = Result_Len + Result_Scale + 3

    ReDim Res(Result_Size - 1)

    Big_sBytes = Big_sValue
    Other_Bytes = Oth_Value
    Res_Bytes = String$(Result_Size, "0")

    For i = -Big_lScale To lLen - 1
	If i < 0 Then
	    X = Big_sBytes((Big_lFrac - i - 2) * 2) - Asc_0
	Else
	    X = Big_sBytes((Big_lDot - i - 2) * 2) - Asc_0
	End If

	For j = -Oth_Scale To oLen - 1
	    If j < 0 Then
		Y = Other_Bytes((Oth_Frac - j - 2) * 2) - Asc_0
	    Else
		Y = Other_Bytes((Oth_Dot - j - 2) * 2) - Asc_0
	    End If
	    Res(i + j + Result_Scale) = Res(i + j + Result_Scale) + (X * Y)

	    Yield.Events
	Next j

	Yield.Events
    Next i

    For i = 0 To Result_Size - 2
	Res(i + 1) = Res(i + 1) + (Res(i) \ 10)
	Res(i) = Res(i) Mod 10

	Yield.Events
    Next i

    pos = Result_Size + 1
    For K = 0 To Result_Scale - 1
	pos = pos - 1
	Res_Bytes((pos - 1) * 2) = Res(K) + Asc_0

	Yield.Events
    Next K

    Res_Scale = Result_Size - pos + 1
    Res_Frac = pos

    If Result_Scale > 0 Then
	pos = pos - 1
	Res_Bytes((pos - 1) * 2) = Asc_Dot
    End If
    Res_Dot = pos

    For i = Result_Scale To Result_Len + Result_Scale - 1
	pos = pos - 1
	Res_Bytes((pos - 1) * 2) = Res(i) + Asc_0

	Yield.Events
    Next i

    Res_Int = pos
    Debug.Assert pos > 0
    While Res_Bytes((Res_Int - 1) * 2) = Asc_0 And Res_Int + 1 < Res_Dot
	Res_Int = Res_Int + 1

	Yield.Events
    Wend

    If Final_Sign < 0 Then
	PrependCh Res_Bytes, "-"
	Res_Int = Res_Int + 1
	Res_Dot = Res_Dot + 1
	Res_Frac = Res_Frac + 1
    End If

    Times_Positive.SetRawB Res_Bytes, Final_Sign, Res_Int, Res_Dot, Res_Frac, Res_Scale

    DebugPrint Times_Positive, "Times_Positive"
End Function

Private Function Divide_Positive(Other As BigDecimal, Final_Sign As Long) As BigDecimal
    Dim Oth_Sign As Long, Oth_Int As Long, Oth_Dot As Long, Oth_Frac As Long, Oth_Scale As Long
    Dim lLen As Long, oLen As Long, Final_Scale As Long
    Dim Res_Int As Long, Res_Dot As Long, Res_Frac As Long, Res_Scale As Long
    Dim Result_Len As Long, Result_Scale As Long, Result_Size As Long
    Dim Big_sBytes() As Byte, Other_Bytes() As Byte, Res_Bytes() As Byte
    Dim dividend_len As Long, divider_len As Long, divider_skip As Long
    Dim i As Long, j As Long, X As Long, Idx As Long
    Dim cur_pow As Long, pos As Long
    Dim beg As Long, real_beg As Long
    Dim Oth_Value As String
    Dim dig As String
    Dim less As Boolean

    Set Divide_Positive = New BigDecimal

    Other.GetRaw Oth_Value, Oth_Sign, Oth_Int, Oth_Dot, Oth_Frac, Oth_Scale

    lLen = Big_lDot - Big_lInt
    oLen = Oth_Dot - Oth_Int

    Final_Scale = Big_lMaxScale + 1

    Result_Len = Application.WorksheetFunction.Max(lLen + Oth_Scale - oLen + 1, 1)
    Result_Scale = Final_Scale
    Result_Size = Result_Len + Result_Scale + 3

    Big_sBytes = Big_sValue
    Other_Bytes = Oth_Value

    If Oth_Scale = 0 And Other_Bytes((Oth_Int - 1) * 2) = Asc_0 Then
	Err.Raise vbObjectError + 1111, "BigDecimal", "Division By ZERO is undefined."
    End If

    dividend_len = lLen + Big_lScale
    divider_len = oLen + Oth_Scale

    ReDim dividend(0 To Result_Size + dividend_len + divider_len - 1) As Long
    ReDim divider(0 To divider_len - 1) As Long

    For i = -Big_lScale To lLen - 1
	If i < 0 Then
	    X = Big_sBytes((Big_lFrac - i - 2) * 2) - Asc_0
	Else
	    X = Big_sBytes((Big_lDot - i - 2) * 2) - Asc_0
	End If
	dividend(lLen - i - 1) = X

	Yield.Events
    Next i

    For i = -Oth_Scale To oLen - 1
	If i < 0 Then
	    X = Other_Bytes((Oth_Frac - i - 2) * 2) - Asc_0
	Else
	    X = Other_Bytes((Oth_Dot - i - 2) * 2) - Asc_0
	End If
	divider(oLen - i - 1) = X

	Yield.Events
    Next i

    Idx = 0
    divider_skip = 0
    While divider_len > 0 And divider(Idx) = 0
	Idx = Idx + 1
	divider_skip = divider_skip + 1
	divider_len = divider_len - 1

	Yield.Events
    Wend
    Debug.Assert divider_len > 0

    cur_pow = lLen - oLen + divider_skip
    pos = 2

    If cur_pow < -Final_Scale Then
	Idx = Idx - divider_skip
	divider_len = divider_len + divider_skip
	Exit Function
    End If

    Res_Bytes = String$(Result_Size, "0")

    Res_Int = pos
    If cur_pow < 0 Then
	Res_Bytes((pos - 1) * 2) = Asc_0
	pos = pos + 1
	Res_Dot = pos
	Res_Bytes((pos - 1) * 2) = Asc_Dot
	pos = pos + 1
	Res_Frac = pos
	j = -1
	While j > cur_pow
	    Res_Bytes((pos - 1) * 2) = Asc_0
	    pos = pos + 1
	    j = j - 1

	    Yield.Events
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
		    If dividend(beg + i) <> divider(i + Idx) Then
			less = (dividend(beg + i) < divider(i + Idx))
			Exit For
		    End If

		    Yield.Events
		Next i
	    End If

	    If less Then
		Exit Do
	    End If

	    For i = divider_len - 1 To 0 Step -1
		dividend(beg + i) = dividend(beg + i) - divider(i + Idx)
		If dividend(beg + i) < 0 Then
		    dividend(beg + i) = dividend(beg + i) + 10
		    dividend(beg + i - 1) = dividend(beg + i - 1) - 1
		End If

		Yield.Events
	    Next i
	    dig = Chr$(Asc(dig) + 1)

	    Yield.Events
	Loop

	Res_Bytes((pos - 1) * 2) = Asc(dig)
	pos = pos + 1

	If cur_pow = 0 Then
	    Res_Dot = pos
	    If Final_Scale > 0 Then
		Res_Bytes((pos - 1) * 2) = Asc_Dot
		pos = pos + 1
	    End If
	    Res_Frac = pos
	End If

	cur_pow = cur_pow - 1
	beg = beg + 1
	Yield.Events
    Wend

    Res_Scale = Result_Size - Res_Frac + 1

    Idx = Idx - divider_skip
    divider_len = divider_len + divider_skip

    While Res_Bytes((Res_Int - 1) * 2) = Asc_0 And Res_Int + 1 < Res_Dot
	Res_Int = Res_Int + 1

	Yield.Events
    Wend

    If Final_Sign < 0 Then
	PrependCh Res_Bytes, "-"
	Res_Int = Res_Int + 1
	Res_Dot = Res_Dot + 1
	Res_Frac = Res_Frac + 1
    End If

    Divide_Positive.SetRawB Res_Bytes, Final_Sign, Res_Int, Res_Dot, Res_Frac, Res_Scale

    DebugPrint Divide_Positive, "divide_positive"
End Function

Private Function Compare_Positive(Other As BigDecimal) As Long
    Dim Oth_Sign As Long, Oth_Int As Long, Oth_Dot As Long, Oth_Frac As Long, Oth_Scale As Long
    Dim lLen As Long, oLen As Long, i As Long
    Dim lchar As Long, ochar As Long, Res_Scale As Long
    Dim Big_sBytes() As Byte, Other_Bytes() As Byte
    Dim Oth_Value As String

    Other.GetRaw Oth_Value, Oth_Sign, Oth_Int, Oth_Dot, Oth_Frac, Oth_Scale

    Res_Scale = Big_lMaxScale

    lLen = Big_lDot - Big_lInt
    oLen = Oth_Dot - Oth_Int

    Big_sBytes = Big_sValue
    Other_Bytes = Oth_Value

    If lLen <> oLen Then
	Compare_Positive = IIf(lLen < oLen, -1, 1)
	Exit Function
    End If

    For i = 0 To lLen - 1
	If Big_sBytes((Big_lInt + i - 1) * 2) <> Other_Bytes((Oth_Int + i - 1) * 2) Then
	    Compare_Positive = IIf(Big_sBytes((Big_lInt + i - 1) * 2) < Other_Bytes((Oth_Int + i - 1) * 2), -1, 1)
	    Exit Function
	End If

	Yield.Events
    Next i

    For i = 0 To Res_Scale - 1
	If i < Big_lScale Then
	    lchar = Big_sBytes((Big_lFrac + i - 1) * 2) - Asc_0
	Else
	    lchar = 0
	End If

	If i < Oth_Scale Then
	    ochar = Other_Bytes((Oth_Frac + i - 1) * 2) - Asc_0
	Else
	    ochar = 0
	End If

	If lchar <> ochar Then
	    Compare_Positive = IIf(lchar < ochar, -1, 1)
	    Exit Function
	End If

	Yield.Events
    Next i

    Compare_Positive = 0
End Function

Friend Sub SetRaw(sValue As String, lSign As Long, lInt As Long, lDot As Long, lFrac As Long, lScale As Long)
    Big_sValue = sValue
    Big_lSign = lSign
    Big_lInt = lInt
    Big_lDot = lDot
    Big_lFrac = lFrac
    Big_lScale = lScale
End Sub

Friend Sub SetRawB(sBytes() As Byte, lSign As Long, lInt As Long, lDot As Long, lFrac As Long, lScale As Long)
    Big_sValue = sBytes
    Big_lSign = lSign
    Big_lInt = lInt
    Big_lDot = lDot
    Big_lFrac = lFrac
    Big_lScale = lScale
End Sub

Friend Sub GetRaw(sValue As String, lSign As Long, lInt As Long, lDot As Long, lFrac As Long, lScale As Long)
    sValue = Big_sValue
    lSign = Big_lSign
    lInt = Big_lInt
    lDot = Big_lDot
    lFrac = Big_lFrac
    lScale = Big_lScale
End Sub

Private Sub PrependCh(ByRef ArrB() As Byte, ByVal Ch As String)
    Dim ChB(0 To 1) As Byte
    Dim Size As Long

    MoveMemory ChB(0), ByVal StrPtr(Ch), 2  'copy Ch

    Size = UBound(ArrB) - LBound(ArrB) + 1  'calc arrb size
    ReDim Preserve ArrB(0 To Size + 1)      'resize arrb
    MoveMemory ArrB(2), ArrB(0), Size       'copy from tail to head

    ArrB(0) = ChB(0)
    ArrB(1) = ChB(1)
End Sub


Private Sub RemoveCh(ByRef ArrB() As Byte)
    Dim Size As Long

    Size = UBound(ArrB) - LBound(ArrB) + 1  'calc arrb size
    If Size <= 2 Then
	Erase ArrB
	Exit Sub
    End If

    Size = Size - 2
    MoveMemory ArrB(0), ArrB(2), Size       'copy from head to tail
    ReDim Preserve ArrB(0 To Size - 1)      'resize arrb
End Sub

Friend Sub Unregister()
    BigOption.Unregister Me
End Sub

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
	    Err.Raise vbObjectError + 1112, "BigDecimal", "BigDecimal can't be nothing."
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
	    Err.Raise vbObjectError + 1113, "BigDecimal", "Can't cast object " & TypeName(Other) & " to BigDecimal."
	End If
    Else
	Select Case VarType(Other)
	    Case vbInteger, vbLong, vbLongLong, vbSingle, vbDouble, vbString
		Big_sValue = CStr(Other)
		If Parse() = -1 Then
		    Err.Raise vbObjectError + 1114, "BigDecimal", TypeName(Other) & " is a malformed BigDecimal number."
		End If
	    Case Else
		Err.Raise vbObjectError + 1115, "BigDecimal", "Can't cast " & TypeName(Other) & " to BigDecimal."
	End Select
    End If
End Property

'ValueEx
Public Property Get ValueEx(Optional RetType As Long = vbString) As Variant
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
	If OptionAutoRound Then
	    If OptionAddTrailingZeroes Then
		ValueEx = CStr(Trim().Round(Big_lTarScale, OptionAddTrailingZeroes).sValue)
	    Else
		ValueEx = CStr(Round(Big_lTarScale).Trim().sValue)
	    End If
	Else
	    ValueEx = CStr(Trim().sValue)
	End If
    Case Else
	Err.Raise vbObjectError + 1116, "BigDecimal", "Unsupported variable type " & RetType
    End Select
End Property

'TarScale
Public Property Get TarScale() As Long
    TarScale = Big_lTarScale
End Property

Public Property Let TarScale(NewScale As Long)
    If NewScale < 0 Then
	Big_lTarScale = 0
	Exit Property
    End If

    Big_lTarScale = NewScale
    Big_lMaxScale = Big_lTarScale + Big_lExtra

    If Big_lScale > Big_lMaxScale Then      'truncate
	Big_lScale = Big_lMaxScale
	Big_sValue = Mid$(Big_sValue, 1, Big_lDot + Big_lScale)
    End If
End Property

'AddTrailingZeroes
Public Property Get AddTrailingZeroes() As Boolean
    AddTrailingZeroes = OptionAddTrailingZeroes
End Property

Public Property Let AddTrailingZeroes(Flag As Boolean)
    OptionAddTrailingZeroes = Flag
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

'Pi
Friend Property Get Pi() As BigDecimal
    Dim Prec As BigDecimal
    Dim Term1 As BigDecimal
    Dim Term2 As BigDecimal

    Set Prec = New BigDecimal

    If Big_lMaxScale <= BigOption.Pi_Scale Then
	Set Pi = New BigDecimal
	Pi = BigOption.Pi_Buf
	If Big_lMaxScale < BigOption.Pi_Scale Then
	    Pi = Pi.Round(Big_lMaxScale)
	End If
	Exit Property
    End If

    Prec = "0." & String$(Big_lMaxScale - 1, "0") & "1"

    Set Term1 = ATanReciprocal(5, Prec).Times(16)
    Set Term2 = ATanReciprocal(239, Prec).Times(4)

    Set Pi = Term1.Minus(Term2)

    BigOption.Pi_Buf = Pi.sValue
    BigOption.Pi_Scale = Pi.lScale

    DebugPrint Pi, "Pi"
End Property

Private Function ATanReciprocal(ByVal X As Long, Prec As BigDecimal) As BigDecimal
    Dim Term As BigDecimal
    Dim Result As BigDecimal
    Dim X2 As BigDecimal
    Dim Y As BigDecimal
    Dim Sign As Long

    Set Result = New BigDecimal
    Set X2 = New BigDecimal
    Set Y = New BigDecimal

    X2 = X * X
    Y = BigOne
    Sign = 1

    Set Term = BigOne.Divide(X)
    Result = Term

    While Term.AbsVal().IsGT(Prec)
	Y = Y.Plus(BigTwo)
	Sign = Sign * -1

	Set Term = Term.Divide(X2)

	If Sign > 0 Then
	    Set Result = Result.Plus(Term.Divide(Y))
	Else
	    Set Result = Result.Minus(Term.Divide(Y))
	End If

	Yield.Events
    Wend

    Set ATanReciprocal = Result
End Function

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
