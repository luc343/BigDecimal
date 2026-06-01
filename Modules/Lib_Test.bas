Attribute VB_Name = "Lib_Test"
Option Explicit

Private TotalTests As Long
Private PassedTests As Long

Dim TargetScale As Long

'---------------------------------------------------
'
'                    Lib_Test
'
' Copyright (c) Lucien Cinc 2025-26
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
'---------------------------------------------------

'***************************************
'
'            Run All Tests
'
'***************************************

Public Sub RunAllTests()
    TargetScale = BigOption.TarScale ' Save current state

    TotalTests = 0
    PassedTests = 0

    Yield.Clear

    Debug.Print " === STARTING AUTOMATED TESTS ====================================="
    Debug.Print

    ' test suites
    Test_Addition
    Test_Subtraction
    Test_Multiplication
    Test_Division
    Test_Logarithms
    Test_ExponentialIdentities
    Test_ComparisonsAndSignChecks
    Test_EdgeCases
    Test_RoundingAndTruncation
    Test_MathBoundaryViolations
    Test_AdvancedMathAndUtilities

    BigOption.TarScale = TargetScale

    Debug.Print
    Debug.Print " =================================================================="
    Debug.Print
    Debug.Print "    TESTS COMPLETE: " & PassedTests & " / " & TotalTests & " PASSED."
    If PassedTests = TotalTests Then
	Debug.Print "    *** ALL GREEN! No features broken."
    Else
	Debug.Print "    >> WARNING: " & (TotalTests - PassedTests) & " TEST(S) FAILED!"
    End If
End Sub

Public Sub AssertEqual(TestName As String, Actual As String, Expected As String)
    Dim DisplayName As String
    Dim MaxLength As Long

    MaxLength = 60
    DisplayName = " [+] " & TestName & " "

    If Len(DisplayName) < MaxLength Then
	DisplayName = DisplayName & String(MaxLength - Len(DisplayName), ".")
    End If

    TotalTests = TotalTests + 1
    If Actual = Expected Then
	PassedTests = PassedTests + 1
	Debug.Print DisplayName & " [PASS]"
    Else
	Debug.Print DisplayName & " [FAIL]"
	Debug.Print "       Expected: " & Expected
	Debug.Print "         Actual: " & Actual
    End If

    Yield.Events
End Sub

Private Sub Test_Addition()
    Dim A As New BigDecimal, B As New BigDecimal
    Dim Expected As String

    ' Test: Basic decimal addition
    A = "1.234"
    B = "5.678"
    AssertEqual "Basic Addition", A.Plus(B).ValueEx(vbString), "6.912"

    ' Test: Carrying over 10
    A = "9.999"
    B = "0.001"
    AssertEqual "Addition Carry-over", A.Plus(B).ValueEx(vbString), "10"

    ' Test: Massive asymmetry test
    ' Verification that addition doesn't misalign the huge integer or the tiny fraction
    A = "123456789012345678901234567890.000000000000000000000000000001"
    B = "2"
    Expected = "123456789012345678901234567892.000000000000000000000000000001"
    AssertEqual "Massive Asymmetric Addition", A.Plus(B).ValueEx(vbString), Expected
End Sub

Private Sub Test_Subtraction()
    Dim Num1 As New BigDecimal, Num2 As New BigDecimal
    Dim Expected As String

    ' Test: Forces a borrow to ripple all the way from the hundreds place
    ' through a wall of zeros down to the 20th decimal place.
    Num1 = "100.0"
    Num2 = "0.00000000000000000001"
    Expected = "99.99999999999999999999"
    AssertEqual "Cascading Ripple Borrow", Num1.Minus(Num2).ValueEx(vbString), Expected

    ' Test: Forcing the engine to flip its internal sign state to negative
    Num1 = "5.5"
    Num2 = "10.7"
    AssertEqual "Sign Reversal Subtraction", Num1.Minus(Num2).ValueEx(vbString), "-5.2"

    ' Test: High-precision variant of sign reversal
    Num1 = "0.00000000000000000001"
    Num2 = "1.0"
    Expected = "-0.99999999999999999999"
    AssertEqual "High-Precision Sign Reversal", Num1.Minus(Num2).ValueEx(vbString), Expected

    ' Test: Mathematically: A - (-B) = A + B
    Num1 = "50.5"
    Num2 = "-10.2"
    AssertEqual "Double Negative Addition Route", Num1.Minus(Num2).ValueEx(vbString), "60.7"

    ' Test: Mathematically: -A - B = -(A + B)
    Num1 = "-25.25"
    Num2 = "75.50"
    AssertEqual "Subtracting From Negative", Num1.Minus(Num2).ValueEx(vbString), "-100.75"

    ' Test: Verifies that arrays cancel out completely and leave a clean, neutral "0"
    ' instead of a broken residual like "-0.00000" or a signed negative zero.
    Num1 = "123456789.987654321"
    Num2 = "123456789.987654321"
    AssertEqual "Perfect Zero Annihilation", Num1.Minus(Num2).ValueEx(vbString), "0"
End Sub

Private Sub Test_Multiplication()
    Dim A As New BigDecimal, B As New BigDecimal

    ' Test: Massive mismatch of scales (forces alignment shift)
    A = "1000000000000"
    B = "0.0000000000001"
    AssertEqual "Scale Alignment Multiplication", A.Times(B).ValueEx(vbString), "0.1"

    ' Test: Nested Carries (9s multiplying 9s cascade carries all the way up)
    A = "99.99"
    B = "99.99"
    AssertEqual "Cascading Carry Multiplication", A.Times(B).ValueEx(vbString), "9998.0001"
End Sub

Private Sub Test_Division()
    Dim Num1 As New BigDecimal, Num2 As New BigDecimal
    Dim Expected As String
    Dim OrgTarScale As Long

    OrgTarScale = BigOption.TarScale

    ' Test: Exact interger division
    Num1 = "100"
    Num2 = "4"
    AssertEqual "Exact Integer Division", Num1.Divide(Num2).ValueEx(vbString), "25"

    ' Test: Exact termination decimals
    Num1 = "1"
    Num2 = "8"
    AssertEqual "Terminating Decimal Division", Num1.Divide(Num2).ValueEx(vbString), "0.125"

    ' Testing: Truncate/round infinite loops.
    ' 1 / 3 = 0.3333333333...
    BigOption.TarScale = 30
    Num1 = "1"
    Num2 = "3"
    Expected = "0.333333333333333333333333333333"
    AssertEqual "Repeating Decimal (1/3)", Num1.Divide(Num2).ValueEx(vbString), Expected

    ' 22 / 7 (Classic Pi approximation check)
    BigOption.TarScale = 25
    Num1 = "22"
    Num2 = "7"
    Expected = "3.1428571428571428571428571"
    AssertEqual "Repeating Decimal Loop (22/7)", Num1.Divide(Num2).ValueEx(vbString), Expected

    ' Test: Dividing small by large (Underflow Hazard)
    ' Forces the engine to accurately pad leading zeros inside the fraction array.
    BigOption.TarScale = 30
    Num1 = "0.0000000001"
    Num2 = "100000"
    Expected = "0.000000000000001"
    AssertEqual "Small Divided by Large", Num1.Divide(Num2).ValueEx(vbString), Expected

    ' Test: Pos / Neg = Neg
    Num1 = "10"
    Num2 = "-2"
    AssertEqual "Positive Divided By Negative", Num1.Divide(Num2).ValueEx(vbString), "-5"

    ' Test: Neg / Neg = Pos
    Num1 = "-15.5"
    Num2 = "-5"
    AssertEqual "Negative Divided By Negative", Num1.Divide(Num2).ValueEx(vbString), "3.1"

    ' Test: Precision & carry rounding boundary
    ' 2 / 3 = 0.666666...66667 (Should round up at the target scale edge!)
    BigOption.TarScale = 20
    Num1 = "2"
    Num2 = "3"
    Expected = "0.66666666666666666667"
    AssertEqual "Division Edge Round-Up (2/3)", Num1.Divide(Num2).ValueEx(vbString), Expected

    BigOption.TarScale = OrgTarScale
End Sub

Private Sub Test_Logarithms()
    Dim X As New BigDecimal
    Dim ExpectedLn2 As String

    ' Test: Log10 of 100 should be 2
    X = "100"
    AssertEqual "Log10(100) = 2", X.Log().ValueEx(vbString), "2"

    ' Test: Verification of Ln2 against a hardcoded known true value (100 digits)
    BigOption.TarScale = 100
    ExpectedLn2 = "0.6931471805599453094172321214581765680755001343602552541206800094933936219696947156058633269964186875"

    ' Test: Explicitly make sure the cache misses
    BigOption.Ln2_Scale = 50
    X = 2
    AssertEqual "High-Precision Ln(2)", X.Ln().ValueEx(vbString), ExpectedLn2
End Sub

Private Sub Test_ExponentialIdentities()
    Dim X As New BigDecimal
    Dim ExpectedExp1 As String

    ' Test: e^0 must equal 1
    X = "0"
    AssertEqual "Exp(0) = 1", X.Exp().ValueEx(vbString), "1"

    ' Test: Round-trip Identity Exp(Ln(5)) should equal 5
    BigOption.TarScale = 50
    X = "5"
    AssertEqual "Identity Exp(Ln(5)) = 5", X.Ln().Exp().ValueEx(vbString), "5"

    ' Test: High Precision Exp(1) against Euler's Constant (50 digits)
    X = "1"
    ExpectedExp1 = "2.71828182845904523536028747135266249775724709369996"
    AssertEqual "High-Precision Exp(1) [e]", X.Exp().ValueEx(vbString), ExpectedExp1
End Sub

Private Sub Test_ComparisonsAndSignChecks()
    Dim Num1 As New BigDecimal, Num2 As New BigDecimal

    ' Test: For Zero
    Num1 = "0.000000"
    AssertEqual "IsZero on True Zero", CStr(Num1.IsZero()), "True"
    AssertEqual "IsPos on True Zero", CStr(Num1.IsPos()), "False"
    AssertEqual "IsNeg on True Zero", CStr(Num1.IsNeg()), "False"

    ' Test: For Positive
    Num1 = "0.000000000000000000000000000001"
    AssertEqual "IsZero on Tiny Pos", CStr(Num1.IsZero()), "False"
    AssertEqual "IsPos on Tiny Pos", CStr(Num1.IsPos()), "True"
    AssertEqual "IsNeg on Tiny Pos", CStr(Num1.IsNeg()), "False"

    ' Test: For Negative
    Num1 = "-123456789.987654321"
    AssertEqual "IsZero on Neg", CStr(Num1.IsZero()), "False"
    AssertEqual "IsPos on Neg", CStr(Num1.IsPos()), "False"
    AssertEqual "IsNeg on Neg", CStr(Num1.IsNeg()), "True"

    ' Test: For IsEq and IsNEq
    ' These have different string lengths, but are mathematically identical!
    Num1 = "100.00500"
    Num2 = "100.005"
    AssertEqual "IsEq (Trailing Zeros Ignored)", CStr(Num1.IsEq(Num2)), "True"
    AssertEqual "IsNEq (Trailing Zeros Ignored)", CStr(Num1.IsNEq(Num2)), "False"

    Num2 = "100.00501"
    AssertEqual "IsEq on Mismatch", CStr(Num1.IsEq(Num2)), "False"
    AssertEqual "IsNEq on Mismatch", CStr(Num1.IsNEq(Num2)), "True"

    ' Setup: Num1 is slightly smaller than Num2
    Num1 = "5.00000000000000000000000000000000000001"
    Num2 = "5.00000000000000000000000000000000000002"

    ' Test: Strictly Less Than / Greater Than
    AssertEqual "IsLT (Strictly Less)", CStr(Num1.IsLT(Num2)), "True"
    AssertEqual "IsGT (Strictly Greater)", CStr(Num1.IsGT(Num2)), "False"

    ' Test: Less-Than-Equal / Greater-Than-Equal (Distinct values)
    AssertEqual "IsLE (Distinct values)", CStr(Num1.IsLE(Num2)), "True"
    AssertEqual "IsGE (Distinct values)", CStr(Num1.IsGE(Num2)), "False"

    ' Edge Case: Check behavior when values are EXACTLY equal
    Num2 = "5.00000000000000000000000000000000000001"

    AssertEqual "IsLT when Equal", CStr(Num1.IsLT(Num2)), "False"
    AssertEqual "IsGT when Equal", CStr(Num1.IsGT(Num2)), "False"
    AssertEqual "IsLE when Equal", CStr(Num1.IsLE(Num2)), "True"
    AssertEqual "IsGE when Equal", CStr(Num1.IsGE(Num2)), "True"
End Sub

Private Sub Test_RoundingAndTruncation()
    Dim A As New BigDecimal

    ' Test: Rounding UP without leading digits
    ' If we force target to add leading Zero before rounding up
    BigOption.TarScale = 10
    A = ".99999999999"
    AssertEqual "Rounding UP without leading digits before dot", A.ValueEx(vbString), "1"

    ' Test: Rounding UP at boundary
    ' If we force target scale to 3, 1.2346 should round UP to 1.235
    BigOption.TarScale = 3
    A = "1.2346"
    AssertEqual "Rounding UP Test", A.ValueEx(vbString), "1.235"

    ' Test: Rounding DOWN at boundary
    A = "1.2344"
    AssertEqual "Rounding DOWN Test", A.ValueEx(vbString), "1.234"

    ' Test: Trailing zeros after a significant digit
    ' Making sure zero-stripping doesn't break on something like 1.0003
    BigOption.TarScale = 4
    A = "1.0003"
    AssertEqual "Embedded Internal Zeros", A.ValueEx(vbString), "1.0003"
End Sub

Private Sub Test_EdgeCases()
    On Error Resume Next
    Dim X As New BigDecimal
    X = "-5"

    ' Test: Does Log base 10 throw an error on negative numbers?
    Err.Clear
    Dim Result As BigDecimal
    Set Result = X.Log()

    If Err.Number = 5 Then
	AssertEqual "Log10 Negative Guard Rail", "Error 5 Thrown", "Error 5 Thrown"
    Else
	AssertEqual "Log10 Negative Guard Rail", "No Error / Wrong Error", "Error 5 Thrown"
    End If
    On Error GoTo 0
End Sub

Private Sub Test_MathBoundaryViolations()
    On Error Resume Next
    Dim Zero As New BigDecimal, One As New BigDecimal, Neg As New BigDecimal
    Dim Result As BigDecimal
    Dim ErrNum As Long

    Zero = "0"
    One = "1"
    Neg = "-1"

    ' Test: Division by zero guard rail
    Err.Clear
    Set Result = One.Divide(Zero)
    If Err.Number <> 0 Then
	AssertEqual "Division by Zero Guard Rail", "Error Caught", "Error Caught"
    Else
	AssertEqual "Division by Zero Guard Rail", "Allowed Unlawful Division", "Error Caught"
    End If

    ' Test: Ln() of 0 (Undefined)
    Err.Clear
    ErrNum = vbObjectError + 1102
    Set Result = Zero.Ln()
    If Err.Number = ErrNum Then
	AssertEqual "Ln(0) Guard Rail", "Error " & ErrNum & " Thrown", "Error " & ErrNum & " Thrown"
    Else
	AssertEqual "Ln(0) Guard Rail", "No Error / Wrong Error", "Error " & ErrNum & " Thrown"
    End If

    ' Test: Ln() of Negative Number (Undefined)
    Err.Clear
    Set Result = Neg.Ln()
    If Err.Number = ErrNum Then
	AssertEqual "Ln(Negative) Guard Rail", "Error " & ErrNum & " Thrown", "Error " & ErrNum & " Thrown"
    Else
	AssertEqual "Ln(Negative) Guard Rail", "No Error / Wrong Error", "Error " & ErrNum & " Thrown"
    End If
    On Error GoTo 0
End Sub

Private Sub Test_AdvancedMathAndUtilities()
    Dim Num1 As New BigDecimal, Num2 As New BigDecimal
    Dim Expected As String

    BigOption.TarScale = 50

    ' Test: Absolute value of a negative number
    Num1 = "-123.456"
    AssertEqual "AbsVal Negative Input", Num1.AbsVal().ValueEx(vbString), "123.456"

    ' Test: Negate a positive number
    Num2 = "123.456"
    AssertEqual "Neg Positive Input", Num2.Neg().ValueEx(vbString), "-123.456"

    ' Test: Negate a Negative number
    Num2 = "-123.456"
    AssertEqual "Neg Negative Input", Num2.Neg().ValueEx(vbString), "123.456"

    ' Test: Truncate and fraction parts of a decimal number
    Num1 = "567.89012"
    AssertEqual "Truncate Pulls Integer", Num1.Trunc().ValueEx(vbString), "567"
    AssertEqual "Frac Pulls Fraction Part", Num1.Frac().ValueEx(vbString), "0.89012"

    ' Test: Negative edge case for Trunc/Frac
    Num2 = "-5.67"
    AssertEqual "Truncate Negative", Num2.Trunc().ValueEx(vbString), "-5"
    AssertEqual "Frac Negative", Num2.Frac().ValueEx(vbString), "-0.67"

    ' Test: Multiple leading zeros (should be stripped, not treated as octal)
    Num1 = "000005.67000"
    AssertEqual "Parser Leading/Trailing Zeros", Num1.ValueEx(vbString), "5.67"

    ' Test: Explicit positive sign
    Num1 = "+89.12"
    AssertEqual "Parser Explicit Plus Sign", Num1.ValueEx(vbString), "89.12"

    ' Test:  explicit precision rounding
    Num1 = "1.234567"
    AssertEqual "Round to 4 places", Num1.Round(4).ValueEx(vbString), "1.2346"

    ' Test: Remainder (Modulo) rule: 10.5 Mod 3 = 1.5
    Num1 = "10.5"
    Num2 = "3"
    AssertEqual "Remainder check", Num1.Remdr(Num2).ValueEx(vbString), "1.5"

    ' Test: 5! = 5 * 4 * 3 * 2 * 1 = 120
    Num1 = "5"
    AssertEqual "Factorial of 5", Num1.Fact().ValueEx(vbString), "120"

    ' Hardcore 30! stress test to make sure large integer carries hold up
    Num1 = "30"
    Expected = "265252859812191058636308480000000"
    AssertEqual "Factorial Large Stress (30!)", Num1.Fact().ValueEx(vbString), Expected

    ' Test: Sqrt of 2 to 50 decimal digits
    BigOption.TarScale = 50
    Num1 = "2"
    Expected = "1.41421356237309504880168872420969807856967187537695"
    AssertEqual "High-Precision Sqrt(2)", Num1.Sqrt().ValueEx(vbString), Expected

    ' Test Case A: Simple Integer Power (3^4 = 81)
    Num1 = "3"
    Num2 = "4"
    AssertEqual "Pow Integer Base/Exp", Num1.Pow(Num2).ValueEx(vbString), "81"

    ' Test Case B: Fractional/Transcendental Power (2.5 ^ 3.5)
    ' Evaluated internally via: Exp(3.5 * Ln(2.5))
    BigOption.TarScale = 40
    Num1 = "2.5"
    Num2 = "3.5"
    Expected = "24.705294220065463531241355815880613544684"
    AssertEqual "Pow Fractional Base/Exp", Num1.Pow(Num2).ValueEx(vbString), Expected
End Sub
