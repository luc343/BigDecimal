Attribute VB_Name = "Lib_Test"
Option Explicit

Private TotalTests As Long
Private PassedTests As Long

Dim TargetScale As Long
Dim Timer As New Timer

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
    BigOption.Stats
    Debug.Print

    TotalTests = 0
    PassedTests = 0

    Timer.Start

    TargetScale = BigOption.TarScale ' Save current state

    BigOption.TarScale = 40

    Yield.Clear

    Debug.Print " ======= STARTING AUTOMATED TESTS ================================="
    Debug.Print

    ' test suites
    Test_Addition
    Test_Subtraction
    Test_Multiplication
    Test_Division
    Test_Logarithms
    Test_Exponential
    Test_Pow10Engine
    Test_RemainderAndModulo
    Test_PiIdentities
    Test_ComparisonsAndSignChecks
    Test_FloorAndCeiling
    Test_GreatestCommonDivisor
    Test_LeastCommonMultiple
    Test_CastingCBgDec
    Test_EdgeCases
    Test_RoundingAndTruncation
    Test_MathBoundaryViolations
    Test_AdvancedMathAndUtilities
    Test_CubeAndCubert
    Test_ShiftDecimal

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

    BigOption.Stats

    Timer.Finish vbLf & "Elaspsed"
End Sub

Public Sub AssertEqual(TestName As String, Actual As String, Expected As String)
    Dim DisplayName As String, Index As String, Expand As String
    Dim MaxLength As Long

    MaxLength = 60
    Index = TotalTests + 1 & " "
    If Len(Index) < 4 Then
	Index = Space$(4 - Len(Index)) & Index
    End If
    If Actual = Expected Then
	Expand = " [+] "
    Else
	Expand = " [-] "
    End If
    DisplayName = Expand & Index & TestName & " "

    If Len(DisplayName) < MaxLength Then
	DisplayName = DisplayName & String(MaxLength - Len(DisplayName), ".")
    End If

    TotalTests = TotalTests + 1
    If Actual = Expected Then
	PassedTests = PassedTests + 1
	Debug.Print DisplayName & " [PASS]"
    Else
	Debug.Print DisplayName & " [FAIL]"
	Debug.Print "           Expected: " & Expected
	Debug.Print "             Actual: " & Actual
    End If

    Yield.Events
End Sub

Private Sub Test_Addition()
    Dim A As BigDecimal, B As BigDecimal
    Dim Expected As String

    Set A = New_BigDecimal()
    Set B = New_BigDecimal()

    ' Test: Basic decimal addition
    A = "1.234"
    B = "5.678"
    AssertEqual "Basic Addition", A.Plus(B).ValueEx, "6.912"

    ' Test: Carrying over 10
    A = "9.999"
    B = "0.001"
    AssertEqual "Addition Carry-over", A.Plus(B).ValueEx, "10"

    ' Test: Positive and negative addition
    A = "05.4561"
    B = "-10.0061"
    AssertEqual "Positive and Negative Addition", A.Plus(B).ValueEx, "-4.55"

    ' Test: Negative and positive addition
    A = "-05.4561"
    B = "+10.0061"
    AssertEqual "Negative and positive Addition", A.Plus(B).ValueEx, "4.55"

    ' Test: Negative and negative addition
    A = "-05.4561"
    B = "-10.0061"
    AssertEqual "Negative and Negative Addition", A.Plus(B).ValueEx, "-15.4622"

    ' Test: Massive asymmetry test
    ' Verification that addition doesn't misalign the huge integer or the tiny fraction
    A = "123456789012345678901234567890.000000000000000000000000000001"
    B = "2"
    Expected = "123456789012345678901234567892.000000000000000000000000000001"
    AssertEqual "Massive Asymmetric Addition", A.Plus(B).ValueEx, Expected

    ' Test: Verifies that arrays cancel out completely and leave a clean, neutral "0"
    ' instead of a broken residual like "-0.00000" or a signed negative zero.
    A = "123456789.987654321"
    B = "-123456789.987654321"
    AssertEqual "Perfect Zero Annihilation Addition", A.Plus(B).ValueEx, "0"
End Sub

Private Sub Test_Subtraction()
    Dim A As BigDecimal, B As BigDecimal
    Dim Expected As String

    Set A = New_BigDecimal()
    Set B = New_BigDecimal()

    ' Test: Forces a borrow to ripple all the way from the hundreds place
    ' through a wall of zeros down to the 20th decimal place.
    A = "100.0"
    B = "0.00000000000000000001"
    Expected = "99.99999999999999999999"
    AssertEqual "Cascading Ripple Borrow", A.Minus(B).ValueEx, Expected

    ' Test: Forcing the engine to flip its internal sign state to negative
    A = "5.5"
    B = "10.7"
    AssertEqual "Sign Reversal Subtraction", A.Minus(B).ValueEx, "-5.2"

    ' Test: High-precision variant of sign reversal
    A = "0.00000000000000000001"
    B = "1.0"
    Expected = "-0.99999999999999999999"
    AssertEqual "High-Precision Sign Reversal", A.Minus(B).ValueEx, Expected

    ' Test: Mathematically: A - (-B) = A + B
    A = "50.5"
    B = "-10.2"
    AssertEqual "Double Negative Addition Route", A.Minus(B).ValueEx, "60.7"

    ' Test: Mathematically: -A - B = -(A + B)
    A = "-25.25"
    B = "75.50"
    AssertEqual "Subtracting From Negative", A.Minus(B).ValueEx, "-100.75"

    ' Test: Verifies that arrays cancel out completely and leave a clean, neutral "0"
    ' instead of a broken residual like "-0.00000" or a signed negative zero.
    A = "123456789.987654321"
    B = "123456789.987654321"
    AssertEqual "Perfect Zero Annihilation Subtraction", A.Minus(B).ValueEx, "0"
End Sub

Private Sub Test_Multiplication()
    Dim A As BigDecimal, B As BigDecimal

    Set A = New_BigDecimal()
    Set B = New_BigDecimal()

    ' Test: Massive mismatch of scales (forces alignment shift)
    A = "1000000000000"
    B = "0.0000000000001"
    AssertEqual "Scale Alignment Multiplication", A.Times(B).ValueEx, "0.1"

    ' Test: Nested Carries (9s multiplying 9s cascade carries all the way up)
    A = "99.99"
    B = "99.99"
    AssertEqual "Cascading Carry Multiplication", A.Times(B).ValueEx, "9998.0001"

    ' Test: Multiplication by Zero (Ensures scale, sign, and string collapse to "0")
    A = "-123456.7890"
    B = "0.000"
    AssertEqual "Zero Multiplication (Negative)", A.Times(B).ValueEx, "0"

    ' Test: Double Negative (Ensures sign tracking correctly flips back to positive)
    A = "-12.5"
    B = "-4"
    AssertEqual "Double Negative Multiplication", A.Times(B).ValueEx, "50"

    ' Test: Mixed Signs (Positive * Negative)
    A = "2.5"
    B = "-1.5"
    AssertEqual "Mixed Sign Multiplication", A.Times(B).ValueEx, "-3.75"

    ' Test: Precision Expansion (Ensures total scale equals sum of individual scales: 4 + 4 = 8 decimals)
    A = "0.1234"
    B = "0.5678"
    AssertEqual "Precision Expansion Scale Sum", A.Times(B).ValueEx, "0.07006652"

    ' Test: Underflow/Micro-decimal Limit (Tests trailing decimal tracking)
    A = "0.00000001"
    B = "0.00000001"
    AssertEqual "Micro-decimal Multiplication", A.Times(B).ValueEx, "0.0000000000000001"

    ' Test: Square of Large Number (Forces string handling to cross memory chunk/array limit boundaries)
    A = "123456789"
    B = "123456789"
    AssertEqual "Large Integer Square", A.Times(B).ValueEx, "15241578750190521"
End Sub

Private Sub Test_Division()
    Dim A As BigDecimal, B As BigDecimal
    Dim Expected As String
    Dim OrgTarScale As Long

    Set A = New_BigDecimal()
    Set B = New_BigDecimal()

    OrgTarScale = BigOption.TarScale

    ' Test: Exact interger division
    A = "100"
    B = "4"
    AssertEqual "Exact Integer Division", A.Divide(B).ValueEx, "25"

    ' Test: Exact termination decimals
    A = "1"
    B = "8"
    AssertEqual "Terminating Decimal Division", A.Divide(B).ValueEx, "0.125"

    ' Testing: Truncate/round infinite loops.
    ' 1 / 3 = 0.3333333333...
    BigOption.TarScale = 30
    A = "1"
    B = "3"
    Expected = "0.333333333333333333333333333333"
    AssertEqual "Repeating Decimal (1/3)", A.Divide(B).ValueEx, Expected

    ' 22 / 7 (Classic Pi approximation check)
    BigOption.TarScale = 25
    A = "22"
    B = "7"
    Expected = "3.1428571428571428571428571"
    AssertEqual "Repeating Decimal Loop (22/7)", A.Divide(B).ValueEx, Expected

    ' Test: Dividing small by large (Underflow Hazard)
    ' Forces the engine to accurately pad leading zeros inside the fraction array.
    BigOption.TarScale = 30
    A = "0.0000000001"
    B = "100000"
    Expected = "0.000000000000001"
    AssertEqual "Small Divided by Large", A.Divide(B).ValueEx, Expected

    ' Test: Pos / Neg = Neg
    A = "10"
    B = "-2"
    AssertEqual "Positive Divided By Negative", A.Divide(B).ValueEx, "-5"

    ' Test: Neg / Neg = Pos
    A = "-15.5"
    B = "-5"
    AssertEqual "Negative Divided By Negative", A.Divide(B).ValueEx, "3.1"

    ' Test: Precision & carry rounding boundary
    ' 2 / 3 = 0.666666...66667 (Should round up at the target scale edge!)
    BigOption.TarScale = 20
    A = "2"
    B = "3"
    Expected = "0.66666666666666666667"
    AssertEqual "Division Edge Round-Up (2/3)", A.Divide(B).ValueEx, Expected

    BigOption.TarScale = OrgTarScale
End Sub

Private Sub Test_Logarithms()
    Dim A As BigDecimal
    Dim Expected As String
    Dim OrgTarScale As Long

    Set A = New_BigDecimal()

    OrgTarScale = BigOption.TarScale
    BigOption.TarScale = 100
    BigOption.Cache = False

    ' Test: Log of 100 should be 2
    A = "100"
    AssertEqual "Logarithm (No Cache) Log(100) = 2", A.Log().ValueEx, "2"

    ' Test: Verification of Ln2 against a hardcoded known true value (100 digits)
    Expected = "0.6931471805599453094172321214581765680755001343602552541206800094933936219696947156058633269964186875"
    A = 2
    AssertEqual "High-Precision (No Cache) Ln(2)", A.Ln().ValueEx, Expected

    BigOption.Cache = True
    BigOption.TarScale = OrgTarScale
End Sub

Private Sub Test_Exponential()
    Dim A As BigDecimal
    Dim Expected As String
    Dim OrgTarScale As Long

    Set A = New_BigDecimal()

    OrgTarScale = BigOption.TarScale

    ' Test: e^0 must equal 1
    A = "0"
    AssertEqual "Exponential of Zero", A.Exp().ValueEx, "1"

    ' Test: Round-trip Identity Exp(Ln(5)) should equal 5
    BigOption.TarScale = 50
    A = "5"
    AssertEqual "Exponential Round-Trip Identity", A.Ln().Exp().ValueEx, "5"

    ' Test: High Precision Exp(1) against Euler's Constant (50 digits)
    A = "1"
    Expected = "2.71828182845904523536028747135266249775724709369996"
    AssertEqual "High-Precision Exp(1) = e", A.Exp().ValueEx, Expected

    BigOption.TarScale = OrgTarScale
End Sub

Private Sub Test_RemainderAndModulo()
    Dim A As BigDecimal, B As BigDecimal

    Set A = New_BigDecimal()
    Set B = New_BigDecimal()

    ' Test: Clean division (No remainder)
    A = "10"
    B = "2"
    AssertEqual "Standard Case: Clean Division", A.Remdr(B).ValueEx, "0"

    ' Test: Standard remainder
    A = "10"
    B = "3"
    AssertEqual "Standard Case: With Remainder", A.Remdr(B).ValueEx, "1"

    ' Test: Small number mod large number
    A = "3"
    B = "10"
    AssertEqual "Fast Path: Small Mod Large", A.Remdr(B).ValueEx, "3"

    ' Test: Decimal remainder
    A = "5.5"
    B = "2"
    AssertEqual "Decimal Case: Half Remainder", A.Remdr(B).ValueEx, "1.5"

    ' Test: Small decimal mod large decimal
    A = "1.2345"
    B = "3"
    AssertEqual "Decimal Case: Small Mod Large", A.Remdr(B).ValueEx, "1.2345"

    ' Test: Remainder (Modulo) rule: 10.5 Mod 3 = 1.5
    A = "10.5"
    B = "3"
    AssertEqual "Remainder check", A.Remdr(B).ValueEx, "1.5"

    ' Test: High precision remainder
    A = "10"
    B = "3.333"
    AssertEqual "High Precision Remainder", A.Remdr(B).ValueEx, "0.001"

    ' Test: Positive A, Negative B
    ' 5 Mod -3 = -1  (Because 5 / -3 floors to -2. Next, -3 * -2 = 6. 5 - 6 = -1)
    A = "5"
    B = "-3"
    AssertEqual "Signed Case: Pos Mod Neg", A.Remdr(B).ValueEx, "-1"

    ' Test: Negative A, Positive B
    ' -5 Mod 3 = 1   (Because -5 / 3 floors to -2. Next, 3 * -2 = -6. -5 - (-6) = 1)
    A = "-5"
    B = "3"
    AssertEqual "Signed Case: Neg Mod Pos", A.Remdr(B).ValueEx, "1"

    ' Test: Both Negative
    ' -5 Mod -3 = -2 (Because -5 / -3 floors to 1. Next, -3 * 1 = -3. -5 - (-3) = -2)
    A = "-5"
    B = "-3"
    AssertEqual "Signed Case: Neg Mod Neg", A.Remdr(B).ValueEx, "-2"
End Sub

Private Sub Test_PiIdentities()
    Dim A As BigDecimal
    Dim Expected As String
    Dim OrgTarScale As Long

    Set A = New_BigDecimal()

    OrgTarScale = BigOption.TarScale
    BigOption.TarScale = 100

    ' Test: Verification of Pi against a hardcoded known true value (100 digits)
    BigOption.Cache = False
    A = BigPi
    Expected = "3.141592653589793238462643383279502884197169399375105820974944592307816406286208998628034825342117068"
    AssertEqual "High-Precision (No Cache) Pi Calculation", A.ValueEx, Expected
    BigOption.Cache = True

    ' Test: Verification of the Caching/Buffering mechanism
    A = BigPi
    AssertEqual "High-Precision (Cached) Pi Fetch", A.ValueEx, Expected

    BigOption.TarScale = OrgTarScale
End Sub

Private Sub Test_ComparisonsAndSignChecks()
    Dim A As BigDecimal, B As BigDecimal

    Set A = New_BigDecimal()
    Set B = New_BigDecimal()

    ' Test: For Zero
    A = "0.000000"
    AssertEqual "IsZero on True Zero", CStr(A.IsZero()), "True"
    AssertEqual "IsPos on True Zero", CStr(A.IsPos()), "False"
    AssertEqual "IsNeg on True Zero", CStr(A.IsNeg()), "False"

    ' Test: For Positive
    A = "0.000000000000000000000000000001"
    AssertEqual "IsZero on Tiny Positive", CStr(A.IsZero()), "False"
    AssertEqual "IsPos on Tiny Positive", CStr(A.IsPos()), "True"
    AssertEqual "IsNeg on Tiny Positive", CStr(A.IsNeg()), "False"

    ' Test: For Negative
    A = "-123456789.987654321"
    AssertEqual "IsZero on Negative", CStr(A.IsZero()), "False"
    AssertEqual "IsPos on Negative", CStr(A.IsPos()), "False"
    AssertEqual "IsNeg on Negative", CStr(A.IsNeg()), "True"

    ' Test: For IsEq and IsNEq
    ' These have different string lengths, but are mathematically identical!
    A = "100.00500"
    B = "100.005"
    AssertEqual "IsEq (Trailing Zeros Ignored)", CStr(A.IsEq(B)), "True"
    AssertEqual "IsNEq (Trailing Zeros Ignored)", CStr(A.IsNEq(B)), "False"

    B = "100.00501"
    AssertEqual "IsEq on Mismatch", CStr(A.IsEq(B)), "False"
    AssertEqual "IsNEq on Mismatch", CStr(A.IsNEq(B)), "True"

    ' Setup: A is slightly smaller than B
    A = "5.00000000000000000000000000000000000001"
    B = "5.00000000000000000000000000000000000002"

    ' Test: Strictly Less Than / Greater Than
    AssertEqual "IsLT (Strictly Less)", CStr(A.IsLT(B)), "True"
    AssertEqual "IsGT (Strictly Greater)", CStr(A.IsGT(B)), "False"

    ' Test: Less-Than-Equal / Greater-Than-Equal (Distinct values)
    AssertEqual "IsLE (Distinct values)", CStr(A.IsLE(B)), "True"
    AssertEqual "IsGE (Distinct values)", CStr(A.IsGE(B)), "False"

    ' Edge Case: Check behavior when values are EXACTLY equal
    B = "5.00000000000000000000000000000000000001"

    AssertEqual "IsLT when Equal", CStr(A.IsLT(B)), "False"
    AssertEqual "IsGT when Equal", CStr(A.IsGT(B)), "False"
    AssertEqual "IsLE when Equal", CStr(A.IsLE(B)), "True"
    AssertEqual "IsGE when Equal", CStr(A.IsGE(B)), "True"
End Sub

Private Sub Test_RoundingAndTruncation()
    Dim A As BigDecimal
    Dim OrgTarScale As Long

    Set A = New_BigDecimal()

    OrgTarScale = BigOption.TarScale

    ' Test: Rounding UP without leading digits
    ' If we force target to add leading Zero before rounding up
    BigOption.TarScale = 10
    A = ".99999999999"
    AssertEqual "Rounding UP without leading digits before dot", A.ValueEx, "1"

    ' Test: Rounding UP at boundary
    ' If we force target scale to 3, 1.2346 should round UP to 1.235
    BigOption.TarScale = 3
    A = "1.2346"
    AssertEqual "Rounding UP Test", A.ValueEx, "1.235"

    ' Test: Rounding DOWN at boundary
    A = "1.2344"
    AssertEqual "Rounding DOWN Test", A.ValueEx, "1.234"

    ' Test: Trailing zeros after a significant digit
    ' Making sure zero-stripping doesn't break on something like 1.0003
    BigOption.TarScale = 4
    A = "1.0003"
    AssertEqual "Embedded Internal Zeros", A.ValueEx, "1.0003"

    BigOption.TarScale = OrgTarScale
End Sub

Private Sub Test_EdgeCases()
    Dim A As BigDecimal

    Set A = New_BigDecimal()

    On Error Resume Next

    ' Test: Does Log base 10 throw an error on negative numbers?
    A = "-5"

    Err.Clear
    Dim Result As BigDecimal
    Set Result = A.Log()

    If Err.Number = vbObjectError + 1104 Then
	AssertEqual "Log Negative Guard Rail", "Error " & vbObjectError + 1104 & " Thrown", "Error " & vbObjectError + 1104 & " Thrown"
    Else
	AssertEqual "Log Negative Guard Rail", "No Error / Wrong Error", "Error " & vbObjectError + 1104 & " Thrown"
    End If
    On Error GoTo 0
End Sub

Private Sub Test_MathBoundaryViolations()
    On Error Resume Next
    Dim Zero As BigDecimal, One As BigDecimal, Neg As BigDecimal
    Dim Result As BigDecimal
    Dim ErrNum As Long

    Set Zero = New_BigDecimal()
    Set One = New_BigDecimal()
    Set Neg = New_BigDecimal()

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
    ErrNum = vbObjectError + 1105
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
    Dim A As BigDecimal, B As BigDecimal
    Dim Expected As String
    Dim OrgTarScale As Long

    Set A = New_BigDecimal()
    Set B = New_BigDecimal()

    OrgTarScale = BigOption.TarScale
    BigOption.TarScale = 50

    ' Test: Absolute value of a negative number
    A = "-123.456"
    AssertEqual "AbsVal Negative Input", A.AbsVal().ValueEx, "123.456"

    ' Test: Negate a positive number
    B = "123.456"
    AssertEqual "Neg Positive Input", B.Neg().ValueEx, "-123.456"

    ' Test: Negate a Negative number
    B = "-123.456"
    AssertEqual "Neg Negative Input", B.Neg().ValueEx, "123.456"

    ' Test: Truncate and fraction parts of a decimal number
    A = "567.89012"
    AssertEqual "Truncate Pulls Integer", A.Trunc().ValueEx, "567"
    AssertEqual "Frac Pulls Fraction Part", A.Frac().ValueEx, "0.89012"

    ' Test: Negative edge case for Trunc/Frac
    B = "-5.67"
    AssertEqual "Truncate Negative", B.Trunc().ValueEx, "-5"
    AssertEqual "Frac Negative", B.Frac().ValueEx, "-0.67"

    ' Test: Multiple leading zeros (should be stripped, not treated as octal)
    A = "000005.67000"
    AssertEqual "Parser Leading/Trailing Zeros", A.ValueEx, "5.67"

    ' Test: Explicit positive sign
    A = "+89.12"
    AssertEqual "Parser Explicit Plus Sign", A.ValueEx, "89.12"

    ' Test:  explicit precision rounding
    A = "1.234567"
    AssertEqual "Round to 4 places", A.Round(4).ValueEx, "1.2346"

    ' Test: 5! = 5 * 4 * 3 * 2 * 1 = 120
    A = "5"
    AssertEqual "Factorial of 5", A.Fact().ValueEx, "120"

    ' Hardcore 30! stress test to make sure large integer carries hold up
    A = "30"
    Expected = "265252859812191058636308480000000"
    AssertEqual "Factorial Large Stress (30!)", A.Fact().ValueEx, Expected

    ' Test: Sqrt of 2 to 50 decimal digits
    BigOption.TarScale = 50
    A = "2"
    Expected = "1.41421356237309504880168872420969807856967187537695"
    AssertEqual "High-Precision Sqrt(2)", A.Sqrt().ValueEx, Expected

    ' Test Case A: Simple Integer Power (3^4 = 81)
    A = "3"
    B = "4"
    AssertEqual "Pow Integer Base/Exp", A.Pow(B).ValueEx, "81"

    ' Test Case B: Fractional/Transcendental Power (2.5 ^ 3.5)
    ' Evaluated internally via: Exp(3.5 * Ln(2.5))
    BigOption.TarScale = 40
    A = "2.5"
    B = "3.5"
    Expected = "24.705294220065463531241355815880613544684"
    AssertEqual "Pow Fractional Base/Exp", A.Pow(B).ValueEx, Expected

    ' Test: Round-Trip Identity Test - Square(Sqr(3)) = 3
    A = "3"
    AssertEqual "Irrational Round-Trip Identity", A.Sqrt().Sqr().ValueEx, "3"

    ' Test: Fractional Round-Trip - Square(Sqr(0.5)) = 0.5
    A = "0.5"
    AssertEqual "Fractional Round-Trip Identity", A.Sqrt().Sqr().ValueEx, "0.5"

    ' Test: Clean terminating fractional root (9/16)
    A = "0.5625"
    AssertEqual "Terminating Fractional Root", A.Sqrt().ValueEx, "0.75"

    ' Test: Deep fractional perfect square (1/1000000)
    A = "0.000001"
    AssertEqual "Deep Precision Fractional Root", A.Sqr().ValueEx, "0.000000000001"

    ' Test: Repeating fractional round-trip via inversion
    A = "9"
    Expected = "0.3333333333333333333333333333333333333333"
    AssertEqual "Inverted Fractional Root", A.Invert().Sqrt().ValueEx, Expected

    ' Test: Micro-fractional irrational root
    A = "0.000002"
    Expected = "0.0014142135623730950488016887242096980786"
    AssertEqual "Micro Irrational Fractional Root", A.Sqrt().ValueEx, Expected

    BigOption.TarScale = OrgTarScale
End Sub

Private Sub Test_Pow10Engine()
    Dim A As BigDecimal
    Dim Expected As String
    Dim OrgTarScale As Long

    Set A = New_BigDecimal()

    OrgTarScale = BigOption.TarScale
    BigOption.TarScale = 50

    ' Test: Base Case 10^0
    A = "0"
    AssertEqual "Zero Exponent Pow10", A.Pow10().ValueEx, "1"

    ' Test: Large integer structural shift (Stresses string padding / array shift)
    A = "25"
    Expected = "10000000000000000000000000"
    AssertEqual "Large Integer Pow10", A.Pow10().ValueEx, Expected

    ' Test: Deep negative shift 10^-8
    A = "-8"
    AssertEqual "Deep Negative Pow10", A.Pow10().ValueEx, "0.00000001"

    BigOption.TarScale = 40

    ' Test: Negative fractional exponent (10^-1.5)
    A = "-1.5"
    Expected = "0.0316227766016837933199889354443271853372"
    AssertEqual "Pow10 Negative Fractional", A.Pow10().ValueEx, Expected

    ' Test: Complex irrational / highly precise exponent
    A = "1.234567"
    Expected = "17.1619645281983487172093706480028667833534"
    AssertEqual "Pow10 Complex Fractional Exponent", A.Pow10().ValueEx, Expected

    BigOption.TarScale = 50

    ' Test: Explicit plus sign on exponent integer
    A = "+2"
    AssertEqual "Explicit Positive Sign Exponent Pow10", A.Pow10().ValueEx, "100"

    ' Test: Extraneous padded zeros on integer exponent
    A = "00004.00000"
    AssertEqual "Dirty Integer String Padded Exponent Pow10", A.Pow10().ValueEx, "10000"

    ' Test: Round-trip consistency (10^x * 10^-x = 1)
    A = "2.5"
    Dim PosPart As BigDecimal
    Dim NegPart As BigDecimal
    Set PosPart = A.Pow10()         ' 10^2.5
    Set NegPart = A.Neg().Pow10()   ' 10^-2.5
    AssertEqual "Pow10 Inverse Round-Trip Identity", PosPart.Times(NegPart).ValueEx, "1"

    BigOption.TarScale = OrgTarScale
End Sub

Private Sub Test_CubeAndCubert()
    Dim A As BigDecimal
    Dim Expected As String

    Set A = New_BigDecimal()

    ' Test: Basic decimal cube
    A = "2.5"
    AssertEqual "Basic Decimal Cube", A.Cube().ValueEx, "15.625"

    ' Test: Large Precision Expansion
    A = "1.00003"
    Expected = "1.000090002700027"
    AssertEqual "Precision Expansion Cube", A.Cube().ValueEx, Expected

    ' Test: Negative Cube Identity
    A = "-3"
    AssertEqual "Negative Base Cube", A.Cube().ValueEx, "-27"

    ' Test: Large Integer Stress Cube
    A = "999999"
    Expected = "999997000002999999"
    AssertEqual "Massive Integer Carry Cube", A.Cube().ValueEx, Expected

    ' Test: Zero Boundary Condition
    A = "0"
    AssertEqual "Zero Cube Annihilation", A.Cube().ValueEx, "0"

    ' Test: Round-Trip Identity Test - Cube(Cubert(5)) = 5
    A = "5"
    AssertEqual "Irrational Cube Round-Trip Identity", A.Cubert().Cube().ValueEx, "5"

    ' Test: Negative Round-Trip Identity - Cube(Cubert(-2)) = -2
    A = "-2"
    AssertEqual "Negative Irrational Round-Trip Identity", A.Cubert().Cube().ValueEx, "-2"

    ' Test: Perfect integer cube root
    A = "1728"
    AssertEqual "Perfect Cube Root", A.Cubert().ValueEx, "12"

    ' Test: Clean terminating fractional cube root
    A = "0.015625"
    AssertEqual "Terminating Fractional Cube Root", A.Cubert().ValueEx, "0.25"

    ' Test: Negative Cube Root Support
    A = "-8"
    AssertEqual "Negative Input Cube Root", A.Cubert().ValueEx, "-2"

    ' Test: Irrational number (Precision Cap Test)
    A = "2"
    Expected = "1.2599210498948731647672106072782283505703"
    AssertEqual "Precision Capped Cube Root", A.Cubert().ValueEx, Expected

    ' Test: Micro-decimal cube root
    A = "0.000000027"
    AssertEqual "Micro Decimal Cube Root", A.Cubert().ValueEx, "0.003"
End Sub

Private Sub Test_FloorAndCeiling()
    Dim A As BigDecimal
    Dim Expected As String

    Set A = New_BigDecimal()

    ' Test: Standard positive decimal
    A = "5.400"
    AssertEqual "Positive Floor", A.Floor().ValueEx, "5"

    ' Test: Standard negative decimal (moves away from zero)
    A = "-5.400"
    AssertEqual "Negative Floor", A.Floor().ValueEx, "-6"

    ' Test: Zero representation
    A = "0.000"
    AssertEqual "Floor of Zero", A.Floor().ValueEx, "0"

    ' Test: Extreme precision asymmetry (Positive)
    A = "999999999999999999999999999999.000000000000000000000000000001"
    Expected = "999999999999999999999999999999"
    AssertEqual "Massive Positive Asymmetry Floor", A.Floor().ValueEx, Expected

    ' Test: Extreme precision asymmetry (Negative)
    A = "-999999999999999999999999999999.000000000000000000000000000001"
    Expected = "-1000000000000000000000000000000"
    AssertEqual "Massive Negative Asymmetry Floor", A.Floor().ValueEx, Expected

    ' Test: Standard positive decimal (moves away from zero)
    A = "5.400"
    AssertEqual "Positive Ceiling", A.Ceiling().ValueEx, "6"

    ' Test: Standard negative decimal (moves toward zero)
    A = "-5.400"
    AssertEqual "Negative Ceiling", A.Ceiling().ValueEx, "-5"

    ' Test: Zero representation
    A = "0.000"
    AssertEqual "Ceiling of Zero", A.Ceiling().ValueEx, "0"

    ' Test: Extreme precision asymmetry (Positive)
    A = "999999999999999999999999999999.000000000000000000000000000001"
    Expected = "1000000000000000000000000000000"
    AssertEqual "Massive Positive Asymmetry Ceiling", A.Ceiling().ValueEx, Expected

    ' Test: Extreme precision asymmetry (Negative)
    A = "-999999999999999999999999999999.000000000000000000000000000001"
    Expected = "-999999999999999999999999999999"
    AssertEqual "Massive Negative Asymmetry Ceiling", A.Ceiling().ValueEx, Expected
End Sub

Private Sub Test_CastingCBgDec()
    Dim Result As BigDecimal
    Dim Expected As String

    ' Test: Integer Type
    Dim vInt As Integer: vInt = -32768
    AssertEqual "Cast Integer Minimum Number", CBgDec(vInt).ValueEx, "-32768"

    ' Test: long Type
    Dim vLong As Long: vLong = 2147483647
    AssertEqual "Cast Long Maximum Number", CBgDec(vLong).ValueEx, "2147483647"

    ' Test: Double Type
    Dim vDouble As Double: vDouble = -0.000123456789
    AssertEqual "Cast Double Precision Number", CBgDec(vDouble).ValueEx, "-0.000123456789"

    ' Test: String Type
    AssertEqual "Cast String Clean Integer Number", CBgDec("12345678901234567890").ValueEx, "12345678901234567890"
    'AssertEqual "Cast String Padded Zeros", CBgDec("  -00125.4500  ").ValueEx, "-125.45"
    'AssertEqual "Cast String Scientific Notation", CBgDec("1.23e4").ValueEx, "12300"

    ' Test: Robustness & Guard Rails ---
    On Error Resume Next

    ' Test: Empty String handling (Should either return 0 or throw a clean error)
    Dim ErrCheck As String
    ErrCheck = ""
    Set Result = CBgDec(ErrCheck)   'returns 0 for empty string
    AssertEqual "Cast Empty String", CBgDec("").ValueEx, "0"

    ' Test: Malformed String Guard Rail
    Err.Clear
    Set Result = CBgDec("123.45.67")
    AssertEqual "Cast Malformed String Error Guard", (Err.Number <> 0), True

    ' Test: Completely Invalid Object Type Guard Rail
    Err.Clear
    Dim InvalidObj As Object
    Set InvalidObj = New Collection ' Collections are completely invalid inputs
    Set Result = CBgDec(InvalidObj)
    AssertEqual "Cast Invalid Object Type Guard", (Err.Number <> 0), True

    On Error GoTo 0
End Sub

Private Sub Test_ShiftDecimal()
    Dim A As BigDecimal
    Dim Expected As String

    Set A = New_BigDecimal()

    ' Test: Shifting Right (Positive Count / Multiplication) ---
    A = "1564.5345"
    AssertEqual "Right Standard Shift", A.Shift(2).ValueEx, "156453.45"

    A = "0.123456"
    AssertEqual "Right Shift From Zero Lead", A.Shift(1).ValueEx, "1.23456"

    ' Test: Shifting Left (Negative Count / Division) ---
    A = "1564.5345"
    AssertEqual "Left Standard Shift", A.Shift(-2).ValueEx, "15.645345"

    A = "12.3"
    AssertEqual "Left Shift To Zero Lead", A.Shift(-1).ValueEx, "1.23"

    ' Test: Boundary Overshoots (Requires Padding) ---
    A = "1.23"
    AssertEqual "Overshoot Padding Right Shift", A.Shift(5).ValueEx, "123000"

    A = "5" ' No existing decimal point
    AssertEqual "Integer Overshoot Right Shift", A.Shift(3).ValueEx, "5000"

    ' Test: Left overshoot: Needs leading zeros inserted after the "0."
    A = "1.23"
    AssertEqual "Overshoot Padding Left Shift", A.Shift(-3).ValueEx, "0.00123"

    A = "0.5"
    AssertEqual "Tiny Decimal Overshoot Shift Left", A.Shift(-2).ValueEx, "0.005"

    ' Test:  Sign Retention and Identity Cases ---
    A = "-123.45"
    AssertEqual "Shift Zero Count Identity", A.Shift(0).ValueEx, "-123.45"

    ' Test: Shifting a true zero should remain zero
    A = "0.0000"
    AssertEqual "Shift Operating on Zero", A.Shift(5).ValueEx, "0"

    ' Test: Negative sign preservation on right shift
    A = "-1.2345"
    AssertEqual "Right Shift Negative Sign", A.Shift(2).ValueEx, "-123.45"

    ' Test: Negative sign preservation on left overshoot
    A = "-1.23"
    AssertEqual "Left Shift Negative Sign Overshoot", A.Shift(-2).ValueEx, "-0.0123"
End Sub

Private Sub Test_GreatestCommonDivisor()
    Dim A As BigDecimal, B As BigDecimal

    Set A = New_BigDecimal()
    Set B = New_BigDecimal()

    ' Test: Standard composite numbers
    A = "24"
    B = "36"
    AssertEqual "Standard Composite GCD", A.GCD(B).ValueEx, "12"

    ' Test: Prime / Coprime numbers (Should yield 1)
    A = "17"
    B = "23"
    AssertEqual "GCD Of Coprime Numbers", A.GCD(B).ValueEx, "1"

    ' Test: One value is zero (GCD(X, 0) = X)
    A = "0"
    B = "45"
    AssertEqual "Zero Edge GCD Case", A.GCD(B).ValueEx, "45"

    ' Test: Negative inputs (GCD is always positive)
    A = "-12"
    B = "18"
    AssertEqual "Negative Inputs GCD", A.GCD(B).ValueEx, "6"

    ' Test: Numbers with trailing scales (Should truncate and evaluate)
    A = "50.000"
    B = "20.450"
    AssertEqual "Scale Shift Check GCD", A.GCD(B).ValueEx, "0.05"

    ' Test: Large Scale Stress
    A = "1234567890123456"
    B = "987654321"
    AssertEqual "Large Scale Stress GCD", A.GCD(B).ValueEx, "3"

    ' Test: Decimals that share a fractional factor
    A = "1.5"
    B = "2.5"
    AssertEqual "Fractional Match GCD", A.GCD(B).ValueEx, "0.5"

    ' Test: High precision decimals needing zero-padding on back-shift
    A = "0.12"
    B = "0.18"
    AssertEqual "Low Decimal Precision GCD", A.GCD(B).ValueEx, "0.06"

    ' Test: Trailing zero asymmetry protection
    A = "0.25"
    B = "0.50000"
    AssertEqual "GCD Of Trailing Zeros Ignored", A.GCD(B).ValueEx, "0.25"

    ' Test: Negative inputs (GCD result must be positive)
    A = "-1.2"
    B = "1.8"
    AssertEqual "Negative Decimal Inputs GCD", A.GCD(B).ValueEx, "0.6"
End Sub

Private Sub Test_LeastCommonMultiple()
    Dim A As BigDecimal, B As BigDecimal

    Set A = New_BigDecimal()
    Set B = New_BigDecimal()

    ' Test: Standard integers
    A = "12"
    B = "18"
    AssertEqual "Standard Integers LCM", A.LCM(B).ValueEx, "36"

    ' Test: Prime / Coprime numbers (LCM is exactly A * B)
    A = "5"
    B = "7"
    AssertEqual "LCM Of Coprime Numbers", A.LCM(B).ValueEx, "35"

    ' Test: Zero edge cases (Should annihilate to 0)
    A = "0"
    B = "999"
    AssertEqual "Zero Edge Case LCM", A.LCM(B).ValueEx, "0"

    ' Test: Fractional decimals (e.g., LCM of 0.25 and 0.75 is 0.75)
    A = "0.25"
    B = "0.75"
    AssertEqual "Fractional Decimals LCM", A.LCM(B).ValueEx, "0.75"

    ' Test: Negative inputs (LCM is always a positive value)
    A = "-6"
    B = "8"
    AssertEqual "Negative Inputs LCM", A.LCM(B).ValueEx, "24"
End Sub
