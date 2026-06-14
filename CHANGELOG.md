# Changelog

All notable changes to the BigNumber Add-in will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

---

## [1.5] - 2026-06-14

### Added

- Added reusable BigDecimal constants:

  - `BigZero`
  - `BigOne`
  - `BigTwo`

  These constants are instantiated once and reused for the lifetime of the application, reducing object creation overhead.

- Added the `BigPi` constant.

  - Honors the current `TarScale` setting when calculating the standard value of π.

- Added the `BigOption.Cache` property.

  - Allows the internal cache to be enabled or disabled.
  - Facilitates automated testing and benchmarking.

- Added an abort mechanism to the Yield Manager.

  - Users can press and hold the **Esc** key to display an abort dialog.
  - Available options include:
    - Continue
    - End
    - Debug
    - Help
  - Improves control during long-running calculations while Excel is displaying a "Not Responding" state.

- Added new BigDecimal methods:

  - `Floor()` – Returns the greatest integer less than or equal to the value.
  - `Ceiling()` – Returns the least integer greater than or equal to the value.
  - `Shift(l)` – Shifts the decimal point left or right.
  - `GCD(x)` – Returns the greatest common divisor.
  - `LCM(x)` – Returns the least common multiple.
  - `Sqr()` – Returns the square of the value.
  - `Invert()` – Returns the reciprocal of the value.
  - `Cube()` – Returns the cube of the value.
  - `Cubert()` – Returns the cube root of the value.
  - `Pow10()` – Returns 10 raised to the power of the value.

### Changed

- Updated `.ValueEx()` so that omitted parameters now default to `vbString`.

  - Simplifies conversion of BigDecimal values to strings.

- Updated internal code to use object assignment (`Set X = Y.Minus(BigOne)`) where appropriate.

  - Avoids copying internal BigDecimal data structures.
  - Improves execution speed by reusing object references.
  - Applied only where immutability guarantees can be maintained safely.

- Increased internal cache precision for:

  - `Ln2()`
  - `Ln10()`
  - `Pi`

  Cache size increased from **150 digits** to **500 digits**.

- Replaced Dictionary-based storage with optimized array structures in the BigOption library.

  - Alternative implementations using Collections and more complex array algorithms were evaluated but resulted in slower overall performance.

- Updated remaining core arithmetic functions to use byte arrays instead of string-based processing:

  - `Plus_Positive()`
  - `Minus_Positive()`
  - `Compare_Positive()`
  - `Divide_Positive()`
  - `Times_Positive()`

### Performance

- Further optimized the Yield Manager.

  - Delivers a significant performance improvement compared to version 1.4.

- Optimized logarithmic and exponential functions:

  - `Ln()`
  - `Ln2()`
  - `Ln10()`
  - `Exp()`

  Testing showed that a reduced Mercator/Taylor series implementation significantly outperformed the Atanh-series approach when cache usage was disabled.

- Reduced internal overhead when updating BigDecimal state by introducing bulk variable update methods.
  
  - `SetRaw()`
  - `SetRawB()`
  - `GetRaw()`

  These methods allow all internal BigDecimal variables to be updated in a single call, eliminating multiple Friend property calls for common internal operations.

### Testing

- Expanded the automated test suite from **69 tests** to **159 tests**.

- Moved automated testing into the **BigCalculator** workbook.

- Tests can be executed by:

  1. Opening the VBA Editor.
  2. Running the `RunAllTests` macro.

- Complete execution time is approximately **3–4 seconds**.

### Documentation

- Updated and expanded the documentation.
- Added new sections and examples to improve usability and clarity.

---

## [1.4] - 2026-05-30

### Added

- Added helper methods to improve code readability and simplify algorithm development:
  - `IsZero()` – Returns `True` if the value equals zero.
  - `IsNeg()` – Returns `True` if the value is negative.
  - `IsPos()` – Returns `True` if the value is positive.
  - `IsEq()` – Returns `True` if two values are equal.
  - `IsNEq()` – Returns `True` if two values are not equal.
  - `IsGT()` – Returns `True` if the value is greater than another value.
  - `IsLT()` – Returns `True` if the value is less than another value.
  - `IsGE()` – Returns `True` if the value is greater than or equal to another value.
  - `IsLE()` – Returns `True` if the value is less than or equal to another value.
  - `Log()` – Returns the base-10 logarithm of the value.

- Added the `AddTrailingZeros` option.
  - Appends trailing zeros when rounding if the requested scale exceeds the configured `TarScale`.
  - Requires `AutoRound` to be enabled.

- Added a Yield class to manage the BigDecimal performance against Excel not responding.

- Added a Scale control to the BigCalculator example, allowing the internal `TarScale` value to be modified interactively.

- Added a comprehensive automated test suite containing 69 tests covering all major BigDecimal functionality.
  - Results are displayed in the VBA Immediate Window.
  - Passed and failed tests are clearly identified.

### Changed

- Renamed `MaxScale` to `TarScale` (*Target Scale*) for improved clarity.
  - `TarScale` now controls the scale used when returning values through `.ValueEx(vbString)`.
  - `MaxScale` has been repurposed internally to store additional digits used during rounding, called Guard digits.

- Increased the number of internal Guard digits to 5 to improve rounding accuracy and reduce propagation of rounding errors.

- Renamed `Cmp()` to `CmpTo()` to improve code readability.

- Modified rounding behaviour so that `Round()` is only applied when returning numeric values.
  - String values are rounded directly to `TarScale` decimal places.
  - The `Round()` parameter now explicitly determines the rounding precision applied to the `BigDecimal`.

### Performance

- Removed the use of `Mid$()` and `Asc()` from the core `Compare_Positive()` and `Round()` methods.
  - Replaced with optimized byte-array processing.
  - Significantly improves performance of comparison and rounding operations.

- Optimized logarithmic calculations used by BigDecimal.
  - Improved algorithm efficiency.
  - Added caching for `Ln2()` and `Ln10()` to accelerate repeated calculations.

### Fixed

- Fixed the BigCalculator example by ensuring `BigDecimal` objects are instantiated before checking for `Nothing`, eliminating runtime errors.

### Documentation

- Updated and expanded the documentation.
- Added new sections and examples to improve usability and clarity.

---

## [1.3] - 2026-05-23

### Added

#### Automatic Type Casting
- Added support for auto-casting standard VBA types (e.g. `vbLong`, `vbDouble`) to a `BigDecimal` object.
- This greatly simplifies writing arithmetic expressions.

**Example**
```vb
Num1 = Num1.Plus(1)   ' The integer 1 is automatically cast to a BigDecimal
```

#### Default Property Assignment
- Assigned the `.Value` property as the default property for the `BigDecimal` class.
- You can now use shorthand assignments instead of explicitly specifying the property.

**Example**
```vb
Num1 = 1      ' Equivalent to Num1.Value = 1
Num1 = Num2   ' Assigns the value of Num2 directly
```

#### Global Configuration via `BigOption`
- Introduced the `BigOption` class to allow global configuration settings across all instantiated `BigDecimal` objects.
- Features seamless background management for:
  - `.MaxScale`
  - `.AutoRound`
  - `.OutputDebug`

**Example**
```vb
BigOption.MaxScale = 200  ' Applies a 200 decimal limit to all instances
```

#### New Mathematical Functions
- Added the `.Pow()` method to return the power of `X^Y`.
- Implemented in both:
  - `BigDecimal`
  - `BigCalculator` example
- Added the `CBgDec()` casting function, which returns an initialized `BigDecimal` object.

---

### Changed

#### Unified Value Retrieval
- Replaced the separate:
  - `.StrValue`
  - `.LngValue`
  - `.DblValue`

with a unified `.ValueEx()` property.

- `.ValueEx()` accepts a parameter specifying the desired return type.

Supported types:
- `vbInteger`
- `vbLong`
- `vbLongLong`
- `vbSingle`
- `vbDouble`
- `vbString`

**Example**
```vb
MyStr = Num1.ValueEx(vbString)
```

#### Method Renaming (Shorthand Syntax)

| Category | Old Method | New Method | VBA Syntax Example |
|---|---|---|---|
| Arithmetic | `Add` | `Plus` | `Num1.Plus(Num2)` |
| Arithmetic | `Subtract` | `Minus` | `Num1.Minus(Num2)` |
| Arithmetic | `Multiply` | `Times` | `Num1.Times(Num2)` |
| Arithmetic | `Divide` | `Divide` | `Num1.Divide(Num2)` |
| Arithmetic | `Modulus` | `Remdr` | `Num1.Remdr(Num2)` |
| Math Functions | `SquareRoot` | `Sqrt` | `Num1.Sqrt()` |
| Math Functions | `Logarithm` | `Ln` | `Num1.Ln()` |
| Math Functions | `Exponential` | `Exp` | `Num1.Exp()` |
| Math Functions | `Factorial` | `Fact` | `Num1.Fact()` |
| Unary Ops | `Absolute` | `AbsVal` | `Num1.AbsVal()` |
| Unary Ops | `Negative` | `Neg` | `Num1.Neg()` |
| Formatting | `Whole` | `Trunc` | `Num1.Trunc()` |
| Formatting | `Round` | `Round` | `Num1.Round(Scale)` |
| Formatting | `Fraction` | `Frac` | `Num1.Frac()` |
| Formatting | `Trim` | `Trim` | `Num1.Trim()` |
| Formatting | `Compare` | `Cmp` | `Num1.Cmp(Num2)` |

---

### Changed — Scale & Truncation Behavior

#### String Initialization
- When instantiating a `BigDecimal` and initializing its value using a string, any decimal places exceeding `MaxScale` are now automatically truncated to the `MaxScale` limit.

#### Dynamic Scale Adjustment
- Decreasing the global or local `.MaxScale` property now actively reduces the decimal precision of the underlying `BigDecimal` number.
- No internal action is taken or required when increasing `MaxScale`.

---

### Performance

#### Core Arithmetic Optimization
- Rewrote fundamental internal core arithmetic functions to utilize byte arrays instead of string manipulation operations such as:
  - `Mid$()`
  - `Chr$()`
- This drastically improves the execution speed of the low-level building blocks upon which all other library functions rely.

---

### Documentation

#### Documentation Updates
- Fully updated the developer documentation to reflect all Version 1.3:
  - Changes
  - Syntax updates
  - New methods

---

## [1.2] - 2026-04-27

### Added

#### Initial Release
- Official public release of the `BigDecimal` add-in for VBA.
- Core architecture established to provide high-precision decimal arithmetic, overcoming the standard limitation of native VBA data types.
- Features:
  - Independent scale tracking
  - Basic mathematical functions
  - String-based initialization to prevent precision loss
