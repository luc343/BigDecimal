# Changelog

All notable changes to the BigNumber Add-in will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

---

## [1.3] - 2026-05-23

### Added

#### Automatic Type Casting
- Added support for auto-casting standard VBA types (e.g. `vbInteger`, `vbLong`, `vbLongLong`, `vbShort`, `vbDouble` and `vbstring`) to a `BigDecimal` object.
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
- Rewrote fundamental internal arithmetic functions to utilize byte arrays instead of string manipulation operations such as:
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
