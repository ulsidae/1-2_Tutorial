# Excel Functions for Beginners Who Want to Go Deeper

A beginner-friendly Excel tutorial covering basic functions for calculations, counting, rounding, and random number generation.

This lesson focuses on understanding how Excel functions are structured and applying them to basic calculations and simple data analysis.

> **Goal:** Understand what each function does, when to use it, and how to combine functions to solve practical problems.

The video is intentionally lightweight and example-driven. This README goes deeper into function syntax, edge cases, and technical details for readers who want more than the video covers.

## Functions Covered

| Function                                                                                  | Purpose                                                              |
| ----------------------------------------------------------------------------------------- | -------------------------------------------------------------------- |
| [`SUM`](https://support.microsoft.com/en-us/excel/functions/sum-function)                 | Adds numbers together                                                |
| [`AVERAGE`](https://support.microsoft.com/en-us/excel/functions/average-function)         | Calculates the arithmetic mean                                       |
| [`MAX`](https://support.microsoft.com/en-us/excel/functions/max-function)                 | Returns the largest value                                            |
| [`MIN`](https://support.microsoft.com/en-us/excel/functions/min-function)                 | Returns the smallest value                                           |
| [`COUNT`](https://support.microsoft.com/en-us/excel/functions/count-function)             | Counts cells containing numbers                                      |
| [`COUNTA`](https://support.microsoft.com/en-us/excel/functions/counta-function)           | Counts non-empty cells                                               |
| [`COUNTBLANK`](https://support.microsoft.com/en-us/excel/functions/countblank-function)   | Counts blank cells                                                   |
| [`ROUND`](https://support.microsoft.com/en-us/excel/functions/round-function)             | Rounds a number to a specified number of digits                      |
| [`ROUNDUP`](https://support.microsoft.com/en-us/excel/functions/roundup-function)         | Rounds a number away from zero                                       |
| [`ROUNDDOWN`](https://support.microsoft.com/en-us/excel/functions/rounddown-function)     | Rounds a number toward zero                                          |
| [`INT`](https://support.microsoft.com/en-us/excel/functions/int-function)                 | Rounds a number down to the nearest integer                          |
| [`MOD`](https://support.microsoft.com/en-us/excel/functions/mod-function)                 | Returns the remainder after division                                 |
| [`RAND`](https://support.microsoft.com/en-us/excel/functions/rand-function)               | Generates a random decimal number from 0 up to, but not including, 1 |
| [`RANDBETWEEN`](https://support.microsoft.com/en-gb/excel/functions/randbetween-function) | Generates a random integer within a specified inclusive range        |

---

## Understanding Excel Function Syntax

Before learning individual functions, it helps to understand how Excel formulas are structured.

For example:

```excel
=SUM(number1, [number2], ...)
```

* `=` indicates that you are entering a formula.
* `SUM` is the function name.
* `number1` is the first argument.
* `[number2]` indicates that the second argument is optional.
* `...` indicates that additional arguments can be provided.
* A comma `,` separates multiple arguments.
* `A1:A5` represents the range from cell `A1` through `A5`.

For example:

```excel
=SUM(A1:A5)
```

can be read as:

> Use the `SUM` function on the range `A1:A5`.

**Important:** Square brackets in function documentation are notation. You do **not** type them when entering an optional argument.

Excel functions may take one or more arguments, while some functions take none.

For example:

```excel
=ROUND(12.56, 1)
```

has two arguments:

```text
12.56 → number
1     → number of decimal places
```

A cell reference such as:

```excel
A1
```

refers to a single cell.

A range such as:

```excel
A1:A10
```

refers to every cell from `A1` through `A10`.

You can also provide multiple references or ranges when a function supports them:

```excel
=SUM(A1:A10, C1:C10)
```

The colon `:` defines a range, while the comma `,` separates arguments.

Understanding this distinction becomes increasingly important when working with more complex formulas.

---

## Functions

<details>
<summary><strong>1. SUM — Add values</strong></summary>

### [`SUM`](https://support.microsoft.com/en-us/excel/functions/sum-function)

`SUM` adds multiple numbers together.

### Syntax

```excel
=SUM(number1, [number2], ...)
```

### Example

```excel
=SUM(A1:A5)
```

If `A1:A5` contains:

```text
10
20
30
40
50
```

the result is:

```text
150
```

You can select an entire range instead of entering each cell individually:

```excel
=SUM(A1:A100)
```

This is easier to read and maintain than:

```excel
=A1+A2+A3+...+A100
```

You can also select multiple ranges:

```excel
=SUM(A1:A10, C1:C10)
```

The function adds the numeric values in both ranges.

> **Tip:** When you need a total, think `SUM` before manually typing `+`.

### Common Forms

```excel
=SUM(10, 20, 30)
=SUM(A1, A2, A3)
=SUM(A1:A3)
=SUM(A1:A3, C1:C3, 100)
```

`SUM` can accept individual numbers, cell references, ranges, and combinations of these.

When a range contains text or logical values directly in worksheet cells, `SUM` generally ignores those non-numeric cells.

For example:

```text
A1 = 10
A2 = Hello
A3 = 20
```

```excel
=SUM(A1:A3)
```

returns:

```text
30
```

However, values supplied directly as function arguments can be treated differently from values stored in referenced cells.

### Why `SUM` Is Usually Better Than Manual Addition

Instead of:

```excel
=A1+A2+A3+A4+A5
```

prefer:

```excel
=SUM(A1:A5)
```

The range-based formula is easier to read and maintain, especially when working with larger datasets.

</details>

<details>
<summary><strong>2. AVERAGE — Calculate the arithmetic mean</strong></summary>

### [`AVERAGE`](https://support.microsoft.com/en-us/excel/functions/average-function)

`AVERAGE` calculates the arithmetic mean of a group of numbers.

### Syntax

```excel
=AVERAGE(number1, [number2], ...)
```

### Example

```excel
=AVERAGE(A1:A5)
```

For:

```text
70
80
90
100
60
```

the result is:

```text
80
```

This is equivalent to:

```text
(70 + 80 + 90 + 100 + 60) ÷ 5 = 80
```

In a referenced range, `AVERAGE` ignores empty cells:

```text
10
20
(empty)
```

Result:

```text
15
```

A numeric zero is included:

```text
10
20
0
```

Result:

```text
10
```

Text entered directly into worksheet cells within a referenced range is generally ignored:

```text
10
20
Apple
```

Result:

```text
15
```

Text supplied directly as an argument can be treated differently.

### AVERAGE Is Not a Weighted Average

`AVERAGE` calculates:

```text
sum of values ÷ number of values
```

It does not automatically account for different weights or sample sizes.

For weighted averages, functions such as `SUMPRODUCT` are commonly used.

</details>

<details>
<summary><strong>3. MAX and MIN — Find extreme values</strong></summary>

### [`MAX`](https://support.microsoft.com/en-us/excel/functions/max-function) / [`MIN`](https://support.microsoft.com/en-us/excel/functions/min-function)

`MAX` returns the largest value, while `MIN` returns the smallest.

```excel
=MAX(A1:A10)
=MIN(A1:A10)
```

For:

```text
12
35
7
24
18
```

the results are:

```text
=MAX(A1:A5) → 35
=MIN(A1:A5) → 7
```

You do not need to sort the data first.

> **Tip:**
> "What is the highest?" → `MAX`
> "What is the lowest?" → `MIN`

`MAX` and `MIN` return the **value itself**, not the location of the cell containing that value.

For example:

```text
A1 = 20
A2 = 50
A3 = 10
```

```excel
=MAX(A1:A3)
```

returns:

```text
50
```

It does not return `A2`.

If you need to determine which cell or record contains the maximum or minimum value, functions such as `MATCH`, `INDEX`, or lookup functions may be required.

</details>

<details>
<summary><strong>4. COUNT, COUNTA, and COUNTBLANK — Count different types of cells</strong></summary>

### [`COUNT`](https://support.microsoft.com/en-us/excel/functions/count-function)

`COUNT` counts cells containing numeric values.

```excel
=COUNT(A1:A10)
```

It does not count text or empty cells.

Dates and times are also counted because Excel stores them internally as numbers.

### [`COUNTA`](https://support.microsoft.com/en-us/excel/functions/counta-function)

`COUNTA` counts cells that contain something.

```excel
=COUNTA(A1:A10)
```

This includes numbers, text, logical values, errors, and formulas.

An important edge case is a formula that returns an empty text string:

```excel
=""
```

Visually, the cell may appear blank, but `COUNTA` still counts it because the cell contains a formula.

### [`COUNTBLANK`](https://support.microsoft.com/en-us/excel/functions/countblank-function)

`COUNTBLANK` counts cells that Excel considers blank.

```excel
=COUNTBLANK(A1:A10)
```

A formula returning `""` can also be counted by `COUNTBLANK`.

### Quick Comparison

| Function     | Counts          |
| ------------ | --------------- |
| `COUNT`      | Numeric values  |
| `COUNTA`     | Non-empty cells |
| `COUNTBLANK` | Blank cells     |

Think of them as:

```text
COUNT      → "How many numbers?"
COUNTA     → "How many cells have something?"
COUNTBLANK → "How many cells are blank?"
```

</details>

<details>
<summary><strong>5. ROUND, ROUNDUP, and ROUNDDOWN — Control rounding</strong></summary>

### [`ROUND`](https://support.microsoft.com/en-us/excel/functions/round-function)

`ROUND` performs standard rounding.

```excel
=ROUND(12.56, 1)
```

Result:

```text
12.6
```

The second argument specifies how many digits to keep.

### [`ROUNDUP`](https://support.microsoft.com/en-us/excel/functions/roundup-function)

`ROUNDUP` rounds a number away from zero.

```excel
=ROUNDUP(12.51, 1)
```

Result:

```text
12.6
```

### [`ROUNDDOWN`](https://support.microsoft.com/en-us/excel/functions/rounddown-function)

`ROUNDDOWN` rounds a number toward zero.

```excel
=ROUNDDOWN(12.59, 1)
```

Result:

```text
12.5
```

### Quick Comparison

| Function    | Behavior          |
| ----------- | ----------------- |
| `ROUND`     | Standard rounding |
| `ROUNDUP`   | Away from zero    |
| `ROUNDDOWN` | Toward zero       |

For negative numbers:

```excel
=ROUNDUP(-12.51, 1)   → -12.6
=ROUNDDOWN(-12.59, 1) → -12.5
```

> **Tip:**
> `ROUNDUP` → **away from zero**
> `ROUNDDOWN` → **toward zero**

This is different from mathematical "rounding down" in the floor-function sense:

```excel
=INT(-12.9)           → -13
=ROUNDDOWN(-12.9, 0)  → -12
```

Therefore:

```text
INT        → toward negative infinity
ROUNDDOWN  → toward zero
```

### Negative `num_digits`

The second argument can also be negative.

```excel
=ROUND(1234, -2) → 1200
=ROUND(1234, -3) → 1000
```

So `num_digits` controls the position at which rounding occurs, not merely the number of decimal places.

</details>

<details>
<summary><strong>6. INT — Round down to an integer</strong></summary>

### [`INT`](https://support.microsoft.com/en-us/excel/functions/int-function)

`INT` returns a number rounded down to the nearest integer.

```excel
=INT(12.9)
```

Result:

```text
12
```

It does **not** simply remove the decimal portion.

```excel
=INT(-12.9)
```

returns:

```text
-13
```

> **Tip:** `INT` means **round toward negative infinity**.

### INT vs. TRUNC

`TRUNC` removes the fractional portion by moving toward zero:

```excel
=INT(12.9)    → 12
=INT(-12.9)   → -13

=TRUNC(12.9)  → 12
=TRUNC(-12.9) → -12
```

Therefore:

```text
INT(-12.9)   → -13
TRUNC(-12.9) → -12
```

This difference matters whenever negative values are possible.

### INT and the Floor Operation

Mathematically, `INT` behaves like the floor function:

```text
floor(x)
```

which returns the greatest integer less than or equal to `x`.

```text
floor(3.7)  = 3
floor(-3.7) = -4
```

</details>

<details>
<summary><strong>7. MOD — Get the remainder</strong></summary>

### [`MOD`](https://support.microsoft.com/en-us/excel/functions/mod-function)

`MOD` returns the remainder after division.

### Syntax

```excel
=MOD(number, divisor)
```

### Example

```excel
=MOD(10, 3)
```

Since:

```text
10 ÷ 3 = 3 remainder 1
```

the result is:

```text
1
```

### Even and Odd Numbers

```excel
=MOD(A1, 2)
```

For an integer:

```text
0 → even
1 → odd
```

### Repeating Cycles

```excel
=MOD(A1, 7)
```

can be used for repeating patterns, cyclic numbering, alternating groups, and periodic calculations.

For example:

```excel
=MOD(A1, 2)
```

can produce:

```text
0
1
0
1
0
1
...
```

> **Tip:** **MOD tells you what is left over.**

### Negative Values

The remainder has the same sign as the divisor:

```excel
=MOD(10, 3)   → 1
=MOD(-10, 3)  → 2
=MOD(10, -3)  → -2
```

### Division by Zero

```excel
=MOD(10, 0)
```

results in a division-by-zero error.

</details>

<details>
<summary><strong>8. RAND — Generate a random decimal</strong></summary>

### [`RAND`](https://support.microsoft.com/en-us/excel/functions/rand-function)

`RAND` generates a pseudo-random decimal number greater than or equal to `0` and less than `1`.

### Syntax

```excel
=RAND()
```

The result satisfies:

```text
0 ≤ RAND() < 1
```

Example results:

```text
0.183742
0.726391
0.054821
```

`RAND()` is volatile, so its result can change when Excel recalculates the worksheet.

> **Tip:** Do not use `RAND()` when you need the generated value to remain unchanged unless you intentionally convert it to a fixed value.

### Converting the Result to a Fixed Value

```text
RAND()
  ↓
Random value
  ↓
Copy
  ↓
Paste Values
  ↓
Fixed number
```

The fixed number is no longer controlled by the `RAND()` formula.

### Another Range

```excel
=RAND()*100
```

produces:

```text
0 ≤ x < 100
```

If you need an integer, `RANDBETWEEN` is usually clearer.

</details>

<details>
<summary><strong>9. RANDBETWEEN — Generate a random integer</strong></summary>

### [`RANDBETWEEN`](https://support.microsoft.com/en-gb/excel/functions/randbetween-function)

`RANDBETWEEN` generates a random integer between a specified minimum and maximum value.

### Syntax

```excel
=RANDBETWEEN(bottom, top)
```

For example:

```excel
=RANDBETWEEN(1, 6)
```

can generate:

```text
1
2
3
4
5
6
```

Both endpoints are included:

```text
bottom ≤ x ≤ top
```

A six-sided die can therefore be simulated with:

```excel
=RANDBETWEEN(1, 6)
```

Like `RAND()`, `RANDBETWEEN` is volatile.

### Negative Bounds

The arguments do not have to be positive:

```excel
=RANDBETWEEN(-10, 10)
```

can generate integers from `-10` through `10`.

### RAND vs. RANDBETWEEN

| Function              | Result                 |
| --------------------- | ---------------------- |
| `RAND()`              | Decimal: `0 ≤ x < 1`   |
| `RANDBETWEEN(1, 100)` | Integer: `1 ≤ x ≤ 100` |

If you generate a dataset that must remain fixed, convert the formulas to values after generation.

</details>

---

## Combining Functions

Excel functions become more useful when they are combined.

For example, `RANDBETWEEN` can generate random values, while `MAX`, `MIN`, `COUNT`, and `AVERAGE` can analyze the generated data.

Suppose `A1:A20` contains:

```excel
=RANDBETWEEN(1, 100)
```

You can then analyze the dataset:

```excel
=MAX(A1:A20)
```

Find the largest value.

```excel
=MIN(A1:A20)
```

Find the smallest value.

```excel
=COUNT(A1:A20)
```

Count numeric observations.

```excel
=AVERAGE(A1:A20)
```

Calculate the arithmetic mean.

Think of the workflow as:

```text
Generate
   ↓
Store
   ↓
Count
   ↓
Measure
   ↓
Compare
```

This is a basic example of separating **data generation** from **data analysis**.

The same pattern can be applied to test scores, sales numbers, survey results, inventory data, and other datasets.

### Important: Volatile Random Data

Because `RANDBETWEEN` is volatile, the generated dataset can change when Excel recalculates.

If the input data must remain constant, convert the generated formulas to values before continuing.

---

## Beginner Escape Route: What Should You Remember?

You do not need to memorize every function immediately.

Start with these mental shortcuts:

| What you want to do       | Function      |
| ------------------------- | ------------- |
| Add numbers               | `SUM`         |
| Find the average          | `AVERAGE`     |
| Find the highest value    | `MAX`         |
| Find the lowest value     | `MIN`         |
| Count numbers             | `COUNT`       |
| Count non-empty cells     | `COUNTA`      |
| Count blank cells         | `COUNTBLANK`  |
| Round normally            | `ROUND`       |
| Round away from zero      | `ROUNDUP`     |
| Round toward zero         | `ROUNDDOWN`   |
| Round down to an integer  | `INT`         |
| Get a remainder           | `MOD`         |
| Generate a random decimal | `RAND`        |
| Generate a random integer | `RANDBETWEEN` |

### The Most Important Beginner Tip

Don't try to memorize the syntax first.

Instead, remember **what problem each function solves**.

```text
"I need a total."       → SUM
"I need an average."    → AVERAGE
"I need the biggest."   → MAX
"I need the smallest."  → MIN
"I need to count."      → COUNT / COUNTA
"I need to round."      → ROUND
"I need the remainder." → MOD
"I need randomness."    → RAND / RANDBETWEEN
```

The syntax becomes much easier to remember through repeated use.

### Quick Reference

```text
SUM
→ Add values

AVERAGE
→ Arithmetic mean

MAX
→ Largest value

MIN
→ Smallest value

COUNT
→ Numeric values

COUNTA
→ Non-empty/content-containing cells

COUNTBLANK
→ Blank cells

ROUND
→ Standard rounding

ROUNDUP
→ Away from zero

ROUNDDOWN
→ Toward zero

INT
→ Toward negative infinity

MOD
→ Remainder

RAND
→ 0 ≤ x < 1, decimal, volatile

RANDBETWEEN
→ Inclusive integer range, volatile
```

A particularly useful set of distinctions is:

```text
ROUNDUP vs. ROUNDDOWN
→ away from zero vs. toward zero

INT vs. TRUNC
→ negative infinity vs. zero

RAND vs. RANDBETWEEN
→ decimal [0, 1) vs. integer [bottom, top]
```

---

## Common Beginner Mistakes

<details>
<summary><strong>1. Forgetting <code>=</code></strong></summary>

Excel formulas normally begin with:

```excel
=
```

For example:

```excel
=SUM(A1:A5)
```

not:

```text
SUM(A1:A5)
```

</details>

<details>
<summary><strong>2. Typing Documentation Notation Literally</strong></summary>

When documentation shows:

```excel
=SUM(number1, [number2], ...)
```

you do not type that literally unless those are actually the values you want.

The notation means:

```text
number1    → required argument
[number2]  → optional argument
...         → additional arguments may be supplied
```

</details>

<details>
<summary><strong>3. Confusing COUNT and COUNTA</strong></summary>

```text
COUNT  → numbers
COUNTA → content
```

Use `COUNT` for numeric observations and `COUNTA` for cells containing something.

</details>

<details>
<summary><strong>4. Thinking INT Simply Removes Decimals</strong></summary>

This is wrong for negative values:

```excel
=INT(-12.9)
```

→ `-13`

If you want:

```text
-12
```

use:

```excel
=TRUNC(-12.9)
```

</details>

<details>
<summary><strong>5. Assuming ROUNDUP Means "Larger"</strong></summary>

`ROUNDUP` means **away from zero**, not "toward positive infinity."

```excel
=ROUNDUP(-12.51, 1)
```

→ `-12.6`

</details>

<details>
<summary><strong>6. Forgetting That Random Functions Recalculate</strong></summary>

Both:

```excel
=RAND()
```

and:

```excel
=RANDBETWEEN(1, 100)
```

can produce new values when Excel recalculates.

If the generated numbers need to remain fixed, convert them to values.

</details>

<details>
<summary><strong>7. Sorting Data Unnecessarily</strong></summary>

If you only need the highest or lowest value, you do not need to sort the dataset.

Use:

```excel
=MAX(range)
=MIN(range)
```

</details>

<details>
<summary><strong>8. Confusing an Empty Cell with Zero</strong></summary>

These are not the same:

```text
(empty)
```

and:

```text
0
```

For example, `AVERAGE` ignores an empty cell but includes `0` in the calculation.

</details>

---

## Summary

In this lesson, we covered several fundamental Excel functions:

* **Calculation:** `SUM`, `AVERAGE`
* **Finding values:** `MAX`, `MIN`
* **Counting:** `COUNT`, `COUNTA`, `COUNTBLANK`
* **Rounding:** `ROUND`, `ROUNDUP`, `ROUNDDOWN`
* **Integer and remainder operations:** `INT`, `MOD`
* **Random number generation:** `RAND`, `RANDBETWEEN`

These functions form a useful foundation for working with numerical data in Excel.

Once these basics are familiar, they can be combined to perform more advanced calculations and data analysis.

The most important step is not memorizing every function. It is learning to recognize **what kind of problem you are trying to solve** and choosing the appropriate function.

---

### References

* [Microsoft Support — Excel functions (alphabetical)](https://support.microsoft.com/en-us/excel/excel-functions-alphabetical)
* [Microsoft Support — Excel functions (by category)](https://support.microsoft.com/en-us/excel/excel-functions-by-category)
