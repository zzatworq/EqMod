TEST_STRING = r"""
To simplify the expression:

Test equation: \[E=mc^{2^{2}}\] and another \( \int_0^1 x^2 dx = \frac{1}{3} \). Normal text follows: \[\sum M_A = 0\].

To simplify the expression:

\[
\left( \frac{20}{x^2 - 36} - \frac{2}{x - 6} \right) \times \frac{1}{4 - x}
\]

we follow these steps:

### Step 1: Factor the Denominators
First, factor the denominators where possible.

\[
x^2 - 36 = (x - 6)(x + 6)
\]

So, the expression becomes:

\[
\left( \frac{20}{(x - 6)(x + 6)} - \frac{2}{x - 6} \right) \times \frac{1}{4 - x}
\]

### Step 2: Combine the Fractions Inside the Parentheses
To combine the fractions, find a common denominator, which is \((x - 6)(x + 6)\).

\[
\frac{20}{(x - 6)(x + 6)} - \frac{2}{x - 6} = \frac{20 - 2(x + 6)}{(x - 6)(x + 6)}
\]

Simplify the numerator:

\[
20 - 2(x + 6) = 20 - 2x - 12 = 8 - 2x
\]

So, the combined fraction is:

\[
\frac{8 - 2x}{(x - 6)(x + 6)}
\]

### Step 3: Factor the Numerator
Factor out a common term from the numerator:

\[
8 - 2x = 2(4 - x)
\]

Now, the expression becomes:

\[
\frac{2(4 - x)}{(x - 6)(x + 6)} \times \frac{1}{4 - x}
\]

### Step 4: Simplify the Expression
Notice that \((4 - x)\) appears in both the numerator and the denominator, so they cancel out:

\[
\frac{2}{(x - 6)(x + 6)}
\]

### Final Answer
The simplified form of the expression is:

\[
\boxed{\frac{2}{(x - 6)(x + 6)}}
\]
that i final
"""