# CheckKOQ
Checks Fortran 90 source-code, to verify the correctness of kind-of-quantity (KOQ) in scientific programs
CheckKOQ outputs the line numbers and variables which have wrong KOQ
For definitions and examples see ISO 80000-1:2009 Quantities and units -- Part 1: General
This is a PROOF-OF-CONCEPT extension to Camfort, a Fortran source-code units-of-measure checker described in:
M. Contrastin, et al., "Units-of-Measure Correctness in Fortran Programs", Computing in Science & Engineering, 18, 102-107, 2016.
The Fortran source must be annotated with comments as follows:
!= KOQRelation :: KOQ1 = KOQ2 [*|/] KOQ3 [*|/] KOQ4...
!= KOQ <KOQ1 > :: variable1 [, variable2]
This version requires Fortran source to have: no labels; no parenthesis in assignment expressions; full spacing
TO DO:
.lots of error checking
.size arrays dynamically to hold the maximum number of tokens on a line (or make a more compact data structure)
.add parsing to deal with parenthesis in assignment expressions
.add parsing to ignore labels
.split the unary operator '**2' into a binary operator '** 2', to keep the number of tokens in Varmath/KOQMath even
.add space padding either side of operators +,-,/,*
