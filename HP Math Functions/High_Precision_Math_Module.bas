Attribute VB_Name = "High_Precision_Math_Module"
  Option Explicit

' A few high precision math functions.

' These functions return results to up to 28 significant digits
' and are primarily designed for special math operations requiring
' more precise values than are built into similar VB functions.

' --------------------------------------
' The functions designed so far include:
'
' Square_Root_Of(ArgX)
' Cube_Root_Of(ArgX)
' SinRad(ArgDeg)
' CosRad(ArgDeg)
' TanRad(ArgDeg)

  Global Const Pi As Variant = "3.1415926535897932384626433832795"

' ====================================================================

' High precision sine (to 28 digits) function for radian arguments.
'
' The argument should be in the range from 0 to ± 2Pi radians.
'
' This function does NOT test for invalid argument input.

  Public Function SinRad(DegArg)
' Level 01

  Dim n   As Integer
  Dim Sum As Variant
  Dim dS  As Variant
  Dim P1  As Variant
  Dim X   As Variant
  Dim Lim As Variant
      
' Initialize variables
   P1 = CDec(1)
    n = 0
  Sum = 0
   dS = CDec(1)
  Lim = CDec(1E-29)
    X = CDec(Trim(DegArg))

' Execute sine computation loop
  Do Until Abs(dS) <= Lim
  dS = P1 * STerm(X, 2 * n + 1)
  Sum = Sum + dS
  n = n + 1
  P1 = -P1
  Loop

  SinRad = Sum

  End Function

' ====================================================================

' High precision cosine (to 28 digits) function for radian arguments.
'
' The argument should be in the range from 0 to ± 2Pi radians.
'
' This function does NOT test for invalid argument input.

  Public Function CosRad(DegArg)
' Level 01

  Dim n   As Integer
  Dim Sum As Variant
  Dim dS  As Variant
  Dim P1  As Variant
  Dim X   As Variant
  Dim Lim As Variant
      
' Initialize variables
   P1 = CDec(1)
    n = 0
  Sum = 0
   dS = CDec(1)
  Lim = CDec(1E-30)
    X = CDec(Trim(DegArg))

' Execute cosine computation loop
  Do Until Abs(dS) <= Lim
  dS = P1 * STerm(X, 2 * n)
  Sum = Sum + dS
  n = n + 1
  P1 = -P1
  Loop

  CosRad = Sum

  End Function

' ====================================================================

' High precision tangent (to 28 digits) function for radian arguments.
'
' The argument should be in the range from 0 to ± 2Pi radians.
' It is possible to oveflow this function.
'
' It simply calls the high precision sine and cosine functions and
' then computes the ratio TanX = SinX / CosX
'
' This function does NOT test for invalid argument input.

  Public Function TanRad(ArgDeg)
' Level 01

  Dim Q As Variant

  Q = CDec(Trim(ArgDeg))

' Compute tangent as Sine/Cosine ratio
  TanRad = SinRad(Q) / CosRad(Q)

  End Function

' ====================================================================

' Compute the value of (x^n) / (n!)
'
' This value occurs frequently in infinite series
' computations and is called by some of the high
' precision math functions.

  Public Function STerm(x_Val, n_Val)
' Level 00

  Dim i  As Integer
  Dim X  As Variant
  Dim Px As Variant
  Dim T  As Variant

      X = CDec(x_Val)
      T = CDec(1)
  For i = 1 To Val(n_Val)
      T = T * X / i
  Next i

  STerm = T

  End Function

' ====================================================================

' Compute square root of (ArgX), by iteration, to up to 29 digits.
' The iteration continues until the limit of precision is reached.
'
' Generally, a square root refers to only positive arguments, but this
' function will accept negative arguments and produce an (imaginary)
' square root by returning a value with " i" attached to the end.
' This has to be checked for prior to using the result in subsequent
' computations as an argument for another function.

  Public Function Square_Root_Of(ArgX)
' Level 00

  Dim X As Variant      ' Argument - Positive or negative value
  
  Dim A As Variant      ' Any general approximation to square root
  Dim B As Variant      ' Next successive approximation to square root
  
  Dim k  As Integer     ' Cycle loop control counter
  Dim i  As String      ' Represents the square root of minus 1
  
' Check for invalid numeric argument
  X = Trim(ArgX): If IsNumeric(X) = False Then GoTo ERROR_HANDLER

  X = CDec(X)  ' Convert argument into decimal data type
  
' Account for a negative argument
  i = "": If X < 0 Then X = -X: i = " i"
  
' Check for zero argument
  If X = 0 Then Square_Root_Of = 0: Exit Function

' Use VB square root as 1st approximation
  A = Sqr(X)
  
'   Very primitive loop to grind out the square root using a series
'   of successive approximations, starting with (A).
    k = 50 ' Set limit of cycles to 50 max - More than enough.
CYCLE:
    B = (A + X / A) / 2 ' Compute next approx (B) from (A)
    
' Check if finished
  If (B = A) Or k <= 0 Then GoTo DONE

' Rinse, lather, repeat until done
  A = B        ' Update approx to current value
  k = k - 1    ' Update limit counter
  GoTo CYCLE
DONE:
  Square_Root_Of = Trim(B & i)
  Exit Function
  
ERROR_HANDLER:
  Square_Root_Of = "ERROR: Invalid numeric argument"
  End Function
  
' ====================================================================

' Compute cube root of (ArgX), by iteration, to up to 29 digits.
' The iteration continues until the limit of precision is reached.
'
' This function is a companion to the SqRoot_Of() function

  Public Function Cube_Root_Of(ArgX)
' Level 00
  
  Dim X  As Variant     ' Argument - May be positive or negative value
  Dim A  As Variant     ' Any general approximation to cube root
  Dim B  As Variant     ' Next successive approximation to cube root
  
  Dim k  As Integer     ' Cycle loop control counter
  Dim Sign As String    ' Sign of argument - Attached to result
      
  A = CDec(0)           ' Initialize (A) as decimal data type
  B = CDec(0)           ' Initialize (B) as decimal data type
  
' Check for invalid numeric argument
  X = Trim(ArgX): If IsNumeric(X) = False Then GoTo ERROR_HANDLER

' Convert argument into decimal data type
  X = CDec(X)

' Consider the sign of the argument
  Sign = "": If X < 0 Then X = -X: Sign = "-"
   
  If X = 0 Then X = 1: GoTo DONE

' Use VB cube root as 1st approximation
  A = X ^ (1 / 3)
   
'   Very primitive loop to grind out the cube root using a series
'   of successive approximations, starting with (A).
    k = 50 ' Set limit of cycles to 50 max - More than enough.
CYCLE:

' Compute next approx (B) from (A)
  B = ((2 * A) + X / (A * A)) / 3
    
' Check if finished
  If (B = A) Or k <= 0 Then GoTo DONE

' Rinse, lather, repeat until done
  A = B        ' Update approx to current value
  k = k - 1    ' Update limit counter
  GoTo CYCLE
DONE:
  Cube_Root_Of = Trim(Sign & B)
  Exit Function
  
ERROR_HANDLER:
  Cube_Root_Of = "ERROR: Invalid numeric argument"
  End Function

' ====================================================================


