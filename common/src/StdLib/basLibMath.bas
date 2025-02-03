Attribute VB_Name = "basLibMath"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' LibMath                                                   '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a library of floating point           '
' mathematical functions.                                   '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' (None)                                                    '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the library has only been tested with     '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

' Enumerates valid values for the Polygon argument in
' procedures that take it.
Public Enum Polygon
    poCircle = 0
    poEllipse = 1
    poRhombus = 2
    poTriangle = 3
    poRectangle = 4
    poPentagon = 5
    poHexagon = 6
    poHeptagon = 7
    poOctagon = 8
    poNonagon = 9
    poDecagon = 10
    poHendecagon = 11
    poDodecagon = 12
End Enum

'''''''''''''''''''''''''''''
' Common Physical Constants '
'''''''''''''''''''''''''''''

Public Const kAvogadro As Double = 6.02214076E+23
Public Const kBoltzmann As Double = 1.380649E-23
Public Const kC As Double = 299792458#
Public Const kCharge As Double = 1.602176634E-19
Public Const kDegsPerRad As Double = 57.2957795130823
Public Const kE As Double = 2.71828182845905
Public Const kEConst As Double = 0.577215664901532
Public Const kFineStruct As Double = 0.0072973525643
Public Const kG As Double = 0.000000000066743
Public Const kGauss As Double = 0.834626841674073
Public Const kOmega As Double = 0.567143290409783
Public Const kPhi As Double = 1.61803398874989
Public Const kPi As Double = 3.14159265358979
Public Const kPlanck As Double = 6.62607015E-34
Public Const kvbTwipsPerInch As Long = 1440

'''''''''''''''''''''
' Library Functions '
'''''''''''''''''''''

Public Function ArcCos( _
    ByVal a As Double _
) As Double
    '
    ' Returns the inverse cosine of a in radians, where
    ' a in [-1,1].
    '
    ArcCos = Atn(-a / Sqr(-a * a + 1)) + 2 * Atn(1)
End Function

Public Function ArcCot( _
    ByVal a As Double _
) As Double
    '
    ' Returns the inverse cotangent of a in radians, where
    ' a in (-inf, inf).
    '
    ArcCot = Atn(-a) + 2 * Atn(1)
End Function

Public Function ArcCsc( _
    ByVal a As Double _
) As Double
    '
    ' Returns the inverse cosecant of a in radians, where
    ' a in (-inf ,-1]U[1,inf).
    '
    ArcCsc = Atn(a / Sqr(a * a - 1)) + (Sgn(a) - 1) * (2 * Atn(1))
End Function

Public Function ArcSec( _
    ByVal a As Double _
) As Double
    '
    ' Returns the inverse secant of a in radians, where
    ' a in (-inf ,-1]U[1,inf).
    '
    ArcSec = Atn(a / Sqr(a * a - 1)) + Sgn((a) - 1) * (2 * Atn(1))
End Function

Public Function ArcSin( _
    ByVal a As Double _
) As Double
    '
    ' Returns the inverse sine of a in radians, where
    ' a in [-1,1]
    '
    ArcSin = Atn(a / Sqr(-a * a + 1))
End Function

Public Function Area( _
    ByVal a As Double, _
    ByVal aPolygon As Polygon, _
    Optional ByVal b As Double = 0 _
) As Double
    '
    ' Returns the area of a regular polygon of 3 to 12 sides.
    ' Three and four-sided polygons take a and optionally the b
    ' argument when base/height differs, all others take only
    ' the a argument. a=base (width), b=height for all polygons.
    '
    Select Case aPolygon
        Case poCircle:
            Area = kPi * a * a
        Case poEllipse:
            If b = 0 Then
                Area = kPi * a * a
            Else
                Area = kPi * a * b
            End If
        Case poTriangle, poRhombus:
            If b = 0 Then
                Area = 0.5 * a * a
            Else
                Area = 0.5 * a * b
            End If
        Case poRectangle:
            If b = 0 Then
                Area = a * a
            Else
                Area = a * b
            End If
        Case poPentagon:
            Area = 1.72047740058897 * a * a ' 0.25 * Sqr(5 * (5 + 2 * Sqr(5))) ~ 1.72047740058897
        Case poHexagon:
            Area = 2.59807621135332 * a * a ' 3 * Sqr(3) / 2 ~ 2.59807621135332
        Case poHeptagon:
            Area = 3.63391244400159 * a * a ' 7 / 4 * Cot(kPi / 7) ~ 3.63391244400159
        Case poOctagon:
            Area = 4.82842712474619 * a * a ' 2 * (1 + Sqr(2)) ~ 4.82842712474619
        Case poNonagon:
            Area = 6.18182419377291 * a * a ' 9 / 4 * Cot(kPi / 9) ~ 6.18182419377291
        Case poDecagon:
            Area = 7.69420884293813 * a * a ' 2.5 * Sqr(5 + 2 * Sqr(5)) ~ 7.69420884293813
        Case poHendecagon:
            Area = 9.36563990694545 * a * a ' 11 / 4 * Cot(kPi / 11) ~ 9.365639907
        Case poDodecagon:
            Area = 11.1961524227066 * a * a ' 3 * (2 + Sqr(3)) ~ 11.1961524227066
        Case Else:
            Err.Raise kvbErrDbInvalidArgument, "LibVBA", "Invalid polygon"
    End Select
End Function

Public Function Cosh( _
    ByVal a As Double _
) As Double
    '
    ' Returns the hyperbolic cosine of a in radians, where
    ' a in (-inf,inf).
    '
    Cosh = (Exp(a) + Exp(-a)) / 2
End Function

Public Function Cot( _
    ByVal a As Double _
) As Double
    '
    ' Returns the cotangent of a in radians, where
    ' a in R -> (n)pi.
    '
    Cot = 1 / Tan(a)
End Function

Public Function Csc( _
    ByVal a As Double _
) As Double
    '
    ' Returns the cosecant of a in radians, where
    ' a in R -> (n)pi.
    '
    Csc = 1 / Sin(a)
End Function

Public Function Degrees( _
    ByVal x As Double _
) As Double
    '
    ' Returns x radians of arc converted to degrees.
    '
    Degrees = x * kDegsPerRad
End Function

Public Function Hypot( _
    ByVal a As Double, _
    ByVal b As Double, _
    Optional ByVal gamma As Double = kPi / 2 _
) As Double
    '
    ' Returns side c of any triangle with sides a and b and
    ' angle gamma in radians opposite to c, using the Law of
    ' Cosines. Gamma can be ommited for right triangles, in
    ' which case the function reduces to Pythagoras.
    '
    Hypot = Sqr(a * a + b * b - 2 * a * b * Cos(gamma))
End Function

Public Function Log10( _
    ByVal x As Double _
) As Variant
    '
    ' Returns the base 10 log of x.
    '
    Log10 = Log(x) / Log(10)
End Function

Public Function Log2( _
    ByVal x As Double _
) As Variant
    '
    ' Returns the base 2 log of x.
    '
    Log2 = Log(x) / Log(2)
End Function

Public Function LogB( _
    ByVal x As Double, _
    ByVal b As Double _
) As Variant
    '
    ' Returns the base b log of x.
    '
    LogB = Log(x) / Log(b)
End Function

Public Function Quadratic( _
    ByVal a As Double, _
    ByVal b As Double, _
    ByVal c As Double _
) As PairT
    '
    ' Returns the roots of a quadratic function as a pair of complex numbers.
    '
    Dim disc As Double, num As ComplexT, denom As Double, root1 As ComplexT, root2 As ComplexT, result As PairT
    
    disc = b ^ 2 - 4 * a * c
    Set num = IIf(disc < 0, Sqrt(-disc), Sqrt(disc))
    denom = 2 * a
    If disc < 0 Then
        Set root1 = NewComplexT(-b / denom, num.RValue / denom)
        Set root2 = NewComplexT(-b / denom, -num.RValue / denom)
    Else
        Set root1 = NewComplexT((-b + num.RValue) / denom, 0)
        Set root2 = NewComplexT((-b - num.RValue) / denom, 0)
    End If
    Set result = NewPairT(root1, root2)
    Set Quadratic = result
    Set result = Nothing
End Function

Public Function Radians( _
    ByVal x As Double _
) As Double
    '
    ' Returns x degrees of arc converted to radians.
    '
    Radians = x / kDegsPerRad
End Function

Public Function Sec( _
    ByVal a As Double _
) As Double
    '
    ' Returns the secant of a in radians, where.
    ' a in R -> (2n+1)pi/2
    '
    Sec = 1 / Cos(a)
End Function

Public Function Sinh( _
    ByVal a As Double _
) As Double
    '
    ' Returns the hyperbolic sine of a in radians, where
    ' a in (-inf, inf).
    '
    Sinh = (Exp(a) - Exp(-a)) / 2
End Function

Public Function Sqrt( _
    ByVal x As Double _
) As ComplexT
    '
    ' Returns the square root of any number as a complex number.
    '
    Dim result As ComplexT
    
    If x < 0 Then
        Set result = NewComplexT(0, Sqr(Abs(x)))
    Else
        Set result = NewComplexT(Sqr(x), 0)
    End If
    Set Sqrt = result
    Set result = Nothing
End Function

Public Function Tanh( _
    ByVal a As Double _
) As Double
    '
    ' Returns the hyperbolic tangent of a in radians, where
    ' a in (-inf, inf).
    '
    Tanh = (Exp(a) - Exp(-a)) / (Exp(a) + Exp(-a))
End Function


