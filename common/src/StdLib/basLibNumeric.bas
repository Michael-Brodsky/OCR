Attribute VB_Name = "basLibNumeric"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' LibNumeric                                                '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a library of arithmatic, statistical  '
' and bit-wise functions.                                   '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' LibArray                                                  '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the library has only been tested with     '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

''''''''''''''''''''
' Numerical Limits '
''''''''''''''''''''

Public Const kvbByteMax As Byte = 255
Public Const kvbByteMin As Byte = 0
Public Const kvbIntegerMax As Integer = 32767
Public Const kvbIntegerMin As Integer = -32768
Public Const kvbLongMax As Long = 2147483647
Public Const kvbLongMin As Long = -2147483648#
#If VBA7 Then
Public Const kvbLongLongMax As LongLong = 9.22337203685477E+18
Public Const kvbLongLongMin As LongLong = -9.22337203685477E+18
#End If
Public Const kvbSingleMax As Single = 3.402823E+38!
Public Const kvbSingleMin As Single = 1.401298E-45!
Public Const kvbDoubleMax As Double = 1.79769313486231E+308
Public Const kvbDoubleMin As Double = -1.79769313486231E+308
Public Const kvbEpsilon As Double = 1E-21

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

Public Function BitIsSet( _
    ByVal x As Long, _
    ByVal aBit As Byte _
) As Boolean
    '
    ' Returns TRUE if the nth bit of x is set,
    ' else returns FALSE. Bit positions are zero-based.
    '
    Dim mask As Long
    
    mask = 2 ^ aBit
    BitIsSet = (mask And x) = mask
End Function

Public Function Constrain( _
    ByVal x As Double, _
    ByVal aMin As Double, _
    ByVal aMax As Double _
) As Double
    '
    ' Returns x constrained between aMin and aMax. The
    ' function is ill-formed if aMin is greater than aMax.
    '
    Constrain = IIf(x < aMin, aMin, IIf(x > aMax, aMax, x))
End Function

Public Function Factorial( _
    ByVal n As Byte _
) As Double
    ' Returns n!
    If n = 0 Or n = 1 Then
        Factorial = 1
    Else
        Factorial = Factorial(n - 1) * n
    End If
End Function

Public Function FibN( _
    ByVal n As Byte _
) As Double
    '
    ' Returns the nth Fibonacci number, excluding zero for n > 0,
    ' using Binet's formula.
    '
    Const k As Double = 2.23606797749979    ' ~ Sqr(5)
    
    FibN = Round((kPhi ^ n - (1 - kPhi) ^ n) / k, 0)
End Function

Public Function Fibonacci( _
    ByVal n As Double _
) As Variant()
    '
    ' Returns the Fibonacci sequence upto and including n
    ' as an array, where n is a positive integral value.
    '
    Dim result() As Variant, i As Long, y As Double
    
    n = Int(Abs(n))
    y = 1
    If n >= 0 Then ArrayPushBack result, 0
    While y <= n
        ArrayPushBack result, y
        y = y + CDbl(result(i))
        i = i + 1
    Wend
    Fibonacci = result
End Function

Public Function IPow2Ge( _
    ByVal n As Long _
) As Long
    '
    ' Returns the smallest integral power-of-two equal
    ' to or greater than the absolute value of x.
    '
    IPow2Ge = Int(2 ^ (Int(Log2(Abs(n) - 1)) + 1))
End Function

Public Function IsEven( _
    ByVal n As Long _
) As Boolean
    '
    ' Returns TRUE if x is an even integral number, else returns FALSE.
    '
    IsEven = (n Mod 2 = 0)
End Function

Public Function IsOdd( _
    ByVal n As Long _
) As Boolean
    '
    ' Returns TRUE if x is an odd integral number, else returns FALSE.
    '
    IsOdd = Not IsEven(n)
End Function

Public Function IsPow2( _
    ByVal n As Long _
) As Boolean
    '
    ' Returns TRUE if the absolute value of x is an
    ' integral power of two, else returns FALSE.
    '
    MakeUnsigned n
    IsPow2 = ((n And (n - 1)) = 0)
End Function

Public Function IsSignNe( _
    ByVal a As Double, _
    ByVal b As Double _
) As Boolean
    '
    ' Returns TRUE if a and b have opposite signs,
    ' else returns FALSE.
    '
    IsSignNe = (SignOf(a) <> SignOf(b))
End Function

Public Function MakeUnsigned( _
    ByRef n As Variant _
) As Variant
    '
    ' Assigns the absolute value of x to x for procedures
    ' that expect arguments of unsigned integral types. An
    ' error occurs if x is the underlying type's minimum
    ' value or non-numeric.
    '
    n = Abs(n)
End Function

Public Function Max( _
    aValues As Variant _
) As Variant
    '
    ' Returns the largest element in a one-dimensional array
    ' using the greater than > operator. No type checking is
    ' performed. The behavior is undefined if aValues is not a
    ' one-dimensional array or the array elements are not
    ' primitive types.
    '
    If ArraySize(aValues) > 0 Then
        Dim v As Variant
        
        Max = aValues(LBound(aValues))
        For Each v In aValues
            If v > Max Then Max = v
        Next
    ElseIf Not IsArray(aValues) Then
         Max = aValues
    End If
End Function

Public Function MaxOf( _
    ParamArray aValues() As Variant _
) As Variant
    '
    ' Returns the largest value from the given aValues.
    '
    If Not IsMissing(aValues) Then MaxOf = Max(CVar(aValues))
End Function

Public Function Mean( _
    aValues As Variant _
) As Variant
    '
    ' Returns the arithmatic mean of elements in a one-
    ' dimensional array. No type checking is performed.
    ' The behavior is undefined if aValues is not a one-
    ' dimensional array or the array elements are non-
    ' numeric.
    '
    Mean = Sum(aValues) / ArraySize(aValues)
End Function

Public Function MeanOf( _
    ParamArray aValues() As Variant _
) As Variant
    '
    ' Returns the arithmatic mean of the given aValues.
    '
    If Not IsMissing(aValues) Then MeanOf = Mean(CVar(aValues))
End Function

Public Function Min( _
    aValues As Variant _
) As Variant
    '
    ' Returns the smallest element in a one-dimensional array
    ' using the less than < operator. No type checking is
    ' performed. The behavior is undefined aValues is not a
    ' one-dimensional array or the array elements are not
    ' primitive types.
    '
    If ArraySize(aValues) > 0 Then
        Dim v As Variant
        
        Min = aValues(LBound(aValues))
        For Each v In aValues
            If v < Min Then Min = v
        Next
    ElseIf Not IsArray(aValues) Then
         Min = aValues
    End If
End Function

Public Function MinMax( _
    aValues As Variant _
) As Pair
    '
    ' Returns the smallest and largest elements in a one-
    ' dimensional array as a Pair, where .First = smallest
    ' and .Second = largest. No type checking is performed.
    ' The behavior is undefined if aValues is not a one-
    ' dimensional array or the array elements are not
    ' primitive types.
    '
    If ArraySize(aValues) > 0 Then
        Dim a As Variant
        
        MinMax.First = aValues(LBound(aValues))
        MinMax.Second = MinMax.First
        For Each a In aValues
            If a < MinMax.First Then
                MinMax.First = a
            ElseIf a > MinMax.Second Then
                MinMax.Second = a
            End If
        Next
    ElseIf Not IsArray(aValues) Then
         MinMax.First = aValues
         MinMax.Second = MinMax.First
    End If
End Function

Public Function MinMaxOf( _
    ParamArray aValues() As Variant _
) As Pair
    '
    ' Returns the smallest and largest values from the given aValues
    ' as a Pair, where .First = smallest, .Second = largest.
    '
    If Not IsMissing(aValues) Then MinMaxOf = MinMax(CVar(aValues))
End Function

Public Function MinOf( _
    ParamArray aValues() As Variant _
) As Variant
    '
    ' Returns the smallest value from the given aValues.
    '
    If Not IsMissing(aValues) Then MinOf = Min(CVar(aValues))
End Function

Public Function Mode( _
    aValues As Variant _
) As Variant
    '
    ' Returns the arithmatic mode of elements in a one-
    ' dimensional array. No type checking is performed.
    ' An error occurs if the elements are non-numeric types.
    ' The behavior is undefined if aValues is not a one-
    ' dimensional array or the array elements are not
    ' primitive types.
    '
    If ArraySize(aValues) > 0 Then
        Dim count As Long, max_count As Long, key As Long, i As Long
        
        count = 1
        key = aValues(LBound(aValues))
        For i = LBound(aValues) + 1 To UBound(aValues)
            If key = aValues(i) Then
                count = count + 1
            Else
                If count > max_count Then
                    Mode = aValues(i - 1)
                    max_count = count
                End If
                key = aValues(i)
                count = 1
            End If
        Next
        If max_count < 2 Then Mode = 0
    End If
End Function

Public Function ModeOf( _
    ParamArray aValues() As Variant _
) As Variant
    '
    ' Returns the arithmatic mode of the given aValues.
    '
    If Not IsMissing(aValues) Then ModeOf = Mode(CVar(aValues))
End Function

Public Sub NegateIf( _
    ByRef x As Double, _
    ByVal aFlag As Boolean _
)
    '
    ' Negates x if aFlag is set.
    '
    x = x * SignOf(aFlag)
End Sub

Public Function Normalize( _
    ByVal x As Double, _
    ByVal xMin As Double, _
    ByVal xMax As Double, _
    ByVal yMin As Double, _
    ByVal yMax As Double _
) As Double
    '
    ' Normalizes x in [xmin,xmax] to x in [ymin,ymax].
    ' The function is ill-formed if xMin >= xMax or
    ' yMin >= yMax.
    '
    Normalize = ((yMax - yMin) / (xMax - xMin) * (x - xMax) + yMax)
End Function

Public Function RandI( _
    Optional ByVal aMin As Integer = kvbIntegerMin, _
    Optional ByVal aMax As Integer = kvbIntegerMax, _
    Optional ByVal aSeed As Variant _
) As Integer
    '
    ' Returns a random Integer between between aMin and aMax inclusive.
    '
    RandI = CInt(Int((CLng(aMax) - CLng(aMin) + 1) * Rnd(aSeed) + CLng(aMin)))
End Function

#If VBA7 Then
Public Function RandL( _
    Optional ByVal aMin As Long = kvbLongMin, _
    Optional ByVal aMax As Long = kvbLongMax, _
    Optional ByVal aSeed As Variant _
) As Long
    '
    ' Returns a random Long between aMin and aMax inclusive.
    '
    RandL = CLng(Int((CDbl(aMax) - CDbl(aMin) + 1) * Rnd(aSeed) + CDbl(aMin)))
End Function
#End If

Public Function Range( _
    aValues As Variant _
) As Variant
    '
    ' Returns the arithmatic range of elements in a one-
    ' dimensional array. No type checking is performed.
    ' An error occurs if the elements are non-numeric types.
    ' The behavior is undefined if aValues is not a one-
    ' dimensional array or the array elements are not
    ' primitive types.
    '
    Dim r As Pair
    
    r = MinMax(aValues)
    Range = r.Second - r.First
End Function

Public Function RangeOf( _
    ParamArray aValues() As Variant _
) As Variant
    '
    ' Returns the arithmatic range of the given aValues.
    '
    If Not IsMissing(aValues) Then RangeOf = Range(CVar(aValues))
End Function

Public Function Sign( _
    ByVal x As Double _
) As Integer
    '
    ' Returns -1 if x is less than 0, else returns 0.
    '
    Sign = (x < 0)
End Function

Public Function SignOf( _
    ByVal x As Double _
) As Integer
    '
    ' Returns -1 if x is less than 0, else returns +1.
    '
    SignOf = 1 Or Sign(x)
End Function

Public Function StdDev( _
    aValues As Variant _
) As Variant
    '
    ' Returns the statistical standard deviation of elements
    ' in a one-dimensional array. No type checking is performed.
    ' The behavior is undefined if aValues is not a one-
    ' dimensional array or the array elements are non-numeric.
    '
    StdDev = Sqr(Variance(aValues))
End Function

Public Function StdDevOf( _
    ParamArray aValues() As Variant _
) As Variant
    '
    ' Returns the statistical standard deviation of the given aValues.
    '
    If Not IsMissing(aValues) Then StdDevOf = StdDev(CVar(aValues))
End Function

Public Function Sum( _
    aValues As Variant _
) As Variant
    '
    ' Returns the arithmatic sum of elements in a one-
    ' dimensional array. No type checking is performed.
    ' The behavior is undefined if aValues is not a one-
    ' dimensional array or the array elements are non-
    ' numeric.
    '
    Dim v As Variant
    
    For Each v In aValues
        Sum = Sum + v
    Next
End Function

Public Function Variance( _
    aValues As Variant _
) As Variant
    '
    ' Returns the statistical variance of elements in a
    ' one-dimensional array. No type checking is performed.
    ' The behavior is undefined if aValues is not a one-
    ' dimensional array or the array elements are non-
    ' numeric.
    '
    If ArraySize(aValues) > 1 Then
        Dim var As Double, avg As Double, v As Variant
        
        avg = Mean(aValues)
        For Each v In aValues
            var = var + (v - avg) * (v - avg)
        Next
        Variance = var / ArraySize(aValues)
    End If
End Function

Public Function VarianceOf( _
    ParamArray aValues() As Variant _
) As Variant
    '
    ' Returns the statistical variance of the given aValues.
    '
    If Not IsMissing(aValues) Then VarianceOf = Variance(CVar(aValues))
End Function

