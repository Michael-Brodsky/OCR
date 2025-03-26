Attribute VB_Name = "basLibNumeric"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' basLibNumeric                                             '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
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
' basLibArray                                               '
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
Public Const kvbCharBits = 8
Public Const kvbSizeOfByte = 1
Public Const kvbSizeOfInt = 2
Public Const kvbSizeOfLong = 4
Public Const kvbSizeOfDouble = 8
#If VBA7 Then
Public Const kvbSizeOfLongLong = 8
#End If

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
    ByVal n As Byte _
) As Boolean
    '
    ' Returns TRUE if the nth bit of x is set,
    ' else returns FALSE. Bit positions are zero-based.
    '
    Dim mask As Long
    
    mask = 2 ^ n
    BitIsSet = (mask And x) = mask
End Function

#If VBA7 Then
Public Function BytesToLong(aBytes() As Byte) As Long
    '
    ' Returns an array of bytes converted to a two's-complement signed Long.
    '
    Dim i As Integer, j As Integer, n As LongLong
    
    For i = LBound(aBytes) To UBound(aBytes)
        n = n + CLng(aBytes(i)) * 256# ^ j
        j = j + 1
    Next
    BytesToLong = CLng(-(n And &H80000000) + (n And &H7FFFFFFF))
End Function
#End If

Public Function BytesToInt(aBytes() As Byte) As Integer
    '
    ' Returns an array of bytes converted to a two's-complement signed Integer.
    '
    Dim i As Integer, j As Integer, n As Long
    
    For i = LBound(aBytes) To UBound(aBytes)
        n = n + CInt(aBytes(i)) * 256 ^ j
        j = j + 1
    Next
    BytesToInt = CInt(-(n And &H8000) + (n And &H7FFF))
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

Public Function IntToBytes(aInteger As Integer) As Byte()
    '
    ' Returns a signed Integer as a two's-complement array of bytes.
    '
    Dim bytes(kvbSizeOfInt - 1) As Byte
    
    bytes(0) = aInteger And &HFF
    bytes(1) = CInt((aInteger And &HFF00) / 256) And &HFF
    IntToBytes = bytes
End Function

Public Function IPow2Ge( _
    ByVal n As Long _
) As Long
    '
    ' Returns the smallest integral power-of-two equal
    ' to or greater than the absolute value of n.
    '
    IPow2Ge = Int(2 ^ (Int(Log2(Abs(n) - 1)) + 1))
End Function

Public Function IsEven( _
    ByVal n As Long _
) As Boolean
    '
    ' Returns TRUE if n is an even integral number, else returns FALSE.
    '
    IsEven = (n Mod 2 = 0)
End Function

Public Function IsOdd( _
    ByVal n As Long _
) As Boolean
    '
    ' Returns TRUE if n is an odd integral number, else returns FALSE.
    '
    IsOdd = Not IsEven(n)
End Function

Public Function IsPow2( _
    ByVal n As Long _
) As Boolean
    '
    ' Returns TRUE if the absolute value of n is an
    ' integral power of two, else returns FALSE.
    '
    MakeUnsigned n
    IsPow2 = ((n And (n - 1)) = 0)
End Function

Public Function IsSignNe( _
    ByVal A As Double, _
    ByVal b As Double _
) As Boolean
    '
    ' Returns TRUE if a and b have opposite signs,
    ' else returns FALSE.
    '
    IsSignNe = (SignOf(A) <> SignOf(b))
End Function

Public Function LongToBytes(aLong As Long) As Byte()
    '
    ' Returns a signed Long as a two's-complement array of bytes.
    '
    Dim bytes(kvbSizeOfLong - 1) As Byte
    
    bytes(0) = aLong And &HFF
    bytes(1) = ((aLong And &HFF00) / 256) And &HFF
    bytes(2) = ((aLong And &HFF0000) / 256 ^ 2) And &HFF
    bytes(3) = ((aLong And &HFF000000) / 256 ^ 3) And &HFF
    LongToBytes = bytes
End Function

Public Function MakeUnsigned( _
    ByRef n As Variant _
) As Variant
    '
    ' Assigns the absolute value of n to n for procedures
    ' that expect arguments of unsigned types. An error
    ' occurs if n is the underlying type's minimum value
    ' or non-numeric.
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
    ' Returns the largest of the given aValues.
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
        Dim A As Variant
        
        MinMax.First = aValues(LBound(aValues))
        MinMax.Second = MinMax.First
        For Each A In aValues
            If A < MinMax.First Then
                MinMax.First = A
            ElseIf A > MinMax.Second Then
                MinMax.Second = A
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
    ' Returns the smallest of the given aValues.
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
        Dim Count As Long, max_count As Long, Key As Long, i As Long
        
        Count = 1
        Key = aValues(LBound(aValues))
        For i = LBound(aValues) + 1 To UBound(aValues)
            If Key = aValues(i) Then
                Count = Count + 1
            Else
                If Count > max_count Then
                    Mode = aValues(i - 1)
                    max_count = Count
                End If
                Key = aValues(i)
                Count = 1
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
    ByRef n As Double, _
    ByVal aFlag As Boolean _
)
    '
    ' Negates n if aFlag is set.
    '
    n = n * SignOf(aFlag)
End Sub

Public Function Normalize( _
    ByVal x As Double, _
    ByVal xMin As Double, _
    ByVal xMax As Double, _
    ByVal yMin As Double, _
    ByVal yMax As Double _
) As Double
    '
    ' Normalizes x in [xMin,xMax] to x in [yMin,yMax].
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
    ' Returns a random Integer in [aMin, aMax].
    '
    RandI = CInt(Int((CLng(aMax) - CLng(aMin) + 1) * Rnd(aSeed) + CLng(aMin)))
End Function

Public Function RandL( _
    Optional ByVal aSeed As Variant _
) As Long
    '
    ' Returns a random Long in [1, 2^31 - 1]. The given aSeed should be
    ' coprime to kM (see below) else the generator's period may be
    ' severely reduced.
    '
    ' VBA's Rnd() function can only generate random single-precision
    ' floating point numbers in [0,1), which have seven significant
    ' digits, and are insufficent to generate values across the entire
    ' range of Longs. Here, we use a variation of the Park-Miller
    ' algorithm capable of generating outputs from 1 to 2,147,483,646.
    '
    '
    Const kSeed As Long = 32765
    Const kM As Long = &H7FFFFFFF
    Const kA As Long = 48271
    Const kQ As Long = kM / kA     ' 44488
    Const kR As Long = kM Mod kA   '  3399
    Dim div As Long, rm As Long, s As Long, t As Long, result As Long
    Static seed As Long
    
    If Not IsMissing(aSeed) Then seed = aSeed
    If seed = 0 Then seed = kSeed
    div = seed / kQ
    rm = seed Mod kQ
    s = rm * kA
    t = div * kR
    result = s - t
    If result < 0 Then result = result + kM
    seed = result
    RandL = seed
End Function

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
    Dim rng As Pair
    
    rng = MinMax(aValues)
    Range = rng.Second - rng.First
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
    ByVal n As Double _
) As Integer
    '
    ' Returns -1 if n is less than 0, else returns 0.
    '
    Sign = (n < 0)
End Function

Public Function SignOf( _
    ByVal n As Double _
) As Integer
    '
    ' Returns -1 if n is less than 0, else returns +1.
    '
    SignOf = 1 Or Sign(n)
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

