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
' This module defines a library of non-mathematical         '
' numeric and bit-wise functions.                           '
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
) As Variant
    '
    ' Returns x constrained between aMin and aMax. The
    ' behavior is undefined if aMin is greater than aMax.
    '
    Constrain = IIf(x < aMin, aMin, IIf(x > aMax, aMax, x))
End Function

Public Function IPow2Ge( _
    ByVal x As Long _
) As Long
    '
    ' Returns the smallest integral power-of-two equal
    ' to or greater than the absolute value of x.
    '
    IPow2Ge = Int(2 ^ (Int(Log2(Abs(x) - 1)) + 1))
End Function

Public Function IsEven( _
    ByVal x As Long _
) As Boolean
    '
    ' Returns TRUE if x is an even integral number, else returns FALSE.
    '
    IsEven = (x Mod 2 = 0)
End Function

Public Function IsOdd( _
    ByVal x As Long _
) As Boolean
    '
    ' Returns TRUE if x is an odd integral number, else returns FALSE.
    '
    IsOdd = Not IsEven(x)
End Function

Public Function IsPow2( _
    ByVal x As Long _
) As Boolean
    '
    ' Returns TRUE if a the absolute value of x is an
    ' integral power of two, else returns FALSE.
    '
    MakeUnsigned x
    IsPow2 = ((x And (x - 1)) = 0)
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
    ByRef x As Variant _
) As Variant
    '
    ' Assigns the absolute value of x to x for procedures
    ' that expect arguments of unsigned integral types. An
    ' error occurs if x is the underlying type's minimum
    ' value or non-numeric.
    '
    x = Abs(x)
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

