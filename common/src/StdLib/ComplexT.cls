VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComplexT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' ComplexT                                                  '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines an aggregate type that holds a        '
' complex number and functions that operate on complex      '
' numbers.                                                  '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' PairT, LibVBA                                             '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the library has only been tested with     '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

''''''''''''''''''
' Public Members '
''''''''''''''''''

Public Real As Double   ' The real part.
Public Imag As Double   ' The imaginary part.

'''''''''''''''''''''
' Private Interface '
'''''''''''''''''''''

Private Function ModSqr( _
    ByVal aReal As Double, _
    ByVal aImag As Double _
) As Double
    '
    ' Returns the sum of the real and imaginary parts squared.
    '
    ModSqr = aReal ^ 2 + aImag ^ 2
End Function

''''''''''''''''''''
' Public Interface '
''''''''''''''''''''

Public Function Add( _
    aOther As ComplexT _
) As ComplexT
    '
    ' Returns the arithmatic sum of this and another complex number.
    '
    Dim result As ComplexT
    
    Set result = NewComplexT(Real + aOther.Real, Imag + aOther.Imag)
    Set Add = result
    Set result = Nothing
End Function

Public Function Clone() As ComplexT
    '
    ' Returns a deep copy of this complex number.
    '
    Dim c As New ComplexT
    
    c.Real = Real
    c.Imag = Imag
    Set Clone = c
    Set c = Nothing
End Function

Public Function Conjugate() As ComplexT
    '
    ' Returns the complex conjugate of this complex number.
    '
    Dim result As ComplexT
    
    Set result = NewComplexT(Real, -Imag)
    Set Conjugate = result
    Set result = Nothing
End Function

Public Function Diff( _
    aOther As ComplexT _
) As ComplexT
    '
    ' Returns the arithmatic difference of this and another complex number.
    '
    Dim result As ComplexT
    
    Set result = NewComplexT(Real, Imag)
    Set Diff = result.Add(aOther.Neg())
    Set result = Nothing
End Function

Public Function Div( _
    aOther As ComplexT _
) As ComplexT
    '
    ' Returns the arithmatic quotient of this and another complex number.
    '
    Dim result As ComplexT, denom As Double
    
    denom = ModSqr(aOther.Real, aOther.Imag)
    Set result = NewComplexT((Real * aOther.Real + Imag * aOther.Imag) / denom, _
    (Imag * aOther.Real - Real * aOther.Imag) / denom)
    Set Div = result
    Set result = Nothing
End Function

Public Function Eq( _
    aOther As ComplexT _
) As Boolean
    '
    ' Returns TRUE if this complex number is equal to another,
    ' else returns FALSE.
    '
    Eq = (Real = aOther.Real And Imag = aOther.Imag)
End Function

Public Sub InitList( _
    ByVal aReal As Double, _
    ByVal aImag As Double _
)
    '
    ' Initializes this complex number with the given arguments.
    '
    Real = aReal
    Imag = aImag
End Sub

Public Function Inv() As ComplexT
    '
    ' Returns the arithmatic inverse of this complex number.
    '
    Dim result As ComplexT, denom As Double
    
    denom = ModSqr(Real, Imag)
    Set result = NewComplexT(Real / denom, -Imag / denom)
    Set Inv = result
    Set result = Nothing
End Function

Public Function Modulus() As Double
    '
    ' Returns the modulus (absolute value) of this complex number.
    '
    Modulus = Sqr(ModSqr(Real, Imag))
End Function

Public Function Mul( _
    aOther As ComplexT _
) As ComplexT
    '
    ' Returns the product of this and another complex number.
    '
    Dim result As ComplexT
    
    Set result = NewComplexT(Real * aOther.Real - (Imag * aOther.Imag), Real * aOther.Imag + Imag * aOther.Real)
    Set Mul = result
    Set result = Nothing
End Function

Public Function Ne( _
    aOther As ComplexT _
) As Boolean
    '
    ' Returns TRUE if this complex number is not equal to another,
    ' else returns FALSE.
    '
    Ne = Not Eq(aOther)
End Function

Public Function Neg() As ComplexT
    '
    ' Returns this complex number negated.
    '
    Dim result As ComplexT
    
    Set result = NewComplexT(-Real, -Imag)
    Set Neg = result
    Set result = Nothing
End Function

Public Function CPolar() As ComplexT
    '
    ' Returns the polar coordinates of this complex number.
    '
    Dim result As ComplexT, p As PairT
    
    Set p = Polar(Real, Imag)
    Set result = NewComplexT(p.First * Cos(p.Second), p.First * Sin(p.Second))
    Set CPolar = result
    Set result = Nothing
End Function

Public Property Get RValue() As Double
    '
    ' Returns the imaginary part if it's non-zero,
    ' else returns the real part.
    '
    RValue = IIf(Imag <> 0, Imag, Real)
End Property

Public Sub Swap( _
    ByRef aOther As ComplexT _
)
    '
    ' Exchanges this complex number with another.
    '
    Dim tmp As Double
    
    tmp = aOther.Real
    aOther.Real = Real
    Real = tmp
    tmp = aOther.Imag
    aOther.Imag = Imag
    Imag = tmp
End Sub

Public Function ToString( _
    Optional ByVal aSigFigs As Integer = 3 _
) As String
    '
    ' Returns a string representation of this complex number
    ' in the form a + bi, optionally with the specified number
    ' of significant figures, upto 15.
    '
    Dim re As String, im As String, op As String
    
    re = IIf(Real <> 0, FmtSigFigs(Real, aSigFigs), IIf(Imag = 0, "0", ""))
    op = IIf(Real <> 0, IIf(Imag < 0, " - ", IIf(Imag > 0, " + ", "")), "")
    im = IIf(Imag <> 0, IIf(Imag <> 1, FmtSigFigs(IIf(Len(op) = 0, Imag, Abs(Imag)), aSigFigs), "") & "i", "")
    ToString = re & op & im
End Function

    

