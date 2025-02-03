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
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines an aggregate type that holds a        '
' complex number.                                           '
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

''''''''''''''''''
' Public Members '
''''''''''''''''''

Public Real As Double   ' The real part.
Public Imag As Double   ' The imaginary part.

Public Function Clone() As ComplexT
    '
    ' Returns a deep copy of the complex number.
    '
    Dim c As New ComplexT
    
    c.Real = Real
    c.Imag = Imag
    Set Clone = c
    Set c = Nothing
End Function

Public Sub InitList( _
    ByVal aReal As Double, _
    ByVal aImag As Double _
)
    '
    ' Initializes a complex number with the given arguments.
    '
    Real = aReal
    Imag = aImag
End Sub

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
    ' Exchanges a complex number with another.
    '
    Dim tmp As Double
    
    tmp = aOther.Real
    aOther.Real = Real
    Real = tmp
    tmp = aOther.Imag
    aOther.Imag = Imag
    Imag = tmp
End Sub
    

