VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AllocatorT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' AllocatorT                                                '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a class used by library container     '
' classes to allocate storage space.                        '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' LibVBA, LibNumeric                                        '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the library has only been tested with     '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

'''''''''''''''''''
' Private Members '
'''''''''''''''''''

Private Buffer As Variant    ' Data storage buffer.

''''''''''''''''''''
' Public Interface '
''''''''''''''''''''

Public Sub Allocate( _
    ByVal aSize As Long _
)
    '
    ' Allocates the smallest power-of-two elements
    ' greater than or equal to aSize.
    '
    MakeUnsigned aSize
    If aSize > Capacity Then
        Select Case aSize
            Case 0, 1:
                Resize aSize
            Case Else:
                Resize IPow2Ge(aSize)
        End Select
    End If
End Sub

Public Property Get Capacity() As Long
    '
    ' Returns the current buffer size in elements.
    '
    Capacity = ArraySize(Buffer)
End Property

Public Property Get Data( _
    Optional aIndex As Variant _
) As Variant
    '
    ' Returns a buffer element, or the entire buffer
    ' if aIndex is omitted.
    '
    If Not IsMissing(aIndex) Then
        Dim n
        
        MakeUnsigned aIndex
        n = LBound(Buffer) + CLng(aIndex)
        If IsObject(Buffer(n)) Then
            Set Data = Buffer(n)
        Else
            Data = Buffer(n)
        End If
    Else
        If IsObject(Buffer) Then
            Set Data = Buffer
        Else
            Data = Buffer
        End If
    End If
End Property

Public Property Let Data( _
    Optional aIndex As Variant, _
    aData As Variant _
)
    '
    ' Assigns a buffer element, or the entire buffer
    ' if aIndex is omitted.
    '
    If Not IsMissing(aIndex) Then
        MakeUnsigned aIndex
        Buffer(LBound(Buffer) + CLng(aIndex)) = aData
    Else
        Buffer = aData
    End If
End Property

Public Property Set Data( _
    Optional aIndex As Variant, _
    aData As Variant _
)
    '
    ' Assigns a buffer element, or the entire buffer
    ' if aIndex is omitted.
    '
    If Not IsMissing(aIndex) Then
        MakeUnsigned aIndex
        Set Buffer(LBound(Buffer) + CLng(aIndex)) = aData
    Else
        Set Buffer = aData
    End If
End Property

Public Sub Deallocate()
    '
    ' Deallocates the buffer and its contents.
    '
    If Not IsEmpty(Buffer) Then ArrayResize Buffer, 0
End Sub

Public Sub Resize( _
    ByVal aSize As Long _
)
    '
    ' Resizes the buffer to aSize elements.
    '
    MakeUnsigned aSize
    ArrayResize Buffer, aSize
End Sub

