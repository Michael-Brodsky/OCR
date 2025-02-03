VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PairT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' PairT                                                     '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a 2-tuple class.                      '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' None                                                      '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the library has only been tested with     '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database
Option Explicit

Public First As Variant
Public Second As Variant

Public Property Get ToArray() As Variant
    '
    ' Returns first and second as a two-element array.
    '
    Dim a(1) As Variant
    
    a(0) = First
    a(1) = Second
    ToArray = a
End Property

Public Function Clone() As PairT
    '
    ' Returns a deep copy of the pair members.
    '
    Dim p As New PairT
    
    If IsObject(First) Then
        Set p.First = First
    Else
        p.First = First
    End If
    If IsObject(Second) Then
        Set p.Second = Second
    Else
        p.Second = Second
    End If
    Set Clone = p
    Set p = Nothing
End Function

Public Sub InitList( _
    aFirst As Variant, _
    aSecond As Variant _
)
    '
    ' Initializes the pair with the given values.
    '
    If IsObject(aFirst) Then
        Set First = aFirst
    Else
        First = aFirst
    End If
    If IsObject(aSecond) Then
        Set Second = aSecond
    Else
        Second = aSecond
    End If
End Sub

Public Sub Swap( _
    ByRef aOther As PairT _
)
    '
    ' Exchanges the pair contents with that of another.
    '
    Dim tmp As Variant
    
    If IsObject(aOther.First) Then
        Set tmp = aOther.First
    Else
        tmp = aOther.First
    End If
    If IsObject(First) Then
        Set aOther.First = First
    Else
        aOther.First = First
    End If
    If IsObject(tmp) Then
        Set First = tmp
    Else
        First = tmp
    End If
    If IsObject(aOther.Second) Then
        Set tmp = aOther.Second
    Else
        tmp = aOther.Second
    End If
    If IsObject(Second) Then
        Set aOther.Second = Second
    Else
        aOther.Second = Second
    End If
    If IsObject(tmp) Then
        Set Second = tmp
    Else
        Second = tmp
    End If
End Sub



