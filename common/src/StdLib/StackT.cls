VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StackT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' StackT                                                    '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a LIFO abstract data type.            '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' ContainerT                                                '
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

Private Container As New ContainerT ' The underlying storage container.

'''''''''''''''''''''''''''''
' Private Member Accessors  '
'                           '
' Workaround for VBA's      '
' inability to access       '
' private class members     '
' in other class instances. '
'''''''''''''''''''''''''''''

Friend Property Get MyContainer() As ContainerT
    Set MyContainer = Container
End Property

Friend Property Set MyContainer(aContainer As ContainerT)
    Set Container = aContainer
End Property

''''''''''''''''''''
' Public Interface '
''''''''''''''''''''

Public Property Get IsEmpty() As Boolean
    '
    ' Returns TRUE if the stack is empty, else returns FALSE.
    '
    IsEmpty = Container.IsEmpty
End Property

Public Function Pop() As Variant
    '
    ' Removes and returns the top element from the stack.
    '
    If IsObject(Container.Back) Then
        Set Pop = Container.PopBack
    Else
        Pop = Container.PopBack
    End If
End Function

Public Sub Push( _
    aValue As Variant _
)
    '
    ' Adds an element with aValue to the top of the stack.
    '
    Container.PushBack aValue
End Sub

Public Property Get Size() As Long
    '
    ' Returns the stack size in elements.
    '
    Size = Container.Size
End Property

Public Property Get Top() As Variant
    '
    ' Returns the element at the top of the stack.
    '
    If IsObject(Container.Back) Then
        Set Top = Container.Back
    Else
        Top = Container.Back
    End If
End Property

Public Sub Swap( _
    ByRef aOther As StackT _
)
    '
    ' Exchanges the stack contents with that of another.
    '
    Container.Swap aOther.MyContainer
End Sub
