VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueueT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' QueueT                                                    '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' * Under Construction *                                    '
' This module defines a FIFO abstract data type.            '
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

Private Container As New CircularContainerT    ' The queue storage container.

'''''''''''''''''''''''''''''
' Private Member Accessors  '
'                           '
' Workaround for VBA's      '
' inability to access       '
' private class members     '
' in other class instances. '
'''''''''''''''''''''''''''''

Friend Property Get MyContainer() As CircularContainerT
    Set MyContainer = Container
End Property

Friend Property Set MyContainer(aContainer As CircularContainerT)
    Set Container = aContainer
End Property

''''''''''''''''''''
' Public Interface '
''''''''''''''''''''

Public Property Get Back() As Variant
    '
    ' Returns the last element.
    '
    If IsObject(Container.Back) Then
        Set Back = Container.Back
    Else
        Back = Container.Back
    End If
End Property

Public Function Clone() As QueueT
    '
    ' Returns a deep copy of the queue.
    '
    Dim q As New QueueT

    Set q.MyContainer = Container.Clone()
    Set Clone = q
    Set q = Nothing
End Function

Public Property Get Front() As Variant
    '
    ' Returns the first element.
    '
    If IsObject(Container.Front) Then
        Set Front = Container.Front
    Else
        Front = Container.Front
    End If
End Property

Public Property Get IsEmpty() As Boolean
    '
    ' Returns TRUE if the queue is empty, else returns FALSE.
    '
    IsEmpty = Container.IsEmpty
End Property

Public Sub Pop()
    '
    ' Removes the first element from the queue.
    '
    Container.PopFront
End Sub

Public Sub Push( _
    aValue As Variant _
)
    '
    ' Appends an element with the given aValue after the
    ' last element.
    '
    Container.PushBack aValue
End Sub

Public Property Get Size() As Long
    '
    ' Returns the queue size in elements.
    '
    Size = Container.Size
End Property

Public Sub Swap( _
    ByRef aOther As QueueT _
)
    '
    ' Exchanges the queue contents with that of another.
    '
    Container.Swap aOther.MyContainer
End Sub



