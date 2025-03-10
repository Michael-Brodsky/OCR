VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VectorT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' VectorT                                                   '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a sequence container class that       '
' encapsulates dynamic-sized arrays.                        '
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

Private Container As New ContainerT    ' The underlying storage container.

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

Public Property Get At( _
    ByVal aIndex As Long _
) As Variant
    '
    ' Returns the element at aIndex.
    '
    If IsObject(Container.At(aIndex)) Then
        Set At = Container.At(aIndex)
    Else
        At = Container.At(aIndex)
    End If
End Property

Public Property Let At( _
    ByVal aIndex As Long, _
    aValue As Variant _
)
    '
    ' Assigns aValue to the element at aIndex.
    '
    Container.At(aIndex) = aValue
End Property

Public Property Set At( _
    ByVal aIndex As Long, _
    aValue As Variant _
)
    '
    ' Assigns aValue to the element aIndex.
    '
    Set Container.At(aIndex) = aValue
End Property

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

Public Property Let Back( _
    aValue As Variant _
)
    '
    ' Assigns aValue to the last element.
    '
    Container.Back = aValue
End Property

Public Property Set Back( _
    aValue As Variant _
)
    '
    ' Assigns aValue to the last element.
    '
    Set Container.Back = aValue
End Property

Public Property Get Capacity() As Long
    '
    ' Returns the number of elements that the vector can currently hold.
    '
    Capacity = Container.Capacity
End Property

Public Sub Clear()
    '
    ' Removes all vector elements.
    '
    Container.Clear
End Sub

Public Function Clone() As VectorT
    '
    ' Returns a deep copy of the vector.
    '
    Dim v As New VectorT

    Set v.MyContainer = Container.Clone()
    Set Clone = v
    Set v = Nothing
End Function

Public Property Get Data() As Variant
    '
    ' Returns the underlying data.
    '
    If IsObject(Container.Data) Then
        Set Data = Container.Data
    Else
        Data = Container.Data
    End If
End Property

Public Property Let Data( _
    aData As Variant _
)
    '
    ' Assigns aData to the underlying data.
    '
    Container.Data = aData
End Property

Public Property Set Data( _
    aData As Variant _
)
    '
    ' Assigns aData to the underlying data.
    '
    Set Container.Data = aData
End Property

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

Public Property Let Front( _
    aValue As Variant _
)
    '
    ' Assigns aValue to the first element.
    '
    Container.Front = aValue
End Property

Public Property Set Front( _
    aValue As Variant _
)
    '
    ' Assigns aValue to the first element.
    '
    Set Container.Front = aValue
End Property

Public Sub InitList( _
    ParamArray aValues() As Variant _
)
    '
    ' Assigns the list of aValues to the vector.
    '
    Dim item As Variant, i As Long
    
    Clear
    Reserve ArraySize(CVar(aValues))
    For Each item In aValues
        PushBack item
    Next
End Sub

Public Function Insert( _
    ByVal aPosition As Long, _
    aValue As Variant, _
    Optional ByVal aCount As Long = 1 _
) As Long
    '
    ' Inserts aCount elements with aValue begining
    ' at aPosition.
    '
    Insert = Container.Insert(aPosition, aValue, aCount)
End Function

Public Property Get IsEmpty() As Boolean
    '
    ' Returns TRUE if the vector is empty, else returns FALSE.
    '
    IsEmpty = Container.IsEmpty
End Property

Public Property Get MaxSize() As Long
    '
    ' Returns the theoretical maximum number of elements the
    ' vector can hold.
    '
    MaxSize = Container.MaxSize
End Property

Public Function PopBack() As Variant
    '
    ' Removes and returns the last element.
    '
    If IsObject(Container.Back) Then
        Set PopBack = Container.PopBack
    Else
        PopBack = Container.PopBack
    End If
End Function

Public Sub PushBack( _
    aValue As Variant _
)
    '
    ' Appends an element with aValue after the last element.
    '
    Container.PushBack aValue
End Sub

Public Sub Remove( _
    ByVal aFirst As Long, _
    Optional ByVal aCount As Long = 1 _
)
    '
    ' Removes aCount elements begining at aPosition.
    '
    Container.Remove aFirst, aCount
End Sub

Public Sub Reserve( _
    ByVal aSize As Long _
)
    '
    ' Increase the vector Capacity to at least aSize elements.
    '
    Container.Reserve aSize
End Sub

Public Sub Resize( _
    ByVal aSize As Long, _
    Optional aValue As Variant _
)
    '
    ' Resizes the vector to exactly aSize elements and
    ' assigns aValue to any new elements.
    '
    Container.Resize aSize, aValue
End Sub

Public Sub ShrinkToFit()
    '
    ' Reduces the vector Capacity to its current Size.
    '
    Container.ShrinkToFit
End Sub

Public Property Get Size() As Long
    '
    ' Returns the vector size in elements.
    '
    Size = Container.Size
End Property

Public Sub Swap( _
    ByRef aOther As VectorT _
)
    '
    ' Exchanges the vector contents with that of another.
    '
    Container.Swap aOther.MyContainer
End Sub


