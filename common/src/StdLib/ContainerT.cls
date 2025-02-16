VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContainerT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' ContainerT                                                '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a class used by library container     '
' classes to implement storage and common functionality.    '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' AllocatorT                                                '
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

Private Allocator As New AllocatorT ' Container storage allocator.
Private CSize As Long               ' Current container size in elements.

'''''''''''''''''''''''''''''
' Private Member Accessors  '
'                           '
' Workaround for VBA's      '
' inability to access       '
' private class members     '
' in other class instances. '
'''''''''''''''''''''''''''''

Friend Property Get MyCSize() As Long
    MyCSize = CSize
End Property

Friend Property Let MyCSize(aCSize As Long)
    CSize = aCSize
End Property

''''''''''''''''''''
' Public Interface '
''''''''''''''''''''

Public Property Get At( _
    ByVal aIndex As Long _
) As Variant
    '
    ' Returns the element at aIndex.
    ' No type or bounds checking is performed.
    '
    If IsObject(Allocator.Data(aIndex)) Then
        Set At = Allocator.Data(aIndex)
    Else
        At = Allocator.Data(aIndex)
    End If
End Property

Public Property Let At( _
    ByVal aIndex As Long, _
    aValue As Variant _
)
    '
    ' Assigns aValue to the element at aIndex.
    ' No type or bounds checking is performed.
    '
    Allocator.Data(aIndex) = aValue
End Property

Public Property Set At( _
    ByVal aIndex As Long, _
    aValue As Variant _
)
    '
    ' Assigns aValue to the element at aIndex.
    ' No type or bounds checking is performed.
    '
    Set Allocator.Data(aIndex) = aValue
End Property

Public Property Get Back() As Variant
    '
    ' Returns the last element.
    ' No type checking is performed.
    ' An error occurs if the container is empty.
    '
    If IsObject(Allocator.Data(Size - 1)) Then
        Set Back = Allocator.Data(Size - 1)
    Else
        Back = Allocator.Data(Size - 1)
    End If
End Property

Public Property Let Back( _
    aValue As Variant _
)
    '
    ' Assigns aValue to the last element.
    ' No type checking is performed.
    ' An error occurs if the container is empty.
    '
    Allocator.Data(Size - 1) = aValue
End Property

Public Property Set Back( _
    aValue As Variant _
)
    '
    ' Assigns aValue to the last element.
    ' No type checking is performed.
    ' An error occurs if the container is empty.
    '
    Set Allocator.Data(Size - 1) = aValue
End Property

Public Property Get Capacity() As Long
    '
    ' Returns the number of elements that can be held
    ' in currently allocated storage.
    '
    Capacity = Allocator.Capacity
End Property

Public Sub Clear()
    '
    ' Removes all elements from the container, sets Size
    ' to zero.
    '
    While Size > 0
        PopBack
    Wend
End Sub

Public Function Clone() As ContainerT
    '
    ' Returns a deep copy of the container.
    '
    Dim c As New ContainerT
    
    c.Data = Data
    c.MyCSize = CSize
    Set Clone = c
    Set c = Nothing
End Function

Public Property Get Data() As Variant
    '
    ' Returns the underlying storage container.
    '
    If IsObject(Allocator.Data) Then
        Set Data = Allocator.Data
    Else
        Data = Allocator.Data
    End If
End Property

Public Property Let Data( _
    aData As Variant _
)
    '
    ' Assigns aData to the underlying storage container.
    ' Sets Size and Capacity to the number of elements assigned.
    ' No type checking is performed.
    '
    Allocator.Data = aData
    Size = Allocator.Capacity
End Property

Public Property Set Data( _
    aData As Variant _
)
    '
    ' Assigns aData to the underlying storage container.
    ' Sets Size and Capacity to the number of elements assigned.
    ' No type checking is performed.
    '
    Set Allocator.Data = aData
    Size = Allocator.Capacity
End Property

Public Sub Fill( _
    ByVal aPosition As Long, _
    aValue As Variant, _
    Optional ByVal aCount As Long = 1 _
)
    '
    ' Assigns aValue to aCount elements begining at aPosition.
    ' No type or bounds checking is performed.
    '
    Dim it As Long
    
    MakeUnsigned aPosition
    MakeUnsigned aCount
    For it = aPosition To aPosition + aCount - 1
        If Not IsMissing(aValue) Then
            If IsObject(aValue) Then
                Set Allocator.Data(it) = aValue
            Else
                Allocator.Data(it) = aValue
            End If
        End If
    Next
End Sub

Public Property Get Front() As Variant
    '
    ' Returns the first element.
    ' No type checking is performed.
    ' An error occurs if the container is empty.
    '
    If IsObject(Allocator.Data(0)) Then
        Set Front = Allocator.Data(0)
    Else
        Front = Allocator.Data(0)
    End If
End Property

Public Property Let Front( _
    aValue As Variant _
)
    '
    ' Assigns aValue to the first element.
    ' No type checking is performed.
    ' An error occurs if the container is empty.
    '
    Allocator.Data(0) = aValue
End Property

Public Property Set Front( _
    aValue As Variant _
)
    '
    ' Assigns aValue to the first element.
    ' No type checking is performed.
    ' An error occurs if the container is empty.
    '
    Set Allocator.Data(0) = aValue
End Property

Public Function Insert( _
    ByVal aPosition As Long, _
    aValue As Variant, _
    Optional ByVal aCount As Long = 1 _
) As Long
    '
    ' Inserts aCount elements of aValue begining at aPosition.
    ' No type checking is performed. Increases Size by aCount.
    ' If the new Size is greater than the old Capacity, a
    ' reallocation occurs. An error occurs if aPosition is past
    ' the end.
    '
    MakeUnsigned aPosition
    MakeUnsigned aCount
    If aPosition < Size Then
        Dim i As Long
    
        Allocator.Allocate Size + aCount
        For i = 0 To aCount - 1
            Dim it As Long
            
            it = i + aPosition
            If IsObject(Allocator.Data(it)) Then
                Set Allocator.Data(Size) = Allocator.Data(it)
            Else
                Allocator.Data(Size) = Allocator.Data(it)
            End If
            Size = Size + 1
        Next
        Fill aPosition, aValue, aCount
        Insert = aPosition
    Else
        Err.Raise kvbErrInvalidProcedureCall, CurrentProject.Name, "Position out of bounds"
    End If
End Function

Public Property Get IsEmpty() As Boolean
    '
    ' Returns TRUE if the container is empty (has zero elements),
    ' else returns FALSE.
    '
    IsEmpty = (Size = 0)
End Property

Public Property Get MaxSize() As Long
    '
    ' Returns the theoretical maximum number of elements the
    ' container can hold.
    '
    MaxSize = kvbLongMax
End Property

Public Function PopBack() As Variant
    '
    ' Removes and returns the last element.
    ' No type checking is performed. Decreases Size by one.
    ' An error occurs if the container is empty.
    '
    Size = Size - 1
    If IsObject(Allocator.Data(Size)) Then
        Set PopBack = Allocator.Data(Size)
    Else
        PopBack = Allocator.Data(Size)
    End If
    Allocator.Data(Size) = Empty
End Function

Public Sub PushBack( _
    aValue As Variant _
)
    '
    ' Appends an element with aValue after the last element.
    ' No type checking is performed. Increases Size by one.
    ' If the new Size is greater than the old Capacity, a
    ' reallocation occurs.
    '
    Allocator.Allocate Size + 1
    If Not IsMissing(aValue) Then
        If IsObject(aValue) Then
            Set Allocator.Data(Size) = aValue
        Else
            Allocator.Data(Size) = aValue
        End If
    End If
    Size = Size + 1
End Sub

Public Function Remove( _
    ByVal aPosition As Long, _
    Optional ByVal aCount As Long = 1 _
) As Long
    '
    ' Removes aCount elements begining at aPosition.
    ' Decreases Size by aCount. No type checking is performed.
    ' An error occurs if aPosition is past the end.
    '
    MakeUnsigned aPosition
    MakeUnsigned aCount
    If aPosition < Size Then
        If aCount > 0 Then
            Dim i As Long, First As Long, last As Long
            
            First = aPosition + aCount
            last = Size - 1
            For i = First To last - 1
                If IsObject(Allocator.Data(i + aCount)) Then
                    Set Allocator.Data(i) = Allocator.Data(i + aCount)
                Else
                    Allocator.Data(i) = Allocator.Data(i + aCount)
                End If
            Next
            Size = Size - aCount
            Fill Size, Empty, ArrayDistance(Size, last)
        End If
        Remove = aPosition
    Else
        Err.Raise kvbErrInvalidProcedureCall, CurrentProject.Name, "Position out of bounds"
    End If
End Function

Public Sub Reserve( _
    ByVal aSize As Long _
)
    '
    ' Increase the container Capacity to at least aSize elements.
    ' If the new Capacity is greater than the old Capacity, a
    ' reallocation occurs, otherwise nothing occurs. The
    ' container Size remains the same.
    '
    Allocator.Resize aSize
End Sub

Public Sub Resize( _
    ByVal aSize As Long, _
    Optional aValue As Variant _
)
    '
    ' Resizes the container to exactly aSize elements. If the
    ' old Size is greater than aSize, the container is reduced
    ' to the first aSize elements. If the old Size is less than
    ' aSize, new elements with the given aValue are appended to
    ' the container. Otherwise, nothing occurs.
    '
    Dim val As Variant
    
    MakeUnsigned aSize
    Allocator.Resize aSize
    If aSize > Size Then Fill Size, aValue, ArrayDistance(Size, aSize)
    Size = aSize
End Sub

Public Sub ShrinkToFit()
    '
    ' Reduces the container Capacity to its current Size.
    '
    Allocator.Resize Size
End Sub

Public Property Get Size() As Long
    '
    ' Returns the container size in elements.
    '
    Size = CSize
End Property

Public Sub Swap( _
    ByRef aOther As ContainerT _
)
    '
    ' Exchanges the container contents with that of another.
    '
    Dim tmpCSize As Long, tmpData As Variant
    
    tmpCSize = aOther.MyCSize
    tmpData = aOther.Data
    aOther.MyCSize = CSize
    aOther.Data = Data
    CSize = CSize
    Data = tmpData
End Sub

'''''''''''''''''''''
' Private Interface '
'''''''''''''''''''''

Private Property Let Size( _
    ByVal aSize As Long _
)
    '
    ' Assigns aSize to the container Size in elements.
    '
    CSize = aSize
End Property

