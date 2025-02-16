Attribute VB_Name = "basLibArray"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' LibArray                                                  '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a library of functions that operate   '
' on arrays.                                                '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' LibVBA                                                    '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the library has only been tested with     '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

Public Function ArrayBack( _
    ByRef aArray As Variant _
) As Variant
    '
    ' Returns the the last array element. No type checking
    ' is performed. The behavior is undefined if aArray is
    ' not an array.
    '
    If IsObject(aArray(UBound(aArray))) Then
        Set ArrayBack = aArray(UBound(aArray))
    Else
        ArrayBack = aArray(UBound(aArray))
    End If
End Function

Public Function ArrayDistance( _
    ByVal aFirst As Long, _
    ByVal aLast As Long _
) As Long
    
    '
    ' Returns the number of hops from aFirst to aLast.
    ' This function is used by container classes for
    ' consistency in implementing iterator-like functionality.
    '
    ArrayDistance = aLast - aFirst
End Function

Public Function ArrayEmpty( _
    aArray As Variant _
) As Boolean
    '
    ' Returns TRUE if an array has zero elements (is unallocated),
    ' else returns FALSE. No type checking is performed. The
    ' behavior is undefined if aArray is not an array.
    '
    ArrayEmpty = (ArraySize(aArray) = 0)
End Function

Public Sub ArrayFill( _
    ByRef aArray As Variant, _
    ByVal aPosition As Long, _
    aValue As Variant, _
    Optional ByVal aCount As Long = 1 _
)
    '
    ' Assigns aValue to aCount array elements begining at aPosition.
    ' No type or bounds checking is performed. The behavior is
    ' undefined if aArray is not an array.
    '
    Dim it As Long
    
    For it = aPosition To aPosition + aCount - 1
        If Not IsMissing(aValue) Then
            If IsObject(aValue) Then
                Set aArray(it) = aValue
            Else
                aArray(it) = aValue
            End If
        End If
    Next
End Sub

Public Function ArrayFirst( _
    ByRef aArray As Variant _
) As Variant
    '
    ' Returns the lower bound of aArray. The behavior is
    ' undefined if aArray is not an array.
    '
    ArrayFirst = LBound(aArray)
End Function

Public Function ArrayFront( _
    ByRef aArray As Variant _
) As Variant
    '
    ' Returns the the first array element. No type checking
    ' is performed. The behavior is undefined if aArray is
    ' not an array.
    '
    If IsObject(aArray(LBound(aArray))) Then
        Set ArrayFront = aArray(LBound(aArray))
    Else
        ArrayFront = aArray(LBound(aArray))
    End If
End Function

Public Function ArrayInsert( _
    ByRef aArray As Variant, _
    ByVal aPosition As Long, _
    aValue As Variant, _
    Optional ByVal aCount As Long = 1 _
) As Long
    '
    ' Inserts aCount array elements of aValue begining at aPosition.
    ' No type checking is performed. Increases the array size by
    ' aCount. An error occurs if aPosition is past the end.
    '
    Dim i As Long, last As Long

    last = ArraySize(aArray)
    ArrayResize aArray, last + aCount
    For i = 0 To aCount - 1
        Dim it As Long
        
        it = i + aPosition
        If IsObject(aArray(it)) Then
            Set aArray(last) = aArray(it)
        Else
            aArray(last) = aArray(it)
        End If
        last = last + 1
    Next
    ArrayFill aArray, aPosition, aValue, aCount
End Function

Public Function ArrayLast( _
    ByRef aArray As Variant _
) As Variant
    '
    ' Returns the upper bound of aArray. The behavior is
    ' undefined if aArray is not an array.
    '
    ArrayLast = UBound(aArray)
End Function

Public Function ArrayPopBack( _
    ByRef aArray As Variant _
) As Variant
    '
    ' Removes and returns the last element from an array.
    ' No type checking is performed. The behavior is
    ' undefined if aArray is not a dynamic array.
    '
    If IsObject(ArrayBack(aArray)) Then
        Set ArrayPopBack = ArrayBack(aArray)
    Else
        ArrayPopBack = ArrayBack(aArray)
    End If
    ArrayResize aArray, ArraySize(aArray) - 1
End Function

Public Function ArrayPopFront( _
    ByRef aArray As Variant _
) As Variant
    '
    ' Removes and returns the first array element.
    ' No type checking is performed. The behavior is
    ' undefined if aArray is not a dynamic array.
    '
    If IsObject(ArrayFront(aArray)) Then
        Set ArrayPopFront = ArrayFront(aArray)
    Else
        ArrayPopFront = ArrayFront(aArray)
    End If
    ArrayRotateLeft aArray
    ArrayResize aArray, ArraySize(aArray) - 1
End Function

Public Sub ArrayPushBack( _
    ByRef aArray As Variant, _
    aValue As Variant _
)
    '
    ' Appends an element with aValue after the last element
    ' of an array. No type checking is performed. The behavior
    ' is undefined if aArray is not a dynamic array.
    '
    Dim sz As Integer
    
    sz = ArraySize(aArray)
    ArrayResize aArray, sz + 1
    If IsObject(aValue) Then
        Set aArray(sz) = aValue
    Else
        aArray(sz) = aValue
    End If
End Sub

Public Sub ArrayPushFront( _
    ByRef aArray As Variant, _
    aValue As Variant _
)
    '
    ' Inserts an element with aValue before the first
    ' array element. No type checking is performed. The
    ' behavior is undefined if aArray is not a dynamic
    ' array.
    '
    ArrayPushBack aArray, aValue
    ArrayRotateRight aArray
End Sub

Public Function ArrayRemove( _
    ByRef aArray As Variant, _
    ByVal aPosition As Long, _
    Optional ByVal aCount As Long = 1 _
) As Long
    '
    ' Removes aCount array elements begining at aPosition.
    ' Decreases the array size by aCount. No type or bounds
    ' checking is performed. The behavior is undefined if
    ' aArray is not a dynamic array.
    '
    If aCount > 0 Then
        Dim i As Long, First As Long, last As Long, sz As Long
        
        sz = ArraySize(aArray)
        First = aPosition + aCount
        last = sz - 1
        For i = First To last - 1
            If IsObject(aArray(i + aCount)) Then
                Set aArray(i) = aArray(i + aCount)
            Else
                aArray(i) = aArray(i + aCount)
            End If
        Next
        sz = sz - aCount
        ArrayFill aArray, sz, Empty, ArrayDistance(sz, last)
        ArrayResize aArray, ArraySize(aArray) - aCount
    End If
    ArrayRemove = aPosition
End Function

Public Function ArrayResize( _
    ByRef aArray As Variant, _
    ByVal aSize As Long _
) As Long
    '
    ' Resizes a dynamic array to the specified size in
    ' elements. Arrays of size zero are deallocated.
    ' No type-checking is performed. The behavior is
    ' undefined if aArray is not a dynamic array.
    '
    aSize = aSize - 1
    If aSize < 0 Then
        If ArraySize(aArray) > 0 Then Erase aArray
    ElseIf ArraySize(aArray) > 0 Then
        ReDim Preserve aArray(aSize)
    Else
        ReDim aArray(aSize)
    End If
    ArrayResize = ArraySize(aArray)
End Function

Public Function ArrayRotateLeft( _
    ByRef aArray As Variant, _
    Optional ByVal aCount As Integer = 1 _
) As Variant
    '
    ' Rotates the contents of aArray aCount elements to the left,
    ' wrapping the left-most element to the right-most element.
    ' If aCount is negative then calls ArrayRotateRight() with the
    ' absolute value of aCount. No type-checking is performed. The
    ' behavior is undefined if aArray is not an array or the array
    ' elements types are not homogeneous.
    '
    Dim i As Integer, j As Integer
    
    If aCount < 0 Then
        aArray = ArrayRotateRight(aArray, Abs(aCount))
    Else
        For i = 1 To aCount
            For j = LBound(aArray) To UBound(aArray) - 1
                Swap aArray(j), aArray(j + 1)
            Next j
        Next i
    End If
    ArrayRotateLeft = aArray
End Function

Public Function ArrayRotateRight( _
    ByRef aArray As Variant, _
    Optional ByVal aCount As Integer = 1 _
) As Variant
    '
    ' Rotates the contents of aArray aCount elements to the right,
    ' wrapping the right-most element to the left-most element.
    ' If aCount is negative then calls ArrayRotateLeft() with the
    ' absolute value of aCount. No type-checking is performed. The
    ' behavior is undefined if aArray is not an array or the array
    ' elements types are not homogeneous.
    '
    Dim i As Integer, j As Integer
    
    If aCount < 0 Then
        aArray = ArrayRotateLeft(aArray, Abs(aCount))
    Else
        For i = 1 To aCount
            For j = UBound(aArray) To LBound(aArray) + 1 Step -1
                Swap aArray(j), aArray(j - 1)
            Next j
        Next i
    End If
    ArrayRotateRight = aArray
End Function

Public Function ArraySearch( _
    aValue As Variant, _
    aArray As Variant, _
    Optional ByVal aSorted As Boolean = False _
) As Variant
    '
    ' Searches for and returns the array element whose value
    ' matches aValue, if any. Uses the most suitable algorithm
    ' based on aSorted. No type-checking is performed. The
    ' behavior is undefined if the aArray is not an array, or
    ' the array elements and aValue types are not a homogeneous.
    '
    If aSorted Then
        ArraySearch = SearchBinary(aValue, aArray)
    Else
        ArraySearch = SearchSequential(aValue, aArray)
    End If
End Function

Public Function ArrayShiftLeft( _
    ByRef aArray As Variant, _
    Optional ByVal aCount As Integer = 1 _
) As Variant
    '
    ' Shifts the contents of aArray aCount elements to the left.
    ' The right-most element is set to its default value. If aCount
    ' is negative then calls ArrayShiftRight() with the absolute
    ' value of aCount. No type checking is performed. The behavior
    ' is undefined if aArray is not an array.
    '
    Dim i As Integer, j As Integer
    
    If aCount < 0 Then
        If IsObject(aArray(j)) Then
            Set aArray = ArrayShiftRight(aArray, Abs(aCount))
        Else
            aArray = ArrayShiftRight(aArray, Abs(aCount))
        End If
    Else
        For i = 1 To aCount
            For j = LBound(aArray) To UBound(aArray) - 1
                If IsObject(aArray(j)) Then
                    Set aArray(j) = aArray(j + 1)
                Else
                    aArray(j) = aArray(j + 1)
                End If
            Next j
            If IsObject(aArray(UBound(aArray))) Then
                Set aArray(UBound(aArray)) = Nothing
            Else
                aArray(UBound(aArray)) = Empty
            End If
        Next i
    End If
    ArrayShiftLeft = aArray
End Function

Public Function ArrayShiftRight( _
    ByRef aArray As Variant, _
    Optional ByVal aCount As Integer = 1 _
) As Variant
    '
    ' Shifts the contents of aArray aCount elements to the right.
    ' The left-most element is set to its default value. If aCount
    ' is negative then calls ArrayShiftLeft() with the absolute
    ' value of aCount. No type checking is performed. The behavior
    ' is undefined if aArray is not an array.
    '
    Dim i As Integer, j As Integer
    
    If aCount < 0 Then
        If IsObject(aArray(j)) Then
            Set aArray = ArrayShiftLeft(aArray, Abs(aCount))
        Else
            aArray = ArrayShiftLeft(aArray, Abs(aCount))
        End If
    Else
        For i = 1 To aCount
            For j = UBound(aArray) To LBound(aArray) + 1 Step -1
                If IsObject(aArray(j)) Then
                    Set aArray(j) = aArray(j - 1)
                Else
                    aArray(j) = aArray(j - 1)
                End If
            Next j
            If IsObject(aArray(LBound(aArray))) Then
                Set aArray(LBound(aArray)) = Nothing
            Else
                aArray(LBound(aArray)) = Empty
            End If
        Next i
    End If
    ArrayShiftRight = aArray
End Function

Public Function ArraySize( _
    aArray As Variant _
) As Long
    '
    ' Returns the size of aArray in elements.
    ' Returns 0 if aArray is an unallocated
    ' array or not an array.
    '
    On Error Resume Next
    ArraySize = UBound(aArray) - LBound(aArray) + 1
    Err.Clear
End Function

Public Sub ArraySort( _
    ByRef aArray As Variant _
)
    '
    ' Lexicographically sorts a one-dimensional array using
    ' the insertion sort algorithm. Array elements must be
    ' contiguous. No type checking is performed. The behavior
    ' is undefined if aArray is not a one-dimensional array
    ' or the array elements are not primitive types.
    '
    Dim i As Integer, n As Integer
    
    n = ArraySize(aArray) - 1
    For i = 1 To n
        Dim j As Integer, low As Integer
        Dim key As Variant
        
        low = LBound(aArray)
        key = aArray(low + i)
        j = low + i - 1
        Do While j >= low
            If aArray(j) <= key Then Exit Do
            aArray(j + 1) = aArray(j)
            j = j - 1
        Loop
        aArray(j + 1) = key
    Next
End Sub

Public Function ArraySorted( _
    aArray As Variant _
) As Boolean
    '
    ' Returns TRUE if a one-dimensional array is lexicographically sorted,
    ' unallocated or not an array, else returns FALSE. No type checking is
    ' performed. The behavior is undefined if aArray is not a one-dimensional
    ' array or the array elements are not primitive types.
    '
    Dim i As Integer, j As Integer
    Dim ai As Variant
    
    j = ArraySize(aArray) - 1
    If j < 1 Then
        ArraySorted = True
    Else
        ai = aArray(LBound(aArray))
        i = LBound(aArray) + 1
        
        Do While i <= j
            If ai <= aArray(i) Then
                ai = aArray(i)
                i = i + 1
            Else
                Exit Do
            End If
        Loop
        ArraySorted = (i > j)
    End If
End Function

Public Function ArrayToCsv( _
    aArray As Variant, _
    Optional ByVal aDelim As String = "," _
) As String
    '
    ' Returns the elements of aArray as a delimited string.
    ' aArray must be convertible to an array of type String.
    ' Uninitialized arrays or objects that are not arrays
    ' return zero-length strings. Uninitialized or Null
    ' elements are assigned as zero-length strings in the
    ' returned string.
    '
    If ArraySize(aArray) > 0 Then
        Dim i As Long, item
        
        For i = LBound(aArray) To UBound(aArray)
            If Not IsSomething(aArray(i)) Then aArray(i) = Empty
        Next
        ArrayToCsv = Join(aArray, aDelim)
    End If
End Function

Public Function ParamToCsv( _
    ParamArray aArgs() As Variant _
) As String
    '
    ' Returns the given arguments as a comma-separated string.
    '
    If Not IsMissing(aArgs) Then ParamToCsv = ArrayToCsv(CVar(aArgs), ",")
End Function

Private Function SearchBinary( _
    aValue As Variant, _
    aArray As Variant _
) As Variant
    '
    ' Searches a sorted one-dimensional array for aValue and
    ' returns the value if found. Uses the binary search algorithm.
    '
    Dim low As Integer, high As Integer, mid As Integer
    
    low = LBound(aArray)
    high = UBound(aArray)
    Do While low <= high
        mid = low + (high - low) / 2
        If aArray(mid) = aValue Then
            SearchBinary = aArray(mid)
            Exit Do
        ElseIf aArray(mid) < aValue Then
            low = mid + 1
        Else
            high = mid - 1
        End If
    Loop
End Function

Private Function SearchSequential( _
    aFind As Variant, _
    aArray As Variant _
) As Variant
    '
    ' Searches each element of a one-dimensional array
    ' for aArray and returns the value if found.
    '
    Dim elem As Variant
    
    For Each elem In aArray
        If elem = aFind Then
            SearchSequential = elem
            Exit For
        End If
    Next
End Function


