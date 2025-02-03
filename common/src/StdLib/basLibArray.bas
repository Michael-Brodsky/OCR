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
    ' aCount. An error occurs if aPosition past the end.
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

Public Function ArrayMax( _
    aArray As Variant _
) As Variant
    '
    ' Returns the largest element in a one-dimensional array
    ' using the greater than > operator. No type checking is
    ' performed. The behavior is undefined if aArray is not a
    ' one-dimensional array or the array elements are not
    ' primitive types.
    '
    If ArraySize(aArray) > 0 Then
        Dim a As Variant
        
        ArrayMax = aArray(LBound(aArray))
        For Each a In aArray
            If a > ArrayMax Then ArrayMax = a
        Next
    ElseIf Not IsArray(aArray) Then
         ArrayMax = aArray
    End If
End Function

Public Function ArrayMean( _
    aArray As Variant _
) As Variant
    '
    ' Returns the arithmatic mean of elements in a one-
    ' dimensional array. No type checking is performed.
    ' The behavior is undefined if aArray is not a one-
    ' dimensional array or the array elements are non-
    ' numeric.
    '
    ArrayMean = ArraySum(aArray) / ArraySize(aArray)
End Function

Public Function ArrayMin( _
    aArray As Variant _
) As Variant
    '
    ' Returns the smallest element in a one-dimensional array
    ' using the less than < operator. No type checking is
    ' performed. The behavior is undefined aArray is not a
    ' one-dimensional array or the array elements are not
    ' primitive types.
    '
    If ArraySize(aArray) > 0 Then
        Dim a As Variant
        
        ArrayMin = aArray(LBound(aArray))
        For Each a In aArray
            If a < ArrayMin Then ArrayMin = a
        Next
    ElseIf Not IsArray(aArray) Then
         ArrayMin = aArray
    End If
End Function

Public Function ArrayMinMax( _
    aArray As Variant _
) As Pair
    '
    ' Returns the smallest and largest elements in a one-
    ' dimensional array as a Pair, where .First = smallest
    ' and .Second = largest. No type checking is performed.
    ' The behavior is undefined if aArray is not a one-
    ' dimensional array or the array elements are not
    ' primitive types.
    '
    If ArraySize(aArray) > 0 Then
        Dim a As Variant
        
        ArrayMinMax.First = aArray(LBound(aArray))
        ArrayMinMax.Second = ArrayMinMax.First
        For Each a In aArray
            If a < ArrayMinMax.First Then
                ArrayMinMax.First = a
            ElseIf a > ArrayMinMax.Second Then
                ArrayMinMax.Second = a
            End If
        Next
    ElseIf Not IsArray(aArray) Then
         ArrayMinMax.First = aArray
         ArrayMinMax.Second = ArrayMinMax.First
    End If
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
    ' Appends an element after the last element of an
    ' array. No type checking is performed. The behavior
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
    ' Inserts an element before the first array element.
    ' No type checking is performed. The behavior is
    ' undefined if aArray is not a dynamic array.
    '
    ArrayPushBack aArray, aValue
    ArrayRotateRight aArray
End Sub

Public Function ArrayRange( _
    aArray As Variant _
) As Variant
    '
    ' Returns the arithmatic range of elements in a one-
    ' dimensional array. No type checking is performed.
    ' An error occurs if the elements are non-numeric types.
    ' The behavior is undefined if aArray is not a one-
    ' dimensional array or the array elements are not
    ' primitive types.
    '
    Dim range As Pair
    
    range = ArrayMinMax(aArray)
    ArrayRange = range.Second - range.First
End Function

Public Function ArrayRemove( _
    ByRef aArray As Variant, _
    ByVal aPosition As Long, _
    Optional ByVal aCount As Long = 1 _
) As Long
    '
    ' Removes aCount array elements begining at aPosition.
    ' Decreases the array size by aCount. No type or bounds
    ' checking is performed.
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
    ' Rotates the contents of an array aCount elements to the left,
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
    ' Rotates the contents of an array aCount elements to the right,
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
    ' the array elements and aValue typesare not a homogeneous.
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
    ' Shifts the contents of an array aCount elements to the left.
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
    ' Shifts the contents of an array aCount elements to the right.
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
    ' Returns the size of an array in elements.
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
    ' is undefined if aArray is not a one- dimensional array
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

Public Function ArraySum( _
    aArray As Variant _
) As Variant
    '
    ' Returns the arithmatic sum of elements in a one-
    ' dimensional array. No type checking is performed.
    ' The behavior is undefined if aArray is not a one-
    ' dimensional array or the array elements are non-
    ' numeric.
    '
    Dim a As Variant
    
    For Each a In aArray
        ArraySum = ArraySum + a
    Next
End Function

Public Function ArrayToCsv( _
    aArray As Variant, _
    Optional ByVal aDelim As String = "," _
) As String
    '
    ' Returns the elements of aArray as a delimited string.
    ' aArray must be convertible to an array of type String.
    '
    ArrayToCsv = Join(aArray, aDelim)
End Function

Public Function MaxOf( _
    ParamArray aArgs() As Variant _
) As Variant
    '
    ' Returns the largest value from the given arguments.
    '
    If Not IsMissing(aArgs) Then MaxOf = ArrayMax(CVar(aArgs))
End Function

Public Function MinMaxOf( _
    ParamArray aArgs() As Variant _
) As Pair
    '
    ' Returns the smallest and largest values from the given argumments
    ' as a Pair, where .First = smallest, .Second = largest.
    '
    If Not IsMissing(aArgs) Then MinMaxOf = ArrayMinMax(CVar(aArgs))
End Function

Public Function MinOf( _
    ParamArray aArgs() As Variant _
) As Variant
    '
    ' Returns the smallest value from the given argumments.
    '
    If Not IsMissing(aArgs) Then MinOf = ArrayMin(CVar(aArgs))
End Function

Public Function ParamArrayParam( _
    aArray As Variant _
) As Variant
    '
    ' Returns the first element of aArray or Empty if
    ' the array size = 0. This function can reverse
    ' the effects of ParamArrayDelegate(), and call
    ' functions with any number of discrete parameters,
    ' by calling this function once for each expected
    ' parameter, including optional parameters. For
    ' example:
    '
    '   Function foo(arg0,arg1,arg2)        ' A target function
    '       ...
    '   End Function
    '   Function bar(arg0)                  ' Another target function
    '       ...
    '   End Function
    '
    '   Function foobar(target, ParamArray args())  ' A function that takes a ParamArray and delegates calls
    '       dim a()
    '       a=ParamArrayDelegate(args)
    '       Select Case target
    '           Case "foo":
    '               foobar = foo(ParamArrayParam(a),ParamArrayParam(a),ParamArrayParam(a))    ' <-- The magick.
    '           Case "bar":
    '               foobar = bar(ParamArrayParam(a))  ' <--
    '       End Select
    '   End Function
    '
    On Error Resume Next
    ParamArrayParam = ArrayPopFront(aArray)
End Function

Public Function ParamToCsv( _
    ParamArray aArgs() As Variant _
) As String
    '
    ' Returns the given arguments as a comma-separated string.
    '
    If Not IsMissing(aArgs) Then ParamToCsv = Join(aArgs, ",")
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
    ' Searches each elelment of a one-dimensional array
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


