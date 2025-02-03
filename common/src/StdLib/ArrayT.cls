VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' ArrayT                                                    '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a container class that encapsulates   '
' fixed-size arrays.                                        '
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

'''''''''''''''''''''''
' Types and Constants '
'''''''''''''''''''''''

Private Const kStrArrayAlreadyDimensioned As String = "Array already dimensioned"

'''''''''''''''''''
' Private Members '
'''''''''''''''''''

Private Container As New ContainerT ' The array storage container.
Private initialized As Boolean      ' Flag indicating whether the array size has been initialized.
                                    ' Since the underlying container is a dynamic array, we use
                                    ' this to limit operations that resize the array to one-time use.

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
    ' Returns the element at the given aIndex.
    '
    MakeUnsigned aIndex
    If IsObject(Container.At(aIndex)) Then
        Set At = Container.At(aIndex)
    Else
        At = Container.At(aIndex)
    End If
End Property

Public Property Let At( _
    ByVal aIndex As Long, _
    ByVal aValue As Variant _
)
    '
    ' Assigns aValue to the element at the given aIndex.
    '
    MakeUnsigned aIndex
    Container.At(aIndex) = aValue
End Property

Public Property Set At( _
    ByVal aIndex As Long, _
    ByVal aValue As Variant _
)
    '
    ' Assigns aValue to the element at the given aIndex.
    '
    MakeUnsigned aIndex
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
    ByVal aValue As Variant _
)
    '
    ' Assigns the given aValue to the last element.
    '
    Container.Back = aValue
End Property

Public Property Set Back( _
    ByVal aValue As Variant _
)
    '
    ' Assigns the given aValue to the last element.
    '
    Set Container.Back = aValue
End Property

Public Function Clone() As ArrayT
    '
    ' Returns a deep copy of the array.
    '
    Dim a As New ArrayT
    
    Set a.MyContainer = Container.Clone()
    Set Clone = a
    Set a = Nothing
End Function

Public Property Get Data() As Variant
    '
    ' Returns the underlying storage container.
    '
    If IsObject(Container.Data) Then
        Set Data = Container.Data
    Else
        Data = Container.Data
    End If
End Property

Public Sub Fill( _
    aValue As Variant _
)
    '
    ' Assigns the given aValue to all elements.
    '
    Container.Fill 0, aValue, Size
End Sub

Public Property Get Front() As Variant
    '
    ' Returns the first element.
    '
    If IsObject(ArrayFront(Container.Data)) Then
        Set Front = Container.Front
    Else
        Front = Container.Front
    End If
End Property

Public Property Let Front( _
    ByVal aValue As Variant _
)
    '
    ' Assigns the given aValue to the first element.
    '
    Container.Front = aValue
End Property

Public Property Set Front( _
    ByVal aValue As Variant _
)
    '
    ' Assigns the given aValue to the first element.
    '
    Set Container.Front = aValue
End Property

Public Sub InitList( _
    aValues As Variant _
)
    '
    ' Initializes the array with the given values.
    '
    Dim i As Long
    
    Size = ArraySize(aValues)
    For i = 0 To ArraySize(aValues) - 1
        Container.At(i) = aValues(i)
    Next
End Sub

Public Property Get IsEmpty() As Boolean
    '
    ' Returns TRUE if the array is empty, else returns FALSE.
    '
    IsEmpty = Container.IsEmpty
End Property

Public Property Get MaxSize() As Long
    '
    ' Returns the maximum number of elements the array
    ' can hold which, in the case of a fixed-size array,
    ' is equal to Size.
    '
    MaxSize = Size
End Property

Public Property Get Size() As Long
    '
    ' Returns the array size in elements.
    '
    Size = Container.Capacity
End Property

Public Property Let Size( _
    ByVal aSize As Long _
)
    '
    ' Allocates the specified number of elements and
    ' marks the array as initialized.
    '
    If Not initialized Then
        Container.Resize aSize
    Else
        Err.Raise kvbErrInvalidProcedureCall, CurrentProject.name, kStrArrayAlreadyDimensioned
    End If
    initialized = True
End Property


Public Sub Swap( _
    ByRef aOther As ArrayT _
)
    '
    ' Exchanges the array contents with that of another.
    '
    Container.Swap aOther.MyContainer
End Sub
