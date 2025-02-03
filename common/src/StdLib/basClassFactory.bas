Attribute VB_Name = "basClassFactory"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' basClassFactory                                           '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines constructors for library clases.      '
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

Public Function NewAllocatorT( _
    Optional ByVal aSize As Long = 0 _
) As AllocatorT
    Dim alloc As New AllocatorT
    
    alloc.Allocate aSize
    Set NewAllocatorT = alloc
    Set alloc = Nothing
End Function

Public Function NewArrayT( _
    ByVal aSize As Long _
) As ArrayT
    Dim obj As New ArrayT, v() As Variant
    
    If aSize <> 0 Then ReDim v(aSize - 1)
    obj.Init v
    Set NewArrayT = obj
    Set obj = Nothing
End Function

Public Function NewCallbackT( _
    Optional ByVal aProcedure As String, _
    Optional ByVal aMethod As VbCallType, _
    Optional aTarget As Object = Nothing _
) As CallbackT
    Dim obj As New CallbackT
    
    obj.Init aProcedure, aMethod, aTarget
    Set NewCallbackT = obj
    Set obj = Nothing
End Function

Public Function NewComplexT( _
    Optional ByVal aReal As Double, _
    Optional ByVal aImag As Double _
) As ComplexT
    Dim obj As New ComplexT
    
    obj.InitList aReal, aImag
    Set NewComplexT = obj
    Set obj = Nothing
End Function

Public Function NewContainerT( _
    Optional aSize As Long _
) As ContainerT
    Dim obj As New ContainerT
    
    obj.Resize aSize
    Set NewContainerT = obj
    Set obj = Nothing
End Function

Public Function NewDAOConnectionT( _
    Optional aConnect As Variant _
) As DAOConnectionT
    Dim obj As New DAOConnectionT
    
    If Not IsMissing(aConnect) Then obj.Connect = aConnect
    Set NewDAOConnectionT = obj
    Set obj = Nothing
End Function

Public Function NewDAORecordsetT() As DAORecordsetT
    Set NewDAORecordsetT = New DAORecordsetT
End Function

Public Function NewFso() As Object
    Set NewFso = CreateObject("Scripting.FileSystemObject")
End Function

Public Function NewListBoxExT( _
    Optional aControl As ListBox _
) As ListBoxExT
    Dim obj As New ListBoxExT
    
    Set obj.Control = aControl
    Set NewListBoxExT = obj
    Set obj = Nothing
End Function

Public Function NewPairT( _
    aFirst As Variant, _
    aSecond As Variant _
) As PairT
    Dim obj As New PairT
    
    obj.InitList aFirst, aSecond
    Set NewPairT = obj
    Set obj = Nothing
End Function

Public Function NewShell() As Object
    Set NewShell = CreateObject("WScript.Shell")
End Function

Public Function NewStackT( _
    Optional aData As Variant _
) As StackT
    Dim obj As New StackT
    
    Set NewStackT = obj
    Set obj = Nothing
End Function

Public Function NewVectorT( _
    Optional ByVal aSize As Long = 0 _
) As VectorT
    Dim obj As New VectorT
    
    If aSize <> 0 Then obj.Resize aSize
    Set NewVectorT = obj
    Set obj = Nothing
End Function



