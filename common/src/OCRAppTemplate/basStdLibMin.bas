Attribute VB_Name = "basStdLibMin"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' basSetup                                                  '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module contains a minimal copy of StdLib objects to  '
' allow the application to start with broken references.    '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' (None)                                                    '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the module has only been tested with      '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

Public Function is_prp( _
    aDatabase As Database, _
    ByVal aProperty As String _
) As Boolean
    On Error GoTo Catch
    With aDatabase.properties(aProperty)
        is_prp = True
    End With
    
Finally:
    Exit Function
    
Catch:
    If Err.Number = 3270 Then Resume Finally
    Err.Raise Err.Number
End Function

Public Sub mk_prp( _
    aDatabase As Database, _
    ByVal aProperty As String, _
    ByVal aValue As Variant, _
    Optional ByVal aType As Integer = dbText _
)
    Dim prp As Property
    
    Set prp = aDatabase.CreateProperty(aProperty, aType, aValue)
    aDatabase.properties.Append prp
    Set prp = Nothing
End Sub

Public Sub rm_prp( _
    aDatabase As Database, _
    ByVal aProperty As String _
)
    If is_prp(aDatabase, aProperty) Then aDatabase.properties.Delete aProperty
End Sub

Public Sub set_prp( _
    aDatabase As Database, _
    ByVal aProperty As String, _
    ByVal aValue As Variant, _
    Optional ByVal aType As Integer = -1 _
)
    If Not is_prp(CurrentDb, aProperty) Then
        If aType <> -1 Then _
        mk_prp aDatabase, aProperty, aValue, aType
    Else
        aDatabase.properties(aProperty).Value = aValue
    End If
End Sub

Public Function get_prp( _
    ByVal aProperty As String, _
    Optional aDatabase As Database _
) As Variant
    Dim db As Database
    
    Set db = aDatabase
    If db Is Nothing Then Set db = CurrentDb
    On Error Resume Next
    get_prp = db.properties(aProperty).Value
End Function

Public Sub log_msg( _
    ByVal aFilePath As String, _
    ByVal aMessage As String _
)
    Dim fileno As Integer
    Dim timeStamp As Date
    
    fileno = FreeFile()
    timeStamp = Now()
    Open aFilePath For Append As #fileno
    Print #fileno, timeStamp & vbTab & aMessage
    Close #fileno
End Sub

Public Function arr_add( _
    aArray As Variant, _
    ByVal aValue As Variant _
) As Integer
    If IsArray(aArray) Then
        On Error Resume Next
        ReDim Preserve aArray(UBound(aArray) + 1)
        If Err.Number = 0 Then GoTo Skip
        ReDim aArray(0)
        Err.Clear
Skip:
        If Not IsMissing(aValue) Then aArray(UBound(aArray)) = aValue
        arr_add = UBound(aArray) + 1
    End If
End Function

Public Function base_name( _
    ByVal aFilePath As String _
) As String
    base_name = CreateObject("Scripting.FileSystemObject").GetBaseName(aFilePath)
End Function

