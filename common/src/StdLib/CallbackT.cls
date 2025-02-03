VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CallbackT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' CallbackT                                                 '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines an aggregate type that contains       '
' information about a callback and that can be coerced to   '
' or from a variant or passed to a late-bound functions.    '
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

Option Compare Database
Option Explicit

Public Target As Object
Public Procedure As String
Public Method As VbCallType

Public Sub Init( _
    aProcedure As String, _
    aMethod As VbCallType, _
    Optional aTarget As Variant = Nothing _
)
    '
    ' Initializes the callback with the given values.
    '
    Set Target = aTarget
    Procedure = aProcedure
    Method = aMethod
End Sub

Public Function Exec( _
    ParamArray aArgs() As Variant _
) As Variant
    '
    ' Executes the callback with the given parameters. Callback procudures
    ' must take a single argument of type variant that is passed as an array.
    '
    If Not Target Is Nothing Then
        Exec = CallByName(Target, Procedure, Method, Array(aArgs)(0))
    Else
        Exec = Application.Run(Procedure, Array(aArgs)(0))
    End If
End Function

