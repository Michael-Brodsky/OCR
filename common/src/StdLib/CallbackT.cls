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
' LibArray                                                  '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the library has only been tested with     '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database
Option Explicit

''''''''''''''''''
' Public Members '
''''''''''''''''''

Public Target As Object     ' The object, if any, the call procedure is defined in.
Public Procedure As String  ' The name of the procedure to call.
Public Method As VbCallType ' VBA call method.

Public Sub Init( _
    ByVal aProcedure As String, _
    ByVal aMethod As VbCallType, _
    Optional aTarget As Object = Nothing _
)
    '
    ' Initializes the callback with the given arguments.
    '
    Set Target = aTarget
    Procedure = aProcedure
    Method = aMethod
End Sub

Public Function Exec( _
    ParamArray aArgs() As Variant _
) As Variant
    '
    ' Executes the callback with the given arguments. Callback procedures
    ' must take a single argument of type variant that is passed as a
    ' variant array.
    '
    If Not Target Is Nothing Then
        If Not IsMissing(aArgs) Then
            Exec = CallByName(Target, Procedure, Method, CVar(aArgs))
        Else
            Exec = CallByName(Target, Procedure, Method)
        End If
    Else
        If Not IsMissing(aArgs) Then
            Exec = Application.Run(Procedure, CVar(aArgs))
        Else
            Exec = Application.Run(Procedure)
        End If
    End If
End Function

