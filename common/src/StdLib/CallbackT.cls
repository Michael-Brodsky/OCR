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
    Optional ByVal aMethod As VbCallType, _
    Optional aTarget As Object = Nothing _
)
    '
    ' Initializes a callback with the given arguments.
    '
    Set Target = aTarget
    Procedure = aProcedure
    Method = aMethod
End Sub

Public Function Exec( _
    ParamArray aArgs() As Variant _
) As Variant
    '
    ' Executes a callback with the given arguments. Callback parameters
    ' are passed as a variant array.
    '
    Dim parms() As Variant, tmp() As Variant, v As Variant
    
    ' This allows passing the paramarray args to the callback procedure
    ' whether any arguments were given or not, without hosing them up,
    ' allegedly.
    On Error GoTo Catch
    parms = aArgs
    Do While VarType(parms(0)) >= vbArray And UBound(parms) < 1
        tmp = parms(0)
        parms = tmp
    Loop
No_Args:
    On Error GoTo 0
    If IsSomething(Target) Then
        If UBound(parms) > -1 Then
            Exec = CallByName(Target, Procedure, Method, parms)
        Else
            Exec = CallByName(Target, Procedure, Method)    ' Specialization for procedures taking no parameters.
        End If
    Else
        If UBound(parms) > -1 Then
            Exec = Application.Run(Procedure, parms)
        Else
            Exec = Application.Run(Procedure)               ' Specialization for procedures taking no parameters.
        End If
    End If
    Exit Function
    
Catch:
    If Err.Number = kvbErrSubscriptOutOfRange Then Resume No_Args
    Err.Raise Err.Number
End Function


