VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OcrConvertT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' OCRConverT                                                '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a class that delegates ocr engine     '
' calls to convert source files into readable text files.   '
' The ocr engine must expose a procedure that executes the  '
' conversion. Each class instance must be initialized with  '
' procedure name and any required arguments.                '
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

Private Procedure As CallbackT  ' Callback object used to invoke the ocr engine conversion procedure.
Private Params As Variant       ' Holds conversion procedure arguments.

Public Property Get Arg( _
    ByVal aIndex As Integer _
) As Variant
    '
    ' Returns an ocr engine procedure argument by index.
    '
    Arg = Params(aIndex)
End Property

Public Property Get Args() As Variant
    '
    ' Returns all ocr engine procedure arguments.
    '
    Args = Params
End Property

Public Function Exec( _
    ByVal aFilePath As String _
) As Variant
    '
    ' Calls the ocr engine procedure with any given arguments.
    '
    Exec = Procedure.Exec(aFilePath, Params)
End Function

Public Sub Init( _
    aProcedure As CallbackT, _
    ParamArray aArgs() As Variant _
)
    '
    ' Initializes a class instance with ocr engine conversion parameters.
    '
    Set Procedure = aProcedure
    Params = CVar(aArgs)
End Sub

