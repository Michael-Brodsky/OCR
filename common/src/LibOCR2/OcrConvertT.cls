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
' OCRConvertT                                                '
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

Public Procedure As CallbackT  ' Callback object that invokes the ocr engine conversion procedure.
Public Params As Variant       ' Conversion procedure arguments.

Public Function Exec( _
    ByVal aFilePath As String _
) As Variant
    '
    ' Calls the ocr engine procedure with the file to convert and any given parameters.
    '
    Exec = Procedure.Exec(aFilePath, Params)
End Function


