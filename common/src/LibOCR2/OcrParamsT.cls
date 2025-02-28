VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OcrParamsT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' OcrParamsT                                                '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a class that aggregates OCR engine    '
' file conversion parameters.                               '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' None                                                      '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the library has only been tested with     '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

Public Filepath As String           ' Full path to input file to convert.
Public OcrParams As Variant         ' Application-specific conversion parameters.
Public WindowStyle As VbAppWinStyle ' OCR engine window style.
Public Timeout As Long              ' Conversion timeout in milliseconds.
