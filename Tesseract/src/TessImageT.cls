VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TessImageT"
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
' This module defines a class that aggregates ImageMagick   '
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

Public Density As Integer         ' Image density in dpi,
Public Format As String           ' Image file format,
Public Resize As String           ' Image resize dimensions as WxH in pixels,
Public Rotate As Integer          ' Image rotation angle in degrees,
Public Sharpen As String          ' Image sharpen sigma value,
Public Trim As Boolean            ' Image trim flag.

