VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertyControlT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' PropertyControlsT                                         '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a class that encapsulates controls    '
' bound to database properties and information on how to    '
' load and store the values.                                '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' basCommon                                                 '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the module has only been tested with      '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

Public Control As Control           ' The client control.
Public PropertyName As String       ' Database property associated with the control.
Public ValueDefault As Variant      ' Database property default value.
Public DisplayDefault As Variant    ' Control default display value.
Public DataType As Integer          ' Database property data type.
Public Required As Boolean          ' Flag indicating whether the control is required to have a value.

Public Sub Init( _
    aControl As Control, _
    ByVal aPropertyName As String, _
    Optional ByVal aValueDefault As Variant, _
    Optional ByVal aDisplayDefault As Variant, _
    Optional ByVal aDataType As Integer = kdbNone, _
    Optional ByVal aRequired As Boolean = False _
)
    Set Control = aControl
    PropertyName = aPropertyName
    DisplayDefault = Null
    If Not IsMissing(aValueDefault) Then ValueDefault = aValueDefault
    If Not IsMissing(aDisplayDefault) Then DisplayDefault = aDisplayDefault
    DataType = aDataType
    Required = aRequired
End Sub


