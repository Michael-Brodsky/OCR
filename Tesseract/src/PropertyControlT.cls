VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertyControlT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Binary
Option Explicit

Public Control As Control
Public PropertyName As String
Public ValueDefault As Variant
Public DisplayDefault As Variant
Public DataType As Integer
Public Required As Boolean

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


