VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConversionControlsT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' ConversionControlsT                                       '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a class that encapsulates controls    '
' in the OCR File Conversion section of the Settings form,  '
' and their group behaviour.                                '
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

Private vecControls As VectorT
Private boolDirty As Boolean

'''''''''''''''''''''''''''
' Public Member Accessors '
'''''''''''''''''''''''''''

Public Property Get Controls() As VectorT
    Set Controls = vecControls
End Property

Public Property Set Controls( _
    aControls As VectorT _
)
    Set vecControls = aControls
End Property

''''''''''''''''''''
' Public Interface '
''''''''''''''''''''

Public Property Get Dirty() As Boolean
    Dirty = boolDirty
End Property

Public Property Let Dirty( _
    ByVal aDirty As Boolean _
)
    If Not aDirty And Dirty Then Save
    boolDirty = aDirty
End Property

Public Property Let Enabled( _
    ByVal aEnabled As Boolean _
)
    ControlsEnable aEnabled
End Property

Public Property Get Invalid() As Control
    Set Invalid = ControlInvalid()
End Property

Public Property Get RequiredControls() As Control()
    RequiredControls = ControlsRequired()
End Property

Public Sub Add( _
    aControl As PropertyControlT _
)
    '
    ' Adds the given control to the collection.
    '
    vecControls.PushBack aControl
End Sub

Public Sub Init( _
    aConversionControls As Variant _
)
    '
    ' Initializes the controls collection.
    '
    Dim item As Variant
    
    Set vecControls = NewVectorT()
    For Each item In aConversionControls
        Dim ctrl As PropertyControlT

        Set ctrl = item
        Add ctrl
    Next
End Sub

Public Sub Load()
    '
    ' Retrieves and displays the current database property
    ' values in the controls.
    '
    Dim it As Long
    
    For it = 0 To vecControls.Size - 1
        ControlLoad vecControls.At(it)
    Next
    boolDirty = False
End Sub

Public Sub Save()
    '
    ' Saves the current control values in their associated
    ' database properties. An error occurrs if all required
    ' controls are not valid.
    '
    Dim it As Long, invld As Control
    
    Set invld = Invalid
    If Not invld Is Nothing Then Err.Raise UsrErr(kOcrAdminErrSettings), CurrentProject.Name, invld.Name & kStrCannotBeBlank
    For it = 0 To vecControls.Size - 1
        ControlSave vecControls.At(it)
    Next
    boolDirty = False
End Sub

Public Sub Undo()
    '
    ' Undoes any changes in control values since the last save.
    '
    If Dirty Then Load
End Sub

Private Function ControlInvalid() As Control
    '
    ' Returns the first control object that has an invalid value, if any.
    '
    Dim it As Long
    
    For it = 0 To vecControls.Size - 1
        Dim ctrl As PropertyControlT
        
        Set ctrl = vecControls.At(it)
        If ctrl.Required And IsNull(ctrl.Control) Then
            Set ControlInvalid = ctrl.Control
            Exit For
        End If
    Next
End Function

Private Sub ControlLoad( _
    aControl As PropertyControlT _
)
    '
    ' Retrieves and displays the current database property value
    ' in the given control.
    '
    aControl.Control = PropertyLoad(aControl.PropertyName, aControl.DisplayDefault)
End Sub

Private Sub ControlSave( _
    aControl As PropertyControlT _
)
    '
    ' Saves the given control's current value in its
    ' associated database property.
    '
    PropertyUpdate aControl.PropertyName, aControl.Control.Value, aControl.ValueDefault, aControl.DataType
End Sub

Private Sub ControlsEnable( _
    ByVal aEnable As Boolean _
)
    '
    ' Enables the controls according to the given argument.
    '
    Dim it As Long
    
    For it = 0 To vecControls.Size - 1
        Dim ctrl As PropertyControlT
        
        Set ctrl = vecControls.At(it)
        ctrl.Control.Enabled = aEnable
    Next
End Sub

Private Function ControlsRequired() As Control()
    '
    ' Returns a list of controls whose Required property is TRUE.
    '
    Dim ctrls() As Control, it As Long
    
    For it = 0 To vecControls.Size - 1
        Dim ctrl As PropertyControlT
        
        Set ctrl = vecControls.At(it)
        If ctrl.Required Then ArrayPushBack ctrls, ctrl.Control
    Next
End Function

