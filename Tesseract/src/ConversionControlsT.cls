VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConversionControlsT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Public Property Set Controls(aControls As VectorT)
    Set vecControls = aControls
End Property

''''''''''''''''''''
' Public Interface '
''''''''''''''''''''

Public Property Get Dirty() As Boolean
    Dirty = boolDirty
End Property

Public Property Let Dirty(aDirty As Boolean)
    If Not aDirty And Dirty Then save
    boolDirty = aDirty
End Property

Public Property Let Enabled(aEnabled As Boolean)
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
    vecControls.PushBack aControl
End Sub

Public Sub Init( _
    aConversionControls As Variant _
)
    Dim item As Variant
    
    Set vecControls = NewVectorT()
    For Each item In aConversionControls
        Dim ctrl As PropertyControlT

        Set ctrl = item
        Add ctrl
    Next
End Sub

Public Sub Load()
    Dim it As Long
    
    For it = 0 To vecControls.Size - 1
        ControlLoad vecControls.At(it)
    Next
    boolDirty = False
End Sub

Public Sub save()
    Dim it As Long
    
    For it = 0 To vecControls.Size - 1
        ControlSave vecControls.At(it)
    Next
    boolDirty = False
End Sub

Public Sub Undo()
    If Dirty Then Load
End Sub

Private Function ControlInvalid() As Control
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
    aControl.Control = PropertyLoad(aControl.PropertyName, aControl.DisplayDefault)
End Sub

Private Sub ControlSave( _
    aControl As PropertyControlT _
)
    PropertyUpdate aControl.PropertyName, aControl.Control.Value, aControl.ValueDefault, aControl.DataType
End Sub

Private Sub ControlsEnable( _
    ByVal aEnable As Boolean _
)
    Dim it As Long
    
    For it = 0 To vecControls.Size - 1
        Dim ctrl As PropertyControlT
        
        Set ctrl = vecControls.At(it)
        ctrl.Control.Enabled = aEnable
    Next
End Sub

Private Function ControlsRequired() As Control()
    Dim ctrls() As Control, it As Long
    
    For it = 0 To vecControls.Size - 1
        Dim ctrl As PropertyControlT
        
        Set ctrl = vecControls.At(it)
        If ctrl.Required Then ArrayPushBack ctrls, ctrl.Control
    Next
End Function

