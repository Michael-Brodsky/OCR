VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmConvertible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Binary
Option Explicit

Private cbCallback As CallbackT

Public Property Set Callback(aCallback As CallbackT)
    Set cbCallback = aCallback
End Property

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo Catch
    If Not cbCallback Is Nothing Then cbCallback.Exec , 42
    Exit Sub
    
Catch:
    ErrMessage
End Sub
