VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmReadable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' Form_frmReadable                                          '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines types and objects that handle events  '
' and procedures for the frmReadable form.                  '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' basConstants                                              '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the module has only been tested with      '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

Private cbCallback As CallbackT ' Parent form callback.

Public Property Set Callback(aCallback As CallbackT)
    Set cbCallback = aCallback
End Property

Private Sub Form_BeforeUpdate(Cancel As Integer)
    '
    ' Execute a callback to the parent form indicating changes were made.
    '
    On Error GoTo Catch
    If Not cbCallback Is Nothing Then cbCallback.Exec , 42
    Exit Sub
    
Catch:
    ErrMessage
End Sub

