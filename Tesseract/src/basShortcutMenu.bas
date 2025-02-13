Attribute VB_Name = "basShortcutMenu"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' basRulesShortcutMenu                                      '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines objects for creating and using the    '
' shortcut menu associated with the OCR Data Rules ListBox  '
' on the Settings form.                                     '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' LibVBA                                                    '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the module has only been tested with      '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

'''''''''''''''''''''''
' Types and Constants '
'''''''''''''''''''''''

' Enumerates valid values for the menu items bitmask.
Public Enum RulesShortCutMask
    rsmCut = 1
    rsmCopy = 2
    rsmPaste = 4
    rsmDelete = 8
    rsmUndo = 16
End Enum

' Enumerates valid shortcut menu actions.
Public Enum RulesShortcutAction
    odsCopy = 1
    odsCut = 2
    odsDelete = 3
    odsPaste = 4
    odsUndo = 5
End Enum

Public Const kRulesShortcutName As String = "barOcrDataEdit"

''''''''''''''''''''''''
' Module-Level Objects '
''''''''''''''''''''''''

Private cbCallback As CallbackT   ' The current client callback.

'''''''''''''''''''''''''''
' Client Public Interface '
'''''''''''''''''''''''''''

Public Sub RulesShortcutCreate( _
    aCallback As CallbackT _
)
    '
    ' Clients must initialize the callback before using the shortcut menu.
    '
    Set cbCallback = aCallback
End Sub

Public Sub RulesShortcutDestroy()
    '
    ' Clients should call this to free up any resources.
    '
    Set cbCallback = Nothing
End Sub

'''''''''''''''''''''''''''''''''
' Menu Command Event Handlers   '
'                               '
' These functions handle events '
' raised by the menu command    '
' buttons and execute client    '
' callbacks with the button     '
' action parameter.             '
'''''''''''''''''''''''''''''''''

Public Function RulesShortcutCopy()
    On Error GoTo Catch
    MenuExecute odsCopy
    Exit Function
    
Catch:
    ErrMessage
End Function

Public Function RulesShortcutCut()
    On Error GoTo Catch
    MenuExecute odsCut
    Exit Function
    
Catch:
    ErrMessage
End Function

Public Function RulesShortcutDelete()
    On Error GoTo Catch
    MenuExecute odsDelete
    Exit Function
    
Catch:
    ErrMessage
End Function

Public Function RulesShortcutPaste()
    On Error GoTo Catch
    MenuExecute odsPaste
    Exit Function
    
Catch:
    ErrMessage
End Function

Public Function RulesShortcutUndo()
    On Error GoTo Catch
    MenuExecute odsUndo
    Exit Function
    
Catch:
    ErrMessage
End Function

'''''''''''''''''''''
' Private Interface '
'''''''''''''''''''''

Private Sub MenuExecute( _
    ByVal aAction As RulesShortcutAction _
)
    '
    ' Executes a client callback.
    '
    If IsSomething(cbCallback) Then cbCallback.Exec aAction
End Sub

'''''''''''''''''''''''''
' Menu Build Procedures '
'''''''''''''''''''''''''

Private Sub BuildShortcut()
    '
    ' Rebuilds the shortcut menu.
    '
    Dim bar As Office.CommandBar, btn As CommandBarButton, i As Integer
        
    Set bar = ResetShortcut(kRulesShortcutName)
    Set btn = bar.Controls.Add(msoControlButton)
    btn.Caption = "Cut"
    btn.OnAction = "=RulesShortcutCut()"
    btn.tag = odsCut
    Set btn = bar.Controls.Add(msoControlButton)
    btn.Caption = "Copy"
    btn.tag = odsCopy
    btn.OnAction = "=RulesShortcutCopy()"
    Set btn = bar.Controls.Add(msoControlButton)
    btn.Caption = "Paste"
    btn.OnAction = "=RulesShortcutPaste()"
    btn.tag = odsPaste
    Set btn = bar.Controls.Add(msoControlButton)
    btn.Caption = "Delete"
    btn.OnAction = "=RulesShortcutDelete()"
    btn.tag = odsDelete
    Set btn = bar.Controls.Add(msoControlButton)
    btn.Caption = "Undo"
    btn.OnAction = "=RulesShortcutUndo()"
    btn.tag = odsUndo
    Set bar = Nothing
End Sub

Private Function ResetShortcut( _
    ByVal aName As String _
) As Office.CommandBar
    '
    ' Removes the current shortcut menu from the
    ' application collection so it can be rebuilt.
    '
    On Error Resume Next
    CommandBars(aName).Delete
    On Error GoTo 0
    Set ResetShortcut = CommandBars.Add(aName, msoBarPopup, False)
End Function


