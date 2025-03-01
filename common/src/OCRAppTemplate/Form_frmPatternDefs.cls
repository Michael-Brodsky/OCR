VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPatternDefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' Form_frmPatternDefs                                       '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines types and objects that handle events  '
' and procedures for the frmPatternDefs form.               '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' basConstants, basCommon                                   '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the module has only been tested with      '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

''''''''''''''''''''''''
' Module-Level Objects '
''''''''''''''''''''''''

Private boolFormUndo As Boolean ' Flag indicating the Form_Undo event fired before the Cancel button.

'''''''''''''''''''''''
' Form Event Handlers '
'''''''''''''''''''''''

Private Sub cmdAddNew_Click()
    On Error GoTo Catch
    DoCmd.GotoRecord , "", acNewRec
    Exit Sub

Catch:
    ErrMessage
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo Catch
    ClickCancel
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo Catch
    FormDeleteRecord Me
    Exit Sub

Catch:
    ErrMessage
End Sub

Private Sub cmdNext_Click()
    On Error GoTo Catch
    FormGotoNext Me
    Exit Sub

Catch:
    ErrMessage
End Sub

Private Sub cmdOK_Click()
    On Error GoTo Catch
    ClickOk
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub cmdPrevious_Click()
    On Error GoTo Catch
    FormGotoPrevious Me
    Exit Sub

Catch:
    ErrMessage
End Sub

Private Sub Form_AfterUpdate()
    On Error GoTo Catch
    Me.cmdOK.Enabled = False
    Exit Sub

Catch:
    ErrMessage
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo Catch
    Cancel = Not MeBeforeUpdate()
    Exit Sub
    
Catch:
    Cancel = True
    ErrMessage
End Sub

Private Sub Form_Current()
    On Error GoTo Catch
    MeCurrent
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
    If DataErr = kvbErrDbCantSaveRecord Then Response = acDataErrContinue
End Sub

Private Sub Form_Load()
    On Error GoTo Catch
    FormLoad Me
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Form_Undo(Cancel As Integer)
    On Error GoTo Catch
    boolFormUndo = True
    If Not Cancel Then MeUndo
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Pattern_BeforeUpdate(Cancel As Integer)
    On Error GoTo Catch
    PatternBeforeUpdate Cancel
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Pattern_Change()
    On Error GoTo Catch
    Me.cmdOK.Enabled = PatternChange()
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Pattern_Name_Change()
    On Error GoTo Catch
    Me.cmdOK.Enabled = PatternNameChange()
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Pattern_Name_BeforeUpdate(Cancel As Integer)
    On Error GoTo Catch
    PatternNameBeforeUpdate Cancel
    Exit Sub
    
Catch:
    ErrMessage
End Sub

'''''''''''''''''''''
' Private Interface '
'''''''''''''''''''''

Private Sub ClickCancel()
    '
    ' Undo changes if form was dirty, else close the form.
    '
    If Me.Dirty Then
        Me.Undo
    ElseIf Not boolFormUndo Then
        FormClose Me
    End If
    boolFormUndo = False
End Sub

Private Sub ClickOk()
    '
    ' Save changes and hide the form.
    '
    Me.Dirty = False
    FormClose Me, True
End Sub

Private Sub ControlsEnable( _
    ByVal aEnable As Boolean _
)
    '
    ' Enable controls in the form's "Pattern" section.
    '
    Me.Pattern.Enabled = aEnable
    Me.Match.Enabled = aEnable
    Me.Submatch.Enabled = aEnable
    Me.Global.Enabled = aEnable
    Me.Ignore_Case.Enabled = aEnable
    Me.Description.Enabled = aEnable
End Sub

Private Function MeBeforeUpdate() As Boolean
    '
    ' Returns TRUE if all required fields have valid entries,
    ' else returns FALSE.
    '
    MeBeforeUpdate = (FormInvalid(Me.Pattern_Name, Me.Pattern) Is Nothing)
End Function

Private Sub MeCurrent()
    '
    ' Enable form controls according to their current state.
    '
    ControlsEnable ((FormInvalid(Me.Pattern_Name, Me.Pattern) Is Nothing))
End Sub

Private Sub MeUndo()
    '
    ' Reset the current control states and disable the OK button.
    '
    MeCurrent
    Me.cmdOK.Enabled = False
End Sub

Private Sub PatternBeforeUpdate( _
    ByRef Cancel As Integer _
)
    Cancel = (Me.Pattern.Text = "")
    If Cancel Then MsgBox Me.Pattern.Name & kStrCannotBeBlank
End Sub

Private Function PatternChange() As Boolean
    If (Me.Pattern.Text <> "") Then
        ControlsEnable True
        PatternChange = (FormInvalid(Me.Pattern_Name) Is Nothing)
    End If
End Function

Private Sub PatternNameBeforeUpdate( _
    ByRef Cancel As Integer _
)
    Cancel = (Me.Pattern_Name.Text = "")
    If Cancel Then MsgBox Me.Pattern_Name.Name & kStrCannotBeBlank
End Sub

Private Function PatternNameChange() As Boolean
    If (Me.Pattern_Name.Text <> "") Then
        ControlsEnable True
        PatternNameChange = (FormInvalid(Me.Pattern) Is Nothing)
    End If
End Function


