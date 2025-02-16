VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSearchDefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' Form_frmSearchDefs                                        '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines types and objects that handle events  '
' and procedures for the frmSearchDefs form.                '
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

Private boolFormUndo As Boolean             ' Flag indicating the Form_Undo event fired before the Cancel button click event.
Private ctrlsSearch As SourceFileControlsT  ' Object that handles search controls behavior.

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
    ctrlsSearch.Current = True
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
    If DataErr = kvbErrDbCantSaveRecord Then Response = acDataErrContinue
End Sub

Private Sub Form_Load()
    On Error GoTo Catch
    MeLoad
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

Private Sub Search_File_Types_Change()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Me.Search_File_Types, ctrlsSearch.RequiredControls)
    ' Also handled in SourceFileControlsT
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Search_File_Types_DblClick(Cancel As Integer)
    ' Handled in SourceFileControlsT
End Sub

Private Sub Search_File_Types_Edit_Click()
    ' Handled in SourceFileControlsT
End Sub

Private Sub Search_Name_BeforeUpdate(Cancel As Integer)
    ' Handled in SourceFileControlsT
End Sub

Private Sub Search_Name_Change()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Me.Search_Name, ctrlsSearch.RequiredControls)
    ' Also handled in SourceFileControlsT
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Search_Path_BeforeUpdate(Cancel As Integer)
    ' Handled in SourceFileControlsT
End Sub

Private Sub Search_Path_Change()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Me.Search_Path, ctrlsSearch.RequiredControls)
    ' Also handled in SourceFileControlsT
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Search_Path_Browse_Click()
    ' Handled in SourceFileControlsT
End Sub

Private Sub Search_Path_DblClick(Cancel As Integer)
    ' Handled in SourceFileControlsT
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

Private Function MeBeforeUpdate() As Boolean
    '
    ' Returns TRUE if all required fields have valid entries,
    ' else returns FALSE.
    '
    Dim Invalid As Control
    
    Set Invalid = ctrlsSearch.Invalid
    MeBeforeUpdate = (Invalid Is Nothing)
    If Not MeBeforeUpdate Then MsgBox Invalid.Name & " cannot be blank."
End Function

Private Sub MeLoad()
    '
    ' Creates any required resources and intializes the form
    ' into a known state.
    '
    Set ctrlsSearch = NewSourceFileControlsT(Me.Search_Name, Me.Search_Path, Me.Search_Path_Browse, _
    Me.Search_File_Types, Me.Search_File_Types_Edit)
End Sub

Private Sub MeUndo()
    '
    ' Sets the control states to their previous state.
    '
    ctrlsSearch.Current = True
    Me.cmdOK.Enabled = False
End Sub

Private Sub MeUnload()
    '
    ' Releases any resources we created.
    '
    Set ctrlsSearch = Nothing
End Sub


