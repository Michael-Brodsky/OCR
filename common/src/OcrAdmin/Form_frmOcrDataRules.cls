VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmOcrDataRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' Form_frmOcrDataRules                                      '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines types and objects that handle events  '
' and procedures for the frmOcrDataRules form.              '
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

Private lngProcessID As Long    ' Process ID to use when inserting new rules.
Private boolFormUndo As Boolean ' Flag indicating the Form_Undo event fired before the Cancel button click event.

''''''''''''''''''''''''''''
' Private Member Accessors '
''''''''''''''''''''''''''''

Private Property Get ProcessId() As Long
    ProcessId = lngProcessID
End Property

Private Property Get StorageControl() As Control
    Set StorageControl = IIf(Me.Storage_Parameter.Visible, Me.Storage_Parameter, _
    IIf(Me.Storage_Field.Visible, Me.Storage_Field, Nothing))
End Property

Private Property Let ProcessId( _
    ByVal aID As Long _
)
    '
    ' Limit the rules list to the current process.
    '
    Me.Filter = "[Process ID] = " & Str(aID)
    Me.FilterOn = True
    lngProcessID = aID
End Property

Private Property Let StorageFields( _
    ByVal aList As String _
)
    Me.Storage_Field.RowSource = aList
End Property

Private Property Let StorageMethod( _
    ByVal aMethod As DataStorageMethod _
)
    '
    ' Display the storage control appropriate for the
    ' current process storage method.
    '
    Select Case aMethod
        Case osmNone:
        Case osmBuiltIn:
            Me.Storage_Field.RowSource = TableFieldsList(Split(Me.OpenArgs, ",")(4))    ' Populate the storage field's list.
            Me.Storage_Field.Visible = True
        Case osmCustom:
            Me.Storage_Field.Visible = False
        Case Else:
    End Select
        Me.Storage_Parameter.Visible = Not Me.Storage_Field.Visible
End Property

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

Private Sub cmdPatternName_Click()
    On Error GoTo Catch
    PatternDefsEdit
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
    'Me.cmdOK.Enabled = False
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Catch
    Me.[Process ID] = ProcessId
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
    If DataErr = kvbErrDbCantSaveRecord Or DataErr = kvbErrDbNoCurrentRecord Then Response = acDataErrContinue
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

Private Sub Pattern_Name_AfterUpdate()
    On Error GoTo Catch
    Me.cmdOK.Enabled = PatternNameAfterUpdate()
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Rule_Name_BeforeUpdate(Cancel As Integer)
    On Error GoTo Catch
    RuleNameBeforeUpdate Cancel
    Exit Sub
    
Catch:
    Cancel = True
    ErrMessage
End Sub

Private Sub Rule_Name_Change()
    On Error GoTo Catch
    Me.cmdOK.Enabled = RuleNameChange()
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Storage_Field_AfterUpdate()
    On Error GoTo Catch
    StorageFieldAfterUpdate
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Storage_Field_BeforeUpdate(Cancel As Integer)
    On Error GoTo Catch

    Exit Sub
    
Catch:
    Cancel = True
    ErrMessage
End Sub

Private Sub Storage_Parameter_BeforeUpdate(Cancel As Integer)
    On Error GoTo Catch
    StorageParameterBeforeUpdate Cancel
    Exit Sub
    
Catch:
    Cancel = True
    ErrMessage
End Sub

Private Sub Storage_Parameter_Change()
    On Error GoTo Catch
    Me.cmdOK.Enabled = StorageParameterChange()
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
    ' Save changes and close the form.
    '
    Dim wasDirty As Boolean
    
    wasDirty = Me.Dirty
    Me.Dirty = False
    If Not wasDirty Then FormClose Me, True
End Sub

Private Sub ControlsEnable( _
    ByVal aEnable As Boolean _
)
    '
    ' Enable all form controls according to the given value.
    '
    EnableStorageControls aEnable
    EnablePatternControls aEnable
End Sub

Private Sub EnablePatternControls( _
    ByVal aEnable As Boolean _
)
    '
    ' Enable the "Pattern" controls according to the given value.
    '
    Me.Pattern_Name.Enabled = aEnable
    Me.Pattern.Enabled = aEnable
    Me.Match.Enabled = aEnable
    Me.Submatch.Enabled = aEnable
    Me.Global.Enabled = aEnable
    Me.Ignore_Case.Enabled = aEnable
    Me.Pattern_Description.Enabled = aEnable
End Sub

Private Sub EnableStorageControls( _
    ByVal aEnable As Boolean _
)
    '
    ' Enable the "Data Storage" controls according to the given value.
    '
    StorageControl.Enabled = aEnable
    Me.Required.Enabled = aEnable
End Sub

Private Function MeBeforeUpdate() As Boolean
    '
    ' Returns TRUE if all required fields have valid entries
    ' and saves the Storage Method entries, else returns FALSE.
    '
    If (FormInvalid(Me.Rule_Name) Is Nothing) Then
        If Me.Storage_Field.Visible Then
            MeBeforeUpdate = (FormInvalid(Me.Storage_Field) Is Nothing)
        ElseIf Me.Storage_Parameter.Visible Then
            MeBeforeUpdate = (FormInvalid(Me.Storage_Parameter) Is Nothing)
        End If
    End If
End Function

Private Sub MeCurrent()
    '
    ' Set the control states according to the current rule's settings.
    '
    ControlsEnable (FormInvalid(StorageControl, Me.Pattern_Name, Me.Pattern) Is Nothing)
End Sub

Private Sub MeLoad()
    '
    ' Initializes the form according to the its OpenArgs.
    '
    On Error Resume Next
    ProcessId = Split(Me.OpenArgs, ",")(2)
    StorageMethod = Split(Me.OpenArgs, ",")(3)
End Sub

Private Sub MeUndo()
    '
    ' Reset the current control states and disable the OK button.
    '
    MeCurrent
    Me.cmdOK.Enabled = False
End Sub

Private Sub PatternDefsEdit()
    '
    ' Open the pattern editor form and reflect any
    ' changes here when done.
    '
    FormOpenDialog kOcrAdminFormOcrDataRules, dlgFind, "[ID]=" & Str(Me.ID)
    Me.Pattern_Description.Requery
End Sub

Private Function PatternNameAfterUpdate() As Boolean
    If (Me.Pattern_Name.Text <> "") Then
        PatternNameAfterUpdate = (FormInvalid(StorageControl, Me.Pattern_Name, Me.Pattern) Is Nothing)
    End If
End Function

Private Sub RuleNameBeforeUpdate( _
    ByRef Cancel As Integer _
)
    Cancel = (Me.Rule_Name.Text = "")
    If Cancel Then MsgBox Me.Rule_Name.Name & " cannot be blank"
End Sub

Private Function RuleNameChange() As Boolean
    If (Me.Rule_Name.Text <> "") Then
        EnableStorageControls True
        RuleNameChange = (FormInvalid(StorageControl, Me.Pattern_Name, Me.Pattern) Is Nothing)
    End If
End Function

Private Sub StorageParameterBeforeUpdate( _
    ByRef Cancel As Integer _
)
    Cancel = (Me.Storage_Parameter.Text = "")
    If Cancel Then MsgBox Me.Storage_Parameter.Name & " cannot be blank"
End Sub

Private Sub StorageFieldAfterUpdate()
    Me.cmdOK.Enabled = True
    EnablePatternControls True
End Sub
Private Function StorageParameterChange() As Boolean
    If (Me.Storage_Parameter.Text <> "") Then
        EnablePatternControls True
        StorageParameterChange = (FormInvalid(Me.Rule_Name, Me.Pattern_Name, Me.Pattern) Is Nothing)
    End If
End Function
