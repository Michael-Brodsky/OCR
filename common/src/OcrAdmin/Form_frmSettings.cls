VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' Form_frmSettings                                          '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines types and objects that handle events  '
' and procedures for the Settings form.                     '
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

'''''''''''''''''''''''
' Types and Constants '
'''''''''''''''''''''''

Private Const kQryProcesses = "qryMyProcesses"
Private Const kPropertyCollapseCatalog As String = "OCRCollapseCatalog"
Private Const kPropertyCollapseDataRules As String = "OCRCollapseDataRules"
Private Const kPropertyCollapseOcrConversion As String = "OCRCollapseOcrConversion"
Private Const kPropertyCollapseSourceFiles As String = "OCRCollapseSourceFiles"
Private Const kStrOnErrorRowSource As String = "3;'Abort';5;'Ignore';4;'Retry'"
Private Const kStrFileNotConvertible As String = "File not convertible"
Private Const kStrSelectFileConvert As String = "Select File To Convert"

''''''''''''''''''''''''
' Module-Level Objects '
''''''''''''''''''''''''

Private boolFormUndo As Boolean                 ' Flag indicating the Form_Undo event fired before the Cancel button click event.
Private ctrlsCatalog As CatalogControlsT        ' Object that encapsulates catalog controls.
Private ctrlsConversion As ConversionControlsT  ' Object that encapsulates ocr conversion controls.
Private ctrlsDataRules As DataRulesControlsT    ' Object that encapsulates data rules controls.
Private ctrlsProcess As ProcessControlsT        ' Object that encapsulates process controls.
Private ctrlsSearch As SourceFileControlsT      ' Object that encapsulates source file controls.
Private listDataRulesEx As ListBoxExT           ' ListBox with custom shortcut menu.

Public Property Get ProcessId() As Long
    '
    ' Used by control classes to get the current process ID.
    '
    If Not Me.NewRecord Then ProcessId = Nz(Me.ID, 0)
End Property

'''''''''''''''''''''''
' Form Event Handlers '
'''''''''''''''''''''''

Private Sub Catalog_Connection_AfterUpdate()
    ' Handled in CatalogControlsT
End Sub

Private Sub Catalog_Connection_Browse_Click()
    ' Handled in CatalogControlsT
End Sub

Private Sub Catalog_Connection_Change()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Nothing, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    ' Also handled in CatalogControlsT
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Catalog_Connection_DblClick(Cancel As Integer)
    ' Handled in CatalogControlsT
End Sub

Private Sub Catalog_Name_AfterUpdate()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Me.Catalog_Name, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    ' Also handled in CatalogControlsT
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Catalog_Name_Change()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Me.Catalog_Name, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    ' Also handled in CatalogControlsT
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Catalog_Name_DblClick(Cancel As Integer)
    cmdCatalogNameEdit_Click
End Sub

Private Sub Catalog_SaveTo_Browse_Click()
    ' Handled in CatalogControlsT
End Sub

Private Sub Catalog_SaveTo_Field_AfterUpdate()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Me.Catalog_SaveTo_Field, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Catalog_SaveTo_Field_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Handled in CatalogControlsT
End Sub

Private Sub Catalog_SaveTo_Method_Click()
    ' Handled in CatalogControlsT
End Sub

Private Sub Catalog_SaveTo_Parameter_BeforeUpdate(Cancel As Integer)
    ' Handled in CatalogControlsT
End Sub

Private Sub Catalog_SaveTo_Parameter_Change()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Me.Catalog_SaveTo_Parameter, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Catalog_SaveTo_Path_Change()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Nothing, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Catalog_SaveTo_Path_DblClick(Cancel As Integer)
    ' Handled in CatalogControlsT
End Sub

Private Sub Catalog_SaveTo_Procedure_BeforeUpdate(Cancel As Integer)
    ' Handled in CatalogControlsT
End Sub

Private Sub Catalog_SaveTo_Procedure_Change()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Me.Catalog_SaveTo_Procedure, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Catalog_SaveTo_Table_AfterUpdate()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Me.Catalog_SaveTo_Table, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    ' Also handled in CatalogControlsT
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Catalog_SaveTo_Table_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Handled in CatalogControlsT
End Sub

Private Sub chkApplicationHide_Click()
    On Error GoTo Catch
    ctrlsConversion.Dirty = True
    Me.cmdOK.Enabled = FormControlChange(Nothing, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub chkProcessActive_Click()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Nothing, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    Exit Sub

Catch:
    ErrMessage
End Sub

Private Sub chkTrim_Click()
    On Error GoTo Catch
    ctrlsConversion.Dirty = True
    Me.cmdOK.Enabled = FormControlChange(Nothing, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    Exit Sub
    
Catch:
    ErrMessage
End Sub

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

Private Sub cmdCatalogDown_Click()
    On Error GoTo Catch
    ToggleCatalog True
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub cmdCatalogNameEdit_Click()
    On Error GoTo Catch
    If FormOpenDialog(kOcrAdminFormCatalogDefs, dlgFind, "[ID]=" & Nz(Me.[Catalog ID], -1)) Then ctrlsCatalog.Requery
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub cmdCatalogUp_Click()
    On Error GoTo Catch
    ToggleCatalog False
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub cmdDataRulesDown_Click()
    On Error GoTo Catch
    ToggleDataRules True
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub cmdDataRulesUp_Click()
    On Error GoTo Catch
    ToggleDataRules False
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo Catch
    DeleteClick
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

Private Sub cmdOcrDown_Click()
    On Error GoTo Catch
    ToggleOcrConversion True
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub cmdOcrUp_Click()
    On Error GoTo Catch
    ToggleOcrConversion False
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub cmdOK_Click()
    On Error GoTo Catch
    ClickOk
    
Finally:
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

Private Sub cmdSearchNameEdit_Click()
    On Error GoTo Catch
    If FormOpenDialog(kOcrAdminFormSearchDefs, dlgFind, "[ID]=" & Nz(Me.[Search ID], -1)) Then ctrlsSearch.Requery
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub cmdSourceFilesDown_Click()
    On Error GoTo Catch
    ToggleSourceFiles True
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub cmdSourceFilesUp_Click()
    On Error GoTo Catch
    ToggleSourceFiles False
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub comboDensity_AfterUpdate()
    On Error GoTo Catch
    ctrlsConversion.Dirty = True
    Me.cmdOK.Enabled = FormControlChange(Nothing, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub comboFormat_AfterUpdate()
    On Error GoTo Catch
    ctrlsConversion.Dirty = True
    Me.cmdOK.Enabled = FormControlChange(Nothing, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub comboOnError_AfterUpdate()
    On Error GoTo Catch
    ctrlsConversion.Dirty = True
    Me.cmdOK.Enabled = FormControlChange(Nothing, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    Exit Sub

Catch:
    ErrMessage

End Sub

Private Sub comboSharpen_AfterUpdate()
    On Error GoTo Catch
    ctrlsConversion.Dirty = True
    Me.cmdOK.Enabled = FormControlChange(Nothing, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Delete_Duplicates_Click()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Nothing, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    Exit Sub

Catch:
    ErrMessage
End Sub

Private Sub Form_AfterInsert()
    On Error GoTo Catch
    MeAfterInsert
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

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Catch
    Me.Ordinal = Nz(DMax("[Ordinal]", "[~tblProcesses]"), 0) + 1    ' Make the new process execution/display order last.
    Exit Sub
    
Catch:
    Cancel = True
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

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Not Cancel Then MeUnload
End Sub

Private Sub Form_Undo(Cancel As Integer)
    On Error GoTo Catch
    boolFormUndo = True
    If Not Cancel Then MeUndo
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub ID_AfterUpdate()
    'Handled in ProcessControlsT
End Sub

Private Sub ID_Change()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Me.ID, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub ID_NotInList(NewData As String, Response As Integer)
    'Handled in ProcessControlsT
End Sub

Private Sub listProcessRules_BeforeUpdate(Cancel As Integer)
    ' Handled in ListBoxExT
End Sub

Private Sub listProcessRules_Click()
    On Error GoTo Catch
    ctrlsDataRules.ListClick
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub listProcessRules_DblClick(Cancel As Integer)
    On Error GoTo Catch
    Cancel = ListProcessRulesDblClick()
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub listProcessRules_Exit(Cancel As Integer)
    ' Handled in ListBoxExT
End Sub

Private Sub listProcessRules_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Handled in ListBoxExT
End Sub

Private Sub Process_Name_Change()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Me.Process_Name, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
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

Private Sub Search_Name_AfterUpdate()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Me.Search_Name, ctrlsSearch.RequiredControls)
    ' Also handled in SourceFileControlsT
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Search_Name_Change()
    ' Handled in SourceFileControlsT
End Sub

Private Sub Search_Name_DblClick(Cancel As Integer)
    cmdSearchNameEdit_Click
End Sub

Private Sub Search_Path_Browse_Click()
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

Private Sub Search_Path_DblClick(Cancel As Integer)
    ' Handled in SourceFileControlsT
End Sub

Private Sub txtResize_Change()
    On Error GoTo Catch
    ctrlsConversion.Dirty = True
    Me.cmdOK.Enabled = FormControlChange(Nothing, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub txtRotate_Change()
    On Error GoTo Catch
    ctrlsConversion.Dirty = True
    Me.cmdOK.Enabled = FormControlChange(Nothing, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub txtTimeout_Change()
    On Error GoTo Catch
    ctrlsConversion.Dirty = True
    Me.cmdOK.Enabled = FormControlChange(Nothing, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
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
    ElseIf ctrlsConversion.Dirty Then
        MeUndo
    ElseIf Not boolFormUndo Then
        FormClose Me
    End If
    boolFormUndo = False
End Sub

Private Sub ClickOk()
    '
    ' Save changes or close form if not dirty.
    '
    Dim wasDirty As Boolean
    
    wasDirty = Me.Dirty Or ctrlsConversion.Dirty
    Me.Dirty = False
    ctrlsConversion.Dirty = False
    If wasDirty And Me.cmdOK.Enabled Then Me.cmdOK.Enabled = False  'Form_AfterUpdate won't fire if only ctrlsConversion were dirty.
    If Not wasDirty Then FormClose Me
End Sub

Private Sub DeleteClick()
    FormDeleteRecord Me
    MeAfterDelete
End Sub

Private Function ListProcessRulesDblClick() As Boolean
    If Me.Dirty Then Me.Dirty = False
    If Not Me.NewRecord Then
        If FormConnectionChanged(Me.Connection) Then FormConnectionSwitch Me.Connection
        ctrlsDataRules.ListDblClick Nz(Me.ID, 0), Me.Catalog_SaveTo_Method.Value, Me.Catalog_SaveTo_Table.Value
    End If
End Function

Private Sub MeAfterInsert()
    Me.Recordset.MoveLast
    ctrlsProcess.Current = True
    ctrlsProcess.Requery
End Sub

Private Sub MeAfterDelete()
    ctrlsProcess.Requery
End Sub

Private Function MeBeforeUpdate() As Boolean
    '
    ' Returns TRUE and saves the corrected values of the "SaveTo" controls
    ' if all required form fields are valid, else returns FALSE.
    '
    MeBeforeUpdate = (ctrlsCatalog.Invalid Is Nothing)
    If MeBeforeUpdate Then ctrlsCatalog.Save
End Function

Private Sub MeCurrent()
    '
    ' Make all control classes current.
    '
    ctrlsProcess.Current = True
    ctrlsCatalog.Current = True
    ctrlsSearch.Current = True
    ctrlsDataRules.Current = True
End Sub

Private Sub MeLoad()
    '
    ' Initializes the form state and control classes.
    '
    Set ctrlsProcess = NewProcessControlsT(Me.ID, Me.Process_Name)
    Set listDataRulesEx = NewListBoxExT(Me.listProcessRules)
    Set ctrlsCatalog = NewCatalogControlsT(Me.Catalog_Name, Me.Catalog_Connection, Me.Catalog_SaveTo_Path, _
    Me.Catalog_SaveTo_Method, Me.Catalog_Connection_Browse, Me.Catalog_SaveTo_Browse, Me.Catalog_SaveTo_Table, _
    Me.Catalog_SaveTo_Field, Me.Catalog_SaveTo_Procedure, Me.Catalog_SaveTo_Parameter, Me.cmdCatalogUp.Visible)
    Set ctrlsConversion = NewConversionControlsT(NewPropertyControlT(Me.txtTimeout, kOcrAdminPropertyAppTimeout, , Null, dbLong), _
    NewPropertyControlT(Me.comboOnError, kOcrAdminPropertyAppOnError, vbIgnore, vbIgnore, dbInteger), NewPropertyControlT( _
    Me.chkApplicationHide, kOcrAdminPropertyAppWindowStyle, False, False, dbBoolean))
    Set ctrlsSearch = NewSourceFileControlsT(Me.Search_Name, Me.Search_Path, Me.Search_Path_Browse, _
    Me.Search_File_Types, Me.Search_File_Types_Edit)
    Set ctrlsDataRules = NewDataRulesControlsT(listDataRulesEx)
    Set ctrlsDataRules.Callback = NewCallbackT("ShortcutCallback", VbMethod, ctrlsDataRules)
    Screen.MousePointer = 1
    FormConnectionReset
    AppLoad
    SectionsInit
    PropertiesInit
End Sub

Private Sub MeUndo()
    '
    ' Reset the current control states and disable the OK button.
    '
    ctrlsConversion.Undo
    MeCurrent
    Me.cmdOK.Enabled = False
End Sub

Private Sub MeUnload()
    '
    ' Save database properties and release any resources we created.
    '
    SectionsUpdate
    FormConnectionClose
    AppUnLoad
    Set ctrlsDataRules.Callback = Nothing
    Set listDataRulesEx = Nothing
    Set ctrlsCatalog = Nothing
    Set ctrlsConversion = Nothing
    Set ctrlsDataRules = Nothing
    Set ctrlsSearch = Nothing
    Set ctrlsProcess = Nothing
    Screen.MousePointer = 0
End Sub

Private Function MeValidate() As Boolean
    '
    ' Returns TRUE if all required controls are valid,
    ' else returns FALSE.
    '
    MeValidate = (FormInvalid(ctrlsProcess.Process) Is Nothing)
    If MeValidate Then MeValidate = (ctrlsSearch.Invalid Is Nothing)
    If MeValidate Then MeValidate = (ctrlsCatalog.Invalid Is Nothing)
End Function

Private Function PropertyControlChange()
    On Error GoTo Catch
    ctrlsConversion.Dirty = True
    Me.cmdOK.Enabled = FormControlChange(Nothing, ctrlsSearch.RequiredControls, ctrlsCatalog.RequiredControls)
    Exit Function
    
Catch:
    ErrMessage
End Function

Private Sub PropertiesInit()
    Me.comboOnError.RowSource = kStrOnErrorRowSource
    PropertiesLoad
End Sub

Private Sub PropertiesLoad()
    ctrlsConversion.Load
End Sub

Private Sub PropertiesUpdate()
    ctrlsConversion.Save
End Sub

Private Sub SectionsInit()
    '
    ' Get the previous collapsible section states from our database properties.
    '
    On Error Resume Next
    If Not PropertyGet(CurrentDb, kPropertyCollapseSourceFiles) Then ToggleSourceFiles False
    If Not PropertyGet(CurrentDb, kPropertyCollapseCatalog) Then ToggleCatalog False
    If Not PropertyGet(CurrentDb, kPropertyCollapseDataRules) Then ToggleDataRules False
    If Not PropertyGet(CurrentDb, kPropertyCollapseOcrConversion) Then ToggleOcrConversion False
End Sub

Private Sub SectionsUpdate()
    '
    ' Save the current collapsible section states in our database properties.
    '
    On Error Resume Next
    PropertySet CurrentDb, kPropertyCollapseSourceFiles, Me.cmdSourceFilesUp.Visible, dbBoolean
    PropertySet CurrentDb, kPropertyCollapseCatalog, Me.cmdCatalogUp.Visible, dbBoolean
    PropertySet CurrentDb, kPropertyCollapseDataRules, Me.cmdDataRulesUp.Visible, dbBoolean
    PropertySet CurrentDb, kPropertyCollapseOcrConversion, Me.cmdOcrUp.Visible, dbBoolean
End Sub

' The following procedures deal with the form's collapsible sections
' and render the form's appearance accordingly. Each section's controls
' encode their relative position and dimensions in their Tag property.
' SS_x defines the section to which the control belongs,
' OF_x indicates the control that defines the section height,
' BL_x indicates which sections the control is below,
' where "x" is the section's id in increasing vertical order, top down.

Private Sub ToggleCatalog( _
    ByVal aVisible As Boolean _
)
    ToggleSection "SS_1", "OF_1", "BL_1", aVisible, Me.cmdCatalogUp, Me.cmdCatalogDown
    'If aVisible Then CurrentMethod
    ctrlsCatalog.Visible = aVisible
End Sub

Private Sub ToggleDataRules( _
    ByVal aVisible As Boolean _
)
    ToggleSection "SS_2", "OF_2", "BL_2", aVisible, Me.cmdDataRulesUp, Me.cmdDataRulesDown
End Sub

Private Sub ToggleOcrConversion( _
    ByVal aVisible As Boolean _
)
    ToggleSection "SS_3", "OF_3", "BL_3", aVisible, Me.cmdOcrUp, Me.cmdOcrDown
End Sub

Private Sub ToggleSourceFiles( _
    ByVal aVisible As Boolean _
)
    ToggleSection "SS_0", "OF_0", "BL_0", aVisible, Me.cmdSourceFilesUp, Me.cmdSourceFilesDown
End Sub

Private Sub ToggleSection( _
    ByVal aSectionTag As Variant, _
    ByVal aOffsetTag As Variant, _
    ByVal aBelowTag As Variant, _
    ByVal aVisible, _
    aButtonUp As CommandButton, _
    aButtonDown As CommandButton _
)
    '
    ' Collapse/expand the given form section.
    '
    Dim ctrl As Control, offset As Variant, factor As Single
    
    For Each ctrl In Me.Controls
        If InStr(1, ctrl.tag, aSectionTag) <> 0 Then
            ctrl.Visible = aVisible
            If InStr(1, ctrl.tag, aOffsetTag) <> 0 Then offset = offset + ctrl.height
        End If
    Next
    offset = offset * IIf(aVisible, 1, -1)
    MoveSectionsBelow aBelowTag, offset
    aButtonUp.Visible = True
    aButtonDown.Visible = True
    DoCmd.GoToControl IIf(aVisible, aButtonUp.Name, aButtonDown.Name)
    aButtonUp.Visible = aVisible
    aButtonDown.Visible = Not aButtonUp.Visible
    Me.InsideHeight = Me.InsideHeight + offset
End Sub

Private Sub MoveSectionsBelow( _
    ByVal aBelowTag As Variant, _
    ByVal aOffset As Long _
)
    '
    ' Keep all form sections and their controls contiguous,
    ' according to the state of sections above them.
    '
    '
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        If InStr(1, ctrl.tag, aBelowTag) <> 0 Then
            ctrl.Top = ctrl.Top + aOffset
        End If
    Next
End Sub

''''''''''''''''''''''''''''''''''
' Application-Specific Interface '
''''''''''''''''''''''''''''''''''

Private Sub AppLoad()
    '
    ' Application-specific initialization procedure.
    '
End Sub

Private Sub AppFileConversionShow()
    '
    ' Displays a file converted to an image according to
    ' the current conversion settings. Useful for tweaking
    ' settings to get a good image for OCR conversion.
    '
    Dim fpath As String
    
    fpath = FilePicker(kStrSelectFileConvert, kOcrAdminStrFilterAllFiles, Nz(Me.Search_Path))
    If fpath <> "" Then
        '
        ' Application-specific code goes here.
        '
    End If
    Exit Sub
    
Catch:
    ShowSpinnyThingForReals False
    Err.Raise Err.Number
End Sub

Private Sub AppUnLoad()
    '
    ' Application-specific termination procedure.
    '
End Sub
