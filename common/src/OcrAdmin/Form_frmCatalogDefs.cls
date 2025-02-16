VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCatalogDefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' Form_frmCatalogDefs                                       '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines types and objects that handle events  '
' and procedures for the frmCatalogDefs form.               '
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
Private ctrlsCatalog As CatalogControlsT    ' Object that handles catalog controls behavior.

'''''''''''''''''''''''
' Form Event Handlers '
'''''''''''''''''''''''

Private Sub Catalog_Connection_AfterUpdate()
    ' Handled in CatalogControlsT
End Sub

Private Sub Catalog_Connection_Change()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Nothing, ctrlsCatalog.RequiredControls)
    ' Also handled in CatalogControlsT
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Catalog_Connection_DblClick(Cancel As Integer)
    ' Handled in CatalogControlsT
End Sub

Private Sub Catalog_Name_BeforeUpdate(Cancel As Integer)
    ' Handled in CatalogControlsT
End Sub

Private Sub Catalog_Name_Change()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Me.Catalog_Name, ctrlsCatalog.RequiredControls)
    ' Also handled in CatalogControlsT
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Catalog_SaveTo_Field_AfterUpdate()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Me.Catalog_SaveTo_Field, ctrlsCatalog.RequiredControls)
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
    Me.cmdOK.Enabled = FormControlChange(Me.Catalog_SaveTo_Parameter, ctrlsCatalog.RequiredControls)
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Catalog_SaveTo_Path_Change()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Nothing, ctrlsCatalog.RequiredControls)
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
    Me.cmdOK.Enabled = FormControlChange(Me.Catalog_SaveTo_Procedure, ctrlsCatalog.RequiredControls)
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Catalog_SaveTo_Table_AfterUpdate()
    On Error GoTo Catch
    Me.cmdOK.Enabled = FormControlChange(Me.Catalog_SaveTo_Table, ctrlsCatalog.RequiredControls)
    ' Also handled in CatalogControlsT
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub Catalog_SaveTo_Table_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Handled in CatalogControlsT
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
    ctrlsCatalog.Current = True
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
    Me.Dirty = False
    FormClose Me, True
End Sub

Private Function MeBeforeUpdate() As Boolean
    '
    ' Returns TRUE and saves the corrected values of the "SaveTo" controls
    ' if all required form fields are valid, else returns FALSE.
    '
    MeBeforeUpdate = (ctrlsCatalog.Invalid Is Nothing)
    If MeBeforeUpdate Then ctrlsCatalog.Save
End Function

Private Sub MeLoad()
    '
    ' Initializes the catalog controls object.
    '
    Set ctrlsCatalog = NewCatalogControlsT(Me.Catalog_Name, Me.Catalog_Connection, Me.Catalog_SaveTo_Path, _
    Me.Catalog_SaveTo_Method, Me.Catalog_Connection_Browse, Me.Catalog_SaveTo_Browse, Me.Catalog_SaveTo_Table, _
    Me.Catalog_SaveTo_Field, Me.Catalog_SaveTo_Procedure, Me.Catalog_SaveTo_Parameter, , True)
End Sub

Private Sub MeUndo()
    '
    ' Updates the catalog controls current state
    ' and disables the OK button.
    '
    ctrlsCatalog.Current = True
    Me.cmdOK.Enabled = False
End Sub

Private Sub MeUnload()
    '
    ' Destroys the catalog controls object.
    '
    Set ctrlsCatalog = Nothing
End Sub
