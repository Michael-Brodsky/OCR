VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CatalogControlsT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' CatalogControlsT                                          '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a class that encapsulates controls    '
' bound to "Catalog" fields and defines common behaviour    '
' accross multiple forms.                                   '
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

'''''''''''''''''''''''
' Types and Constants '
'''''''''''''''''''''''

Private Const kStrStorageField As String = "Storage Field"
Private Const kStrStorageParameter As String = "Storage Parameter"

'''''''''''''''''''
' Private Members '
'''''''''''''''''''

Private WithEvents comboCatalogName As ComboBox
Attribute comboCatalogName.VB_VarHelpID = -1
Private WithEvents txtCatalogName As TextBox
Attribute txtCatalogName.VB_VarHelpID = -1
Private WithEvents txtConnection As TextBox
Private WithEvents cmdConnectionBrowse As CommandButton
Private WithEvents txtSaveToPath As TextBox
Private WithEvents cmdSaveToPathBrowse As CommandButton
Private WithEvents optionMethod As OptionGroup
Attribute optionMethod.VB_VarHelpID = -1
Private WithEvents comboTable As ComboBox
Private txtProcedure As TextBox
Private WithEvents comboField As ComboBox
Private txtParameter As TextBox
Private boolVisible As Boolean

'''''''''''''''''''''''''''
' Public Member Accessors '
'''''''''''''''''''''''''''

Public Property Get CatalogControl() As Control
    Set CatalogControl = IIf(Not comboCatalogName Is Nothing, comboCatalogName, txtCatalogName)
End Property

Public Property Set CatalogControl(aControl As Control)
    Set comboCatalogName = IIf(TypeName(aControl) = "ComboBox", aControl, Nothing)
    Set txtCatalogName = IIf(TypeName(aControl) = "TextBox", aControl, Nothing)
End Property

Public Property Get ConnectionBrowse() As CommandButton
    Set ConnectionBrowse = cmdConnectionBrowse
End Property

Public Property Set ConnectionBrowse(aButton As CommandButton)
    Set cmdConnectionBrowse = aButton
End Property

Public Property Get ConnectionControl() As TextBox
    Set ConnectionControl = txtConnection
End Property

Public Property Set ConnectionControl(aTextBox As TextBox)
    Set txtConnection = aTextBox
End Property

Public Property Get FieldControl() As ComboBox
    Set FieldControl = comboField
End Property

Public Property Set FieldControl(aListBox As ComboBox)
    Set comboField = aListBox
End Property

Public Property Get MethodControl() As OptionGroup
    Set MethodControl = optionMethod
End Property

Public Property Set MethodControl(aOptionGroup As OptionGroup)
    Set optionMethod = aOptionGroup
End Property

Public Property Get ParameterControl() As TextBox
    Set ParameterControl = txtParameter
End Property

Public Property Set ParameterControl(aTextBox As TextBox)
    Set txtParameter = aTextBox
End Property

Public Property Get ProcedureControl() As TextBox
    Set ProcedureControl = txtProcedure
End Property

Public Property Set ProcedureControl(aTextBox As TextBox)
    Set txtProcedure = aTextBox
End Property

Public Property Get SaveToBrowse() As CommandButton
    Set SaveToBrowse = cmdSaveToPathBrowse
End Property

Public Property Set SaveToBrowse(aButton As CommandButton)
    Set cmdSaveToPathBrowse = aButton
End Property

Public Property Get SaveToControl() As TextBox
    Set SaveToControl = txtSaveToPath
End Property

Public Property Set SaveToControl(aTextBox As TextBox)
    Set txtSaveToPath = aTextBox
End Property

Public Property Get TableControl() As ComboBox
    Set TableControl = comboTable
End Property

Public Property Set TableControl(aComboBox As ComboBox)
    Set comboTable = aComboBox
End Property

''''''''''''''''''''
' Public Interface '
''''''''''''''''''''

Public Property Let Current(aCurrent As Boolean)
    If aCurrent Then StorageMethod = OcrStorageMethod(TableControl, ProcedureControl)
End Property

Public Property Let Enabled(aEnabled As Boolean)
    ControlsEnable True
End Property

Public Property Get Invalid() As Control
    Set Invalid = ControlsInvalid()
End Property

Public Property Get RequiredControls() As Control()
    RequiredControls = ControlsRequired()
End Property

Public Property Get StorageMethod() As DataStorageMethod
    If Not MethodControl Is Nothing Then StorageMethod = MethodControl.Value
End Property

Public Property Let StorageMethod( _
    ByVal aStorageMethod As DataStorageMethod _
)
    If boolVisible Then ControlsShow aStorageMethod
    If Not MethodControl Is Nothing Then MethodControl.Value = aStorageMethod
End Property

Public Property Get Visible() As Boolean
    Visible = boolVisible
End Property

Public Property Let Visible(aVisible As Boolean)
    If aVisible Then ControlsShow StorageMethod
    boolVisible = aVisible
End Property

Public Sub Init( _
    aCatalog As Control, _
    aConnection As TextBox, _
    aSaveToPath As TextBox, _
    aMethod As OptionGroup, _
    aBrowseConnection As CommandButton, _
    aBrowseSaveTo As CommandButton, _
    aTable As ComboBox, _
    aField As ComboBox, _
    aProcedure As TextBox, _
    aParameter As TextBox, _
    aVisible As Boolean _
)
    Set CatalogControl = aCatalog
    Set ConnectionControl = aConnection
    Set SaveToControl = aSaveToPath
    Set MethodControl = aMethod
    Set ConnectionBrowse = aBrowseConnection
    Set SaveToBrowse = aBrowseSaveTo
    Set TableControl = aTable
    Set FieldControl = aField
    Set ProcedureControl = aProcedure
    Set ParameterControl = aParameter
    Visible = aVisible
    StorageMethod = MethodControl.Value
End Sub

Public Sub Requery()
    CatalogControl.Requery
    ConnectionControl.Requery
    SaveToControl.Requery
    TableControl.Requery
    FieldControl.Requery
    ProcedureControl.Requery
    ParameterControl.Requery
End Sub

Public Sub save()
    '
    ' Save only those settings used by the current storage method and
    ' clear the unused ones.
    '
    Select Case StorageMethod
        Case osmBuiltIn:
            ProcedureControl = Null
            ParameterControl = Null
        Case osmCustom:
            TableControl = Null
            FieldControl = Null
        Case Else:
    End Select
End Sub

''''''''''''''''''
' Event Handlers '
''''''''''''''''''

Private Sub cmdConnectionBrowse_Click()
    On Error GoTo Catch
    If ControlPickPath(ConnectionControl, pkFile, kOcrAdminStrSelectDatabase, kOcrAdminStrFilterDatabase) Then _
    txtConnection_AfterUpdate
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub cmdSaveToPathBrowse_Click()
    On Error GoTo Catch
    Call ControlPickPath(SaveToControl, pkFolder, kOcrAdminStrSelectSaveToPath)
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub comboCatalogName_AfterUpdate()
    On Error GoTo Catch
    StorageMethod = OcrStorageMethod(TableControl, ProcedureControl)
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub comboCatalogName_Change()
    On Error GoTo Catch
    EnableDataEntry comboCatalogName
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub comboField_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Catch
    If FormConnectionChanged(ConnectionControl) Then ControlsUpdateLists
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub comboTable_AfterUpdate()
    On Error GoTo Catch
    FieldsUpdateList
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub comboTable_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Catch
    If FormConnectionChanged(ConnectionControl) Then ControlsUpdateLists
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub optionMethod_Click()
    On Error GoTo Catch
    ControlsShow MethodControl.Value
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub txtCatalogName_BeforeUpdate(Cancel As Integer)
    On Error GoTo Catch
    Cancel = (txtCatalogName.Text = "")
    If Cancel Then MsgBox txtCatalogName.Name & " cannot be blank"
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub txtCatalogName_Change()
    On Error GoTo Catch
    EnableDataEntry txtCatalogName
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub txtConnection_AfterUpdate()
    On Error GoTo Catch
    ControlsClear
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub txtConnection_DblClick(Cancel As Integer)
    On Error GoTo Catch
    cmdConnectionBrowse_Click
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub txtParameter_BeforeUpdate(Cancel As Integer)
    On Error GoTo Catch
    Cancel = (txtParameter.Text = "")
    If Cancel Then MsgBox txtParameter.Name & " cannot be blank"
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub txtProcedure_BeforeUpdate(Cancel As Integer)
    On Error GoTo Catch
    Cancel = (txtProcedure.Text = "")
    If Cancel Then MsgBox txtProcedure.Name & " cannot be blank"
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub txtSaveToPath_DblClick(Cancel As Integer)
    cmdSaveToPathBrowse_Click
End Sub

'''''''''''''''''''''
' Private Interface '
'''''''''''''''''''''

Private Sub ControlsClear()
    '
    ' Clear the Storage Table and Field values
    ' and their dropdown lists.
    '
    TableControl.Value = Null
    TableControl.RowSource = ""
    FieldControl.Value = Null
    FieldControl.RowSource = ""
End Sub

Private Sub ControlsEnable( _
    ByVal aEnabled As Boolean, _
    Optional ByVal aAll As Boolean = True _
)
    '
    ' Enable the Connection, SaveTo and Storage Method controls,
    ' and the Storage Table and Field if the aAll flag is set.
    '
    If aAll Then
        TableControl.Enabled = aEnabled
        FieldControl.Enabled = aEnabled
        ProcedureControl.Enabled = aEnabled
        ParameterControl.Enabled = aEnabled
    End If
    ConnectionControl.Enabled = aEnabled
    ConnectionBrowse.Enabled = aEnabled
    SaveToControl.Enabled = aEnabled
    SaveToBrowse.Enabled = aEnabled
    MethodControl.Enabled = aEnabled
End Sub

Private Sub ControlsShow( _
    ByVal aStorageMethod As DataStorageMethod _
)
    '
    ' Show only those storage controls used by the current storage method.
    '
    On Error Resume Next
    Select Case aStorageMethod
        Case osmNone:
            ControlsEnable False
        Case osmBuiltIn:
            ControlsEnable True
            If Not TableControl Is Nothing Then
                TableControl.Visible = True
                ControlsSwitch TableControl, ProcedureControl, ParameterControl
                If Not FieldControl Is Nothing Then FieldControl.Visible = True
            ElseIf Not FieldControl Is Nothing Then
                FieldControl.Visible = True
                ControlsSwitch FieldControl, ProcedureControl, ParameterControl
            End If
            If Not ProcedureControl Is Nothing Then ProcedureControl.Visible = Not TableControl.Visible
            If Not ParameterControl Is Nothing Then ParameterControl.Visible = ProcedureControl.Visible
        Case osmCustom:
            ControlsEnable True
            If Not ProcedureControl Is Nothing Then
                ProcedureControl.Visible = True
                ControlsSwitch ProcedureControl, TableControl, FieldControl
                If Not ParameterControl Is Nothing Then ParameterControl.Visible = True
            ElseIf Not ParameterControl Is Nothing Then
                ParameterControl.Visible = True
                ControlsSwitch ParameterControl, TableControl, FieldControl
            End If
            If Not TableControl Is Nothing Then TableControl.Visible = Not ProcedureControl.Visible
            If Not FieldControl Is Nothing Then FieldControl.Visible = TableControl.Visible
        Case Else:
    End Select
End Sub

Private Sub ControlsSwitch( _
    aToControl As Variant, _
    ParamArray aFromControls() As Variant _
)
    '
    ' Switch the focus away from the From controls
    ' and to the To control, if necessary.
    '
    Dim frm As Form, ctrl As Variant
    
    Set frm = aToControl.Parent
    For Each ctrl In aFromControls
        If Not ctrl Is Nothing Then
            If ctrl Is frm.ActiveControl Then
                aToControl.SetFocus
                Exit For
            End If
        End If
    Next
    Set frm = Nothing
End Sub
    
Private Sub ControlsUpdateLists()
    '
    ' Update the Storage Table and Field dropdown lists
    ' for the current database.
    '
    FormConnectionSwitch ConnectionControl
    TableControl.RowSource = DatabaseTablesList()
    FieldsUpdateList
End Sub

Private Function ControlsInvalid() As Control
    '
    ' Returns the first control not having a valid entry,
    ' or Nothing if all controls are valid.
    '
    Set ControlsInvalid = FormInvalid(CatalogControl)
    If ControlsInvalid Is Nothing Then
        Select Case StorageMethod
            Case osmBuiltIn:
                Set ControlsInvalid = IIf(IsNull(TableControl), TableControl, _
                IIf(IsNull(FieldControl), FieldControl, Nothing))
            Case osmCustom:
                Set ControlsInvalid = IIf(IsNull(ProcedureControl), ProcedureControl, _
                IIf(IsNull(ParameterControl), ParameterControl, Nothing))
            Case Else:
        End Select
    End If
End Function

Private Sub EnableDataEntry( _
    aSender As Control _
)
    '
    ' Enable controls in stages, according to what has
    ' and hasn't been entered on the form.
    '
    If aSender.Text <> "" Then ControlsEnable True, False
End Sub

Private Sub FieldsUpdateList( _
    Optional ByVal aClear As Boolean = False _
)
    '
    ' Update the Storage Field dropdown list with the currently selected
    ' Storage Table's available fields.
    '
    If aClear Then FieldControl = Null
    FieldControl.RowSource = TableFieldsList(TableControl)
End Sub

Private Function ControlsRequired() As Control()
    '
    ' Returns a list of controls that require valid entries.
    '
    Dim ctrls() As Control
    
    ArrayPushBack ctrls, CatalogControl
    Select Case StorageMethod
        Case osmBuiltIn:
            ArrayPushBack ctrls, TableControl
            ArrayPushBack ctrls, FieldControl
        Case osmCustom:
            ArrayPushBack ctrls, ParameterControl
            ArrayPushBack ctrls, ProcedureControl
        Case Else:
    End Select
    ControlsRequired = ctrls
End Function


