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
Attribute txtConnection.VB_VarHelpID = -1
Private WithEvents cmdConnectionBrowse As CommandButton
Attribute cmdConnectionBrowse.VB_VarHelpID = -1
Private WithEvents txtSaveToPath As TextBox
Attribute txtSaveToPath.VB_VarHelpID = -1
Private WithEvents cmdSaveToPathBrowse As CommandButton
Attribute cmdSaveToPathBrowse.VB_VarHelpID = -1
Private WithEvents optionMethod As OptionGroup
Attribute optionMethod.VB_VarHelpID = -1
Private WithEvents comboTable As ComboBox
Attribute comboTable.VB_VarHelpID = -1
Private txtProcedure As TextBox
Private WithEvents comboField As ComboBox
Attribute comboField.VB_VarHelpID = -1
Private txtParameter As TextBox
Private labMethod As Label      ' Data Rules storage method/parameter column header label.
Private boolVisible As Boolean

'''''''''''''''''''''''''''
' Public Member Accessors '
'''''''''''''''''''''''''''

Public Property Get CatalogControl() As Control
    Set CatalogControl = IIf(Not comboCatalogName Is Nothing, comboCatalogName, txtCatalogName)
End Property

Public Property Set CatalogControl( _
    aControl As Control _
)
    Set comboCatalogName = IIf(TypeName(aControl) = "ComboBox", aControl, Nothing)
    Set txtCatalogName = IIf(TypeName(aControl) = "TextBox", aControl, Nothing)
End Property

Public Property Get ConnectionBrowse() As CommandButton
    Set ConnectionBrowse = cmdConnectionBrowse
End Property

Public Property Set ConnectionBrowse( _
    aButton As CommandButton _
)
    Set cmdConnectionBrowse = aButton
End Property

Public Property Get ConnectionControl() As TextBox
    Set ConnectionControl = txtConnection
End Property

Public Property Set ConnectionControl( _
    aTextBox As TextBox _
)
    Set txtConnection = aTextBox
End Property

Public Property Get FieldControl() As ComboBox
    Set FieldControl = comboField
End Property

Public Property Set FieldControl( _
    aListBox As ComboBox _
)
    Set comboField = aListBox
End Property

Public Property Get MethodControl() As OptionGroup
    Set MethodControl = optionMethod
End Property

Public Property Set MethodControl( _
    aOptionGroup As OptionGroup _
)
    Set optionMethod = aOptionGroup
End Property

Public Property Get MethodLabel() As Label
    Set MethodLabel = labMethod
End Property

Public Property Set MethodLabel(aLabel As Label)
    Set labMethod = aLabel
End Property

Public Property Get ParameterControl() As TextBox
    Set ParameterControl = txtParameter
End Property

Public Property Set ParameterControl( _
    aTextBox As TextBox _
)
    Set txtParameter = aTextBox
End Property

Public Property Get ProcedureControl() As TextBox
    Set ProcedureControl = txtProcedure
End Property

Public Property Set ProcedureControl( _
    aTextBox As TextBox _
)
    Set txtProcedure = aTextBox
End Property

Public Property Get SaveToBrowse() As CommandButton
    Set SaveToBrowse = cmdSaveToPathBrowse
End Property

Public Property Set SaveToBrowse( _
    aButton As CommandButton _
)
    Set cmdSaveToPathBrowse = aButton
End Property

Public Property Get SaveToControl() As TextBox
    Set SaveToControl = txtSaveToPath
End Property

Public Property Set SaveToControl( _
    aTextBox As TextBox _
)
    Set txtSaveToPath = aTextBox
End Property

Public Property Get TableControl() As ComboBox
    Set TableControl = comboTable
End Property

Public Property Set TableControl( _
    aComboBox As ComboBox _
)
    Set comboTable = aComboBox
End Property

''''''''''''''''''''
' Public Interface '
''''''''''''''''''''

Public Property Let Current( _
    ByVal aCurrent As Boolean _
)
    If aCurrent Then StorageMethod = OcrStorageMethod(TableControl, ProcedureControl)
End Property

Public Property Let Enabled( _
    ByVal aEnabled As Boolean _
)
    ControlsEnable True
End Property

Public Property Get Invalid() As Control
    Set Invalid = ControlsInvalid()
End Property

Public Property Get RequiredControls() As Control()
    RequiredControls = ControlsRequired()
End Property

Public Property Get StorageMethod() As DataStorageMethod
    StorageMethod = MethodControl.Value
End Property

Public Property Let StorageMethod( _
    ByVal aStorageMethod As DataStorageMethod _
)
    If boolVisible Then ControlsShow aStorageMethod
    MethodControl.Value = aStorageMethod
    If Not MethodLabel Is Nothing Then
        If MethodLabel.Visible Then MethodLabel.Caption = "Storage " & IIf(StorageMethod = osmBuiltIn, "Field", "Parameter")
    End If
End Property

Public Property Get Visible() As Boolean
    Visible = boolVisible
End Property

Public Property Let Visible( _
    ByVal aVisible As Boolean _
)
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
    Optional aStorageLabel As Label = Nothing, _
    Optional ByVal aVisible As Boolean = False _
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
    Set MethodLabel = aStorageLabel
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

Public Sub Save()
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

Private Sub comboField_MouseDown( _
    Button As Integer, _
    Shift As Integer, _
    X As Single, _
    Y As Single _
)
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

Private Sub comboTable_MouseDown( _
    Button As Integer, _
    Shift As Integer, _
    X As Single, _
    Y As Single _
)
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

Private Sub txtCatalogName_BeforeUpdate( _
    Cancel As Integer _
)
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

Private Sub txtConnection_DblClick( _
    Cancel As Integer _
)
    On Error GoTo Catch
    cmdConnectionBrowse_Click
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub txtParameter_BeforeUpdate( _
    Cancel As Integer _
)
    On Error GoTo Catch
    Cancel = (txtParameter.Text = "")
    If Cancel Then MsgBox txtParameter.Name & " cannot be blank"
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub txtProcedure_BeforeUpdate( _
    Cancel As Integer _
)
    On Error GoTo Catch
    Cancel = (txtProcedure.Text = "")
    If Cancel Then MsgBox txtProcedure.Name & " cannot be blank"
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub txtSaveToPath_DblClick( _
    Cancel As Integer _
)
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
    ' and the Storage Table/Field and Procedure/Parameter controls
    ' if the aAll flag is set.
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
            TableControl.Visible = True
            ControlsSwitch TableControl, ProcedureControl, ParameterControl
            ProcedureControl.Visible = False
        Case osmCustom:
            ControlsEnable True
            ProcedureControl.Visible = True
            ControlsSwitch ProcedureControl, TableControl, FieldControl
            TableControl.Visible = False
        Case Else:
    End Select
    FieldControl.Visible = TableControl.Visible
    ParameterControl.Visible = ProcedureControl.Visible
End Sub

Private Sub ControlsSwitch( _
    aToControl As Variant, _
    ParamArray aFromControls() As Variant _
)
    '
    ' Switch the focus away from any From controls
    ' to the To control, if necessary.
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


