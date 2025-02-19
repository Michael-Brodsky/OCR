VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SourceFileControlsT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' SourceFileControlsT                                       '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a class that encapsulates controls    '
' bound to "Source File" fields and defines common          '
' behaviour accross multiple forms.                         '
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

'''''''''''''''''''
' Private Members '
'''''''''''''''''''

Private WithEvents comboSearchName As ComboBox
Attribute comboSearchName.VB_VarHelpID = -1
Private WithEvents txtSearchName As TextBox
Attribute txtSearchName.VB_VarHelpID = -1
Private WithEvents txtSearchPath As TextBox
Attribute txtSearchPath.VB_VarHelpID = -1
Private WithEvents cmdSearchPathBrowse As CommandButton
Attribute cmdSearchPathBrowse.VB_VarHelpID = -1
Private WithEvents comboFileTypes As ComboBox
Attribute comboFileTypes.VB_VarHelpID = -1
Private WithEvents cmdFileTypesEdit As CommandButton
Attribute cmdFileTypesEdit.VB_VarHelpID = -1

'''''''''''''''''''''''''''
' Public Member Accessors '
'''''''''''''''''''''''''''

Public Property Get SearchFileTypes() As ComboBox
    Set SearchFileTypes = comboFileTypes
End Property

Public Property Set SearchFileTypes( _
    aComboBox As ComboBox _
)
    Set comboFileTypes = aComboBox
End Property

Public Property Get SearchFileTypesEdit() As CommandButton
    Set SearchFileTypesEdit = cmdFileTypesEdit
End Property

Public Property Set SearchFileTypesEdit( _
    aButton As CommandButton _
)
    Set cmdFileTypesEdit = aButton
End Property

Public Property Get SearchName() As Control
    Set SearchName = IIf(Not comboSearchName Is Nothing, comboSearchName, txtSearchName)
End Property

Public Property Set SearchName( _
    aControl As Control _
)
    Set comboSearchName = IIf(TypeName(aControl) = "ComboBox", aControl, Nothing)
    Set txtSearchName = IIf(TypeName(aControl) = "TextBox", aControl, Nothing)
End Property

Public Property Get SearchPath() As TextBox
    Set SearchPath = txtSearchPath
End Property

Public Property Set SearchPath( _
    aTextBox As TextBox _
)
    Set txtSearchPath = aTextBox
End Property

Public Property Get SearchPathBrowse() As CommandButton
    Set SearchPathBrowse = cmdSearchPathBrowse
End Property

Public Property Set SearchPathBrowse( _
    aButton As CommandButton _
)
    Set cmdSearchPathBrowse = aButton
End Property

''''''''''''''''''''
' Public Interface '
''''''''''''''''''''

Public Property Let Current( _
    ByVal aCurrent As Boolean _
)
    If aCurrent Then EnableControls
End Property

Public Property Get Invalid() As Control
    Set Invalid = FormInvalid(SearchName, SearchPath, SearchFileTypes)
End Property

Public Property Get RequiredControls() As Control()
    RequiredControls = ControlsRequired()
End Property

Public Sub Init( _
    aSearchName As Control, _
    aSearchPath As TextBox, _
    aSearchPathBrowse As CommandButton, _
    aFileTypes As ComboBox, _
    aFileTypesEdit As CommandButton _
)
    Set SearchName = aSearchName
    Set SearchPath = aSearchPath
    Set SearchPathBrowse = aSearchPathBrowse
    Set SearchFileTypes = aFileTypes
    Set SearchFileTypesEdit = aFileTypesEdit
End Sub

Public Sub Requery()
    SearchName.Requery
    SearchPath.Requery
    SearchPathBrowse.Requery
    SearchFileTypes.Requery
    SearchFileTypesEdit.Requery
End Sub

''''''''''''''''''
' Event Handlers '
''''''''''''''''''

Private Sub cmdSearchPathBrowse_Click()
    On Error GoTo Catch
    ControlPickPath txtSearchPath, pkFolder, kOcrAdminStrSelectSearchPath
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub cmdFileTypesEdit_Click()
    On Error GoTo Catch
    If FormOpenDialog(kOcrAdminFormFileTypes, dlgAddNew) Then comboFileTypes.Requery
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub comboFileTypes_DblClick( _
    Cancel As Integer _
)
    cmdFileTypesEdit_Click
End Sub

Private Sub comboSearchName_AfterUpdate()
    On Error GoTo Catch
    ControlsEnable (comboSearchName.Text <> "")
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub comboSearchName_BeforeUpdate( _
    Cancel As Integer _
)
    On Error GoTo Catch
    Cancel = (comboSearchName.Text = "")
    If Cancel Then MsgBox comboSearchName.Name & kStrCannotBeBlank
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub comboSearchName_Change()
    On Error GoTo Catch
    If (comboSearchName.Text <> "") Then ControlsEnable True
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub txtSearchName_BeforeUpdate( _
    Cancel As Integer _
)
    On Error GoTo Catch
    Cancel = (txtSearchName.Text = "")
    If Cancel Then MsgBox txtSearchName.Name & kStrCannotBeBlank
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub txtSearchName_Change()
    On Error GoTo Catch
    If (txtSearchName.Text <> "") Then ControlsEnable True
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub txtSearchPath_BeforeUpdate( _
    Cancel As Integer _
)
    On Error GoTo Catch
    Cancel = (txtSearchPath.Text = "")
    If Cancel Then MsgBox txtSearchPath.Name & kStrCannotBeBlank
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub txtSearchPath_Change()
    On Error GoTo Catch
    If (txtSearchPath.Text <> "") Then ControlsEnable True
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub txtSearchPath_DblClick( _
    Cancel As Integer _
)
    cmdSearchPathBrowse_Click
End Sub

'''''''''''''''''''''
' Private Interface '
'''''''''''''''''''''

Private Sub ControlsEnable( _
    ByVal aEnabled As Boolean _
)
    SearchPath.Enabled = aEnabled
    SearchPathBrowse.Enabled = aEnabled
    SearchFileTypes.Enabled = aEnabled
    SearchFileTypesEdit.Enabled = aEnabled
End Sub

Private Sub EnableControls()
    ControlsEnable (Not IsNull(SearchName))
End Sub

Private Function ControlsRequired() As Control()
    '
    ' Returns a list of controls that require valid entries.
    '
    Dim ctrls() As Control
    
    ArrayPushBack ctrls, SearchName
    ArrayPushBack ctrls, SearchPath
    ArrayPushBack ctrls, SearchFileTypes
    ControlsRequired = ctrls
End Function

