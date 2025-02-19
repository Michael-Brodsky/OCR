VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessControlsT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' ProcessControlsT                                          '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a class that encapsulates and defines '
' the behaviour of the "Process" 'controls on the Settings  '
' form.                                                     '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' basCommon                                '
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

Private WithEvents comboProcessId As ComboBox
Attribute comboProcessId.VB_VarHelpID = -1
Private WithEvents txtProcessName As TextBox
Attribute txtProcessName.VB_VarHelpID = -1

'''''''''''''''''''''''''''
' Public Member Accessors '
'''''''''''''''''''''''''''

Public Property Get Process() As Control
    Set Process = IIf(comboProcessId.Visible, comboProcessId, IIf(txtProcessName.Visible, txtProcessName, Nothing))
End Property

Public Property Get ProcessId() As ComboBox
    Set ProcessId = comboProcessId
End Property

Public Property Set ProcessId( _
    aComboBox As ComboBox _
)
    Set comboProcessId = aComboBox
End Property

Public Property Get ProcessName() As TextBox
    Set ProcessName = txtProcessName
End Property

Public Property Set ProcessName( _
    aTextBox As TextBox _
)
    Set txtProcessName = aTextBox
End Property

''''''''''''''''''''
' Public Interface '
''''''''''''''''''''

Public Property Let Current( _
    ByVal aCurrent As Boolean _
)
    If aCurrent Then ProcessCurrent
End Property

Public Sub Init( _
    aProcessId As ComboBox, _
    aProcessName As TextBox _
)
    Const kSqlProcessIdRowSource As String = _
    "SELECT [ID], [Process Name] " & _
    "FROM " & "qryMyProcessInfo" & ";"

    aProcessId.RowSource = kSqlProcessIdRowSource
    Set ProcessId = aProcessId
    Set ProcessName = aProcessName
End Sub

Public Sub Requery()
    Process.Requery
End Sub

''''''''''''''''''
' Event Handlers '
''''''''''''''''''

Private Sub comboProcessId_AfterUpdate()
    On Error GoTo Catch
    ProcessForm.Recordset.FindFirst "[ID] = " & Str(comboProcessId)
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub comboProcessId_NotInList( _
    NewData As String, _
    Response As Integer _
)
    On Error GoTo Catch
    Response = ProcessNotInList(NewData)
    Exit Sub
    
Catch:
    Response = acDataErrContinue
    ErrMessage
End Sub

'''''''''''''''''''''
' Private Interface '
'''''''''''''''''''''

Private Sub ProcessCurrent()
    Dim frm As Form
    
    Set frm = ProcessForm()
    txtProcessName.Visible = frm.NewRecord
    comboProcessId.Visible = Not txtProcessName.Visible
    If comboProcessId.Visible Then
        Dim rst As DAO.Recordset
        
        Set rst = frm.Recordset
        If Not RecordsetEmpty(rst) Then comboProcessId.Value = rst.Fields("ID").Value
        Set rst = Nothing
    End If
    Set frm = Nothing
End Sub

Private Function ProcessForm() As Form
    Set ProcessForm = Process.Parent
End Function

Private Function ProcessNotInList( _
    ByVal aNewName As String _
) As Integer
    FormRunUpdate kOcrTableProcesses, "[Process Name]", StringQuote(aNewName), "[ID] = " & Str(comboProcessId)
    ProcessNotInList = acDataErrAdded
End Function

