VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataRulesControlsT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' DataRulesControlsT                                        '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a class that encapsulates the Data    '
' Rules list control on the Settings form and handles its   '
' shortcut menu events.                                     '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' basCommon, basShortcutMenu                                '
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

Private Const kParameterProcessId As String = "ProcessID"
Private Const kParameterRuleId As String = "RuleID"
Private Const kQueryClear As String = "qryRulesClear"
Private Const kQueryCopy As String = "qryRulesCopy"
Private Const kQueryDelete As String = "qryRulesDelete"
Private Const kQueryPaste As String = "qryRulesPaste"
Private Const kQueryUndoPaste As String = "qryRulesUndoPaste"
Private Const kTableClipboard As String = "~tblClipboard"
Private Const kSqlRulesProcessId As String = _
    "SELECT [ID], [Rule Name], [Storage ParameterField], [Required], [Pattern Name] " & _
    "FROM " & "qryMyProcessDataRules" & " " & _
    "WHERE [Process ID] = <ProcessId> " & _
    "ORDER BY [Required], [ID];"

'''''''''''''''''''
' Private Members '
'''''''''''''''''''
    
Private listRules As ListBoxExT ' Client Listbox control.
Private menuHistory As StackT   ' LIFO menuHistory of shortcut menu actions.
Private boolCanPaste As Boolean ' Flag indicating whether the paste action is enabled.
                                ' We use this to determine what is available after Cut
                                ' (can paste) and Delete (no paste).
                                
'''''''''''''''''''''''''''
' Public Member Accessors '
'''''''''''''''''''''''''''

Public Property Get list() As ListBoxExT
    Set list = listRules
End Property

Public Property Set list(aListBox As ListBoxExT)
    Set listRules = aListBox
End Property

Public Property Set Callback(aCallback As CallbackT)
    '
    ' Initialize/destroy the shortcut menu callback.
    '
    If Not aCallback Is Nothing Then
        RulesShortcutCreate aCallback
    Else
        RulesShortcutDestroy
    End If
End Property

''''''''''''''''''''
' Public Interface '
''''''''''''''''''''

Public Property Let Current(ByVal aCurrent As Boolean)
    If aCurrent Then ListCurrent
End Property

Public Sub Init( _
    aList As ListBoxExT _
)
    Set list = aList
    Set menuHistory = NewStackT()
End Sub

Public Sub ListClick()
    ShortcutCurrent
End Sub

Public Sub ListDblClick( _
    ByVal aProcessId As Long, _
    ByVal aStorageMethod As DataStorageMethod, _
    Optional ByVal aStorageTable As Variant _
)
    '
    ' Open the catalog settings editor form.
    '
    Dim rule_id As Long, list As ListBox, update As Boolean
    
    Set list = listRules.Control
    With list
        If .ItemsSelected.Count > 0 Then rule_id = .ItemData(.ItemsSelected(0))
    End With
    If FormOpenDialog(kOcrAdminFormOcrDataRules, dlgFind, "[ID]=" & Str(rule_id), aProcessId, aStorageMethod, aStorageTable) Then
        ListSave
        Current = True
    Else
        ListUndo
    End If
End Sub

Public Function ShortcutCallback( _
    aButtonTag As Variant _
)
    '
    ' Call the shortcut action procedure according to the
    ' button clicked.
    '
    Dim tag As RulesShortcutAction
    
    tag = aButtonTag(0) ' Call back args are variant arrays.
    Select Case tag
        Case odsCopy:
            ShortcutCopy
        Case odsCut:
            ShortcutCut
        Case odsDelete
            ShortcutDelete
        Case odsPaste:
            ShortcutPaste RuleProcessId()
        Case odsUndo:
            ShortcutUndo
        Case Else:
    End Select
    If tag <> odsCopy Then ListCurrent
End Function

'''''''''''''''''''''
' Private Interface '
'''''''''''''''''''''

Private Sub ListCurrent()
    '
    ' Display the current process' rules and update the
    ' shortcut menu status.
    '
    With listRules.Control
        .RowSource = ReplaceTags(kSqlRulesProcessId, "<ProcessId>", RuleProcessId())
        .Requery
    End With
    ShortcutCurrent
End Sub

Private Sub ListSave()
    FormRunUpdate kOcrTableProcessOcrData, "[Saved]", True, "[Saved]=False"
End Sub

Private Sub ListUndo()
    FormRunDelete kOcrTableProcessOcrData, "[Saved]=False"
End Sub

Private Function RuleProcessId() As Long
    '
    ' Returns the current process id displayed on the Settings form.
    '
    RuleProcessId = listRules.Control.Parent.ProcessId
End Function

Private Function ShortcutCanPaste() As Boolean
    '
    ' Returns TRUE if the paste action is available,
    ' else returns FALSE.
    '
    ShortcutCanPaste = boolCanPaste And (DCount("*", kTableClipboard) > 0)
End Function

Private Function ShortcutCanUndo() As Boolean
    '
    ' Returns TRUE if the undo action is available,
    ' else returns FALSE.
    '
    ShortcutCanUndo = Not menuHistory.IsEmpty
End Function

Private Sub ShortcutClear()
    '
    ' Clears the clipboard and menuHistory.
    '
    While menuHistory.Size > 0
        menuHistory.Pop
    Wend
    CurrentDb.QueryDefs(kQueryClear).Execute dbFailOnError
End Sub

Private Sub ShortcutCurrent()
    '
    ' Enable menu items according to what actions are available.
    '
    Dim mask As RulesShortCutMask
    
    If Not IsEmpty(listRules.Selected()) Then mask = rsmCopy Or rsmCut Or rsmDelete
    If ShortcutCanPaste() Then mask = mask Or rsmPaste
    If ShortcutCanUndo() Then mask = mask Or rsmUndo
    ShortcutMask mask
End Sub

Private Sub ShortcutMask( _
    ByVal aMask As RulesShortCutMask _
)
    '
    ' Apply the given display mask to the menu.
    '
    Application.CommandBars(kRulesShortcutName).Controls("Cut").Enabled = aMask And rsmCut
    Application.CommandBars(kRulesShortcutName).Controls("Copy").Enabled = aMask And rsmCopy
    Application.CommandBars(kRulesShortcutName).Controls("Paste").Enabled = aMask And rsmPaste
    Application.CommandBars(kRulesShortcutName).Controls("Delete").Enabled = aMask And rsmDelete
    Application.CommandBars(kRulesShortcutName).Controls("Undo").Enabled = aMask And rsmUndo
End Sub

'''''''''''''''''''''''''
' Shortcut Menu Actions '
'''''''''''''''''''''''''

Private Sub ShortcutCopy()
    '
    ' Copies the selected rules to the clipboard and
    ' pushes a copy action onto the menu menuHistory.
    '
    Dim qdf As DAO.QueryDef
    Dim item As Variant
    
    Set qdf = CurrentDb.QueryDefs(kQueryCopy)
    ShortcutClear
    For Each item In listRules.Control.ItemsSelected
        qdf.Parameters(kParameterRuleId) = CLng(listRules.Control.ItemData(item))
        qdf.Execute dbFailOnError
    Next
    menuHistory.Push NewPairT(odsCopy, RuleProcessId())
    boolCanPaste = True
    Set qdf = Nothing
End Sub

Private Function ShortcutCut( _
    Optional ByVal aDelete As Boolean = False _
) As Long
    '
    ' Copies the selected rules to the clipboard, deletes
    ' them from the current process and, if not a delete
    ' action, pushes a cut action onto the menuHistory.
    '
    Dim qdf As DAO.QueryDef
    Dim item As Variant
    
    ShortcutCut = RuleProcessId()
    Set qdf = CurrentDb.QueryDefs(kQueryCopy)
    ShortcutClear
    For Each item In listRules.Control.ItemsSelected
        qdf.Parameters(kParameterRuleId) = CLng(listRules.Control.ItemData(item))
        qdf.Execute dbFailOnError
    Next
    Set qdf = CurrentDb.QueryDefs(kQueryDelete)
    For Each item In listRules.Control.ItemsSelected
        qdf.Parameters(kParameterRuleId) = CLng(listRules.Control.ItemData(item))
        qdf.Execute dbFailOnError
    Next
    If Not aDelete Then
        boolCanPaste = True
        menuHistory.Push NewPairT(odsCut, ShortcutCut)
    End If
    Set qdf = Nothing
End Function

Private Sub ShortcutDelete()
    '
    ' Calls ShortcutCut but modifies the behavior after the call.
    '
    Dim pid As Long
    
    pid = ShortcutCut(True)
    boolCanPaste = False
    menuHistory.Push NewPairT(odsDelete, pid)
End Sub

Private Sub ShortcutPaste( _
    ByVal aProcessId As Long, _
    Optional ByVal aUndo As Boolean = False _
)
    '
    ' Adds the rules currently on the clipboard to the
    ' given process and, if this isn't an undo action,
    ' pushes a paste action onto the menuHistory.
    '
    Dim qdf As DAO.QueryDef
    Dim pid As Long
    
    Set qdf = CurrentDb.QueryDefs(kQueryPaste)
    qdf.Parameters(kParameterProcessId) = aProcessId
    qdf.Execute dbFailOnError
    If Not aUndo Then
        menuHistory.Push NewPairT(odsPaste, aProcessId)
    End If
    Set qdf = Nothing
End Sub

Private Sub ShortcutUndo()
    '
    ' Pops off and undoes each action on the menuHistory
    ' in LIFO order.
    '
    Dim qdf As DAO.QueryDef
    
    While Not menuHistory.IsEmpty
        Dim action As PairT
        
        Set action = menuHistory.Pop
        Select Case CInt(action.First)
            Case odsCopy:
            Case odsCut, odsDelete:
                ShortcutPaste CLng(action.Second), True
            Case odsPaste:
                Set qdf = CurrentDb.QueryDefs(kQueryUndoPaste)
                qdf.Parameters(kParameterProcessId) = CLng(action.Second)
                qdf.Execute dbFailOnError
            Case Else:
        End Select
    Wend
    Set qdf = Nothing
End Sub

