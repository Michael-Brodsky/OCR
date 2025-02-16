VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSearchFileTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' Form_frmSearchFileTypes                                   '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines types and objects that handle events  '
' and procedures for the frmSearchFileTypes form.           '
'                                                           '
' The form's user interface uses subforms. Any changes made '
' in the subform would immediately be reflected in its      '
' underlying table. To allow users to cancel these changes, '
' this module simulates transactions using temporary        '
' tables.                                                   '
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

Private Const kFormConvertible As String = "frmConvertible"             ' Convertible subform.
Private Const kFormReadableble As String = "frmReadable"                ' Readable subform.
Private Const kSqlDelete As String = "DELETE * FROM [<Table>];"         ' Delete query template.
Private Const kSqlCopy As String = "INSERT INTO [<Into Table>] " & _
    "SELECT * FROM [<From Table>];"                                     ' Copy query template.
Private Const kTableConvertible As String = "~tblTessConvertibleFiles"  ' Convertible subform's underlying table.
Private Const kTableConvertibleTemp As String = "~tmpConvertible"       ' Convertible subform's temporary table.
Private Const kTableReadable As String = "~tblTessReadableFiles"        ' Readable subform's underlying table.
Private Const kTableReadableTemp As String = "~tmpReadable"             ' Readable subform's temporary table.

''''''''''''''''''''''''
' Module-Level Objects '
''''''''''''''''''''''''

Private inTrans As Boolean  ' Flag indicating whether any changes are pending.

'''''''''''''''''''''''
' Form Event Handlers '
'''''''''''''''''''''''

Private Sub cmdCancel_Click()
    On Error GoTo Catch
    inTrans = False
    FormClose Me
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub cmdOK_Click()
    On Error GoTo Catch
    CommitTransaction
    Me.Visible = False
    Exit Sub
        
Catch:
    ErrMessage
End Sub

Private Sub Form_Load()
    BeginTransaction
    Me.cmdOK.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = ChangesSave()
End Sub

Public Sub TransactionCallback( _
    Optional aArgs As Variant _
)
    '
    ' Subform callback function to indicate changes are pending.
    '
    inTrans = True
End Sub

'''''''''''''''''''''
' Private Interface '
'''''''''''''''''''''

Private Sub BeginTransaction()
    '
    ' Make a copy of the permanent data in the temp tables
    ' and set the subforms' record source to work with the copies.
    '
    Dim cb As CallbackT
    
    Set cb = NewCallbackT("TransactionCallback", VbMethod, Form_frmSearchFileTypes)
    Set Me.frmConvertible.Form.Callback = cb
    Set Me.frmReadable.Form.Callback = cb
    Application.SetHiddenAttribute acTable, kTableConvertibleTemp, True
    Application.SetHiddenAttribute acTable, kTableReadableTemp, True
    CurrentDb.Execute ReplaceTags(kSqlDelete, "<Table>", kTableConvertibleTemp)
    CurrentDb.Execute ReplaceTags(kSqlDelete, "<Table>", kTableReadableTemp)
    CurrentDb.Execute ReplaceTags(kSqlCopy, "<Into Table>", kTableConvertibleTemp, "<From Table>", kTableConvertible)
    CurrentDb.Execute ReplaceTags(kSqlCopy, "<Into Table>", kTableReadableTemp, "<From Table>", kTableReadable)
    Me.frmConvertible.Form.Requery
    Me.frmReadable.Form.Requery
End Sub

Private Sub CommitTransaction()
    '
    ' Copy the temp data back to the permanent tables.
    '
    DoCmd.SetWarnings False
    CurrentDb.Execute ReplaceTags(kSqlDelete, "<Table>", kTableConvertible)
    CurrentDb.Execute ReplaceTags(kSqlDelete, "<Table>", kTableReadable)
    CurrentDb.Execute ReplaceTags(kSqlCopy, "<Into Table>", kTableConvertible, "<From Table>", kTableConvertibleTemp)
    CurrentDb.Execute ReplaceTags(kSqlCopy, "<Into Table>", kTableReadable, "<From Table>", kTableReadableTemp)
    DoCmd.SetWarnings True
    inTrans = False
End Sub

Private Function ChangesSave() As Boolean
    '
    ' If changes are pending, warn user before closing.
    '
    If inTrans Then
        Select Case MsgBox("Save changes?", vbYesNoCancel, FileBaseName(CurrentProject.Name))
            Case vbYes:
                CommitTransaction
                ChangesSave = True
                Me.Visible = False
            Case vbNo:
                
            Case vbCancel:
                ChangesSave = True
            Case Else:
        End Select
    End If
End Function
