Attribute VB_Name = "basCommon"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' basCommon                                                 '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines objects and event handlers common to  '
' all application forms.                                    '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' LibAIOOcr                                                 '
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

' Enumerates valid values for a form open action (see FormOpenDialog() below).
Public Enum FormOpenAction
    dlgAddNew = 1
    dlgFind = 2
End Enum

''''''''''''''''''''''''
' Module-Level Objects '
''''''''''''''''''''''''

Private dbCurrent As DAO.Database   ' The current catalog database, if any.

'''''''''''''''''''''
' Common Procedures '
'''''''''''''''''''''

Public Function TableFieldsList( _
    ByVal aTableName As Variant _
) As Variant
    '
    ' Returns a semicolon-delimited list of field names
    ' in the given table.
    '
    If Not dbCurrent Is Nothing Then
        If Not IsNull(aTableName) Then
            Dim fld As DAO.Field
            
            For Each fld In dbCurrent.TableDefs(aTableName).Fields
                TableFieldsList = TableFieldsList & fld.Name & ";"
            Next
        End If
    End If
End Function

Public Function DatabaseTablesList() As Variant
    '
    ' Returns a semicolon-delimited list of available
    ' catalog table names in the current database.
    '
    If Not dbCurrent Is Nothing Then
        Dim tbl As Variant
        
        For Each tbl In TablesList((dbSystemObject Or dbHiddenObject), dbCurrent)
            If Left(tbl, 1) <> "~" Then DatabaseTablesList = DatabaseTablesList & tbl & ";"
        Next
    End If
End Function

Public Sub FormClose( _
    ByVal aForm As Form, _
    Optional ByVal aHide As Boolean = False _
)
    '
    ' Closes or hides the given form.
    '
    If aHide Then
        aForm.Visible = False
    Else
        DoCmd.Close acForm, aForm.Name
    End If
End Sub

Public Function FormConnectionChanged( _
    aConnection As Variant _
) As Boolean
    '
    ' Returns TRUE if the current connection (database)
    ' is different than the one given, else returns FALSE.
    '
    If dbCurrent Is Nothing Then
        FormConnectionChanged = True
    Else
        With NewDAOConnectionT(aConnection)
            FormConnectionChanged = (dbCurrent.Name <> .Database)
        End With
    End If
End Function

Public Sub FormConnectionClose()
    '
    ' Closes the current connection (database).
    '
    If Not dbCurrent Is Nothing Then dbCurrent.Close
    FormConnectionReset
End Sub

Private Sub FormConnectionOpen( _
    aConnection As Variant _
)
    '
    ' Closes the current and opens the given database connection.
    '
    FormConnectionClose
    On Error GoTo Catch
    With NewDAOConnectionT(aConnection)
        Set dbCurrent = OpenDatabase(.Database, IIf((.Name = "ODBC") Or (InStr(1, .Driver, "ODBC", _
        vbDatabaseCompare) > 0), False, Nothing), False, .connect)
    End With
    Exit Sub
    
Catch:
    If Err.Number = kvbErrDbOdbcCallFailed Then
        Dim e As DAO.Error
        
        For Each e In DBEngine.Errors
            Err.Description = e.Description
            Exit For
        Next
    End If
    Err.Raise Err.Number
End Sub

Public Sub FormConnectionReset()
    '
    ' Sets the current connection to nothing.
    '
    Set dbCurrent = Nothing
End Sub

Public Sub FormConnectionSwitch( _
    aConnection As Variant _
)
    '
    ' Switches the connection to the given database.
    '
    On Error GoTo Catch
    ShowSpinnyThingForReals True
    FormConnectionOpen aConnection
    ShowSpinnyThingForReals False
    DoEvents
    Exit Sub
    
Catch:
    ShowSpinnyThingForReals False
    Err.Raise Err.Number
End Sub

Public Sub ShowSpinnyThingForReals( _
    ByVal aShow As Boolean _
)
    '
    ' There's no guarantee the busy mousepointer
    ' will show, but I tried.
    '
    Const kCount As Integer = 500
    Dim real As Integer
    
    DoCmd.Hourglass aShow
    For real = 1 To kCount
        DoEvents
    Next
End Sub

Public Function FormControlChange( _
    aActiveControl As Control, _
    ParamArray aRequiredControls() As Variant _
) As Boolean
    '
    ' Returns TRUE if all required controls are valid and,
    ' the active control is either unspecified or valid,
    ' else returns FALSE.
    '
    Dim ctrls As Variant
    
    For Each ctrls In aRequiredControls
        Dim ctrl As Variant
        
        For Each ctrl In ctrls
            If Not aActiveControl Is Nothing Then
                If ctrl.Name = aActiveControl.Name Then GoTo Continue_For
            End If
            If IsNull(ctrl) Then Exit Function
Continue_For:
        Next
    Next
    If aActiveControl Is Nothing Then FormControlChange = True
    If Not FormControlChange Then FormControlChange = (aActiveControl.Text <> "")
End Function

Public Sub FormDeleteRecord( _
    aForm As Form _
)
    '
    ' Deletes the current record from the underlying recordset
    ' and moves to the previous record, if any.
    '
    Dim rst As DAO.Recordset, mark As String
    
    With aForm
        If Not aForm.NewRecord Then
            Set rst = .RecordsetClone
            rst.Bookmark = .Bookmark
            rst.MovePrevious
            If Not rst.BOF Then mark = rst.Bookmark
            .Recordset.Delete
            .Requery
            If mark <> "" Then .Bookmark = mark
        End If
    End With
    Set rst = Nothing
End Sub

Public Sub FormGotoNext( _
    aForm As Form _
)
    '
    ' Moves to the next record in the underlying recordset
    ' and wraps around at the first/last records.
    '
    With aForm
        If Not RecordsetEmpty(.Recordset) Then
            Dim last As Long
            
            last = .Recordset.RecordCount - 1
            .Recordset.Move IIf(.Recordset.AbsolutePosition <> last, 1, -last)
        End If
    End With
End Sub

Public Sub FormGotoPrevious( _
    aForm As Form _
)
    '
    ' Moves to the previous record in the underlying recordset
    ' and wraps around at the first/last records.
    '
    With aForm
        If Not RecordsetEmpty(.Recordset) Then
            Dim last As Long
            
            last = .Recordset.RecordCount - 1
            .Recordset.Move IIf(.Recordset.AbsolutePosition <> 0, -1, last)
        End If
    End With
End Sub

Public Function FormInvalid( _
    ParamArray aControls() As Variant _
) As Control
    '
    ' Returns the first control in the list that is blank,
    ' or Nothing if no controls are blank.
    '
    Dim ctrl As Variant
    
    For Each ctrl In aControls
        If Not ctrl Is Nothing Then
            If IsNull(ctrl) Then
                Set FormInvalid = ctrl
                Exit For
            End If
        End If
    Next
End Function

Public Sub FormLoad( _
    aForm As Form _
)
    '
    ' Initialize the form into a known state based on its OpenArgs, if any.
    '
    If Not IsNull(aForm.OpenArgs) Then
        Dim Args() As String, vals() As String
    
        On Error Resume Next        ' We don't care about errors because the form may not have been passed valid OpenArgs.
        Args = Split(aForm.OpenArgs, ",")
        Select Case CInt(Args(0))   ' args(0) is the "action", args(1) is the "parameter".
            Case dlgAddNew: ' Action: add new record.
                DoCmd.GotoRecord acDataForm, aForm.Name, acNewRec
                If UBound(Args) > 0 Then
                    vals = Split(Args(1), "=")          ' Parameter is a Key value pair in the form key=value.
                    aForm.Controls(vals(0)) = vals(1)   ' Assign "value" to form control "Key"
                End If
            Case dlgFind:   ' Action: find record. Parameter must be valid for Recordset.FindFirst.
                If Not RecordsetEmpty(aForm.Recordset) Then aForm.Recordset.FindFirst Args(1)
            Case Else:
        End Select
        Err.Clear
    End If
End Sub

Public Function FormOpenDialog( _
    ByVal aFormName As String, _
    ParamArray aOpenArgs() As Variant _
) As Boolean
    '
    ' Opens the given form in dialog view and returns TRUE if the form
    ' remained loaded after being dismissed, else returns FALSE.
    '
    On Error GoTo Catch
    DoCmd.OpenForm aFormName, , , , , acDialog, ArrayToCsv(CVar(aOpenArgs))
    FormOpenDialog = CurrentProject.AllForms(aFormName).IsLoaded
    
Catch:
    DoCmd.Close acForm, aFormName, acSaveNo
    If Err.Number <> 0 Then Err.Raise Err.Number
End Function

Public Sub FormRunDelete( _
    ByVal aTableName As String, _
    Optional ByVal aWhere As String _
)
    '
    ' Generates and executes a delete query from the given arguments.
    ' Generally used by forms to delete data outside their own recordset.
    '
    Const kSql As String = _
        "DELETE * " & _
        "FROM [<Table>] " & _
        "<Where>;"
    Dim strSql As String, strValue As String
    
    strSql = ReplaceTags(kSql, "<Table>", aTableName, "<Where>", IIf(aWhere <> "", "WHERE " & aWhere, ""))
    DoCmd.SetWarnings False
    CurrentDb.Execute strSql, dbFailOnError
    DoCmd.SetWarnings True
End Sub

Public Sub FormRunUpdate( _
    ByVal aTableName As String, _
    ByVal aFieldName As String, _
    ByVal aValue As Variant, _
    Optional ByVal aWhere As String _
)
    '
    ' Generates and executes an update query from the given arguments.
    ' Generally used by forms to update data outside their own recordset.
    '
    Const kSql As String = _
        "UPDATE [<Table>] " & _
        "SET <Field> = <Value> " & _
        "<Where>;"
    Dim strSql As String, strValue As String
    
    strSql = ReplaceTags(kSql, "<Table>", aTableName, "<Field>", aFieldName, _
    "<Value>", aValue, "<Where>", IIf(aWhere <> "", "WHERE " & aWhere, ""))
    DoCmd.SetWarnings False
    CurrentDb.Execute strSql, dbFailOnError
    DoCmd.SetWarnings True
End Sub

''''''''''''''''''''''''''''''
' Control Class Constructors '
''''''''''''''''''''''''''''''

Public Function NewCatalogControlsT( _
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
) As CatalogControlsT
    Dim ctrls As New CatalogControlsT
    
    ctrls.Init aCatalog, aConnection, aSaveToPath, aMethod, aBrowseConnection, aBrowseSaveTo, _
    aTable, aField, aProcedure, aParameter, aVisible
    Set NewCatalogControlsT = ctrls
    Set ctrls = Nothing
End Function

Public Function NewConversionControlsT( _
    ParamArray aConversionControls() As Variant _
) As ConversionControlsT
    Dim obj As New ConversionControlsT
    
    obj.Init CVar(aConversionControls)
    Set NewConversionControlsT = obj
    Set obj = Nothing
End Function

Public Function NewDataRulesControlsT( _
    aControl As ListBoxExT _
) As DataRulesControlsT
    Dim obj As New DataRulesControlsT
    
    obj.Init aControl
    Set NewDataRulesControlsT = obj
    Set obj = Nothing
End Function

Public Function NewProcessControlsT( _
    aProcessId As ComboBox, _
    aProcessName As TextBox _
) As ProcessControlsT
    Dim obj As New ProcessControlsT
    
    obj.Init aProcessId, aProcessName
    Set NewProcessControlsT = obj
    Set obj = Nothing
End Function

Public Function NewPropertyControlT( _
    aControl As Control, _
    ByVal aPropertyName As String, _
    Optional ByVal aValueDefault As Variant, _
    Optional ByVal aDisplayDefault As Variant, _
    Optional ByVal aDataType As Integer = kdbNone, _
    Optional ByVal aRequired As Boolean = False _
) As PropertyControlT
    Dim obj As New PropertyControlT
    
    obj.Init aControl, aPropertyName, aValueDefault, aDisplayDefault, aDataType, aRequired
    Set NewPropertyControlT = obj
    Set obj = Nothing
End Function

Public Function NewSourceFileControlsT( _
    aSearchName As Control, _
    aSearchPath As TextBox, _
    aSeachPathBrowse As CommandButton, _
    aFileTypes As ComboBox, _
    aFileTypesEdit As CommandButton _
) As SourceFileControlsT
    Dim ctrls As New SourceFileControlsT
    
    ctrls.Init aSearchName, aSearchPath, aSeachPathBrowse, aFileTypes, aFileTypesEdit
    Set NewSourceFileControlsT = ctrls
    Set ctrls = Nothing
End Function

