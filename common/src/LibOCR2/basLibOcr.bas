Attribute VB_Name = "basLibOcr"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' basLibOcr                                                 '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines an application programming interface  '
' (API) that clients can use to electroniaclly catalog      '
' paper documents in a database from scans, photos, etc.    '
' The API handles text parsing, data extraction, and        '
' database and file management. It searches for and reads   '
' source files, extracts identifying information and, if    '
' found, stores the information in a catalog (database) and '
' moves the files to a permanent location. The API only     '
' reads text files. Source files stored as images or other  '
' non-text formats must be converted to text by a client-   '
' supplied application (the ocr engine), such as an optical '
' character recognition (OCR) utility, which the API can    '
' control.                                                  '
'                                                           '
' The API can execute multiple user-defined processes, each '
' with its own search parameters, catalog database, and     '
' data extraction and storage rules. It is compatible with  '
' MS Access, SQL Server or any DAO compatible back-end      '
' database. It contains logic to ensure data integrity,     '
' employs client callbacks for realtime status monitoring,  '
' and is configurable for maximum performance and           '
' reliability.                                              '
'                                                           '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' StdLib, CallbackT, DAOConnectionT, RsActiveProcessesT,    '
' RsProcessInfoT, OcrConvertT                               '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the API has only been tested on Windows   '
' 10 and 11, with MS OFFICE 365 (64-bit) implementations.   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

'''''''''''''''''''''''
' Types and Constants '
'''''''''''''''''''''''

' Enumerates valid client callback types.
Public Enum OcrCallbackType
    octBegin = 1    ' Begin process,
    octSearch = 2   ' Search for source files,
    octFile = 3     ' Catalog files,
    octEnd = 4      ' End process,
    octMessage = 5  ' API message.
End Enum

' Enumerates valid catalog data storage methods.
Public Enum DataStorageMethod
    osmNone = 0     ' Indicates a bad or missing method,
    osmBuiltIn = 1  ' Use database recordset methods,
    osmCustom = 2   ' Call a user-defined procedure.
End Enum

Private Const kRegexNoSubmatch As Integer = -1  ' Indicates a regular expression has no submatch parameter.
Private Const kStrFileConversion As String = "File conversion failed"
Private Const kStrDuplicateIndex As String = "Duplicate source file"
Private Const kStrFileEmpty As String = "Converted file is empty"
Private Const kStrFileExists As String = "Catalog file already exists"
Private Const kStrNotFound As String = " not found"
Private Const kStrNoDataRules As String = "Process has no data storage rules"
Private Const kStrNoStorageMethod As String = "No storage method"
Private Const kStrNoSaveToMethod As String = "No save to method"
Private Const kStrNoRequiredData As String = "No required ocr data defined"
Private Const kStrPatternInvalid As String = "Bad or missing pattern parameters"
Private Const kStrStorageMethodFailed As String = "Storage method failed: "

Public Const kOcrErrDataNotFound = 10000        ' Required data not found in source file.
Public Const kOcrErrFileNoData = 10001          ' Source file has no text.
Public Const kOcrErrTimeout = 10002             ' Process timed out.
Public Const kOcrErrFileExists = 10003          ' Source file already exists in the catalog SaveTo Path.
Public Const kOcrErrFileConversion = 10004      ' A source file conversion error occurred.
Public Const kOcrErrProcessNoRules = 10005      ' Process has no data storage rules.
Public Const kOcrErrNoStorageMethod = 10006     ' Ocr data rule has no storage method.
Public Const kOcrErrNoRequiredData = 10007      ' Process has no required ocr data rule.
Public Const kOcrErrPatternInvalid = 10008      ' Bad or missing pattern match parameters.
Public Const kOcrErrStorageMethodFailed = 10009 ' Data storage method failed.

Public Const kOcrQueryActiveProcesses As String = "qryActiveProcesses"
Public Const kOcrQueryProcessInfo As String = "qryProcessInfo"
Public Const kOcrQueryProcessDataRules As String = "qryProcessDataRules"
Public Const kOcrTableCatalogDefs As String = "~tblCatalogDefs"
Public Const kOcrTablePatternDefs As String = "~tblPatternDefs"
Public Const kOcrTableProcessOcrData As String = "~tblProcessOcrData"
Public Const kOcrTableProcesses As String = "~tblProcesses"
Public Const kOcrTableSearchDefs As String = "~tblSearchDefs"

''''''''''''''''''''''''
' Module-Level Objects '
''''''''''''''''''''''''

Private boolInterrupt As Boolean    ' Flag indicating that the process has been interrupted.

''''''''''''''''''''
' Public Interface '
''''''''''''''''''''

Public Sub ocrReset()
    '
    ' Resets the program into a know state.
    '
    boolInterrupt = False
End Sub

Public Function ocrStart( _
    aOcrConvert As OcrConvertT, _
    Optional aCallback As CallbackT = Nothing, _
    Optional ByVal aOnError As VbMsgBoxResult = vbIgnore _
) As Long
    '
    ' Starts the program with the given arguments.
    '
    Dim processes As New RsActiveProcessesT
    
    processes.Open_
    While Not (processes.BOF Or processes.EOF Or boolInterrupt)
        ocrStart = ocrStart + ocrProcessStart(processes, aOcrConvert, aCallback, aOnError)
        processes.LastUpdate = Now()
        DoEvents
        processes.MoveNext
    Wend
    If IsSomething(processes) Then processes.Close_
    Set processes = Nothing
End Function

Public Sub ocrStop()
    '
    ' Stops the program.
    '
    boolInterrupt = True
End Sub

Public Function ocrStorageMethod( _
    ByVal aTableName As Variant, _
    ByVal aProcedureName As Variant _
) As DataStorageMethod
    '
    ' Returns the data storage method according to the given parameters.
    ' Returns osmNone if neither or both arguments have valid values.
    '
    If Nz(aTableName) <> "" And Nz(aProcedureName) = "" Then
        ocrStorageMethod = osmBuiltIn
    ElseIf Nz(aTableName) = "" And Nz(aProcedureName) <> "" Then
        ocrStorageMethod = osmCustom
    End If
End Function

'''''''''''''''''''''
' Private Interface '
'''''''''''''''''''''

Private Function ocrCatalogFile( _
    ByVal aSourceFile As String, _
    aProcessInfo As RsProcessInfoT, _
    aProcessRules As RsProcessDataRulesT, _
    aOcrConvert As OcrConvertT, _
    aWorkspace As DAO.Workspace, _
    aCatalog As Object, _
    Optional aCallback As CallbackT = Nothing, _
    Optional ByVal aOnError As VbMsgBoxResult = vbIgnore _
) As Boolean
    '
    ' Top-level file cataloging procedure. Converts a source
    ' file to text, extracts and stores any data and moves
    ' the file to a storage location.
    '
    Dim sourcePath As String, destinationPath As String, ocrdata() As Variant, ocrText As String, _
    intrans As Boolean, retry As Boolean, msg As String
    
    On Error GoTo Catch
    retry = False
    ' Generate the full source and destination file paths.
    sourcePath = PathBuild(aProcessInfo.SearchPath, aSourceFile)
    destinationPath = IIf(aProcessInfo.SaveToPath <> "", PathBuild(aProcessInfo.SaveToPath, aSourceFile), sourcePath)
    ' Make sure the destination file doesn't already exist.
    If (destinationPath <> sourcePath) And FileExists(destinationPath) Then _
    Err.Raise UsrErr(kOcrErrFileExists), CurrentProject.Name, kStrFileExists
Try:
    ' Generate source file ocr text using the OCR engine.
    ocrText = ocrFileConvert(aOcrConvert, sourcePath)
    aWorkspace.BeginTrans
    intrans = True
    ' Extract any ocr data and move the source file to the destination path.
    ocrdata = ocrStoreData(ocrText, destinationPath, aProcessInfo, aProcessRules, aCatalog)
    If sourcePath <> destinationPath Then FileMove sourcePath, destinationPath
    ' Only commit the transaction if the source file was successfully cataloged and moved.
    aWorkspace.CommitTrans
    intrans = False
    ocrCatalogFile = True
Finally:
    ocrClientCallBack aCallback, octFile, aSourceFile, Trim(Join(ocrdata, ",")), msg
    Exit Function

Catch:
    If intrans Then aWorkspace.Rollback
    intrans = False
    If Err.Number = kvbErrDbDuplicateIndex Then Err.Description = kStrDuplicateIndex    ' The default description is too verbose.
    Select Case ocrErrorHandler(aOnError, retry)
        Case vbCancel, vbAbort:
            Err.Raise Err.Number
        Case vbIgnore:
            msg = Err.Description
            Resume Finally
        Case vbRetry:
            Resume Try
    End Select
End Function

Private Function ocrClientCallBack( _
    aCallback As CallbackT, _
    ParamArray aArgs() As Variant _
) As Variant
    '
    ' Executes a client callback with the given arguments.
    '
    If Not aCallback Is Nothing Then ocrClientCallBack = aCallback.Exec(CVar(aArgs))
End Function

Private Function ocrErrorHandler( _
    ByVal aOnError As VbMsgBoxResult, _
    ByRef aRetry As Boolean _
) As VbMsgBoxResult
    '
    ' Limits retries to one and determines the type of error
    ' and action to take. Certain types of errors may override
    ' the specified OnError action when it would be inappropriate:
    '
    ' Duplicate indexes and errors converting, reading, or
    ' moving files are always ignored because a retry would
    ' return the same result,
    '
    ' Database errors dealing with permissions, write access,
    ' invalid table or field names, or invalid calls, as well
    ' as file path errors always cancel the current process
    ' because the error would recur with each database or file
    ' operation,
    '
    ' Untrapped errors are assumed to be unrecoverable
    ' database, VBA or system errors and always terminate the
    ' program and pass the error back to the client.
    '
    If aRetry Then
        ocrErrorHandler = vbCancel          ' If retry already attempted, cancel the current process, ...
    ElseIf aOnError = vbAbort Then
        ocrErrorHandler = vbAbort           ' ... if the OnError action is vbAbort, terminate the program, ...
    Else
        Select Case Err.Number              ' ... otherwise check the error code.
            Case kvbErrDbDuplicateIndex, UsrErr(kOcrErrFileExists) To UsrErr(kOcrErrFileConversion), _
            kvbErrFileNotFound, kvbErrFileAlreadyOpen, kvbErrFileAlreadyExists, UsrErr(kOcrErrFileExists), _
            kvbErrInputPastEndOfFile, kvbErrPermissionDenied:
                ocrErrorHandler = vbIgnore  ' Errors that shouldn't be retried.
            Case UsrErr(kOcrErrStorageMethodFailed), kvbErrBadFilenameOrNumber, kvbErrDbItemNotFound, UsrErr(kOcrErrProcessNoRules), _
            UsrErr(kOcrErrPatternInvalid), kvbErrCantRenameWithDifferentDrive, kvbErrPathNotFound, kvbErrPathFileAccessError, _
            kvbErrDbInvalidArgument To kvbErrDbFieldNotFound, kvbErrDbCantUpdateReadOnly To kvbErrDbObjectNoPermissions, _
            kvbErrDbTableQueryNotFound, kvbErrDbOdbcCallFailed, UsrErr(kOcrErrNoStorageMethod), UsrErr(kOcrErrNoRequiredData), _
            kvbErrDbFieldRequired To kvbErrDbValidationRuleViolation:
                ocrErrorHandler = vbCancel  ' Errors that cancel the current process.
            Case UsrErr(kOcrErrDataNotFound) To UsrErr(kOcrErrTimeout)
                aRetry = True
                ocrErrorHandler = aOnError  ' Errors that can be retried or ignored.
            Case Else:
                ocrErrorHandler = vbAbort   ' Db/VBA/system or other fatal errors.
        End Select
    End If
End Function

Private Function ocrExtractData( _
    aProcessRules As RsProcessDataRulesT, _
    ByVal aOcrText As String, _
    ByRef aData As VectorT _
) As Variant()
    '
    ' Appends aData with aProcessRules.StorageField and corresponding
    ' data extracted from aOcrText, if any, and returns a list of
    ' required data found.
    '
    If aProcessRules.Count > 0 Then
        Dim found() As Variant
        
        aProcessRules.MoveFirst
        While Not aProcessRules.EOF
            Dim ocrValue As Variant
            
            ocrValue = ocrGetValue(aOcrText, aProcessRules.Pattern, aProcessRules.Match, aProcessRules.Submatch, _
            aProcessRules.IgnoreCase, aProcessRules.Global_)
            If IsEmpty(ocrValue) And aProcessRules.Required Then _
            Err.Raise UsrErr(kOcrErrDataNotFound), CurrentProject.Name, aProcessRules.Name & kStrNotFound
            ' Only keep data where the rule's storage parameter/field isn't empty.
            If Not (IsEmpty(ocrValue) Or Nz(aProcessRules.StorageParameterField) = "") Then _
                aData.PushBack NewPairT(aProcessRules.StorageParameterField, ocrValue)
            If aProcessRules.Required Then ArrayPushBack found, ocrValue
            aProcessRules.MoveNext
        Wend
        ocrExtractData = found
    End If
End Function

Private Function ocrFileConvert( _
    aOcrConvert As OcrConvertT, _
    ByVal aFilePath As String, _
    Optional aCallback As CallbackT = Nothing, _
    Optional ByVal aOnError As VbMsgBoxResult = vbIgnore _
) As String
    '
    ' Returns any text generated by the ocr engine from the given file.
    '
    On Error Resume Next    ' Any ocr engine errors are converted to our kOcrErrFileConversion error.
    ocrFileConvert = aOcrConvert.Exec(aFilePath)
    On Error GoTo 0
    If Err.Number <> 0 Then Err.Raise UsrErr(kOcrErrFileConversion), CurrentProject.Name, Err.Description
    If ocrFileConvert = "" Then Err.Raise UsrErr(kOcrErrFileNoData), CurrentProject.Name, kStrFileEmpty
End Function

Private Function ocrGetCatalog( _
    aProcessInfo As RsProcessInfoT, _
    aConnection As DAOConnectionT, _
    ByRef aWorkspace As DAO.Workspace _
) As Object
    '
    ' Returns a catalog object based on the connection and process info
    ' settings. Processes that use the built-in storage methods return
    ' a DAO.Recordset, custom storage methods return an application object.
    ' If the connection is to an external MS Access database, then a new
    ' application object is created and its CurrentDb is set to the
    ' connection database so that the storage procedure can be called in
    ' that database. Otherwise we return the current application instance
    ' since the catalog database is either the current or an ODBC database,
    ' and in those cases storage procedures can only be called from the the
    ' current db.
    '
    Select Case ocrStorageMethod(aProcessInfo.SaveToTable, aProcessInfo.SaveToProcedure)
        Case osmBuiltIn:
            Set aWorkspace = DBEngine(0)
            Set ocrGetCatalog = aWorkspace.OpenDatabase(aConnection.Database, _
            IIf(aConnection.Name = "ODBC", False, Nothing), False, aConnection.connect). _
            OpenRecordset(aProcessInfo.SaveToTable, dbOpenDynaset, dbSeeChanges)
        Case osmCustom:
            If Not (aConnection.Database = CurrentDb.Name Or aConnection.Name = "ODBC") Then
                Set ocrGetCatalog = New Access.Application
                ocrGetCatalog.Visible = False
                Set aWorkspace = ocrGetCatalog.DBEngine(0)
                ocrGetCatalog.OpenCurrentDatabase aConnection.Database
            Else
                Set ocrGetCatalog = Application
                Set aWorkspace = ocrGetCatalog.DBEngine(0)
            End If
        Case osmNone:
            Err.Raise UsrErr(kOcrErrNoStorageMethod), CurrentProject.Name, kStrNoSaveToMethod
        Case Else:
    End Select
End Function

Private Function ocrGetValue( _
    ByVal aText As String, _
    ByVal aPattern As String, _
    ByVal aMatch As Integer, _
    ByVal aSubmatch As Integer, _
    ByVal aIgnoreCase As Boolean, _
    ByVal aGlobal As Boolean _
) As Variant
    '
    ' Returns data extracted from the given text, if any.
    '
    Dim regex As Object
    Dim matches As MatchCollection
    Dim m As Match
    
    On Error GoTo Catch
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = aPattern
    regex.IgnoreCase = aIgnoreCase
    regex.global = aGlobal
    If regex.Test(aText) Then
        Set matches = regex.Execute(aText)
        Set m = matches.item(aMatch)
        If aSubmatch > kRegexNoSubmatch Then
            ocrGetValue = m.SubMatches(aSubmatch)
        Else
            ocrGetValue = m.Value
        End If
    End If
    Set matches = Nothing
    Set regex = Nothing
    Exit Function
    
Catch:
    Err.Raise UsrErr(kOcrErrPatternInvalid), CurrentProject.Name, kStrPatternInvalid
End Function

Private Function ocrProcessFiles( _
    aProcessInfo As RsProcessInfoT, _
    aProcessRules As RsProcessDataRulesT, _
    aCatalog As Object, _
    aOcrConvert As OcrConvertT, _
    aWorkspace As DAO.Workspace, _
    Optional aCallback As CallbackT = Nothing, _
    Optional ByVal aOnError As VbMsgBoxResult = vbIgnore _
) As Long
    '
    ' Searches for source files and passes them on for cataloging.
    '
    Dim fTypes() As String, ftype As Variant, intrans As Boolean, msg As String, retry As Boolean
    fTypes = Split(aProcessInfo.FileTypes, ",")
    ' For each specified file type ...
    For Each ftype In fTypes
        Dim fpath As String, SourceFile As String
        
        ftype = "*." & ftype
        fpath = PathBuild(aProcessInfo.SearchPath, CStr(ftype))
        ocrClientCallBack aCallback, octSearch, aProcessInfo.SearchName, aProcessInfo.SearchPath, FilesCount(fpath)
        SourceFile = Dir(fpath)
        ' ... search the source path and ...
        While Not (Len(SourceFile) = 0 Or boolInterrupt)
            ' ... pass each file on for cataloging.
            If ocrCatalogFile(SourceFile, aProcessInfo, aProcessRules, aOcrConvert, aWorkspace, aCatalog, aCallback, aOnError) Then _
            ocrProcessFiles = ocrProcessFiles + 1
            DoEvents
            SourceFile = Dir
        Wend
        DoEvents
    Next
End Function

Private Function ocrProcessStart( _
    aProcess As RsActiveProcessesT, _
    aOcrConvert As OcrConvertT, _
    Optional aCallback As CallbackT = Nothing, _
    Optional ByVal aOnError As VbMsgBoxResult = vbIgnore _
) As Long
    '
    ' Executes a cataloging process.
    '
    Dim connection As DAOConnectionT, procinfo As New RsProcessInfoT, rules As New RsProcessDataRulesT, _
    catalog As Object, ws As DAO.Workspace, retry As Boolean
    
    On Error GoTo Catch
    procinfo.Open_ aProcess.ID
    rules.Open_ aProcess.ID
    If rules.Count = 0 Then Err.Raise UsrErr(kOcrErrProcessNoRules), CurrentProject.Name, kStrNoDataRules
    Set connection = NewDAOConnectionT(procinfo.connection)
Try:
    Set catalog = ocrGetCatalog(procinfo, connection, ws)
    If IsSomething(catalog) Then
        ocrClientCallBack aCallback, octBegin, procinfo.Name
        ocrProcessStart = ocrProcessFiles(procinfo, rules, catalog, aOcrConvert, ws, aCallback, aOnError)
        If TypeName(catalog) = "Application" Then
            If catalog.CurrentDb.Name <> CurrentDb.Name Then catalog.Quit
        Else
            catalog.Close
        End If
    End If
Finally:
    ocrClientCallBack aCallback, octEnd, procinfo.CatalogName, ocrProcessStart
    If IsSomething(procinfo) Then procinfo.Close_
    If IsSomething(rules) Then rules.Close_
    Set procinfo = Nothing
    Set rules = Nothing
    Set catalog = Nothing
    Set ws = Nothing
    Exit Function
Catch:
    Select Case ocrErrorHandler(aOnError, retry)
        Case vbAbort:
            Err.Raise Err.Number
        Case vbCancel, vbIgnore:
            ocrClientCallBack aCallback, octMessage, aProcess.Name, Err.Description
            Resume Finally
        Case vbRetry:
            Resume Try
    End Select
End Function

Private Function ocrStorageDelegate( _
    aCatalog As Object, _
    aProcessInfo As RsProcessInfoT, _
    aData As VectorT _
) As Long
    '
    ' Calls the storage delegate according to the
    ' specified parameters.
    '
    If TypeName(aCatalog) = "Application" Then
        Dim proc As String
        
        proc = FileBaseName(aCatalog.CurrentDb.Name) & "." & aProcessInfo.SaveToProcedure
        ocrStorageDelegate = ocrStoreCustom(aCatalog, proc, aData)
    Else
        ocrStorageDelegate = ocrStoreBuiltIn(aCatalog, aData)
    End If
    If Err.Number <> 0 Then Err.Raise Err.Number
End Function

Private Function ocrStoreBuiltIn( _
    aCatalog As DAO.Recordset, _
    aCatalogData As VectorT _
) As Long
    '
    ' Stores catalog data by inserting records
    ' into a DAO.Recordset. Data is a VectorT
    ' of PairTs: field,value.
    '
    Dim item As Variant
    
    aCatalog.AddNew
    On Error GoTo Catch
    While aCatalogData.Size > 0
        Dim p As PairT, fld As String, val As Variant
        
        Set p = aCatalogData.PopBack()
        fld = CStr(p.First)
        val = p.Second
        aCatalog(fld) = val
    Wend
    ocrStoreBuiltIn = Nz(aCatalog("ID"), 0) ' SQL Server doesn't return a row ID.
    aCatalog.Update
    Exit Function
    
Catch:
    ' Add some useful info to the error message.
    Err.Description = fld & ": " & Err.Description
    Err.Raise Err.Number
End Function

Private Function ocrStoreCustom( _
    aCatalog As Access.Application, _
    ByVal aProcedure As String, _
    aValues As VectorT _
) As Long
    '
    ' Calls a user-defined procedure to store catalog data.
    ' aValues is converted from a VectorT of PairTs to a
    ' an array with dimensions (n)(2), where n is the vector
    ' size.
    '
    Dim values() As Variant
    
    While aValues.Size() > 0
        ArrayPushBack values, aValues.PopBack.ToArray
    Wend
    On Error GoTo Catch
    ocrStoreCustom = aCatalog.Run(aProcedure, values)
    Exit Function
    
Catch:
    ' Change the error number to one of ours.
    Err.Raise UsrErr(kOcrErrStorageMethodFailed), CurrentProject.Name, kStrStorageMethodFailed & Err.Description
End Function

Private Function ocrStoreData( _
    ByVal aOcrText As String, _
    ByVal aDestinationPath As String, _
    aProcessInfo As RsProcessInfoT, _
    aProcessRules As RsProcessDataRulesT, _
    aCatalog As Object _
) As Variant()
    '
    ' Stores data extracted from an ocr text file in the catalog, and
    ' returns a list of required data found in the file.
    '
    Dim fileData As VectorT, recID As Long
    
    On Error GoTo Catch
    Set fileData = NewVectorT()
    fileData.PushBack NewPairT(aProcessInfo.SaveToParameterField, aDestinationPath)
    ocrStoreData = ocrExtractData(aProcessRules, aOcrText, fileData)
    recID = ocrStorageDelegate(aCatalog, aProcessInfo, fileData)
    If IsSomething(fileData) Then fileData.Clear
    Set fileData = Nothing
    Exit Function
    
Catch:
    ' ODBC connections can return kvbErrDbOdbcCallFailed for any number of reasons.
    ' We check if it's a recoverable error and reraise it as such.
    If Err.Number = kvbErrDbOdbcCallFailed Then
        Dim e As DAO.Error
        
        For Each e In DBEngine.Errors
            Dim done As Boolean
            
            done = True
            Select Case e.Number
                Case kvbErrDbOdbcKeyViolation:
                    Err.Number = kvbErrDbDuplicateIndex
                Case Else:
                    done = False
            End Select
            If done Then Exit For
        Next
    End If
    Err.Raise Err.Number
End Function

''''''''''''''''''''''''''''''
' basClassFactory Extensions '
''''''''''''''''''''''''''''''

Public Function NewOcrConvertT( _
    aProcedure As CallbackT, _
    ParamArray aArgs() As Variant _
) As OcrConvertT
    Dim obj As New OcrConvertT
    
    Set obj.Procedure = aProcedure
    obj.Params = CVar(aArgs)
    Set NewOcrConvertT = obj
    Set obj = Nothing
End Function

