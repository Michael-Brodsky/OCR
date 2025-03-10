Attribute VB_Name = "basLibVBA"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' basLibVBA                                                 '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a library of generic, reusable types  '
' and objects that provided solutions to commonly           '
' encountered programming tasks in a single line of code.   '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' basLibWin, basLibArray, basLibNumeric                     '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the library has only been tested with     '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database
Option Explicit

'''''''''''''''''''''''
' Types and Constants '
'''''''''''''''''''''''

' 2-tuple storage type (see also PairT).
Public Type Pair
    First As Variant
    Second As Variant
End Type

' Enumerates valid values for the ObjectType argument in
' procedures that take it.
Public Enum ObjectType
    otForms = 0
    otReports = 1
    otMacros = 2
    otModules = 3
End Enum

' Enumerates valid values for the PathType argument in
' procedures that take it.
Public Enum PathType
    ptNone = 0
    ptDrive = 1
    ptFolder = 2
    ptFile = 3
End Enum

' Enumerates valid values for the PickerType argument in
' procedures that take it.
Public Enum PickerType
    pkFile = 1
    pkFolder = 2
End Enum

' Enumerates valid values for the RecordsetOperation argument in
' procedures that take it.
Public Enum RecordsetOperation
    roAddnew = 0
    roEdit = 1
    roDelete = 2
End Enum

''''''''''''''''''''''''''''''
' VBA and Access Error Codes '
''''''''''''''''''''''''''''''

Public Const kvbErrInvalidProcedureCall As Long = 5
Public Const kvbErrOutOfMemory As Long = 7
Public Const kvbErrSubscriptOutOfRange As Long = 9
Public Const kvbErrTypeMismatch As Long = 13
Public Const kvbErrUserInterruptOccurred As Long = 18
Public Const kvbErrBadFilenameOrNumber As Long = 52
Public Const kvbErrFileNotFound As Long = 53
Public Const kvbErrBadFileMode As Long = 54
Public Const kvbErrFileAlreadyOpen As Long = 55
Public Const kvbErrDeviceIOError As Long = 57
Public Const kvbErrFileAlreadyExists As Long = 58
Public Const kvbErrBadRecordLength As Long = 59
Public Const kvbErrDiskFull As Long = 61
Public Const kvbErrInputPastEndOfFile As Long = 62
Public Const kvbErrBadRecordNumber As Long = 63
Public Const kvbErrTooManyFiles As Long = 67
Public Const kvbErrDeviceUnavailable As Long = 68
Public Const kvbErrPermissionDenied As Long = 70
Public Const kvbErrDiskNotReady As Long = 71
Public Const kvbErrCantRenameWithDifferentDrive As Long = 74
Public Const kvbErrPathFileAccessError As Long = 75
Public Const kvbErrPathNotFound As Long = 76
Public Const kvbErrObjectNotFound As Long = 91
Public Const kvbErrInvalidFileFormat As Long = 321
Public Const kvbErrInvalidPropertyValue As Long = 380
Public Const kvbErrInvalidPropertyArrayIndex As Long = 381
Public Const kvbErrInvalidPicture As Long = 481
Public Const kvbErrPrinterError As Long = 482
Public Const kvbErrPrinterUnsupportedProperty As Long = 483
Public Const kvbErrPrinterSysInfo As Long = 484
Public Const kvbErrInvalidPictureType As Long = 485
Public Const kvbErrPrinterInvalidImageType As Long = 486
Public Const kvbErrReplacementsTooLong As Long = 746
Public Const kvbErrDbCantSaveRecord As Long = 2169
Public Const kvbErrDbSyntaxError As Long = 2433
Public Const kvbErrDbApplicationDefined As Long = 2467
Public Const kvbErrDbOdbcKeyViolation As Long = 2627
Public Const kvbErrDbInvalidArgument As Long = 3001
Public Const kvbErrDbInvalidName As Long = 3005
Public Const kvbErrDbExclusivelyLocked As Long = 3006
Public Const kvbErrDbCantOpenLibrary As Long = 3007
Public Const kvbErrDbTableExclusivelyOpen As Long = 3008
Public Const kvbErrDbTableInUse As Long = 3009
Public Const kvbErrDbTableExists As Long = 3010
Public Const kvbErrDbObjectNotFound As Long = 3011
Public Const kvbErrDbObjectExists As Long = 3012
Public Const kvbErrDbCantOpenMoreTables As Long = 3014
Public Const kvbErrDbNotAnIndex As Long = 3015
Public Const kvbErrDbFieldWontFit As Long = 3016
Public Const kvbErrDbFieldSizeTooLong As Long = 3017
Public Const kvbErrDbFieldNotFound As Long = 3018
Public Const kvbErrDbNoCurrentIndex As Long = 3019
Public Const kvbErrDbUpdateWithoutEdit As Long = 3020
Public Const kvbErrDbNoCurrentRecord As Long = 3021
Public Const kvbErrDbDuplicateIndex As Long = 3022
Public Const kvbErrDbEditAlreadyUsed As Long = 3023
Public Const kvbErrDbFileNotFound As Long = 3024
Public Const kvbErrDbCantUpdateReadOnly As Long = 3027
Public Const kvbErrDbObjectNoPermissions As Long = 3033
Public Const kvbErrDbPermissionDenied As Long = 3051
Public Const kvbErrDbInvalidFileName As Long = 3055
Public Const kvbErrDbIndexNullValue As Long = 3058
Public Const kvbErrDbOperationCancelled As Long = 3059
Public Const kvbErrDbQueryExpression As Long = 3075
Public Const kvbErrDbCriteriaExpression As Long = 3076
Public Const kvbErrDbExpression As Long = 3077
Public Const kvbErrDbTableQueryNotFound As Long = 3078
Public Const kvbErrDbOdbcCallFailed As Long = 3146
Public Const kvbErrDbOperationNotSupported As Long = 3251
Public Const kvbErrDbItemNotFound As Long = 3265
Public Const kvbErrDbPropertyNotFound As Long = 3270
Public Const kvbErrDbInvalidPropertyValue As Long = 3271
Public Const kvbErrDbFieldRequired As Long = 3314
Public Const kvbErrDbFieldZeroLength As Long = 3315
Public Const kvbErrDbTableLevelValidation As Long = 3316
Public Const kvbErrDbValidationRuleViolation As Long = 3317
Public Const kvbErrDbObjectInvalid As Long = 3420
Public Const kvbErrSavingToFile As Long = 31036
Public Const kvbErrLoadingFromFile As Long = 31037

'''''''''''''''''''''''''''
' Miscellaneous Constants '
'''''''''''''''''''''''''''

Public Const kvbInvalid As Long = -1                    ' Indicates an invalid value.
Public Const kdbNone As Integer = -1                    ' Indicates no value.
Public Const kstrAllFiles As String = "All Files,*.*"   ' Default value for certain optional procedure parameters.
Public Const kstrSelectFile As String = "Select File"   ' Default value for certain optional procedure parameters.
Public Const kstrSelectPath As String = "Select Path"   ' Default value for certain optional procedure parameters.

'''''''''''''''''''''
' Library Functions '
'''''''''''''''''''''

Public Function ControlPickPath( _
    aControl As Control, _
    ByVal aPickerType As PickerType, _
    Optional ByVal aTitle As String = kstrSelectPath, _
    Optional ByVal aFilters As String = kstrAllFiles, _
    Optional ByVal aInitFolder As String = "" _
) As Boolean
    '
    ' Sets the value of a control with the path, if any, picked
    ' from a standard file dialog of the specified picker type,
    ' initialized with the given title, filters and initial folder.
    ' Returns TRUE if anything was picked, else returns FALSE.
    '

    Dim path As String
    
    path = PathPicker(aPickerType, aTitle, aFilters, aInitFolder)
    If path <> "" Then
        aControl = path
        ControlPickPath = True
    End If
End Function

Public Function DatabaseGet( _
    Optional ByVal aPath As String = "", _
    Optional ByVal aOptions As Variant, _
    Optional ByVal aReadOnly As Boolean = False, _
    Optional ByVal aConnect As String = "" _
) As DAO.Database
    '
    ' Opens and returns a database object using the given arguments.
    ' If path is omitted, then returns the current database.
    '
    If aPath = "" Then
        Set DatabaseGet = CurrentDb
    Else
        Set DatabaseGet = OpenDatabase(aPath, IIf(Nz(aOptions) <> "", aOptions, Nothing), aReadOnly, aConnect)
    End If
End Function
    
Public Function DllErrDescription( _
    ByVal aErrno As Long _
) As String
    '
    ' Returns a human-readable description for certain Windows
    ' API error codes.
    '
    Select Case aErrno
        Case 0:
        
        Case SE_ERR_ACCESSDENIED:
            DllErrDescription = kwinErrDescAccessDenied
        Case SE_ERR_ASSOCINCOMPLETE:
            DllErrDescription = kwinErrDescAssocIncomplete
        Case SE_ERR_BAD_FORMAT:
            DllErrDescription = kwinErrDescBadFormat
        Case SE_ERR_DDEBUSY:
            DllErrDescription = kwinErrDescDdeBusy
        Case SE_ERR_DDEFAIL:
            DllErrDescription = kwinErrDescDdeFail
        Case SE_ERR_DDETIMEOUT:
            DllErrDescription = kwinErrDescDdeTimeout
        Case SE_ERR_DLLNOTFOUND:
            DllErrDescription = kwinErrDescDllNotFound
        Case SE_ERR_FILE_NOT_FOUND:
            DllErrDescription = kwinErrDescFileNotFound
        Case SE_ERR_NOASSOC:
            DllErrDescription = kwinErrDescNoAssoc
        Case SE_ERR_OOM:
            DllErrDescription = kwinErrDescOutOfMemory
        Case SE_ERR_PATH_NOT_FOUND:
            DllErrDescription = kwinErrDescPathNotFound
        Case SE_ERR_SHARE:
            DllErrDescription = kwinErrDescShare
        Case Else:
            DllErrDescription = "An unknown error occured [" & CStr(aErrno) & "]"
    End Select
End Function

Public Sub ErrMessage()
    '
    ' Displays an error dialog in a standard format.
    '
    MsgBox "[" & CStr(Err.Number) & "] " & Err.Description, , Err.Source
End Sub

Public Function ExecutablePath( _
    ByVal aName As String _
) As String
    '
    ' Returns the full path to the named Windows executable file.
    '
    ExecutablePath = CreateObject("WScript.Shell") _
    .RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & aName & "\")
End Function

Public Function FieldExists( _
    aTableDef As DAO.TableDef, _
    ByVal aFieldName As String _
) As Boolean
    '
    ' Returns TRUE if the named field exists in the
    ' specified TableDef, else returns FALSE.
    '
    On Error Resume Next
    FieldExists = IsObject(aTableDef.Fields(aFieldName))
    Err.Clear
End Function

Public Function FileBaseName( _
    ByVal aFilePath As String _
) As String
    '
    ' Returns the file base name from the given file path.
    '
    FileBaseName = CreateObject("Scripting.FileSystemObject").GetBaseName(aFilePath)
End Function

Public Function FileEmpty( _
    ByVal aFilePath As String _
) As Boolean
    '
    ' Returns TRUE if the given file is empty, else returns FALSE.
    '
    FileEmpty = (FileLen(aFilePath) = 0)
End Function

Public Function FileDrive( _
    ByVal aPath As String _
) As String
    '
    ' Returns the drive name from the given path.
    '
    FileDrive = CreateObject("Scripting.FileSystemObject").GetDriveName(aPath)
End Function

Public Function FileExists( _
    ByVal aFilePath As String _
) As Boolean
    '
    ' Returns TRUE if the specified file exists, else returns FALSE.
    '
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(aFilePath)
End Function

Public Function FileExtension( _
    ByVal aFilePath As String _
) As String
    '
    ' Returns the file extension component from the given path.
    '
    FileExtension = CreateObject("Scripting.FileSystemObject").GetExtensionName(aFilePath)
End Function

Public Function FileFolder( _
    ByVal aFilePath As String _
) As String
    '
    ' Returns the file folder component from the given path.
    '
    FileFolder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(aFilePath)
End Function

Public Function FileLocked( _
    ByVal aFile As String _
) As Boolean
    '
    ' Returns TRUE if a file is locked (in use), else returns FALSE.
    '
    Dim fileno As Integer
    
    fileno = FreeFile()
    On Error GoTo Catch
    Open aFile For Binary Access Read Write Lock Read Write As fileno
    Close fileno
    
Finally:
    Exit Function
    
Catch:
    FileLocked = True
    Resume Finally
End Function

Public Function FileLastModified( _
    ByVal aFile As String _
) As Date
    '
    ' Returns a file's last modified date-time.
    '
    FileLastModified = CreateObject("Scripting.FileSystemObject").GetFile(aFile).DateLastModified
End Function

Public Sub FileMove( _
    ByVal aSource As String, _
    ByVal aDestination As String _
)
    '
    ' Moves a source file to the specified destination.
    '
    CreateObject("Scripting.FileSystemObject").MoveFile Source:=aSource, Destination:=aDestination
End Sub

Public Function FileName( _
    ByVal aFilePath As String _
) As String
    '
    ' Returns the file name component from the given path.
    '
    FileName = CreateObject("Scripting.FileSystemObject").GetFileName(aFilePath)
End Function

Public Function FileOpen( _
    ByVal aFile As String, _
    Optional ByVal aShowCmd As VbAppWinStyle = vbNormalFocus, _
    Optional ByVal aHwnd As LongPtr = 0 _
) As LongPtr
    '
    ' Opens a file and returns the result code (hInstance) of
    ' the associated application that opened the file.
    '
    FileOpen = ShellExecute(aHwnd, "open", aFile, vbNullString, vbNullString, aShowCmd)
End Function

Public Function FilePicker( _
    Optional ByVal aTitle As String = kstrSelectFile, _
    Optional ByVal aFilters As String = kstrAllFiles, _
    Optional ByVal aInitFolder As String = "" _
) As String
    '
    ' Opens a standard file picker dialog with the given title,
    ' filters and initial folder, and returns the selected file,
    ' if any. Filters must have the standard file dialog filter
    ' format: "Description1,Filter1[;Description2,Filter2...;DescriptionN,FilterN]"
    '
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .title = aTitle
        .InitialFileName = aInitFolder
        If Not IsMissing(aFilters) Then
            If IsArray(aFilters) Then
                Dim fltrs() As String, fltr As Variant
                Dim Description As String
                
                .Filters.Clear
                fltrs = Split(aFilters, ";")
                If (Not Not fltrs) <> 0 Then
                    For Each fltr In fltrs
                        Dim info() As String
                        
                        info = Split(fltr, ",")
                        If (Not Not fltrs) <> 0 Then .Filters.Add info(0), info(1)
                    Next
                End If
            End If
        End If
        If .show Then FilePicker = .SelectedItems(1)
    End With
End Function

Public Function FilesCount( _
    ByVal aFilePath As String _
) As Long
    '
    ' Returns the number of files found in the given path.
    ' The path can contain wildcards as the last element to
    ' count only certain file extensions. If no extensions
    ' are specified, then all files are counted.
    '
    If FileName(aFilePath) = "" Then aFilePath = PathBuild(aFilePath, "*.*")
    If Dir(aFilePath) <> "" Then
        Do
            FilesCount = FilesCount + 1
        Loop While Dir <> ""
    End If
End Function

Public Function FilesRecordset( _
    ByVal aFolder As String, _
    ByVal aFileTypes As String, _
    ByVal aField As String, _
    aFiles As DAO.Recordset _
) As Long
    '
    ' Appends a recordset with specific files found in a specific folder.
    ' File names are stored in the given recordset field. Returns the number
    ' of files found. Multiple file extensions can be specified in a comma-
    ' separated list, e.g. "doc,docx,docm,...".
    '
    Dim sourceFile As String
    Dim FileTypes() As String
    Dim ftype As Variant
    
    FileTypes = Split(aFileTypes, ",")
    For Each ftype In FileTypes
        Dim path As String
        
        ftype = "*." & ftype
        path = aFolder & CStr(ftype)
        sourceFile = Dir(path)
        While Len(sourceFile) > 0
            RecordsetAddNew aFiles, aField, PathBuild(aFolder, sourceFile)
            sourceFile = Dir
        Wend
    Next
    FilesRecordset = RecordsetCount(aFiles)
End Function

Public Function FmtSigFigs(aNumber As Double, aSigFigs As Integer) As String
    '
    ' Returns a number formatted with the specified number of
    ' siginificant figures, upto 15.
    '
    FmtSigFigs = Format(aNumber, "0" & IIf((Int(aNumber) <> aNumber), "." & _
    String(Constrain(aSigFigs, 0, 15), "0"), ""))
End Function

Public Sub FolderCreate( _
    ByVal aPath As String _
)
    '
    ' Creates a folder in the given path.
    '
    CreateObject("Scripting.FileSystemObject").CreateFolder aPath
End Sub

Public Function FolderExists( _
    ByVal aPath As String _
) As Boolean
    '
    ' Returns TRUE if a folder in the given path exists,
    ' else returns FALSE.
    '
    FolderExists = CreateObject("Scripting.FileSystemObject").FolderExists(aPath)
End Function

Public Function FolderPicker( _
    Optional ByVal aTitle As String = "Select Folder", _
    Optional ByVal aInitFolder As String = "" _
) As String
    '
    ' Opens a standard folder picker dialog and returns the selected folder, if any.
    '
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .title = aTitle
        .InitialFileName = aInitFolder
        If .show Then FolderPicker = .SelectedItems(1)
    End With
End Function

Public Function FormatParam( _
    ParamArray aArgs() As Variant _
) As String
    '
    ' Returns the given arguments as a tab-delimited string.
    ' Arguments can be of any type, or array of type,
    ' convertible to a string.
    '
    Dim Arg As Variant
    
    If Not IsMissing(aArgs) Then
        Dim i As Integer
    
        Arg = aArgs(LBound(aArgs))
        GoSub Unpack
        For i = LBound(aArgs) + 1 To UBound(aArgs)
            Arg = aArgs(i)
            GoSub Unpack
        Next
    End If
    FormatParam = Trim(FormatParam)
    Exit Function
    
Unpack:
    If IsArray(Arg) Then
        Dim j As Integer
        
        For j = LBound(Arg) To UBound(Arg)
            FormatParam = FormatParam & StringTab(FormatParam) & Arg(j)
        Next
    Else
        FormatParam = FormatParam & StringTab(FormatParam) & aArgs(i)
    End If
    Return
End Function

Public Function IsSomething( _
    Optional aArg As Variant _
) As Boolean
    '
    ' Like "Not Is Nothing", but with less typing
    ' and more checks.
    '
    On Error Resume Next
    If IsObject(aArg) Then
        If aArg Is Nothing Then Exit Function
    End If
    IsSomething = Not (IsMissing(aArg) Or IsEmpty(aArg) Or IsNull(aArg))
End Function

Public Function IsValue( _
    aValue As Variant _
) As Boolean
    '
    ' Returns TRUE if the given value is not nothing, null
    ' empty or a zero-length string, else returns FALSE.
    '
    IsValue = IsSomething(aValue) And Nz(aValue) <> ""
End Function

Public Sub MessageLog( _
    ByVal aFilePath As String, _
    ByVal aMessage As String _
)
    '
    ' Appends a message string to the specified file prepended
    ' with a timestamp.
    '
    Dim fileno As Integer
    Dim timeStamp As Date
    
    fileno = FreeFile()
    timeStamp = Now()
    Open aFilePath For Append As #fileno
    Print #fileno, timeStamp & vbTab & aMessage
    Close #fileno
End Sub

Public Function ObjectExists( _
    aApplication As Access.Application, _
    ByVal aObjName As String, _
    ByVal aObjType As ObjectType _
) As Boolean
    '
    ' Returns TRUE if a named object of the given type exists
    ' in the application's current project, else returns FALSE
    ' (see ObjectType enum above).
    '
    On Error GoTo Catch   ' IsObject() throws an error if the object doesn't exist.
    Select Case aObjType
        Case otForms:
            ObjectExists = IsObject(aApplication.CurrentProject.AllForms(aObjName))
        Case otReports:
            ObjectExists = IsObject(aApplication.CurrentProject.AllReports(aObjName))
        Case otMacros:
            ObjectExists = IsObject(aApplication.CurrentProject.AllMacros(aObjName))
        Case otModules:
            ObjectExists = IsObject(aApplication.CurrentProject.AllModules(aObjName))
        Case Else:
    End Select
    
Finally:
    Exit Function
    
Catch:
    If Err.Number = kvbErrDbApplicationDefined Then Resume Finally
    Err.Raise Err.Number
End Function

Public Function PathBuild( _
    ByVal aFirst As String, _
    ByVal aSecond As String _
) As String
    '
    ' Returns a properly formatted path from two tokens.
    '
    PathBuild = CreateObject("Scripting.FileSystemObject").buildpath(aFirst, aSecond)
End Function

Public Function PathGetType( _
    ByVal aPath As String _
) As PathType
    '
    ' Returns a value indicating whether the a path is a properly
    ' formatted file, folder, drive or none of these (see PathType
    ' enum above). This function only checks for formatting, and the
    ' path need not actually exist. Only paths containing an extended
    ' file name as the last component will return ptFile. Paths
    ' containing unsupported characters always return ptNone.
    '
    If Not FileName(aPath) Like "*[!A-Za-z0-9\.\(\),_-]*" Then
        If Not (FileName(aPath) = "" Or FileExtension(aPath) = "") Then
            PathGetType = ptFile
        ElseIf FileFolder(aPath) <> "" Then
            PathGetType = ptFolder
        ElseIf FileDrive(aPath) <> "" Then
            PathGetType = ptDrive
        End If
    End If
End Function

Public Function PathValid( _
    ByVal aPath As String _
) As Boolean
    '
    ' Returns TRUE if the given file, folder or drive path exists,
    ' else returns FALSE.
    '
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    PathValid = fso.FileExists(aPath) Or fso.FolderExists(aPath) Or fso.DriveExists(aPath)
    Set fso = Nothing
End Function

Public Function PathPicker( _
    ByVal aPickerType As PickerType, _
    Optional ByVal aTitle As String = kstrSelectPath, _
    Optional ByVal aFilters As String = kstrAllFiles, _
    Optional ByVal aInitFolder As String = "" _
) As String
    '
    ' Opens a standard file dialog of the given type, with the
    ' specified title, filters and initial folder, and returns
    ' the picked object's name, if any.
    '
    Select Case aPickerType
        Case pkFile:
            PathPicker = FilePicker(aTitle, aFilters, aInitFolder)
        Case pkFolder:
            PathPicker = FolderPicker(aTitle, aInitFolder)
        Case Else:
    End Select
End Function

Public Function PathTerminate( _
    ByVal aPath As String, _
    Optional ByVal aChar As String = "\" _
) As String
    '
    ' Returns a string appended with a trailing backslash, or
    ' optionally another character.
    '
    PathTerminate = aPath
    If Right(PathTerminate, 1) <> aChar Then PathTerminate = PathTerminate & aChar
End Function

Public Sub PropertyCreate( _
    aDatabase As Database, _
    ByVal aProperty As String, _
    aValue As Variant, _
    Optional ByVal aType As Integer = dbText _
)
    '
    ' Creates a new database property with the given name, value and
    ' optionally of the given type.
    '
    Dim prp As Property
    
    Set prp = aDatabase.CreateProperty(aProperty, aType, aValue)
    aDatabase.properties.Append prp
    Set prp = Nothing
End Sub

Public Sub PropertyDelete( _
    aDatabase As Database, _
    ByVal aProperty As String _
)
    '
    ' Deletes the named database property, if it exists.
    '
    If PropertyExists(aDatabase, aProperty) Then aDatabase.properties.Delete aProperty
End Sub

Public Function PropertyExists( _
    aDatabase As Database, _
    ByVal aProperty As String _
) As Boolean
    '
    ' Returns TRUE if the named database property exists,
    ' else returns FALSE.
    '
    On Error GoTo Catch
    With aDatabase.properties(aProperty)
        PropertyExists = True
    End With
    
Finally:
    Exit Function
    
Catch:
    If Err.Number = kvbErrDbPropertyNotFound Then Resume Finally
    Err.Raise Err.Number
End Function

Public Function PropertyGet( _
    aDatabase As Database, _
    ByVal aProperty As String _
) As Variant
    '
    ' Returns the current value of a database property in the
    ' given database.
    '
    On Error Resume Next
    PropertyGet = aDatabase.properties(aProperty).Value
End Function

Public Function PropertyLoad( _
    ByVal aPropertyName As String, _
    Optional ByVal aDefault As Variant, _
    Optional aDatabase As DAO.Database _
) As Variant
    '
    ' Returns the current value of a database property, or
    ' the default value, if the property doesn't exist. If
    ' aDatabase is omitted, the current database is used.
    '
    If aDatabase Is Nothing Then Set aDatabase = CurrentDb
    PropertyLoad = aDefault
    If PropertyExists(aDatabase, aPropertyName) Then PropertyLoad = PropertyGet(aDatabase, aPropertyName)
End Function

Public Sub PropertySet( _
    aDatabase As Database, _
    ByVal aProperty As String, _
    aValue As Variant, _
    Optional ByVal aType As Integer = kdbNone _
)
    '
    ' Sets the value of a database property to the given value.
    ' If the property does not exist, it is created using the
    ' optional type.
    '
    If Not PropertyExists(CurrentDb, aProperty) Then
        If aType <> kdbNone Then _
        PropertyCreate aDatabase, aProperty, aValue, aType
    Else
        aDatabase.properties(aProperty).Value = aValue
    End If
End Sub

Public Sub PropertyUpdate( _
    ByVal aPropertyName As String, _
    Optional aValue As Variant = Empty, _
    Optional ByVal aDefault As Variant = Empty, _
    Optional ByVal aType As Integer, _
    Optional aDatabase As DAO.Database _
)
    '
    ' Sets the current value of a database property to the
    ' given or default value, or deletes the property if
    ' neither value is specified. If aDatabase is omitted,
    ' the current database is used.
    '
    If aDatabase Is Nothing Then Set aDatabase = CurrentDb
    aValue = IIf(IsValue(aValue), aValue, aDefault)
    If IsValue(aValue) Then
        PropertySet aDatabase, aPropertyName, aValue, aType
    Else
        PropertyDelete aDatabase, aPropertyName
    End If
End Sub

Public Sub RecordsetClear( _
    aRst As DAO.Recordset _
)
    '
    ' Deletes all records from a recordset.
    '
    With aRst
        If Not RecordsetEmpty(aRst) Then .MoveFirst
        While Not (.BOF Or .EOF)
            .Delete
            .MoveNext
        Wend
    End With
End Sub

Public Sub RecordsetAddNew( _
    aRst As DAO.Recordset, _
    ParamArray aArgs() As Variant _
)
    '
    ' Adds a new record to a recordset object with the given
    ' arguments. The args must be given as field-value pairs,
    ' e.g. "Field1",Value1,"Field2",Value2, ... ,"FieldN",ValueN.
    '
    If Not IsMissing(aArgs) Then RecordsetDo aRst, roAddnew, CVar(aArgs)
End Sub

Public Sub RecordsetDelete( _
    aRst As DAO.Recordset _
)
    '
    ' Deletes the current record in the given recordset object.
    '
    RecordsetDo aRst, roDelete
End Sub

Private Sub RecordsetDo( _
    aRst As DAO.Recordset, _
    ByVal aOperation As RecordsetOperation, _
    Optional aArgs As Variant _
)
    '
    ' Helper procedure that performs an operation on a recordset
    ' object (see RecordsetAddNew, RecordsetDelete, RecordsetEdit).
    '
    Dim Arg As Variant
    Dim fld As String
    Dim flag As Boolean
    
    Select Case aOperation
        Case roAddnew:
            aRst.AddNew
        Case roEdit:
            aRst.Edit
        Case roDelete:
            aRst.Delete
            Exit Sub
        Case Else:
    End Select
    If IsArray(aArgs) Then
        For Each Arg In aArgs
            If flag Then
                aRst.Fields(fld) = Arg
            Else
                fld = CStr(Arg)
            End If
            flag = Not flag
        Next
        aRst.Update
    Else
        Err.Raise kvbErrDbInvalidArgument, "LibVBA:RecordsetDo()", "aArgs must be an array of type Variant"
    End If
End Sub

Public Sub RecordsetEdit( _
    aRst As DAO.Recordset, _
    ParamArray aArgs() As Variant _
)
    '
    ' Changes the value of the given fields to the given values.
    ' The args must given as field-value pairs,
    ' e.g. "Field1",Value1,"Field2",Value2, ... ,"FieldN",ValueN.
    '
    If Not IsMissing(aArgs) Then RecordsetDo aRst, roEdit, CVar(aArgs)
End Sub

Public Function RecordsetCount( _
    aRst As DAO.Recordset _
) As Long
    '
    ' Returns the current number of records in the given recordset object.
    '
    If Not RecordsetEmpty(aRst) Then
        With aRst.Clone
            .MoveLast
            RecordsetCount = .RecordCount
        End With
    End If
End Function

Public Function RecordsetEmpty( _
    aRst As DAO.Recordset _
) As Boolean
    '
    ' Returns TRUE if the given recordset object is empty (has a record count of 0),
    ' else returns FALSE.
    '
    RecordsetEmpty = (aRst.BOF And aRst.EOF)
End Function

Public Function RecordsetReopen( _
    aRst As DAO.Recordset, _
    ByVal aQuery As String, _
    Optional ByVal aType As Variant, _
    Optional ByVal aOptions As Variant, _
    Optional ByVal aLockEdit As Variant _
) As Long
    '
    ' (Re)opens a recordset object using the given query string
    ' and returns the record count.
    '
    If Not aRst Is Nothing Then aRst.Close
    Set aRst = CurrentDb.OpenRecordset(aQuery, aType, aOptions, aLockEdit)
    RecordsetReopen = RecordsetCount(aRst)
End Function

Public Function RemoteCall( _
    aDatabaseName As String, _
    ByVal aProcedureName As String, _
    ParamArray aArgs() As Variant _
) As Variant
    '
    ' Executes a remote procedure call in another database
    ' and returns the value returned by the called procedure.
    '
    Dim app As New Access.Application
    Dim dbproc As String
    
    dbproc = FileBaseName(aDatabaseName) & "." & aProcedureName
    app.OpenCurrentDatabase aDatabaseName, False
    app.Visible = False
    If Not IsMissing(aArgs) Then
        RemoteCall = app.Run(dbproc, CVar(aArgs))
    Else
        RemoteCall = app.Run(dbproc)
    End If
    app.Quit
    Set app = Nothing
End Function

Public Function ReplaceTags( _
    ByVal aString As String, _
    ParamArray aArgs() As Variant _
) As String
    '
    ' Convenience function that is easier to use than multiple
    ' nested Replace() functions. Search and replace arguments
    ' must be supplied in pairs, e.g."search for", "replace with",
    ' otherwise an error may occur.
    '
    Dim i As Integer
    
    For i = LBound(aArgs) To UBound(aArgs) Step 2
        aString = Replace(aString, CStr(aArgs(i)), CStr(aArgs(i + 1)))
    Next
    ReplaceTags = aString
End Function

Public Function StringDelim( _
    ByVal aString As String, _
    Optional ByVal aLeft As String = """", _
    Optional ByVal aRight As String = """", _
    Optional ByVal aForce As Boolean = False _
) As String
    '
    ' Optionally returns a string surrounded by the given delimiting
    ' strings. If the force flag is set, the delimiters are applied
    ' automatically, else they are only applied if the string is not
    ' already delimited. The default delimiter is a double-quote.
    '
    If Not aForce Then
        If Left(aString, Len(aLeft)) = aLeft Then aLeft = ""
        If Right(aString, Len(aRight)) = aRight Then aRight = ""
    End If
    StringDelim = aLeft & aString & aRight
End Function

Public Function StringQuote( _
    ByVal aString As String, _
    Optional ByVal aChar As String = """" _
) As String
    '
    ' Returns a string surrounded by the given char.
    ' If omitted, char defaults to a double-quote.
    '
    StringQuote = aChar & aString & aChar
End Function

Public Function StringTab( _
    ByVal aToken As String, _
    Optional ByVal aSpacing As Integer = 4 _
) As String
    '
    ' Appends an emulated tab character to the end of a
    ' string using spaces. Useful for displaying strings
    ' in controls that do not display tab characters (e.g.
    ' TextBoxes). The spacing argument is the tab spacing
    ' in characters.
    '
    StringTab = Space(MaxOf((Len(aToken) Mod aSpacing), 1))
End Function

Public Sub Swap( _
    ByRef a As Variant, _
    ByRef b As Variant _
)
    '
    ' Swaps the contents of a and b. No type checking is performed.
    '
    Dim temp As Variant
    
    If IsObject(a) Then
        Set temp = a
        Set a = b
        Set b = temp
    Else
        temp = a
        a = b
        b = temp
    End If
End Sub

Public Function TableExists( _
    aDatabase As Database, _
    ByVal aTableName As String _
) As Boolean
    '
    ' Returns TRUE if the named table exists in the given database,
    ' else returns FALSE.
    '
    On Error Resume Next
    TableExists = IsObject(aDatabase.TableDefs(aTableName))
    Err.Clear
End Function

Public Function TableLink( _
    ByVal aSrcDatabase As String, _
    ByVal aSrcTable As String, _
    ByVal aDestTable As String, _
    Optional ByVal aSrcDbType As String = "Microsoft Access" _
) As String
    '
    ' Links the given source table from the source database
    ' to the destination table in the current database.
    ' Returns the destination table name.
    '
    DoCmd.TransferDatabase TransferType:=acLink, _
        DatabaseType:=aSrcDbType, _
        DatabaseName:=aSrcDatabase, _
        ObjectType:=acTable, _
        Source:=aSrcTable, _
        Destination:=aDestTable
    TableLink = aDestTable
End Function

Public Function TablesList( _
    Optional ByVal aExcludeMask As Long = 0, _
    Optional aDatabase As DAO.Database _
) As Variant
    '
    ' Returns a list of database tables names,
    ' excluding any types specified by the exclude
    ' mask. See DAO.TableDef.Attributes for mask values.
    '
    Dim tdf As DAO.TableDef, tbls() As Variant
    
    If aDatabase Is Nothing Then Set aDatabase = CurrentDb
    For Each tdf In aDatabase.TableDefs
        Dim exclude As Boolean
        
        exclude = (tdf.Attributes And aExcludeMask)
        If Not exclude Then ArrayPushBack tbls, tdf.Name
    Next
    TablesList = tbls
End Function

Public Function TextBoxAppend( _
    ByRef aTextbox As TextBox, _
    ByVal aText As String, _
    Optional ByVal aPrepend As Boolean = False _
) As Long
    '
    ' Appends text to a multiline textbox and scrolls to the
    ' bottom so that the latest text is always visible. The
    ' newline can be optionally prepended or appended to the text.
    ' NOTE: Multi-line TextBoxes are limited to 65535 charcters.
    ' Attempting to append more characters causes an error.
    '
    With aTextbox
        .SetFocus
        If aPrepend Then
            .Value = .Value & vbNewLine & aText
        Else
            .Value = .Value & aText & vbNewLine
        End If
        .SelStart = Len(.Value)
        .SelLength = 0
        TextBoxAppend = Len(.Value)
    End With
End Function

Public Function TextConcat( _
    ByVal aDelim As String, _
    ParamArray aTokens() As Variant _
) As String
    '
    ' Concatenates the given tokens into a delimited string.
    '
    If Not IsMissing(aTokens) Then TextConcat = Join(aTokens, aDelim)
End Function

Public Function TimeoutMs( _
    ByVal aCurrent As Double, _
    ByVal aStart As Double, _
    ByVal aTimeout As Long _
) As Boolean
    '
    ' Returns TRUE if difference between the current time and start time
    ' is greater than or equal to the given timeout value in milliseconds,
    ' else returns FALSE. Always returns FALSE if the timeout is WAIT_INFINITE.
    ' Works well with the built-in Timer() function.
    '
    If aTimeout <> WAIT_INFINITE Then TimeoutMs = ((aCurrent - aStart) * 1000 >= aTimeout)
End Function

Public Function TimerStart( _
    aForm As Form, _
    aInterval As Long _
) As Long
    '
    ' Starts the form timer with the given interval.
    '
    aForm.TimerInterval = aInterval
    TimerStart = aInterval
End Function

Public Sub TimerStop( _
    aForm As Form _
)
    '
    ' Stops the form timer.
    '
    aForm.TimerInterval = 0
End Sub

Public Function UsrErr( _
    ByVal aErrNum As Long _
) As Long
    '
    ' Returns a "user error" code guaranteed to be outside the
    ' range of any VBA or system error codes.
    '
    UsrErr = vbObjectError + aErrNum
End Function

Public Function VarTypeText( _
    aVar As Variant _
) As String
    '
    ' Returns the human-readable VBA variable type.
    '
    Dim vt As Integer
    
    vt = VarType(aVar)
    Select Case vt
        Case vbEmpty:
            VarTypeText = "vbEmpty"
        Case vbNull:
            VarTypeText = "vbNull"
        Case vbInteger:
            VarTypeText = "vbInteger"
        Case vbLong:
            VarTypeText = "vbLong"
        Case vbSingle:
            VarTypeText = "vbSingle"
        Case vbDouble:
            VarTypeText = "vbDouble"
        Case vbCurrency:
            VarTypeText = "vbCurrency"
        Case vbDate:
            VarTypeText = "vbDate"
        Case vbString:
            VarTypeText = "vbString"
        Case vbObject:
            VarTypeText = "vbObject"
        Case vbError:
            VarTypeText = "vbError"
        Case vbBoolean:
            VarTypeText = "vbBoolean"
        Case vbVariant:
            VarTypeText = "vbVariant"
        Case vbDataObject:
            VarTypeText = "vbDataObject"
        Case vbDecimal:
            VarTypeText = "vbDecimal"
        Case vbByte:
            VarTypeText = "vbByte"
        Case vbLongLong:
            VarTypeText = "vbLongLong"
        Case vbUserDefinedType:
            VarTypeText = "vbUserDefinedType"
        Case Is >= vbArray:
            VarTypeText = "vbArray"
        Case Else:
            VarTypeText = CStr(vt)
    End Select
End Function

Public Function DbTypeText( _
    aObject As Variant _
) As String
    '
    ' Returns the human-readable operational data type of the given object.
    ' If object is numeric, returns the db type name of the given value,
    ' (e.g. db property type), else returns the field's db type name.
    '
    Dim vt As Integer
    
    If IsNumeric(aObject) Then
        vt = CInt(aObject)
    Else
        vt = aObject.Type
    End If
    Select Case vt
        Case dbAttachment:
            DbTypeText = "dbAttachment"
        Case dbMemo:
            DbTypeText = "dbMemo"
        Case dbNumeric:
            DbTypeText = "dbNumeric"
        Case dbBigInt:
            DbTypeText = "dbBigInt"
        Case dbBinary:
            DbTypeText = "dbBinary"
        Case dbLongBinary:
            DbTypeText = "dbLongBinary"
        Case dbInteger:
            DbTypeText = "dbInteger"
        Case dbLong:
            DbTypeText = "dbLong"
        Case dbSingle:
            DbTypeText = "dbSingle"
        Case dbDouble:
            DbTypeText = "dbDouble"
        Case dbFloat:
            DbTypeText = "dbFloat"
        Case dbGUID:
            DbTypeText = "dbGUID"
        Case dbCurrency:
            DbTypeText = "dbCurrency"
        Case dbDate:
            DbTypeText = "dbDate"
        Case dbTime:
            DbTypeText = "dbTime"
        Case dbTimeStamp:
            DbTypeText = "dbTimeStamp"
        Case dbText:
            DbTypeText = "dbText"
        Case dbChar:
            DbTypeText = "dbChar"
        Case dbBoolean:
            DbTypeText = "dbBoolean"
        Case dbDecimal:
            DbTypeText = "dbDecimal"
        Case dbByte:
            DbTypeText = "dbByte"
        Case dbDateTimeExtended:
            DbTypeText = "dbDateTimeExtended"
        Case Else:
            DbTypeText = CStr(vt)
    End Select
End Function
