Attribute VB_Name = "basSetup"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' basSetup                                                  '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module downloads and installs any dependencies and   '
' configures the application to run in the current          '
' environment.                                              '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' StdLibMin                                                 '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the module has only been tested with      '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

' Type that aggregates information about a library database.
Private Type AppReference
    Name As String      ' The library name as it appears in the project exploder,
    FilePath As String  ' The library's file name,
    Tables() As String  ' List of library tables to link to the current database.
End Type

' Type that aggregates information about an error.
Private Type SetupError
    Number As Long          ' Error number,
    Source As String        ' Error source,
    Description As String   ' Error description.
End Type

' Type that aggregates information about a setup procedure.
Private Type SetupProcedure
    Name As String      ' The procedure identifier,
    Args() As Variant   ' List of procedure parameters.
End Type

' Type that aggregates information about a setup setting.
Private Type SetupProperty
    Name As String      ' The setting property name,
    Value As Variant    ' The setting value,
    Type As Integer     ' The setting data type.
End Type

' Type that aggregates information about an application setting.
Private Type SetupSetting
    TableName As String ' The setting table name,
    FieldName As String ' field name and
    Value As String     ' value.
End Type

'''''''''''''''''''
' Setup Constants '
'''''''''''''''''''

' Setup procedure constants.
Private Const kStrSetupWarning As String = _
    "The application needs to be configured for your environment. Run Setup now?"
Private Const kStrSetupRequired As String = _
    "You must run the Setup before using the application."
Private Const kSetupTags As String = _
    "<ProjectDir>;IIf(Len(Nz(get_prp(""OCRAppInstallDir""))) = 0, CurrentProject.Path, get_prp(""OCRAppInstallDir""));" & _
    "<ProjectName>;base_name(CurrentProject.Name)"

' Application dependencies: "ProcedureName[,Args0,Args1,...,ArgsN]"
' Here Arg0 is the local procedure name (see Dependency Installers below),
' Args1 is the installer download URL,
' Args2 is our local save to folder.

' Application references: "ReferenceName,FilePath[,TableName_0,TableName_1,...,TableName_N]"

' Default application settings: "PropertyName,Value,DataType"

' Application files & folders: "FilePath0,FilePath1,...,FilePath1"

' Navigation pane object hidden attribute: "ObjectName,ObjectType,TRUE/FALSE"

Private boolInSetup As Boolean  ' Flag to prevent multiple re-entrant calls to the setup procedure.
Private errError As SetupError  ' Stores error info from procedures called with Application.Run().

''''''''''''''''''''
' Public Interface '
''''''''''''''''''''

Public Function OcrAdminRunSetup()
    '
    ' Prompts the user to run setup if not already.
    '
End Function

'''''''''''''''''''''
' Private Interface '
'''''''''''''''''''''

Private Sub OcrSetupReset()

End Sub

Private Function OcrAdminRmRefs( _
    ParamArray aReferences() As Variant _
)
    '
    ' Remove our library references. This is to circumvent MS Access'
    ' attempts at locating broken references by removing our references
    ' and reloading them from the known install folder.
    '
    Dim item As Variant

    For Each item In aReferences
        Dim aRef As Reference, ref As AppReference
        
        ref = RefInfoUnpack(item)
        For Each aRef In Application.References
            If aRef.Name = ref.Name Then
                Application.References.Remove aRef
            End If
        Next
    Next
End Function

Private Sub OcrAdminSetup()
    '
    ' Runs the initial setup procedures.
    '
End Sub

Private Sub FixReferences( _
    ParamArray aReferences() As Variant _
)
    '
    ' Relinks any referenced objects.
    '
    Dim aRef As Reference, ref As Variant
    
    ' Remove any broken references.
    For Each aRef In Application.References
        If aRef.IsBroken Then
            Application.References.Remove aRef
        End If
    Next
    ' Remove our library references so Access doesn't search for them on startup.
    OcrAdminRmRefs
    For Each ref In aReferences
        Dim appRef As AppReference
        
        appRef = RefInfoUnpack(ref)
        Set aRef = GetReference(appRef.Name)                        ' Relink our library references.
        If aRef Is Nothing Then
            Application.References.AddFromFile ReplaceSetupTags(kSetupTags, appRef.FilePath)
        End If
        On Error GoTo 0
        If (Not Not appRef.Tables) <> 0 Then                        ' If it has linked tables, ...
            Dim tbl As Variant
            
            On Error Resume Next
            For Each tbl In appRef.Tables
                Dim tblExists As Boolean
                
                tblExists = IsObject(CurrentDb.TableDefs(tbl))      ' (Throws an error if table not found.)
                If tblExists Then DoCmd.DeleteObject acTable, tbl   ' ... remove and ....
                DoCmd.TransferDatabase TransferType:=acLink, _
                    DatabaseType:="Microsoft Access", _
                    DatabaseName:=ReplaceSetupTags(kSetupTags, appRef.FilePath), _
                    ObjectType:=acTable, _
                    Source:=tbl, _
                    Destination:=tbl                                ' ... relink them.
                Application.SetHiddenAttribute acTable, tbl, True
            Next
            On Error GoTo 0
        End If
    Next
End Sub

Private Sub CheckDependencies( _
    ParamArray aDependencies() As Variant _
)
    '
    ' Check for any application dependencies.
    '
    If Not IsMissing(aDependencies) Then
        Dim dep As Variant
        
        For Each dep In aDependencies
            Dim proc As SetupProcedure
            
            proc = ProcInfoUnpack(dep)
            RunProcedure proc.Name, proc.Args
        Next
    End If
End Sub

Private Sub CreateFileSystem( _
    ParamArray aFolders() As Variant _
)
    '
    ' Creates any required folders.
    '
    If Not IsMissing(aFolders) Then
        Dim fldr As Variant, fso As Object
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        For Each fldr In aFolders
            Dim fpath As String
            
            fpath = ReplaceSetupTags(kSetupTags, fldr)
            If Not fso.FolderExists(fpath) Then
                fso.CreateFolder fpath
            End If
        Next
        Set fso = Nothing
    End If
End Sub

Private Sub SetAppDefaults( _
    ParamArray aSettings() As Variant _
)
    '
    ' Set default application settings.
    '
    If Not IsMissing(aSettings) Then
        Dim setting As Variant
        
        For Each setting In aSettings
            Dim appset As SetupProperty
            
            appset = AppPropertyUnpack(setting)
            rm_prp CurrentDb, appset.Name
            If Nz(appset.Value) <> "" Then
                appset.Value = ReplaceSetupTags(kSetupTags, appset.Value)
                set_prp CurrentDb, appset.Name, appset.Value, appset.Type
            End If
        Next
    End If
End Sub

Private Sub SetSettingsDefaults( _
    ParamArray aSettings() As Variant _
)
    '
    ' Set default settings in our tables.
    '
    If Not IsMissing(aSettings) Then
        Dim setting As Variant
        
        DoCmd.SetWarnings False
        For Each setting In aSettings
            Dim appset As SetupSetting
            
            appset = AppSettingUnpack(setting)
            DoCmd.RunSQL "UPDATE [" & appset.TableName & "] SET " & appset.FieldName & " = " & StringQuote(appset.Value) & ";"
        Next
        DoCmd.SetWarnings True
   End If

End Sub

Private Sub SetNavOptions( _
    ByVal aOptions As Boolean _
)
    '
    ' Set the nav pane and ribbon behavior according to given options.
    '
    Application.SetOption "Show Hidden Objects", aOptions
    PropertySet CurrentDb, "AllowBuiltinToolbars", aOptions, dbBoolean
    PropertySet CurrentDb, "AllowFullMenus", aOptions, dbBoolean
    PropertySet CurrentDb, "AllowShortcutMenus", aOptions, dbBoolean
    PropertySet CurrentDb, "AllowToolbarChanges", aOptions, dbBoolean
End Sub

Private Sub SetNavPaneAttribs( _
    ParamArray aHidden() As Variant _
)
    '
    ' Set nav pane object attributes.
    '
    If Not IsMissing(aHidden) Then
        Dim hidden As Variant
        
        For Each hidden In aHidden
            Dim params() As String, objname As String, objtype As Integer, objhide As Boolean
        
            params = Split(hidden, ",")
            objname = params(0)
            objtype = CInt(params(1))
            objhide = CBool(params(2))
            Application.SetHiddenAttribute objtype, objname, objhide
        Next
    End If
End Sub

''''''''''''''''''''
' Helper Functions '
''''''''''''''''''''

Private Function RefInfoUnpack( _
    ByVal aInfo As String _
) As AppReference
    Dim info() As String, ref As AppReference, i As Integer
    
    info = Split(aInfo, ",")
    ref.Name = info(0)
    ref.FilePath = info(1)
    For i = 2 To UBound(info)
        arr_add ref.Tables, info(i)
    Next
    RefInfoUnpack = ref
End Function

Private Function AppPropertyUnpack( _
    ByVal aProperty As String _
) As SetupProperty
    Dim prop As SetupProperty
    Dim info() As String
    
    info = Split(aProperty, ",")
    prop.Name = info(0)
    prop.Value = CVar(info(1))
    prop.Type = CInt(info(2))
    AppPropertyUnpack = prop
End Function

Private Function AppSettingUnpack( _
    ByVal aSetting As String _
) As SetupSetting
    Dim setting As SetupSetting
    Dim info() As String
    
    info = Split(aSetting, ",")
    setting.TableName = info(0)
    setting.FieldName = info(1)
    setting.Value = ReplaceSetupTags(kSetupTags, info(2))
    AppSettingUnpack = setting
End Function

Private Function ProcInfoUnpack( _
    ByVal aProc As String _
) As SetupProcedure
    Dim proc As SetupProcedure
    Dim info() As String, i As Integer
    
    info = Split(aProc, ",")
    proc.Name = info(0)
    For i = 1 To UBound(info)
        arr_add proc.Args, CVar(info(i))
    Next
    ProcInfoUnpack = proc
End Function

Private Function GetReference( _
    ByVal aReferenceName As String _
) As Reference
    On Error Resume Next
    Set GetReference = Application.References(aReferenceName)
    If Err.Number <> 0 Then
        If Err.Number = 9 Then
            Err.Clear
        Else
            Err.Raise Err.Number
        End If
    End If
End Function

Public Function ReplaceSetupTags( _
    ByVal aTags As String, _
    ByVal aString As String _
) As String
    Dim tags() As String, i As Integer
    
    tags = Split(aTags, ";")
    For i = LBound(tags) To UBound(tags) Step 2
        If InStr(1, aString, tags(i)) > 0 Then aString = Replace(aString, tags(i), Eval(tags(i + 1)))
    Next
    ReplaceSetupTags = aString
End Function

Private Sub ErrorClear( _
    ByRef aSetupError As SetupError _
)
    aSetupError.Number = 0
    aSetupError.Source = ""
    aSetupError.Description = ""
End Sub

Private Function RunProcedure( _
    ByVal aProcName As String, _
    Optional aArgs As Variant _
) As Variant
    '
    ' Run a setup procedure. The Application.Run call decouples
    ' any error handling in the called procedure. The procedure
    ' has to handle any errors and "pass" them back using the
    ' errError object, otherwise a runtime error will occur and
    ' halt execution.
    '
    RunProcedure = Application.Run(aProcName, aArgs)
    If errError.Number <> 0 Then Err.Raise errError.Number, errError.Source, errError.Description
End Function

Private Sub DownloadInstaller( _
    ByVal aUrl As String, _
    ByVal aSaveAs As String, _
    Optional ByVal aAppName As String _
)
    Dim fso As Object, fName As String, saveAs As String
    
    MsgBox "Setup needs to download and install " & aUrl, , base_name(CurrentProject.Name)
    If Not GetInternetConnectedState() Then _
    Err.Raise UsrErr(kOcrAdminErrSetup), CurrentProject.Name, "Setup requires an active internet connection. " & _
    "Connect this computer to the internet and run Setup again."
    Set fso = CreateObject("Scripting.FileSystemObject")
    fName = fso.GetFileName(aUrl)
    saveAs = fso.BuildPath(ReplaceSetupTags(kSetupTags, aSaveAs), fName)
    If URLDownloadToFile(0, aUrl, saveAs, 0, 0) <> 0 Then _
    Err.Raise UsrErr(kOcrAdminErrSetup), CurrentProject.Name, aUrl & " Download failed. You'll need to download and " & _
    "install " & fName & " before you can run this application."
    Set fso = Nothing
End Sub

'''''''''''''''''''''''''
' Dependency Installers '
'''''''''''''''''''''''''
