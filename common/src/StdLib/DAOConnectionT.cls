VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DAOConnectionT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' DAOConnectionT                                            '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a class that encapsulates a database  '
' connection string accessible using dot notation.          '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' (None)                                                    '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the library has only been tested with     '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database
Option Explicit

'''''''''''''''''''
' Private Members '
'''''''''''''''''''

Private CString As String   ' The current connection string.

''''''''''''''''''''
' Public Interface '
''''''''''''''''''''

Public Property Get Address() As String
    Address = GetOption("Addr")
End Property

Public Property Let Address( _
    aAddress As String _
)
    SetOption "Addr", aAddress
End Property

Public Property Get AnsiNpw() As String
    AnsiNpw = GetOption("AnsiNPW")
End Property

Public Property Let AnsiNpw( _
    aEnabled As String _
)
    SetOption "AnsiNPW", aEnabled
End Property

Public Property Get Application() As String
    Application = GetOption("APP")
End Property

Public Property Let Application( _
    aApplication As String _
)
    SetOption "APP", aApplication
End Property

Public Property Get ApplicationIntent() As String
    ApplicationIntent = GetOption("ApplicationIntent")
End Property

Public Property Let ApplicationIntent( _
    aIntent As String _
)
    SetOption "ApplicationIntent", aIntent
End Property

Public Property Get AttachDbFileName() As String
    AttachDbFileName = GetOption("AttachDBFileName")
End Property

Public Property Let AttachDbFileName( _
    aFileName As String _
)
    SetOption "AttachDBFileName", aFileName
End Property

Public Property Get AutoTranslate() As String
    AutoTranslate = GetOption("AutoTranslate")
End Property

Public Property Let AutoTranslate( _
    aEnabled As String _
)
    SetOption "AutoTranslate", aEnabled
End Property

Public Property Get Connect() As Variant
    Connect = CString
End Property

Public Property Let Connect( _
    ByVal aConnect As Variant _
)
    If Nz(aConnect) = "" Then aConnect = CurrentDb.Name
    CString = IIf(PathGetType(aConnect) = ptFile, "Microsoft Access;DATABASE=" & aConnect, aConnect)
End Property

Public Property Get ConnectionTimeout() As String
    ConnectionTimeout = GetOption("Connection Timeout")
End Property

Public Property Let ConnectionTimeout( _
    ByVal aTimeout As String _
)
    SetOption "Connection Timeout", aTimeout
End Property

Public Property Get Database() As String
    Database = GetOption("DATABASE")
End Property

Public Property Let Database( _
    ByVal aDatabase As String _
)
    SetOption "DATABASE", aDatabase
End Property

Public Property Get Driver() As String
    Driver = GetOption("DRIVER")
End Property

Public Property Let Driver( _
    ByVal aDriver As String _
)
    SetOption "DRIVER", aDriver
End Property

Public Property Get DataSource() As String
    DataSource = GetOption("DSN")
End Property

Public Property Let DataSource( _
    ByVal aDataSource As String _
)
    SetOption "DSN", aDataSource
End Property

Public Property Get Encrypt() As String
    Encrypt = GetOption("ENCRYPT")
End Property

Public Property Let Encrypt( _
    ByVal aEncrypt As String _
)
    SetOption "ENCRYPT", aEncrypt
End Property

Public Property Get FailoverPartner() As String
    FailoverPartner = GetOption("Failover_Partner")
End Property

Public Property Let FailoverPartner( _
    ByVal aPartner As String _
)
    SetOption "Failover_Partner", aPartner
End Property

Public Property Get FailoverPartnerSpn() As String
    FailoverPartnerSpn = GetOption("FailoverPartnerSPN")
End Property

Public Property Let FailoverPartnerSpn( _
    ByVal aSpn As String _
)
    SetOption "FailoverPartnerSPN", aSpn
End Property

Public Property Get FileDsn() As String
    FileDsn = GetOption("FileDSN")
End Property

Public Property Let FileDsn( _
    ByVal aDsn As String _
)
    SetOption "FileDSN", aDsn
End Property

Public Property Get MarsConnection() As String
    MarsConnection = GetOption("MARS_Connection")
End Property

Public Property Let MarsConnection( _
    ByVal aConnection As String _
)
    SetOption "MARS_Connection", aConnection
End Property

Public Property Get MultiSubnetFailover() As String
    MultiSubnetFailover = GetOption("MultiSubnetFailover")
End Property

Public Property Let MultiSubnetFailover( _
    ByVal aFailover As String _
)
    SetOption "MultiSubnetFailover", aFailover
End Property

Public Property Get Name() As String
    Name = GetOption("")
End Property

Public Property Let Name( _
    ByVal aName As String _
)
    SetOption "", aName
End Property

Public Property Get Net() As String
    Net = GetOption("Net")
End Property

Public Property Let Net( _
    ByVal aNet As String _
)
    SetOption "Net", aNet
End Property

Public Property Get Password() As String
    Password = GetOption("PWD")
End Property

Public Property Let Password( _
    ByVal aPassword As String _
)
    SetOption "PWD", aPassword
End Property

Public Property Get QueryLogFile() As String
    QueryLogFile = GetOption("QueryLogFile")
End Property

Public Property Let QueryLogFile( _
    ByVal aFilePath As String _
)
    SetOption "QueryLogFile", aFilePath
End Property

Public Property Get QueryLogOn() As String
    QueryLogOn = GetOption("QueryLog_On")
End Property

Public Property Let QueryLogOn( _
    ByVal aEnabled As String _
)
    SetOption "QueryLog_On", aEnabled
End Property

Public Property Get QueryLogTime() As String
    QueryLogTime = GetOption("QueryLogTime")
End Property

Public Property Let QueryLogTime( _
    ByVal aTime As String _
)
    SetOption "QueryLogTime", aTime
End Property

Public Property Get QuotedId() As String
    QuotedId = GetOption("QuotedId")
End Property

Public Property Let QuotedId( _
    ByVal aId As String _
)
    SetOption "QuotedId", aId
End Property

Public Property Get Regional() As String
    Regional = GetOption("Regional")
End Property

Public Property Let Regional( _
    ByVal aRegion As String _
)
    SetOption "Regional", aRegion
End Property

Public Property Get SaveFile() As String
    SaveFile = GetOption("SaveFile")
End Property

Public Property Let SaveFile( _
    ByVal aDsFile As String _
)
    SetOption "SaveFile", aDsFile
End Property

Public Property Get Server() As String
    Server = GetOption("SERVER")
End Property

Public Property Let Server( _
    ByVal aServer As String _
)
    SetOption "SERVER", aServer
End Property

Public Property Get ServerSpn() As String
    ServerSpn = GetOption("ServerSPN")
End Property

Public Property Let ServerSpn( _
    ByVal aSpn As String _
)
    SetOption "ServerSPN", aSpn
End Property

Public Property Get StatsLogFile() As String
    StatsLogFile = GetOption("StatsLogFile")
End Property

Public Property Let StatsLogFile( _
    ByVal aFilePath As String _
)
    SetOption "StatsLogFile", aFilePath
End Property

Public Property Get StatsLogOn() As String
    StatsLogOn = GetOption("StatsLog_On")
End Property

Public Property Let StatsLogOn( _
    ByVal aEnabled As String _
)
    SetOption "StatsLog_On", aEnabled
End Property

Public Property Get TrustedConnection() As String
    TrustedConnection = GetOption("TRUSTED_CONNECTION")
End Property

Public Property Let TrustedConnection( _
    ByVal aTrusted As String _
)
    SetOption "TRUSTED_CONNECTION", aTrusted
End Property

Public Property Get TrustServerCertificate() As String
    TrustServerCertificate = GetOption("TrustServerCertificate")
End Property

Public Property Let TrustServerCertificate( _
    ByVal aTrust As String _
)
    SetOption "TrustServerCertificate", aTrust
End Property

Public Property Get UserId() As String
    UserId = GetOption("UID")
End Property

Public Property Let UserId( _
    ByVal aUserId As String _
)
    SetOption "UID", aUserId
End Property

Public Property Get WorkstationId() As String
    WorkstationId = GetOption("WSID")
End Property

Public Property Let WorkstationId( _
    ByVal aWorkstationId As String _
)
    SetOption "WSID", aWorkstationId
End Property

Private Function IsMatch( _
    ByVal aKey As String, _
    ByVal aKeyVal As String, _
    Optional ByVal aIsType As Boolean = False _
) As Boolean
    '
    ' Returns TRUE if the given key-value string contains the
    ' key string or, the key is blank and the key-value string
    ' is a value only, else returns FALSE. This logic assumes
    ' that connection type specifiers are always values only,
    ' e.g. "ODBC" or "Microsoft Access", and any other
    ' connection options are always key-value pairs, e.g.
    ' "DATABASE=mydatabase" or "SERVER=(local)".
    '
    IsMatch = IIf((aKey <> ""), (InStr(1, aKeyVal, aKey) > 0), (InStr(1, aKeyVal, "=") = 0))
End Function

Private Function GetOption( _
    ByVal aOption As String, _
    Optional ByVal aString As Variant _
) As Variant
    '
    ' Returns the value of aOption from the current
    ' connection string, if any.
    '
    Dim items() As String, item As Variant
    
    If aOption <> "" Then aOption = aOption & "="
    items = Split(IIf(Not IsMissing(aString), aString, CString), ";")
    For Each item In items
        If Len(Trim(item)) > 0 Then
            If IsMatch(aOption, item) Then
                Dim keyval() As String
                
                keyval = Split(item, "=")
                GetOption = keyval(UBound(keyval))
                Exit For
            End If
        End If
    Next
End Function

Private Sub SetOption( _
    ByVal aOption As String, _
    ByVal aValue As String _
)
    '
    ' Replaces the option value in or appends it to the current
    ' connection string depending on whether or not it exists
    ' in the string.
    '
    Dim items() As String, item As Variant, keyval As String
    
    If aOption <> "" Then aOption = aOption & "="
    keyval = aOption & aValue
    items = Split(CString, ";")
    For Each item In items
        If Len(Trim(item)) > 0 Then
            If IsMatch(aOption, item) Then
                CString = Replace(CString, item, IIf(aValue = "", "", keyval))
                Exit Sub
            End If
        End If
    Next
    CString = CString & IIf(aValue = "", "", keyval & ";")
End Sub
