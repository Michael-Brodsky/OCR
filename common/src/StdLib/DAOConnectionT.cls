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
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
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
' None                                                      '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the library has only been tested with     '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database
Option Explicit

Private cnstr As String     ' The current connection string.

Public Property Get Application() As String
    Application = GetOption("APP")
End Property

Public Property Let Application( _
    aApplication As String _
)
    SetOption "APP", aApplication
End Property

Public Property Get Connect() As Variant
    Connect = cnstr
End Property

Public Property Let Connect(aConnect As Variant)
    If Nz(aConnect) = "" Then aConnect = CurrentDb.name
    cnstr = IIf(PathGetType(aConnect) = ptFile, "Microsoft Access;DATABASE=" & aConnect, aConnect)
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

Public Property Get Password() As String
    Password = GetOption("PWD")
End Property

Public Property Let Password( _
    ByVal aPassword As String _
)
    SetOption "PWD", aPassword
End Property

Public Property Get Server() As String
    Server = GetOption("SERVER")
End Property

Public Property Let Server( _
    ByVal aServer As String _
)
    SetOption "SERVER", aServer
End Property

Public Property Get name() As String
    name = GetOption("")
End Property

Public Property Let name( _
    ByVal aName As String _
)
    SetOption "", aName
End Property

Public Property Get TrustedConnection() As String
    TrustedConnection = GetOption("TRUSTED_CONNECTION")
End Property

Public Property Let TrustedConnection( _
    ByVal aTrusted As String _
)
    SetOption "TRUSTED_CONNECTION", aTrusted
End Property

Public Property Get UserId() As String
    UserId = GetOption("UID")
End Property

Public Property Let UserId( _
    ByVal aUserId As String _
)
    SetOption "UID", aUserId
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
    ' Returns the value of the given option from the current
    ' connection string, if any.
    '
    Dim items() As String, item As Variant
    
    If aOption <> "" Then aOption = aOption & "="
    items = Split(IIf(Not IsMissing(aString), aString, cnstr), ";")
    For Each item In items
        If IsMatch(aOption, item) Then
            Dim keyval() As String
            
            keyval = Split(item, "=")
            GetOption = keyval(UBound(keyval))
            Exit For
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
    items = Split(cnstr, ";")
    For Each item In items
        If IsMatch(aOption, item) Then
            cnstr = Replace(cnstr, item, IIf(aValue = "", "", keyval))
            Exit Sub
        End If
    Next
    cnstr = cnstr & IIf(aValue = "", "", keyval & ";")
End Sub
