VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RsProcessDataRulesT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' ProcessRulesT                                              '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a class that encapsulates a DAO.      '
' Recordset holding information about a cataloging process  '
' data extraction and storage rules, that is accessible     '
' using dot notation.                                       '
'                                                           '
' DEPENDENCIES:                                             '
'                                                           '
' DAORecordsetT                                             '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the library has only been tested with     '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

Private Data As DAORecordsetT

Public Property Get BOF() As Boolean
    BOF = Data.BOF
End Property

Public Property Get Count() As Long
    Count = Data.Count
End Property

Public Property Get Name() As String
    Name = Data.Value("Rule Name")
End Property

Public Property Get EOF() As Boolean
    EOF = Data.EOF
End Property

Public Property Get StorageParameterField() As Variant
    StorageParameterField = Nz(Data.Value("Storage ParameterField"))
End Property

Public Property Get Global_() As Boolean
    Global_ = Data.Value("Global")
End Property

Public Property Get IgnoreCase() As Boolean
    IgnoreCase = Data.Value("Ignore Case")
End Property

Public Property Get Match() As Integer
    Match = Data.Value("Match")
End Property

Public Property Get Pattern() As String
    Pattern = Data.Value("Pattern")
End Property

Public Property Get Required() As Boolean
    Required = Data.Value("Required")
End Property

Public Property Get Submatch() As Integer
    Submatch = Data.Value("Submatch")
End Property

Public Sub MoveFirst()
    Data.MoveFirst
End Sub

Public Sub MoveNext()
    Data.MoveNext
End Sub

Public Sub Close_()
    If Not Data Is Nothing Then Data.Close_
    Set Data = Nothing
End Sub

Public Sub Open_( _
    ByVal aProcessId As Long _
)
    Dim qry As DAO.QueryDef
    
    Close_
    Set Data = NewDAORecordsetT()
    ' TODO: The qry needs to be replaced with something that doesn't replace text in another query.
    Set qry = CurrentDb.QueryDefs(kOcrQueryProcessDataRules)
    qry.Parameters("ProcessId") = aProcessId
    Data.Open_ qry
    qry.Close
    Set qry = Nothing
End Sub
