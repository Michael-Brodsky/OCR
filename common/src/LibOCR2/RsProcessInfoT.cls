VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RsProcessInfoT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' RsProcessInfoT                                            '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a class that encapsulates a DAO       '
' Recordset holding information about a catalog database    '
' and accessible using dot notation.                        '
'                                                           '
' RsProcessInfoT                                            '
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

Private Data As DAORecordsetT   ' The underlying recordset.

Public Property Get BOF() As Boolean
    BOF = Data.BOF
End Property

Public Property Get CatalogName() As String
    CatalogName = Data.Value("Catalog Name")
End Property

Public Property Get connection() As String
    connection = Nz(Data.Value("Connection"))
End Property

Public Property Get Count() As Long
    Count = Data.Count
End Property

Public Property Get DeleteDuplicates() As Boolean
    DeleteDuplicates = Data.Value("Delete Duplicates")
End Property

Public Property Get Name() As String
    Name = Data.Value("Process Name")
End Property

Public Property Get EOF() As Boolean
    EOF = Data.EOF
End Property

Public Property Get FileTypes() As String
    FileTypes = Data.Value("File Types")
End Property

Public Property Get ID() As Long
    ID = Data.Value("ID")
End Property

Public Property Get SaveToPath() As String
    SaveToPath = Nz(Data.Value("SaveTo Path"))
End Property

Public Property Get SaveToParameterField() As String
    SaveToParameterField = Nz(IIf(Not IsNull(Data.Value("SaveTo Table")), _
    Data.Value("SaveTo Field"), Data.Value("SaveTo Parameter")))
End Property

Public Property Get SaveToProcedure() As String
    SaveToProcedure = Nz(Data.Value("SaveTo Procedure"))
End Property

Public Property Get SaveToTable() As String
    SaveToTable = Nz(Data.Value("SaveTo Table"))
End Property

Public Property Get SearchPath() As String
    SearchPath = Data.Value("Search Path")
End Property

Public Property Get SearchName() As String
    SearchName = Data.Value("Search Name")
End Property

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
    Set qry = CurrentDb.QueryDefs(kOcrQueryProcessInfo)
    qry.Parameters("ProcessId") = aProcessId
    Data.Open_ qry
    qry.Close
    Set qry = Nothing
End Sub


