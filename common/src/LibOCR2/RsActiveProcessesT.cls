VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RsActiveProcessesT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' RsActiveProcessesT                                                '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a class that encapsulates a DAO       '
' Recordset holding information about active cataloging     '
' processes and accessible using dot notation.              '
'                                                           '
' DEPENDENCIES                                              '
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

Public Property Get Active() As Long
    Active = Data.Value("Active")
End Property

Public Property Get CatalogId() As Long
    CatalogId = Data.Value("Catalog ID")
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

Public Property Get ID() As Long
    ID = Data.Value("ID")
End Property

Public Property Get LastUpdate() As Variant
    LastUpdate = Data.Value("Last Update")
End Property

Public Property Let LastUpdate(ByVal aDate As Variant)
    Data.Value("Last Update") = aDate
End Property

Public Property Get SearchId() As Long
    SearchId = Data.Value("Search ID")
End Property

Public Sub MoveNext()
    Data.MoveNext
End Sub

Public Sub Close_()
    If Not Data Is Nothing Then Data.Close_
    Set Data = Nothing
End Sub

Public Sub Open_()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    
    Close_
    Set db = CurrentDb
    ' TODO: The qdf needs to be replaced with a defined query.
    Set qdf = CurrentDb.QueryDefs(kOcrQueryActiveProcesses)
    Set Data = NewDAORecordsetT()
    Data.Open_ qdf, dbOpenDynaset
    Set qdf = Nothing
    Set db = Nothing
End Sub

