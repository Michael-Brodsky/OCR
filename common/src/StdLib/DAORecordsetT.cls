VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DAORecordsetT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' DAORecordsetT                                             '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a class that encapsulates a DAO.      '
' Recordset accessible using dot notation. DAORecordsetT is '
' intended to encapsulate forward-only or or snapshot type  '
' recordsets and thus has limited access methods.           '                                               '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' LibVBA                                                    '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the library has only been tested with     '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

'''''''''''''''''''
' Private Members '
'''''''''''''''''''

Private rst As DAO.Recordset    ' The underlying recordset.

''''''''''''''''''''
' Public Interface '
''''''''''''''''''''

Public Property Get BOF() As Boolean
    BOF = rst.BOF
End Property

Public Property Get count() As Long
    count = RecordsetCount(rst)
End Property

Public Property Get EOF() As Boolean
    EOF = rst.EOF
End Property

Public Property Get Value( _
    ByVal aFieldName As String _
) As Variant
    Value = rst.Fields(aFieldName).Value
End Property

Public Property Let Value( _
    ByVal aFieldName As String, _
    aValue As Variant _
)
    rst.Edit
    rst.Fields(aFieldName) = aValue
    rst.Update
End Property

Public Sub Close_()
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
End Sub

Public Sub MoveFirst()
    rst.MoveFirst
End Sub

Public Sub MoveNext()
    rst.MoveNext
End Sub

Public Sub Open_( _
    aQueryDef As DAO.QueryDef, _
    Optional ByVal aType As RecordsetTypeEnum = dbOpenSnapshot _
)
    Close_
    Set rst = aQueryDef.OpenRecordset(aType)
End Sub
