Attribute VB_Name = "basLibOneNote"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' basLibOneNote                                             '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a library of generic, reusable        '
' objects for dealing with common, repetitive MS OneNote    '
' tasks using single lines of self-documenting code.        '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' None                                                      '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the library has only been tested with     '
' MS ONENOTE 365 (64-bit) implementations.                  '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database
Option Explicit

''''''''''''''''''''
' MS XML Constants '
''''''''''''''''''''

Public Const ONENOTE_SCHEMA As String = "xmlns:one='http://schemas.microsoft.com/office/onenote/2013/onenote'"
Public Const XML_SCHEMA = xs2013

'''''''''''''''''''''
' Library Functions '
'''''''''''''''''''''

Public Function oneCreateNotebook( _
    ByVal aNotebookName As String, _
    aOneNote As onenote.Application, _
    Optional ByVal aLocation As String = "" _
) As String
    '
    ' Creates a OneNote notebook with the given aNotebookName,
    ' optionally in the specified aLocation. If omitted, OneNote
    ' uses its default location. Returns the created notebook name.
    '
    If Len(aLocation) = 0 Then aOneNote.GetSpecialLocation slDefaultNotebookFolder, aLocation
    aLocation = aLocation & "\" & aNotebookName
    aOneNote.OpenHierarchy aLocation, vbNullString, oneCreateNotebook, cftNotebook
    oneCreateNotebook = aNotebookName
End Function

Public Function oneCreateSection( _
    ByVal aSectionName As String, _
    ByVal aNotebookName As String, _
    aOneNote As onenote.Application _
) As String
    '
    ' Creates a OneNote notebook section with the given aSectionName,
    ' in the notebook named aNotebookName. Returns the created section
    ' name.
    '
    Dim notebook As MSXML2.IXMLDOMNode
    
    Set notebook = oneGetNotebook(aNotebookName, aOneNote)
    If Not notebook Is Nothing Then
        Dim path As String
        
        path = oneGetNodeAttr(notebook, "path") & "\" & aSectionName & ".one"
        aOneNote.OpenHierarchy path, vbNullString, oneCreateSection, cftSection
        oneCreateSection = aSectionName
    End If
    Set notebook = Nothing
End Function

Public Function oneGetDocument( _
    ByVal aXml As String, _
    aOneNote As onenote.Application _
) As MSXML2.DOMDocument60
    '
    ' Returns a DOMDocument generated from the given aXml string.
    '
    Dim doc As New MSXML2.DOMDocument60
    
    doc.SetProperty "SelectionNamespaces", ONENOTE_SCHEMA
    If doc.loadXML(aXml) Then
        Set oneGetDocument = doc
    End If
    Set doc = Nothing
End Function

Public Function oneGetHeirarchy( _
    ByVal aNodeId As String, _
    ByVal aScope As HierarchyScope, _
    aOneNote As onenote.Application _
) As MSXML2.DOMDocument60
    '
    ' Returns a DOMDocument containing the heirarchy of the given
    ' aScope starting at the given aNodeId.
    '
    Dim xml As String
    
    aOneNote.GetHierarchy aNodeId, aScope, xml, XML_SCHEMA
    Set oneGetHeirarchy = oneGetDocument(xml, aOneNote)
End Function

Public Function oneGetNodeAttr( _
    aNode As MSXML2.IXMLDOMNode, _
    Optional ByVal aAttribute As String = "ID" _
) As String
    '
    ' Returns the given aAttribute value of the given aNode. If
    ' attribute is omitted, then returns the "ID" attribute.
    '
    oneGetNodeAttr = aNode.Attributes.getNamedItem(aAttribute).Text
End Function

Public Function oneSetNodeAttr( _
    aNode As MSXML2.IXMLDOMNode, _
    ByVal aAttribute As String, _
    ByVal aText As String _
) As String
    '
    ' Sets the given aAttribute value of the given aNode to aText
    ' and returns the value.
    '
    aNode.Attributes.getNamedItem(aAttribute).Text = aText
    oneSetNodeAttr = aNode.Attributes.getNamedItem(aAttribute).Text
End Function

Public Function oneGetNotebook( _
    ByVal aNotebookName As String, _
    aOneNote As onenote.Application _
) As MSXML2.IXMLDOMNode
    '
    ' Returns the notebook node whose name attribute matches the given
    ' aNotebookName.
    '
    Set oneGetNotebook = oneGetHeirarchy("", hsNotebooks, aOneNote).documentElement _
    .selectSingleNode("//one:Notebook[@name='" & aNotebookName & "']")
End Function

Public Function oneGetNotebookId( _
    ByVal aNotebookName As String, _
    aOneNote As onenote.Application _
) As String
    '
    ' Returns the notebook node id whose name attribute matches the given
    ' aNotebookName.
    '
    oneGetNotebookId = oneGetNodeAttr(oneGetHeirarchy("", hsNotebooks, aOneNote).documentElement _
    .selectSingleNode("//one:Notebook[@name='" & aNotebookName & "']"))
End Function

Public Function oneGetNotebooks( _
    aOneNote As onenote.Application _
) As MSXML2.IXMLDOMNodeList
    '
    ' Returns a collection of all current notebook nodes.
    '
    Set oneGetNotebooks = oneGetHeirarchy("", hsNotebooks, aOneNote).documentElement.selectNodes("//one:Notebook")
End Function

Public Function oneGetSection( _
    aNotebookID As String, _
    aSectionName As String, _
    aOneNote As onenote.Application _
) As MSXML2.IXMLDOMNode
    '
    ' Returns the section node whose name attribute matches the given aSectionName
    ' and that is a child of the notebook node with the given aNotebookID.
    '
    Set oneGetSection = oneGetHeirarchy(aNotebookID, hsSections, aOneNote).documentElement _
    .selectSingleNode("//one:Section[@name='" & aSectionName & "']")
End Function

Public Function oneGetSectionId( _
    aNotebookID As String, _
    aSectionName As String, _
    aOneNote As onenote.Application _
) As String
    '
    ' Returns the section node id whose name attribute matches the given aSectionName
    ' and that is a child of the notebook node with the given aNotebookID.
    '
    oneGetSectionId = oneGetNodeAttr(oneGetHeirarchy(aNotebookID, hsSections, aOneNote).documentElement _
    .selectSingleNode("//one:Section[@name='" & aSectionName & "']"))
End Function

Public Function oneGetSections( _
    ByVal aNotebookID As String, _
    aOneNote As onenote.Application _
) As MSXML2.IXMLDOMNodeList
    '
    ' Returns a collection of section nodes that are children of the
    ' notebook node with the given aNotebookID.
    '
    Set oneGetSections = oneGetHeirarchy(aNotebookID, hsSections, aOneNote).documentElement _
    .selectNodes("//one:Section")
End Function

Public Function oneGetPage( _
    ByVal aSectionId As String, _
    ByVal aPageName As String, _
    aOneNote As onenote.Application _
) As MSXML2.IXMLDOMNode
    '
    ' Returns a page node whose name attribute matches the given aPageName
    ' and that is a child of the section node with the given aSectionId.
    '
    Set oneGetPage = oneGetHeirarchy(aSectionId, hsPages, aOneNote).documentElement _
    .selectSingleNode("//one:Page[@name='" & aPageName & "']")
End Function

Public Function oneGetPageId( _
    ByVal aSectionId As String, _
    ByVal aPageName As String, _
    aOneNote As onenote.Application _
) As String
    '
    ' Returns a page node id whose name attribute matches the given aPageName
    ' and that is a child of the section node with the given aSectionId.
    '
    oneGetPageId = oneGetNodeAttr(oneGetHeirarchy(aSectionId, hsPages, aOneNote).documentElement _
    .selectSingleNode("//one:Page[@name='" & aPageName & "']"))
End Function

Public Function oneGetPages( _
    ByVal aSectionId As String, _
    aOneNote As onenote.Application _
) As MSXML2.IXMLDOMNodeList
    '
    ' Returns a collection of page nodes that are children of the
    ' section node with the given aSectionId.
    '
    Set oneGetPages = oneGetHeirarchy(aSectionId, hsPages, aOneNote).documentElement _
    .selectNodes("//one:Page")
End Function

Public Function oneGetChildren( _
    ByVal aPageId As String, _
    aOneNote As onenote.Application, _
    Optional ByVal aNodeType As String = "" _
) As MSXML2.IXMLDOMNodeList
    '
    ' Returns a collection of nodes that are children of the
    ' node with the given aPageId.
    '
    Set oneGetChildren = oneGetHeirarchy(aPageId, hsChildren, aOneNote).documentElement _
    .selectNodes(aNodeType)
End Function

Public Sub oneDeletePage( _
    ByVal aPageId As String, _
    aOneNote As onenote.Application, _
    Optional ByVal aPermanetly As Boolean = False _
)
    '
    ' Deletes the page node whose ID attribute matches the
    ' given aPageId.
    '
    aOneNote.DeleteHierarchy aPageId, , aPermanetly
End Sub

Public Sub oneDeletePages( _
    ByVal aSectionId As String, _
    aOneNote As onenote.Application, _
    Optional ByVal aPermanetly As Boolean = False _
)
    '
    ' Deletes all page nodes that are children of the section
    ' whose ID attribute matches the given aSectionId.
    '
    Dim pages As MSXML2.IXMLDOMNodeList
    Dim page As MSXML2.IXMLDOMNode
    
    On Error Resume Next
    Set pages = oneGetPages(aSectionId, aOneNote)
    If Not pages Is Nothing Then
        For Each page In pages
            oneDeletePage oneGetNodeAttr(page), aOneNote, aPermanetly
        Next
    End If
End Sub
