Attribute VB_Name = "basLibOneNote"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' LibOneNote                                                '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
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
    ByVal aNotebook As String, _
    aOneNote As OneNote.Application, _
    Optional ByVal aLocation As String = "" _
) As String
    '
    ' Creates a OneNote notebook with the given notebook name,
    ' optionally in the specified location. If omitted, OneNote
    ' uses its default location. Returns the notebook name.
    '
    If Len(aLocation) = 0 Then aOneNote.GetSpecialLocation slDefaultNotebookFolder, aLocation
    aLocation = aLocation & "\" & aNotebook
    aOneNote.OpenHierarchy aLocation, vbNullString, oneCreateNotebook, cftNotebook
    oneCreateNotebook = aNotebook
End Function

Public Function oneCreateSection( _
    ByVal aSection As String, _
    ByVal aNotebook As String, _
    aOneNote As OneNote.Application _
) As String
    '
    ' Creates a OneNote notebook section with the given section name,
    ' in the named notebook. Returns the section name.
    '
    Dim notebook As MSXML2.IXMLDOMNode
    
    Set notebook = oneGetNotebook(aNotebook, aOneNote)
    If Not notebook Is Nothing Then
        Dim path As String
        
        path = oneGetNodeAttr(notebook, "path") & "\" & aSection & ".one"
        aOneNote.OpenHierarchy path, vbNullString, oneCreateSection, cftSection
        oneCreateSection = aSection
    End If
    Set notebook = Nothing
End Function

Public Function oneGetDocument( _
    ByVal aXml As String, _
    aOneNote As OneNote.Application _
) As MSXML2.DOMDocument60
    '
    ' Returns a DOMDocument from the given xml string.
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
    aOneNote As OneNote.Application _
) As MSXML2.DOMDocument60
    '
    ' Returns a DOMDocument containing the heirarchy of the given
    ' scope starting at the given node id.
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
    ' Returns the given attribute value of the given node. If
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
    ' Sets the given attribute value of the given node and returns
    ' the given value.
    '
    aNode.Attributes.getNamedItem(aAttribute).Text = aText
    oneSetNodeAttr = aNode.Attributes.getNamedItem(aAttribute).Text
End Function

Public Function oneGetNotebook( _
    ByVal aNotebook As String, _
    aOneNote As OneNote.Application _
) As MSXML2.IXMLDOMNode
    '
    ' Returns the notebook node whose name attribute matches the given
    ' notebook name.
    '
    Set oneGetNotebook = oneGetHeirarchy("", hsNotebooks, aOneNote).documentElement _
    .selectSingleNode("//one:Notebook[@name='" & aNotebook & "']")
End Function

Public Function oneGetNotebooks( _
    aOneNote As OneNote.Application _
) As MSXML2.IXMLDOMNodeList
    '
    ' Returns a collection of all current notebook nodes.
    '
    Set oneGetNotebooks = oneGetHeirarchy("", hsNotebooks, aOneNote).documentElement.selectNodes("//one:Notebook")
End Function

Public Function oneGetSection( _
    aNotebookID As String, _
    aSection As String, _
    aOneNote As OneNote.Application _
) As MSXML2.IXMLDOMNode
    '
    ' Returns the section node whose name attribute matches the given section
    ' and that is a child of the given notebook id.
    '
    Set oneGetSection = oneGetHeirarchy(aNotebookID, hsSections, aOneNote).documentElement _
    .selectSingleNode("//one:Section[@name='" & aSection & "']")
End Function

Public Function oneGetSections( _
    ByVal aNotebookID As String, _
    aOneNote As OneNote.Application _
) As MSXML2.IXMLDOMNodeList
    '
    ' Returns a collection of section nodes that are children of the
    ' given notebook id.
    '
    Set oneGetSections = oneGetHeirarchy(aNotebookID, hsSections, aOneNote).documentElement _
    .selectNodes("//one:Section")
End Function

Public Function oneGetPage( _
    ByVal aSectionId As String, _
    ByVal aPage As String, _
    aOneNote As OneNote.Application _
) As MSXML2.IXMLDOMNode
    '
    ' Returns a page node whose name attribute matches the given name
    ' and that is a child of the given section id.
    '
    Set oneGetPage = oneGetHeirarchy(aSectionId, hsPages, aOneNote).documentElement _
    .selectSingleNode("//one:Page[@name='" & aPage & "']")
End Function

Public Function oneGetPages( _
    ByVal aSectionId As String, _
    aOneNote As OneNote.Application _
) As MSXML2.IXMLDOMNodeList
    '
    ' Returns a collection of page nodes that are children of the
    ' given section id.
    '
    Set oneGetPages = oneGetHeirarchy(aSectionId, hsPages, aOneNote).documentElement _
    .selectNodes("//one:Page")
End Function

Public Function oneGetChildren( _
    ByVal aPageId As String, _
    aOneNote As OneNote.Application, _
    Optional ByVal aNodeType As String = "" _
) As MSXML2.IXMLDOMNodeList
    '
    ' Returns a collection of nodes that are children of the
    ' given page id.
    '
    Set oneGetChildren = oneGetHeirarchy(aPageId, hsChildren, aOneNote).documentElement _
    .selectNodes(aNodeType)
End Function

Public Sub oneDeletePage( _
    ByVal aPageId As String, _
    aOneNote As OneNote.Application, _
    Optional ByVal aPermanetly As Boolean = False _
)
    '
    ' Deletes the page node whose ID attribute matches the
    ' given page id.
    '
    aOneNote.DeleteHierarchy aPageId, , aPermanetly
End Sub

Public Sub oneDeletePages( _
    ByVal aSectionId As String, _
    aOneNote As OneNote.Application, _
    Optional ByVal aPermanetly As Boolean = False _
)
    '
    ' Deletes all page nodes that are children of the section
    ' whose ID attribute matches the given section id.
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
