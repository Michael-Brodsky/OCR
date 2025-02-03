VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBoxExT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' ListBoxExT                                                '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a wrapper class that extends the      '
' capabilities of a ListBox control.                        '
'                                                           '
' Right-Click:                                              '
'   Saves all of the currently selected items so they don't '
'   disappear if a shortcut menu opens.                     '
'                                                           '
' OnExit:                                                   '
'   Deselects all list items.                               '
'                                                           '
' Selected():                                               '
'   Returns a list of selected items each with a list of    '
'   column values.                                          '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' LibVBA                                                    '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the module has only been tested with      '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

'''''''''''''''''''''''
' Types and Constants '
'''''''''''''''''''''''

' Enumerates valid values for the button shift mask in mouse
' and key event handlers.
Public Enum ButtonShiftMask
    bsmShift = 1
    bsmCtrl = 2
    bsmAlt = 4
End Enum

' Enumerates valid values for the ListBox.MultiSelect property.
Public Enum ListMultiSelect
    lmsNone = 0
    lmsSimple = 1
    lmsExtended = 2
End Enum

'''''''''''''''''''
' Private Members '
'''''''''''''''''''

Private WithEvents ListBoxEx As ListBox ' Client control.
Attribute ListBoxEx.VB_VarHelpID = -1
Private itemSelected() As Boolean       ' Current selection status of each client ListItem.

''''''''''''''''''''
' Public Interface '
''''''''''''''''''''

Public Property Set Control( _
    aListBox As ListBox _
)
    '
    ' Sets the client ListBox control.
    '
    Set ListBoxEx = aListBox
End Property

Public Property Get Selected( _
    Optional ByVal aNext As Variant _
) As Variant
    '
    ' Returns a list of selected client ListItems each
    ' with a list of column values, or Empty, if no
    ' ListItems are selected.
    '
    Dim i As Integer, j As Integer, items As Variant
    
    For i = 0 To ListBoxEx.ListCount - 1
        If ListBoxEx.Selected(i) Then
            Dim item As Variant
            
            ArrayResize item, 0
            For j = 0 To ListBoxEx.ColumnCount - 1
                ArrayPushBack item, ListBoxEx.Column(j, i)
            Next
            ArrayPushBack items, item
        End If
    Next
    Selected = items
End Property

Private Sub ListBoxEx_BeforeUpdate( _
    Cancel As Integer _
)
    '
    ' Restores any selected client ListItems that may
    ' have been unselected by a superceding event.
    '
    While ArraySize(itemSelected) > 0
        ListBoxEx.Selected(ArraySize(itemSelected) - 1) = ArrayPopFront(itemSelected)
    Wend
End Sub

Private Sub ListBoxEx_Exit( _
    Cancel As Integer _
)
    '
    ' Deselect all client ListItems.
    '
    Dim i As Integer
    
    For i = 0 To ListBoxEx.ListCount - 1
        ListBoxEx.Selected(i) = False
    Next
End Sub

Private Sub ListBoxEx_MouseDown( _
    Button As Integer, _
    Shift As Integer, _
    x As Single, _
    y As Single _
)
    '
    ' Takes an action according the the button clicked and shift mask.
    '
    Select Case Button
        Case vbKeyLButton:
        
        Case vbKeyRButton:
            If MultiSelect() Then SaveSelected
        Case vbKeyMButton:
        
        Case Else:
    End Select
End Sub

Private Function MultiSelect() As Boolean
    '
    ' Returns TRUE if the client is multi-select enabled,
    ' else returns FALSE.
    '
    MultiSelect = (ListBoxEx.MultiSelect <> lmsNone)
End Function

Private Sub SaveSelected()
    '
    ' Saves all currently selected client ListItems.
    '
    Dim i As Integer
    
    ArrayResize itemSelected, ListBoxEx.ListCount
    For i = 0 To ListBoxEx.ListCount - 1
        itemSelected(i) = ListBoxEx.Selected(i)
    Next
End Sub

