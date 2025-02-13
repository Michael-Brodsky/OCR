VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' Form_Admin                                                '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines types and objects that handle events  '
' and procedures for the Admin form.                        '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' basConstants                                              '
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

'Private Const kProcessListCount As Integer = 6 ' The number of visible rows in the processes list box.
Private Const kOcrProcedure As String = "tessGenerate"  ' OCR conversion function name.
Private Const kAppCallbackProcedure As String = "OcrTesseractCallback"  ' The application callback function.
Private Const kStrFilesProcessed As String = " file(s) successfully cataloged"
Private Const kStrFilesCataloged As String = " total file(s) cataloged"
Private Const kStrFilesFound As String = " file(s) found"
Private Const kStrFolderCatalogNotFound As String = "Catalog folder not found"
Private Const kStrFolderSearchNotFound As String = "Search folder not found"
Private Const kStrFormatTimeElapsed As String = "#.00"
Private Const kStrSettingsError As String = _
    "Bad or missing settings. Open Settings " & _
    "and ensure that all required settings are valid. " & _
    "Click the green check button to save."
Private Const kStrSetupRequired As String = _
    "You must run Setup before using the application."
    
''''''''''''''''''''''''
' Module-Level Objects '
''''''''''''''''''''''''

'Public Event OnExit(Cancel As Integer)  ' Clear selected items in processes list on mouse exit.

Private listProcessesEx As ListBoxExT   ' Listbox that allows drag and drop process ordering.
Private boolClosing As Boolean          ' Flag indicating that the application is closing.

''''''''''''''''''
' Event Handlers '
''''''''''''''''''

Private Sub cmdAppSettings_Click()
    On Error GoTo Catch
    SettingsOpen
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub cmdProcessUp_Click()
    On Error GoTo Catch
    ProcessListReorder -1
    Exit Sub
    
Catch:
    DoCmd.Hourglass False
    ErrMessage
End Sub

Private Sub cmdReset_Click()
    On Error GoTo Catch
    ProcessReset
    Exit Sub
    
Catch:
    DoCmd.Hourglass False
    ErrMessage
End Sub

Private Sub cmdSettings_Click()
    On Error GoTo Catch
    SettingsOpen
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub cmdStart_Click()
    On Error GoTo Catch
    ProcessStart
    
Catch:
    UISetState False
    If Err.Number <> 0 Then ErrMessage
End Sub

Private Sub cmdStop_Click()
    On Error GoTo Catch
    DoCmd.Hourglass True    ' Visial effect while waiting for the API call to return.
    ProcessStop
    Exit Sub
    
Catch:
    DoCmd.Hourglass False
    ErrMessage
End Sub

Private Sub Form_Close()
    On Error Resume Next
    boolClosing = True
    ProcessStop
End Sub

Private Sub Form_Load()
    On Error GoTo Catch
    MeLoad
    Exit Sub

Catch:
    ErrMessage
End Sub

Private Sub Form_Open(Cancel As Integer)
    On Error Resume Next
    DoCmd.Hourglass False   ' You can never have enough of these.
    Cancel = Not PropertyExists(CurrentDb, kOcrAdminPropertyAppDir) ' Prevents app from starting before setup complete.
    If Cancel Then MsgBox kStrSetupRequired
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    DoCmd.Hourglass False
    Cancel = Not MeUnload()
End Sub

Private Sub listProcesses_DblClick(Cancel As Integer)
    On Error GoTo Catch
    SettingsOpen
    Exit Sub

Catch:
    ErrMessage
End Sub

Private Sub listProcesses_Exit(Cancel As Integer)
    On Error GoTo Catch
    Exit Sub

Catch:
    ErrMessage
End Sub

'''''''''''''''''''''
' Private Interface '
'''''''''''''''''''''

Private Sub MeLoad()
    '
    ' Initializes our form.
    '
    On Error Resume Next
    Set listProcessesEx = NewListBoxExT(Me.listProcesses)
    AppInit ' Do any application-specific initialization.
    ProcessReset
End Sub

Private Function MeUnload() As Boolean
    '
    ' Releases the custom ListBox resources.
    '
    If AppUnLoad() Then
        Set listProcessesEx = Nothing
        MeUnload = True
    End If
End Function

Private Sub ProcessListReorder( _
    ByVal aDirection As Integer _
)
    Dim i As Integer, n As Integer, idI As Long, idN As Long, ordI As Long, ordN As Long
    
    If Me.listProcesses.ListCount = 0 Then Exit Sub
    For i = 0 To Me.listProcesses.ListCount - 1
        If Me.listProcesses.Selected(i) Then Exit For
    Next
    n = i + aDirection
    idI = Me.listProcesses.Column(0, i)
    idN = Me.listProcesses.Column(0, Constrain(n, 0, Me.listProcesses.ListCount - 1))
    If idI <> idN Then
        ordI = DLookup("[Ordinal]", "[~tblProcesses]", "[ID] = " & Str(idI))
        ordN = DLookup("[Ordinal]", "[~tblProcesses]", "[ID] = " & Str(idN))
        Swap ordI, ordN
        CurrentDb.Execute "UPDATE [~tblProcesses] SET [Ordinal] = " & Str(ordI) & " WHERE [ID] = " & Str(idI) & ";"
        CurrentDb.Execute "UPDATE [~tblProcesses] SET [Ordinal] = " & Str(ordN) & " WHERE [ID] = " & Str(idN) & ";"
        Me.listProcesses.Requery
        Me.listProcesses.Selected(n) = False
        Me.listProcesses.Selected(n - aDirection) = True
    End If
End Sub

Private Sub ProcessReset()
    '
    ' Resets the application into a known state.
    '
    DoCmd.Hourglass False
    AppOcrReset ' Do any application-specific reset procedures.
    ProcessResultsShow
End Sub

Private Sub ProcessResultsShow( _
    Optional ByVal aResults As String = "" _
)
    '
    ' Shows cataloging process results in our form's controls.
    '
    On Error Resume Next    ' We don't care about errors as this sub might get called during the form unload event.
    If aResults <> "" Then
        Me.listProcesses.Requery
        TextBoxAppend Me.txtMessages, aResults
    Else
        Me.txtMessages = ""
        TextBoxAppend Me.txtMessages, CurrentProject.Name & " " & kOcrAdminStrVersion
        Me.txtMessages.Requery
    End If
    Err.Clear
End Sub

Private Sub ProcessStart()
    '
    ' Starts the cataloging process. Generates the common OCR start function
    ' arguments from the current common settings and/or defaults and calls
    ' the application-specific function with those arguments.
    '
    Dim result As Long, start As Double
    
    On Error GoTo Catch
    start = Timer()
    result = AppProcessStart( _
        PropertyGet(CurrentDb, kOcrAdminPropertyAppFolderTemp), _
        PropertyGet(CurrentDb, kOcrAdminPropertyAppTimeout), _
        PropertyGet(CurrentDb, kOcrAdminPropertyAppOnError), _
        IIf(PropertyGet(CurrentDb, kOcrAdminPropertyAppWindowStyle) = True, vbHide, vbNormalFocus) _
    )
    ProcessResultsShow CStr(result) & " " & kStrFilesCataloged & " (" & _
    Format(Timer() - start, kStrFormatTimeElapsed) & "s)"
    Exit Sub
    
Catch:
    ErrMessage
End Sub

Private Sub ProcessStop( _
    Optional ByVal aKill As Boolean = False _
)
    '
    ' Stops the cataloging process.
    '
    AppOcrStop  ' Perform any application-specific stop procedures.
End Sub

Private Sub SettingsOpen()
    '
    ' Open the settings form and display any relevant
    ' changes in our form's controls.
    '
    FormOpenDialog kOcrAdminFormSettings, dlgFind, "[ID]=" & Nz(Me.listProcesses, -1)
    Me.listProcesses.Requery
    'DoCmd.OpenForm kOcrAdminFormSettings, , , , , acDialog, ParamToCsv(dlgFind, "[ID]=" & Str(Nz(Me.listProcesses, -1)))
    'If CurrentProject.AllForms(kOcrAdminFormSettings).IsLoaded Then Me.listProcesses.Requery
    'DoCmd.Close acForm, kOcrAdminFormSettings, acSaveNo
End Sub

Private Sub UISetState( _
    ByVal aRun As Boolean _
)
    '
    ' Sets the state of our form's controls according to
    ' whether the process is running or not.
    '
    If Not aRun Then DoCmd.Hourglass False
    AppUISetState aRun      ' Do any application-specific UI tasks.
    If Not boolClosing Then ' If we're not closing, set the Start, Stop and Reset button states.
        Me.cmdStart.Enabled = Not aRun
        Me.cmdStop.Enabled = Not Me.cmdStart.Enabled
        Me.cmdReset.Enabled = Me.cmdStart.Enabled
    End If
    DoEvents                ' This ensures a visual transition.
    boolClosing = False     ' Clear the flag set in the Form_Close event.
End Sub

' ===================================================================================================================

'''''''''''''''''''''''''''''''''''
' Application-Specific Procedures '
'''''''''''''''''''''''''''''''''''

Private Function AppInit( _
    ParamArray aArgs() As Variant _
)
    '
    ' Application-specific initialization procedure.
    '
End Function

Private Function AppProcessStart( _
    Optional ByVal aFolderTemp As String = "", _
    Optional ByVal aTimeout As Long = WAIT_INFINITE, _
    Optional ByVal aOnError As Integer = vbIgnore, _
    Optional ByVal aShowCmd As VbAppWinStyle = vbNormalFocus _
) As Long
    '
    ' Generates any application-specific arguments and calls
    ' the application's start function.
    '
    Dim img As TessImageT, ocr As OcrConvertT, cb As CallbackT
    
    ' Tesseract image conversion params.
    Set img = NewTessImageT( _
        PropertyLoad(kOcrTesseractPropertyImageDensity, kTessImageDensityDefault), _
        PropertyLoad(kOcrTesseractPropertyImageFormat, kTessImageFormatDefault), _
        PropertyLoad(kOcrTesseractPropertyImageResize), _
        PropertyLoad(kOcrTesseractPropertyImageRotate), _
        PropertyLoad(kOcrTesseractPropertyImageSharpen), _
        PropertyLoad(kOcrTesseractPropertyImageTrim, False) _
    )
    ' Ocr conversion params.
    Set ocr = NewOcrConvertT( _
        NewCallbackT(kOcrProcedure, VbMethod), img, aFolderTemp, aTimeout, aShowCmd _
    )
    ' Application callback params.
    Set cb = NewCallbackT( _
        kAppCallbackProcedure, VbMethod, Form_Admin _
    )
    ' Start the process.
    UISetState True
    AppProcessStart = ocrStart(ocr, cb, aOnError)
    ' Cleanup.
    Set img = Nothing
    Set ocr = Nothing
    Set cb = Nothing
End Function

Private Sub AppOcrReset( _
    ParamArray aArgs() As Variant _
)
    '
    ' Generates any application-specific arguments and calls
    ' the application's reset function.
    '
    ocrReset
End Sub

Private Sub AppOcrStop( _
    ParamArray aArgs() As Variant _
)
    '
    ' Generates any application-specific arguments and calls
    ' the application's stop function.
    '
    ocrStop
End Sub

Private Function AppUnLoad( _
    ParamArray aArgs() As Variant _
) As Boolean
    '
    ' Performs any application-specific tasks prior to
    ' closing this window. Returns "Cancel"
    '
    AppUnLoad = True
End Function

Private Sub AppUISetState( _
    ParamArray aArgs() As Variant _
)
    '
    ' Sets the application-specific user interface state.
    '
End Sub

Public Sub OcrTesseractCallback( _
    aArgs As Variant _
)
    Dim callback_t As OcrCallbackType, msg As String, i As Integer
    
    On Error Resume Next    ' We don't care about errors since we're just displaying info msgs.
    callback_t = aArgs(0)
    Select Case callback_t
        Case octBegin:
            TextBoxAppend Me.txtMessages, aArgs(1)
        Case octSearch:
            TextBoxAppend Me.txtMessages, aArgs(2) & ": " & aArgs(3) & kStrFilesFound
        Case octFile:
            msg = IIf(aArgs(2) <> "", " (" & aArgs(2) & ") ", " ") & aArgs(3)
            TextBoxAppend Me.txtMessages, aArgs(1) & msg
        Case octEnd:
            TextBoxAppend Me.txtMessages, aArgs(1) & ": " & aArgs(2) & kStrFilesProcessed
        Case octMessage:
            ArrayPopFront aArgs
            TextBoxAppend Me.txtMessages, Join(aArgs, " ")
        Case Else:
    End Select
    
Finally:
    Err.Clear
End Sub

Private Sub listProcesses_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And acShiftMask) <> 0 Then
        Select Case KeyCode
            Case 40:
                ProcessListReorder 1
            Case 38:
                ProcessListReorder -1
            Case Else:
        End Select
    End If
End Sub
