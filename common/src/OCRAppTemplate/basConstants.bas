Attribute VB_Name = "basConstants"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' basConstants                                              '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines global application constants.         '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' (None)                                                    '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the module has only been tested with      '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

''''''''''''''''''''''''''''''''
' Common Application Constants '
''''''''''''''''''''''''''''''''

' These come from the OCRAdmin.accdb template UI and are
' used in any front-end application.

Public Const kOcrAdminConvertErrDefault As Integer = vbIgnore
Public Const kOcrAdminErrSetup As Long = 99
Public Const kOcrAdminErrSettings As Long = 100
Public Const kOcrAdminFormAdmin As String = "Admin"
Public Const kOcrAdminFormAppSettings As String = "frmSettings"
Public Const kOcrAdminFormCatalog As String = "Catalog"
Public Const kOcrAdminFormCatalogDefs As String = "frmCatalogDefs"
Public Const kOcrAdminFormFileTypes As String = "frmSearchFileTypes"
Public Const kOcrAdminFormSettings As String = "frmSettings"
Public Const kOcrAdminFormOcrDataRules As String = "frmOcrDataRules"
Public Const kOcrAdminFormPatternDefs As String = "frmPatternDefs"
Public Const kOcrAdminFormProcesses As String = "frmProcessDefs"
Public Const kOcrAdminFormSearchDefs As String = "frmSearchDefs"
Public Const kOcrAdminMacroSetup As String = "Setup"
Public Const kOcrAdminPropertyAppDir As String = "OCRAppInstallDir"
Public Const kOcrAdminPropertyAppIsDemo As String = "OCRAppIsDemo"
Public Const kOcrAdminPropertyAppOnError As String = "OCRApplicationOnError"
Public Const kOcrAdminPropertyAppTimeout As String = "OCRApplicationTimeout"
Public Const kOcrAdminPropertyAppWindowStyle As String = "OCRApplicationHide"
Public Const kOcrAdminPropertyAppFolderTemp As String = "OCRTempDir"
Public Const kStrCannotBeBlank As String = " cannot be blank"
Public Const kOcrAdminStrFilterAllFiles As String = "All Files,*.*"
Public Const kOcrAdminStrFilterDatabase As String = "Microsoft Access,*.accdb; *.accde ,All Files,*.*"
Public Const kOcrAdminStrSelectDatabase As String = "Select Database"
Public Const kOcrAdminStrSelectSaveToPath As String = "Select Save To Path"
Public Const kOcrAdminStrSelectSearchPath As String = "Select Search Path"
Public Const kOcrAdminStrVersion As String = "(v) 20250225"

'-------------------------------------------------------------------------------------------------------

''''''''''''''''''''''''''''''''''
' Application-Specific Constants '
''''''''''''''''''''''''''''''''''

Public Const kOcrTesseractPropertyImageDensity As String = "OCRImageDensity"
Public Const kOcrTesseractPropertyImageFormat As String = "OCRImageFormat"
Public Const kOcrTesseractPropertyImageResize As String = "OCRImageResize"
Public Const kOcrTesseractPropertyImageRotate As String = "OCRImageRotate"
Public Const kOcrTesseractPropertyImageSharpen As String = "OCRImageSharpen"
Public Const kOcrTesseractPropertyImageTrim As String = "OCRImageTrim"

