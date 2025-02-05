Attribute VB_Name = "basTessOcrEngine"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' basOcrTesseract                                           '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines an ocr engine based on the OCR-       '
' Tesseract and ImageMagick applications. Tesseract         '
' generates ocr text files from image files and ImageMagick '
' converts certain image files into a format readable by    '
' Tesseract.                                                '
'                                                           '
' Tesseract is a free, open source optical character        '
' recognition engine, originally developed by               '
' Hewlett-Packard in the 1980s and sponsored by Google in   '
' 2006. It is considered one of the most accurate open-     '
' source OCR engines available.                             '
'                                                           '
' Tesseract only supports certain image file formats as     '
' input. To support Adobe "pdf" files, the API uses the     '
' the ImageMagick open-source software suite to convert     '
' "pdf" files into a format readable by Tesseract.          '
' ImageMagick relies on the GhostScript interpreter for     '
' its capabilities.                                         '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' LibOCR, TessImageT                                        '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the API has only been tested with         '
' MS OFFICE 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

'''''''''''''''''''''''
' Types and Constants '
'''''''''''''''''''''''

Public Const kTessErrFileToText = 11000                     ' OCR output file error.
Public Const kTessErrFileConvert = 11001                    ' Input file conversion error.
Public Const kTessOcrOutputTypeDefault As String = "txt"    ' Tesseract default output file format.
Public Const kTessImageFormatDefault As String = "png"      ' ImageMagick default conversion format.
Public Const kTessImageDensityDefault As Integer = 300      ' ImageMagick default image density in dpi.

''''''''''''''''''''
' Public Interface '
''''''''''''''''''''

Public Function tessConvert( _
    ByVal aFilePath As String, _
    ByVal aFolderTemp As String, _
    aImage As TessImageT, _
    Optional ByVal aTimeout As Long = WAIT_INFINITE, _
    Optional ByVal aShowCmd As VbAppWinStyle = vbNormalFocus _
) As String
    '
    ' Converts a file into a type readable by Tesseract.
    ' The file is created by ImageMagick in the given
    ' temp folder.
    '
    Dim Params As String, Options As String, hInstance As Long
    Dim fmt As String
    Dim dense As Integer
    
    ' Generate command line from the given file and image parameters.
    fmt = aImage.Format
    If fmt = "" Then fmt = kTessImageFormatDefault
    dense = aImage.Density
    If dense = 0 Then dense = kTessImageDensityDefault
    Options = " "
    If aImage.Trim Then Options = Options & "-trim "
    If aImage.Resize <> "" Then Options = Options & "-resize " & aImage.Resize & " "
    If aImage.Rotate <> 0 Then Options = Options & "-rotate " & CStr(aImage.Rotate) & " "
    If aImage.Sharpen <> "" Then Options = Options & "-sharpen " & aImage.Sharpen & " "
    tessConvert = PathBuild(aFolderTemp, FileBaseName(aFilePath) & "." & fmt)
    Params = " -density " & CStr(dense) & " " & StringQuote(aFilePath) & Options & StringQuote(tessConvert)
    ' Execute the command and wait for it to finish.
    hInstance = CLng(ShellEx(StringQuote(tessPathExeConvert()), aShowCmd, "open", Params, aFolderTemp, 0, aTimeout))
    If hInstance <> 0 Then Err.Raise UsrErr(kTessErrFileConvert), _
    CurrentProject.Name, "File conversion error " & CStr(hInstance)
End Function

Public Function tessGenerate( _
    aArgs As Variant _
) As String
    '
    ' Generates an output ocr text file from an input source file.
    '
    Dim fin As String, fldr As String, img As TessImageT, timeout As Long, show As VbAppWinStyle, fout As String
    
    fin = aArgs(0)
    Set img = aArgs(1)(0)
    fldr = aArgs(1)(1)
    timeout = aArgs(1)(2)
    show = aArgs(1)(3)
    fout = fin
    ' If necessary, convert the source file to a format readable by Tesseract.
    If Not tessFileReadable(fin) Then _
    fout = tessConvert(fin, fldr, img, timeout, show)
    ' Run Tesseract to generate the OCR text file from the readable file.
    tessGenerate = tessFileToText(fout, fldr, timeout, , show)
    If fout <> fin Then Kill fout
End Function

Public Function tessPathExeConvert() As String
    '
    ' Returns the full path to the ImageMagick executable.
    '
    Const kRegKey As String = "HKLM\SOFTWARE\ImageMagick\Current\BinPath"
    Const kExeName As String = "magick.exe"
    Dim fso As Object
    Dim binPath As String
    
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    binPath = CreateObject("WScript.Shell").RegRead(kRegKey)
    tessPathExeConvert = fso.buildpath(binPath, kExeName)
    Set fso = Nothing
End Function

Public Function tessPathExeOcr() As String
    '
    ' Returns the full path to the Tesseract executable.
    '
    Const kExeName As String = "tesseract.exe"
    Const kRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Tesseract-OCR\InstallDir"
    Dim fso As Object
    Dim binPath As String
    
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    binPath = CreateObject("WScript.Shell").RegRead(kRegKey)
    tessPathExeOcr = fso.buildpath(binPath, kExeName)
    Set fso = Nothing
End Function

Public Function tessPathExeGs() As String
    '
    ' Returns the full path to the GhostScript dll executable.
    '
    Const HKLM = &H80000002
    Const kRegKey = "SOFTWARE\GPL Ghostscript"
    Const kKeyVal = "GS_DLL"
    Dim reg As Object, subkeys() As Variant

    On Error Resume Next
    Set reg = GetObject("winmgmts://./root/default:StdRegProv")
    If reg.EnumKey(HKLM, kRegKey, subkeys) = 0 Then
        Dim gsLib As String, sk As Variant

        For Each sk In subkeys
            If reg.GetStringValue(HKLM, kRegKey & "\" & sk, kKeyVal, gsLib) = 0 Then
                tessPathExeGs = CreateObject("Scripting.FileSystemObject").GetParentFolderName(gsLib)
                If tessPathExeGs <> "" Then Exit For
            End If
        Next
    End If
    Set reg = Nothing
End Function

'''''''''''''''''''''
' Private Interface '
'''''''''''''''''''''

Private Function tessFileToText( _
    ByVal aFilePath As String, _
    ByVal aFolderTemp As String, _
    Optional ByVal aTimeout As Long = WAIT_INFINITE, _
    Optional ByVal aOut As String = kTessOcrOutputTypeDefault, _
    Optional ByVal aShowCmd As VbAppWinStyle = vbNormalFocus _
) As String
    '
    ' Runs Tesseract to generate OCR text from the given file.
    ' The OCR text file is created in the given folder.
    '
    Dim Params As String, hInstance As Long, fout As String
    
    If FileExtension(aFilePath) = kTessOcrOutputTypeDefault Then
        ' If the file is already in the ocr output format, just return the file name.
        tessFileToText = aFilePath
    Else
        ' Generate the Tesseract command line from the given parameters.
        fout = FileBaseName(aFilePath)
        tessFileToText = PathBuild(aFolderTemp, fout & "." & aOut)
        Params = " " & StringQuote(aFilePath) & " " & fout
        ' Execute the command line and wait for the command to finish.
        hInstance = CLng(ShellEx(StringQuote(tessPathExeOcr()), aShowCmd, "open", Params, aFolderTemp, 0, aTimeout))
        If hInstance <> 0 Then Err.Raise UsrErr(kTessErrFileToText), _
        CurrentProject.Name, "OCR text File creation error " & CStr(hInstance)
    End If
End Function

Private Function tessFileExtension( _
    ByVal aFile As String, _
    aExtensions As DAO.Recordset _
) As Boolean
    '
    ' Returns TRUE if the given file's extension is found in the list,
    ' else returns FALSE.
    '
    aExtensions.FindFirst "[File Extension] Like '*" & FileExtension(aFile) & "*'"
    tessFileExtension = Not aExtensions.NoMatch
End Function

Private Function tessFileConvertible( _
    ByVal aFile As String _
) As Boolean
    '
    ' Returns TRUE if the given file's extension is found in the list
    ' of convertible file types, else returns FALSE.
    '
    tessFileConvertible = tessFileExtension(aFile, CurrentDb.OpenRecordset("~tblTessConvertibleFiles", dbOpenDynaset))
End Function

Private Function tessFileReadable( _
    ByVal aFile As String _
) As Boolean
    '
    ' Returns TRUE if the given file's extension is found in the list
    ' of readable file types, else returns FALSE.
    '
    tessFileReadable = tessFileExtension(aFile, CurrentDb.OpenRecordset("~tblTessReadableFiles", dbOpenDynaset))
End Function

''''''''''''''''''''''''''''''
' basClassFactory Extensions '
''''''''''''''''''''''''''''''

Public Function NewTessImageT( _
    Optional aDensity As Variant, _
    Optional aFormat As Variant, _
    Optional aResize As Variant, _
    Optional aRotate As Variant, _
    Optional aSharpen As Variant, _
    Optional aTrim As Variant _
) As TessImageT
    Dim obj As New TessImageT
    
    If Not IsMissing(aDensity) Then obj.Density = aDensity
    If Not IsMissing(aFormat) Then obj.Format = aFormat
    If Not IsMissing(aResize) Then obj.Resize = aResize
    If Not IsMissing(aRotate) Then obj.Rotate = aRotate
    If Not IsMissing(aSharpen) Then obj.Sharpen = aSharpen
    If Not IsMissing(aTrim) Then obj.Trim = aTrim
    Set NewTessImageT = obj
    Set obj = Nothing
End Function


