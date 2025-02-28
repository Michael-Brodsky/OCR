VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OcrConvertT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' OcrParamsT                                                '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a class that encapsulates the OCR     '
' engine file conversion function call and parameters.      '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' CallbackT, OcrParamsT                                     '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the library has only been tested with     '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

Implements OcrParamsT

Private m_ocrparams As OcrParamsT   ' The ocr engine file conversion parameters.
Private m_callback As CallbackT     ' The ocr engine conversion function call.

Private Property Get OcrParamsT_FilePath() As String
    OcrParamsT_FilePath = m_ocrparams.Filepath
End Property

Private Property Let OcrParamsT_FilePath(ByVal aFilePath As String)
    m_ocrparams.Filepath = aFilePath
End Property

Private Property Get OcrParamsT_OcrParams() As Variant
    If IsObject(m_ocrparams.OcrParams) Then
        Set OcrParamsT_OcrParams = m_ocrparams.OcrParams
    Else
        OcrParamsT_OcrParams = m_ocrparams.OcrParams
    End If
End Property

Private Property Let OcrParamsT_OcrParams(aOcrParams As Variant)
    m_ocrparams.OcrParams = aOcrParams
End Property

Private Property Set OcrParamsT_OcrParams(aOcrParams As Variant)
    Set m_ocrparams.OcrParams = aOcrParams
End Property

Private Property Let OcrParamsT_Timeout(ByVal aTimeout As Long)

End Property

Private Property Get OcrParamsT_WindowStyle() As VbAppWinStyle
    OcrParamsT_WindowStyle = m_ocrparams.WindowStyle
End Property

Private Property Let OcrParamsT_WindowStyle(ByVal aWindowStyle As VbAppWinStyle)
    m_ocrparams.WindowStyle = aWindowStyle
End Property

Private Property Get OcrParamsT_Timeout() As Long
    OcrParamsT_Timeout = m_ocrparams.Timeout
End Property

Private Property Get Callback() As CallbackT
    Set Callback = m_callback
End Property

Private Property Set Callback(ByVal aCallback As CallbackT)
    Set m_callback = aCallback
End Property

Public Function Exec(ByVal aFilePath As String) As Variant
    OcrParamsT_FilePath = aFilePath
    If Not m_callback.target Is Nothing Then
        Exec = CallByName(m_callback.target, m_callback.Procedure, m_callback.Method, m_ocrparams)
    Else
        Exec = Application.Run(m_callback.Procedure, m_ocrparams)
    End If
End Function

Public Sub Init( _
    aCallback As CallbackT, _
    aOcrParams As OcrParamsT _
)
    Set m_callback = aCallback
    Set m_ocrparams = aOcrParams
End Sub

