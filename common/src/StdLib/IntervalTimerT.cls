VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IntervalTimerT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' IntervalTimerT                                            '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a an interval timer class used to     '
' track elapsed time intervals and timeouts.                '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' (None)                                                    '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the library has only been tested with     '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

Private timeStart As Double ' The current interval start time.

Public Property Get Elapsed() As Double
    '
    ' Returs the current interval's elapsed time, in seconds.
    '
    Elapsed = Timer() - timeStart
End Property

Public Function Expired( _
    ByVal aTimeout As Double _
) As Boolean
    '
    ' Returns TRUE if the current elapsed time is greater than
    ' the given timeout period in seconds, else returns FALSE.
    '
    Expired = (Elapsed > aTimeout)
End Function

Public Sub Reset()
    '
    ' Resets the current interval.
    '
    timeStart = 0
End Sub

Public Sub Start()
    '
    ' Starts a new interval.
    '
    timeStart = Timer()
End Sub

