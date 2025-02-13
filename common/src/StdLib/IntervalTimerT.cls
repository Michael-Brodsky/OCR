VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IntervalTimerT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Binary
Option Explicit

Private timeStart As Double

Public Function Elapsed() As Double
    Elapsed = Timer() - timeStart
End Function

Public Function Expired( _
    ByVal aTimeout As Double _
) As Boolean
    Expired = Elapsed >= aTimeout
End Function

Public Sub Reset()
    timeStart = 0
End Sub

Public Sub Start()
    timeStart = Timer()
End Sub

