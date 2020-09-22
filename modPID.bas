Attribute VB_Name = "modPID"
Option Explicit
Dim MaxPower As Double
Dim ConPos As Double
Dim MassPos As Double
Dim Error As Double
Dim Mass As Double
Dim PropGain As Double
Dim IntGain As Double
Dim DerGain As Double
Dim Force As Double
Dim Accel As Double
Dim PropResult As Double
Dim IntResult As Double
Dim DerResult As Double
Dim ForcePercent As Double
Dim MaxAlloableFinalError As Double
Dim SignChange  As Boolean
Dim OldError As Double

Public Type PIDReturn
    PowerPercent As Double
    PropRet As Double
    IntRet As Double
    DerRet As Double
End Type


Private Function CalcPID(CurError As Double, PropGain As Double, IntGain As Double, DerGain As Double, DeadBand As Double) As PIDReturn
    

End Function



