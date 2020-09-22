Attribute VB_Name = "ModGraph"
Option Explicit

Public GraphValues() As Long
Public ValuesToGraph() As Long
Public GraphNumber As Long

Public Sub AddtoGraphArray(ControlPos As Double, _
                            MassPos As Double, _
                            CurError As Double, _
                            PropResult As Double, _
                            DerResult As Double, _
                            IntResult As Double, _
                            CurVelocity As Double, _
                            Force As Double)
   
    Dim UpperBound As Integer
    UpperBound = UBound(GraphValues, 2)
   
    If UpperBound >= 5000 Then
        frmMain.tmrGraph.Enabled = False
    End If
   
    GraphValues(0, UpperBound) = ControlPos
    GraphValues(1, UpperBound) = MassPos
    GraphValues(2, UpperBound) = CurError
    GraphValues(3, UpperBound) = PropResult
    GraphValues(4, UpperBound) = DerResult
    GraphValues(5, UpperBound) = IntResult
    GraphValues(6, UpperBound) = CurVelocity
    GraphValues(7, UpperBound) = Force
    
    'debug.print "Force: " & Force
    ReDim Preserve GraphValues(7, UpperBound + 1)
End Sub
