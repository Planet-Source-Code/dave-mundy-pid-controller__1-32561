VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'************* clsPID ***************
'Proportional integral derivative controller
'Author:  Dave Mundy
'Date: March 2002
'************************************
Option Explicit

'local variable(s) to hold property value(s)
Private mvarPropGain As Double 'local copy
Private mvarDerGain As Double 'local copy
Private mvarIntGain As Double 'local copy
Private mvarDeadBand As Double 'local copy
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent CalcUpdates[(arg1, arg2, ... , argn)]
Public Event CalcUpdates(PropResult As Double, IntResult As Double, DerResult As Double, ForcePercent As Double)




Public Property Let DeadBand(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DeadBand = 5
    mvarDeadBand = vData
End Property

Public Property Get DeadBand() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DeadBand
    DeadBand = mvarDeadBand
End Property

Public Property Let IntGain(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IntGain = 5
    mvarIntGain = vData
End Property

Public Property Get IntGain() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IntGain
    IntGain = mvarIntGain
End Property

Public Property Let DerGain(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DerGain = 5
    mvarDerGain = vData
End Property

Public Property Get DerGain() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DerGain
    DerGain = mvarDerGain
End Property

Public Property Let PropGain(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PropGain = 5
    mvarPropGain = vData
End Property

Public Property Get PropGain() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PropGain
    PropGain = mvarPropGain
End Property

Public Function CalcPID(CurError As Double) As Double
    Dim PidResults As PIDReturn
    Dim PropRet As Double
    Dim IntResult As Double
    Dim DerResult As Double
    Dim ForcePercent As Double
    
    PropRet = (CurError * PropGain)
    IntResult = Integral(CurError, 0.001, , False)
    DerResult = Deravative(CurError, 0.001)
    


    
    ForcePercent = ((PropRet + DerResult * DerGain + IntResult * IntGain)) / _
    ((PropGain + (IntGain + DerGain)) * 3)
    
    RaiseEvent CalcUpdates(PropRet, IntResult, DerResult, ForcePercent)
    
    If ForcePercent > 1 Or ForcePercent < -1 Then
        CalcPID = ForcePercent / Abs(ForcePercent)
    Else
        CalcPID = ForcePercent
    End If
    
    
End Function

Private Function Deravative(Value As Double, Optional DeltaTime As Double = -1, Optional AbsTime As Double = -1, Optional Initialize As Boolean = False) As Double
    Static LastValue As Double
    Static LastAbsTime As Double
    Dim TimeDiff As Double
    
    If DeltaTime = -1 Then
        TimeDiff = AbsTime - LastAbsTime
    ElseIf AbsTime = -1 Then
        TimeDiff = DeltaTime
    Else
        MsgBox "Either a DeltaTime or an AbsTime argument must be supplied", vbOKOnly, "Deravative Error"
        Exit Function
    End If
    If TimeDiff = 0 Then
        MsgBox "Time Difference Cannot Be Zero, Exiting Deravative Calculation", vbOKOnly, "Deravative Error"
        Exit Function
    End If

    Deravative = (Value - LastValue) / TimeDiff
    
    LastAbsTime = AbsTime
    LastValue = Value
        
End Function

Private Function Integral(Value As Double, Optional DeltaTime As Double = -1, Optional AbsTime As Double = -1, Optional Reset As Boolean = False) As Double
    Static LastValue As Double
    Static LastAbsTime As Double
    Static ValueCarry As Double
    Dim TimeDiff As Double
    
    If Reset = True Then
        ValueCarry = 0
        Exit Function
    End If
    
    If DeltaTime = -1 Then
        TimeDiff = AbsTime - LastAbsTime
    ElseIf AbsTime = -1 Then
        TimeDiff = DeltaTime
    Else
        MsgBox "Either a DeltaTime or an AbsTime argument must be supplied", vbOKOnly, "Integral Error"
        Exit Function
    End If
    
    If TimeDiff = 0 Then
        MsgBox "Time Difference Cannot Be Zero, Exiting Integral Calculation", vbOKOnly, "Integral Error"
        Exit Function
    End If
    
    Integral = (0.5 * (TimeDiff * (Value - LastValue))) + (LastValue * TimeDiff) + ValueCarry
    ValueCarry = Integral
    LastAbsTime = AbsTime
    LastValue = Value
     If TimeDiff = 0 Then
        MsgBox "Time Difference Cannot Be Zero, Exiting Integral Calculation", vbOKOnly, "Integral Error"
        Exit Function
    End If

End Function
