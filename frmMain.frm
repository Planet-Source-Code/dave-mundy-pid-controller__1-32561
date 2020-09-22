VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "PID (Proportional Integral Derivative) Demo"
   ClientHeight    =   9195
   ClientLeft      =   375
   ClientTop       =   975
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   14610
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   1215
      Left            =   13200
      TabIndex        =   66
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtMaxSineAmplitude 
      Height          =   285
      Left            =   3510
      TabIndex        =   62
      Text            =   "30000"
      Top             =   3600
      Width           =   825
   End
   Begin VB.CommandButton cmdSine 
      Caption         =   "Perform Sine Function"
      Height          =   345
      Left            =   450
      TabIndex        =   61
      Top             =   3600
      Width           =   1725
   End
   Begin VB.Timer tmrSetPointMover 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7260
      Top             =   2820
   End
   Begin VB.TextBox txtParabolaPeriod 
      Height          =   285
      Left            =   5640
      TabIndex        =   59
      Text            =   "3"
      Top             =   3270
      Width           =   825
   End
   Begin VB.TextBox txtRampPeriod 
      Height          =   285
      Left            =   5640
      TabIndex        =   57
      Text            =   "3"
      Top             =   2940
      Width           =   825
   End
   Begin VB.TextBox txtParabolaAmplitude 
      Height          =   285
      Left            =   3510
      TabIndex        =   55
      Text            =   "30000"
      Top             =   3270
      Width           =   825
   End
   Begin VB.CommandButton cmdParabola 
      Caption         =   "Perform Parabola"
      Height          =   345
      Left            =   450
      TabIndex        =   54
      Top             =   3270
      Width           =   1725
   End
   Begin VB.TextBox txtRampAmplitude 
      Height          =   285
      Left            =   3510
      TabIndex        =   52
      Text            =   "25000"
      Top             =   2940
      Width           =   825
   End
   Begin VB.CommandButton cmdRamp 
      Caption         =   "Perform Ramp"
      Height          =   345
      Left            =   450
      TabIndex        =   51
      Top             =   2940
      Width           =   1725
   End
   Begin VB.TextBox txtStepAmplitude 
      Height          =   285
      Left            =   3510
      TabIndex        =   49
      Text            =   "50000"
      Top             =   2610
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Graph"
      Height          =   345
      Left            =   7620
      TabIndex        =   48
      Top             =   3300
      Width           =   1725
   End
   Begin VB.TextBox txtFriction 
      Height          =   285
      Left            =   4380
      TabIndex        =   46
      Text            =   "100"
      Top             =   60
      Width           =   1695
   End
   Begin VB.TextBox txtGraphSampleTime 
      Height          =   285
      Left            =   11370
      TabIndex        =   43
      Text            =   "50"
      Top             =   3000
      Width           =   1305
   End
   Begin VB.Timer tmrGraph 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7260
      Top             =   2400
   End
   Begin VB.CommandButton cmdShowGraph 
      Caption         =   "Show Graph"
      Height          =   345
      Left            =   7620
      TabIndex        =   42
      Top             =   2970
      Width           =   1725
   End
   Begin VB.CommandButton cmdStepFcn 
      Caption         =   "Perform Step Function"
      Height          =   345
      Left            =   450
      TabIndex        =   32
      Top             =   2610
      Width           =   1725
   End
   Begin VB.CommandButton cmdResetController 
      Caption         =   "Velocity = 0"
      Height          =   345
      Left            =   450
      TabIndex        =   21
      Top             =   2220
      Width           =   1725
   End
   Begin VB.TextBox txtDerGain 
      Height          =   285
      Left            =   5550
      TabIndex        =   30
      Text            =   "200000"
      Top             =   1890
      Width           =   1695
   End
   Begin VB.TextBox txtIntGain 
      Height          =   285
      Left            =   5550
      TabIndex        =   28
      Text            =   "1000"
      Top             =   2250
      Width           =   1695
   End
   Begin VB.Timer tmrVelocity 
      Interval        =   1
      Left            =   7260
      Top             =   1560
   End
   Begin VB.Timer tmrPID 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7260
      Top             =   1980
   End
   Begin VB.CommandButton cmdStopPID 
      Caption         =   "Stop PID Control"
      Height          =   345
      Left            =   450
      TabIndex        =   26
      Top             =   1890
      Width           =   1725
   End
   Begin VB.CommandButton cmdStartPID 
      Caption         =   "Start PID Control"
      Height          =   345
      Left            =   450
      TabIndex        =   27
      Top             =   1560
      Width           =   1725
   End
   Begin VB.TextBox txtMaxV 
      Height          =   285
      Left            =   8730
      TabIndex        =   24
      Text            =   "100"
      Top             =   2250
      Width           =   825
   End
   Begin VB.TextBox txtDeadband 
      Height          =   285
      Left            =   8730
      TabIndex        =   22
      Text            =   "10"
      Top             =   2610
      Width           =   825
   End
   Begin VB.TextBox txtVel 
      Height          =   285
      Left            =   8040
      TabIndex        =   19
      Text            =   "0"
      Top             =   60
      Width           =   1695
   End
   Begin VB.CommandButton cmeResetMaxError 
      Caption         =   "Reset Max Error"
      Height          =   345
      Left            =   2280
      TabIndex        =   18
      Top             =   1560
      Width           =   1725
   End
   Begin VB.TextBox txtMaxError 
      Height          =   285
      Left            =   11280
      TabIndex        =   16
      Text            =   "0"
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtPropGain 
      Height          =   285
      Left            =   5550
      TabIndex        =   14
      Text            =   "4500000"
      Top             =   1530
      Width           =   1695
   End
   Begin VB.TextBox txtMinPower 
      Height          =   285
      Left            =   8730
      TabIndex        =   12
      Text            =   "5"
      Top             =   1890
      Width           =   825
   End
   Begin VB.TextBox txtMaxPower 
      Height          =   285
      Left            =   8730
      TabIndex        =   10
      Text            =   "1000000"
      Top             =   1530
      Width           =   825
   End
   Begin VB.TextBox txtError 
      Height          =   285
      Left            =   11280
      TabIndex        =   8
      Text            =   "0"
      Top             =   2190
      Width           =   1695
   End
   Begin VB.TextBox txtContPos 
      Height          =   285
      Left            =   11280
      TabIndex        =   6
      Text            =   "0"
      Top             =   1860
      Width           =   1695
   End
   Begin VB.TextBox txtMassPos 
      Height          =   285
      Left            =   11280
      TabIndex        =   4
      Text            =   "0"
      Top             =   1530
      Width           =   1695
   End
   Begin VB.TextBox txtMass 
      Height          =   285
      Left            =   1230
      TabIndex        =   2
      Text            =   "5000"
      Top             =   60
      Width           =   1695
   End
   Begin MSComctlLib.Slider sldMass 
      Height          =   345
      Left            =   930
      TabIndex        =   0
      Top             =   480
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   609
      _Version        =   393216
      Max             =   100000
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider sldCont 
      Height          =   630
      Left            =   930
      TabIndex        =   1
      Top             =   840
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   1111
      _Version        =   393216
      Max             =   100000
      TickFrequency   =   1000
      TextPosition    =   1
   End
   Begin VB.Label Label22 
      Caption         =   "Set Position"
      Height          =   375
      Left            =   0
      TabIndex        =   64
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label23 
      Caption         =   "Mass Position"
      Height          =   375
      Left            =   0
      TabIndex        =   65
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      Caption         =   "Max Amplitude:"
      Height          =   285
      Left            =   2250
      TabIndex        =   63
      Top             =   3630
      Width           =   1185
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      Caption         =   "Time Period (s):"
      Height          =   285
      Left            =   4380
      TabIndex        =   60
      Top             =   3300
      Width           =   1185
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "Time Period (s):"
      Height          =   285
      Left            =   4380
      TabIndex        =   58
      Top             =   2970
      Width           =   1185
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "Max Amplitude:"
      Height          =   285
      Left            =   2250
      TabIndex        =   56
      Top             =   3300
      Width           =   1185
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Caption         =   "Max Amplitude:"
      Height          =   285
      Left            =   2250
      TabIndex        =   53
      Top             =   2970
      Width           =   1185
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Max Amplitude:"
      Height          =   285
      Left            =   2250
      TabIndex        =   50
      Top             =   2640
      Width           =   1185
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Frictional Force:"
      Height          =   285
      Left            =   3000
      TabIndex        =   47
      Top             =   90
      Width           =   1305
   End
   Begin VB.Label lblMaxSampleTime 
      Caption         =   "Max Sample Time = 50 Seconds"
      Height          =   285
      Left            =   9420
      TabIndex        =   45
      Top             =   3330
      Width           =   3285
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Graph Sample Time (ms):"
      Height          =   285
      Left            =   9420
      TabIndex        =   44
      Top             =   3000
      Width           =   1875
   End
   Begin VB.Label lblSetPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   90
      TabIndex        =   41
      ToolTipText     =   "Derivative Result"
      Top             =   5910
      Width           =   1215
   End
   Begin VB.Label lblError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   6480
      TabIndex        =   40
      ToolTipText     =   "Derivative Result"
      Top             =   8820
      Width           =   3105
   End
   Begin VB.Label txtForce 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   10530
      TabIndex        =   39
      ToolTipText     =   "Derivative Result"
      Top             =   5910
      Width           =   1605
   End
   Begin VB.Label lblPropandGain 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   2670
      TabIndex        =   38
      ToolTipText     =   "Derivative Result"
      Top             =   6600
      Width           =   2745
   End
   Begin VB.Label txtProp 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   3510
      TabIndex        =   37
      ToolTipText     =   "Derivative Result"
      Top             =   5700
      Width           =   1995
   End
   Begin VB.Label lblIntandGain 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   6450
      TabIndex        =   36
      ToolTipText     =   "Derivative Result"
      Top             =   8040
      Width           =   3105
   End
   Begin VB.Label lblDerandGain 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   6450
      TabIndex        =   35
      ToolTipText     =   "Derivative Result"
      Top             =   5160
      Width           =   3075
   End
   Begin VB.Label txtDer 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   7530
      TabIndex        =   33
      ToolTipText     =   "Derivative Result"
      Top             =   4440
      Width           =   2025
   End
   Begin VB.Label txtInt 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   7560
      TabIndex        =   34
      ToolTipText     =   "Derivative Result"
      Top             =   7410
      Width           =   1965
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Deravative Gain:"
      Height          =   285
      Left            =   4080
      TabIndex        =   31
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Integral Gain:"
      Height          =   285
      Left            =   4080
      TabIndex        =   29
      Top             =   2280
      Width           =   1395
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Max Velocity:"
      Height          =   285
      Left            =   7650
      TabIndex        =   25
      Top             =   2280
      Width           =   1005
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Deadband:"
      Height          =   285
      Left            =   7650
      TabIndex        =   23
      Top             =   2640
      Width           =   1005
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Velocity (pulses/ms) :"
      Height          =   285
      Left            =   6300
      TabIndex        =   20
      Top             =   90
      Width           =   1665
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Max Error:"
      Height          =   285
      Left            =   9960
      TabIndex        =   17
      Top             =   2550
      Width           =   1245
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Proportional Gain:"
      Height          =   285
      Left            =   4080
      TabIndex        =   15
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Min Power:"
      Height          =   285
      Left            =   7650
      TabIndex        =   13
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Max Power:"
      Height          =   285
      Left            =   7650
      TabIndex        =   11
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Position Error:"
      Height          =   285
      Left            =   9960
      TabIndex        =   9
      Top             =   2220
      Width           =   1245
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Control Position:"
      Height          =   285
      Left            =   9960
      TabIndex        =   7
      Top             =   1890
      Width           =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Mass Position:"
      Height          =   285
      Left            =   9990
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Mass (g):"
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   90
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   5250
      Left            =   900
      Picture         =   "frmMain.frx":0000
      Top             =   3750
      Width           =   12885
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************* PID Demo ***************
'Proportional integral derivative controller Demo
'Author:  Dave Mundy
'Date: March 2002
'**************************************
Dim CurV As Double
Dim SetPointCurV As Double
Dim IsActive As Boolean
Dim Iteration As Double
Dim PerformingRamp As Boolean
Dim PerformingParabola As Boolean
Dim ParabolaCount As Double
Dim ParabolaDirection As Boolean    'true = positive
                                  'false = negative
Dim PerformingSin As Boolean
Dim ResultOfProp As Double
Dim ResultOfDer As Double
Dim ResultOfInt As Double
Dim ResultOfForcePercent As Double
Dim WithEvents PIDController As clsPID
Attribute PIDController.VB_VarHelpID = -1

Private Sub chkToFile_Click()

End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show
End Sub

Private Sub cmdParabola_Click()
    CurV = 0
    ParabolaCount = 1
    tmrGraph.Enabled = True
    Me.cmdStopPID.Enabled = False
    ParabolaDirection = True
    Me.tmrPID.Enabled = False
    Me.sldMass.SelStart = 0
    Me.sldCont.SelStart = 0
    Me.sldCont.SelStart = (Val(Me.sldMass.Max) / 2) - (Val(Me.txtParabolaAmplitude))
    Me.sldMass.SelStart = (Val(Me.sldMass.Max) / 2) - (Val(Me.txtParabolaAmplitude))
    InitializePID
    Me.tmrPID.Enabled = True
    PerformingParabola = True
    SetPointCurV = Val(Me.txtParabolaAmplitude) / ((Val(Me.txtParabolaPeriod) / 2) * 100)
'    MsgBox (SetPointCurV)
    Me.tmrSetPointMover.Enabled = True
    
    Me.txtContPos = sldCont.SelStart
    Me.lblSetPos = sldCont.SelStart
End Sub

Private Sub cmdRamp_Click()
    CurV = 0
    tmrGraph.Enabled = True
    Me.cmdStopPID.Enabled = False
    Me.tmrPID.Enabled = False
    Me.sldMass.SelStart = 0
    Me.sldCont.SelStart = 0
    Me.sldCont.SelStart = (Val(Me.sldMass.Max) / 2) - (Val(Me.txtRampAmplitude))
    Me.sldMass.SelStart = (Val(Me.sldMass.Max) / 2) - (Val(Me.txtRampAmplitude))
    InitializePID
    Me.tmrPID.Enabled = True
    PerformingRamp = True
    SetPointCurV = Val(Me.txtRampAmplitude) / ((Val(Me.txtRampPeriod) / 2) * 100)
'    MsgBox (SetPointCurV)
    Me.tmrSetPointMover.Enabled = True
    
    Me.txtContPos = sldCont.SelStart
    Me.lblSetPos = sldCont.SelStart
End Sub

Private Sub cmdSine_Click()
    If PerformingSin = False Then
        tmrGraph.Enabled = True
        PerformingSin = True
        Me.sldCont.SelStart = (Val(Me.sldMass.Max) / 2)
        Me.sldMass.SelStart = (Val(Me.sldMass.Max) / 2)
        ParabolaCount = 0
        Me.tmrPID.Enabled = True
        Me.tmrSetPointMover.Enabled = True
        cmdSine.Caption = "Stop Sine Function"
    Else
        tmrGraph.Enabled = False
        PerformingSin = False
        'Me.sldCont.SelStart = (Val(Me.sldMass.Max) / 2)
        'Me.sldMass.SelStart = (Val(Me.sldMass.Max) / 2)
        'ParabolaCount = 0
        Me.tmrPID.Enabled = False
        Me.tmrSetPointMover.Enabled = False
        cmdSine.Caption = "Perform Sine Function"
    End If
End Sub

Private Sub Command1_Click()
    ReDim GraphValues(7, 0)
End Sub



Private Sub tmrGraph_Timer()
    ModGraph.AddtoGraphArray Me.sldCont.SelStart, Me.sldMass.SelStart, (Me.sldCont - Me.sldMass), ResultOfProp, ResultOfDer, ResultOfInt, CurV, ResultOfForcePercent * Val(Me.txtMaxPower)
End Sub

Private Sub tmrSetPointMover_Timer()
    If PerformingRamp = True Then
        If Me.sldCont.SelStart >= Me.sldCont.Max / 2 Then
            SetPointCurV = -1 * Val(Me.txtRampAmplitude) / ((Val(Me.txtRampPeriod) / 2) * 100)
            Me.sldCont.SelStart = Val(Me.sldCont.SelStart) + SetPointCurV
        Else
            Me.sldCont.SelStart = Val(Me.sldCont.SelStart) + SetPointCurV
        End If
        If Me.sldCont.SelStart <= (Val(Me.sldMass.Max) / 2) - (Val(Me.txtRampAmplitude)) Then
            Me.tmrSetPointMover.Enabled = False
            PerformingRamp = False
            Me.cmdStopPID.Enabled = True
        End If
    End If
    
    If PerformingParabola = True Then
        If ParabolaDirection = True Then 'positive
            ParabolaCount = ParabolaCount + 0.1
            If ((Val(Me.sldMass.Max) / 2) - Me.sldCont.SelStart) < 5000 Then
                SetPointCurV = (ParabolaCount * ParabolaCount) - (1000 / ((Val(Me.sldCont.Max) / 2) - Me.sldCont.SelStart)) ^ 2
            Else
                SetPointCurV = (ParabolaCount * ParabolaCount)
            End If
            Me.sldCont.SelStart = Val(Me.sldCont.SelStart) + SetPointCurV
            If Me.sldCont.SelStart >= Me.sldCont.Max / 2 Then
                ParabolaDirection = False
            End If
            'debug.print SetPointCurV
        Else 'negative
            ParabolaCount = ParabolaCount - 0.1
            If ((Val(Me.sldMass.Max) / 2) - Me.sldCont.SelStart) < 5000 Then
                SetPointCurV = (ParabolaCount * ParabolaCount) - (1000 / ((Val(Me.sldCont.Max) / 2) - Me.sldCont.SelStart)) ^ 2
            Else
                SetPointCurV = (ParabolaCount * ParabolaCount)
            End If
            Me.sldCont.SelStart = Val(Me.sldCont.SelStart) - SetPointCurV
            'debug.print SetPointCurV
        End If

        If Me.sldCont.SelStart <= (Val(Me.sldMass.Max) / 2) - (Val(Me.txtParabolaAmplitude)) Then
            Me.tmrSetPointMover.Enabled = False
            PerformingParabola = False
            Me.cmdStopPID.Enabled = True
        End If
    End If
    
    If PerformingSin = True Then
        ParabolaCount = ParabolaCount + 0.05
        Me.sldCont.SelStart = ((Val(Me.sldMass.Max) / 2) + 10000 * Sin(ParabolaCount))
        
    End If
    
End Sub

Private Sub cmdResetController_Click()
    CurV = 0
    'PIDControllerIntegral 0, , , True
    
End Sub

Private Sub cmdShowGraph_Click()
    'MsgBox (UBound(GraphValues, 2))
    tmrGraph.Enabled = False
    Me.tmrPID.Enabled = False
    frmGraph.Show vbModal
End Sub

Private Sub cmdStartPID_Click()
    GraphNumber = 0
    tmrGraph.Enabled = True
    tmrPID.Enabled = True
    tmrVelocity.Enabled = True
End Sub

Private Sub cmdStepFcn_Click()
    CurV = 0
    'Integral 0, , , True
    Me.tmrPID.Enabled = False
    Me.sldMass.SelStart = 0
    Me.sldCont.SelStart = 0
    Me.sldCont.SelStart = Val(Me.txtStepAmplitude)
    InitializePID
    tmrGraph.Enabled = True
    Me.tmrPID.Enabled = True
    Me.txtContPos = sldCont.SelStart
    Me.lblSetPos = sldCont.SelStart
End Sub

Private Sub InitializePID()
    'Set PIDController = New clsPID
    Dim CurError As Double
    Dim RetForce As Double
    
    PIDController.PropGain = Val(Me.txtPropGain)
    PIDController.DerGain = Val(Me.txtDerGain)
    PIDController.IntGain = Val(Me.txtIntGain)
    PIDController.DeadBand = Val(Me.txtDeadband)
    CurError = (Me.sldCont - Me.sldMass) / Me.sldMass.Max
    
    RetForce = PIDController.CalcPID(CurError)
End Sub

Private Sub cmdStopPID_Click()
    tmrGraph.Enabled = False
    tmrPID.Enabled = False
    'tmrVelocity.Enabled = False
End Sub

Private Sub cmeResetMaxError_Click()
    Me.txtMaxError.Text = 0
    txtMaxError.Text = 0
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Set PIDController = New clsPID
    ReDim GraphValues(7, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub PIDController_CalcUpdates(PropResult As Double, IntResult As Double, DerResult As Double, ForcePercent As Double)
    'Dim RedimDim As Long
    
    Me.txtContPos = Me.sldCont.SelStart
    Me.txtMassPos = Me.sldMass.SelStart
    Me.txtDer = DerResult * Me.txtDerGain
    Me.txtInt = IntResult * Me.txtIntGain
    Me.txtProp = PropResult * Me.txtPropGain
    Me.lblError = Me.sldMass.SelStart
    
    ResultOfProp = PropResult
    ResultOfDer = DerResult
    ResultOfInt = IntResult
    ResultOfForcePercent = ForcePercent
    
    Me.lblDerandGain = Me.txtDerGain & " * " & Str(DerResult)
    Me.lblIntandGain = Me.txtIntGain & " * " & Str(IntResult)
    Me.lblPropandGain = Me.txtProp & " * " & Str((Me.sldCont - Me.sldMass) / Me.sldMass.Max)
    Me.txtForce = Format((ForcePercent * Val(Me.txtMaxPower)), "########.###")
    Me.txtError = Str(Me.sldCont - Me.sldMass)
    Me.lblSetPos = sldCont.SelStart
End Sub

Private Sub sldCont_Change()
    Me.txtContPos = sldCont.SelStart
    Me.lblSetPos = sldCont.SelStart
    DoEvents
End Sub

Private Sub tmrPID_Timer()

    Dim CurError As Double
    Dim RetForce As Double
    
    PIDController.PropGain = Val(Me.txtPropGain)
    PIDController.DerGain = Val(Me.txtDerGain)
    PIDController.IntGain = Val(Me.txtIntGain)
    PIDController.DeadBand = Val(Me.txtDeadband)
    CurError = (Me.sldCont - Me.sldMass) / Me.sldMass.Max
    
    RetForce = PIDController.CalcPID(CurError)
    
    RetForce = RetForce * Me.txtMaxPower
    
    Me.ChangeVelocity (RetForce / Me.txtMass)
    
    DoEvents
    
End Sub

Private Sub tmrVelocity_Timer()
    Dim FinalPos As Double
    If CurV <> 0 Then
        CurV = CurV - ((Val(Me.txtFriction) * Abs(CurV) / CurV) / Val(Me.txtMass))
        FinalPos = Me.sldMass.SelStart + CurV
    
    
        If FinalPos > Me.sldMass.Min And FinalPos < Me.sldMass.Max Then
            Me.sldMass.SelStart = FinalPos
        ElseIf FinalPos < Me.sldMass.Min Then
            Me.sldMass.SelStart = Me.sldMass.Min
        ElseIf FinalPos < Me.sldMass.Min Then
            Me.sldMass.SelStart = Me.sldMass.Max
        End If
    End If
    Me.txtVel.Text = CurV
End Sub

Public Sub ChangeVelocity(Acceleration As Double)
    CurV = CurV + Acceleration
End Sub

Private Sub txtGraphSampleTime_Change()
    
    If Val(Me.txtGraphSampleTime.Text) <> 0 Then
        lblMaxSampleTime.Caption = "Max Sample Time = " & Str(5000 / (1000 / Val(Me.txtGraphSampleTime.Text))) & " Seconds"
        Me.tmrGraph.Interval = Val(Me.txtGraphSampleTime)
    Else
        MsgBox ("Invalid Value")
    End If
End Sub
