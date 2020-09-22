VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmGraph 
   Caption         =   "Graph"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkGraphSeries 
      Alignment       =   1  'Right Justify
      Caption         =   "Force"
      Height          =   315
      Index           =   8
      Left            =   7440
      TabIndex        =   16
      Top             =   7350
      Width           =   885
   End
   Begin VB.CheckBox chkGraphSeries 
      Alignment       =   1  'Right Justify
      Caption         =   "Velocity"
      Height          =   315
      Index           =   7
      Left            =   5880
      TabIndex        =   14
      Top             =   7350
      Width           =   885
   End
   Begin VB.CheckBox chkGraphSeries 
      Alignment       =   1  'Right Justify
      Caption         =   "Integral Component"
      Height          =   435
      Index           =   6
      Left            =   4170
      TabIndex        =   12
      Top             =   7290
      Width           =   1125
   End
   Begin VB.CheckBox chkGraphSeries 
      Alignment       =   1  'Right Justify
      Caption         =   "Derivative Component"
      Height          =   435
      Index           =   5
      Left            =   2310
      TabIndex        =   10
      Top             =   7290
      Width           =   1125
   End
   Begin VB.CheckBox chkGraphSeries 
      Alignment       =   1  'Right Justify
      Caption         =   "Proportional Component"
      Height          =   465
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   7290
      Width           =   1335
   End
   Begin VB.CheckBox chkGraphSeries 
      Alignment       =   1  'Right Justify
      Caption         =   "Error"
      Height          =   315
      Index           =   3
      Left            =   4560
      TabIndex        =   7
      Top             =   6780
      Value           =   1  'Checked
      Width           =   705
   End
   Begin VB.CheckBox chkGraphSeries 
      Alignment       =   1  'Right Justify
      Caption         =   "Mass Position"
      Height          =   315
      Index           =   2
      Left            =   2100
      TabIndex        =   5
      Top             =   6810
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkGraphSeries 
      Alignment       =   1  'Right Justify
      Caption         =   "Set Position"
      Height          =   315
      Index           =   1
      Left            =   300
      TabIndex        =   3
      Top             =   6810
      Value           =   1  'Checked
      Width           =   1155
   End
   Begin VB.CommandButton cmdClearGraph 
      Caption         =   "Clear Graph"
      Height          =   345
      Left            =   9180
      TabIndex        =   1
      Top             =   7140
      Width           =   1185
   End
   Begin MSChart20Lib.MSChart graphPID 
      Height          =   6735
      Left            =   0
      OleObjectBlob   =   "frmGraph.frx":0000
      TabIndex        =   0
      Top             =   30
      Width           =   12225
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8370
      TabIndex        =   17
      Top             =   7380
      Width           =   285
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6840
      TabIndex        =   15
      Top             =   7380
      Width           =   285
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5340
      TabIndex        =   13
      Top             =   7380
      Width           =   285
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3480
      TabIndex        =   11
      Top             =   7380
      Width           =   285
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1500
      TabIndex        =   9
      Top             =   7380
      Width           =   285
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5310
      TabIndex        =   6
      Top             =   6810
      Width           =   285
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Top             =   6840
      Width           =   285
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1500
      TabIndex        =   2
      Top             =   6840
      Width           =   285
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkDer_Click()
    If chkDer.Value = 1 Then
        graphPID.Plot.SeriesCollection(5).ShowLine = True
    Else
        graphPID.Plot.SeriesCollection(5).ShowLine = False
    End If
    graphPID.Refresh
End Sub

Private Sub chkError_Click()
    If chkError.Value = 1 Then
        graphPID.Plot.SeriesCollection(3).ShowLine = True
    Else
        graphPID.Plot.SeriesCollection(3).ShowLine = False
    End If
End Sub

Private Sub chkInt_Click()
    If chkInt.Value = 1 Then
        graphPID.Plot.SeriesCollection(6).ShowLine = True
    Else
        graphPID.Plot.SeriesCollection(6).ShowLine = False
    End If
End Sub

Private Sub chkMassPosition_Click()
    If chkMassPosition.Value = 1 Then
        graphPID.Plot.SeriesCollection(2).ShowLine = True
    Else
        graphPID.Plot.SeriesCollection(2).ShowLine = False
    End If
End Sub

Private Sub chkProp_Click()
    If chkProp.Value = 1 Then
        graphPID.Plot.SeriesCollection(4).ShowLine = True
    Else
        graphPID.Plot.SeriesCollection(4).ShowLine = False
    End If
    graphPID.EditCopy
End Sub

Private Sub chkSetPosition_Click()
    If chkSetPosition.Value = 1 Then
        graphPID.Plot.SeriesCollection(1).ShowLine = True
    Else
        graphPID.Plot.SeriesCollection(1).ShowLine = False
    End If
End Sub

Private Sub chkVelocity_Click()
    If chkVelocity.Value = 1 Then
        graphPID.Plot.SeriesCollection(7).ShowLine = True
    Else
        graphPID.Plot.SeriesCollection(7).ShowLine = False
    End If
End Sub

Private Sub chkGraphSeries_Click(Index As Integer)
    If chkGraphSeries(Index).Value = 1 Then
        graphPID.Plot.SeriesCollection(Index).ShowLine = True
    Else
        graphPID.Plot.SeriesCollection(Index).ShowLine = False
    End If
    PlotActiveGraphData
    
End Sub

Private Sub cmdClearGraph_Click()
    ReDim GraphValues(6, 0)
    graphPID.ChartData = ModGraph.GraphValues
    
End Sub



Private Sub Form_Load()
    PlotActiveGraphData
    ShowActiveData
End Sub

Public Sub PlotActiveGraphData()
    Dim i As Integer
    Dim j As Integer
    ReDim ValuesToGraph(7, 0)
    For i = 0 To UBound(GraphValues, 2)
        For j = 1 To 8
            If Me.chkGraphSeries(j).Value = 1 Then
                ValuesToGraph(j - 1, i) = GraphValues(j - 1, i)
            Else
                ValuesToGraph(j - 1, i) = 0
            End If
        Next j
        ReDim Preserve ValuesToGraph(7, i + 1)
    Next i
    graphPID.ChartData = ModGraph.ValuesToGraph
    With graphPID.Plot.Axis(VtChAxisIdX).ValueScale
        .Auto = False
        .Minimum = 0
        .Maximum = UBound(GraphValues, 2)
        
    End With
End Sub

Private Sub ShowActiveData()
    Dim i As Integer
    For i = 1 To 8
        If chkGraphSeries(i).Value = 1 Then
            graphPID.Plot.SeriesCollection(i).ShowLine = True
        Else
            graphPID.Plot.SeriesCollection(i).ShowLine = False
        End If
    Next i
End Sub
