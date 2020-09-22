VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtAbout 
      Height          =   4935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmAbout.frx":0000
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub
