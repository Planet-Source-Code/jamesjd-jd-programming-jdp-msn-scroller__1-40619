VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JD Programming - JDP Scroller"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3000
      Top             =   840
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Text            =   "http://www.jdprogramming.cjb.net"
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton ScrollStopButton 
      Caption         =   "Stop Scrolling"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton ScrollStartButton 
      Caption         =   "Start Scrolling"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox ScrollingText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "--> Your Scrolling Text Goes Here <--"
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = False
End Sub

Private Sub ScrollStartButton_Click()
Timer1.Enabled = True
End Sub

Private Sub ScrollStopButton_Click()
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
Call SendMessage(ScrollingText.Text)
End Sub
