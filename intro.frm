VERSION 5.00
Object = "{8BF1EB17-2379-11D5-80DE-C8AE202D4E0E}#1.0#0"; "ZProgBar.ocx"
Begin VB.Form frmIntro 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15390
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "intro.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   11520
   ScaleWidth      =   15390
   ShowInTaskbar   =   0   'False
   Begin ZealProgressBar.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   9120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   450
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mangal"
         Size            =   8.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   960
      Top             =   1560
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Loading....."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   7680
      Width           =   3615
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   11340
      Left            =   480
      Picture         =   "intro.frx":030A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15600
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Integer

Private Sub Timer1_Timer()
t = t + 1
If t = 1 Then
ProgressBar1.Value = 20
End If
If t = 2 Then
ProgressBar1.Value = 50
End If
If t = 3 Then
ProgressBar1.Value = 70
End If
If t = 4 Then
ProgressBar1.Value = 90
End If
If t = 5 Then
ProgressBar1.Value = 100
Unload frmIntro
Load frmname
frmname.Show
End If
End Sub
