VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmname 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Play"
   ClientHeight    =   7920
   ClientLeft      =   3075
   ClientTop       =   1515
   ClientWidth     =   9240
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   9240
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2640
      TabIndex        =   39
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   495
      Left            =   360
      TabIndex        =   38
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Frame frame1 
      Height          =   615
      Left            =   360
      TabIndex        =   31
      Top             =   4440
      Width           =   8655
      Begin VB.OptionButton optr1 
         Caption         =   "Red"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optg1 
         Caption         =   "Green"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   2640
         TabIndex        =   34
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton opty1 
         Caption         =   "Yellow"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optb1 
         Caption         =   "Blue"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5160
         TabIndex        =   32
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "1st Player"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame farme2 
      Height          =   615
      Left            =   360
      TabIndex        =   25
      Top             =   4920
      Width           =   8655
      Begin VB.OptionButton optr2 
         Caption         =   "Red"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optg2 
         Caption         =   "Green"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   2640
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton opty2 
         Caption         =   "Yellow"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optb2 
         Caption         =   "Blue"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5160
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "2nd Player"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   360
      TabIndex        =   19
      Top             =   5400
      Width           =   8655
      Begin VB.OptionButton optr3 
         Caption         =   "Red"
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optg3 
         Caption         =   "Green"
         Enabled         =   0   'False
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton opty3 
         Caption         =   "Yellow"
         Enabled         =   0   'False
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optb3 
         Caption         =   "Blue"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5160
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "3rd Player"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   360
      TabIndex        =   13
      Top             =   5880
      Width           =   8655
      Begin VB.OptionButton optr4 
         Caption         =   "Red"
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optg4 
         Caption         =   "Green"
         Enabled         =   0   'False
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton opty4 
         Caption         =   "Yellow"
         Enabled         =   0   'False
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optb4 
         Caption         =   "Blue"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5160
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "4th Player"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox txtp4name 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   2400
      Width           =   3495
   End
   Begin VB.TextBox txtp3name 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox txtp2name 
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   1440
      Width           =   3495
   End
   Begin VB.TextBox txtp1name 
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
   Begin VB.Frame framenp 
      Caption         =   "Name of the players"
      Height          =   2655
      Left            =   3000
      TabIndex        =   8
      Top             =   600
      Width           =   6015
      Begin VB.Label Label4 
         Caption         =   "4th Player"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "3rd player"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "2nd Player"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "1st Player"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.OptionButton opt4 
      Caption         =   "Four Players"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.OptionButton opt3 
      Caption         =   "Three Players"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin VB.OptionButton opt2 
      Caption         =   "Double players"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.Frame framepm 
      Caption         =   "Select the play mode"
      Height          =   2655
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   240
      Left            =   2640
      TabIndex        =   40
      Top             =   120
      Width           =   240
      URL             =   "C:\WINDOWS\Media\onestop.mid"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   999
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "invisible"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   423
      _cy             =   423
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      Height          =   255
      Left            =   5520
      Shape           =   3  'Circle
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      Height          =   255
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      Height          =   255
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "Select colour"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   37
      Top             =   3720
      Width           =   2895
   End
End
Attribute VB_Name = "frmname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmmain.Show
frmname.Hide

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
frmname.Width = 9360
frmname.Height = 8490
opt2.Value = True
End Sub

Private Sub opt2_Click()
txtp3name.Text = ""
txtp4name.Text = ""
Label4.Enabled = False
Label3.Enabled = False
txtp3name.Enabled = False
txtp4name.Enabled = False
Label9.Enabled = True
opty3.Enabled = False
optb3.Enabled = False
optg3.Enabled = False
optr3.Enabled = False
opty4.Enabled = False
optb4.Enabled = False
optg4.Enabled = False
optr4.Enabled = False
If optr1.Value = True Then
 optr2.Value = False
 optr3.Value = False
 optr4.Value = False
End If
End Sub

Private Sub opt3_Click()
txtp4name.Text = ""
Label4.Enabled = False
Label3.Enabled = True
txtp3name.Enabled = True
txtp4name.Enabled = False
optr2.Enabled = True
optr3.Enabled = True
opty3.Enabled = True
optb3.Enabled = True
optg3.Enabled = True
opty4.Enabled = False
optb4.Enabled = False
optg4.Enabled = False
optr4.Enabled = False
If optr1.Value = True Then
optr2.Value = False
optr3.Value = False
optr4.Value = False
optr2.Enabled = False
optr3.Enabled = False
optr4.Enabled = False
End If
End Sub

Private Sub opt4_Click()
Label3.Enabled = True
opty3.Enabled = True
optb3.Enabled = True
optg3.Enabled = True
optr3.Enabled = True
optr2.Enabled = True
Label4.Enabled = True
opty4.Enabled = True
optb4.Enabled = True
optg4.Enabled = True
optr4.Enabled = True
txtp3name.Enabled = True
txtp4name.Enabled = True
If optr1.Value = True Then
optr2.Value = False
optr3.Value = False
optr4.Value = False
optr2.Enabled = False
optr3.Enabled = False
optr4.Enabled = False
End If

End Sub

Private Sub optb1_Click()
If opt4.Value = True Then
    optb2.Value = False
    optb3.Value = False
    optb4.Value = False
    optb2.Enabled = False
    optb3.Enabled = False
    optb4.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = True
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = True
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = True
ElseIf opt3.Value = True Then
    optb2.Value = False
    optb3.Value = False
    optb4.Value = False
    optb2.Enabled = False
    optb3.Enabled = False
    optb4.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = False
    optr4.Value = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = False
    optg4.Value = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = False
    opty4.Value = False
ElseIf opt2.Value = True Then
    optb2.Value = False
    optb3.Value = False
    optb4.Value = False
    optb2.Enabled = False
    optb3.Enabled = False
    optb4.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = False
    optr3.Value = False
    optr4.Value = False
    optr4.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = False
    optg3.Value = False
    optg4.Value = False
    optg4.Enabled = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = False
    opty3.Value = False
    opty4.Value = False
    opty4.Enabled = False
End If
End Sub

Private Sub optb2_Click()
If opt4.Value = True Then
    optb1.Value = False
    optb3.Value = False
    optb4.Value = False
    optb1.Enabled = False
    optb3.Enabled = False
    optb4.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = True
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = True
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = True
ElseIf opt3.Value = True Then
    optb1.Value = False
    optb3.Value = False
    optb4.Value = False
    optb1.Enabled = False
    optb3.Enabled = False
    optb4.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = False
    optr4.Value = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = False
    optg4.Value = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = False
    opty4.Value = False
ElseIf opt2.Value = True Then
    optb1.Value = False
    optb3.Value = False
    optb4.Value = False
    optb1.Enabled = False
    optb3.Enabled = False
    optb4.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = False
    optr3.Value = False
    optr4.Value = False
    optr4.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = False
    optg3.Value = False
    optg4.Value = False
    optg4.Enabled = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = False
    opty3.Value = False
    opty4.Value = False
    opty4.Enabled = False
End If

End Sub

Private Sub optb3_Click()
If opt4.Value = True Then
    optb2.Value = False
    optb1.Value = False
    optb4.Value = False
    optb2.Enabled = False
    optb1.Enabled = False
    optb4.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = True
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = True
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = True
ElseIf opt3.Value = True Then
    optb2.Value = False
    optb1.Value = False
    optb4.Value = False
    optb2.Enabled = False
    optb1.Enabled = False
    optb4.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = False
    optr4.Value = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = False
    optg4.Value = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = False
    opty4.Value = False
ElseIf opt2.Value = True Then
    optb2.Value = False
    optb1.Value = False
    optb4.Value = False
    optb2.Enabled = False
    optb1.Enabled = False
    optb4.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = False
    optr3.Value = False
    optr4.Value = False
    optr4.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = False
    optg3.Value = False
    optg4.Value = False
    optg4.Enabled = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = False
    opty3.Value = False
    opty4.Value = False
    opty4.Enabled = False
End If

End Sub

Private Sub optb4_Click()
If opt4.Value = True Then
    optb2.Value = False
    optb3.Value = False
    optb1.Value = False
    optb2.Enabled = False
    optb3.Enabled = False
    optb1.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = True
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = True
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = True
ElseIf opt3.Value = True Then
    optb2.Value = False
    optb3.Value = False
    optb1.Value = False
    optb2.Enabled = False
    optb3.Enabled = False
    optb1.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = False
    optr4.Value = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = False
    optg4.Value = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = False
    opty4.Value = False
ElseIf opt2.Value = True Then
    optb2.Value = False
    optb3.Value = False
    optb1.Value = False
    optb2.Enabled = False
    optb3.Enabled = False
    optb1.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = False
    optr3.Value = False
    optr4.Value = False
    optr4.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = False
    optg3.Value = False
    optg4.Value = False
    optg4.Enabled = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = False
    opty3.Value = False
    opty4.Value = False
    opty4.Enabled = False
End If

End Sub

Private Sub optg1_Click()
If opt4.Value = True Then
    optg2.Value = False
    optg3.Value = False
    optg4.Value = False
    optg3.Enabled = False
    optg4.Enabled = False
    optg2.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = True
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = True
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = True
ElseIf opt3.Value = True Then
    optg2.Value = False
    optg3.Value = False
    optg4.Value = False
    optg3.Enabled = False
    optg4.Enabled = False
    optg2.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = False
    optr4.Value = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = False
    opty4.Value = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = False
    optb4.Value = False
ElseIf opt2.Value = True Then
    optg2.Value = False
    optg3.Value = False
    optg4.Value = False
    optg3.Enabled = False
    optg4.Enabled = False
    optg2.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = False
    optr3.Value = False
    optr4.Enabled = False
    optr4.Value = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = False
    opty3.Value = False
    opty4.Value = False
    opty4.Enabled = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = False
    optb3.Value = False
    optb4.Value = False
    optb4.Enabled = False
End If
End Sub

Private Sub optg2_Click()
If opt4.Value = True Then
    optg1.Value = False
    optg3.Value = False
    optg4.Value = False
    optg3.Enabled = False
    optg4.Enabled = False
    optg1.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = True
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = True
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = True
ElseIf opt3.Value = True Then
    optg1.Value = False
    optg3.Value = False
    optg4.Value = False
    optg3.Enabled = False
    optg4.Enabled = False
    optg1.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = False
    optr4.Value = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = False
    opty4.Value = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = False
    optb4.Value = False
ElseIf opt2.Value = True Then
    optg1.Value = False
    optg3.Value = False
    optg4.Value = False
    optg3.Enabled = False
    optg4.Enabled = False
    optg1.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = False
    optr3.Value = False
    optr4.Enabled = False
    optr4.Value = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = False
    opty3.Value = False
    opty4.Value = False
    opty4.Enabled = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = False
    optb3.Value = False
    optb4.Value = False
    optb4.Enabled = False
End If
End Sub

Private Sub optg3_Click()
If opt4.Value = True Then
    optg2.Value = False
    optg1.Value = False
    optg4.Value = False
    optg1.Enabled = False
    optg4.Enabled = False
    optg2.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = True
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = True
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = True
ElseIf opt3.Value = True Then
    optg2.Value = False
    optg1.Value = False
    optg4.Value = False
    optg1.Enabled = False
    optg4.Enabled = False
    optg2.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = False
    optr4.Value = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = False
    opty4.Value = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = False
    optb4.Value = False
ElseIf opt2.Value = True Then
    optg2.Value = False
    optg1.Value = False
    optg4.Value = False
    optg1.Enabled = False
    optg4.Enabled = False
    optg2.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = False
    optr3.Value = False
    optr4.Enabled = False
    optr4.Value = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = False
    opty3.Value = False
    opty4.Value = False
    opty4.Enabled = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = False
    optb3.Value = False
    optb4.Value = False
    optb4.Enabled = False
End If
End Sub

Private Sub optg4_Click()
If opt4.Value = True Then
    optg2.Value = False
    optg3.Value = False
    optg1.Value = False
    optg3.Enabled = False
    optg1.Enabled = False
    optg2.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = True
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = True
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = True
ElseIf opt3.Value = True Then
    optg2.Value = False
    optg3.Value = False
    optg1.Value = False
    optg3.Enabled = False
    optg1.Enabled = False
    optg2.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = False
    optr4.Value = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = False
    opty4.Value = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = False
    optb4.Value = False
ElseIf opt2.Value = True Then
    optg2.Value = False
    optg3.Value = False
    optg1.Value = False
    optg3.Enabled = False
    optg1.Enabled = False
    optg2.Enabled = False
    optr1.Enabled = True
    optr2.Enabled = True
    optr3.Enabled = False
    optr3.Value = False
    optr4.Enabled = False
    optr4.Value = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = False
    opty3.Value = False
    opty4.Value = False
    opty4.Enabled = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = False
    optb3.Value = False
    optb4.Value = False
    optb4.Enabled = False
End If

End Sub

Private Sub optr1_Click()
If opt4.Value = True Then
    optr2.Value = False
    optr3.Value = False
    optr4.Value = False
    optr2.Enabled = False
    optr3.Enabled = False
    optr4.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = True
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = True
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = True
ElseIf opt3.Value = True Then
    optr2.Value = False
    optr3.Value = False
    optr4.Value = False
    optr2.Enabled = False
    optr3.Enabled = False
    optr4.Enabled = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Value = False
    opty4.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = False
    optg4.Value = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Value = False
    optb4.Enabled = False
ElseIf opt2.Value = True Then
    optr2.Value = False
    optr3.Value = False
    optr4.Value = False
    optr2.Enabled = False
    optr3.Enabled = False
    optr4.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Value = False
    optg3.Enabled = False
    optg4.Value = False
    optg4.Enabled = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Value = False
    opty4.Value = False
    opty3.Enabled = False
    opty4.Enabled = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Value = False
    optb4.Value = False
    optb3.Enabled = False
    optb4.Enabled = False
End If
End Sub

Private Sub optr2_Click()
If opt4.Value = True Then
    optr1.Value = False
    optr3.Value = False
    optr4.Value = False
    optr1.Enabled = False
    optr3.Enabled = False
    optr4.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = True
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = True
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = True
ElseIf opt3.Value = True Then
    optr1.Value = False
    optr3.Value = False
    optr4.Value = False
    optr1.Enabled = False
    optr3.Enabled = False
    optr4.Enabled = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Value = False
    opty4.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = False
    optg4.Value = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Value = False
    optb4.Enabled = False
ElseIf opt2.Value = True Then
    optr1.Value = False
    optr3.Value = False
    optr4.Value = False
    optr1.Enabled = False
    optr3.Enabled = False
    optr4.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Value = False
    optg3.Enabled = False
    optg4.Value = False
    optg4.Enabled = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Value = False
    opty4.Value = False
    opty3.Enabled = False
    opty4.Enabled = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Value = False
    optb4.Value = False
    optb3.Enabled = False
    optb4.Enabled = False
End If
End Sub

Private Sub optr3_Click()
If opt4.Value = True Then
    optr2.Value = False
    optr1.Value = False
    optr4.Value = False
    optr2.Enabled = False
    optr1.Enabled = False
    optr4.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = True
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = True
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = True
ElseIf opt3.Value = True Then
    optr2.Value = False
    optr1.Value = False
    optr4.Value = False
    optr2.Enabled = False
    optr1.Enabled = False
    optr4.Enabled = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Value = False
    opty4.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = False
    optg4.Value = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Value = False
    optb4.Enabled = False
ElseIf opt2.Value = True Then
    optr2.Value = False
    optr1.Value = False
    optr4.Value = False
    optr2.Enabled = False
    optr1.Enabled = False
    optr4.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Value = False
    optg3.Enabled = False
    optg4.Value = False
    optg4.Enabled = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Value = False
    opty4.Value = False
    opty3.Enabled = False
    opty4.Enabled = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Value = False
    optb4.Value = False
    optb3.Enabled = False
    optb4.Enabled = False
End If
End Sub

Private Sub optr4_Click()
If opt4.Value = True Then
    optr2.Value = False
    optr3.Value = False
    optr1.Value = False
    optr2.Enabled = False
    optr3.Enabled = False
    optr1.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = True
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Enabled = True
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = True
ElseIf opt3.Value = True Then
    optr2.Value = False
    optr3.Value = False
    optr1.Value = False
    optr2.Enabled = False
    optr3.Enabled = False
    optr1.Enabled = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Enabled = True
    opty4.Value = False
    opty4.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = False
    optg4.Value = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Value = False
    optb4.Enabled = False
ElseIf opt2.Value = True Then
    optr2.Value = False
    optr3.Value = False
    optr1.Value = False
    optr2.Enabled = False
    optr3.Enabled = False
    optr1.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Value = False
    optg3.Enabled = False
    optg4.Value = False
    optg4.Enabled = False
    opty1.Enabled = True
    opty2.Enabled = True
    opty3.Value = False
    opty4.Value = False
    opty3.Enabled = False
    opty4.Enabled = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Value = False
    optb4.Value = False
    optb3.Enabled = False
    optb4.Enabled = False
End If
End Sub

Private Sub opty1_Click()
If opt4.Value = True Then
    opty2.Value = False
    opty2.Enabled = False
    opty3.Value = False
    opty3.Enabled = False
    opty4.Value = False
    opty4.Enabled = False
    optr2.Enabled = True
    optr1.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = True
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = True
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = True
ElseIf opt3.Value = True Then
    opty2.Value = False
    opty2.Enabled = False
    opty3.Value = False
    opty3.Enabled = False
    opty4.Value = False
    opty4.Enabled = False
    optr2.Enabled = True
    optr1.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = False
    optr4.Value = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = False
    optg4.Value = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = False
    optb4.Value = False
ElseIf opt2.Value = True Then
    opty2.Value = False
    opty2.Enabled = False
    opty3.Value = False
    opty3.Enabled = False
    opty4.Value = False
    opty4.Enabled = False
    optr2.Enabled = True
    optr1.Enabled = True
    optr3.Enabled = False
    optr3.Value = False
    optr4.Value = False
    optr4.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = False
    optg3.Value = False
    optg4.Value = False
    optg4.Enabled = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = False
    optb3.Value = False
    optb4.Value = False
    optb4.Enabled = False
End If
End Sub

Private Sub opty2_Click()
If opt4.Value = True Then
    opty1.Value = False
    opty1.Enabled = False
    opty3.Value = False
    opty3.Enabled = False
    opty4.Value = False
    opty4.Enabled = False
    optr2.Enabled = True
    optr1.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = True
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = True
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = True
ElseIf opt3.Value = True Then
    opty1.Value = False
    opty1.Enabled = False
    opty3.Value = False
    opty3.Enabled = False
    opty4.Value = False
    opty4.Enabled = False
    optr2.Enabled = True
    optr1.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = False
    optr4.Value = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = False
    optg4.Value = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = False
    optb4.Value = False
ElseIf opt2.Value = True Then
    opty1.Value = False
    opty1.Enabled = False
    opty3.Value = False
    opty3.Enabled = False
    opty4.Value = False
    opty4.Enabled = False
    optr2.Enabled = True
    optr1.Enabled = True
    optr3.Enabled = False
    optr3.Value = False
    optr4.Value = False
    optr4.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = False
    optg3.Value = False
    optg4.Value = False
    optg4.Enabled = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = False
    optb3.Value = False
    optb4.Value = False
    optb4.Enabled = False
End If
End Sub

Private Sub opty3_Click()
If opt4.Value = True Then
    opty2.Value = False
    opty2.Enabled = False
    opty1.Value = False
    opty1.Enabled = False
    opty4.Value = False
    opty4.Enabled = False
    optr2.Enabled = True
    optr1.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = True
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = True
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = True
ElseIf opt3.Value = True Then
    opty2.Value = False
    opty2.Enabled = False
    opty1.Value = False
    opty1.Enabled = False
    opty4.Value = False
    opty4.Enabled = False
    optr2.Enabled = True
    optr1.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = False
    optr4.Value = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = False
    optg4.Value = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = False
    optb4.Value = False
ElseIf opt2.Value = True Then
    opty2.Value = False
    opty2.Enabled = False
    opty1.Value = False
    opty1.Enabled = False
    opty4.Value = False
    opty4.Enabled = False
    optr2.Enabled = True
    optr1.Enabled = True
    optr3.Enabled = False
    optr3.Value = False
    optr4.Value = False
    optr4.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = False
    optg3.Value = False
    optg4.Value = False
    optg4.Enabled = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = False
    optb3.Value = False
    optb4.Value = False
    optb4.Enabled = False
End If
End Sub

Private Sub opty4_Click()
If opt4.Value = True Then
    opty2.Value = False
    opty2.Enabled = False
    opty3.Value = False
    opty3.Enabled = False
    opty1.Value = False
    opty1.Enabled = False
    optr2.Enabled = True
    optr1.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = True
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = True
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = True
ElseIf opt3.Value = True Then
    opty2.Value = False
    opty2.Enabled = False
    opty3.Value = False
    opty3.Enabled = False
    opty1.Value = False
    opty1.Enabled = False
    optr2.Enabled = True
    optr1.Enabled = True
    optr3.Enabled = True
    optr4.Enabled = False
    optr4.Value = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = True
    optg4.Enabled = False
    optg4.Value = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = True
    optb4.Enabled = False
    optb4.Value = False
ElseIf opt2.Value = True Then
    opty2.Value = False
    opty2.Enabled = False
    opty3.Value = False
    opty3.Enabled = False
    opty1.Value = False
    opty1.Enabled = False
    optr2.Enabled = True
    optr1.Enabled = True
    optr3.Enabled = False
    optr3.Value = False
    optr4.Value = False
    optr4.Enabled = False
    optg1.Enabled = True
    optg2.Enabled = True
    optg3.Enabled = False
    optg3.Value = False
    optg4.Value = False
    optg4.Enabled = False
    optb1.Enabled = True
    optb2.Enabled = True
    optb3.Enabled = False
    optb3.Value = False
    optb4.Value = False
    optb4.Enabled = False
End If
End Sub
