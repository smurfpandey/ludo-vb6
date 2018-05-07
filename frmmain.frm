VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9165
   ClientLeft      =   3030
   ClientTop       =   1080
   ClientWidth     =   9150
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Moveable        =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   9150
   Begin VB.Timer Timer4 
      Interval        =   5000
      Left            =   480
      Top             =   6600
   End
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   7680
      Top             =   7320
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   7920
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   840
      Top             =   1560
   End
   Begin VB.TextBox txtp3val 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   100
      Top             =   8640
      Width           =   1215
   End
   Begin VB.TextBox txtp4val 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   99
      Top             =   8640
      Width           =   1215
   End
   Begin VB.TextBox txtp2val 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   98
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox icoblu 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DragMode        =   1  'Automatic
      Height          =   540
      Index           =   0
      Left            =   2520
      Picture         =   "frmmain.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   97
      Top             =   6600
      Width           =   540
   End
   Begin VB.PictureBox icoblu 
      AutoSize        =   -1  'True
      DragMode        =   1  'Automatic
      Height          =   540
      Index           =   1
      Left            =   1920
      Picture         =   "frmmain.frx":04B4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   96
      Top             =   6600
      Width           =   540
   End
   Begin VB.PictureBox icoblu 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      DragMode        =   1  'Automatic
      Height          =   540
      Index           =   2
      Left            =   2520
      Picture         =   "frmmain.frx":0968
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   95
      Top             =   6000
      Width           =   540
   End
   Begin VB.PictureBox icoblu 
      AutoSize        =   -1  'True
      DragMode        =   1  'Automatic
      Height          =   540
      Index           =   3
      Left            =   1920
      Picture         =   "frmmain.frx":0E1C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   94
      Top             =   6000
      Width           =   540
   End
   Begin VB.PictureBox icoyel 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      DragMode        =   1  'Automatic
      Height          =   540
      Index           =   0
      Left            =   6120
      Picture         =   "frmmain.frx":12D0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   93
      Top             =   6000
      Width           =   540
   End
   Begin VB.PictureBox icoyel 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      DragMode        =   1  'Automatic
      Height          =   540
      Index           =   1
      Left            =   6720
      Picture         =   "frmmain.frx":15DA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   92
      Top             =   6000
      Width           =   540
   End
   Begin VB.PictureBox icoyel 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      DragMode        =   1  'Automatic
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Index           =   2
      Left            =   6720
      Picture         =   "frmmain.frx":18E4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   91
      Top             =   6600
      Width           =   540
   End
   Begin VB.PictureBox icoyel 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      DragMode        =   1  'Automatic
      Height          =   540
      Index           =   3
      Left            =   6120
      Picture         =   "frmmain.frx":1BEE
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   90
      Top             =   6600
      Width           =   540
   End
   Begin VB.PictureBox icogre 
      AutoSize        =   -1  'True
      DragMode        =   1  'Automatic
      Height          =   540
      Index           =   1
      Left            =   6720
      Picture         =   "frmmain.frx":1EF8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   89
      Top             =   2040
      Width           =   540
   End
   Begin VB.PictureBox icogre 
      AutoSize        =   -1  'True
      DragMode        =   1  'Automatic
      Height          =   540
      Index           =   3
      Left            =   6120
      Picture         =   "frmmain.frx":2432
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   88
      Top             =   2640
      Width           =   540
   End
   Begin VB.PictureBox icogre 
      AutoSize        =   -1  'True
      DragMode        =   1  'Automatic
      Height          =   540
      Index           =   0
      Left            =   6120
      Picture         =   "frmmain.frx":296C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   87
      Top             =   2040
      Width           =   540
   End
   Begin VB.PictureBox icogre 
      AutoSize        =   -1  'True
      DragMode        =   1  'Automatic
      Height          =   540
      Index           =   2
      Left            =   6720
      Picture         =   "frmmain.frx":2EA6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   86
      Top             =   2640
      Width           =   540
   End
   Begin VB.PictureBox icored 
      AutoSize        =   -1  'True
      DragMode        =   1  'Automatic
      Height          =   540
      Index           =   3
      Left            =   2520
      Picture         =   "frmmain.frx":33E0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   85
      Top             =   2640
      Width           =   540
   End
   Begin VB.PictureBox icored 
      AutoSize        =   -1  'True
      DragMode        =   1  'Automatic
      Height          =   540
      Index           =   2
      Left            =   1920
      Picture         =   "frmmain.frx":383D
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   84
      Top             =   2640
      Width           =   540
   End
   Begin VB.PictureBox icored 
      AutoSize        =   -1  'True
      DragMode        =   1  'Automatic
      Height          =   540
      Index           =   0
      Left            =   1920
      Picture         =   "frmmain.frx":3C9A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   83
      Top             =   2040
      Width           =   540
   End
   Begin VB.PictureBox icored 
      AutoSize        =   -1  'True
      DragMode        =   1  'Automatic
      Height          =   540
      Index           =   1
      Left            =   2520
      Picture         =   "frmmain.frx":40F7
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   82
      Top             =   2040
      Width           =   540
   End
   Begin VB.PictureBox picgre 
      BackColor       =   &H0000C000&
      Height          =   495
      Index           =   1
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   81
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox picgre 
      BackColor       =   &H0000C000&
      Height          =   495
      Index           =   2
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   80
      Top             =   1920
      Width           =   495
   End
   Begin VB.PictureBox picgre 
      BackColor       =   &H0000C000&
      Height          =   495
      Index           =   3
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   79
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox picgre 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   4
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   78
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox picblu 
      BackColor       =   &H00FF0000&
      Height          =   495
      Index           =   4
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   77
      Top             =   5520
      Width           =   495
   End
   Begin VB.PictureBox picblu 
      BackColor       =   &H00FF0000&
      Height          =   495
      Index           =   3
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   76
      Top             =   6120
      Width           =   495
   End
   Begin VB.PictureBox picblu 
      BackColor       =   &H00FF0000&
      Height          =   495
      Index           =   2
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   75
      Top             =   6720
      Width           =   495
   End
   Begin VB.PictureBox picblu 
      BackColor       =   &H00FF0000&
      Height          =   495
      Index           =   1
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   74
      Top             =   7320
      Width           =   495
   End
   Begin VB.PictureBox picblu 
      BackColor       =   &H00FF0000&
      Height          =   495
      Index           =   0
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   73
      Top             =   7920
      Width           =   495
   End
   Begin VB.PictureBox picyel 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Index           =   0
      Left            =   7920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   72
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox picyel 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Index           =   1
      Left            =   7320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   71
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox picyel 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Index           =   2
      Left            =   6720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   70
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox picyel 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Index           =   3
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   69
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox picyel 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Index           =   4
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   68
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   24
      Left            =   8520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   67
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   25
      Left            =   8520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   66
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   26
      Left            =   7920
      Picture         =   "frmmain.frx":4554
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   65
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   27
      Left            =   7320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   64
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   28
      Left            =   6720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   63
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   29
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   62
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   30
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   61
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   31
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   60
      Top             =   5520
      Width           =   495
   End
   Begin VB.PictureBox picred 
      BackColor       =   &H000000C0&
      Height          =   495
      Index           =   0
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   59
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox picred 
      BackColor       =   &H000000C0&
      Height          =   495
      Index           =   1
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   58
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox picred 
      BackColor       =   &H000000C0&
      Height          =   495
      Index           =   2
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   57
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox picred 
      BackColor       =   &H000000C0&
      Height          =   495
      Index           =   3
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   56
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox picred 
      BackColor       =   &H000000C0&
      Height          =   495
      Index           =   4
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   55
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox picgre 
      BackColor       =   &H0000C000&
      Height          =   495
      Index           =   0
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   54
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   6
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   53
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   10
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   52
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   9
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   51
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   8
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   50
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   7
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   49
      Top             =   1920
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   11
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   48
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   12
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   47
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   13
      Left            =   4920
      Picture         =   "frmmain.frx":4E1E
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   46
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   14
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   45
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   15
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   44
      Top             =   1920
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   16
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   43
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   17
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   42
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   19
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   41
      Top             =   3720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   33
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   40
      Top             =   6720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   32
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   39
      Top             =   6120
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   18
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   38
      Top             =   3720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   20
      Left            =   6720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   37
      Top             =   3720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   21
      Left            =   7320
      Picture         =   "frmmain.frx":56E8
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   36
      Top             =   3720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   22
      Left            =   7920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   35
      Top             =   3720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   23
      Left            =   8520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   34
      Top             =   3720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   1
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   33
      Top             =   3720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   0
      Left            =   720
      Picture         =   "frmmain.frx":5FB2
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   32
      Top             =   3720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   2
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   31
      Top             =   3720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   3
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   30
      Top             =   3720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   4
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   29
      Top             =   3720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   5
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   28
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   50
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   27
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   490
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   26
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   480
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   25
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   470
      Left            =   1320
      Picture         =   "frmmain.frx":687C
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   24
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   460
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   23
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   450
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   22
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   440
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   21
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   42
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   20
      Top             =   6120
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   43
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   19
      Top             =   5520
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   41
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   18
      Top             =   6720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   40
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   17
      Top             =   7320
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   39
      Left            =   3720
      Picture         =   "frmmain.frx":7146
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   16
      Top             =   7920
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   38
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   15
      Top             =   8520
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   37
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   14
      Top             =   8520
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   36
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   8520
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   35
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   7920
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   34
      Left            =   4920
      Picture         =   "frmmain.frx":7A10
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   7320
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   495
      Index           =   51
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox txtp1val 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdp4click 
      BackColor       =   &H00FF0000&
      Caption         =   "Player4 Click Me"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton cmdp2click 
      BackColor       =   &H0000C000&
      Caption         =   "Player2 Click Me"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdp3click 
      BackColor       =   &H0000FFFF&
      Caption         =   "Player3Click Me"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton cmdp1click 
      BackColor       =   &H000000C0&
      Caption         =   "Player1 Click Me"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.Frame f4 
      BackColor       =   &H80000007&
      Caption         =   "Player4"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Left            =   0
      TabIndex        =   2
      Top             =   5760
      Width           =   3375
   End
   Begin VB.Frame f3 
      BackColor       =   &H80000007&
      Caption         =   "Player3"
      ForeColor       =   &H0000FFFF&
      Height          =   3375
      Left            =   5760
      TabIndex        =   1
      Top             =   5760
      Width           =   3375
   End
   Begin VB.Frame f2 
      BackColor       =   &H80000007&
      Caption         =   "Player2"
      ForeColor       =   &H0000C000&
      Height          =   3375
      Left            =   5760
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
   Begin VB.Frame f1 
      BackColor       =   &H80000007&
      Caption         =   "Player1"
      ForeColor       =   &H000000FF&
      Height          =   3375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3375
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   3720
      Picture         =   "frmmain.frx":82DA
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1695
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p1name, p2name, p3name, p4name As String

Private Sub cmdExit_Click()
MsgBox "Okay bye bye, hope you all might have enjoyed."
End
End Sub

Private Sub cmdP1click_Click()
If b = True And c = True And d = True And a = False Then
txtp1val.Text = ""
p1val = Int((6 - 1 + 1) * Rnd + 1)
txtp1val.Text = p1val
a = True
b = False
Select Case p1val
Case 1
    MsgBox "It's one. Now " + p2name + " it's your chance."
Case 2
    MsgBox "It's two. Now " + p2name + " it's your chance."
Case 3
    MsgBox "It's three. Now " + p2name + " it's your chance."
Case 4
    MsgBox "It's four. Now " + p2name + " it's your chance."
Case 5
    MsgBox "It's five. Now " + p2name + " it's your chance."
Case 6
    MsgBox "Wow it's six. " + p1name + " it's your chance once again."
    txtp1val.Text = " "
    a = False
    b = True
End Select
If cmdp2click.Enabled = False Then
b = True
End If
Else
    MsgBox "It's not your turn " + p1name + "."
End If
End Sub

Private Sub cmdp2click_Click()
If a And c And d = True And b = False Then
txtp2val.Text = ""
p1val = Int((6 - 1 + 1) * Rnd + 1)
txtp2val.Text = p1val
b = True
c = False
Select Case p1val
Case 1
    MsgBox "It's one. Now " + p3name + " it's your chance."
Case 2
    MsgBox "It's two. Now " + p3name + " it's your chance."
Case 3
    MsgBox "It's three. Now " + p3name + " it's your chance."
Case 4
    MsgBox "It's four. Now " + p3name + " it's your chance."
Case 5
    MsgBox "It's five. Now " + p3name + " it's your chance."
Case 6
    MsgBox " Wow it's six. " + p2name + " it's your chance once again."
    txtp2val.Text = " "
    b = False
    a = True
End Select
Else
    MsgBox "It's not your turn " + p2name + "."
End If
End Sub

Private Sub cmdp3click_Click()
If a And b And d = True And c = False Then
txtp3val.Text = ""
p1val = Int((6 - 1 + 1) * Rnd + 1)
txtp3val.Text = p1val
c = True
d = False
Select Case p1val
Case 1
    MsgBox "It's one. Now " + p4name + " it's your chance."
Case 2
    MsgBox "It's two. Now " + p4name + " it's your chance."
Case 3
    MsgBox "It's three. Now " + p4name + " it's your chance."
Case 4
    MsgBox "It's four. Now " + p4name + " it's your chance."
Case 5
    MsgBox "It's five. Now " + p4name + " it's your chance."
Case 6
    MsgBox "Wow it's six. " + p3name + " it's your chance once again."
    txtp3val.Text = " "
    c = False
    b = True
End Select
Else
    MsgBox "It's not your turn " + p3name + "."
End If
End Sub

Private Sub cmdp4click_Click()
If a And c And b = True And d = False Then
txtp4val.Text = ""
p1val = Int((6 - 1 + 1) * Rnd + 1)
txtp4val.Text = p1val
d = True
a = False
Select Case p1val
Case 1
    MsgBox "It's one. Now " + p1name + " it's your chance."
Case 2
    MsgBox "It's two. Now " + p1name + " it's your chance."
Case 3
    MsgBox "It's three. Now " + p1name + " it's your chance."
Case 4
    MsgBox "It's four. Now " + p1name + " it's your chance."
Case 5
    MsgBox "It's five. Now " + p1name + " it's your chance."
Case 6
    MsgBox "Wow it's six. Now " + p4name + " it's your chance once again."
    txtp4val.Text = " "
    d = False
    c = True
End Select
Else
    MsgBox "It's not your turn " + p4name + "."
End If
End Sub

Private Sub Form_Load()
If frmname.optr1.Value = True Then
p1name = Trim(frmname.txtp1name.Text)
ElseIf frmname.optr2.Value = True Then
p1name = Trim(frmname.txtp2name.Text)
ElseIf frmname.optr3.Value = True Then
p1name = Trim(frmname.txtp3name.Text)
ElseIf frmname.optr4.Value = True Then
p1name = Trim(frmname.txtp4name.Text)
Else
f1.Enabled = False
cmdp1click.Enabled = False
End If
If frmname.optg1.Value = True Then
p2name = Trim(frmname.txtp1name.Text)
ElseIf frmname.optg2.Value = True Then
p2name = Trim(frmname.txtp2name.Text)
ElseIf frmname.optg3.Value = True Then
p2name = Trim(frmname.txtp3name.Text)
ElseIf frmname.optg4.Value = True Then
p1name = Trim(frmname.txtp4name.Text)
Else
cmdp2click.Enabled = False
f2.Enabled = False
End If
If frmname.opty1.Value = True Then
p3name = Trim(frmname.txtp1name.Text)
ElseIf frmname.opty2.Value = True Then
p3name = Trim(frmname.txtp2name.Text)
ElseIf frmname.opty3.Value = True Then
p3name = Trim(frmname.txtp3name.Text)
ElseIf frmname.opty4.Value = True Then
p3name = Trim(frmname.txtp4name.Text)
Else
cmdp3click.Enabled = False
f3.Enabled = False
End If
If frmname.optb1.Value = True Then
p4name = Trim(frmname.txtp1name.Text)
ElseIf frmname.optb2.Value = True Then
p4name = Trim(frmname.txtp2name.Text)
ElseIf frmname.optb3.Value = True Then
p4name = Trim(frmname.txtp3name.Text)
ElseIf frmname.optb4.Value = True Then
p4name = Trim(frmname.txtp4name.Text)
Else
cmdp4click.Enabled = False
f4.Enabled = False
End If
f1.Caption = p1name
f2.Caption = p2name
f3.Caption = p3name
f4.Caption = p4name
cmdp1click.Caption = p1name + " click me"
cmdp2click.Caption = p2name + " click me"
cmdp3click.Caption = p3name + " click me"
cmdp4click.Caption = p4name + " click me"
frmmain.Show
a = False
b = True
c = True
d = True
If cmdp1click.Enabled = False Then
a = True
b = False
End If
End Sub

Private Sub icored_Click(Index As Integer)
icored(Index).Drag
End Sub

Private Sub Timer1_Timer()
Dim i As Integer
i = i + 1
If i = 1 Then
txtp1val.Text = " "
End If
End Sub

Private Sub Timer2_Timer()
Dim w As Integer
w = w + 1
If w = 1 Then
txtp2val.Text = " "
End If
End Sub

Private Sub Timer3_Timer()
Dim q As Integer
q = q + 1
If q = 1 Then
txtp3val.Text = " "
End If
End Sub

Private Sub Timer4_Timer()
Dim r As Integer
r = r + 1
If r = 1 Then
txtp4val.Text = " "
End If
End Sub
