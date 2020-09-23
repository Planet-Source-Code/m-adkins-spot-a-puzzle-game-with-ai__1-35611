VERSION 5.00
Begin VB.Form frmGrid 
   Caption         =   "Spot"
   ClientHeight    =   5715
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8865
   Icon            =   "frmGrid.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picNothing 
      Height          =   465
      Left            =   3150
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   161
      TabStop         =   0   'False
      Top             =   6300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picHappy 
      Height          =   465
      Left            =   2550
      Picture         =   "frmGrid.frx":0442
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   160
      TabStop         =   0   'False
      Top             =   6300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picYellow 
      Height          =   465
      Left            =   1995
      Picture         =   "frmGrid.frx":0851
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   159
      TabStop         =   0   'False
      Top             =   6300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picRed 
      Height          =   465
      Left            =   765
      Picture         =   "frmGrid.frx":0B8F
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   158
      TabStop         =   0   'False
      Top             =   6300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picGreen 
      Height          =   465
      Left            =   1380
      Picture         =   "frmGrid.frx":0F3F
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   157
      TabStop         =   0   'False
      Top             =   6300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picBlue 
      Height          =   465
      Left            =   150
      Picture         =   "frmGrid.frx":12A3
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   156
      TabStop         =   0   'False
      Top             =   6300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Timer tmrGrid 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   8400
      Top             =   0
   End
   Begin VB.Frame fraP4 
      Caption         =   "Player 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6000
      TabIndex        =   153
      Top             =   4350
      Width           =   2715
      Begin VB.Label lblName4 
         Caption         =   "Label1"
         Height          =   315
         Left            =   150
         TabIndex        =   155
         Top             =   300
         Width           =   2415
      End
      Begin VB.Label lblScore4 
         Caption         =   "Label1"
         Height          =   315
         Left            =   150
         TabIndex        =   154
         Top             =   750
         Width           =   2415
      End
   End
   Begin VB.Frame fraP3 
      Caption         =   "Player 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6000
      TabIndex        =   150
      Top             =   2950
      Width           =   2715
      Begin VB.Label lblScore3 
         Caption         =   "Label1"
         Height          =   315
         Left            =   150
         TabIndex        =   152
         Top             =   750
         Width           =   2415
      End
      Begin VB.Label lblName3 
         Caption         =   "Label1"
         Height          =   315
         Left            =   150
         TabIndex        =   151
         Top             =   300
         Width           =   2415
      End
   End
   Begin VB.Frame fraP2 
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6000
      TabIndex        =   147
      Top             =   1550
      Width           =   2715
      Begin VB.Label lblName2 
         Caption         =   "Label1"
         Height          =   315
         Left            =   150
         TabIndex        =   149
         Top             =   300
         Width           =   2415
      End
      Begin VB.Label lblScore2 
         Caption         =   "Label1"
         Height          =   315
         Left            =   150
         TabIndex        =   148
         Top             =   750
         Width           =   2415
      End
   End
   Begin VB.Frame fraP1 
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6000
      TabIndex        =   144
      Top             =   150
      Width           =   2715
      Begin VB.Label lblScore1 
         Caption         =   "Label1"
         Height          =   315
         Left            =   150
         TabIndex        =   146
         Top             =   750
         Width           =   2415
      End
      Begin VB.Label lblName1 
         Caption         =   "Label1"
         Height          =   315
         Left            =   150
         TabIndex        =   145
         Top             =   300
         Width           =   2415
      End
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   132
      Left            =   150
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   143
      TabStop         =   0   'False
      Tag             =   "1,12"
      Top             =   5100
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   134
      Left            =   1080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   142
      TabStop         =   0   'False
      Tag             =   "3,12"
      Top             =   5100
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   133
      Left            =   615
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   141
      TabStop         =   0   'False
      Tag             =   "2,12"
      Top             =   5100
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   135
      Left            =   1545
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   140
      TabStop         =   0   'False
      Tag             =   "4,12"
      Top             =   5100
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   137
      Left            =   2475
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   139
      TabStop         =   0   'False
      Tag             =   "6,12"
      Top             =   5100
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   136
      Left            =   2010
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   138
      TabStop         =   0   'False
      Tag             =   "5,12"
      Top             =   5100
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   138
      Left            =   2940
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   137
      TabStop         =   0   'False
      Tag             =   "7,12"
      Top             =   5100
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   140
      Left            =   3870
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   136
      TabStop         =   0   'False
      Tag             =   "9,12"
      Top             =   5100
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   139
      Left            =   3405
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   135
      TabStop         =   0   'False
      Tag             =   "8,12"
      Top             =   5100
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   141
      Left            =   4335
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   134
      TabStop         =   0   'False
      Tag             =   "10,12"
      Top             =   5100
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   143
      Left            =   5265
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   133
      TabStop         =   0   'False
      Tag             =   "12,12"
      Top             =   5100
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   142
      Left            =   4800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   132
      TabStop         =   0   'False
      Tag             =   "11,12"
      Top             =   5100
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   120
      Left            =   150
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   131
      TabStop         =   0   'False
      Tag             =   "1,11"
      Top             =   4650
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   122
      Left            =   1080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   130
      TabStop         =   0   'False
      Tag             =   "3,11"
      Top             =   4650
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   121
      Left            =   615
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   129
      TabStop         =   0   'False
      Tag             =   "2,11"
      Top             =   4650
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   123
      Left            =   1545
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   128
      TabStop         =   0   'False
      Tag             =   "4,11"
      Top             =   4650
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   125
      Left            =   2475
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   127
      TabStop         =   0   'False
      Tag             =   "6,11"
      Top             =   4650
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   124
      Left            =   2010
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   126
      TabStop         =   0   'False
      Tag             =   "5,11"
      Top             =   4650
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   126
      Left            =   2940
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   125
      TabStop         =   0   'False
      Tag             =   "7,11"
      Top             =   4650
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   128
      Left            =   3870
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   124
      TabStop         =   0   'False
      Tag             =   "9,11"
      Top             =   4650
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   127
      Left            =   3405
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   123
      TabStop         =   0   'False
      Tag             =   "8,11"
      Top             =   4650
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   129
      Left            =   4335
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   122
      TabStop         =   0   'False
      Tag             =   "10,11"
      Top             =   4650
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   131
      Left            =   5265
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   121
      TabStop         =   0   'False
      Tag             =   "12,11"
      Top             =   4650
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   130
      Left            =   4800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   120
      TabStop         =   0   'False
      Tag             =   "11,11"
      Top             =   4650
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   108
      Left            =   150
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   119
      TabStop         =   0   'False
      Tag             =   "1,10"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   110
      Left            =   1080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   118
      TabStop         =   0   'False
      Tag             =   "3,10"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   109
      Left            =   615
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   117
      TabStop         =   0   'False
      Tag             =   "2,10"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   111
      Left            =   1545
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   116
      TabStop         =   0   'False
      Tag             =   "4,10"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   113
      Left            =   2475
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   115
      TabStop         =   0   'False
      Tag             =   "6,10"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   112
      Left            =   2010
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   114
      TabStop         =   0   'False
      Tag             =   "5,10"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   114
      Left            =   2940
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   113
      TabStop         =   0   'False
      Tag             =   "7,10"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   116
      Left            =   3870
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   112
      TabStop         =   0   'False
      Tag             =   "9,10"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   115
      Left            =   3405
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   111
      TabStop         =   0   'False
      Tag             =   "8,10"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   117
      Left            =   4335
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   110
      TabStop         =   0   'False
      Tag             =   "10,10"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   119
      Left            =   5265
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   109
      TabStop         =   0   'False
      Tag             =   "12,10"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   118
      Left            =   4800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   108
      TabStop         =   0   'False
      Tag             =   "11,10"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   96
      Left            =   150
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   107
      TabStop         =   0   'False
      Tag             =   "1,9"
      Top             =   3750
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   98
      Left            =   1080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   106
      TabStop         =   0   'False
      Tag             =   "3,9"
      Top             =   3750
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   97
      Left            =   615
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   105
      TabStop         =   0   'False
      Tag             =   "2,9"
      Top             =   3750
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   99
      Left            =   1545
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   104
      TabStop         =   0   'False
      Tag             =   "4,9"
      Top             =   3750
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   101
      Left            =   2475
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   103
      TabStop         =   0   'False
      Tag             =   "6,9"
      Top             =   3750
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   100
      Left            =   2010
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   102
      TabStop         =   0   'False
      Tag             =   "5,9"
      Top             =   3750
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   102
      Left            =   2940
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   101
      TabStop         =   0   'False
      Tag             =   "7,9"
      Top             =   3750
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   104
      Left            =   3870
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   100
      TabStop         =   0   'False
      Tag             =   "9,9"
      Top             =   3750
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   103
      Left            =   3405
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   99
      TabStop         =   0   'False
      Tag             =   "8,9"
      Top             =   3750
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   105
      Left            =   4335
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   98
      TabStop         =   0   'False
      Tag             =   "10,9"
      Top             =   3750
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   107
      Left            =   5265
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   97
      TabStop         =   0   'False
      Tag             =   "12,9"
      Top             =   3750
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   106
      Left            =   4800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   96
      TabStop         =   0   'False
      Tag             =   "11,9"
      Top             =   3750
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   84
      Left            =   150
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   95
      TabStop         =   0   'False
      Tag             =   "1,8"
      Top             =   3300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   86
      Left            =   1080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   94
      TabStop         =   0   'False
      Tag             =   "3,8"
      Top             =   3300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   85
      Left            =   615
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   93
      TabStop         =   0   'False
      Tag             =   "2,8"
      Top             =   3300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   87
      Left            =   1545
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   92
      TabStop         =   0   'False
      Tag             =   "4,8"
      Top             =   3300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   89
      Left            =   2475
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   91
      TabStop         =   0   'False
      Tag             =   "6,8"
      Top             =   3300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   88
      Left            =   2010
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   90
      TabStop         =   0   'False
      Tag             =   "5,8"
      Top             =   3300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   90
      Left            =   2940
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   89
      TabStop         =   0   'False
      Tag             =   "7,8"
      Top             =   3300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   92
      Left            =   3870
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   88
      TabStop         =   0   'False
      Tag             =   "9,8"
      Top             =   3300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   91
      Left            =   3405
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   87
      TabStop         =   0   'False
      Tag             =   "8,8"
      Top             =   3300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   93
      Left            =   4335
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   86
      TabStop         =   0   'False
      Tag             =   "10,8"
      Top             =   3300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   95
      Left            =   5265
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   85
      TabStop         =   0   'False
      Tag             =   "12,8"
      Top             =   3300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   94
      Left            =   4800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   84
      TabStop         =   0   'False
      Tag             =   "11,8"
      Top             =   3300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   72
      Left            =   150
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   83
      TabStop         =   0   'False
      Tag             =   "1,7"
      Top             =   2850
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   74
      Left            =   1080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   82
      TabStop         =   0   'False
      Tag             =   "3,7"
      Top             =   2850
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   73
      Left            =   615
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   81
      TabStop         =   0   'False
      Tag             =   "2,7"
      Top             =   2850
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   75
      Left            =   1545
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   80
      TabStop         =   0   'False
      Tag             =   "4,7"
      Top             =   2850
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   77
      Left            =   2475
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   79
      TabStop         =   0   'False
      Tag             =   "6,7"
      Top             =   2850
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   76
      Left            =   2010
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   78
      TabStop         =   0   'False
      Tag             =   "5,7"
      Top             =   2850
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   78
      Left            =   2940
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   77
      TabStop         =   0   'False
      Tag             =   "7,7"
      Top             =   2850
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   80
      Left            =   3870
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   76
      TabStop         =   0   'False
      Tag             =   "9,7"
      Top             =   2850
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   79
      Left            =   3405
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   75
      TabStop         =   0   'False
      Tag             =   "8,7"
      Top             =   2850
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   81
      Left            =   4335
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   74
      TabStop         =   0   'False
      Tag             =   "10,7"
      Top             =   2850
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   83
      Left            =   5265
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   73
      TabStop         =   0   'False
      Tag             =   "12,7"
      Top             =   2850
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   82
      Left            =   4800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   72
      TabStop         =   0   'False
      Tag             =   "11,7"
      Top             =   2850
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   60
      Left            =   150
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   71
      TabStop         =   0   'False
      Tag             =   "1,6"
      Top             =   2400
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   62
      Left            =   1080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   70
      TabStop         =   0   'False
      Tag             =   "3,6"
      Top             =   2400
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   61
      Left            =   615
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   69
      TabStop         =   0   'False
      Tag             =   "2,6"
      Top             =   2400
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   63
      Left            =   1545
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   68
      TabStop         =   0   'False
      Tag             =   "4,6"
      Top             =   2400
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   65
      Left            =   2475
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   67
      TabStop         =   0   'False
      Tag             =   "6,6"
      Top             =   2400
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   64
      Left            =   2010
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   66
      TabStop         =   0   'False
      Tag             =   "5,6"
      Top             =   2400
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   66
      Left            =   2940
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   65
      TabStop         =   0   'False
      Tag             =   "7,6"
      Top             =   2400
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   68
      Left            =   3870
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   64
      TabStop         =   0   'False
      Tag             =   "9,6"
      Top             =   2400
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   67
      Left            =   3405
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   63
      TabStop         =   0   'False
      Tag             =   "8,6"
      Top             =   2400
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   69
      Left            =   4335
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   62
      TabStop         =   0   'False
      Tag             =   "10,6"
      Top             =   2400
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   71
      Left            =   5265
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   61
      TabStop         =   0   'False
      Tag             =   "12,6"
      Top             =   2400
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   70
      Left            =   4800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   60
      TabStop         =   0   'False
      Tag             =   "11,6"
      Top             =   2400
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   48
      Left            =   150
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   59
      TabStop         =   0   'False
      Tag             =   "1,5"
      Top             =   1950
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   50
      Left            =   1080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   58
      TabStop         =   0   'False
      Tag             =   "3,5"
      Top             =   1950
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   49
      Left            =   615
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   57
      TabStop         =   0   'False
      Tag             =   "2,5"
      Top             =   1950
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   51
      Left            =   1545
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   56
      TabStop         =   0   'False
      Tag             =   "4,5"
      Top             =   1950
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   53
      Left            =   2475
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   55
      TabStop         =   0   'False
      Tag             =   "6,5"
      Top             =   1950
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   52
      Left            =   2010
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   54
      TabStop         =   0   'False
      Tag             =   "5,5"
      Top             =   1950
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   54
      Left            =   2940
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   53
      TabStop         =   0   'False
      Tag             =   "7,5"
      Top             =   1950
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   56
      Left            =   3870
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   52
      TabStop         =   0   'False
      Tag             =   "9,5"
      Top             =   1950
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   55
      Left            =   3405
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   51
      TabStop         =   0   'False
      Tag             =   "8,5"
      Top             =   1950
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   57
      Left            =   4335
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   50
      TabStop         =   0   'False
      Tag             =   "10,5"
      Top             =   1950
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   59
      Left            =   5265
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   49
      TabStop         =   0   'False
      Tag             =   "12,5"
      Top             =   1950
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   58
      Left            =   4800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   48
      TabStop         =   0   'False
      Tag             =   "11,5"
      Top             =   1950
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   36
      Left            =   150
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   47
      TabStop         =   0   'False
      Tag             =   "1,4"
      Top             =   1500
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   38
      Left            =   1080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   46
      TabStop         =   0   'False
      Tag             =   "3,4"
      Top             =   1500
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   37
      Left            =   615
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   45
      TabStop         =   0   'False
      Tag             =   "2,4"
      Top             =   1500
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   39
      Left            =   1545
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   44
      TabStop         =   0   'False
      Tag             =   "4,4"
      Top             =   1500
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   41
      Left            =   2475
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   43
      TabStop         =   0   'False
      Tag             =   "6,4"
      Top             =   1500
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   40
      Left            =   2010
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   42
      TabStop         =   0   'False
      Tag             =   "5,4"
      Top             =   1500
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   42
      Left            =   2940
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   41
      TabStop         =   0   'False
      Tag             =   "7,4"
      Top             =   1500
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   44
      Left            =   3870
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   40
      TabStop         =   0   'False
      Tag             =   "9,4"
      Top             =   1500
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   43
      Left            =   3405
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   39
      TabStop         =   0   'False
      Tag             =   "8,4"
      Top             =   1500
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   45
      Left            =   4335
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   38
      TabStop         =   0   'False
      Tag             =   "10,4"
      Top             =   1500
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   47
      Left            =   5265
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   37
      TabStop         =   0   'False
      Tag             =   "12,4"
      Top             =   1500
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   46
      Left            =   4800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   36
      TabStop         =   0   'False
      Tag             =   "11,4"
      Top             =   1500
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   24
      Left            =   150
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   35
      TabStop         =   0   'False
      Tag             =   "1,3"
      Top             =   1050
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   26
      Left            =   1080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   34
      TabStop         =   0   'False
      Tag             =   "3,3"
      Top             =   1050
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   25
      Left            =   615
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   "2,3"
      Top             =   1050
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   27
      Left            =   1545
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   32
      TabStop         =   0   'False
      Tag             =   "4,3"
      Top             =   1050
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   29
      Left            =   2475
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "6,3"
      Top             =   1050
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   28
      Left            =   2010
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "5,3"
      Top             =   1050
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   30
      Left            =   2940
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   29
      TabStop         =   0   'False
      Tag             =   "7,3"
      Top             =   1050
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   32
      Left            =   3870
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "9,3"
      Top             =   1050
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   31
      Left            =   3405
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   27
      TabStop         =   0   'False
      Tag             =   "8,3"
      Top             =   1050
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   33
      Left            =   4335
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "10,3"
      Top             =   1050
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   35
      Left            =   5265
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   25
      TabStop         =   0   'False
      Tag             =   "12,3"
      Top             =   1050
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   34
      Left            =   4800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "11,3"
      Top             =   1050
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   12
      Left            =   150
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "1,2"
      Top             =   600
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   14
      Left            =   1080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "3,2"
      Top             =   600
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   13
      Left            =   615
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   "2,2"
      Top             =   600
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   15
      Left            =   1545
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   20
      TabStop         =   0   'False
      Tag             =   "4,2"
      Top             =   600
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   17
      Left            =   2475
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "6,2"
      Top             =   600
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   16
      Left            =   2010
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "5,2"
      Top             =   600
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   18
      Left            =   2940
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "7,2"
      Top             =   600
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   20
      Left            =   3870
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "9,2"
      Top             =   600
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   19
      Left            =   3405
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "8,2"
      Top             =   600
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   21
      Left            =   4335
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "10,2"
      Top             =   600
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   23
      Left            =   5265
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "12,2"
      Top             =   600
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   22
      Left            =   4800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "11,2"
      Top             =   600
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   10
      Left            =   4800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "11,1"
      Top             =   150
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   11
      Left            =   5265
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "12,1"
      Top             =   150
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   9
      Left            =   4335
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "10,1"
      Top             =   150
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   7
      Left            =   3405
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "8,1"
      Top             =   150
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   8
      Left            =   3870
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "9,1"
      Top             =   150
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   0
      Left            =   150
      ScaleHeight     =   405
      ScaleMode       =   0  'User
      ScaleWidth      =   405
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "1,1"
      Top             =   150
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   5
      Left            =   2475
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "6,1"
      Top             =   150
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   6
      Left            =   2940
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "7,1"
      Top             =   150
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   4
      Left            =   2010
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "5,1"
      Top             =   150
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   2
      Left            =   1080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "3,1"
      Top             =   150
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   3
      Left            =   1545
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "4,1"
      Top             =   150
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic0 
      Height          =   465
      Index           =   1
      Left            =   615
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "2,1"
      Top             =   150
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Spot"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&New Game"
      End
      Begin VB.Menu mnuGameLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnSquareClicked As Boolean
Private mstrClickedCoord As String
Private mbytIndex As Byte

'find X coordinate in "X,Y" system
Private Function CoordX(ByVal strCoord As String) As Byte
    Dim strX As String
    On Error Resume Next
    strX = strCoord
    Do
        If InStr(strX, ",") > 0 Then
            strX = Right(strX, Len(strX) - InStr(strX, ","))
        End If
    Loop Until InStr(strX, ",") = 0
    CoordX = CByte(Left(strCoord, Len(strCoord) - Len(strX)))
End Function

'find Y coordinate in "X,Y" system
Private Function CoordY(ByVal strCoord As String) As Byte
    Dim strY As String
    On Error Resume Next
    strY = strCoord
    Do
        If InStr(strY, ",") > 0 Then
            strY = Right(strY, Len(strY) - InStr(strY, ","))
        End If
    Loop Until InStr(strY, ",") = 0
    CoordY = CByte(strY)
End Function

'returns boolean specifying whether the passed coordinate system is 1 space
'adjacent to mstrClickedCoord
Private Function IsAdjacent(ByVal strCoord As String) As Boolean
    Dim pic As PictureBox
    IsAdjacent = False
    If Abs(CInt(CoordX(mstrClickedCoord)) - CInt(CoordX(strCoord))) <= 1 Then
        If Abs(CInt(CoordY(mstrClickedCoord)) - CInt(CoordY(strCoord))) <= 1 Then
            Set pic = FindBox(strCoord)
            If pic.Visible = True Then
                IsAdjacent = True
            End If
        End If
    End If
End Function

'returns the PictureBox object at the passed coordinate system
Private Function FindBox(ByVal strCoord As String) As PictureBox
    Dim bytIndex As Byte
    For bytIndex = 0 To 143
        If pic0(bytIndex).Tag = strCoord Then
            Set FindBox = pic0(bytIndex)
            Exit For
        End If
    Next bytIndex
End Function

'scans all CGame.Value settings for all players and refreshes screen accordingly
Private Sub RefreshSquares()
    Dim bytIndex As Byte
    Dim bytPlayer As Byte
    'set Player spots
    For bytPlayer = 1 To ggamGame.Players
        For bytIndex = 0 To 143
            If pic0(bytIndex).Tag <> "" Then
                If ggamGame.Value(CoordX(pic0(bytIndex).Tag), _
                        CoordY(pic0(bytIndex).Tag)) = bytPlayer Then
                    pic0(bytIndex).Picture = gpicPlayerImg(bytPlayer)
                End If
            End If
        Next bytIndex
    Next bytPlayer
    'set blank spots
    For bytIndex = 0 To 143
        If ggamGame.Value(CoordX(pic0(bytIndex).Tag), _
                CoordY(pic0(bytIndex).Tag)) = 0 Then
            pic0(bytIndex).Picture = picNothing.Picture
        End If
    Next bytIndex
End Sub

'captures spots surrounding space after legal move and advances to NextPlayer
Private Sub SetAdjacentValues()
    Dim bytIndex As Byte
    Dim bytVal As Byte
    For bytIndex = 0 To 143
        If IsAdjacent(pic0(bytIndex).Tag) Then
            bytVal = ggamGame.Value(CoordX(pic0(bytIndex).Tag), _
                    CoordY(pic0(bytIndex).Tag))
            If bytVal <> ggamGame.PlayerTurn And bytVal <> 0 Then
                ggamGame.Value CoordX(pic0(bytIndex).Tag), _
                        CoordY(pic0(bytIndex).Tag), ggamGame.PlayerTurn
            End If
        End If
    Next bytIndex
    tmrGrid.Enabled = False
    RefreshSquares
    mblnSquareClicked = False
    ggamGame.NextPlayer
    If DisplayScores = False Then
        Do While IsPossible(ggamGame.PlayerTurn) = False
            If ggamGame.IsHuman(ggamGame.PlayerTurn) = True Then _
                    MsgBox "No possible moves for " & _
                    gstrPlayerNames(ggamGame.PlayerTurn) & ".", _
                    vbInformation, "No Possible Moves"
            ggamGame.NextPlayer
        Loop
    Else
        Call mnuGameNew_Click
        Exit Sub
    End If
    Me.Caption = "Spot - " & gstrPlayerNames(ggamGame.PlayerTurn) & "'s Turn"
    If ggamGame.IsHuman(ggamGame.PlayerTurn) = False Then DoBestMove
End Sub

'returns boolean expressing legality of initial spot activate
Private Function ClickCheck() As Boolean
    ClickCheck = (CurrentClickValue = ggamGame.PlayerTurn)
End Function

'compares strCoord to mstrClickedCoord
'returns: 0 for illegal, 1 for clone, 2 for jump
Private Function MoveCheck(ByVal strCoord As String) As Byte
    Dim bytX As Byte, bytY As Byte
    Dim pic As PictureBox
    bytX = CoordX(strCoord)
    bytY = CoordY(strCoord)
    Set pic = FindBox(strCoord)
    If ggamGame.Value(bytX, bytY) <> 0 Or pic.Visible = False Then
        MoveCheck = 0
    ElseIf IsAdjacent(strCoord) Then
        MoveCheck = 1
    ElseIf Abs(CInt(CoordX(mstrClickedCoord)) - CInt(bytX)) <= 2 Then
        If Abs(CInt(CoordY(mstrClickedCoord)) - CInt(bytY)) <= 2 Then
            MoveCheck = 2
        End If
    Else
        MoveCheck = 0
    End If
End Function

'returns CGame.Value of coordinates in mstrClickedCoord
Private Function CurrentClickValue() As Byte
    CurrentClickValue = ggamGame.Value(CoordX(mstrClickedCoord), _
            CoordY(mstrClickedCoord))
End Function

'sets module-level variables specifying an active click and coordinates
'and activates timer
Private Function SetClick(ByVal strCoord As String, _
        Optional ByVal blnActivateClick As Boolean = True) As Boolean
    mstrClickedCoord = strCoord
    If ClickCheck = False Then
        Beep
'        MsgBox "This is not your spot.  Please select a spot that is yours.", _
'                vbInformation, "Not Your Spot"
        SetClick = False
        Exit Function
    End If
    mblnSquareClicked = blnActivateClick
    If blnActivateClick = True Then _
            FindBox(mstrClickedCoord).Picture = picNothing.Picture
    tmrGrid.Enabled = blnActivateClick
    SetClick = blnActivateClick
End Function

Private Sub Form_Load()
    Dim bytX As Byte, bytY As Byte, bytPic As Byte
    Dim picVis As PictureBox
    Randomize       'start random number generator for DoBestMove
    With ggamGame
        'show appropriate picture boxes
        For bytX = 1 To .MaxX
            For bytY = 1 To .MaxY
                Set picVis = FindBox(CStr(bytX) & "," & CStr(bytY))
                picVis.Visible = True
            Next bytY
        Next bytX
        'redimension player picture boxes
        ReDim gpicPlayerImg(1 To .Players)
        'set picture boxes for Players
        For bytPic = 1 To .Players
            Select Case gstrPlayerImg(bytPic)
                Case "Blue"
                    Set gpicPlayerImg(bytPic) = picBlue
                Case "Red"
                    Set gpicPlayerImg(bytPic) = picRed
                Case "Green"
                    Set gpicPlayerImg(bytPic) = picGreen
                Case "Yellow"
                    Set gpicPlayerImg(bytPic) = picYellow
                Case "Happy"
                    Set gpicPlayerImg(bytPic) = picHappy
            End Select
        Next bytPic
        'set Player names in labels
        lblName1 = gstrPlayerNames(1)
        lblName2 = gstrPlayerNames(2)
        lblName3 = gstrPlayerNames(3)
        lblName4 = gstrPlayerNames(4)
        'hide player frames as necessary
        If .Players < 3 Then fraP3.Visible = False
        If .Players < 4 Then fraP4.Visible = False
        'set initial locations and display spots
        .Value 1, 1, 1
        .Value .MaxX, .MaxY, 2
        If .Players > 2 Then .Value .MaxX, 1, 3
        If .Players > 3 Then .Value 1, .MaxY, 4
        mblnSquareClicked = False
        RefreshSquares
        DisplayScores
    End With
    Me.Caption = "Spot - " & lblName1.Caption & "'s Turn"
    If ggamGame.IsHuman(ggamGame.PlayerTurn) = False Then
        tmrGrid.Interval = 3000
        tmrGrid.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMenu.Show
    frmMenu.cboPlayers.SetFocus
End Sub

Private Sub mnuGameExit_Click()
    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next frm
    End
End Sub

Private Sub mnuGameNew_Click()
    If MsgBox("Do you want to start a new game?", vbQuestion + vbYesNo) = vbYes Then
        Unload frmGrid
    End If
End Sub

Private Sub pic0_Click(Index As Integer)
    Call AllClicks(pic0(Index).Tag)
End Sub

Private Sub tmrGrid_Timer()
    Static blnFlash As Boolean
    Dim picFlash As PictureBox
    Set picFlash = FindBox(mstrClickedCoord)
    If tmrGrid.Interval = 3000 Then
        tmrGrid.Enabled = False
        tmrGrid.Interval = 400
        DoBestMove
    Else
        If blnFlash = True Then
            picFlash.Picture = picNothing.Picture
            blnFlash = False
        Else
            picFlash.Picture = gpicPlayerImg(ggamGame.PlayerTurn).Picture
            blnFlash = True
        End If
    End If
    Exit Sub
End Sub

'returns True if game is over, False if game should continue
Private Function DisplayScores() As Boolean
    Dim bytPlayers As Byte, bytNoScoreCount As Byte, strWinner As String
    Dim bytHighScore As Byte, bytTotalScores As Byte
    Dim blnTieGame As Boolean
    
    bytNoScoreCount = 0
    bytHighScore = 0
    bytTotalScores = 0
    blnTieGame = False
    With ggamGame
        lblScore1.Caption = .Score(1)
        lblScore2.Caption = .Score(2)
        If lblScore3.Visible = True Then lblScore3.Caption = .Score(3)
        If lblScore4.Visible = True Then lblScore4.Caption = .Score(4)
        
        For bytPlayers = 1 To .Players
            If .Score(bytPlayers) = 0 Then
                bytNoScoreCount = bytNoScoreCount + 1
            Else
                If .Score(bytPlayers) > bytHighScore Then
                    bytHighScore = .Score(bytPlayers)
                    strWinner = gstrPlayerNames(bytPlayers)
                ElseIf .Score(bytPlayers) = bytHighScore Then
                    strWinner = strWinner & " and " & gstrPlayerNames(bytPlayers)
                    blnTieGame = True
                End If
                bytTotalScores = bytTotalScores + .Score(bytPlayers)
            End If
        Next bytPlayers
        If (bytNoScoreCount = .Players - 1) Or (bytTotalScores = .MaxX * .MaxY) Then
            MsgBox strWinner & " wins!", vbInformation, _
                    IIf(blnTieGame = False, "We Have A Winner!", "We Have A Tie!")
            DisplayScores = True
        Else
            DisplayScores = False
        End If
    End With
End Function

Private Sub AllClicks(ByVal strCoord As String)
    Dim bytX As Byte, bytY As Byte
    bytX = CoordX(strCoord)
    bytY = CoordY(strCoord)
    Select Case mblnSquareClicked
        Case False
            If SetClick(strCoord) = False Then Exit Sub
        Case True
            If strCoord = mstrClickedCoord Then
                mblnSquareClicked = False
                tmrGrid.Enabled = False
                RefreshSquares
                Exit Sub
            End If
            Select Case MoveCheck(strCoord)
                Case 0
                    Beep
                    'MsgBox "Cannot move here.  Please try again.", vbInformation, _
                    '        "Illegal Move"
                    Exit Sub
                Case 1
                    ggamGame.Value bytX, bytY, ggamGame.PlayerTurn
                Case 2
                    ggamGame.Value bytX, bytY, ggamGame.PlayerTurn
                    ggamGame.Value CoordX(mstrClickedCoord), _
                            CoordY(mstrClickedCoord), 0
            End Select
        mstrClickedCoord = strCoord
        SetAdjacentValues
    End Select
End Sub

'cycle through all possible From squares and compare to all possible To squares
'return true if at least one possible move is found
Private Function IsPossible(ByVal bytPlayer As Byte) As Boolean
    Dim bytFrom As Byte, bytTo As Byte
    IsPossible = False
    For bytFrom = 0 To 143
        With pic0(bytFrom)
            If ggamGame.Value(CoordX(.Tag), CoordY(.Tag)) = bytPlayer Then
                mstrClickedCoord = .Tag
                For bytTo = 0 To 143
                    If MoveCheck(pic0(bytTo).Tag) <> 0 Then
                        IsPossible = True
                        Exit Function
                    End If
                Next bytTo
            End If
        End With
    Next bytFrom
End Function

'return total number of squares adjacent to the current square that belong to
'the current player (used for computer-controlled players)
Private Function SquaresOpenIfJump() As Byte
    Dim bytCounter As Byte
    Dim bytSquares As Byte
    bytSquares = 0
    For bytCounter = 0 To 143
        With pic0(bytCounter)
            If IsAdjacent(.Tag) And ggamGame.Value(CoordX(.Tag), CoordY(.Tag)) = _
                    ggamGame.PlayerTurn And .Tag <> mstrClickedCoord Then _
                    bytSquares = bytSquares + 1
        End With
    Next bytCounter
    SquaresOpenIfJump = bytSquares
End Function

'return total number of squares that would be taken by the current player if a
'square was taken (used for computer-controlled players)
'strCoord signifies the coordinates of the square to be "taken"
Private Function SquaresTakenIfJump(ByVal strCoord As String) As Byte
    Dim bytCounter As Byte
    Dim bytSquares As Byte
    Dim strCurrentCoord As String
    bytSquares = 0
    strCurrentCoord = mstrClickedCoord      'save current coordinates
    For bytCounter = 0 To 143
        With pic0(bytCounter)
            mstrClickedCoord = strCoord     'temporarily set module-level variable
            If IsAdjacent(.Tag) And ggamGame.Value(CoordX(.Tag), CoordY(.Tag)) _
                    <> ggamGame.PlayerTurn And _
                    ggamGame.Value(CoordX(.Tag), CoordY(.Tag)) <> 0 _
                    And .Tag <> mstrClickedCoord Then _
                    bytSquares = bytSquares + 1
        End With
    Next bytCounter
    mstrClickedCoord = strCurrentCoord      'restore original value
    SquaresTakenIfJump = bytSquares
End Function

'calculates the potential net points earned or lost from a single move
'strCoord signifies square to be "taken", and is passed to SquaresTakenIfJump
'mstrClickedCoord is the "from" square, and strCoord is the "to" sqare
'this function is used for computer-controlled players
Private Function NetPointChange(ByVal strCoord As String) As Integer
    NetPointChange = CInt(SquaresTakenIfJump(strCoord)) - CInt(SquaresOpenIfJump)
End Function

'find the best move and do it (used for computer-controlled players)
Private Sub DoBestMove()
    Dim strFrom As String
    Dim strTo As String
    Dim strFromBest As String
    Dim strToBest As String
    
    Dim intPointsHigh As Integer
    Dim intPointsCurrent As Integer
    
    Dim bytStart(1) As Integer
    Dim bytFinish(1) As Integer
    Dim intStep(1) As Integer
    Dim bytFromCounter As Integer
    Dim bytToCounter As Integer
    
    Screen.MousePointer = vbHourglass
    
    'find move search order
    '(top to bottom or vice-versa; left to right or vice-versa)
    If CInt(Rnd * 1) = 0 Then
        bytStart(0) = 0
        bytFinish(0) = 143
        intStep(0) = 1
    Else
        bytStart(0) = 143
        bytFinish(0) = 0
        intStep(0) = -1
    End If
    If CInt(Rnd * 1) = 0 Then
        bytStart(1) = 0
        bytFinish(1) = 143
        intStep(1) = 1
    Else
        bytStart(1) = 143
        bytFinish(1) = 0
        intStep(1) = -1
    End If
    
    intPointsHigh = -9  '-8 is the worst possible, so start below that
        
    For bytFromCounter = bytStart(0) To bytFinish(0) Step intStep(0)
        strFrom = pic0(bytFromCounter).Tag
        If CoordX(strFrom) <= ggamGame.MaxX And _
                CoordY(strFrom) <= ggamGame.MaxY And _
                ggamGame.Value(CoordX(strFrom), CoordY(strFrom)) = _
                ggamGame.PlayerTurn Then
            mstrClickedCoord = strFrom
            For bytToCounter = bytStart(1) To bytFinish(1) Step intStep(1)
                '-8 is the worst possible, so start below that
                intPointsCurrent = -9
                strTo = pic0(bytToCounter).Tag
                If CoordX(strTo) <= ggamGame.MaxX And _
                        CoordY(strTo) <= ggamGame.MaxY Then
                    Select Case MoveCheck(strTo)
                        Case 0      'illegal
                            'do nothing
                        Case 1      'clone
                            intPointsCurrent = 1 + CInt(SquaresTakenIfJump(strTo))
                            If intPointsCurrent >= intPointsHigh Then
                                intPointsHigh = intPointsCurrent
                                strFromBest = strFrom
                                strToBest = strTo
                            End If
                        Case 2      'jump
                            intPointsCurrent = NetPointChange(strTo)
                            If intPointsCurrent >= intPointsHigh Then
                                intPointsHigh = intPointsCurrent
                                strFromBest = strFrom
                                strToBest = strTo
                            End If
                    End Select
                End If
            Next bytToCounter
        End If
    Next bytFromCounter
    
    Call AllClicks(strFromBest)
    Call AllClicks(strToBest)
    
    Screen.MousePointer = vbDefault
End Sub
