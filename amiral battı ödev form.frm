VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   ScaleHeight     =   9765
   ScaleWidth      =   14670
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List11 
      Height          =   5910
      Left            =   6840
      TabIndex        =   126
      Top             =   3240
      Width           =   1455
   End
   Begin VB.ListBox List10 
      Height          =   5910
      Left            =   8520
      TabIndex        =   125
      Top             =   3240
      Width           =   855
   End
   Begin VB.ListBox List9 
      Height          =   6105
      Left            =   12720
      TabIndex        =   124
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "yeniden baþla"
      Height          =   1215
      Left            =   12720
      TabIndex        =   123
      Top             =   720
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "gemilerin"
      Height          =   3015
      Left            =   9840
      TabIndex        =   112
      Top             =   3480
      Width           =   1935
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         Caption         =   "teðmen"
         Height          =   375
         Left            =   720
         TabIndex        =   122
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         Caption         =   "üstteðmen2"
         Height          =   375
         Left            =   960
         TabIndex        =   121
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         Caption         =   "üstteðmen1"
         Height          =   375
         Left            =   360
         TabIndex        =   120
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         Caption         =   "yüzbaþý3"
         Height          =   375
         Left            =   1320
         TabIndex        =   119
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         Caption         =   "yüzbaþý2"
         Height          =   375
         Left            =   720
         TabIndex        =   118
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         Caption         =   "yüzbaþý1"
         Height          =   375
         Left            =   120
         TabIndex        =   117
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         Caption         =   "amiral4"
         Height          =   375
         Left            =   720
         TabIndex        =   116
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   "amiral3"
         Height          =   375
         Left            =   1320
         TabIndex        =   115
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "amiral2"
         Height          =   375
         Left            =   720
         TabIndex        =   114
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "amiral1"
         Height          =   375
         Left            =   120
         TabIndex        =   113
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   3960
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   480
      Width           =   2415
   End
   Begin VB.ListBox List3 
      Height          =   840
      Left            =   10200
      TabIndex        =   10
      Top             =   480
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   8880
      TabIndex        =   9
      Top             =   480
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   7560
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "konumgirme"
      Height          =   855
      Left            =   600
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "bot konum ata"
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "baþla"
      Height          =   735
      Left            =   600
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ListBox List8 
      Height          =   1035
      Left            =   11400
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List7 
      Height          =   840
      Left            =   10320
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List6 
      Height          =   840
      Left            =   9000
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List5 
      Height          =   1035
      Left            =   7440
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List4 
      Height          =   1035
      Left            =   11280
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "100"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   100
      Left            =   6000
      TabIndex        =   111
      Top             =   8880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "99"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   99
      Left            =   5400
      TabIndex        =   110
      Top             =   8880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "98"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   98
      Left            =   4800
      TabIndex        =   109
      Top             =   8880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "97"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   97
      Left            =   4200
      TabIndex        =   108
      Top             =   8880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "96"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   96
      Left            =   3600
      TabIndex        =   107
      Top             =   8880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "95"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   95
      Left            =   3000
      TabIndex        =   106
      Top             =   8880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "94"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   94
      Left            =   2400
      TabIndex        =   105
      Top             =   8880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "93"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   93
      Left            =   1800
      TabIndex        =   104
      Top             =   8880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "92"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   92
      Left            =   1200
      TabIndex        =   103
      Top             =   8880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "91"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   91
      Left            =   600
      TabIndex        =   102
      Top             =   8880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "90"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   90
      Left            =   6000
      TabIndex        =   101
      Top             =   8280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "89"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   89
      Left            =   5400
      TabIndex        =   100
      Top             =   8280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "88"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   88
      Left            =   4800
      TabIndex        =   99
      Top             =   8280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "87"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   87
      Left            =   4200
      TabIndex        =   98
      Top             =   8280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "86"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   86
      Left            =   3600
      TabIndex        =   97
      Top             =   8280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "85"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   85
      Left            =   3000
      TabIndex        =   96
      Top             =   8280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "84"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   84
      Left            =   2400
      TabIndex        =   95
      Top             =   8280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "83"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   83
      Left            =   1800
      TabIndex        =   94
      Top             =   8280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "82"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   82
      Left            =   1200
      TabIndex        =   93
      Top             =   8280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "81"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   81
      Left            =   600
      TabIndex        =   92
      Top             =   8280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "80"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   80
      Left            =   6000
      TabIndex        =   91
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "79"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   79
      Left            =   5400
      TabIndex        =   90
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "78"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   78
      Left            =   4800
      TabIndex        =   89
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "77"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   77
      Left            =   4200
      TabIndex        =   88
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "76"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   76
      Left            =   3600
      TabIndex        =   87
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "75"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   75
      Left            =   3000
      TabIndex        =   86
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "74"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   74
      Left            =   2400
      TabIndex        =   85
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "73"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   73
      Left            =   1800
      TabIndex        =   84
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "72"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   72
      Left            =   1200
      TabIndex        =   83
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "71"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   71
      Left            =   600
      LinkTimeout     =   0
      TabIndex        =   82
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "70"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   70
      Left            =   6000
      TabIndex        =   81
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "69"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   69
      Left            =   5400
      TabIndex        =   80
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "68"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   68
      Left            =   4800
      TabIndex        =   79
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "67"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   67
      Left            =   4200
      TabIndex        =   78
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "66"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   66
      Left            =   3600
      TabIndex        =   77
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "65"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   65
      Left            =   3000
      TabIndex        =   76
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "64"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   64
      Left            =   2400
      TabIndex        =   75
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "63"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   63
      Left            =   1800
      TabIndex        =   74
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "62"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   62
      Left            =   1200
      TabIndex        =   73
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "61"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   61
      Left            =   600
      TabIndex        =   72
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "60"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   60
      Left            =   6000
      TabIndex        =   71
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "59"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   59
      Left            =   5400
      TabIndex        =   70
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "58"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   58
      Left            =   4800
      TabIndex        =   69
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "57"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   57
      Left            =   4200
      TabIndex        =   68
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "56"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   56
      Left            =   3600
      TabIndex        =   67
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "55"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   55
      Left            =   3000
      TabIndex        =   66
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "54"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   54
      Left            =   2400
      TabIndex        =   65
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "53"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   53
      Left            =   1800
      TabIndex        =   64
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "52"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   52
      Left            =   1200
      TabIndex        =   63
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "51"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   51
      Left            =   600
      TabIndex        =   62
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "10"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   10
      Left            =   6000
      TabIndex        =   21
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "9"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   9
      Left            =   5400
      TabIndex        =   20
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "8"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   8
      Left            =   4800
      TabIndex        =   19
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "7"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   7
      Left            =   4200
      TabIndex        =   18
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "6"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   6
      Left            =   3600
      TabIndex        =   17
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "5"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   5
      Left            =   3000
      TabIndex        =   16
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "4"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   4
      Left            =   2400
      TabIndex        =   15
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "3"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   3
      Left            =   1800
      TabIndex        =   14
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "2"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   2
      Left            =   1200
      TabIndex        =   13
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "1"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   12
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "11"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   11
      Left            =   600
      TabIndex        =   22
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "50"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   50
      Left            =   6000
      TabIndex        =   61
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "49"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   49
      Left            =   5400
      TabIndex        =   60
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "48"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   48
      Left            =   4800
      TabIndex        =   59
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "47"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   47
      Left            =   4200
      TabIndex        =   58
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "46"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   46
      Left            =   3600
      TabIndex        =   57
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "45"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   45
      Left            =   3000
      TabIndex        =   56
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "44"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   44
      Left            =   2400
      TabIndex        =   55
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "43"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   43
      Left            =   1800
      TabIndex        =   54
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "42"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   42
      Left            =   1200
      TabIndex        =   53
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "41"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   41
      Left            =   600
      TabIndex        =   52
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "40"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   40
      Left            =   6000
      TabIndex        =   51
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "39"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   39
      Left            =   5400
      TabIndex        =   50
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "38"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   38
      Left            =   4800
      TabIndex        =   49
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "37"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   37
      Left            =   4200
      TabIndex        =   48
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "36"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   36
      Left            =   3600
      TabIndex        =   47
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "35"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   35
      Left            =   3000
      TabIndex        =   46
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "34"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   34
      Left            =   2400
      TabIndex        =   45
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "33"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   33
      Left            =   1800
      TabIndex        =   44
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "32"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   32
      Left            =   1200
      TabIndex        =   43
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "31"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   31
      Left            =   600
      TabIndex        =   42
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "30"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   30
      Left            =   6000
      TabIndex        =   41
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "29"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   29
      Left            =   5400
      TabIndex        =   40
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "28"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   28
      Left            =   4800
      TabIndex        =   39
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "27"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   27
      Left            =   4200
      TabIndex        =   38
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "26"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   26
      Left            =   3600
      TabIndex        =   37
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "25"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   25
      Left            =   3000
      TabIndex        =   36
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "24"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   24
      Left            =   2400
      TabIndex        =   35
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "23"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   23
      Left            =   1800
      TabIndex        =   34
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "22"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   22
      Left            =   1200
      TabIndex        =   33
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "21"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   21
      Left            =   600
      TabIndex        =   32
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "20"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   20
      Left            =   6000
      TabIndex        =   31
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "19"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   19
      Left            =   5400
      TabIndex        =   30
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "18"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   18
      Left            =   4800
      TabIndex        =   29
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "17"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   17
      Left            =   4200
      TabIndex        =   28
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "16"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   16
      Left            =   3600
      TabIndex        =   27
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "15"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   15
      Left            =   3000
      TabIndex        =   26
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "14"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   14
      Left            =   2400
      TabIndex        =   25
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "13"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   13
      Left            =   1800
      TabIndex        =   24
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "12"
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   12
      Left            =   1200
      TabIndex        =   23
      Top             =   4080
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   9735
      Left            =   0
      Picture         =   "amiral battý ödev form.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim durum1 As Boolean
Dim durum2 As Boolean
Dim durum3 As Boolean
Dim durum4 As Boolean
Dim durum5 As Boolean
Dim durum6 As Boolean
Dim durum7 As Boolean
Dim durum8 As Boolean
Dim durum9 As Boolean
Dim durum10 As Boolean
Dim durum11 As Boolean
Dim durum12 As Boolean
Dim durum13 As Boolean
Dim durum14 As Boolean
Dim durum15 As Boolean
Dim durum16 As Boolean
Dim durum17 As Boolean
Dim durum18 As Boolean
Dim durum19 As Boolean
Dim durum20 As Boolean
Dim sayac As Integer

 Dim Jhin(101) As Integer
Dim tur(101) As Integer

Dim f As Integer

Dim bt As Integer
Dim but As Integer
Dim but1 As Integer
Dim byb As Integer
Dim byb2 As Integer
Dim byb1 As Integer
Dim bamr As Integer
Dim bamr1 As Integer
Dim bamr2 As Integer
Dim bamr3 As Integer
Dim O As Integer


Dim bkonumgir(100) As Integer

Dim b(10, 10) As Integer
Dim s As Integer

Dim t As Integer
Dim ut As Integer
Dim ut1 As Integer
Dim yb1 As Integer
Dim yb2 As Integer

Dim yb As Integer
Dim amr As Integer
Dim amr1 As Integer
Dim amr2 As Integer
Dim amr3 As Integer



Dim a(10, 10) As Integer


Private Sub Command1_Click()
Randomize Timer
bt = Int(Rnd(1) * 100)


'hocam burada botun gemilerine konum atýyoruz aþaðýda ise amiral gibi 4 tane kutusu olanlarýn
'gelemiyceði yerleri tek tek girdim daha kýsa nasýl yapýcaðýmý bulamadým


tekrarata3:

but = Int(Rnd(1) * 100)
If but = 20 Then GoTo tekrarata3
If but = 20 Then GoTo tekrarata3
If but = 30 Then GoTo tekrarata3
If but = 40 Then GoTo tekrarata3
If but = 50 Then GoTo tekrarata3
If but = 60 Then GoTo tekrarata3
If but = 70 Then GoTo tekrarata3
If but = 80 Then GoTo tekrarata3
If but = 90 Then GoTo tekrarata3
If but = 100 Then GoTo tekrarata3
but1 = but + 1
If but1 = bt Then GoTo tekrarata3
If but = bt Then GoTo tekrarata3
tekrarata2:

byb = Int(Rnd(1) * 100)

If byb = 9 Then GoTo tekrarata2
If byb = 19 Then GoTo tekrarata2
If byb = 29 Then GoTo tekrarata2
If byb = 39 Then GoTo tekrarata2
If byb = 49 Then GoTo tekrarata2
If byb = 59 Then GoTo tekrarata2
If byb = 69 Then GoTo tekrarata2
If byb = 79 Then GoTo tekrarata2
If byb = 89 Then GoTo tekrarata2
If byb = 99 Then GoTo tekrarata2
If byb = 20 Then GoTo tekrarata2
If byb = 30 Then GoTo tekrarata2
If byb = 40 Then GoTo tekrarata2
If byb = 50 Then GoTo tekrarata2
If byb = 60 Then GoTo tekrarata2
If byb = 70 Then GoTo tekrarata2
If byb = 80 Then GoTo tekrarata2
If byb = 90 Then GoTo tekrarata2
If byb = 100 Then GoTo tekrarata2
byb1 = byb + 1
byb2 = byb1 + 1
If byb1 = bt Then GoTo tekrarata2
If byb2 = but Then GoTo tekrarata2
If byb1 = but Then GoTo tekrarata2
If byb2 = bt Then GoTo tekrarata2
If byb = bt Then GoTo tekrarata2
If byb = but Then GoTo tekrarata2
If byb = but1 Then GoTo tekrarata2




tekrarata:

bamr = Int(Rnd(1) * 100)
If bamr <= 10 Then GoTo tekrarata
If bamr = 9 Then GoTo tekrarata
If bamr = 19 Then GoTo tekrarata
If bamr = 29 Then GoTo tekrarata
If bamr = 39 Then GoTo tekrarata
If bamr = 49 Then GoTo tekrarata
If bamr = 59 Then GoTo tekrarata
If bamr = 69 Then GoTo tekrarata
If bamr = 79 Then GoTo tekrarata
If bamr = 89 Then GoTo tekrarata
If bamr = 99 Then GoTo tekrarata
If bamr = 30 Then GoTo tekrarata
If bamr = 40 Then GoTo tekrarata
If bamr = 50 Then GoTo tekrarata
If bamr = 60 Then GoTo tekrarata
If bamr = 70 Then GoTo tekrarata
If bamr = 80 Then GoTo tekrarata
If bamr = 90 Then GoTo tekrarata
If bamr = 100 Then GoTo tekrarata

bamr1 = bamr + 1
bamr2 = bamr1 + 1
bamr3 = bamr - 9
If bamr = bt Then GoTo tekrarata
If bamr1 = bt Then GoTo tekrarata
If bamr2 = bt Then GoTo tekrarata
If bamr3 = bt Then GoTo tekrarata
If bamr1 = but Then GoTo tekrarata
If bamr1 = but1 Then GoTo tekrarata
If bamr2 = byb Then GoTo tekrarata
If bamr1 = byb2 Then GoTo tekrarata
If bamr1 = byb1 Then GoTo tekrarata
If bamr1 = byb Then GoTo tekrarata
If bamr3 = byb2 Then GoTo tekrarata
If bamr3 = byb Then GoTo tekrarata
If bamr3 = byb1 Then GoTo tekrarata
If bamr = but1 Then GoTo tekrarata
If bamr = byb2 Then GoTo tekrarata
If bamr3 = but1 Then GoTo tekrarata
If bamr3 = but Then GoTo tekrarata


f = 0

For i = 0 To 10 Step 1
    For j = 0 To 10 Step 1
    f = f + 1
    b(i, j) = f
    
    If bt = b(i, j) Then btegmen = bt
    If but = b(i, j) Then busttegmen = but
    If but1 = b(i, j) Then busttegmen1 = but1
    If byb = b(i, j) Then byuzbasi = byb
    If byb1 = b(i, j) Then byuzbasi1 = byb1
    If byb2 = b(i, j) Then byüzbasi2 = byb2
    If bamr = b(i, j) Then bamiral = bamr
    If bamr1 = b(i, j) Then bamiral1 = bamr1
    If bamr2 = b(i, j) Then bamiral2 = bamr2
    If bamr3 = b(i, j) Then bamiral3 = bamr3
    Dim kontrol As String
    Dim sorgu As Boolean
    
    Next
Next
'burda botun bana vurucaðý konumlarý önceden bir diziye girdiriyorum sonra tura göre gelen
'diziyi konum olarak seçtiriyorum
For n = 1 To 100 Step 1
 
e7:
Randomize Timer
bkonumgir(n) = Int(Rnd(1) * 100)
If bkonumgir(n) = 1 Then saya1ç = saya1ç + 1
If bkonumgir(n) = 1 And saya1ç = 1 Then GoTo e8
If bkonumgir(n) = 2 Then saya2ç = saya2ç + 1

If bkonumgir(n) = 2 And saya2ç = 1 Then GoTo e8
If bkonumgir(n) = 3 Then saya3ç = saya3ç + 1

If bkonumgir(n) = 3 And saya3ç = 1 Then GoTo e8
If bkonumgir(n) = 4 Then saya4ç = saya4ç + 1

If bkonumgir(n) = 4 And saya4ç = 1 Then GoTo e8
If bkonumgir(n) = 5 Then saya5ç = saya5ç + 1

If bkonumgir(n) = 5 And saya5ç = 1 Then GoTo e8
If bkonumgir(n) = 6 Then saya6ç = saya6ç + 1

If bkonumgir(n) = 6 And saya6ç = 1 Then GoTo e8
If bkonumgir(n) = 7 Then saya7ç = saya7ç + 1
If bkonumgir(n) = 7 And saya7ç = 1 Then GoTo e8
If bkonumgir(n) = 8 Then saya8ç = saya8ç + 1
If bkonumgir(n) = 8 And saya8ç = 1 Then GoTo e8
If bkonumgir(n) = 9 Then saya9ç = saya9ç + 1
If bkonumgir(n) = 9 And saya9ç = 1 Then GoTo e8



sorgu = InStr(1, kontrol, bkonumgir(n))
jet = jet + 1


If sorgu = True And jet <= 1001 Then GoTo e7

kontrol = kontrol & "," & bkonumgir(n)

jet = 0
e8:
List9.AddItem bkonumgir(n)
Next







O = 1
'burda ise botun gemilerini atadýðý konumlarý listeye aktarýyorum oyun sonunda visible olucak þekilde aþaðýyada kodladým
MsgBox "düþman konumlarýný seçti"
List8.AddItem "düþman teðmen"
List7.AddItem "düþman üstteðmen"
List6.AddItem "düþman yüzbaþý"
List5.AddItem "düþman amiral"

List8.AddItem bt
List7.AddItem but
List7.AddItem but1
List6.AddItem byb
List6.AddItem byb1
List6.AddItem byb2
List5.AddItem bamr
List5.AddItem bamr1
List5.AddItem bamr2
List5.AddItem bamr3
List9.AddItem "botun vurduðu konumlar"
List10.AddItem "Tur"
List11.AddItem "senin vurduðun konumlar"


End Sub



Private Sub Command2_Click()
'burda ise kendi gemilerimin konumlarýný atýyorum
t = InputBox("girin teðmen için")
ut = InputBox("üst teðmen için seçtiðin kutunun yanýndaki kutu 2.kutu olarak sayýlýr hesaplama için üstteðmen1 = usteðmen + 1")
ut1 = ut + 1
yb = InputBox("yüzbaþý için seçtiðin kutunun yanindaki 2.ve 3. kutu geminin devamý sayýlýr hesaplama için yüzbaþý1 = yüzbaþý + 1   yüzbaþý2 = yüzbaþý1 + 1 ")
yb1 = yb + 1
yb2 = yb1 + 1
amr = InputBox("amiral için seçtiðin kutunun yanindaki 2,3 ve 2'nin üstündeki kutu geminin devamý sayýlýr hesaplama için amr1 = amr + 1 amr2 = amr1 + 1 amr3 = amr - 9")
amr1 = amr + 1
amr2 = amr1 + 1
amr3 = amr - 9

s = 0

For i = 0 To 10 Step 1
    For j = 0 To 10 Step 1
    s = s + 1
    a(i, j) = s
    
    If t = a(i, j) Then tegmen = t
    If ut = a(i, j) Then usttegmen = ut
    If ut = a(i, j) Then usttegmen1 = ut1
    If yb = a(i, j) Then yuzbasi = yb
    If yb1 = a(i, j) Then yuzbasi1 = yb1
    If yb2 = a(i, j) Then yüzbasi2 = yb2
    If amr = a(i, j) Then amiral = amr
    If amr1 = a(i, j) Then amiral1 = amr1
    If amr2 = a(i, j) Then amiral2 = amr2
    If amr3 = a(i, j) Then amiral3 = amr3
    
    Next
Next
MsgBox "konumlarýný seçtin"

' burda kendi gemi konumlarýmý listeye yazdýrýyorum oyun içerisinde göre bilmek için
List4.AddItem "teðmen"
List3.AddItem "üstteðmen"
List2.AddItem "yüzbaþý"
List1.AddItem "amiral"

List4.AddItem t
List3.AddItem ut
List3.AddItem ut1
List2.AddItem yb
List2.AddItem yb1
List2.AddItem yb2
List1.AddItem amr
List1.AddItem amr1
List1.AddItem amr2
List1.AddItem amr3


End Sub

Private Sub Command3_Click()

'burda düþman gemilerini vurucaðým konumlarý seçiyorum ve düþmaný vurup vuramadýðýmý vurduðumda olucaklarý
'tek tek yazdým
konumgir = InputBox("konumgir")
List11.AddItem konumgir

For i = 0 To 10 Step 1
    For j = 0 To 10 Step 1
If b(i, j) = bamr And konumgir = bamr Then 'burda vurduðum konumla düþman konumu uyuþuyormu onu kontrol ediyor
Text1.Text = "vurdun" & Chr(13) & b(i, j) ' burda metin kutusuna vurduðum konum yazýyor
Text1.BackColor = vbRed
Index = b(i, j)
Label1(Index).BackColor = vbRed 'burda vurduðum konumun rengi deðiþiyor
durum1 = True 'burda geminin bu bölümünün vurulup vurulmadýðýný kontrol ediyor


GoTo s
End If
If b(i, j) = bamr1 And konumgir = bamr1 Then
Text1.Text = "vurdun" & Chr(13) & b(i, j)
Text1.BackColor = vbRed
Index = b(i, j)
Label1(Index).BackColor = vbRed
durum2 = True

GoTo s
End If

If b(i, j) = bamr2 And konumgir = bamr2 Then
Text1.Text = "vurdun" & Chr(13) & b(i, j)
Text1.BackColor = vbRed
Index = b(i, j)
Label1(Index).BackColor = vbRed

durum3 = True

GoTo s

End If
If b(i, j) = bamr3 And konumgir = bamr3 Then
Text1.Text = "vurdun" & Chr(13) & b(i, j)
Text1.BackColor = vbRed
Index = b(i, j)
Label1(Index).BackColor = vbRed
durum4 = True
GoTo s
End If
If b(i, j) = but And konumgir = but Then
Text1.Text = "vurdun" & Chr(13) & b(i, j)
Text1.BackColor = vbRed
Index = b(i, j)
Label1(Index).BackColor = vbRed
durum5 = True
GoTo s

End If
If b(i, j) = but1 And konumgir = but1 Then
Text1.Text = "vurdun" & Chr(13) & b(i, j)
Text1.BackColor = vbRed
Index = b(i, j)
Label1(Index).BackColor = vbRed
durum6 = True
GoTo s
End If
If b(i, j) = byb And konumgir = byb Then
Text1.Text = "vurdun" & Chr(13) & b(i, j)
Text1.BackColor = vbRed
Index = b(i, j)
Label1(Index).BackColor = vbRed
durum7 = True
GoTo s
End If
If b(i, j) = byb1 And konumgir = byb1 Then
Text1.Text = "vurdun" & Chr(13) & b(i, j)
Text1.BackColor = vbRed
Index = b(i, j)
Label1(Index).BackColor = vbRed
durum8 = True
GoTo s
End If
If b(i, j) = byb2 And konumgir = byb2 Then
Text1.Text = "vurdun" & Chr(13) & b(i, j)
Text1.BackColor = vbRed
Index = b(i, j)
Label1(Index).BackColor = vbRed
durum9 = True
GoTo s
End If
If b(i, j) = bt And konumgir = bt Then
Text1.Text = "vurdun" & Chr(13) & b(i, j)
Text1.BackColor = vbRed
Index = b(i, j)
Label1(Index).BackColor = vbRed
durum10 = True
GoTo s
End If
s:

  Next
Next
 ' burdan sonrasý bot'un  kýsmý



For i = 0 To 10 Step 1
    For j = 0 To 10 Step 1
If a(i, j) = amr And bkonumgir(O) = amr Then
Text1.Text = "vurdun" & Chr(13) & a(i, j)
Text1.BackColor = vbGreen
Index = a(i, j)
Label1(Index).BackColor = vbGreen
Label2.BackColor = vbGreen
durum11 = True
GoTo l
End If
If a(i, j) = amr1 And bkonumgir(O) = amr1 Then
Text1.Text = "vurdun" & Chr(13) & a(i, j)
Text1.BackColor = vbGreen
Index = a(i, j)
Label1(Index).BackColor = vbGreen
Label3.BackColor = vbGreen
durum12 = True
GoTo l
End If

If a(i, j) = amr2 And bkonumgir(O) = amr2 Then
Text1.Text = "vurdun" & Chr(13) & a(i, j)
Text1.BackColor = vbGreen
Index = a(i, j)
Label1(Index).BackColor = vbGreen
Label4.BackColor = vbGreen
durum13 = True
GoTo l

End If
If a(i, j) = amr3 And bkonumgir(O) = amr3 Then
Text1.Text = "vurdun" & Chr(13) & a(i, j)
Text1.BackColor = vbGreen
Index = a(i, j)
Label1(Index).BackColor = vbGreen
Label5.BackColor = vbGreen
durum14 = True
GoTo l
End If
If a(i, j) = ut And bkonumgir(O) = ut Then
Text1.Text = "vurdun" & Chr(13) & a(i, j)
Text1.BackColor = vbGreen
Index = a(i, j)
Label1(Index).BackColor = vbGreen
Label9.BackColor = vbGreen
durum15 = True
GoTo l

End If
If a(i, j) = ut1 And bkonumgir(O) = ut1 Then
Text1.Text = "vurdun" & Chr(13) & a(i, j)
Text1.BackColor = vbGreen
Index = a(i, j)
Label1(Index).BackColor = vbGreen
Label10.BackColor = vbGreen
durum16 = True
GoTo l
End If
If a(i, j) = yb And bkonumgir(O) = yb Then
Text1.Text = "vurdun" & Chr(13) & a(i, j)
Text1.BackColor = vbGreen
Index = a(i, j)
Label1(Index).BackColor = vbGreen
Label6.BackColor = vbGreen
durum17 = True
GoTo l
End If
If a(i, j) = yb1 And bkonumgir(O) = yb1 Then
Text1.Text = "vurdun" & Chr(13) & a(i, j)
Text1.BackColor = vbGreen
Index = a(i, j)
Label1(Index).BackColor = vbGreen
Label7.BackColor = vbGreen
durum18 = True
GoTo l
End If
If a(i, j) = yb2 And bkonumgir(O) = yb2 Then
Text1.Text = "vurdun" & Chr(13) & a(i, j)
Text1.BackColor = vbGreen
Index = a(i, j)
Label1(Index).BackColor = vbGreen
Label8.BackColor = vbGreen
durum19 = True
GoTo l
End If
If a(i, j) = t And bkonumgir(O) = t Then
Text1.Text = "vurdun" & a(i, j)
Text1.BackColor = vbGreen
Index = a(i, j)
Label1(Index).BackColor = vbGreen
Label11.BackColor = vbGreen
durum20 = True
GoTo l
End If
l:

  Next
Next

'burasý tüm konumlarýn vurulup vurulmadýðýný ve düþmanýn konum ve vurduðu yerleri oyun sonunda
'göstermesi için ayaladýðým yer
If durum4 = True And durum5 = True And durum3 = True And durum2 = True And durum1 = True And durum6 = True And durum7 = True And durum9 = True And durum8 = True And durum10 = True Then MsgBox "kazandýn"
If durum4 = True And durum5 = True And durum3 = True And durum2 = True And durum1 = True And durum6 = True And durum7 = True And durum9 = True And durum8 = True And durum10 = True Then List5.Visible = True

If durum4 = True And durum5 = True And durum3 = True And durum2 = True And durum1 = True And durum6 = True And durum7 = True And durum9 = True And durum8 = True And durum10 = True Then List6.Visible = True

If durum4 = True And durum5 = True And durum3 = True And durum2 = True And durum1 = True And durum6 = True And durum7 = True And durum9 = True And durum8 = True And durum10 = True Then List7.Visible = True
If durum4 = True And durum5 = True And durum3 = True And durum2 = True And durum1 = True And durum6 = True And durum7 = True And durum9 = True And durum8 = True And durum10 = True Then List8.Visible = True
If durum4 = True And durum5 = True And durum3 = True And durum2 = True And durum1 = True And durum6 = True And durum7 = True And durum9 = True And durum8 = True And durum10 = True Then List9.Visible = True


If durum14 = True And durum15 = True And durum13 = True And durum12 = True And durum11 = True And durum16 = True And durum17 = True And durum19 = True And durum18 = True And durum20 = True Then MsgBox "kaybettin"
If durum14 = True And durum15 = True And durum13 = True And durum12 = True And durum11 = True And durum16 = True And durum17 = True And durum19 = True And durum18 = True And durum20 = True Then List5.Visible = True
If durum14 = True And durum15 = True And durum13 = True And durum12 = True And durum11 = True And durum16 = True And durum17 = True And durum19 = True And durum18 = True And durum20 = True Then List6.Visible = True
If durum14 = True And durum15 = True And durum13 = True And durum12 = True And durum11 = True And durum16 = True And durum17 = True And durum19 = True And durum18 = True And durum20 = True Then List7.Visible = True
If durum14 = True And durum15 = True And durum13 = True And durum12 = True And durum11 = True And durum16 = True And durum17 = True And durum19 = True And durum18 = True And durum20 = True Then List8.Visible = True
If durum14 = True And durum15 = True And durum13 = True And durum12 = True And durum11 = True And durum16 = True And durum17 = True And durum19 = True And durum18 = True And durum20 = True Then List9.Visible = True

'burasý tur sayýsýný gösteren yer
List10.AddItem O

Command3.Caption = "sonraki tura geç"

O = O + 1

End Sub








Private Sub Command4_Click()
'oyunu sýfýrlayan yer
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
List8.Clear
List9.Clear
List10.Clear
O = 0
List11.Clear
List5.Visible = False
List6.Visible = False
List7.Visible = False
List8.Visible = False
List9.Visible = False
For t = 0 To 100 Step 1


bkonumgir(t) = 0
Next

For h = 1 To 100 Step 1
Index = h
Label1(Index).BackColor = vbBlue
Next
Label2.BackColor = vbBlue
Label3.BackColor = vbBlue
Label4.BackColor = vbBlue
Label5.BackColor = vbBlue
Label6.BackColor = vbBlue
Label7.BackColor = vbBlue
Label8.BackColor = vbBlue
Label9.BackColor = vbBlue
Label10.BackColor = vbBlue
Label11.BackColor = vbBlue

durum1 = False
durum2 = False
durum6 = False
durum3 = False
durum5 = False
durum4 = False
durum7 = False
durum8 = False
durum9 = False
durum10 = False
durum11 = False
durum12 = False
durum14 = False
durum13 = False
durum15 = False
durum16 = False
durum17 = False
durum18 = False
durum19 = False
durum20 = False


End Sub


