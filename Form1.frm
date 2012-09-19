VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modbus TCP/IP Client"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdStopPoll 
      Caption         =   "ENDP"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   112
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmd_poll 
      Caption         =   "Poll"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   111
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   1215
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   110
      Top             =   6240
      Width           =   10095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Automation"
      Height          =   615
      Left            =   10320
      TabIndex        =   109
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   9360
      TabIndex        =   106
      Text            =   "1"
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox port 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   104
      Text            =   "502"
      Top             =   4800
      Width           =   975
   End
   Begin VB.CheckBox Check 
      Caption         =   "FLOAD"
      Height          =   255
      Left            =   11040
      TabIndex        =   103
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   89
      Left            =   11040
      TabIndex        =   102
      Text            =   "0"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   88
      Left            =   11040
      TabIndex        =   101
      Text            =   "0"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   87
      Left            =   11040
      TabIndex        =   100
      Text            =   "0"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   86
      Left            =   11040
      TabIndex        =   99
      Text            =   "0"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   85
      Left            =   11040
      TabIndex        =   98
      Text            =   "0"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   84
      Left            =   11040
      TabIndex        =   97
      Text            =   "0"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   83
      Left            =   11040
      TabIndex        =   96
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   82
      Left            =   11040
      TabIndex        =   95
      Text            =   "0"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   81
      Left            =   11040
      TabIndex        =   94
      Text            =   "0"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   80
      Left            =   11040
      TabIndex        =   93
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   79
      Left            =   9720
      TabIndex        =   92
      Text            =   "0"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   78
      Left            =   9720
      TabIndex        =   91
      Text            =   "0"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   77
      Left            =   9720
      TabIndex        =   90
      Text            =   "0"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   76
      Left            =   9720
      TabIndex        =   89
      Text            =   "0"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   75
      Left            =   9720
      TabIndex        =   88
      Text            =   "0"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   74
      Left            =   9720
      TabIndex        =   87
      Text            =   "0"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   73
      Left            =   9720
      TabIndex        =   86
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   72
      Left            =   9720
      TabIndex        =   85
      Text            =   "0"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   71
      Left            =   9720
      TabIndex        =   84
      Text            =   "0"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   70
      Left            =   9720
      TabIndex        =   83
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   69
      Left            =   8400
      TabIndex        =   82
      Text            =   "0"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   68
      Left            =   8400
      TabIndex        =   81
      Text            =   "0"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   67
      Left            =   8400
      TabIndex        =   80
      Text            =   "0"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   66
      Left            =   8400
      TabIndex        =   79
      Text            =   "0"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   65
      Left            =   8400
      TabIndex        =   78
      Text            =   "0"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   64
      Left            =   8400
      TabIndex        =   77
      Text            =   "0"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   63
      Left            =   8400
      TabIndex        =   76
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   62
      Left            =   8400
      TabIndex        =   75
      Text            =   "0"
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "save"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   74
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   0
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      TabIndex        =   72
      Text            =   "Disconnected"
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox ip 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   70
      Text            =   "192.168.1.103"
      Top             =   4800
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "192.168.1.103"
      RemotePort      =   502
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Write Registers"
      Height          =   495
      Left            =   5880
      TabIndex        =   69
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   61
      Left            =   8400
      TabIndex        =   68
      Text            =   "0"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   60
      Left            =   8400
      TabIndex        =   67
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   59
      Left            =   7080
      TabIndex        =   66
      Text            =   "0"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   58
      Left            =   7080
      TabIndex        =   65
      Text            =   "0"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   57
      Left            =   7080
      TabIndex        =   64
      Text            =   "0"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   56
      Left            =   7080
      TabIndex        =   63
      Text            =   "0"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   55
      Left            =   7080
      TabIndex        =   62
      Text            =   "0"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   54
      Left            =   7080
      TabIndex        =   61
      Text            =   "0"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   53
      Left            =   7080
      TabIndex        =   60
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   52
      Left            =   7080
      TabIndex        =   59
      Text            =   "0"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   51
      Left            =   7080
      TabIndex        =   58
      Text            =   "0"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   50
      Left            =   7080
      TabIndex        =   57
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   49
      Left            =   5760
      TabIndex        =   56
      Text            =   "0"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   48
      Left            =   5760
      TabIndex        =   55
      Text            =   "0"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   47
      Left            =   5760
      TabIndex        =   54
      Text            =   "0"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   46
      Left            =   5760
      TabIndex        =   53
      Text            =   "0"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   45
      Left            =   5760
      TabIndex        =   52
      Text            =   "0"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   44
      Left            =   5760
      TabIndex        =   51
      Text            =   "0"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   43
      Left            =   5760
      TabIndex        =   50
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   42
      Left            =   5760
      TabIndex        =   49
      Text            =   "0"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   41
      Left            =   5760
      TabIndex        =   48
      Text            =   "0"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   40
      Left            =   5760
      TabIndex        =   47
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   39
      Left            =   4440
      TabIndex        =   46
      Text            =   "0"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   38
      Left            =   4440
      TabIndex        =   45
      Text            =   "0"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   37
      Left            =   4440
      TabIndex        =   44
      Text            =   "0"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   36
      Left            =   4440
      TabIndex        =   43
      Text            =   "0"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   35
      Left            =   4440
      TabIndex        =   42
      Text            =   "0"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   34
      Left            =   4440
      TabIndex        =   41
      Text            =   "0"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   33
      Left            =   4440
      TabIndex        =   40
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   32
      Left            =   4440
      TabIndex        =   39
      Text            =   "0"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   31
      Left            =   4440
      TabIndex        =   38
      Text            =   "0"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   30
      Left            =   4440
      TabIndex        =   37
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   29
      Left            =   3120
      TabIndex        =   36
      Text            =   "0"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   28
      Left            =   3120
      TabIndex        =   35
      Text            =   "0"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   27
      Left            =   3120
      TabIndex        =   34
      Text            =   "0"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   26
      Left            =   3120
      TabIndex        =   33
      Text            =   "0"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   25
      Left            =   3120
      TabIndex        =   32
      Text            =   "0"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   24
      Left            =   3120
      TabIndex        =   31
      Text            =   "0"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   23
      Left            =   3120
      TabIndex        =   30
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   22
      Left            =   3120
      TabIndex        =   29
      Text            =   "0"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   21
      Left            =   3120
      TabIndex        =   28
      Text            =   "0"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   20
      Left            =   3120
      TabIndex        =   27
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   19
      Left            =   1800
      TabIndex        =   26
      Text            =   "0"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   18
      Left            =   1800
      TabIndex        =   25
      Tag             =   "T19"
      Text            =   "0"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   1800
      TabIndex        =   24
      Text            =   "0"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   1800
      TabIndex        =   23
      Text            =   "0"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   1800
      TabIndex        =   22
      Text            =   "0"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   1800
      TabIndex        =   21
      Text            =   "0"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   1800
      TabIndex        =   20
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   1800
      TabIndex        =   19
      Text            =   "0"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   1800
      TabIndex        =   18
      Text            =   "0"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   1800
      TabIndex        =   17
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   480
      TabIndex        =   16
      Text            =   "0"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   480
      TabIndex        =   15
      Text            =   "0"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   480
      TabIndex        =   14
      Text            =   "0"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   480
      TabIndex        =   13
      Text            =   "0"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   480
      TabIndex        =   12
      Text            =   "0"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   11
      Text            =   "0"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   10
      Text            =   "0"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   480
      TabIndex        =   9
      Text            =   "0"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Text            =   "0"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Text            =   "0"
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmd_read 
      Caption         =   "Read "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   8040
      TabIndex        =   4
      Text            =   "10"
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   6480
      TabIndex        =   2
      Text            =   "100"
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmd_disconnect 
      Caption         =   "Disconnect"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmd_connect 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Data type"
      Height          =   375
      Left            =   11040
      TabIndex        =   108
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Device address"
      Height          =   375
      Left            =   9360
      TabIndex        =   107
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "port"
      Height          =   375
      Left            =   3000
      TabIndex        =   105
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Status"
      Height          =   255
      Left            =   4320
      TabIndex        =   73
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Adrress IP"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1320
      TabIndex        =   71
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Length"
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Start register"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim MbusQuery(11) As Byte
Public MbusResponse As String
Dim MbusByteArray(500) 'As Byte
Public MbusRead As Boolean
Public MbusWrite As Boolean
Dim ModbusTimeOut As Integer
Dim ModbusWait As Boolean
Dim mbpolling  As Boolean

Private Sub Check1_Click()
Timer1.Interval = 1000
If Timer1.Enabled Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
End If
End Sub







Private Sub cmd_poll_Click()
    lStartTime = GetTickCount
    mbpolling = True
'    lTime_Max = 0
'    lTime_Min = 1000
'    txtTimeMax.Text = 0
'    txtTimeMin.Text = 1000
'    cmdStartPolling.Enabled = False
'    cmdStopPolling.Enabled = True
'    txtStartTime.Text = Now
    cmd_read_Click
End Sub

'Sub check1_DataChanged()
'Timer1.Enabled = True
'Timer1.Interval = 1000
'End Sub


Private Sub cmd_connect_Click()
Dim StartTime

If (Winsock1.State <> sckClosed) Then
    Winsock1.Close
End If
Winsock1.RemoteHost = ip.Text
Winsock1.RemotePort = port.Text
Winsock1.Connect

StartTime = Timer

Do While ((Timer < StartTime + 2) And (Winsock1.State <> 7))
DoEvents
Loop
If (Winsock1.State = 7) Then
   Text5.Text = "Connected"
   Text5.BackColor = &HFF00&
Else
   Text5.Text = "Can't connect"
   Text5.BackColor = &HFF
End If
End Sub




Private Sub cmd_disconnect_Click()
If (Winsock1.State <> sckClosed) Then
Winsock1.Close
End If
Do While (Winsock1.State <> sckClosed)
DoEvents
Loop
Text5.Text = "Disconnected"
Text5.BackColor = &HFF
End Sub
Public Function read_dat()
Dim StartLow As Byte
Dim StartHigh As Byte
Dim LengthLow As Byte
Dim LengthHigh As Byte
If (Winsock1.State = 7) Then

StartLow = Val(Text2.Text - 1) Mod 256
StartHigh = Val(Text2.Text - 1) \ 256
LengthLow = Val(Text3.Text) Mod 256
LengthHigh = Val(Text3.Text) \ 256
MbusQuery(0) = 0
MbusQuery(1) = 0
MbusQuery(2) = 0
MbusQuery(3) = 0
MbusQuery(4) = 0
MbusQuery(5) = 6
MbusQuery(6) = Val(Text1.Text)
MbusQuery(7) = 3
MbusQuery(8) = StartHigh
MbusQuery(9) = StartLow
MbusQuery(10) = LengthHigh
MbusQuery(11) = LengthLow
MbusRead = True
MbusWrite = False
Winsock1.SendData MbusQuery
ModbusWait = True
ModbusTimeOut = 0
'Timer1.Enabled = True
'Else
'MsgBox ("Device not connected via TCP/IP")
End If
End Function

Private Sub cmd_read_Click()
Call read_dat
End Sub

Private Sub VScroll1_Change()

End Sub


Private Sub cmdStopPoll_Click()
mbpolling = False
End Sub

Private Sub Command4_Click()
Dim MbusWriteCommand As String
Dim StartLow As Byte
Dim StartHigh As Byte
Dim ByteLow As Byte
Dim ByteHigh As Byte
Dim i As Integer
If (Winsock1.State = 7) Then
StartLow = Val(Text2.Text - 1) Mod 256
StartHigh = Val(Text2.Text - 1) \ 256
LengthLow = Val(Text3.Text) Mod 256
LengthHigh = Val(Text3.Text) \ 256


MbusWriteQuery = Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(7 + 2 * Val(Text3.Text)) + Chr(1) + Chr(16) + Chr(StartHigh) + Chr(StartLow) + Chr(0) + Chr(Val(Text3.Text)) + Chr(2 * Val(Text3.Text))
For i = 0 To Val(Text3.Text) - 1
ByteLow = Val(Text4(i).Text) Mod 256
ByteHigh = Val(Text4(i).Text) \ 256
MbusWriteQuery = MbusWriteQuery + Chr(ByteHigh) + Chr(ByteLow)
Next i
MbusRead = False
MbusWrite = True
Winsock1.SendData MbusWriteQuery
ModbusWait = True
ModbusTimeOut = 0
Timer1.Enabled = True
Else
MsgBox ("Device not connected via TCP/IP")
End If
End Sub





Private Sub cmd_save_Click()

Dim i
Open "d:\LTP.txt" For Output As #1
For i = 0 To 89
'Print #1, NO; Str(i + 1); Text4(i).Text
Write #1, Val(Text4(i).Text)

Next i
'Print #1, "浓度"
Close #1
Call save
End Sub


Public Function save()
On Error GoTo pj1
pj2: Open "c:\tcp_set.ini" For Output As #1
Print #1, ip.Text
Print #1, port.Text
Print #1, Text2.Text
Print #1, Text3.Text
Print #1, Text1.Text
Print #1, Check.Value

Close #1
Exit Function
pj1: Close #1
    GoTo pj2
End Function

Private Sub Form_Load()
Dim tem As String
On Error GoTo xt1
xt2: Open "c:\tcp_set.ini" For Input As #1
Line Input #1, tem '
ip.Text = tem
Line Input #1, tem
port.Text = tem
Line Input #1, tem
Text2.Text = tem
Line Input #1, tem
Text3.Text = tem
Line Input #1, tem
Text1.Text = tem
Line Input #1, tem
Check.Value = Val(tem)
'On Error GoTo xt1
Close #1
'Call Command1_Click
'Call Command3_Click

'Call Command5_Click
'Unload Me
Exit Sub

xt1: Close #1
     Open "c:\tcp_set.ini" For Output As #1
     Print #1, "192.168.1.103"
     Print #1, "502"
     Print #1, "100"
     Print #1, "10"
     Print #1, "1"
     Print #1, "1"
     Close #1
     GoTo xt2


End Sub



'
Private Sub Timer1_Timer()
Call read_dat
End Sub

'Private Sub Timer1_Timer()
'
'ModbusTimeOut = ModbusTimeOut + 1
'If ModbusTimeOut > 2 Then
'ModbusWait = False
'ModbusTimeOut = 0
'Text5.Text = "Modbus Time Out"
'Text5.BackColor = &HFF
'Timer1.Enabled = False
'End If
'End Sub

Private Sub Winsock1_DataArrival(ByVal datalength As Long)
Text6.Text = ""
Dim b As Byte
Dim j As Byte
Dim fl As Double
For i = 1 To datalength
    Winsock1.GetData b
    MbusByteArray(i) = b
      strxx = strxx & b & " "
        Text6.Text = strxx
Next
j = 0
If MbusRead Then
    If Check.Value Then
   'If MbusByteArray(10) = 201 Then
         For i = 10 To MbusByteArray(9) + 9 Step 4
        ' For i = 13 To 200 + 12 Step 4
        Text4(j).Text = hex2float(MbusByteArray(i), MbusByteArray(i + 1), MbusByteArray(i + 2), MbusByteArray(i + 3))
        j = j + 1
         Next i
       '  End If
Else
       For i = 10 To MbusByteArray(9) + 9 Step 2
'For i = 1 To datalength
'Text1.Text = Str(j) + ": " + " [ " + Str((MbusByteArray(i) * 255) + MbusByteArray(i + 1)) + " ]"
'Text1.Text = Str(j) + ": " + " [ " + Str(MbusByteArray(i)) + " ]"
'List1.AddItem (Text1.Text)
        Text4(j).Text = Str((MbusByteArray(i) * 256) + MbusByteArray(i + 1))
        j = j + 1
        Next i
End If

Text4(59).Text = mbpolling
Text5.Text = "Registers read"
Text5.BackColor = &HFF00&
'For l = j To 89
'Text4(l).Text = "*****"
'Next l
ModbusWait = False
ModbusTimeOut = 0
'Timer1.Enabled = False

DoEvents
    Sleep 10
    If mbpolling = True Then
        Call cmd_read_Click
    End If
End If


    
If MbusWrite Then
If (MbusByteArray(8) = 16) And (MbusByteArray(12) = Val(Text3.Text)) Then
Text5.Text = "Registers written"
Text5.BackColor = &HFF00&
ModbusWait = False
ModbusTimeOut = 0
Timer1.Enabled = False
Else
Text5.Text = "Error writting registers"
Text5.BackColor = &HFF
End If

End If

End Sub
