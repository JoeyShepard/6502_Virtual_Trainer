VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form MemManager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memory Manager"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   524
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LstBreakpoints 
      Height          =   1815
      ItemData        =   "MemManager.frx":0000
      Left            =   4800
      List            =   "MemManager.frx":0002
      TabIndex        =   113
      Top             =   4680
      Width           =   2895
   End
   Begin VB.ListBox LstFunctions 
      Height          =   1815
      ItemData        =   "MemManager.frx":0004
      Left            =   4800
      List            =   "MemManager.frx":0006
      TabIndex        =   110
      Top             =   2520
      Width           =   2895
   End
   Begin VB.ListBox LstSections 
      Height          =   1815
      ItemData        =   "MemManager.frx":0008
      Left            =   4800
      List            =   "MemManager.frx":000A
      TabIndex        =   109
      Top             =   360
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Memory"
      Height          =   2535
      Left            =   120
      TabIndex        =   64
      Top             =   4080
      Width           =   4575
      Begin VB.CommandButton BtnDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   3720
         TabIndex        =   115
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox TxtTitle 
         Height          =   285
         Left            =   720
         MaxLength       =   14
         TabIndex        =   16
         Top             =   720
         Width           =   2775
      End
      Begin VB.PictureBox ColorSelect2 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2160
         ScaleHeight     =   255
         ScaleWidth      =   1335
         TabIndex        =   107
         Top             =   1200
         Width           =   1335
         Begin VB.Shape ColorShape2 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   0
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.TextBox TxtValue 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3120
         MaxLength       =   2
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton BtnSetValue 
         Caption         =   "Set"
         Height          =   375
         Left            =   3720
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtAddress2 
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton BtnSetColor 
         Caption         =   "Set"
         Height          =   375
         Left            =   3720
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox ChkSetCode 
         Caption         =   "Mark as code"
         Height          =   195
         Left            =   2160
         TabIndex        =   22
         Top             =   2220
         Width           =   1335
      End
      Begin VB.TextBox TxtAddress1 
         Height          =   285
         Left            =   720
         MaxLength       =   4
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox ChkSetRO 
         Caption         =   "Mark read-only"
         Height          =   195
         Left            =   2160
         TabIndex        =   21
         Top             =   1920
         Width           =   1575
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   79
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   1200
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   78
         Left            =   360
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   1200
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00002040&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   77
         Left            =   600
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   1200
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00004040&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   76
         Left            =   840
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   1200
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   75
         Left            =   1080
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   1200
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00404000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   74
         Left            =   1320
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   1200
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   73
         Left            =   1560
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   1200
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00400040&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   72
         Left            =   1800
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   1200
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   71
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   1440
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00000080&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   70
         Left            =   360
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   1440
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   69
         Left            =   600
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   1440
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00008080&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   68
         Left            =   840
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   1440
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   67
         Left            =   1080
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   1440
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   66
         Left            =   1320
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   1440
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   65
         Left            =   1560
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   1440
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00800080&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   64
         Left            =   1800
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   1440
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   63
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   1680
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   62
         Left            =   360
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   1680
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H000040C0&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   61
         Left            =   600
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1680
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   60
         Left            =   840
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   1680
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   59
         Left            =   1080
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   1680
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   58
         Left            =   1320
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   1680
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00C00000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   57
         Left            =   1560
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   1680
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00C000C0&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   56
         Left            =   1800
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   1680
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   55
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   1920
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   54
         Left            =   360
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   1920
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   53
         Left            =   600
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   1920
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   52
         Left            =   840
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   1920
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   51
         Left            =   1080
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   1920
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   50
         Left            =   1320
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   1920
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00FF4040&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   49
         Left            =   1560
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   1920
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   48
         Left            =   1800
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   1920
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   47
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   2160
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H006060FF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   46
         Left            =   360
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   2160
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H0040C0FF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   45
         Left            =   600
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2160
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   44
         Left            =   840
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2160
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H0040FF80&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   43
         Left            =   1080
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   2160
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   42
         Left            =   1320
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   2160
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   41
         Left            =   1560
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   2160
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00FF60FF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   40
         Left            =   1800
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   2160
         Width           =   240
      End
      Begin VB.TextBox TxtRGB2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   0
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   17
         Text            =   "FF"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox TxtRGB2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   3120
         MaxLength       =   2
         TabIndex        =   19
         Text            =   "FF"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox TxtRGB2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   285
         Index           =   1
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   18
         Text            =   "FF"
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Title:"
         Height          =   195
         Left            =   120
         TabIndex        =   108
         Top             =   720
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   195
         Left            =   2640
         TabIndex        =   106
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Range:"
         Height          =   195
         Left            =   120
         TabIndex        =   105
         Top             =   270
         Width           =   525
      End
      Begin VB.Shape ShpColor2 
         Height          =   1230
         Left            =   105
         Top             =   1185
         Width           =   1950
      End
   End
   Begin VB.PictureBox ColorSelect 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2280
      ScaleHeight     =   255
      ScaleWidth      =   1335
      TabIndex        =   63
      Top             =   2640
      Width           =   1335
      Begin VB.Shape ColorShape 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Frame FrmAssemble 
      Caption         =   "Assemble"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.TextBox TxtRGB 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   285
         Index           =   1
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "FF"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox TxtRGB 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   3120
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "FF"
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox TxtRGB 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   0
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "FF"
         Top             =   2880
         Width           =   375
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00FF60FF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   39
         Left            =   1800
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   3480
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   38
         Left            =   1560
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   3480
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   37
         Left            =   1320
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   3480
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H0040FF80&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   36
         Left            =   1080
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   3480
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   35
         Left            =   840
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   3480
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H0040C0FF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   34
         Left            =   600
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   3480
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H006060FF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   33
         Left            =   360
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   3480
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   32
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   3480
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   31
         Left            =   1800
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   3240
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00FF4040&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   30
         Left            =   1560
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   3240
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   29
         Left            =   1320
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   3240
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   28
         Left            =   1080
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3240
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   27
         Left            =   840
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   3240
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   26
         Left            =   600
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   3240
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   25
         Left            =   360
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   3240
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   24
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   3240
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00C000C0&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   23
         Left            =   1800
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3000
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00C00000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   22
         Left            =   1560
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   3000
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   21
         Left            =   1320
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   3000
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   20
         Left            =   1080
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   3000
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   19
         Left            =   840
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   3000
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H000040C0&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   18
         Left            =   600
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   3000
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   17
         Left            =   360
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   3000
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   16
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   3000
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00800080&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   15
         Left            =   1800
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2760
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   14
         Left            =   1560
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   2760
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   13
         Left            =   1320
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2760
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   12
         Left            =   1080
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2760
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00008080&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   11
         Left            =   840
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2760
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   10
         Left            =   600
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2760
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00000080&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   9
         Left            =   360
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2760
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   8
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2760
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00400040&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   7
         Left            =   1800
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2520
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   6
         Left            =   1560
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2520
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00404000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   5
         Left            =   1320
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2520
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   4
         Left            =   1080
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2520
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00004040&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   3
         Left            =   840
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2520
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00002040&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   600
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2520
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   360
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2520
         Width           =   240
      End
      Begin VB.PictureBox ColorBox 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2520
         Width           =   240
      End
      Begin VB.CheckBox ChkLoadRO 
         Caption         =   "Mark read-only"
         Height          =   195
         Left            =   2160
         TabIndex        =   8
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox TxtAddress 
         Height          =   285
         Left            =   3720
         MaxLength       =   4
         TabIndex        =   4
         Top             =   2760
         Width           =   615
      End
      Begin VB.CheckBox ChkLoadCode 
         Caption         =   "Mark as code"
         Height          =   195
         Left            =   2160
         TabIndex        =   9
         Top             =   3540
         Width           =   1335
      End
      Begin VB.CommandButton BtnPath 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TxtPath 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "C:\Users\Joey\Desktop\Electronics\6502\compile\list.txt"
         Top             =   240
         Width           =   3735
      End
      Begin VB.CommandButton BtnLoad 
         Caption         =   "Load"
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   3360
         Width           =   735
      End
      Begin RichTextLib.RichTextBox RtfAssemble 
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3201
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         RightMargin     =   35000
         TextRTF         =   $"MemManager.frx":000C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape ShpColor 
         Height          =   1230
         Left            =   105
         Top             =   2505
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Address:"
         Height          =   195
         Left            =   3720
         TabIndex        =   10
         Top             =   2520
         Width           =   615
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Breakpoints"
      Height          =   195
      Left            =   4800
      TabIndex        =   114
      Top             =   4440
      Width           =   840
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Labels"
      Height          =   195
      Left            =   4800
      TabIndex        =   112
      Top             =   2280
      Width           =   465
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Memory regions:"
      Height          =   195
      Left            =   4800
      TabIndex        =   111
      Top             =   120
      Width           =   1155
   End
End
Attribute VB_Name = "MemManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private AddressBegin, AddressEnd

Private Sub BtnDelete_Click()
   Dim flag As Integer
   If Not CheckAddress Then
      For i = AddressBegin To AddressEnd
         RamTitles(i) = ""
         RamColors(i) = vbWhite
         RamAttribs(i) = 0
         RamLabels(i) = ""
         RamOps(i) = ""
      Next i
      RefreshAddressList
      RefreshBreakpointList
      RefreshLabelList
      UpdateTable
   End If
End Sub

Public Sub BtnLoad_Click()
   'On Error GoTo eh
   Dim buff As String
   Dim started As Boolean, quoted As Boolean, comment As Boolean
   Dim bytes(3) As String
   Dim TextRTF As String
   Dim StartedRTF As Boolean
   
   If Emulating = True Then
      ToStopEmulating = True
      ToLoad = True
      Exit Sub
   End If
   
   TextRTF = "{\rtf1\ansi\ansicpg1251\deff0\deflang1049{\fonttbl{\f0\fnil\fcharset0 Courier New;}}"
   TextRTF = TextRTF & vbCrLf & "{\colortbl ;\red96\green96\blue96;\red0\green0\blue192;\red128\green128\blue128;\red0\green0\blue0;\red255\green0\blue0;\red192\green0\blue0;}"
   TextRTF = TextRTF & vbCrLf & "\viewkind4\uc1\pard\cf1\f0\fs17 "
   StartedRTF = False
   
   Open TxtPath.text For Input As #2
      For i = 1 To 4
         Line Input #2, buff
      Next i
      
      'RtfAssemble.text = ""
      Do While Not EOF(2)
         Line Input #2, buff
         oldaddress = Address
         Address = Mid(buff, 3, 4)
         
         If oldaddress <> Address Then
            lblname = ""
         End If
         
         For i = 0 To 3
            bytes(i) = Mid(buff, 12 + i * 3, 2)
         Next i
         opcode = Mid(buff, 25)
         started = False
         quoted = False
         comment = False
         ocfinal = ""
         comfinal = ""

         For i = 1 To Len(opcode)
            a = Mid(opcode, i, 1)
            If Not started And a <> " " Then started = True
            If started Then
               If a = """" Then quoted = Not quoted
               If quoted = False And comment = False And a = ";" Then comment = True
               If comment Then
                  comfinal = comfinal & a
               Else
                  If quoted = True Then
                     ocfinal = ocfinal & a
                  ElseIf a <> " " Or Right(ocfinal, 1) <> " " Then
                     ocfinal = ocfinal & a
                     If a = ":" And lblname = "" Then
                        lblname = ocfinal
                        ocfinal = ""
                        started = False
                     End If
                  End If
               End If
            End If
         Next i
         
         If bytes(0) <> "  " Then RamOps(HexToInt(Address)) = ocfinal
         'else then = "" ?
         For i = 1 To 3
            If bytes(i) <> "  " And HexToInt(Address) + i < 65536 Then
               RamOps(HexToInt(Address) + i) = ""
               RamLabels(HexToInt(Address) + i) = ""
            End If
         Next i
         
         If bytes(0) = "  " And lblname <> "" Then
            TextRTF = TextRTF & "\par \cf2\b " & Left(lblname, Len(lblname) - 1)
            TextRTF = TextRTF & "\cf5 " & ":\b0 " & vbCrLf
            'AddTextAll RtfAssemble, Left(lblname, Len(lblname) - 1), RGB(0, 0, 192), True, rtfLeft, False, False
            'AddTextAll RtfAssemble, ":" & vbCrLf, RGB(255, 0, 0), True, rtfLeft, False, False
            
            'is this right?
            'lblname = ""
         Else
            If StartedRTF Then
               TextRTF = TextRTF & "\par \cf1 "
            End If
            TextRTF = TextRTF & Address & " "
            
            'AddTextAll RtfAssemble, Address & " ", RGB(96, 96, 96), False, rtfLeft, False, False
            For i = 0 To 3
               If bytes(i) = "xx" Then
                  TextRTF = TextRTF & "\cf6\b xx "
                  'AddTextAll RtfAssemble, bytes(i) & " ", RGB(192, 0, 0), True, rtfLeft, False, False
               Else
                  TextRTF = TextRTF & "\cf2\b " & bytes(i) & " "
                  'AddTextAll RtfAssemble, bytes(i) & " ", RGB(0, 0, 192), True, rtfLeft, False, False
               End If
            Next i
            
            If ocfinal <> "" Then
               TextRTF = TextRTF & "\cf4\b0 " & ocfinal
            End If
            TextRTF = TextRTF & "\cf3\b0 " & comfinal & vbCrLf
            
            'AddTextAll RtfAssemble, ocfinal, vbBlack, False, rtfLeft, False, False
            'AddTextAll RtfAssemble, comfinal & vbCrLf, RGB(128, 128, 128), False, rtfLeft, False, False
         End If
         StartedRTF = True
            
            
         For i = 0 To 3
            If bytes(i) = "xx" Then
               SetRAM HexToInt(Address) + i, -1
               If Address <> oldaddress Then RamLabels(HexToInt(Address) + i) = ""
            ElseIf bytes(i) <> "  " Then
               SetRAM HexToInt(Address) + i, HexToInt(bytes(i))
               If Address <> oldaddress Then RamLabels(HexToInt(Address) + i) = ""
            End If
         Next i
         
         If lblname <> "" Then
            RamLabels(HexToInt(Address)) = lblname
            lblname = ""
         End If
      Loop
   Close #2
   RtfAssemble.SelStart = 0
   RefreshLabelList
   
   TextRTF = TextRTF & "\par " & vbCrLf & "\par } & vbcrlf"
   RtfAssemble.TextRTF = TextRTF
   
   Exit Sub
eh:
   If Err = 75 Then
      MsgBox "File not found.", vbCritical
   Else
      MsgBox "Unknown error: " & Error, vbCritical
   End If
   Close #2
End Sub

Private Sub BtnPath_Click()
   On Error GoTo eh
   frmMain.CD1.Filter = "Assembly Listings (*.txt)|*.txt|All Files (*.*)|*.*"
   frmMain.CD1.ShowOpen
   TxtPath.text = frmMain.CD1.filename
   Exit Sub
eh:
   If Err <> 32755 Then
      MsgBox "There was an error opening the file: " & Error, vbCritical
      Close #1
   End If
End Sub

Private Sub BtnSetColor_Click()
   Dim flag As Integer
   If ChkSetRO.value Then flag = AttribReadonly
   If ChkSetCode.value Then flag = flag + AttribCode
   If Not CheckAddress Then
      For i = AddressBegin To AddressEnd
         RamTitles(i) = TxtTitle.text
         RamColors(i) = ColorShape2.BackColor
         If RamAttribs(i) And AttribBreakpoint Then
            RamAttribs(i) = AttribBreakpoint + flag
         Else
            RamAttribs(i) = flag
         End If
      Next i
      RefreshAddressList
      UpdateTable
   End If
End Sub

Private Sub BtnSetValue_Click()
   Dim failed As Boolean
   failed = CheckAddress
   
   If UCase(TxtValue.text) = "X" Then
      value = -1
   Else
      value = HexToInt(TxtValue.text)
      If value > 255 Then value = -1
      If value < 0 Then value = -1
      If value = -1 Then
         failed = True
         MsgBox TxtValue.text & " is not a valid value. Enter a hexadecimal value between 0 and FF, or enter X to mark memory as uninitialized.", vbCritical
      End If
   End If
   
   If failed = False Then
      For i = AddressBegin To AddressEnd
         'should also update peripherals
         SetRAM i, value
         RamOps(i) = ""
         RamLabels(i) = ""
      Next i
      UpdateTable
   End If
End Sub

Private Sub ColorBox_Click(Index As Integer)
   temp = Hex(ColorBox(Index).BackColor \ 2 ^ 16) & " "
   temp = temp & Hex((ColorBox(Index).BackColor \ 2 ^ 8) Mod 256) & " "
   temp = temp & Hex(ColorBox(Index).BackColor And &HFF)
   If Index < 40 Then
      ColorShape.BackColor = ColorBox(Index).BackColor
      ColorSelect.ToolTipText = temp
   Else
      ColorShape2.BackColor = ColorBox(Index).BackColor
      ColorSelect2.ToolTipText = temp
   End If
End Sub

Private Sub Form_Load()
   Dim temp As String
   For i = 0 To 39
      temp = Hex(ColorBox(i).BackColor \ 2 ^ 16) & " "
      temp = temp & Hex((ColorBox(i).BackColor \ 2 ^ 8) Mod 256) & " "
      temp = temp & Hex(ColorBox(i).BackColor And &HFF)
      ColorBox(i).ToolTipText = temp
   Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
   MemManager.Hide
   Cancel = True
End Sub

Private Sub LstBreakpoints_DblClick()
   JumpTable HexToInt(LstBreakpoints.text)
End Sub

Private Sub LstFunctions_DblClick()
   JumpTable HexToInt(Left(LstFunctions.text, 4))
End Sub

Private Sub LstSections_Click()
   TxtAddress1.text = HexBig(RamStarts(SectionList(LstSections.ListIndex)))
   TxtAddress2.text = HexBig(RamEnds(SectionList(LstSections.ListIndex)))
   TxtTitle.text = RamTitles(SectionList(LstSections.ListIndex))
   ColorShape2.BackColor = RamColors(SectionList(LstSections.ListIndex))
   If RamAttribs(SectionList(LstSections.ListIndex)) And AttribReadonly Then
      ChkSetRO.value = 1
   Else
      ChkSetRO.value = 0
   End If
   
   If RamAttribs(SectionList(LstSections.ListIndex)) And AttribCode Then
      ChkSetCode.value = 1
   Else
      ChkSetCode.value = 0
   End If
End Sub

Private Sub LstSections_DblClick()
   If LstSections.ListIndex <> -1 Then
      JumpTable SectionList(LstSections.ListIndex)
   End If
End Sub

Private Sub TxtRGB_Change(Index As Integer)
   Dim failed As Boolean
   failed = False
   For i = 0 To 2
      If HexToInt(TxtRGB(i)) = -1 Then failed = True
   Next i
   If failed = False Then
      ColorSelect.ToolTipText = Hex(Val(TxtRGB(0).text)) & Hex(Val(TxtRGB(1).text)) & Hex(Val(TxtRGB(2).text))
      ColorShape.BackColor = RGB(HexToInt(TxtRGB(0).text), HexToInt(TxtRGB(1).text), HexToInt(TxtRGB(2).text))
   End If
End Sub

Private Sub TxtRGB2_Change(Index As Integer)
   Dim failed As Boolean
   failed = False
   For i = 0 To 2
      If HexToInt(TxtRGB2(i)) = -1 Then failed = True
   Next i
   If failed = False Then
      ColorSelect2.ToolTipText = Hex(Val(TxtRGB2(0).text)) & Hex(Val(TxtRGB2(1).text)) & Hex(Val(TxtRGB2(2).text))
      ColorShape2.BackColor = RGB(HexToInt(TxtRGB2(0).text), HexToInt(TxtRGB2(1).text), HexToInt(TxtRGB2(2).text))
   End If
End Sub

Private Function CheckAddress() As Boolean
   Dim failed As Boolean
   AddressBegin = HexToInt(TxtAddress1.text)
   If AddressBegin > 65535 Then AddressBegin = -1
   If AddressBegin < 0 Then AddressBegin = -1
   If AddressBegin = -1 Then
      failed = True
      MsgBox TxtAddress1.text & " is not a valid value. Enter a hexadecimal value between 0 and FFFF.", vbCritical
   End If
   
   If TxtAddress2.text = "" Then
      AddressEnd = AddressBegin
   Else
      AddressEnd = HexToInt(TxtAddress2.text)
      If AddressEnd > 65535 Then AddressEnd = -1
      If AddressEnd < 0 Then AddressEnd = -1
      If AddressEnd = -1 Then
         failed = True
         MsgBox TxtAddress2.text & " is not a valid value. Enter a hexadecimal value between 0 and FFFF.", vbCritical
      End If
   End If
   
   If AddressBegin > AddressEnd Then
      MsgBox "The starting range must be less than or equal to the end range.", vbCritical
      failed = True
   End If
   
   If failed Then
      CheckAddress = True
   Else
      CheckAddress = False
   End If
End Function
