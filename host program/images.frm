VERSION 5.00
Begin VB.Form images 
   Caption         =   "Form2"
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5280
   LinkTopic       =   "Form2"
   Picture         =   "images.frx":0000
   ScaleHeight     =   323
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   352
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Slider 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   2040
      Picture         =   "images.frx":024A
      ScaleHeight     =   960
      ScaleWidth      =   480
      TabIndex        =   29
      Top             =   2400
      Width           =   480
   End
   Begin VB.PictureBox code_mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3600
      Picture         =   "images.frx":1A8C
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   28
      Top             =   480
      Width           =   195
   End
   Begin VB.PictureBox code 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3360
      Picture         =   "images.frx":1CD6
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   27
      Top             =   480
      Width           =   195
   End
   Begin VB.PictureBox readonly 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3360
      Picture         =   "images.frx":1F20
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   26
      Top             =   240
      Width           =   195
   End
   Begin VB.PictureBox reset_down 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   4440
      Picture         =   "images.frx":216A
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   25
      Top             =   1080
      Width           =   360
   End
   Begin VB.PictureBox reset 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   3960
      Picture         =   "images.frx":286C
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   24
      Top             =   1080
      Width           =   360
   End
   Begin VB.PictureBox HalfCycleDown 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   4440
      Picture         =   "images.frx":2F6E
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   23
      Top             =   720
      Width           =   360
   End
   Begin VB.PictureBox HalfCycle 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   3960
      Picture         =   "images.frx":3670
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   22
      Top             =   720
      Width           =   360
   End
   Begin VB.PictureBox CyclingDown 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   4440
      Picture         =   "images.frx":3D72
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   21
      Top             =   360
      Width           =   360
   End
   Begin VB.PictureBox Cycling 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   3960
      Picture         =   "images.frx":4474
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   20
      Top             =   360
      Width           =   360
   End
   Begin VB.PictureBox play_on 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   4440
      Picture         =   "images.frx":4B76
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   19
      Top             =   0
      Width           =   360
   End
   Begin VB.PictureBox play 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   3960
      Picture         =   "images.frx":5278
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   18
      Top             =   0
      Width           =   360
   End
   Begin VB.PictureBox breakpoint_mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3600
      Picture         =   "images.frx":597A
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   17
      Top             =   0
      Width           =   195
   End
   Begin VB.PictureBox breakpoint 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3360
      Picture         =   "images.frx":5BC4
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   16
      Top             =   0
      Width           =   195
   End
   Begin VB.PictureBox DipSingle 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3120
      Picture         =   "images.frx":5E0E
      ScaleHeight     =   210
      ScaleWidth      =   90
      TabIndex        =   15
      Top             =   1800
      Width           =   90
   End
   Begin VB.PictureBox DipSwitch 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2040
      Picture         =   "images.frx":5F68
      ScaleHeight     =   480
      ScaleWidth      =   990
      TabIndex        =   14
      Top             =   1800
      Width           =   990
   End
   Begin VB.PictureBox LED8bar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   3480
      Picture         =   "images.frx":78AA
      ScaleHeight     =   390
      ScaleWidth      =   90
      TabIndex        =   13
      Top             =   1320
      Width           =   90
   End
   Begin VB.PictureBox LED8off 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2400
      Picture         =   "images.frx":7AF4
      ScaleHeight     =   480
      ScaleWidth      =   1020
      TabIndex        =   12
      Top             =   1200
      Width           =   1020
   End
   Begin VB.PictureBox LEDoff 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   1800
      Picture         =   "images.frx":94B6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   1200
      Width           =   480
   End
   Begin VB.PictureBox LEDon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   1200
      Picture         =   "images.frx":A0F8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   1200
      Width           =   480
   End
   Begin VB.PictureBox keypad 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   0
      Picture         =   "images.frx":AD3A
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   9
      Top             =   1800
      Width           =   1920
   End
   Begin VB.PictureBox switchoff 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   600
      Picture         =   "images.frx":16D7C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   1200
      Width           =   480
   End
   Begin VB.PictureBox switchon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      Picture         =   "images.frx":179BE
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   1200
      Width           =   480
   End
   Begin VB.PictureBox timersmall 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3000
      Picture         =   "images.frx":18600
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   0
      Width           =   240
   End
   Begin VB.PictureBox test 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2400
      Picture         =   "images.frx":18942
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox SevenSegNone 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   1800
      Picture         =   "images.frx":19584
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox SevenSegAll 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      Picture         =   "images.frx":1A1C6
      ScaleHeight     =   390
      ScaleWidth      =   3345
      TabIndex        =   3
      Top             =   600
      Width           =   3345
   End
   Begin VB.PictureBox SevenSeg 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   1200
      Picture         =   "images.frx":1E648
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox buttondown 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   600
      Picture         =   "images.frx":1F28A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox buttonup 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      Picture         =   "images.frx":1FECC
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "images"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
