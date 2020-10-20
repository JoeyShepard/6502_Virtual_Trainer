VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trainer 6502"
   ClientHeight    =   9000
   ClientLeft      =   11040
   ClientTop       =   2670
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   Begin VB.CheckBox ChkNOP 
      Caption         =   "NOP Test"
      Height          =   195
      Left            =   7920
      TabIndex        =   44
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox TxtUpdate 
      Height          =   285
      Left            =   7200
      MaxLength       =   4
      TabIndex        =   42
      Text            =   "100"
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Emulate 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   5280
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   41
      ToolTipText     =   "Start execution"
      Top             =   1260
      Width           =   360
   End
   Begin VB.CheckBox ChkUpdate 
      Caption         =   "Update memory"
      Height          =   195
      Left            =   9240
      TabIndex        =   40
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox ChkJump 
      Caption         =   "Follow access"
      Height          =   195
      Left            =   9240
      TabIndex        =   38
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.PictureBox BtnReset 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   6720
      Picture         =   "Form1.frx":0702
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   37
      ToolTipText     =   "Reset"
      Top             =   1260
      Width           =   360
   End
   Begin VB.CheckBox ChkHispeed 
      Caption         =   "1Mbps"
      Height          =   195
      Left            =   10800
      TabIndex        =   36
      Top             =   1680
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.PictureBox PicCycles 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   5280
      Picture         =   "Form1.frx":0E04
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   35
      Top             =   0
      Width           =   4455
      Begin VB.Label LblPhase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phase: down"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   930
      End
      Begin VB.Line ShpCycles 
         BorderColor     =   &H00808080&
         Tag             =   "free"
         X1              =   293
         X2              =   293
         Y1              =   0
         Y2              =   80
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Break on"
      Height          =   1215
      Left            =   9840
      TabIndex        =   30
      Top             =   0
      Width           =   2055
      Begin VB.CheckBox ChkBreakDatafetch 
         Caption         =   "Data fetch of code"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox ChkBreakROWrite 
         Caption         =   "Write to read-only"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox ChkBreakNoncode 
         Caption         =   "Non-code execution"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox ChkBreakUninit 
         Caption         =   "Uninitialized read"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Value           =   1  'Checked
         Width           =   1575
      End
   End
   Begin VB.ComboBox ComboPorts 
      Height          =   315
      ItemData        =   "Form1.frx":12506
      Left            =   10800
      List            =   "Form1.frx":12543
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   1290
      Width           =   1095
   End
   Begin VB.PictureBox BtnFullCycle 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   6240
      Picture         =   "Form1.frx":125C3
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   28
      ToolTipText     =   "Single step"
      Top             =   1260
      Width           =   360
   End
   Begin VB.PictureBox BtnPlay 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   5760
      Picture         =   "Form1.frx":12CC5
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   27
      ToolTipText     =   "Start stepping"
      Top             =   1260
      Width           =   360
   End
   Begin VB.PictureBox PicPeriphToolbar 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   5280
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   440
      TabIndex        =   7
      Top             =   8340
      Width           =   6600
      Begin VB.PictureBox PeripheralButton 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   11
         Left            =   6090
         Picture         =   "Form1.frx":133C7
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   45
         Tag             =   "multiplier"
         ToolTipText     =   "Timer"
         Top             =   30
         Width           =   480
      End
      Begin VB.PictureBox PeripheralButton 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   10
         Left            =   2070
         Picture         =   "Form1.frx":14009
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   23
         Tag             =   "screen"
         ToolTipText     =   "Graphic Display"
         Top             =   30
         Width           =   480
      End
      Begin VB.PictureBox PeripheralButton 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   9
         Left            =   1560
         Picture         =   "Form1.frx":14C4B
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   22
         Tag             =   "text"
         ToolTipText     =   "Text Display"
         Top             =   30
         Width           =   480
      End
      Begin VB.PictureBox PeripheralButton 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   8
         Left            =   5070
         Picture         =   "Form1.frx":1588D
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   20
         Tag             =   "keyboard"
         ToolTipText     =   "Keyboard"
         Top             =   30
         Width           =   480
      End
      Begin VB.PictureBox PeripheralButton 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   7
         Left            =   4050
         Picture         =   "Form1.frx":164CF
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   19
         Tag             =   "dip"
         ToolTipText     =   "Dip Switch"
         Top             =   30
         Width           =   480
      End
      Begin VB.PictureBox PeripheralButton 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   6
         Left            =   1050
         Picture         =   "Form1.frx":17111
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   18
         Tag             =   "bar"
         ToolTipText     =   "LED Bar"
         Top             =   30
         Width           =   480
      End
      Begin VB.PictureBox PeripheralButton 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   5
         Left            =   30
         Picture         =   "Form1.frx":17D53
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   17
         Tag             =   "led"
         ToolTipText     =   "LED"
         Top             =   30
         Width           =   480
      End
      Begin VB.PictureBox PeripheralButton 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   4
         Left            =   4560
         Picture         =   "Form1.frx":18995
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   16
         Tag             =   "keypad"
         ToolTipText     =   "Keypad"
         Top             =   30
         Width           =   480
      End
      Begin VB.PictureBox PeripheralButton 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   3
         Left            =   3540
         Picture         =   "Form1.frx":195D7
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   15
         Tag             =   "switch"
         ToolTipText     =   "On/Off Switch"
         Top             =   30
         Width           =   480
      End
      Begin VB.PictureBox PeripheralButton 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   2
         Left            =   5580
         Picture         =   "Form1.frx":1A219
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   14
         Tag             =   "ticker"
         ToolTipText     =   "Timer"
         Top             =   30
         Width           =   480
      End
      Begin VB.PictureBox PeripheralButton 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   1
         Left            =   540
         Picture         =   "Form1.frx":1AE5B
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   13
         Tag             =   "7seg"
         ToolTipText     =   "7-Segment Panel"
         Top             =   30
         Width           =   480
      End
      Begin VB.PictureBox PeripheralButton 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   0
         Left            =   3030
         Picture         =   "Form1.frx":1BA9D
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   8
         Tag             =   "button"
         ToolTipText     =   "Push Button"
         Top             =   30
         Width           =   480
      End
   End
   Begin VB.PictureBox PicPeripherals 
      AutoRedraw      =   -1  'True
      BackColor       =   &H007F7F7F&
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   5280
      ScaleHeight     =   425
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   441
      TabIndex        =   6
      Top             =   1920
      Width           =   6615
      Begin MSCommLib.MSComm MSComm1 
         Left            =   1200
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   0   'False
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   600
         Top             =   120
      End
      Begin VB.PictureBox Peripheral 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   0
         Left            =   120
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   9
         Tag             =   "button"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
         Begin VB.Label PeripheralTextOld 
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   165
            Index           =   0
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Visible         =   0   'False
            Width           =   105
         End
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   1920
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Filter          =   "All Files (*.*)|*.*"
      End
      Begin VB.Label PeriphLabel3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label PeriphLabel2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label PeriphLabel1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   525
      End
   End
   Begin VB.VScrollBar ScrollTable 
      Height          =   8625
      LargeChange     =   10
      Left            =   4890
      Max             =   16374
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox PicTable 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   8625
      Left            =   0
      ScaleHeight     =   575
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   326
      TabIndex        =   0
      Top             =   360
      Width           =   4890
      Begin VB.PictureBox BreakPoint 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   0
         Left            =   15
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   26
         Top             =   15
         Width           =   195
      End
      Begin VB.TextBox TxtLbl 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   15
         Width           =   1440
      End
      Begin VB.TextBox TxtDis 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00202020&
         Height          =   195
         Index           =   0
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   15
         Width           =   1800
      End
      Begin VB.TextBox TxtChar 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2670
         MaxLength       =   1
         TabIndex        =   4
         Text            =   "FF"
         Top             =   15
         Width           =   360
      End
      Begin VB.TextBox TxtData 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2250
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "FF"
         Top             =   15
         Width           =   375
      End
      Begin VB.TextBox TxtAddress 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "FFFF"
         Top             =   15
         Width           =   495
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   203
         X2              =   203
         Y1              =   0
         Y2              =   576
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   176
         X2              =   176
         Y1              =   0
         Y2              =   576
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   49
         X2              =   49
         Y1              =   0
         Y2              =   576
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   148
         X2              =   148
         Y1              =   0
         Y2              =   576
      End
      Begin VB.Shape ShpHighlight 
         BorderColor     =   &H0000FF00&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   225
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   4890
      End
   End
   Begin VB.Shape ShpUpload 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   5280
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape ShpUploadBG 
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   5280
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Update"
      Height          =   195
      Left            =   7200
      TabIndex        =   43
      Top             =   1320
      Width           =   525
   End
   Begin VB.Label LblHeading 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Screen Buffer (4000-4FFF)"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   5145
   End
   Begin VB.Menu FileMnu 
      Caption         =   "File"
      Begin VB.Menu OpenItm 
         Caption         =   "Open..."
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu SaveMemItm 
         Caption         =   "Save Memory As..."
      End
      Begin VB.Menu SavePeriphItm 
         Caption         =   "Save Peripherals As..."
      End
   End
   Begin VB.Menu MemMnu 
      Caption         =   "Memory"
      Begin VB.Menu MemManItm 
         Caption         =   "Memory Manager"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu PeripheralMnu 
      Caption         =   "Peripherals"
      Begin VB.Menu UnlockItm 
         Caption         =   "Unlock All"
         Shortcut        =   ^U
      End
      Begin VB.Menu LockItm 
         Caption         =   "Lock All"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu ContextMnu 
      Caption         =   "Context Menu"
      Visible         =   0   'False
      Begin VB.Menu DeleteItm 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu RightClickMnu 
      Caption         =   "Right Click Menu"
      Visible         =   0   'False
      Begin VB.Menu ROItm 
         Caption         =   "Read-only"
      End
      Begin VB.Menu CodeItm 
         Caption         =   "Code"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu BPItm 
         Caption         =   "Breakpoint"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BPItm_Click()
      If RamAttribs(RightClickPtr) And AttribBreakpoint Then
         RamAttribs(RightClickPtr) = RamAttribs(RightClickPtr) - AttribBreakpoint
         UpdateCell RightClickIndex
      Else
         RamAttribs(RightClickPtr) = RamAttribs(RightClickPtr) Or AttribBreakpoint
         UpdateCell RightClickIndex
      End If
      RefreshBreakpointList
End Sub

Private Sub BreakPoint_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim j As Long
   j = ScrollTable.value
   j = j * 4 + Index
   If Button = vbLeftButton Then
      If RamAttribs(j) And AttribBreakpoint Then
         'is this even right?
         'RamAttribs(j) = RamAttribs(j) And (AttribReadonly Or AttribCode)
         RamAttribs(j) = RamAttribs(j) - AttribBreakpoint
         UpdateCell Index
      Else
         RamAttribs(j) = RamAttribs(j) Or AttribBreakpoint
         UpdateCell Index
      End If
      RefreshBreakpointList
   ElseIf Button = vbRightButton Then
      If RamAttribs(j) And AttribReadonly Then
         ROItm.Checked = True
      Else
         ROItm.Checked = False
      End If
      
      If RamAttribs(j) And AttribCode Then
         CodeItm.Checked = True
      Else
         CodeItm.Checked = False
      End If
      
      If RamAttribs(j) And AttribBreakpoint Then
         BPItm.Checked = True
      Else
         BPItm.Checked = False
      End If
      RightClickPtr = j
      RightClickIndex = Index
      PopupMenu RightClickMnu
   End If
End Sub

Private Sub BtnFullCycle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   BtnFullCycle.Picture = images.HalfCycleDown.Picture
End Sub

Private Sub BtnFullCycle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If ExecuteUART = False Then
      If EnableUART Then
         JustCycling = True
         SingleCycle
         JustCycling = False
      End If
      DisableUART
   End If
   BtnFullCycle.Picture = images.HalfCycle.Picture
End Sub

Private Sub BtnPlay_Click()
   If ExecuteUART = False Then
      If EnableUART Then
         GlobalTime = GetTickCount
         CycleCount = 0
         frmMain.BtnPlay.Picture = images.CyclingDown.Picture
         frmMain.BtnPlay.ToolTipText = "Stop stepping"
      End If
   Else
      'DisableUART
      ToDisableUART = True
   End If
End Sub

Private Sub BtnReset_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   BtnReset.Picture = images.reset_down.Picture
   Open App.Path & "\crash log.txt" For Output As #10
   Close #10
End Sub

Public Sub BtnReset_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Emulating = True Then
      ToStopEmulating = True
      ToReset = True
      Exit Sub
   End If
   
   Dim buff As String
   dbug "DisableUART"
   DisableUART
   If EnableUART Then
      JustCycling = True
      dbug "Set output"
      frmMain.MSComm1.Output = FormMsg(RESET_CPU, "", PACKET_SIZE)
      dbug "WaitInput"
      If WaitInput Then 'successfully got input
         dbug "Get input"
         buff = frmMain.MSComm1.Input
         'MsgBox Hex(Asc(Mid(buff, 1, 1))) & Hex(Asc(Mid(buff, 2, 1))) & vbCrLf & Hex(Asc(Mid(buff, 3, 1))) & Hex(Asc(Mid(buff, 4, 1)))
         HighlightPtr = HexToInt("FFFC")
         If frmMain.ChkJump.value = 1 Then JumpTable HighlightPtr
         MoveHighlightPtr
         If ChkNOP.value = 1 Then
            Open App.Path & "\NOP Test.txt" For Output As #25
               Print #25, Date & " " & Time
            Close #25
         End If
      Else
         MsgBox "Reset command timed out.", vbCritical
      End If
   End If
   JustCycling = False
   dbug "DisableUART 2"
   DisableUART
   BtnReset.Picture = images.reset.Picture
   
End Sub

Private Sub BtnTest_Click()
   Dim failed As Boolean
   failed = False
   BtnTest.Caption = "Running..."
   If EnableUART Then
      If True Then
         frmMain.MSComm1.Output = FormMsg(CUSTOM_CHECK, "", PACKET_SIZE)
         If WaitInput(10000) Then
            buff = frmMain.MSComm1.Input
            If Asc(Left(buff, 1)) = 1 Then
               MsgBox "Test successful"
            Else
               MsgBox "Failed" & vbCrLf & "Expected: " & HexBig(Asc(Mid(buff, 2, 1)) + Asc(Mid(buff, 3, 1)) * 256) & vbCrLf & "Received: " & HexBig(Asc(Mid(buff, 4, 1)) + Asc(Mid(buff, 5, 1)) * 256) & vbCrLf & "Count: " & Asc(Mid(buff, 6, 1))
            End If
         Else
            MsgBox "Timeout", vbCritical
         End If
      ElseIf False Then
         For i = 0 To 500
            testmsg = ""
            For j = 1 To 16
               testmsg = testmsg & Chr(65 + 26 * Rnd)
            Next j
            MSComm1.Output = FormMsg(COMM_CHECK, ByVal testmsg, PACKET_SIZE)
            If WaitInput Then
               buff = MSComm1.Input
               If Left(buff, 16) <> Left(testmsg, 16) Then
                  failed = True
                  Exit For
               End If
            Else
               MsgBox "Connection to device timed out.", vbCritical
               Exit For
            End If
         Next i
         If failed Then
            MsgBox "Mismatch at " & i, vbCritical
         Else
            MsgBox "All tests passed"
         End If
      End If
      DisableUART
   End If
   BtnTest.Caption = "Test"
End Sub

Private Sub CodeItm_Click()
   If RamAttribs(RightClickPtr) And AttribCode Then
      RamAttribs(RightClickPtr) = RamAttribs(RightClickPtr) - AttribCode
      UpdateCell RightClickIndex
   Else
      RamAttribs(RightClickPtr) = RamAttribs(RightClickPtr) Or AttribCode
      UpdateCell RightClickIndex
   End If
End Sub

Private Sub DeleteItm_Click()
   PeriphData(PeriphToDelete).Deleted = True
   Peripheral(PeriphToDelete).Visible = False
   PeriphLabel1(PeriphToDelete).Visible = False
   PeriphLabel2(PeriphToDelete).Visible = False
   PeriphLabel3(PeriphToDelete).Visible = False
End Sub

Private Sub Emulate_Click()
   'Put button down before starting
   Dim temp As String, crc(4) As Long
   Dim i As Long, j As Long
   Dim buff As String
   Dim failed As Boolean
   Dim SendLimit As String
   Dim ProgressStep As Integer
   Dim DataLengths(5) As Long
   Dim DataCount As Long
   Dim Index As Integer
   
   ProgressStep = ShpUploadBG.Width / 10
   If ExecuteUART = False Then
      ShpUpload.Width = 0
      ShpUploadBG.Visible = True
      ShpUpload.Visible = True
      If Val(TxtUpdate.text) < 0 Or Val(TxtUpdate.text) > 5000 Then
         MsgBox "Enter an update period between 0 and 5000.", vbCritical
      ElseIf EnableUART Then
         ShpUpload.Width = ProgressStep
         failed = False
         frmMain.Emulate.Picture = images.play_on.Picture
         frmMain.Emulate.ToolTipText = "Stop execution"
         TxtUpdate.Enabled = False
         j = HexToInt("FFFF")
         
         For i = 0 To j
            Index = i \ 14336
            If RAM(i) = -1 Then
               temp = temp & Chr((RamAttribs(i) And &HFF) Or AttribUninitialized)
               If Not CompactMode Then temp = temp & Chr(0)
               crc(Index) = crc(Index) + ((RamAttribs(i) And &HFF) Or AttribUninitialized)
               If Not CompactMode Then DataCount = DataCount + 1
            Else
               temp = temp & Chr(RamAttribs(i) And &HFF)
               temp = temp & Chr(RAM(i))
               'crc(i \ HexToInt("3800")) = crc(i \ HexToInt("3800")) + RAM(i)
               crc(Index) = crc(Index) + RAM(i)
               crc(Index) = crc(Index) + (RamAttribs(i)) ' And &HFF)
               DataCount = DataCount + 1
            End If
            DataCount = DataCount + 1
            'crc(i \ 14336) = crc(i \ 14336) + Asc(Left(Right(temp, 2), 1))
            
            If (i + 1) Mod 14336 = 0 Then
               DataLengths(Index + 1) = DataCount
            End If
            
            If i = j \ 4 Then
               ShpUpload.Width = ProgressStep * 2
            ElseIf i = 2 * j \ 4 Then
               ShpUpload.Width = ProgressStep * 3
            ElseIf i = 3 * j \ 4 Then
               ShpUpload.Width = ProgressStep * 4
            End If
         Next i
         DataLengths(5) = DataCount
         
         ShpUpload.Width = ProgressStep * 5
         
         i = Val(TxtUpdate.text)
         frmMain.MSComm1.Output = FormMsg(UPDATE_RAM, "", PACKET_SIZE)
         
         If WaitInput Then
            buff = MSComm1.Input
            For i = 0 To 4
               If failed = False Then
                  If i = 4 Then
                     SendLimit = "4000"
                  Else
                     SendLimit = "7000"
                  End If
                  
                  'bloop = bloop & DataLengths(i + 1) - DataLengths(i) & vbCrLf
                  'If i = 4 Then MsgBox bloop
                  
                  'MsgBox DataLengths(i + 1) - DataLengths(i)
                  'frmMain.MSComm1.Output = Mid(temp, 1 + i * HexToInt("7000"), HexToInt("7000")) & Chr(crc(i) Mod 256)
                  
                  frmMain.MSComm1.Output = Mid(temp, 1 + DataLengths(i), DataLengths(i + 1) - DataLengths(i)) & Chr(crc(i) Mod 256)
                  Do While frmMain.MSComm1.OutBufferCount > 0
                     'This causes it to service sub main
                     'DoEvents
                  Loop
                  
                  If WaitInput(1000) Then
                     buff = MSComm1.Input
                     If Asc(Mid(buff, PACKET_SIZE, 1)) = UPDATE_RAM_CRC Then
                        j = Asc(Mid(buff, 3, 1)) + Asc(Mid(buff, 4, 1)) * 256
                        'Maybe this puts less stress on mcu
                        If False And j <> DataLengths(i + 1) - DataLengths(i) Then
                           failed = True
                           MsgBox j & " bytes reported received. " & DataLengths(i + 1) - DataLengths(i) & " bytes expected.", vbCritical
                           DisableUART
                        ElseIf Asc(Left(buff, 1)) <> crc(i) Mod 256 Then
                           failed = True
                           MsgBox "Checksums for the transferred memory did not match." & vbCrLf & "Crc[" & HexBig(i * &H3800) & ":" & HexBig(i * &H3800 + HexToInt(SendLimit) / 2 - 1) & "] = " & crc(i) Mod 256 & vbCrLf & "Received crc = " & Asc(Left(buff, 1)), vbCritical
                           DisableUART
                        Else

                        End If
                     Else
                        'MsgBox "derp what?" & vbCrLf & Asc(Mid(buff, PACKET_SIZE, 1))
                     End If
                  Else
                     failed = True
                     MsgBox "Connection to device timed out.", vbCritical
                     DisableUART
                  End If
               End If
               If i = 4 Then
                  ShpUpload.Width = ShpUploadBG.Width
               Else
                  ShpUpload.Width = ProgressStep * (6 + i)
               End If
            Next i
            
            'check if failed
            
            MSComm1.Output = FormMsg(GET_RAM_CRC, "", PACKET_SIZE)
            If WaitInput(1000) Then
               buff = MSComm1.Input
               
               For i = 0 To 4
                  If failed = False Then
                     If crc(i) Mod 256 <> Asc(Mid(buff, 1 + i, 1)) Then
                        failed = True
                        MsgBox "Checksums failed rereading RAM.", vbCritical
                        DisableUART
                     End If
                  End If
               Next i
            Else
               failed = True
               MsgBox "Connection to device timed out.", vbCritical
               DisableUART
            End If
             
            If failed = False Then
               i = 0
               If ChkBreakUninit.value = 1 Then i = i Or AttribUninitialized
               If frmMain.ChkBreakROWrite.value = 1 Then i = i Or AttribReadonly
                                                          
               j = Val(TxtUpdate.text) * 12
                                                          
               frmMain.MSComm1.Output = FormMsg(BEGIN_EMULATING, Chr(i) & Chr(j And &HFF) & Chr(j \ 256), PACKET_SIZE)
               EmuCount = 0
               'DoEvents 'needed?
               GlobalTime = GetTickCount
               CycleCount = 0
               For i = 0 To 10
                  TextsToUpdate(i) = -1
               Next i
               Emulating = True
               
            End If
         Else
            MsgBox "Connection to device timed out.", vbCritical
            DisableUART
         End If
      End If
      ShpUploadBG.Visible = False
      ShpUpload.Visible = False
   Else
      ToStopEmulating = True
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Not IgnoreKey Then
      For i = 1 To PeriphCount
         If PeriphData(i).ptype = Keyboard Then
            If PeriphData(i).Address <> -1 Then
               If RAM(PeriphData(i).Address) <> KeyCode Then
                  If ExecuteUART And Emulating Then AddToEmuBuff = True
                  SetRAM PeriphData(i).Address, KeyCode
                  AddToEmuBuff = False
               End If
            End If
         End If
      Next i
      If ExecuteUART = False And KeyCode = vbKeySpace Then
         Call BtnFullCycle_MouseUp(vbLeftButton, 0, 0, 0)
      End If
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If Not IgnoreKey Then
      For i = 1 To PeriphCount
         If PeriphData(i).ptype = Keyboard Then
            If PeriphData(i).Address <> -1 Then
               If ExecuteUART And Emulating Then AddToEmuBuff = True
               SetRAM PeriphData(i).Address, 0
               AddToEmuBuff = False
            End If
         End If
      Next i
   End If
End Sub

Private Sub Form_Load()
   Dim t2 As String
   
   Load images
   Load MemManager
   
   'Is this supposed to be here?
   MoveHighlightPtr
   
   For i = 1 To RowCount
      Load TxtAddress(i)
      TxtAddress(i).Top = (TxtAddress(0).Height + 1) * i + 1
      TxtAddress(i).Visible = True
      
      Load TxtData(i)
      TxtData(i).Top = (TxtData(0).Height + 1) * i + 1
      TxtData(i).Visible = True
      
      Load TxtChar(i)
      TxtChar(i).Top = (TxtChar(0).Height + 1) * i + 1
      TxtChar(i).Visible = True
   
      Load TxtDis(i)
      TxtDis(i).Top = (TxtDis(0).Height + 1) * i + 1
      TxtDis(i).Visible = True
      
      Load TxtLbl(i)
      TxtLbl(i).Top = (TxtLbl(0).Height + 1) * i + 1
      TxtLbl(i).Visible = True
      
      Load breakpoint(i)
      breakpoint(i).Top = (breakpoint(0).Height + 1) * i + 1
      breakpoint(i).Visible = True
   Next i
   
   For i = 0 To PicPeripherals.Width Step 16
      For j = 0 To PicPeripherals.Height Step 16
         PicPeripherals.PSet (i, j), vbBlack
      Next j
   Next i
   
   frmMain.ComboPorts.ListIndex = 0
   
   For i = 0 To 65535
      RAM(i) = -1
      RamColors(i) = vbWhite
   Next i
   
   UpdateTable
   
   OpenFile App.Path & "\save\test_mem.mem"
   OpenFile App.Path & "\save\test_per.per"
   
   Open App.Path & "\log.txt" For Output As #100
   Close #100
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload images
   Unload MemManager
   If frmMain.MSComm1.PortOpen = True Then frmMain.MSComm1.PortOpen = False
   Close #100
   End
End Sub

Private Sub LockItm_Click()
 For i = 1 To PeriphCount
      If PeriphData(i).Deleted = False Then
         PeriphData(i).Locked = True
      End If
   Next i
End Sub

Private Sub MemManItm_Click()
   RefreshAddressList
   RefreshLabelList
   RefreshBreakpointList
   MemManager.Show
End Sub

Private Sub OpenItm_Click()
   On Error GoTo eh
   CD1.Filter = "Trainer Files (*.mem; *.per)|*.mem;*.per|All Files (*.*)|*.*"
   CD1.ShowOpen
   OpenFile CD1.filename
   Exit Sub
eh:
   If Err <> 32755 Then
      MsgBox "There was an error opening the file: " & Error, vbCritical
      Close #1
   End If
End Sub



Private Sub Peripheral_Click(Index As Integer)
   Exit Sub
   t = GetTickCount
   For j = 0 To 5000
      For i = 0 To 100
         Peripheral(Index).Line (i, i)-(i + 1, i + 1), vbRed, BF
         'Peripheral(Index).PSet (i, i), vbRed
      Next i
   Next j
   MsgBox (GetTickCount - t) / 1000
   
   t = GetTickCount
   For j = 0 To 5000
      For i = 0 To 100
         SetPixel Peripheral(Index).hDC, i, i, RGB(0, 0, 255)
         SetPixel Peripheral(Index).hDC, i + 1, i, vbBlue
         SetPixel Peripheral(Index).hDC, i, i + 1, vbBlue
         SetPixel Peripheral(Index).hDC, i + 1, i + 1, vbBlue
      Next i
   Next j
   Peripheral(Index).Refresh
   MsgBox (GetTickCount - t) / 1000
   
End Sub

Private Sub Peripheral_DblClick(Index As Integer)
   If PeriphData(Index).Address <> -1 Then JumpTable PeriphData(Index).Address
End Sub

Private Sub Peripheral_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
   Call PicPeripherals_DragDrop(Peripheral(Index), Peripheral(Index).Left + x, Peripheral(Index).Top + y)
End Sub

Private Sub Peripheral_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim update As Boolean
   update = False
   If Button = vbLeftButton Then
      If PeriphData(Index).Locked Then
         If PeriphData(Index).ptype = PushButton Then
            Peripheral(Index).Picture = images.buttondown.Picture
            For i = 0 To 4
               If PeriphData(Index).DownAddress(i) <> -1 Then
                  If PeriphData(Index).DownValue(i) <> -1 Then
                     If ExecuteUART And Emulating Then AddToEmuBuff = True
                     SetRAM PeriphData(Index).DownAddress(i), PeriphData(Index).DownValue(i)
                     AddToEmuBuff = False
                  End If
               End If
            Next i
         ElseIf PeriphData(Index).ptype = SwitchButton Then
            If PeriphData(Index).switchon Then
               PeriphData(Index).switchon = False
               Peripheral(Index).Picture = images.switchoff
               For i = 0 To 4
               If PeriphData(Index).UpAddress(i) <> -1 Then
                  If PeriphData(Index).UpValue(i) <> -1 Then
                     If ExecuteUART And Emulating Then AddToEmuBuff = True
                     SetRAM PeriphData(Index).UpAddress(i), PeriphData(Index).UpValue(i)
                     AddToEmuBuff = False
                  End If
               End If
            Next i
            Else
               PeriphData(Index).switchon = True
               Peripheral(Index).Picture = images.switchon
               For i = 0 To 4
               If PeriphData(Index).DownAddress(i) <> -1 Then
                  If PeriphData(Index).DownValue(i) <> -1 Then
                     If ExecuteUART And Emulating Then AddToEmuBuff = True
                     SetRAM PeriphData(Index).DownAddress(i), PeriphData(Index).DownValue(i)
                     AddToEmuBuff = False
                  End If
               End If
            Next i
            End If
         ElseIf PeriphData(Index).ptype = keypad Then
            If PeriphData(Index).Address <> -1 Then
               If ExecuteUART And Emulating Then AddToEmuBuff = True
               SetRAM PeriphData(Index).Address, 16 - ((3 - Int(x / Peripheral(Index).Width * 4)) + Int(y / Peripheral(Index).Height * 4) * 4)
               AddToEmuBuff = False
            End If
         ElseIf PeriphData(Index).ptype = DipSwitch Then
            
            i = 7 - Int(x / Peripheral(Index).Width * 8)
            
            If PeriphData(Index).DipValue And 2 ^ i Then
               PeriphData(Index).DipValue = PeriphData(Index).DipValue - 2 ^ i
            Else
               PeriphData(Index).DipValue = PeriphData(Index).DipValue + 2 ^ i
            End If
            
            If PeriphData(Index).Address <> -1 Then
               If ExecuteUART And Emulating Then AddToEmuBuff = True
               SetRAM PeriphData(Index).Address, PeriphData(Index).DipValue
               AddToEmuBuff = False
            End If
            
            UpdateDip Index
         End If
      Else
         StartX = x
         StartY = y
         Peripheral(Index).Drag
      End If
   ElseIf Button = vbRightButton Then
      If ExecuteUART = False Then
         If Shift = 0 Then
            DlgPtr = Index
            
            If PeriphData(Index).ptype = PushButton Then
               Load DlgButton
               
               For i = 0 To 4
                  If PeriphData(DlgPtr).DownAddress(i) <> -1 Then DlgButton.DownAddress(i).text = HexBig(PeriphData(DlgPtr).DownAddress(i))
                  If PeriphData(DlgPtr).UpAddress(i) <> -1 Then DlgButton.UpAddress(i).text = HexBig(PeriphData(DlgPtr).UpAddress(i))
                  If PeriphData(DlgPtr).DownValue(i) <> -1 Then DlgButton.DownValue(i).text = Hex(PeriphData(DlgPtr).DownValue(i))
                  If PeriphData(DlgPtr).UpValue(i) <> -1 Then DlgButton.UpValue(i).text = Hex(PeriphData(DlgPtr).UpValue(i))
               Next i
               
               If PeriphData(Index).Locked Then
                  DlgButton.LockedCheck.value = 1
               Else
                  DlgButton.LockedCheck.value = 0
               End If
               
               DlgButton.Captions(0).text = PeriphData(DlgPtr).Labels(0)
               DlgButton.Captions(1).text = PeriphData(DlgPtr).Labels(1)
               DlgButton.Captions(2).text = PeriphData(DlgPtr).Labels(2)
                
               DlgButton.Left = frmMain.Left + 3000
               DlgButton.Top = frmMain.Top + 3000
               DlgButton.Show 1
            ElseIf PeriphData(Index).ptype = SevenSeg Then
               Load Dlg7seg
               
               If PeriphData(Index).Locked Then
                  Dlg7seg.LockedCheck.value = 1
               Else
                  Dlg7seg.LockedCheck.value = 0
               End If
               
               If PeriphData(DlgPtr).Address <> -1 Then Dlg7seg.Address.text = HexBig(PeriphData(DlgPtr).Address)
               
               Dlg7seg.Captions(0).text = PeriphData(DlgPtr).Labels(0)
               Dlg7seg.Captions(1).text = PeriphData(DlgPtr).Labels(1)
               Dlg7seg.Captions(2).text = PeriphData(DlgPtr).Labels(2)
                
               Dlg7seg.Left = frmMain.Left + 3000
               Dlg7seg.Top = frmMain.Top + 3000
               Dlg7seg.Show 1
            ElseIf PeriphData(Index).ptype = Ticker Then
               Load DlgTimer
               
               If PeriphData(Index).Locked Then
                  DlgTimer.LockedCheck.value = 1
               Else
                  DlgTimer.LockedCheck.value = 0
               End If
               
               If PeriphData(Index).Ticker16 Then
                  DlgTimer.Check16bit.value = 1
               Else
                  DlgTimer.Check16bit.value = 0
               End If
               
               If PeriphData(DlgPtr).Address <> -1 Then DlgTimer.Address.text = HexBig(PeriphData(DlgPtr).Address)
               If PeriphData(DlgPtr).TickerInterval <> -1 Then DlgTimer.Interval.text = PeriphData(DlgPtr).TickerInterval
               
               DlgTimer.Captions(0).text = PeriphData(DlgPtr).Labels(0)
               DlgTimer.Captions(1).text = PeriphData(DlgPtr).Labels(1)
               DlgTimer.Captions(2).text = PeriphData(DlgPtr).Labels(2)
               
               Timer1.Enabled = False
               
               DlgTimer.Left = frmMain.Left + 3000
               DlgTimer.Top = frmMain.Top + 3000
               DlgTimer.Show 1
            ElseIf PeriphData(Index).ptype = SwitchButton Then
               Load DlgSwitch
               
               For i = 0 To 4
                  If PeriphData(DlgPtr).DownAddress(i) <> -1 Then DlgSwitch.DownAddress(i).text = HexBig(PeriphData(DlgPtr).DownAddress(i))
                  If PeriphData(DlgPtr).UpAddress(i) <> -1 Then DlgSwitch.UpAddress(i).text = HexBig(PeriphData(DlgPtr).UpAddress(i))
                  If PeriphData(DlgPtr).DownValue(i) <> -1 Then DlgSwitch.DownValue(i).text = Hex(PeriphData(DlgPtr).DownValue(i))
                  If PeriphData(DlgPtr).UpValue(i) <> -1 Then DlgSwitch.UpValue(i).text = Hex(PeriphData(DlgPtr).UpValue(i))
               Next i
               
               If PeriphData(Index).Locked Then
                  DlgSwitch.LockedCheck.value = 1
               Else
                  DlgSwitch.LockedCheck.value = 0
               End If
               
               DlgSwitch.Captions(0).text = PeriphData(DlgPtr).Labels(0)
               DlgSwitch.Captions(1).text = PeriphData(DlgPtr).Labels(1)
               DlgSwitch.Captions(2).text = PeriphData(DlgPtr).Labels(2)
                
               DlgSwitch.Left = frmMain.Left + 3000
               DlgSwitch.Top = frmMain.Top + 3000
               DlgSwitch.Show 1
            ElseIf PeriphData(Index).ptype = keypad Then
               Load DlgKeypad
                        
               If PeriphData(Index).Locked Then
                  DlgKeypad.LockedCheck.value = 1
               Else
                  DlgKeypad.LockedCheck.value = 0
               End If
               
               If PeriphData(DlgPtr).Address <> -1 Then DlgKeypad.Address.text = HexBig(PeriphData(DlgPtr).Address)
               
               DlgKeypad.Captions(0).text = PeriphData(DlgPtr).Labels(0)
               DlgKeypad.Captions(1).text = PeriphData(DlgPtr).Labels(1)
               DlgKeypad.Captions(2).text = PeriphData(DlgPtr).Labels(2)
                
               DlgKeypad.Left = frmMain.Left + 3000
               DlgKeypad.Top = frmMain.Top + 3000
               DlgKeypad.Show 1
            ElseIf PeriphData(Index).ptype = LED Then
               Load DlgLED
               
               For i = 0 To 5
                  If PeriphData(DlgPtr).LEDvalue(i) <> -1 Then DlgLED.Limit(i).text = Hex(PeriphData(DlgPtr).LEDvalue(i))
                  DlgLED.Compare(i).ListIndex = PeriphData(DlgPtr).LEDrelation(i)
               Next i
               
               If PeriphData(Index).Locked Then
                  DlgLED.LockedCheck.value = 1
               Else
                  DlgLED.LockedCheck.value = 0
               End If
               
               If PeriphData(DlgPtr).Address <> -1 Then DlgLED.Address.text = HexBig(PeriphData(DlgPtr).Address)
               
               DlgLED.Captions(0).text = PeriphData(DlgPtr).Labels(0)
               DlgLED.Captions(1).text = PeriphData(DlgPtr).Labels(1)
               DlgLED.Captions(2).text = PeriphData(DlgPtr).Labels(2)
                
               DlgLED.Left = frmMain.Left + 3000
               DlgLED.Top = frmMain.Top + 3000
               DlgLED.Show 1
            ElseIf PeriphData(Index).ptype = LED8 Then
               Load DlgLED8
               
               If PeriphData(Index).Locked Then
                  DlgLED8.LockedCheck.value = 1
               Else
                  DlgLED8.LockedCheck.value = 0
               End If
               
               If PeriphData(DlgPtr).Address <> -1 Then DlgLED8.Address.text = HexBig(PeriphData(DlgPtr).Address)
               
               DlgLED8.Captions(0).text = PeriphData(DlgPtr).Labels(0)
               DlgLED8.Captions(1).text = PeriphData(DlgPtr).Labels(1)
               DlgLED8.Captions(2).text = PeriphData(DlgPtr).Labels(2)
                
               DlgLED8.Left = frmMain.Left + 3000
               DlgLED8.Top = frmMain.Top + 3000
               DlgLED8.Show 1
            ElseIf PeriphData(Index).ptype = DipSwitch Then
               Load DlgDip
                        
               If PeriphData(Index).Locked Then
                  DlgDip.LockedCheck.value = 1
               Else
                  DlgDip.LockedCheck.value = 0
               End If
               
               If PeriphData(DlgPtr).Address <> -1 Then DlgDip.Address.text = HexBig(PeriphData(DlgPtr).Address)
               
               DlgDip.Captions(0).text = PeriphData(DlgPtr).Labels(0)
               DlgDip.Captions(1).text = PeriphData(DlgPtr).Labels(1)
               DlgDip.Captions(2).text = PeriphData(DlgPtr).Labels(2)
                
               DlgDip.Left = frmMain.Left + 3000
               DlgDip.Top = frmMain.Top + 3000
               DlgDip.Show 1
            ElseIf PeriphData(Index).ptype = Keyboard Then
               Load DlgKeyboard
               
               With DlgKeyboard
               
               If PeriphData(Index).Locked Then
                  .LockedCheck.value = 1
               Else
                  .LockedCheck.value = 0
               End If
               
               If PeriphData(DlgPtr).Address <> -1 Then .Address.text = HexBig(PeriphData(DlgPtr).Address)
               
               .Captions(0).text = PeriphData(DlgPtr).Labels(0)
               .Captions(1).text = PeriphData(DlgPtr).Labels(1)
               .Captions(2).text = PeriphData(DlgPtr).Labels(2)
                
               .Left = frmMain.Left + 3000
               .Top = frmMain.Top + 3000
               .Show 1
               End With
            ElseIf PeriphData(Index).ptype = TextDisplay Then
               Load DlgText
               
               With DlgText
               
               If PeriphData(Index).Locked Then
                  .LockedCheck.value = 1
               Else
                  .LockedCheck.value = 0
               End If
               
               If PeriphData(DlgPtr).Address <> -1 Then .Address.text = HexBig(PeriphData(DlgPtr).Address)
               
               .HeightValue.text = PeriphData(DlgPtr).TextHeight
               .WidthValue.text = PeriphData(DlgPtr).TextWidth
               
               .Captions(0).text = PeriphData(DlgPtr).Labels(0)
               .Captions(1).text = PeriphData(DlgPtr).Labels(1)
               .Captions(2).text = PeriphData(DlgPtr).Labels(2)
                
               .Left = frmMain.Left + 3000
               .Top = frmMain.Top + 3000
               .Show 1
               End With
            ElseIf PeriphData(Index).ptype = Multiplier Then
               Load DlgMultiplier
               
               With DlgMultiplier
               
               If PeriphData(Index).Locked Then
                  .LockedCheck.value = 1
               Else
                  .LockedCheck.value = 0
               End If
               
               If PeriphData(DlgPtr).Address <> -1 Then .Output.text = HexBig(PeriphData(DlgPtr).Address)
               If PeriphData(DlgPtr).DownAddress(0) <> -1 Then .Input1.text = HexBig(PeriphData(DlgPtr).DownAddress(0))
               If PeriphData(DlgPtr).UpAddress(0) <> -1 Then .Input2.text = HexBig(PeriphData(DlgPtr).UpAddress(0))
               
               If PeriphData(DlgPtr).UseBCD Then
                  .BCDCheck = 1
               Else
                  .BCDCheck = 0
               End If
               
               .Captions(0).text = PeriphData(DlgPtr).Labels(0)
               .Captions(1).text = PeriphData(DlgPtr).Labels(1)
               .Captions(2).text = PeriphData(DlgPtr).Labels(2)
                
               .Left = frmMain.Left + 3000
               .Top = frmMain.Top + 3000
               .Show 1
               End With
            ElseIf PeriphData(Index).ptype = ScreenDisplay Then
               Load DlgScreen
               
               With DlgScreen
               
               If PeriphData(Index).Locked Then
                  .LockedCheck.value = 1
               Else
                  .LockedCheck.value = 0
               End If
               
               If PeriphData(DlgPtr).Address <> -1 Then .Address.text = HexBig(PeriphData(DlgPtr).Address)
               
               .HeightValue.text = PeriphData(DlgPtr).ScreenHeight
               .WidthValue.text = PeriphData(DlgPtr).ScreenWidth
               
               .Resolution.ListIndex = PeriphData(DlgPtr).ScreenRes
               
               .Captions(0).text = PeriphData(DlgPtr).Labels(0)
               .Captions(1).text = PeriphData(DlgPtr).Labels(1)
               .Captions(2).text = PeriphData(DlgPtr).Labels(2)
                
               .Left = frmMain.Left + 3000
               .Top = frmMain.Top + 3000
               .Show 1
               End With
            End If
         Else
            PeriphToDelete = Index
            PopupMenu ContextMnu
         End If
      End If
   End If
End Sub

Private Sub Peripheral_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
      If PeriphData(Index).Locked Then
         If PeriphData(Index).ptype = PushButton Then
            Peripheral(Index).Picture = images.buttonup.Picture
            For i = 0 To 4
               If PeriphData(Index).UpAddress(i) <> -1 Then
                  If PeriphData(Index).UpValue(i) <> -1 Then
                     If ExecuteUART And Emulating Then AddToEmuBuff = True
                     SetRAM PeriphData(Index).UpAddress(i), PeriphData(Index).UpValue(i)
                     AddToEmuBuff = False
                  End If
               End If
            Next i
         ElseIf PeriphData(Index).ptype = keypad Then
            If ExecuteUART And Emulating Then AddToEmuBuff = True
            SetRAM PeriphData(Index).Address, 0
            AddToEmuBuff = False
         End If
      End If
   End If
End Sub

Private Sub PeripheralButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Not ExecuteUART Then
      StartX = x
      StartY = y
      PeripheralButton(Index).Drag
   End If
End Sub

Private Sub PeripheralText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Call Peripheral_MouseDown(Index, Button, Shift, 0, 0)
End Sub

Private Sub PicCycles_Click()
   If ShpCycles.tag = "free" Then
      ShpCycles.BorderColor = RGB(0, 255, 0)
      ShpCycles.tag = "locked"
   Else
      ShpCycles.BorderColor = RGB(128, 128, 128)
      ShpCycles.tag = "free"
   End If
End Sub

Private Sub PicCycles_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If ShpCycles.tag = "free" Then
      ShpCycles.X1 = x
      ShpCycles.X2 = x
   End If
End Sub

Private Sub PicPeripherals_DragDrop(Source As Control, x As Single, y As Single)
   Dim standard As Boolean
   standard = False
   
   If Source.tag = "button" Then standard = True
   If Source.tag = "7seg" Then standard = True
   If Source.tag = "ticker" Then standard = True
   If Source.tag = "switch" Then standard = True
   If Source.tag = "keypad" Then standard = True
   If Source.tag = "led" Then standard = True
   If Source.tag = "bar" Then standard = True
   If Source.tag = "dip" Then standard = True
   If Source.tag = "keyboard" Then standard = True
   If Source.tag = "text" Then standard = True
   If Source.tag = "screen" Then standard = True
   If Source.tag = "multiplier" Then standard = True
   If standard = True Then
      CreatePeripheral Source.tag, Int((x - StartX + 8) / 16) * 16, Int((y - StartY + 8) / 16) * 16
   Else
      Source.Move Int((x - StartX + 8) / 16) * 16, Int((y - StartY + 8) / 16) * 16
      RefreshLabels Source.tag
   End If
End Sub

Private Sub ROItm_Click()
   If RamAttribs(RightClickPtr) And AttribReadonly Then
      RamAttribs(RightClickPtr) = RamAttribs(RightClickPtr) - AttribReadonly
      UpdateCell RightClickIndex
   Else
      RamAttribs(RightClickPtr) = RamAttribs(RightClickPtr) Or AttribReadonly
      UpdateCell RightClickIndex
   End If
End Sub

Private Sub SaveMemItm_Click()
   On Error GoTo eh
   CD1.Filter = "Memory Files (*.mem)|*.mem|All Files (*.*)|*.*"
   CD1.ShowSave
   Open CD1.filename For Output As #1
   Close #1
   Open CD1.filename For Binary As #1
      BinWrite "6502 Trainer Memory File"
      Put #1, , RAM
      For i = 0 To 65535
         BinWrite RamTitles(i)
         Put #1, , RamColors(i)
         Put #1, , RamAttribs(i)
         BinWrite RamLabels(i)
         BinWrite RamOps(i)
      Next i
   Close #1
   Exit Sub
eh:
   If Err <> 32755 Then
      MsgBox "There was an error saving the file: " & Error, vbCritical
      Close #1
   End If
End Sub

Private Sub SavePeriphItm_Click()
   On Error GoTo eh
   Dim ptypestr As String
   Dim intbuff As Integer
   Dim LblBuff(2) As String
   Dim TempCount As Integer
   CD1.Filter = "Peripheral Files (*.per)|*.per|All Files (*.*)|*.*"
   CD1.ShowSave
   Open CD1.filename For Output As #1
   Close #1
   Open CD1.filename For Binary As #1
      BinWrite "6502 Trainer Peripheral File"
      TempCount = PeriphCount
      For i = 1 To PeriphCount
         If PeriphData(i).Deleted = True Then TempCount = TempCount - 1
      Next i
      Put #1, , TempCount
      For i = 1 To PeriphCount
         If PeriphData(i).Deleted = False Then
            Select Case PeriphData(i).ptype
            Case PushButton: ptypestr = "button"
            Case SevenSeg: ptypestr = "7seg"
            Case Ticker: ptypestr = "ticker"
            Case SwitchButton: ptypestr = "switch"
            Case keypad: ptypestr = "keypad"
            Case LED: ptypestr = "led"
            Case LED8: ptypestr = "bar"
            Case DipSwitch: ptypestr = "dip"
            Case Keyboard: ptypestr = "keyboard"
            Case TextDisplay: ptypestr = "text"
            Case ScreenDisplay: ptypestr = "screen"
            Case Multiplier: ptypestr = "multiplier"
            End Select
            
            BinWrite ptypestr
            intbuff = Peripheral(i).Left
            Put #1, , intbuff
            intbuff = Peripheral(i).Top
            Put #1, , intbuff
            
            LblBuff(0) = PeriphData(i).Labels(0)
            LblBuff(1) = PeriphData(i).Labels(1)
            LblBuff(2) = PeriphData(i).Labels(2)
            
            PeriphData(i).Labels(0) = ""
            PeriphData(i).Labels(1) = ""
            PeriphData(i).Labels(2) = ""
            
            Put #1, , PeriphData(i)
            BinWrite LblBuff(0)
            BinWrite LblBuff(1)
            BinWrite LblBuff(2)
            
            PeriphData(i).Labels(0) = LblBuff(0)
            PeriphData(i).Labels(1) = LblBuff(1)
            PeriphData(i).Labels(2) = LblBuff(2)
         End If
      Next i
   Close #1
   Exit Sub
eh:
   If Err <> 32755 Then
      MsgBox "There was an error saving the file: " & Error, vbCritical
      Close #1
   End If
End Sub

Private Sub ScrollTable_Change()
   UpdateTable
End Sub

Private Sub ScrollTable_Scroll()
   UpdateTable
End Sub

Private Sub Timer1_Timer()
   Dim k As Single, temp As Single
   Dim update As Boolean
   Dim update2 As Boolean
   Dim value As Integer
   
   For i = 1 To PeriphCount
      If PeriphData(i).Deleted = False Then
         If PeriphData(i).ptype = Ticker Then
            If PeriphData(i).Address <> -1 Then
               If PeriphData(i).TickerInterval <> -1 Then
                  update = False
                  update2 = False
                  k = PeriphData(i).TickerInterval '/ 1000
                  value = RAM(PeriphData(i).Address)
                  value2 = RAM(PeriphData(i).Address + 1)
                  While PeriphData(i).TickerStart + k < GetTickCount 'Timer
                     PeriphData(i).TickerStart = PeriphData(i).TickerStart + k
                     value = value + 1
                     If value > 255 Then
                        value = value - 256
                        If PeriphData(i).Ticker16 Then
                           value2 = value2 + 1
                           If value2 > 255 Then
                              value2 = value2 - 256
                           End If
                           update2 = True
                        End If
                     End If
                     update = True
                  Wend
                  If update Then
                     If ExecuteUART And Emulating Then AddToEmuBuff = True
                     SetRAM PeriphData(i).Address, value
                     AddToEmuBuff = False
                  End If
                  If update2 Then
                     If ExecuteUART And Emulating Then AddToEmuBuff = True
                     SetRAM PeriphData(i).Address + 1, value2
                     AddToEmuBuff = False
                  End If
               End If
            End If
         End If
      End If
   Next i
End Sub

Private Sub TxtChar_GotFocus(Index As Integer)
   IgnoreKey = True
   TxtChar(Index).text = ""
   If Index = RowCount And ScrollTable.value < (16374) Then
      ScrollTable.value = ScrollTable.value + 1
      TxtChar(Index - 4).SetFocus
   End If
End Sub

Private Sub TxtChar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode <> 16 Then
      If Index <> RowCount Then
         TxtChar(Index + 1).SetFocus
      End If
   End If
End Sub

Private Sub TxtChar_LostFocus(Index As Integer)
   IgnoreKey = False
   k = ScrollTable.value
   k = k * 4
   If Len(TxtChar(Index).text) <> 0 Then
      If ExecuteUART And Emulating Then AddToEmuBuff = True
      SetRAM k + Index, Asc(TxtChar(Index).text)
      AddToEmuBuff = False
   Else
      UpdateCell Index
   End If
End Sub

Private Sub TxtData_GotFocus(Index As Integer)
   IgnoreKey = True
   TxtData(Index).SelStart = 0
   TxtData(Index).SelLength = Len(TxtData(Index).text)
   If Index = RowCount And ScrollTable.value < (16374) Then
      ScrollTable.value = ScrollTable.value + 1
      TxtData(Index - 4).SetFocus
   End If
End Sub

Private Sub TxtData_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Index <> RowCount Then
         TxtData(Index + 1).SetFocus
      End If
   End If
End Sub

Private Sub TxtData_LostFocus(Index As Integer)
   IgnoreKey = False
   k = ScrollTable.value
   k = k * 4
   If Len(TxtData(Index).text) = 0 Or UCase(TxtData(Index).text) = "X" Then
      If ExecuteUART And Emulating Then AddToEmuBuff = True
      SetRAM k + Index, -1
      AddToEmuBuff = False
   Else
      value = HexToInt(TxtData(Index).text)
      If value = -1 Or value < 0 Or value > 255 Then
         MsgBox TxtData(Index).text & " is not a valid value. Enter a hexadecimal value between 0 and FF.", vbCritical
         'TxtData(index).SetFocus
         If ExecuteUART And Emulating Then AddToEmuBuff = True
         SetRAM k + Index, RAM(k + Index)
         AddToEmuBuff = False
      Else
         If ExecuteUART And Emulating Then AddToEmuBuff = True
         SetRAM k + Index, value
         AddToEmuBuff = False
      End If
   End If
End Sub

Private Sub UnlockItm_Click()
   For i = 1 To PeriphCount
      If PeriphData(i).Deleted = False Then
         PeriphData(i).Locked = False
      End If
   Next i
End Sub
