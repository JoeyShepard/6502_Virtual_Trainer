VERSION 5.00
Begin VB.Form DlgLED 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LED Properties"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Compare 
      Height          =   315
      Index           =   5
      ItemData        =   "DlgLED.frx":0000
      Left            =   120
      List            =   "DlgLED.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   2160
      Width           =   735
   End
   Begin VB.ComboBox Compare 
      Height          =   315
      Index           =   4
      ItemData        =   "DlgLED.frx":004C
      Left            =   120
      List            =   "DlgLED.frx":006E
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1800
      Width           =   735
   End
   Begin VB.ComboBox Compare 
      Height          =   315
      Index           =   3
      ItemData        =   "DlgLED.frx":0098
      Left            =   120
      List            =   "DlgLED.frx":00BA
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1440
      Width           =   735
   End
   Begin VB.ComboBox Compare 
      Height          =   315
      Index           =   2
      ItemData        =   "DlgLED.frx":00E4
      Left            =   120
      List            =   "DlgLED.frx":0106
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1080
      Width           =   735
   End
   Begin VB.ComboBox Compare 
      Height          =   315
      Index           =   1
      ItemData        =   "DlgLED.frx":0130
      Left            =   120
      List            =   "DlgLED.frx":0152
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Limit 
      Height          =   285
      Index           =   5
      Left            =   960
      MaxLength       =   2
      TabIndex        =   15
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox Limit 
      Height          =   285
      Index           =   4
      Left            =   960
      MaxLength       =   2
      TabIndex        =   13
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox Limit 
      Height          =   285
      Index           =   3
      Left            =   960
      MaxLength       =   2
      TabIndex        =   12
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Limit 
      Height          =   285
      Index           =   2
      Left            =   960
      MaxLength       =   2
      TabIndex        =   11
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Limit 
      Height          =   285
      Index           =   1
      Left            =   960
      MaxLength       =   2
      TabIndex        =   10
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Limit 
      Height          =   285
      Index           =   0
      Left            =   960
      MaxLength       =   2
      TabIndex        =   9
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Address 
      Height          =   285
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   7
      Top             =   2160
      Width           =   615
   End
   Begin VB.ComboBox Compare 
      Height          =   315
      Index           =   0
      ItemData        =   "DlgLED.frx":017C
      Left            =   120
      List            =   "DlgLED.frx":019E
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Captions 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
   Begin VB.CheckBox LockedCheck 
      Caption         =   "Locked"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   120
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.TextBox Captions 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Captions 
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Value"
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Address"
      Height          =   195
      Left            =   1680
      TabIndex        =   8
      Top             =   1920
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Captions:"
      Height          =   195
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   660
   End
End
Attribute VB_Name = "DlgLED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnOK_Click()
   Dim failed As Boolean
   Dim value As Long
   failed = False
   
   If Len(Address.Text) = 0 Then
      PeriphData(DlgPtr).Address = -1
   Else
      value = HexToInt(Address.Text)
      If value > 65535 Then value = -1
      If value < 0 Then value = -1
      If value = -1 Then
         failed = True
         MsgBox Address.Text & " is not a valid value. Enter a hexadecimal value between 0 and FFFF.", vbCritical
      Else
         PeriphData(DlgPtr).Address = value
      End If
   End If
   
   For i = 0 To 5
      If Len(Limit(i).Text) = 0 Then
         PeriphData(DlgPtr).LEDvalue(i) = -1
      Else
         value = HexToInt(Limit(i).Text)
         If value > 255 Then value = -1
         If value < 0 Then value = -1
         If value = -1 Then
            failed = True
            MsgBox Limit(i).Text & " is not a valid value. Enter a hexadecimal value between 0 and FF.", vbCritical
         Else
            PeriphData(DlgPtr).LEDvalue(i) = value
         End If
      End If
      PeriphData(DlgPtr).LEDrelation(i) = Compare(i).ListIndex
   Next i
   
   If Not failed Then
      If LockedCheck.value = 1 Then
         PeriphData(DlgPtr).Locked = True
      Else
         PeriphData(DlgPtr).Locked = False
      End If
      
      PeriphData(DlgPtr).Labels(0) = Captions(0).Text
      PeriphData(DlgPtr).Labels(1) = Captions(1).Text
      PeriphData(DlgPtr).Labels(2) = Captions(2).Text
      RefreshLabels DlgPtr
      
      UpdateLED DlgPtr
      
      Unload DlgLED
   End If
End Sub

Private Sub Form_Load()
   For i = 0 To 5
      Compare(i).ListIndex = 0
   Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload DlgLED
End Sub

