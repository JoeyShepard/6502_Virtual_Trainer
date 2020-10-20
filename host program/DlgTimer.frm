VERSION 5.00
Begin VB.Form DlgTimer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Timer Properties"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check16bit 
      Caption         =   "16-bit"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Interval 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Address 
      Height          =   285
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Captions 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.CheckBox LockedCheck 
      Caption         =   "Locked"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   120
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.TextBox Captions 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Captions 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Interval"
      Height          =   195
      Left            =   1920
      TabIndex        =   10
      Top             =   840
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Address"
      Height          =   195
      Left            =   1920
      TabIndex        =   7
      Top             =   120
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Captions:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "DlgTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnOK_Click()
   Dim failed As Boolean
   Dim value As Long
   failed = False
   
   If Len(Address.text) = 0 Then
      PeriphData(DlgPtr).Address = -1
   Else
      value = HexToInt(Address.text)
      If value > 65535 Then value = -1
      If value < 0 Then value = -1
      If value = -1 Then
         failed = True
         MsgBox Address.text & " is not a valid value. Enter a hexadecimal value between 0 and FFFF.", vbCritical
      Else
         PeriphData(DlgPtr).Address = value
      End If
   End If
   
   If Len(Interval.text) = 0 Then
      PeriphData(DlgPtr).TickerInterval = -1
   Else
      value = Val(Interval.text)
      If value = 0 Then value = -1
      If value = -1 Then
         failed = True
         MsgBox Interval.text & " is not a valid value. Enter a decimal value between 0 and 4,000,000,000.", vbCritical
      Else
         PeriphData(DlgPtr).TickerInterval = value
      End If
   End If
   
   If Not failed Then
      If LockedCheck.value = 1 Then
         PeriphData(DlgPtr).Locked = True
      Else
         PeriphData(DlgPtr).Locked = False
      End If
      
      If Check16bit.value = 1 Then
         PeriphData(DlgPtr).Ticker16 = True
      Else
         PeriphData(DlgPtr).Ticker16 = False
      End If
      
      PeriphData(DlgPtr).Labels(0) = Captions(0).text
      PeriphData(DlgPtr).Labels(1) = Captions(1).text
      PeriphData(DlgPtr).Labels(2) = Captions(2).text
      RefreshLabels DlgPtr
      
      frmMain.Timer1.Enabled = False
      If PeriphData(DlgPtr).Address <> -1 Then
         If PeriphData(DlgPtr).TickerInterval <> -1 Then
            PeriphData(DlgPtr).TickerStart = GetTickCount
            frmMain.Timer1.Enabled = True
         End If
      End If
      
      Unload DlgTimer
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload DlgTimer
End Sub

