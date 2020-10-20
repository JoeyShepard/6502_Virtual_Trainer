VERSION 5.00
Begin VB.Form DlgMultiplier 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multiplier Properties"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox BCDCheck 
      Caption         =   "BCD"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Output 
      Height          =   285
      Left            =   2760
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Input2 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Input1 
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
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Captions 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
   Begin VB.CheckBox LockedCheck 
      Caption         =   "Locked"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   120
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.TextBox Captions 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   5
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Output"
      Height          =   195
      Left            =   2760
      TabIndex        =   12
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Input 2"
      Height          =   195
      Left            =   1920
      TabIndex        =   11
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Input 1"
      Height          =   195
      Left            =   1920
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Captions:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "DlgMultiplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnOK_Click()
   Dim failed As Boolean
   Dim value As Long
   failed = False
   
   If Len(Output.text) = 0 Then
      PeriphData(DlgPtr).Address = -1
   Else
      value = HexToInt(Output.text)
      If value > 65535 Then value = -1
      If value < 0 Then value = -1
      If value = -1 Then
         failed = True
         MsgBox Output.text & " is not a valid value. Enter a hexadecimal value between 0 and FFFF.", vbCritical
      Else
         PeriphData(DlgPtr).Address = value
      End If
   End If
   
   If Len(Input1.text) = 0 Then
      PeriphData(DlgPtr).DownAddress(0) = -1
   Else
      value = HexToInt(Input1.text)
      If value > 65535 Then value = -1
      If value < 0 Then value = -1
      If value = -1 Then
         failed = True
         MsgBox Input1.text & " is not a valid value. Enter a hexadecimal value between 0 and FFFF.", vbCritical
      Else
         PeriphData(DlgPtr).DownAddress(0) = value
      End If
   End If
   
   If Len(Input2.text) = 0 Then
      PeriphData(DlgPtr).UpAddress(0) = -1
   Else
      value = HexToInt(Input2.text)
      If value > 65535 Then value = -1
      If value < 0 Then value = -1
      If value = -1 Then
         failed = True
         MsgBox Input2.text & " is not a valid value. Enter a hexadecimal value between 0 and FFFF.", vbCritical
      Else
         PeriphData(DlgPtr).UpAddress(0) = value
      End If
   End If
   
   If Not failed Then
      If LockedCheck.value = 1 Then
         PeriphData(DlgPtr).Locked = True
      Else
         PeriphData(DlgPtr).Locked = False
      End If
      
      If BCDCheck.value = 1 Then
         PeriphData(DlgPtr).UseBCD = True
      Else
         PeriphData(DlgPtr).UseBCD = False
      End If
      
      PeriphData(DlgPtr).Labels(0) = Captions(0).text
      PeriphData(DlgPtr).Labels(1) = Captions(1).text
      PeriphData(DlgPtr).Labels(2) = Captions(2).text
      RefreshLabels DlgPtr
      
      UpdateMultiplier DlgPtr
      
      Unload DlgMultiplier
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload DlgTimer
End Sub

