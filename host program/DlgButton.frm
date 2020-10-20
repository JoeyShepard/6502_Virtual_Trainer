VERSION 5.00
Begin VB.Form DlgButton 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Button Properties"
   ClientHeight    =   2535
   ClientLeft      =   7485
   ClientTop       =   7290
   ClientWidth     =   5415
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Up"
      Height          =   2295
      Left            =   1920
      TabIndex        =   24
      Top             =   120
      Width           =   1575
      Begin VB.TextBox UpAddress 
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   4
         TabIndex        =   14
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox UpAddress 
         Height          =   285
         Index           =   1
         Left            =   120
         MaxLength       =   4
         TabIndex        =   16
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox UpAddress 
         Height          =   285
         Index           =   2
         Left            =   120
         MaxLength       =   4
         TabIndex        =   18
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox UpAddress 
         Height          =   285
         Index           =   3
         Left            =   120
         MaxLength       =   4
         TabIndex        =   20
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox UpAddress 
         Height          =   285
         Index           =   4
         Left            =   120
         MaxLength       =   4
         TabIndex        =   22
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox UpValue 
         Height          =   285
         Index           =   0
         Left            =   960
         MaxLength       =   2
         TabIndex        =   15
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox UpValue 
         Height          =   285
         Index           =   1
         Left            =   960
         MaxLength       =   2
         TabIndex        =   17
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox UpValue 
         Height          =   285
         Index           =   2
         Left            =   960
         MaxLength       =   2
         TabIndex        =   19
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox UpValue 
         Height          =   285
         Index           =   3
         Left            =   960
         MaxLength       =   2
         TabIndex        =   21
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox UpValue 
         Height          =   285
         Index           =   4
         Left            =   960
         MaxLength       =   2
         TabIndex        =   23
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label4 
         Caption         =   "Value"
         Height          =   255
         Left            =   960
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Down"
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
      Begin VB.TextBox DownValue 
         Height          =   285
         Index           =   4
         Left            =   960
         MaxLength       =   2
         TabIndex        =   13
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox DownValue 
         Height          =   285
         Index           =   3
         Left            =   960
         MaxLength       =   2
         TabIndex        =   11
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox DownValue 
         Height          =   285
         Index           =   2
         Left            =   960
         MaxLength       =   2
         TabIndex        =   9
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox DownValue 
         Height          =   285
         Index           =   1
         Left            =   960
         MaxLength       =   2
         TabIndex        =   7
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox DownValue 
         Height          =   285
         Index           =   0
         Left            =   960
         MaxLength       =   2
         TabIndex        =   5
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox DownAddress 
         Height          =   285
         Index           =   4
         Left            =   120
         MaxLength       =   4
         TabIndex        =   12
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox DownAddress 
         Height          =   285
         Index           =   3
         Left            =   120
         MaxLength       =   4
         TabIndex        =   10
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox DownAddress 
         Height          =   285
         Index           =   2
         Left            =   120
         MaxLength       =   4
         TabIndex        =   8
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox DownAddress 
         Height          =   285
         Index           =   1
         Left            =   120
         MaxLength       =   4
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox DownAddress 
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   4
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Value"
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   570
      End
   End
   Begin VB.TextBox Captions 
      Height          =   285
      Index           =   2
      Left            =   3720
      TabIndex        =   29
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Captions 
      Height          =   285
      Index           =   1
      Left            =   3720
      TabIndex        =   27
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CheckBox LockedCheck 
      Caption         =   "Locked"
      Height          =   255
      Left            =   3720
      TabIndex        =   31
      Top             =   120
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox Captions 
      Height          =   285
      Index           =   0
      Left            =   3720
      TabIndex        =   25
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   30
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Captions:"
      Height          =   195
      Left            =   3720
      TabIndex        =   0
      Top             =   600
      Width           =   660
   End
End
Attribute VB_Name = "DlgButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnOK_Click()
   Dim failed As Boolean
   Dim value As Long
   failed = False
   
   For i = 0 To 4
      If Len(DownAddress(i).text) = 0 Then
         PeriphData(DlgPtr).DownAddress(i) = -1
      Else
         value = HexToInt(DownAddress(i).text)
         If value > 65535 Then value = -1
         If value < 0 Then value = -1
         If value = -1 Then
            failed = True
            MsgBox DownAddress(i).text & " is not a valid value. Enter a hexadecimal value between 0 and FFFF.", vbCritical
         Else
            PeriphData(DlgPtr).DownAddress(i) = value
         End If
      End If
      
      If Len(DownValue(i).text) = 0 Then
         PeriphData(DlgPtr).DownValue(i) = -1
      Else
         value = HexToInt(DownValue(i).text)
         If value > 255 Then value = -1
         If value < 0 Then value = -1
         If value = -1 Then
            failed = True
            MsgBox DownValue(i).text & " is not a valid value. Enter a hexadecimal value between 0 and FF.", vbCritical
         Else
            PeriphData(DlgPtr).DownValue(i) = value
         End If
      End If
      
      If Len(UpAddress(i).text) = 0 Then
         PeriphData(DlgPtr).UpAddress(i) = -1
      Else
         value = HexToInt(UpAddress(i).text)
         If value > 65535 Then value = -1
         If value < 0 Then value = -1
         If value = -1 Then
            failed = True
            MsgBox UpAddress(i).text & " is not a valid value. Enter a hexadecimal value between 0 and FFFF.", vbCritical
         Else
            PeriphData(DlgPtr).UpAddress(i) = value
         End If
      End If
      
      If Len(UpValue(i).text) = 0 Then
         PeriphData(DlgPtr).UpValue(i) = -1
      Else
         value = HexToInt(UpValue(i).text)
         If value > 255 Then value = -1
         If value < 0 Then value = -1
         If value = -1 Then
            failed = True
            MsgBox UpValue(i).text & " is not a valid value. Enter a hexadecimal value between 0 and FF.", vbCritical
         Else
            PeriphData(DlgPtr).UpValue(i) = value
         End If
      End If
   Next i
   
   If Not failed Then
      If LockedCheck.value = 1 Then
         PeriphData(DlgPtr).Locked = True
      Else
         PeriphData(DlgPtr).Locked = False
      End If
      
      PeriphData(DlgPtr).Labels(0) = Captions(0).text
      PeriphData(DlgPtr).Labels(1) = Captions(1).text
      PeriphData(DlgPtr).Labels(2) = Captions(2).text
      
      RefreshLabels DlgPtr
      
      Unload DlgButton
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload DlgButton
End Sub

