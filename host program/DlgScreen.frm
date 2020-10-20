VERSION 5.00
Begin VB.Form DlgScreen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Screen Properties"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Resolution 
      Height          =   315
      ItemData        =   "DlgScreen.frx":0000
      Left            =   120
      List            =   "DlgScreen.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox WidthValue 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox HeightValue 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Address 
      Height          =   285
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Captions 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   1575
   End
   Begin VB.CheckBox LockedCheck 
      Caption         =   "Locked"
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   120
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.TextBox Captions 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Captions 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Height"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Width"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Address"
      Height          =   195
      Left            =   2520
      TabIndex        =   2
      Top             =   600
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Captions:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "DlgScreen"
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
   
   If Val(WidthValue.text) < 1 Or Val(WidthValue.text) > 500 Then
      MsgBox WidthValue.text & " is not a valid value. Enter a decimal value between 1 and 500.", vbCritical
      failed = True
   Else
      If Val(HeightValue.text) < 1 Or Val(HeightValue.text) > 500 Then
         MsgBox HeightValue.text & " is not a valid value. Enter a decimal value between 1 and 500.", vbCritical
         failed = True
      Else
         PeriphData(DlgPtr).ScreenWidth = Val(WidthValue.text)
         PeriphData(DlgPtr).ScreenHeight = Val(HeightValue.text)
      End If
   End If
   
   If Not failed Then
      If LockedCheck.value = 1 Then
         PeriphData(DlgPtr).Locked = True
      Else
         PeriphData(DlgPtr).Locked = False
      End If
      
      PeriphData(DlgPtr).ScreenRes = Resolution.ListIndex
      
      PeriphData(DlgPtr).Labels(0) = Captions(0).text
      PeriphData(DlgPtr).Labels(1) = Captions(1).text
      PeriphData(DlgPtr).Labels(2) = Captions(2).text
      RefreshLabels DlgPtr
      
      ResizeScreen DlgPtr
      
      Unload DlgScreen
   End If
End Sub

Private Sub Form_Load()
   Resolution.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload DlgScreen
End Sub

