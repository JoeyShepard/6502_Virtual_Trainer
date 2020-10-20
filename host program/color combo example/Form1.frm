VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim iColor As Integer
    Combo1.Width = 975
    Combo1.AddItem "Black"
    Combo1.AddItem "Navy"
    Combo1.AddItem "Green"
    Combo1.AddItem "Teal"
    Combo1.AddItem "Maroon"
    Combo1.AddItem "Burgandy"
    Combo1.AddItem "Olive"
    Combo1.AddItem "Silver"
    Combo1.AddItem "Grey"
    Combo1.AddItem "Blue"
    Combo1.AddItem "Green"
    Combo1.AddItem "Cyan"
    Combo1.AddItem "Red"
    Combo1.AddItem "Purple"
    Combo1.AddItem "Yellow"
    Combo1.AddItem "White"
    For iColor = 0 To 15
        Combo1.itemData(iColor) = QBColor(iColor)
    Next
    Combo1.ListIndex = 0
    lPrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SubClassedForm)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SetWindowLong(hwnd, GWL_WNDPROC, lPrevWndProc)
End Sub


