Attribute VB_Name = "Module1"
Option Explicit

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type DRAWITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemAction As Long
        itemState As Long
        hwndItem As Long
        hdc As Long
        rcItem As RECT
        itemData As Long
End Type

Public Type CWPSTRUCT
        lParam As Long
        wParam As Long
        message As Long
        hwnd As Long
End Type

Public Type CREATESTRUCT
        lpCreateParams As Long
        hInstance As Long
        hMenu As Long
        hWndParent As Long
        cy As Long
        cx As Long
        y As Long
        x As Long
        style As Long
        'These next 2 are Normaly String, but need to be a fixed length
        'so we know how long they are when using CopyMemory,
        'We're only interested in the Style Property anyway.
        lpszName As Long
        lpszClass As Long
        ExStyle As Long
End Type

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Const WH_CALLWNDPROC = 4

Public Const CBS_OWNERDRAWVARIABLE = &H20&
Public Const CB_GETLBTEXT = &H148
Public Const CB_SETITEMHEIGHT = &H153

Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWTEXT = 8

Public Const GWL_WNDPROC = (-4)
Public Const GWL_STYLE = (-16)

Public Const ODS_SELECTED = &H1

Public Const ODT_COMBOBOX = 3

Public Const WM_CREATE = &H1
Public Const WM_DRAWITEM = &H2B

Public lPrevWndProc As Long
Public lHook As Long
Public lSubCombo As Long

Public Function SubClassedForm(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim tItem As DRAWITEMSTRUCT
    Dim sItem As String
    Dim lBack As Long
    
    If Msg = WM_DRAWITEM Then
        'This function only passes the Address of the DrawItem Structure, so we need to
        'use the CopyMemory API to Get a Copy into the Variable we setup:
        Call CopyMemory(tItem, ByVal lParam, Len(tItem))
        'If it's our Combobox..
        If tItem.CtlType = ODT_COMBOBOX Then
            'Get the Item Text
            sItem = Space(255)
            Call SendMessage(tItem.hwndItem, CB_GETLBTEXT, tItem.itemID, ByVal sItem)
            sItem = Left(sItem, InStr(sItem, Chr(0)) - 1)
            'Select the Highlight Colors if this Item is currently selected
            If (tItem.itemState And ODS_SELECTED) Then
                lBack = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT))
                Call FillRect(tItem.hdc, tItem.rcItem, lBack)
                Call SetBkColor(tItem.hdc, GetSysColor(COLOR_HIGHLIGHT))
                Call SetTextColor(tItem.hdc, GetSysColor(COLOR_HIGHLIGHTTEXT))
            Else
                'Otherwise, use the default Colors, unless the Item Data contains a new Color
                lBack = CreateSolidBrush(GetSysColor(COLOR_WINDOW))
                Call FillRect(tItem.hdc, tItem.rcItem, lBack)
                Call SetBkColor(tItem.hdc, GetSysColor(COLOR_WINDOW))
                Call SetTextColor(tItem.hdc, IIf(tItem.itemData, tItem.itemData, GetSysColor(COLOR_WINDOWTEXT)))
            End If
            Call DeleteObject(lBack)
            'Display the Item
            TextOut tItem.hdc, tItem.rcItem.Left, tItem.rcItem.Top, ByVal sItem, Len(sItem)
            'Don't Return a Value as we've dealt with this Message ourselves
            SubClassedForm = 0
            Exit Function
        End If
    End If
    'Not our Combobox, so just process the Message as Normal
    SubClassedForm = CallWindowProc(lPrevWndProc, hwnd, Msg, wParam, lParam)
End Function

Sub Main()
    'The Combobox is a little more tricky to manipulate than a Listbox
    'So we need to do a little extra work to make it an "Owner Drawn" Control.
    lHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf HookApp, App.hInstance, App.ThreadID)
    Form1.Show
    'Once the Control. etc are Drawn, we can release the Hook
    Call UnhookWindowsHookEx(lHook)
End Sub

Private Function HookApp(ByVal lHookID As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    'This Function will get called when Initializing the Form
    'We want to Interupt it when it tries to create our Combobox..
    Dim tCWP As CWPSTRUCT
    Dim sClass As String

    Call CopyMemory(tCWP, ByVal lParam, Len(tCWP))

    If tCWP.message = WM_CREATE Then
        'Get the Control Classname
        sClass = Space(128)
        Call GetClassName(tCWP.hwnd, ByVal sClass, 128)
        sClass = Left(sClass, InStr(sClass, Chr(0)) - 1)
        'If it's our Combobox, Sub-class it to Modify the Create Message..
        If sClass = "ComboLBox" Then
            lSubCombo = SetWindowLong(tCWP.hwnd, GWL_WNDPROC, AddressOf SubComboCreate)
        End If
    End If
    'Continue the Hook Processing
    HookApp = CallNextHookEx(lHook, lHookID, wParam, ByVal lParam)

End Function

Private Function SubComboCreate(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'This Function will be called when the Combobox is about to be created
    Dim tCreate As CREATESTRUCT
    
    If Msg = WM_CREATE Then
        'Grab the Data that's going to be used to Create the Combobox
        Call CopyMemory(tCreate, ByVal lParam, Len(tCreate))
        'Alter it, to make the Combobox an "Owner Drawn" Control
        tCreate.style = tCreate.style Or CBS_OWNERDRAWVARIABLE
        'Copy the modified data back
        Call CopyMemory(ByVal lParam, tCreate, Len(tCreate))
        'Alter the Style to OwnerDrawn
        Call SetWindowLong(hwnd, GWL_STYLE, tCreate.style)
        'Release this Subclassing Function
        Call SetWindowLong(hwnd, GWL_WNDPROC, lSubCombo)
    End If
    'Let Windows Process the Modified Data
    SubComboCreate = CallWindowProc(lSubCombo, hwnd, Msg, wParam, lParam)
    
End Function
