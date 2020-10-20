Attribute VB_Name = "Module1"
Declare Function SetCommState Lib "kernel32" (ByVal hCommDev As Long, lpDCB As DCB) As Long
Declare Function GetCommState Lib "kernel32" (ByVal nCid As Long, lpDCB As DCB) As Long
Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function SetPixel Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Declare Function SetPixelV Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Byte
Declare Function GetTickCount Lib "kernel32" () As Long

Public Const SRCPAINT = &HEE0086
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046 'Good for cloaked ships
Public Const SRCCOPY = &HCC0020

Public Const RowCount = 40
Public Const PACKET_SIZE = 20
Public Const CompactMode = True

Private Type DCB
   DCBlength As Long
   BaudRate As Long
   Bits1 As Long
   wReserved As Integer
   XonLim As Integer
   XoffLim As Integer
   ByteSize As Byte
   Parity As Byte
   StopBits As Byte
   XonChar As Byte
   XoffChar As Byte
   ErrorChar As Byte
   EofChar As Byte
   EvtChar As Byte
   wReserved2 As Integer
End Type

Enum PeripheralTypes
   PushButton
   SwitchButton
   SevenSeg
   Ticker
   keypad
   LED
   LED8
   DipSwitch
   Keyboard
   TextDisplay
   ScreenDisplay
   Multiplier
End Enum

Enum RamAttributes
   AttribReadonly = 1
   AttribCode = 2
   AttribBreakpoint = 4
   AttribUninitialized = 128 'only used when communicating with SRAM
End Enum

Enum MessageTypes
   PING = 10
   PONG = 15
   ERROR_ = 20
   UNKNOWN = 25
   RESET_CPU = 30
   RESET_ACK = 35
   DOWN_CYCLE = 40
   UP_CYCLE = 45
   DOWN_ACK = 50
   UP_ACK_READ = 55
   UP_ACK_WRITE = 60
   UPDATE_RAM = 65
   UPDATE_RAM_ACK = 70
   UPDATE_RAM_CRC = 75
   GET_RAM_CRC = 80
   SEND_RAM_CRC = 85
   BEGIN_EMULATING = 90
   UPDATE_DIRTY_RAM = 95
   STOP_EMULATING = 100
   KEEP_EMULATING = 105
   COMM_CHECK = 110
   COMM_CHECK_ACK = 115
   CUSTOM_CHECK = 120
   CUSTOM_CHECK_ACK = 125
End Enum

Enum InputPinTypes
   CPU_VPB = 2
   CPU_MLB = 4
   CPU_VPA = 8
   CPU_VDA = 16
   CPU_MX = 32
   CPU_E = 64
   CPU_RWB = 128
End Enum

Type PeripheralRecord
   ptype As Integer
   UpAddress(4) As Long
   DownAddress(4) As Long
   UpValue(4) As Long
   DownValue(4) As Long
   Labels(2) As String
   Locked As Boolean
   Address As Long
   Ticker16 As Boolean
   TickerStart As Long
   TickerInterval As Long
   switchon As Boolean
   LEDvalue(5) As Long
   LEDrelation(5) As Long
   DipValue As Integer
   TextWidth As Integer
   TextHeight As Integer
   ScreenWidth As Integer
   ScreenHeight As Integer
   ScreenRes As Integer
   Deleted As Boolean
   UseBCD As Boolean
End Type

Public RamColors(65535) As Long
Public RamTitles(65535) As String
Public RamStarts(65535) As Long
Public RamEnds(65535) As Long
Public RamAttribs(65535) As Integer
Public RamLabels(65535) As String
Public RamOps(65535) As String
Public RAM(65535) As Integer
Public PeriphCount As Integer
Public StartX As Long, StartY As Long
Public PeriphData() As PeripheralRecord
Public DlgPtr As Integer
Public IgnoreKey As Boolean
Public SectionList() As Long
Public ExecuteUART As Boolean
Public PeriphToDelete As Integer
Public StopUART As Boolean
Public HighlightPtr As Long
Public WaitingForInput As Boolean
Public WaitingSingleCycle As Boolean
Public ToDisableUART As Boolean
Public JustCycling As Boolean
Public CycleCount As Long
Public GlobalTime As Long
Public Emulating As Boolean
Public AddToEmuBuff As Boolean
Public EmuData(63) As Integer
Public EmuAddress(63) As Long
Public EmuCount As Integer
Public ToStopEmulating As Boolean
Public RightClickPtr As Long
Public RightClickIndex As Long
Public ToReset As Boolean
Public ToLoad As Boolean
'This limits the max number of textboxes!
Public TextsToUpdate(10) As Integer

Public Sub UpdateTable()
   Dim i As Long, j As Long, k As Long
   j = -1
   For i = 0 To RowCount
      UpdateCell i
      k = frmMain.ScrollTable.value
      k = k * 4 + i
      If k < 65536 Then
         If j = -1 Then
            If RamTitles(k) <> "" Then j = k
         End If
      End If
   Next i
   
   If j <> -1 Then
      frmMain.LblHeading.Caption = RamTitles(j) & "(" & HexBig(RamStarts(j))
      If RamStarts(j) <> RamEnds(j) Then
         frmMain.LblHeading.Caption = frmMain.LblHeading.Caption & "-" & HexBig(RamEnds(j)) & ")"
      Else
         frmMain.LblHeading.Caption = frmMain.LblHeading.Caption & ")"
      End If
      frmMain.LblHeading.BackColor = RamColors(j)
      frmMain.LblHeading.Visible = True
   Else
      frmMain.LblHeading.Visible = False
   End If
   MoveHighlightPtr
End Sub

Sub UpdateCell(ByVal i As Integer)
   Dim k As Long
   Dim j As String
   k = frmMain.ScrollTable.value
   k = k * 4 + i
   If k > 65535 Then
      frmMain.TxtLbl(i).text = ""
      frmMain.TxtAddress(i).text = ""
      frmMain.TxtData(i).text = ""
      frmMain.TxtChar(i).text = ""
      frmMain.TxtDis(i).text = ""
      
      frmMain.TxtLbl(i).BackColor = vbWhite
      frmMain.TxtAddress(i).BackColor = vbWhite
      frmMain.TxtData(i).BackColor = vbWhite
      frmMain.TxtChar(i).BackColor = vbWhite
      frmMain.TxtDis(i).BackColor = vbWhite
      
      frmMain.TxtLbl(i).ToolTipText = ""
      frmMain.TxtAddress(i).ToolTipText = ""
      frmMain.TxtData(i).ToolTipText = ""
      frmMain.TxtChar(i).ToolTipText = ""
      frmMain.TxtDis(i).ToolTipText = ""
   Else
      frmMain.TxtAddress(i).text = HexBig(k)
      If RAM(k) = -1 Then
         frmMain.TxtData(i).ForeColor = vbRed
         frmMain.TxtData(i).text = "X"
         frmMain.TxtChar(i).text = ""
      Else
         frmMain.TxtData(i).ForeColor = vbBlack
         j = Hex(RAM(k))
         frmMain.TxtData(i).text = String(2 - Len(j), "0") + j
         frmMain.TxtChar(i).text = Chr(RAM(k))
      End If
      
      frmMain.TxtLbl(i).text = RamLabels(k)
      frmMain.TxtDis(i).text = RamOps(k)
      
      frmMain.TxtLbl(i).BackColor = RamColors(k)
      frmMain.TxtAddress(i).BackColor = RamColors(k)
      frmMain.TxtData(i).BackColor = RamColors(k)
      frmMain.TxtChar(i).BackColor = RamColors(k)
      frmMain.TxtDis(i).BackColor = RamColors(k)
      
      If RamTitles(k) <> "" Then
         frmMain.TxtLbl(i).ToolTipText = RamTitles(k)
         frmMain.TxtAddress(i).ToolTipText = RamTitles(k)
         frmMain.TxtData(i).ToolTipText = RamTitles(k)
         frmMain.TxtChar(i).ToolTipText = RamTitles(k)
         frmMain.TxtDis(i).ToolTipText = RamTitles(k)
      Else
         frmMain.TxtLbl(i).ToolTipText = ""
         frmMain.TxtAddress(i).ToolTipText = ""
         frmMain.TxtData(i).ToolTipText = ""
         frmMain.TxtChar(i).ToolTipText = ""
         frmMain.TxtDis(i).ToolTipText = ""
      End If
   
      frmMain.BreakPoint(i).Cls
      If RamAttribs(k) And AttribReadonly Then
         BitBlt frmMain.BreakPoint(i).hDC, 0, 0, 13, 13, images.ReadOnly.hDC, 0, 0, SRCCOPY
      End If
      
      If RamAttribs(k) And AttribCode Then
         BitBlt frmMain.BreakPoint(i).hDC, 0, 0, 13, 13, images.code_mask.hDC, 0, 0, SRCAND
         BitBlt frmMain.BreakPoint(i).hDC, 0, 0, 13, 13, images.code.hDC, 0, 0, SRCPAINT
      End If
      
      If RamAttribs(k) And AttribBreakpoint Then
         BitBlt frmMain.BreakPoint(i).hDC, 0, 0, 13, 13, images.breakpoint_mask.hDC, 0, 0, SRCAND
         BitBlt frmMain.BreakPoint(i).hDC, 0, 0, 13, 13, images.BreakPoint.hDC, 0, 0, SRCPAINT
      End If
   End If
End Sub

Function HexBig(ByVal number As Long)
   j = Hex(number \ 256)
   final = String(2 - Len(j), "0") + j
   j = Hex(number Mod 256)
   HexBig = final + String(2 - Len(j), "0") + j
End Function

Function HexSmall(ByVal number As Long)
   j = Hex(number Mod 256)
   HexSmall = String(2 - Len(j), "0") + j
End Function

Function HexToInt(ByVal number As String) As Long
   Dim failed As Boolean
   Dim temp As String
   Dim value As Long
   value = -1
   failed = False
   For j = 1 To Len(number)
      temp = UCase(Mid(number, j, 1))
      If Asc(temp) >= Asc("0") And Asc(temp) <= Asc("9") Then
         If value < 0 Then value = 0
         value = value * 16
         value = value + Asc(temp) - Asc("0")
      ElseIf Asc(temp) >= Asc("A") And Asc(temp) <= Asc("F") Then
         If value < 0 Then value = 0
         value = value * 16
         value = value + Asc(temp) - Asc("A") + 10
      Else
         failed = True
      End If
   Next j
   If failed = True Then
      HexToInt = -1
   Else
      HexToInt = value
   End If
End Function

Sub RefreshLabels(Index As Integer)
   If PeriphData(Index).Deleted = False Then
      For i = 0 To 2
         If PeriphData(Index).Labels(i) = "" Then
            If i = 0 Then frmMain.PeriphLabel1(Index).Visible = False
            If i = 1 Then frmMain.PeriphLabel2(Index).Visible = False
            If i = 2 Then frmMain.PeriphLabel3(Index).Visible = False
         Else
            If i = 0 Then
               frmMain.PeriphLabel1(Index).Visible = True
               frmMain.PeriphLabel1(Index).Caption = PeriphData(Index).Labels(0)
               frmMain.PeriphLabel1(Index).Move frmMain.Peripheral(Index).Left + frmMain.Peripheral(Index).Width / 2 - frmMain.PeriphLabel1(Index).Width / 2, frmMain.Peripheral(Index).Top + frmMain.Peripheral(Index).Height
            End If
            
            If i = 1 Then
               frmMain.PeriphLabel2(Index).Visible = True
               frmMain.PeriphLabel2(Index).Caption = PeriphData(Index).Labels(1)
               frmMain.PeriphLabel2(Index).Move frmMain.Peripheral(Index).Left + frmMain.Peripheral(Index).Width / 2 - frmMain.PeriphLabel2(Index).Width / 2, frmMain.Peripheral(Index).Top + frmMain.Peripheral(Index).Height + frmMain.PeriphLabel1(Index).Height
            End If
            
            If i = 2 Then
               frmMain.PeriphLabel3(Index).Visible = True
               frmMain.PeriphLabel3(Index).Caption = PeriphData(Index).Labels(2)
               frmMain.PeriphLabel3(Index).Move frmMain.Peripheral(Index).Left + frmMain.Peripheral(Index).Width / 2 - frmMain.PeriphLabel3(Index).Width / 2, frmMain.Peripheral(Index).Top + frmMain.Peripheral(Index).Height + frmMain.PeriphLabel1(Index).Height + frmMain.PeriphLabel2(Index).Height
            End If
         End If
      Next i
   End If
End Sub

Sub Update7Seg(ByVal Index As Integer)
   frmMain.Peripheral(Index).Picture = images.SevenSegNone.Picture
   If PeriphData(Index).Address <> -1 Then
      If RAM(PeriphData(Index).Address) <> -1 Then
         i = RAM(PeriphData(Index).Address) \ 16
         j = RAM(PeriphData(Index).Address) Mod 16
         BitBlt frmMain.Peripheral(Index).hDC, 2, 3, 13, 26, images.SevenSegAll.hDC, i * 14, 0, SRCCOPY
         BitBlt frmMain.Peripheral(Index).hDC, 17, 3, 13, 26, images.SevenSegAll.hDC, j * 14, 0, SRCCOPY
      End If
   End If
End Sub
Sub UpdateLED(ByVal Index As Integer)
   Dim paint As Boolean
   paint = False
   frmMain.Peripheral(Index).Picture = images.LEDoff.Picture
   If PeriphData(Index).Address <> -1 Then
      If RAM(PeriphData(Index).Address) <> -1 Then
         For i = 0 To 5
            Select Case PeriphData(Index).LEDrelation(i)
               Case 1 'GT
                  If RAM(PeriphData(Index).Address) > PeriphData(Index).LEDvalue(i) Then paint = True
               Case 2 'GTE
                  If RAM(PeriphData(Index).Address) >= PeriphData(Index).LEDvalue(i) Then paint = True
               Case 3 'EQU
                  If RAM(PeriphData(Index).Address) = PeriphData(Index).LEDvalue(i) Then paint = True
               Case 4 'LTE
                  If RAM(PeriphData(Index).Address) <= PeriphData(Index).LEDvalue(i) Then paint = True
               Case 5 'LT
                  If RAM(PeriphData(Index).Address) < PeriphData(Index).LEDvalue(i) Then paint = True
               Case 6 'NE
                  If RAM(PeriphData(Index).Address) <> PeriphData(Index).LEDvalue(i) Then paint = True
               Case 7 'AND
                  If RAM(PeriphData(Index).Address) And PeriphData(Index).LEDvalue(i) Then paint = True
               Case 8 'OR
                  If RAM(PeriphData(Index).Address) Or PeriphData(Index).LEDvalue(i) Then paint = True
               Case 9 'XOR
                  If RAM(PeriphData(Index).Address) Xor PeriphData(Index).LEDvalue(i) Then paint = True
            End Select
         Next i
         If paint Then frmMain.Peripheral(Index).Picture = images.LEDon.Picture
      End If
   End If
End Sub
Sub UpdateMultiplier(ByVal Index As Integer)
   Dim i As Long
   Dim temp As String, comp As String
   Dim failed As Boolean
   If PeriphData(Index).DownAddress(0) <> PeriphData(Index).Address Then
      If PeriphData(Index).DownAddress(0) <> PeriphData(Index).Address + 1 Then
         If PeriphData(Index).UpAddress(0) <> PeriphData(Index).Address Then
            If PeriphData(Index).UpAddress(0) <> PeriphData(Index).Address + 1 Then
               If PeriphData(Index).DownAddress(0) <> -1 Then
                  If PeriphData(Index).UpAddress(0) <> -1 Then
                     If PeriphData(Index).Address <> -1 Then
                        If RAM(PeriphData(Index).DownAddress(0)) <> -1 Then
                           If RAM(PeriphData(Index).UpAddress(0)) <> -1 Then
                              If PeriphData(Index).UseBCD Then
                                 failed = False
                                 temp = Hex(RAM(PeriphData(Index).UpAddress(0))) & Hex(RAM(PeriphData(Index).DownAddress(0)))
                                 comp = "ABCDEF"
                                 For i = 1 To 6
                                    If InStr(temp, Mid(comp, i, 1)) <> 0 Then
                                       failed = True
                                       Exit For
                                    End If
                                 Next i
                                 If failed = True Then
                                    SetRAM PeriphData(Index).Address, 0
                                    SetRAM PeriphData(Index).Address + 1, 0
                                 Else
                                    i = Val(Hex(RAM(PeriphData(Index).DownAddress(0))))
                                    i = i * Val(Hex(RAM(PeriphData(Index).UpAddress(0))))
                                    SetRAM PeriphData(Index).Address, HexToInt(Mid(Str(i Mod 100), 2))
                                    SetRAM PeriphData(Index).Address + 1, HexToInt(Mid(Str(i \ 100), 2))
                                 End If
                              Else
                                 i = RAM(PeriphData(Index).DownAddress(0))
                                 i = i * RAM(PeriphData(Index).UpAddress(0))
                                 SetRAM PeriphData(Index).Address, i Mod 256
                                 SetRAM PeriphData(Index).Address + 1, i \ 256
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
End Sub
Sub SetRAM(ByVal Address As Long, ByVal value As Integer)
   Dim found As Boolean
   found = False
   If AddToEmuBuff Then
      For i = 0 To EmuCount - 1
         If EmuAddress(i) = Address Then
            EmuData(i) = value
            found = True
            Exit For
         End If
      Next i
      If Not found Then
         EmuAddress(EmuCount) = Address
         EmuData(EmuCount) = value
         EmuCount = EmuCount + 1
      End If
   End If
   
   RAM(Address) = value
   
   For i = 1 To PeriphCount
      If PeriphData(i).Deleted = False Then
         If PeriphData(i).ptype = SevenSeg Then
            If PeriphData(i).Address = Address Then
               Update7Seg i
            End If
         ElseIf PeriphData(i).ptype = LED Then
            If PeriphData(i).Address = Address Then
               UpdateLED i
            End If
         ElseIf PeriphData(i).ptype = LED8 Then
            If PeriphData(i).Address = Address Then
               UpdateLED8 i
            End If
         ElseIf PeriphData(i).ptype = TextDisplay Then
            If PeriphData(i).Address <> -1 Then
               If PeriphData(i).Address <= Address Then
                  If Address < PeriphData(i).Address + PeriphData(i).TextWidth * PeriphData(i).TextHeight Then
                     If Emulating Then
                        For j = 0 To 10
                           If TextsToUpdate(j) = i Then
                              Exit For
                           ElseIf TextsToUpdate(j) = -1 Then
                              TextsToUpdate(j) = i
                              Exit For
                           End If
                        Next j
                     Else
                        RefreshTextBox i
                     End If
                  End If
               End If
            End If
         ElseIf PeriphData(i).ptype = Multiplier Then
            If PeriphData(i).DownAddress(0) = Address Or PeriphData(i).UpAddress(0) = Address Then
               UpdateMultiplier i
            End If
         ElseIf PeriphData(i).ptype = ScreenDisplay Then
            If PeriphData(i).Address <> -1 Then
               If PeriphData(i).ScreenRes = 0 Then
                  If PeriphData(i).Address <= Address Then
                     j = PeriphData(i).ScreenWidth * PeriphData(i).ScreenHeight
                     If j Mod 8 <> 0 Then
                        j = j \ 8 + 1
                     Else
                        j = j / 8
                     End If
                     
                     If Address < PeriphData(i).Address + j Then
                        For j = 0 To 7
                           RefreshScreenPixel i, (((Address - PeriphData(i).Address) * 8) + j) Mod PeriphData(i).ScreenWidth, (((Address - PeriphData(i).Address) * 8) + j) \ PeriphData(i).ScreenWidth
                        Next j
                     End If
                  End If
               ElseIf PeriphData(i).ScreenRes = 1 Or PeriphData(i).ScreenRes = 2 Then
                  If PeriphData(i).Address <= Address Then
                     If Address < (PeriphData(i).Address + PeriphData(i).ScreenWidth * PeriphData(i).ScreenHeight) Then
                        RefreshScreenPixel i, (Address - PeriphData(i).Address) Mod PeriphData(i).ScreenWidth, (Address - PeriphData(i).Address) \ PeriphData(i).ScreenWidth
                     End If
                  End If
               ElseIf PeriphData(i).ScreenRes = 3 Then
                  If PeriphData(i).Address <= Address Then
                     If Address < (PeriphData(i).Address + PeriphData(i).ScreenWidth * PeriphData(i).ScreenHeight * 3) Then
                        RefreshScreenPixel i, ((Address - PeriphData(i).Address) \ 3) Mod PeriphData(i).ScreenWidth, ((Address - PeriphData(i).Address) \ 3) \ PeriphData(i).ScreenWidth
                     End If
                  End If
               End If
            End If
         End If
      End If
   Next i
   If (ExecuteUART And frmMain.ChkUpdate.value = 1) Or ExecuteUART = False Or JustCycling Then
      k = frmMain.ScrollTable.value
      k = k * 4
      If Address >= k Then
         If Address <= (k + RowCount) Then
            UpdateCell Address - k
         End If
      End If
   End If
End Sub

Sub UpdateLED8(ByVal Index As Integer)
   frmMain.Peripheral(Index).Picture = images.LED8off.Picture
   If PeriphData(Index).Address <> -1 Then
      If RAM(PeriphData(Index).Address) <> -1 Then
         i = RAM(PeriphData(Index).Address)
         For j = 7 To 0 Step -1
            If i Mod 2 = 1 Then
               BitBlt frmMain.Peripheral(Index).hDC, 3 + j * 8, 3, 6, 26, images.LED8bar.hDC, 0, 0, SRCCOPY
            End If
            i = i \ 2
         Next j
      End If
   End If
End Sub

Sub ResizeTextBox(ByVal Index As Integer)
   frmMain.Peripheral(Index).Width = PeriphData(Index).TextWidth * 7
   frmMain.Peripheral(Index).Height = PeriphData(Index).TextHeight * 11
   'frmMain.PeripheralText(index).Width = PeriphData(index).TextWidth * 7
   'frmMain.PeripheralText(index).Height = PeriphData(index).TextHeight * 11
   RefreshTextBox Index
End Sub

Sub RefreshTextBox(ByVal Index As Integer)
   Dim temp As String
   Dim counter As Integer
   
   frmMain.Peripheral(Index).Cls
   
   If PeriphData(Index).Address = -1 Then
      'frmMain.PeripheralText(index).Caption = ""
   Else
      For i = 0 To PeriphData(Index).TextWidth * PeriphData(Index).TextHeight - 1
         If i + PeriphData(Index).Address <= 65535 Then
            'counter = counter + 1
            'If counter Mod PeriphData(Index).TextWidth = 0 Then
            '   frmMain.Peripheral(Index).Print temp
            '   temp = ""
            If Len(temp) = PeriphData(Index).TextWidth Then
               frmMain.Peripheral(Index).Print temp
               temp = ""
            End If
            
            If RAM(i + PeriphData(Index).Address) = -1 Then
               temp = temp + " "
            ElseIf RAM(i + PeriphData(Index).Address) = 0 Then
               temp = temp + " "
            ElseIf RAM(i + PeriphData(Index).Address) = 9 Then
               temp = temp + " "
            ElseIf RAM(i + PeriphData(Index).Address) = 10 Then
               temp = temp + " "
            'Could keep track of new lines here instead
            ElseIf RAM(i + PeriphData(Index).Address) = 13 Then
               temp = temp + " "
            ElseIf RAM(i + PeriphData(Index).Address) >= 28 And RAM(i + PeriphData(Index).Address) <= 31 Then
               temp = temp + " "
            Else
               temp = temp + Chr(RAM(i + PeriphData(Index).Address))
            End If
         End If
      Next i
      If temp <> "" Then frmMain.Peripheral(Index).Print temp
   End If
End Sub

Sub ResizeScreen(ByVal Index As Integer)
   frmMain.Peripheral(Index).Width = PeriphData(Index).ScreenWidth * 2
   frmMain.Peripheral(Index).Height = PeriphData(Index).ScreenHeight * 2
   RefreshScreen Index
End Sub

Sub RefreshScreen(ByVal Index As Integer)
   Dim i As Long, j As Long
   
   frmMain.Peripheral(Index).Cls
   If PeriphData(Index).Address <> -1 Then
      For j = 0 To PeriphData(Index).ScreenHeight - 1
         For i = 0 To PeriphData(Index).ScreenWidth - 1
            RefreshScreenPixel Index, i, j
         Next i
      Next j
   End If
End Sub

Sub RefreshScreenPixel(ByVal Index As Long, ByVal x As Integer, ByVal y As Long)
   Dim failed As Boolean
   failed = False
   
   Select Case PeriphData(Index).ScreenRes
      Case 0 '1 bit
         j = (x + y * PeriphData(Index).ScreenWidth) \ 8 + PeriphData(Index).Address
         If j <= 65535 Then
            If RAM(j) = -1 Then
               frmMain.Peripheral(Index).Line (x * 2, y * 2)-(x * 2 + 1, y * 2 + 1), vbGreen, BF
            Else
               i = (x + y * PeriphData(Index).ScreenWidth) Mod 8
               If RAM(j) And (2 ^ i) Then
                  frmMain.Peripheral(Index).Line (x * 2, y * 2)-(x * 2 + 1, y * 2 + 1), vbBlack, BF
               Else
                  frmMain.Peripheral(Index).Line (x * 2, y * 2)-(x * 2 + 1, y * 2 + 1), vbWhite, BF
               End If
            End If
         End If
      Case 1 '8 bit gray
         If x + y * PeriphData(Index).ScreenWidth + PeriphData(Index).Address <= 65535 Then
            If RAM(x + y * PeriphData(Index).ScreenWidth + PeriphData(Index).Address) = -1 Then
               frmMain.Peripheral(Index).Line (x * 2, y * 2)-(x * 2 + 1, y * 2 + 1), vbGreen, BF
            Else
               i = RAM(x + y * PeriphData(Index).ScreenWidth + PeriphData(Index).Address)
               frmMain.Peripheral(Index).Line (x * 2, y * 2)-(x * 2 + 1, y * 2 + 1), RGB(i, i, i), BF
            End If
         End If
      Case 2 '8 bit color
         If x + y * PeriphData(Index).ScreenWidth + PeriphData(Index).Address <= 65535 Then
            If RAM(x + y * PeriphData(Index).ScreenWidth + PeriphData(Index).Address) = -1 Then
               frmMain.Peripheral(Index).Line (x * 2, y * 2)-(x * 2 + 1, y * 2 + 1), vbGreen, BF
            Else
               i = RAM(x + y * PeriphData(Index).ScreenWidth + PeriphData(Index).Address)
               frmMain.Peripheral(Index).Line (x * 2, y * 2)-(x * 2 + 1, y * 2 + 1), RGB256(i), BF
            End If
         End If
      Case 3 '24 bit color
         If x * 3 + y * 3 * PeriphData(Index).ScreenWidth + PeriphData(Index).Address + 2 <= 65535 Then
            For i = 0 To 2
               If RAM(x * 3 + y * 3 * PeriphData(Index).ScreenWidth + PeriphData(Index).Address + i) = -1 Then failed = True
            Next i
            
            If failed Then
               frmMain.Peripheral(Index).Line (x * 2, y * 2)-(x * 2 + 1, y * 2 + 1), vbGreen, BF
            Else
               i = x * 3 + y * 3 * PeriphData(Index).ScreenWidth + PeriphData(Index).Address
               frmMain.Peripheral(Index).Line (x * 2, y * 2)-(x * 2 + 1, y * 2 + 1), RGB(RAM(i), RAM(i + 1), RAM(i + 2)), BF
            End If
         End If
   End Select
End Sub

Function RGB256(ByVal color As Integer) As Long
   Dim colors(3) As Long
   colors(0) = color \ 32 'red
   colors(1) = (color And &H1C) \ 4 'green
   colors(2) = (color And &H3) 'blue
   
   RGB256 = RGB(colors(0) * 36, colors(1) * 36, colors(2) * 85)
End Function

Sub OpenFile(filename As String)
'On Error GoTo eh
   Dim TempCount As Integer, ptype As String
   Dim xin As Integer, yin As Integer
   Dim buff As Integer
   Open filename For Binary As #1
      header = BinRead
      If header = "6502 Trainer Memory File" Then
         For i = 0 To 65535
            Get #1, , buff
            SetRAM i, buff
         Next i
         For i = 0 To 65535
            RamTitles(i) = BinRead
            Get #1, , RamColors(i)
            Get #1, , RamAttribs(i)
            RamLabels(i) = BinRead
            RamOps(i) = BinRead
         Next i
         RefreshAddressList
         RefreshBreakpointList
         RefreshLabelList
         UpdateTable
         For i = 1 To PeriphCount
            TotalRefresh i
         Next i
      ElseIf header = "6502 Trainer Peripheral File" Then
         For i = 1 To PeriphCount
            Unload frmMain.Peripheral(i)
            Unload frmMain.PeriphLabel1(i)
            Unload frmMain.PeriphLabel2(i)
            Unload frmMain.PeriphLabel3(i)
         Next i
         ReDim PeriphData(0)
         PeriphCount = 0
         
         Get #1, , TempCount
         For i = 1 To TempCount
            ptype = BinRead
            Get #1, , xin
            Get #1, , yin
            CreatePeripheral ptype, xin, yin
            
            Get #1, , PeriphData(PeriphCount)
            PeriphData(i).Labels(0) = BinRead
            PeriphData(i).Labels(1) = BinRead
            PeriphData(i).Labels(2) = BinRead
            TotalRefresh i
         Next i
      Else
         MsgBox "This file is corrupt or not compatible.", vbCritical
      End If
   Close #1
   Exit Sub
eh:
   MsgBox "There was an error opening the file: " & Error, vbCritical
   Close #1
End Sub

Function CreatePeripheral(tag As String, x As Integer, y As Integer) As Boolean
   Dim standard As Boolean
   standard = False
   
   If tag = "button" Then standard = True
   If tag = "7seg" Then standard = True
   If tag = "ticker" Then standard = True
   If tag = "switch" Then standard = True
   If tag = "keypad" Then standard = True
   If tag = "led" Then standard = True
   If tag = "bar" Then standard = True
   If tag = "dip" Then standard = True
   If tag = "keyboard" Then standard = True
   If tag = "text" Then standard = True
   If tag = "screen" Then standard = True
   If tag = "multiplier" Then standard = True
   
   If standard = False Then
      MsgBox "Unknown peripheral type: " & tag, vbCritical
      CreatePeripheral = False
      Exit Function
   End If
   
   PeriphCount = PeriphCount + 1
   Load frmMain.Peripheral(PeriphCount)
   frmMain.Peripheral(PeriphCount).Left = x
   frmMain.Peripheral(PeriphCount).Top = y
   frmMain.Peripheral(PeriphCount).Visible = True
   frmMain.Peripheral(PeriphCount).tag = PeriphCount
   
   Load frmMain.PeriphLabel1(PeriphCount)
   Load frmMain.PeriphLabel2(PeriphCount)
   Load frmMain.PeriphLabel3(PeriphCount)
         
   'Load frmMain.PeripheralText(PeriphCount)
   'Set frmMain.PeripheralText(PeriphCount).Container = frmMain.Peripheral(PeriphCount)
   
   ReDim Preserve PeriphData(PeriphCount)
   For i = 0 To 4
      PeriphData(PeriphCount).DownAddress(i) = -1
      PeriphData(PeriphCount).UpAddress(i) = -1
      PeriphData(PeriphCount).DownValue(i) = -1
      PeriphData(PeriphCount).UpValue(i) = -1
   Next i
   
   PeriphData(PeriphCount).Address = -1
   PeriphData(PeriphCount).Locked = True
   PeriphData(PeriphCount).Ticker16 = False
   PeriphData(PeriphCount).TickerInterval = -1
   PeriphData(PeriphCount).switchon = True
   PeriphData(PeriphCount).Deleted = False
   PeriphData(PeriphCount).UseBCD = False
   For i = 0 To 5
      PeriphData(PeriphCount).LEDvalue(i) = -1
      PeriphData(PeriphCount).LEDrelation(i) = 0
   Next i
   
   Select Case tag
   Case "button"
      PeriphData(PeriphCount).ptype = PushButton
      frmMain.Peripheral(PeriphCount).Picture = frmMain.PeripheralButton(0).Picture
   Case "7seg"
      PeriphData(PeriphCount).ptype = SevenSeg
      frmMain.Peripheral(PeriphCount).Picture = images.SevenSegNone.Picture
   Case "ticker"
      PeriphData(PeriphCount).ptype = Ticker
      frmMain.Peripheral(PeriphCount).Picture = frmMain.PeripheralButton(2).Picture
   Case "switch"
      PeriphData(PeriphCount).ptype = SwitchButton
      frmMain.Peripheral(PeriphCount).Picture = images.switchon.Picture
   Case "keypad"
      PeriphData(PeriphCount).ptype = keypad
      frmMain.Peripheral(PeriphCount).Picture = images.keypad.Picture
   Case "led"
      PeriphData(PeriphCount).ptype = LED
      frmMain.Peripheral(PeriphCount).Picture = images.LEDoff.Picture
   Case "bar"
      PeriphData(PeriphCount).ptype = LED8
      frmMain.Peripheral(PeriphCount).Picture = images.LED8off.Picture
   Case "dip"
      PeriphData(PeriphCount).ptype = DipSwitch
      frmMain.Peripheral(PeriphCount).Picture = images.DipSwitch.Picture
   Case "keyboard"
      PeriphData(PeriphCount).ptype = Keyboard
      frmMain.Peripheral(PeriphCount).Picture = frmMain.PeripheralButton(8).Picture
   Case "text"
      PeriphData(PeriphCount).ptype = TextDisplay
      PeriphData(PeriphCount).TextWidth = 20
      PeriphData(PeriphCount).TextHeight = 10
      'frmMain.PeripheralText(PeriphCount).Caption = ""
      'frmMain.PeripheralText(PeriphCount).Visible = True
      ResizeTextBox PeriphCount
   Case "screen"
      PeriphData(PeriphCount).ptype = ScreenDisplay
      PeriphData(PeriphCount).ScreenWidth = 80
      PeriphData(PeriphCount).ScreenHeight = 60
      PeriphData(PeriphCount).ScreenRes = 0
      ResizeScreen PeriphCount
   Case "multiplier"
      PeriphData(PeriphCount).ptype = Multiplier
      frmMain.Peripheral(PeriphCount).Picture = frmMain.PeripheralButton(11).Picture
   End Select
End Function

Sub BinWrite(mess As String)
   Dim buff As Byte
   For i = 1 To Len(mess)
      buff = Asc(Mid(mess, i, 1))
      Put #1, , buff
   Next i
   buff = 0
   Put #1, , buff
End Sub

Function BinRead() As String
   Dim final As String
   Dim buff As Byte
   Do
      Get #1, , buff
      If buff <> 0 Then final = final & Chr(buff)
   Loop While buff <> 0
   BinRead = final
End Function

Sub TotalRefresh(ByVal i As Integer)
   Select Case PeriphData(i).ptype
   Case SevenSeg
      Update7Seg i
   Case DipSwitch
      If PeriphData(i).Address <> -1 Then
         If ExecuteUART And Emulating Then AddToEmuBuff = True
         SetRAM PeriphData(i).Address, PeriphData(i).DipValue
         AddToEmuBuff = False
      End If
      UpdateDip i
   Case LED
      UpdateLED i
   Case LED8
      UpdateLED8 i
   Case ScreenDisplay
      ResizeScreen i
   Case SwitchButton
      If PeriphData(i).switchon = True Then
         frmMain.Peripheral(i).Picture = images.switchon
      Else
         frmMain.Peripheral(i).Picture = images.switchoff
      End If
   Case TextDisplay
      ResizeTextBox i
   Case Ticker
      frmMain.Timer1.Enabled = False
      If PeriphData(i).Address <> -1 Then
         If PeriphData(i).TickerInterval <> -1 Then
            PeriphData(i).TickerStart = GetTickCount()
            frmMain.Timer1.Enabled = True
         End If
      End If
   Case Multiplier
      UpdateMultiplier i
   End Select
   RefreshLabels i
End Sub

Sub UpdateDip(Index As Integer)
   frmMain.Peripheral(Index).Picture = images.DipSwitch.Picture
   For i = 7 To 0 Step -1
      If 2 ^ (7 - i) And PeriphData(Index).DipValue Then
         BitBlt frmMain.Peripheral(Index).hDC, 2 + 8 * i, 9, 6, 14, images.DipSingle.hDC, 0, 0, SRCCOPY
      End If
   Next i
End Sub

Sub AddTextAll(rtf As RichTextBox, ByVal msg As String, color As Long, bold As Boolean, align As Variant, underline As Boolean, italics As Boolean)
    Dim i As Integer, j As Integer
    i = rtf.SelStart
    j = rtf.SelLength
    rtf.SelStart = Len(rtf.text)
    rtf.SelText = msg '+ vbCrLf
    rtf.SelStart = Len(rtf.text) - Len(msg) '+ vbCrLf)
    rtf.SelLength = Len(msg)
    rtf.SelColor = color
    rtf.SelAlignment = align
    rtf.SelBold = bold
    rtf.SelUnderline = underline
    rtf.SelItalic = italics
    rtf.SelLength = 0
    rtf.SelStart = i
    rtf.SelLength = j
    If j = 0 Then rtf.SelStart = Len(rtf.text)
    DoEvents
End Sub

Sub RefreshAddressList()
   Dim i As Long, j As Long, k As Integer
   Dim Address As Long
   Dim color As Long, text As String
   Dim Address2 As Long
   Dim color2 As Long, text2 As String
   
   color = RamColors(0)
   text = RamTitles(0)
   Address = 0
   RamStarts(0) = 0
   
   color2 = RamColors(65535)
   text2 = RamTitles(65535)
   Address2 = 65535
   RamEnds(65535) = 65535
   
   For i = 1 To 65535
      If (RamColors(i) <> color) Or (RamTitles(i) <> text) Then
         color = RamColors(i)
         text = RamTitles(i)
         Address = i
      End If
      RamStarts(i) = Address
      
      j = 65535 - i
      If (RamColors(j) <> color2) Or (RamTitles(j) <> text2) Then
         color2 = RamColors(j)
         text2 = RamTitles(j)
         Address2 = j
      End If
      RamEnds(j) = Address2
   Next i
   
   oldindex = MemManager.LstSections.ListIndex
   'oldsize = MemManager.LstSections.ListCount
   MemManager.LstSections.Clear
   Address = 0
   j = 0
   ReDim SectionList(0)
   Do
      If RamTitles(Address) <> "" Then
         MemManager.LstSections.AddItem HexBig(RamStarts(Address)) & "-" & HexBig(RamEnds(Address)) & " " & RamTitles(Address)
         SectionList(j) = RamStarts(Address)
         j = j + 1
         ReDim Preserve SectionList(j)
      End If
      If Address <> 65535 Then
         'Address = RamEnds(Address + 1)
         Address = RamEnds(Address) + 1
      Else
         Address = 65536
      End If
   Loop While Address <> 65536
   
   If oldindex >= MemManager.LstSections.ListCount Then
      MemManager.LstSections.ListIndex = oldindex - 1
   Else
      MemManager.LstSections.ListIndex = oldindex
   End If
End Sub

Sub RefreshBreakpointList()
   j = MemManager.LstBreakpoints.ListIndex
   MemManager.LstBreakpoints.Clear
   For i = 0 To 65535
      If RamAttribs(i) And AttribBreakpoint Then
         MemManager.LstBreakpoints.AddItem HexBig(i)
      End If
   Next i
   If j < MemManager.LstBreakpoints.ListCount Then
      MemManager.LstBreakpoints.ListIndex = j
   Else
      MemManager.LstBreakpoints.ListIndex = MemManager.LstBreakpoints.ListCount - 1
   End If
End Sub

Sub RefreshLabelList()
   j = MemManager.LstFunctions.ListIndex
   MemManager.LstFunctions.Clear
   For i = 0 To 65535
      If RamLabels(i) <> "" Then
         MemManager.LstFunctions.AddItem HexBig(i) & " " & RamLabels(i)
      End If
   Next i
   If j < MemManager.LstFunctions.ListCount Then
      MemManager.LstFunctions.ListIndex = j
   Else
      MemManager.LstFunctions.ListIndex = MemManager.LstBreakpoints.ListCount - 1
   End If
End Sub

Sub main()
   Dim buff As String, ptr As Long
   Dim temp As String
   ExecuteUART = False
   StopUART = False
   IgnoreKey = False
   WaitingForInput = False
   WaitingSingleCycle = False
   ToDisableUART = False
   Emulating = False
   AddToEmulationBuff = False
   ToStopEmulating = False
   ToReset = False
   ToLoad = False
   
   HighlightPtr = HexToInt("FFFC")
   
   Load frmMain
   frmMain.Show
   
   Do
      If ExecuteUART Then
         If Emulating Then
            'WaitInput
            'buff = frmMain.MSComm1.Input
            'MsgBox Hex(Asc(Mid(buff, 1, 1))) & "=" & Hex(Asc(Mid(buff, 2, 1)))
            'MsgBox Hex(Asc(Mid(buff, 3, 1))) & "=" & Hex(Asc(Mid(buff, 4, 1)))
            'MsgBox Hex(Asc(Mid(buff, 5, 1))) & "=" & Hex(Asc(Mid(buff, 6, 1)))
            'MsgBox Hex(Asc(Mid(buff, 7, 1))) & "=" & Hex(Asc(Mid(buff, 8, 1)))
            'End
            
            'If WaitInput(500, 199) Then ' message, flag, 2 flag addresses, 2 cycle counts, dirty count, 64 addresses+data
            If WaitInput(500 + Val(frmMain.TxtUpdate.text), 7) Then
               t = GetTickCount
               buff = frmMain.MSComm1.Input
               If Asc(Mid(buff, 1, 1)) = UPDATE_DIRTY_RAM Then
                  '1=message
                  '2=extra flag
                  '3=flag address
                  '4=flag address
                  '5=cycle count
                  '6=cycle count
                  '7=DirtyCount
                  CycleCount = CycleCount + Asc(Mid(buff, 5, 1))
                  CycleCount = CycleCount + Asc(Mid(buff, 6, 1)) * 256
                  j = Asc(Mid(buff, 7, 1)) - 1
                  
                  
                  For i = 0 To j
                     Do While Len(buff) < i * 3 + 3 + 7
                        DoEvents
                        If frmMain.MSComm1.InBufferCount > 0 Then buff = buff + frmMain.MSComm1.Input
                     Loop
                     
                     ptr = Asc(Mid(buff, 9 + i * 3, 1))
                     ptr = ptr * 256
                     ptr = ptr + Asc(Mid(buff, 8 + i * 3, 1))
                     
                     SetRAM ptr, Asc(Mid(buff, 10 + i * 3, 1))
                  Next i
                  
                  'textcounter = 0
                  For i = 0 To 10
                     If TextsToUpdate(i) = -1 Then
                        'Exit For
                     Else
                        RefreshTextBox TextsToUpdate(i)
                        TextsToUpdate(i) = -1
                        'textcounter = textcounter + 1
                     End If
                  Next i
                  DoEvents 'To refresh textboxes
                  
                  If frmMain.MSComm1.InBufferCount > 0 Then
                     ToStopEmulating = True
                  ElseIf Asc(Mid(buff, 2, 1)) <> 0 Then
                     ToStopEmulating = True
                  End If
                  
                  If ToStopEmulating Then
                     'DoEvents
                     finaltime = (GetTickCount - GlobalTime) / 1000
                     If finaltime < 0 Then MsgBox "Error: " & finaltime
                     frmMain.MSComm1.Output = Chr(STOP_EMULATING)
                     WaitInput 500, 1
                     temp = frmMain.MSComm1.Input
                  Else
                     temp = Chr(KEEP_EMULATING) & Chr(EmuCount)
                     For i = 0 To EmuCount - 1
                        temp = temp & Chr(EmuAddress(i) Mod 256) & Chr(EmuAddress(i) \ 256) & Chr(EmuData(i))
                     Next i
                     frmMain.MSComm1.Output = temp
                     EmuCount = 0
                  End If
                  
                  If frmMain.MSComm1.InBufferCount > 0 Then
                     MsgBox "Too much data tansferred from chip.", vbCritical
                  ElseIf Asc(Mid(buff, 2, 1)) <> 0 Then
                     HighlightPtr = Asc(Mid(buff, 4, 1))
                     HighlightPtr = HighlightPtr * 256 + Asc(Mid(buff, 3, 1))
                     MoveHighlightPtr
                     JumpTable HighlightPtr
                     i = Asc(Mid(buff, 2, 1))
                     If i And AttribBreakpoint Then
                        'MsgBox "Breakpoint at " & HexSmall(Asc(Mid(buff, 4, 1))) & HexSmall(Asc(Mid(buff, 3, 1))) & "."
                     ElseIf i And AttribReadonly Then
                        MsgBox "Write to read-only at " & HexSmall(Asc(Mid(buff, 4, 1))) & HexSmall(Asc(Mid(buff, 3, 1))) & ".", vbCritical
                     ElseIf i And AttribUninitialized Then
                        MsgBox "Uninitialized read at " & HexSmall(Asc(Mid(buff, 4, 1))) & HexSmall(Asc(Mid(buff, 3, 1))) & ".", vbCritical
                     End If
                  End If
                  
                  If ToStopEmulating Then
                     If finaltime <> 0 Then
                        MsgBox "Time: " & finaltime & vbCrLf & "Cycles: " & CycleCount & vbCrLf & Round(CycleCount / finaltime, 2) & " hz"
                     Else
                        MsgBox "Time: " & finaltime & vbCrLf & "Cycles: " & CycleCount
                     End If
                     DisableUART
                     UpdateTable
                     If ToReset Then
                        Call frmMain.BtnReset_MouseUp(vbLeftButton, 0, 0, 0)
                        ToReset = False
                     End If
                     If ToLoad Then
                        Call MemManager.BtnLoad_Click
                        ToLoad = False
                     End If
                  End If
                  
               End If
            Else
               MsgBox "Connection to device timed out.", vbCritical
               DisableUART
            End If
         Else
            SingleCycle
            If ToDisableUART Then
               t = (GetTickCount - GlobalTime) / 1000
               MsgBox "Time: " & t & vbCrLf & "Cycles: " & CycleCount & vbCrLf & Round(CycleCount / t, 2) & " hz"
               DisableUART
               UpdateTable
            End If
         End If
      End If
      DoEvents
   Loop
End Sub

Function EnableUART()
   'On Error GoTo ehandler
   Dim failed As Boolean
   failed = False
   
   EnableUART = False
   
   If frmMain.MSComm1.PortOpen Then frmMain.MSComm1.PortOpen = False
   
   frmMain.MSComm1.CommPort = frmMain.ComboPorts.ListIndex + 1
   If frmMain.ChkHispeed.value = 1 Then
      frmMain.MSComm1.Settings = "57600,N,8,1"
   Else
      frmMain.MSComm1.Settings = "9600,N,8,1"
   End If
   frmMain.MSComm1.PortOpen = True
   
   If True Then
      Dim TempDCB As DCB
      If GetCommState(frmMain.MSComm1.CommID, TempDCB) = 0 Then
         MsgBox "Unable to set the baud rate.", vbCritical
         failed = True
      Else
         If CompactMode Then
            TempDCB.BaudRate = 500000
         Else
            TempDCB.BaudRate = 1000000
         End If
         If SetCommState(frmMain.MSComm1.CommID, TempDCB) = 0 Then
            MsgBox "Unable to set the baud rate.", vbCritical
            failed = True
         End If
      End If
   End If
   
   If failed = False Then
      frmMain.MSComm1.Output = FormMsg(PING, "PING!", PACKET_SIZE)
      
      If WaitInput = False Then 'timeout
         Sleep 200
         frmMain.MSComm1.Output = FormMsg(PING, "PING!", PACKET_SIZE)
         If WaitInput = False Then
            MsgBox "Unable to connect to the device.", vbCritical
            failed = True
         End If
      End If
   End If
   
   If failed = False Then
      buff = frmMain.MSComm1.Input
      
      If Asc(Mid(buff, PACKET_SIZE, 1)) = PONG And Left(buff, 5) = "PONG!" Then
         EnableUART = True
         ExecuteUART = True
      Else
         MsgBox "Device connected but did not authenticate.", vbCritical
         failed = True
      End If
   End If
   Exit Function
ehandler:
   If Err = 8002 Then
      MsgBox "Unable to open COM" & frmMain.ComboPorts.ListIndex + 1 & ".", vbCritical
   Else
      MsgBox "Unknown error connecting: " & Error, vbCritical
   End If
End Function

Sub DisableUART()
'somehow all these debugs stop it from crashing
   ToDisableUART = False
   ToStopEmulating = False
   dbug "DU1"
   If frmMain.MSComm1.PortOpen Then
      dbug "DU2"
      'Do While WaitingForInput
      '   a = WaitingSingleCycle
      '   DoEvents
      'Loop
      dbug "DU3"
      frmMain.MSComm1.PortOpen = False
      Sleep 100
      dbug "DU4"
   End If
   dbug "DU5"
   If frmMain.MSComm1.InputLen > 0 Then buff = frmMain.MSComm1.Input
   dbug "DU6"
   frmMain.BtnPlay.Picture = images.Cycling.Picture
   frmMain.BtnPlay.ToolTipText = "Start stepping"
   frmMain.Emulate.Picture = images.play.Picture
   frmMain.BtnPlay.ToolTipText = "Start execution"
   frmMain.TxtUpdate.Enabled = True
   dbug "DU7"
   Emulating = False
   dbug "DU8"
   ExecuteUART = False
   dbug "DU9"
End Sub

Function FormMsg(msg_type As Integer, ByVal msg As String, length As Integer)
   Dim buff As String
   FormMsg = msg & Space(length - Len(msg) - 1) & Chr(msg_type)
End Function

Function WaitInput(Optional TimeToWait As Integer, Optional SizeOfPacket As Integer)
   Dim LimitTime As Single
   
   If SizeOfPacket = 0 Then SizeOfPacket = PACKET_SIZE
   
   If TimeToWait <> 0 Then
      LimitTime = TimeToWait '/ 1000
   Else
      LimitTime = 500 '0.5
   End If
   WaitingForInput = True
   t = GetTickCount
   Do While frmMain.MSComm1.InBufferCount < SizeOfPacket
      DoEvents
      If GetTickCount - t > LimitTime Then
         WaitingForInput = False
         WaitInput = False
         Exit Function
      End If
   Loop
   WaitInput = True
   WaitingForInput = False
End Function

Sub Sleep(ms As Integer)
   t = GetTickCount
   While GetTickCount - t < ms '(ms / 1000)
      DoEvents
   Wend
End Sub

Sub dbug(msg As String)
   Open App.Path & "\log.txt" For Append As #100
      Print #100, msg
   Close #100
End Sub

Sub SingleCycle()
   Dim SendData As String
   Dim CycleThisTime As Boolean
   Dim report As String
   CycleCount = CycleCount + 1
   WaitingSingleCycle = True
   frmMain.MSComm1.Output = FormMsg(DOWN_CYCLE, "", PACKET_SIZE)
   If WaitInput Then
      buff = frmMain.MSComm1.Input
      ptr = Asc(Mid(buff, 1, 1))
      ptr = (ptr * 256 + Asc(Mid(buff, 2, 1)))
      
      If frmMain.ChkNOP.value = 1 Then
         Open App.Path & "\NOP test.txt" For Append As #25
            Print #25, HexBig(ptr)
         Close #25
      End If
      
      'report = "Down cycle: " & HexBig(ptr) & vbCrLf & "Flags: " & Hex(Asc(Mid(buff, 3, 1))) & vbCrLf
      
      HighlightPtr = ptr
      MoveHighlightPtr
      If frmMain.ChkJump.value = 1 Then JumpTable ptr
      
      SendData = ""
      
      CycleThisTime = False
      If frmMain.ChkNOP.value = 1 Then CycleThisTime = True
      For i = 0 To 11
         If frmMain.ChkNOP.value = 1 Then
            SendData = SendData & Chr(&HEA)
         Else
            If ptr + i > 65535 Then
               SendData = SendData & Chr(0)
               CycleThisTime = True
            ElseIf RAM(ptr + i) = -1 Then
               SendData = SendData & Chr(0)
               CycleThisTime = True
            Else
               SendData = SendData & Chr(RAM(ptr + i))
            End If
         End If
      Next i
      If JustCycling Or CycleThisTime Then
         SendData = SendData & Chr(&H55) 'no speed up
      Else
         SendData = SendData & Chr(&HAA) 'speed up
      End If
      frmMain.MSComm1.Output = FormMsg(UP_CYCLE, SendData, PACKET_SIZE)
      
      If WaitInput Then
         buff = frmMain.MSComm1.Input
         input_flags = Asc(Mid(buff, 3, 1))
         Open App.Path & "\crash log.txt" For Append As #10
            Print #10, HexBig(ptr) & " " & Hex(input_flags)
         Close #10
         If Asc(Mid(buff, PACKET_SIZE, 1)) = UP_ACK_READ Then
            'report = report & "Read: " & Hex(RAM(ptr))
            CycleCount = CycleCount + Asc(Mid(buff, 4, 1))
            If frmMain.ChkNOP.value = 0 Then
               If (input_flags And CPU_VPA) Or (input_flags And CPU_VDA) Then
                  If RAM(ptr) = -1 And frmMain.ChkBreakUninit.value = 1 Then
                     MsgBox "Unitialized read at " & HexBig(ptr) & ".", vbCritical, "Break"
                     DisableUART
                  End If
               End If
            End If
         ElseIf Asc(Mid(buff, PACKET_SIZE, 1)) = UP_ACK_WRITE Then
            If frmMain.ChkNOP.value = 0 Then
               If (RamAttribs(ptr) And AttribReadonly) And frmMain.ChkBreakROWrite.value = 1 Then
                  MsgBox "Write to read-only at " & HexBig(ptr) & ".", vbCritical, "Break"
                  DisableUART
               Else
                  'report = report & "Write: " & Hex(Asc(Mid(buff, 4, 1)))
                  SetRAM ptr, Asc(Mid(buff, 4, 1))
               End If
            Else
               MsgBox "Write operation while NOP testing.", vbCritical
               DisableUART
            End If
         End If
         If frmMain.ChkNOP.value = 0 Then
            If RamAttribs(ptr) And AttribBreakpoint Then
               If JustCycling = False Then
                  t = (GetTickCount - GlobalTime) / 1000
                  MsgBox "Time: " & t & vbCrLf & "Cycles: " & CycleCount & vbCrLf & Round(CycleCount / t, 2) & " hz"
                  DisableUART
                  JumpTable ptr
               End If
            End If
         End If
         
         'MsgBox report
      Else
         MsgBox "Connection to device timed out.", vbCritical
         DisableUART
      End If
   Else
      MsgBox "Connection to device timed out.", vbCritical
      DisableUART
   End If
   WaitingSingleCycle = False
End Sub

Sub JumpTable(ByVal ptr As Long)
   j = ptr \ 4
   If j > frmMain.ScrollTable.Max Then
      frmMain.ScrollTable.value = frmMain.ScrollTable.Max
      'should it be RowCount or RowCount/4?
   ElseIf j > frmMain.ScrollTable.value And j <= frmMain.ScrollTable.value + RowCount / 4 Then
      
   Else
      frmMain.ScrollTable.value = j
   End If
End Sub

Sub MoveHighlightPtr()
   Dim j As Long
   If (ExecuteUART And frmMain.ChkUpdate.value = 1) Or ExecuteUART = False Or JustCycling Then
      If HighlightPtr \ 4 >= frmMain.ScrollTable.value Then
         If HighlightPtr \ 4 <= frmMain.ScrollTable.value + RowCount Then
            j = frmMain.ScrollTable.value
            frmMain.ShpHighlight.Top = (HighlightPtr - j * 4) * (frmMain.ShpHighlight.Height - 1)
            frmMain.ShpHighlight.Visible = True
         Else
            frmMain.ShpHighlight.Visible = False
         End If
      Else
         frmMain.ShpHighlight.Visible = False
      End If
   End If
End Sub
