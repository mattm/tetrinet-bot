Attribute VB_Name = "modUniversal"
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVK As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetTickCount& Lib "kernel32" ()
Public Declare Function IsWindowEnabled& Lib "user32" (ByVal hwnd As Long)
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal CMD As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetActiveWindow& Lib "user32" (ByVal hwnd As Long)
Public Declare Function GetForegroundWindow& Lib "user32" ()
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String)
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_CHAR = &H102
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const GW_NEXT = 2
Public Const GW_CHILD = 5
Public Const WM_SETTEXT = &HC
Public Const SWP_NOSIZE = &H1
Public Const VK_UP = &H26
Public Const VK_LEFT = &H25
Public Const VK_RIGHT = &H27
Public Const VK_SPACE = &H20
Public Const VK_DOWN = &H28
Public Const VK_ESCAPE = &H1B

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_GETTEXT = &H189
Public Const SW_NORMAL = &H1
Public Const KEY_DOWN = -32767


Type POINTAPI
    x As Long
    y As Long
End Type

Public Sub Timeout(sSeconds As Variant)
    Dim startTime&, curTime&
    
    sSeconds = sSeconds * 1000
    startTime = GetTickCount()
    
    Do:  DoEvents
        curTime = GetTickCount()
    Loop Until ((curTime - startTime) > sSeconds)

End Sub

Public Function GetText(hwnd As Long) As String
    
    Dim getLen%
    
    Dim StringText As String
    
    getLen = SendMessageByNum(hwnd, WM_GETTEXTLENGTH, 0&, 0&)
    StringText = String(getLen + 1, Chr(0))
    Call SendMessageByString(hwnd, WM_GETTEXT, getLen + 1, StringText$)
    
    GetText = Left(StringText, InStr(StringText, Chr(0)) - 1)
    
 End Function

Public Sub ClickIcon(iconHWnd As Long)
    
    Call SendMessage(iconHWnd&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(iconHWnd&, WM_LBUTTONUP, 0&, 0&)

End Sub

Public Function GetColor(xCoor As Integer, yCoor As Integer) As String
    
    Dim pt As POINTAPI
    Dim hWind&, dc&, clr&
    
    pt.x = xCoor
    pt.y = yCoor
    hWind = WindowFromPoint(pt.x, pt.y)
    dc = GetDC(hWind)
    ScreenToClient hWind, pt
    clr = GetPixel(dc, pt.x, pt.y)
    ReleaseDC Wind&, dc
 
    GetColor = Hex(clr&)
End Function

Public Function CountChr(sourceString As String, sChr As String)
    Dim GetChar&, Times&
    
    For GetChar = 1 To Len(sourceString)
        If Mid(sourceString, GetChar, 1) = sChr Then Times = Times + 1
    Next GetChar
    CountChr = Times

End Function

Public Function tForm1() As Long
    tForm1& = FindWindow("TForm1", vbNullString)
End Function

Public Function tForm2() As Long
    tForm2& = FindWindow("TForm2", vbNullString)
End Function

Public Sub SetForm2()
    If GetForegroundWindow <> tForm2 Then
        ClickCategory "Show Fields"
        BringWindowToTop tForm2
        SetWindowPos tForm1, 0, 10, 10, 0, 0, SWP_NOSIZE
        SetWindowPos tForm2, 0, 10, 100, 0, 0, SWP_NOSIZE
    End If
End Sub

Public Sub ClickCategory(CategoryName As String)
    Dim tButton&
    tButton = FindWindowEx(tForm1, 0, "TPanel", vbNullString)
    Do: DoEvents
        tButton& = GetWindow(tButton, GW_NEXT)
        If tButton = 0 Then Exit Sub
    Loop Until Trim(GetText(tButton)) = CategoryName
    ClickIcon tButton
End Sub

Public Function InList(lstBox As ListBox, strSearch As String) As Boolean
    Dim searchList As Integer
    For searchList = 0 To lstBox.ListCount - 1
        If lstBox.List(searchList) = strSearch Then InList = True: Exit Function
    Next searchList
    InList = False
End Function

Public Sub PressKey(Key As Integer)
    keybd_event Key, 0, 0, 0
End Sub

Public Function GameOn() As Boolean
    Dim tPanel&, tButton&
    
    tPanel = FindWindowEx(tForm1, 0, "TPanel", vbNullString)
    tButton = FindWindowEx(tPanel, 0, "TButton", "Stop Current Game")
    If tButton Then GameOn = True Else GameOn = False
End Function

Public Sub UpdateStatus(vColor, vUpdate As String, Optional vBold As Boolean = False)
    With frmMain.rtbStatus
        .SelColor = vColor
        .SelBold = vBold
        .SelText = vUpdate
        .SelStart = Len(.Text)
    End With
End Sub

Public Sub ShowNormDisplay(bBlock As String, AddInfo As String)
    Dim eRow%, aRow%, clrHoles%
    Dim tRow$, rChr$
    With frmNormDisplay
        For clrHoles = 0 To 15
            .picBlock(clrHoles).BackColor = vbWhite
        Next clrHoles
            For eRow = 1 To MatrixRowCount(bBlock)
                tRow = MatrixRow(bBlock, eRow)
                For aRow = 1 To Len(tRow)
                    rChr = Mid(tRow, aRow, 1)
                    If rChr = "P" Or rChr = "1" Then
                       .picBlock(eRow - 1 + 4 * (aRow - 1)).BackColor = &HC95331
                    End If
                Next aRow
            Next eRow
            .rtfInfo.Text = vbNullString
            .rtfInfo.SelText = AddInfo
    End With
End Sub

Public Sub CheckSticks(Block As String)

    Dim stickStats$, sSpot%, sSoFar%, sPerc%
    If Block = " 0 " Or _
        Block = "PPPP" & vbCrLf Or _
        Block = " 0 0 0 0 " Or _
        Block = "P" & vbCrLf & "P" & vbCrLf & "P" & vbCrLf & "P" & vbCrLf Then
        stickStats = frmMain.lblSticks.Caption
        sSpot = InStr(stickStats, Chr(32))
        sSoFar = Val(Left(stickStats, sSpot - 1))
        frmMain.lblSticks.Caption = (sSoFar + 1) & Chr(32)
    End If
        stickStats = frmMain.lblSticks.Caption
        sSpot = InStr(stickStats, Chr(32))
        sSoFar = Val(Left(stickStats, sSpot - 1))
        sPerc = (sSoFar / (frmMain.lblBlocksDropped.Caption + 1)) * 100
        frmMain.lblSticks.Caption = sSoFar & " (" & sPerc & "%)"
    
End Sub

Public Sub UpdateSeconds()
    Dim timeDiff
    
    timeDiff = (GetTickCount - frmMain.lblStart.Caption) / 1000
    frmMain.lblSecondsPlaying.Caption = Val(Format(timeDiff, "00.0")) & "s"
    
End Sub

Public Function APM() As Integer

    Dim tsRichEdit&
    Dim weapText$
    Dim spotParse%, spotLast%, pSpot%
    Dim gameLength%, rVal%, total%, ratioAPM%

    tsRichEdit = FindWindowEx(tForm2, 0, "TSRichEdit", vbNullString)
    weapText = GetText(tsRichEdit)

    spotLast = 0
    Do: DoEvents
    
        spotParse = InStr(spotLast + 1, weapText, Chr(13))
        If spotParse = 0 Then Exit Do
        
        rLine = Mid(weapText, spotLast + 1, spotParse - spotLast - 1)
        
        If InStr(rLine, "Added to All from " & myNick) Then
            pSpot = InStr(rLine, ".")
            rVal = Val(Mid(rLine, pSpot + 2, 1))
            total = total + rVal
        End If
        
        spotLast = spotParse + 1
    Loop Until spotLast > Len(weapText)
    
    gameLength = Val(frmMain.lblSecondsPlaying.Caption)
    If gameLength <> 0 Then ratioAPM = (60 * total) / gameLength
    
    APM = ratioAPM
End Function

Public Sub FixIfDied()
    Dim xCoor%, yCoor%
    Dim sColor$
    
    xCoor = 6 + 16 * 1
    yCoor = 198 + 16 * -3
    sColor = GetColor(xCoor, yCoor)
    If sColor = "FF8000" Then
        PressKey VK_SPACE
        PressKey VK_SPACE
        SetForegroundWindow tForm1
    End If
    
End Sub
