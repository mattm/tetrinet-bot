Attribute VB_Name = "modSpecials"
Option Explicit

Const OPP_OTHER = "FF0000"
Const OPP_O = "FF"
Const OPP_NG = "FF00"
Const OPP_BLOCK = "C0C0C0"
Const OPP_FIELD = "FFFFFF"
Const OPP_S = "FF8080"
Const OPP_BACK = "6B5A58"

Const myNick = "Pawn"

Public Const NOT_PLAYING = "0 0 0 0 0 0 0 0 0 0 0 0 "

Public Function OpponentFieldMatrix(Position As Integer) As String
    Dim eachPlayer%, column%, Row%
    Dim sColor$, heights$
    Dim bFound As Boolean
        
        heights = vbNullString
        For column = 0 To (12 - 1)
            bFound = False
            For Row = 0 To (22 - 1)
                Select Case Position
                    Case 1: sColor = GetColor(319 + 8 * column, 144 + 8 * Row)
                    Case 2: sColor = GetColor(419 + 8 * column, 144 + 8 * Row)
                    Case 3: sColor = GetColor(517 + 8 * column, 144 + 8 * Row)
                    Case 4: sColor = GetColor(419 + 8 * column, 342 + 8 * Row)
                    Case 5: sColor = GetColor(517 + 8 * column, 342 + 8 * Row)
                End Select
                If sColor <> OPP_BACK And sColor <> "0" And sColor <> OPP_FIELD Then
                   heights = heights & 22 - Row & Chr(32)
                   bFound = True
                   Exit For
                End If
            Next Row
            If bFound = False Then heights = heights & "0 "
        Next column
        OpponentFieldMatrix = heights
End Function


Public Function AnalyzeWRTBQR() As Integer
    Dim sPlayers$, sNum$, playerMatrix$, fColor$
    Dim getEachPlayer%, eachPlayer%, cLow%, cHigh%, weaps%
    Dim column%, Row%, oHeight%, oNumber%, nNumber%, sNumber%, wRecord%, wNumber%

    sPlayers = NumbersPlaying
        For getEachPlayer = 1 To Len(sPlayers)
            sNum = Mid(sPlayers, getEachPlayer, 1)
            eachPlayer = GetPositionOfNumber(Val(sNum))
            
            If IsPlaying(eachPlayer) Then
                playerMatrix = OpponentFieldMatrix(eachPlayer)
                cLow = MatrixItemLowest(playerMatrix)
                cHigh = MatrixItemHighest(playerMatrix)
                weaps = 0
                For column = 0 To (12 - 1)
                    For Row = 0 To (22 - 1) - (cLow) ' - 1)
                        Select Case eachPlayer
                            Case 1: fColor = GetColor(319 + 8 * column, 144 + 8 * Row)
                            Case 2: fColor = GetColor(419 + 8 * column, 144 + 8 * Row)
                            Case 3: fColor = GetColor(517 + 8 * column, 144 + 8 * Row)
                            Case 4: fColor = GetColor(419 + 8 * column, 342 + 8 * Row)
                            Case 5: fColor = GetColor(517 + 8 * column, 342 + 8 * Row)
                        End Select
                        Select Case fColor
                            Case OPP_OTHER
                                weaps = weaps + 1
                            
                            Case OPP_O
                                If cHigh < oHeight Or oHeight = 0 Then
                                    oHeight = cHigh
                                    oNumber = eachPlayer
                                End If
                            Case OPP_NG
                                nNumber = eachPlayer
                            Case OPP_S
                                sNumber = eachPlayer
                        End Select
                    Next Row
            Next column
            If weaps >= wRecord Then
                wRecord = weaps
                wNumber = eachPlayer
            End If
        End If
    Next getEachPlayer

    If sNumber <> 0 Then AnalyzeWRTBQR = GetNumberofPosition(sNumber): Exit Function
    If nNumber <> 0 Then AnalyzeWRTBQR = GetNumberofPosition(nNumber): Exit Function
    If oNumber <> 0 Then AnalyzeWRTBQR = GetNumberofPosition(oNumber): Exit Function
    If wRecord <> 0 Then AnalyzeWRTBQR = GetNumberofPosition(wNumber): Exit Function
    AnalyzeWRTBQR = 0
    
End Function

Public Function AnalyzeWRTA() As Integer
        
        Dim sPlayers$, sNum$, playerMatrix$
        Dim getEachPlayer%, eachPlayer%, cLow%, lRecord%, lPlayer%

        sPlayers = NumbersPlaying
        For getEachPlayer = 1 To Len(sPlayers)
            
            sNum = Mid(sPlayers, getEachPlayer, 1)
            eachPlayer = GetPositionOfNumber(Val(sNum))
            playerMatrix = OpponentFieldMatrix(eachPlayer)
            cLow = MatrixItemHighest(playerMatrix)
            
            If (cLow < lRecord Or lRecord = 0) And playerMatrix <> NOT_PLAYING Then
                lRecord = cLow
                lPlayer = eachPlayer
            End If
       
       Next getEachPlayer
       AnalyzeWRTA = GetNumberofPosition(lPlayer)
End Function

Public Function AnalyzeWRTO() As Integer

    Dim sPlayers$, playerMatrix$, sColor$, sNum$
    Dim getEachPlayer%, eachPlayer%, cHigh%, column%, Row%
    Dim oHeight%, oNumber%
    
    sPlayers = NumbersPlaying
    For getEachPlayer = 1 To Len(sPlayers)
        sNum = Mid(sPlayers, getEachPlayer, 1)
        eachPlayer = GetPositionOfNumber(Val(sNum))
        playerMatrix = OpponentFieldMatrix(eachPlayer)
        cHigh = MatrixItemHighest(playerMatrix)
        For column = 0 To (12 - 1)
            For Row = 0 To (22 - 1)
                Select Case eachPlayer
                    Case 1: sColor = GetColor(319 + 8 * column, 144 + 8 * Row)
                    Case 2: sColor = GetColor(419 + 8 * column, 144 + 8 * Row)
                    Case 3: sColor = GetColor(517 + 8 * column, 144 + 8 * Row)
                    Case 4: sColor = GetColor(419 + 8 * column, 342 + 8 * Row)
                    Case 5: sColor = GetColor(517 + 8 * column, 342 + 8 * Row)
                End Select
                Select Case sColor
                    Case OPP_O
                        If cHigh < oHeight Or oHeight = 0 Then
                            oHeight = cHigh
                            oNumber = eachPlayer
                        End If
                    End Select
            Next Row
        Next column
    Next getEachPlayer
    AnalyzeWRTO = GetNumberofPosition(oNumber)

End Function

Public Function AnalyzeWRTC() As Long
    
    Dim sPlayers$, sNum$, playerMatrix$, sColor$
    Dim getEachPlayer%, eachPlayer%, cLow%, column%
    Dim oNumber%, wNumber%

    sPlayers = NumbersPlaying
    For getEachPlayer = 1 To Len(sPlayers)
        sNum = Mid(sPlayers, getEachPlayer, 1)
        eachPlayer = GetPositionOfNumber(Val(sNum))
        playerMatrix = OpponentFieldMatrix(eachPlayer)
        cLow = MatrixItemLowest(playerMatrix)
        If playerMatrix <> NOT_PLAYING And cLow = 0 Then
            For column = 0 To (12 - 1)
                Select Case eachPlayer
                    Case 1: sColor = GetColor(319 + 8 * column, 144 + 8 * 21)
                    Case 2: sColor = GetColor(419 + 8 * column, 144 + 8 * 21)
                    Case 3: sColor = GetColor(517 + 8 * column, 144 + 8 * 21)
                    Case 4: sColor = GetColor(419 + 8 * column, 342 + 8 * 21)
                    Case 5: sColor = GetColor(517 + 8 * column, 342 + 8 * 21)
                End Select
                Select Case sColor
                    Case OPP_S, OPP_NG
                        AnalyzeWRTC = eachPlayer
                        Exit Function
                    Case OPP_O
                        oNumber = eachPlayer
                    Case OPP_OTHER
                        wNumber = eachPlayer
                End Select
            Next column
        End If
    Next getEachPlayer
    
    If oNumber <> 0 Then AnalyzeWRTC = GetNumberofPosition(oNumber): Exit Function

    If wNumber <> 0 Then AnalyzeWRTC = GetNumberofPosition(wNumber): Exit Function

    If SelfClearNecessary = True Then
        AnalyzeWRTC = 7
        Exit Function
    End If
    
    AnalyzeWRTC = 0
End Function


Public Function NumbersPlaying() As String

    Dim tPanel&, tList&
    
    Dim tCount%, getItem%, sLength%, verify%, rNum%, rPos%
    Dim bString$, numList$, rFinal$


    tPanel = FindWindowEx(tForm1, 0, "TPanel", vbNullString)
    tList = FindWindowEx(tPanel, 0, "TListBox", vbNullString)
    tCount = SendMessage(tList, LB_GETCOUNT, 0, 0)
    
    For getItem = 0 To tCount - 1
        sLength = SendMessage(tList, LB_GETTEXTLEN, getItem, 0)
        bString = String(32, Chr(0))
        Call SendMessageByString(tList, LB_GETTEXT, getItem, bString)
        If Left(bString, 1) = "*" Then bString = Mid(bString, 2)
        numList = numList & Left(bString, 1)
    Next getItem
    
    numList = Replace(numList, myNumber, vbNullString)
    
    For verify = 1 To Len(numList)
        
        rNum = Val(Mid(numList, verify, 1))
        rPos = GetPositionOfNumber(rNum)
        If IsPlaying(rPos) = True Then rFinal = rFinal & rNum
    
    Next verify
    If rFinal = "" Then rFinal = "0"
    NumbersPlaying = rFinal
    
End Function

Public Function myNumber() As Integer
    
    Dim tPanel&
    Dim goThrough%, tNum%
    Dim tCaption$

    tPanel = FindWindowEx(tForm2, 0, "TPanel", vbNullString)
    tPanel = GetWindow(tPanel, GW_NEXT)
    
    For goThrough = 1 To 6
        tPanel = GetWindow(tPanel, GW_NEXT)
        tCaption = GetText(tPanel)
        If tCaption <> "Not Playing" Then tNum = Left(tCaption, 1)
        If InStr(tCaption, myNick) Then myNumber = Val(tNum): Exit Function
    Next goThrough

End Function

Public Function IsPlaying(Position As Integer) As Boolean
    
    Dim column%
    Dim sColor$
    
    Dim GRAY As Boolean, BACKGROUND As Boolean
        
    GRAY = False: BACKGROUND = False
    For column = 0 To (12 - 1)
        
        Select Case Position
            Case 1: sColor = GetColor(319 + 8 * column, 144 + 8 * 21)
            Case 2: sColor = GetColor(419 + 8 * column, 144 + 8 * 21)
            Case 3: sColor = GetColor(517 + 8 * column, 144 + 8 * 21)
            Case 4: sColor = GetColor(419 + 8 * column, 342 + 8 * 21)
            Case 5: sColor = GetColor(517 + 8 * column, 342 + 8 * 21)
        End Select
        
        If sColor = OPP_BLOCK Then GRAY = True
        If sColor = OPP_BACK Then BACKGROUND = True

        If GRAY = True And BACKGROUND = True Then Exit For
    Next column
        
    If BACKGROUND = True And GRAY = True Then IsPlaying = True Else IsPlaying = False

End Function

Public Function GetPositionOfNumber(Number As Integer) As Integer
    If Number < myNumber Then GetPositionOfNumber = Number Else GetPositionOfNumber = Number - 1
End Function

Public Function GetNumberofPosition(Position As Integer) As Integer
    Dim orig$, revised$
    
    orig = "123456"
    revised = Replace(orig, myNumber, vbNullString)
    If Position = 0 Then GetNumberofPosition = 0: Exit Function
    GetNumberofPosition = Val(Mid(revised, Position, 1))
    
End Function

Public Function SelfClearNecessary() As Boolean
    
    Dim Row1$, Row2$
    Dim rLeft%
    
    Row1 = AnalyzeRow(1)
    Row2 = AnalyzeRow(2)
    
    For rLeft = 1 To Len(Row1)
        If Mid(Row2, rLeft, 1) = "1" And Mid(Row1, rLeft, 1) = "0" Then
            SelfClearNecessary = True
            Exit Function
        End If
    Next rLeft
    
    SelfClearNecessary = False
End Function

Public Function LowestPlayer(Optional ForSwitch As Boolean = False) As Integer

    Dim eachPlayer%, pHigh%, pRecord%
    Dim pMatrix$, pPlayer%
    
    For eachPlayer = 1 To 5
        pMatrix = OpponentFieldMatrix(eachPlayer)
        
        If pMatrix <> NOT_PLAYING Then
            pHigh = MatrixItemHighest(pMatrix)
            
            If pHigh < pRecord Or pRecord = 0 Then
                pRecord = pHigh
                pPlayer = eachPlayer
            End If
            
        End If
        
    Next eachPlayer
    
    If ForSwitch = True And pRecord > 10 Then pPlayer = 0


    LowestPlayer = GetNumberofPosition(pPlayer)

End Function

Public Function WeaponList() As String
    Dim rEach%
    Dim sColor$, sWeap$
    
    For rEach = 0 To 17
        sColor = GetColor(130 + 16 * rEach, 509)
        
        Select Case sColor
            Case "808000"   'Add Line
                sWeap = sWeap & "A"
            Case "4080FF"   'Clear Special Blocks
                sWeap = sWeap & "B"
            Case "80"       'Quake
                sWeap = sWeap & "Q"
            Case "800080"   'Random Clear Blocks
                sWeap = sWeap & "R"
            Case "FFFF00"   'Switch
                sWeap = sWeap & "S"
            Case "FF"       'Block Bomb
                sWeap = sWeap & "O"
            Case "FFFF"     'Clear Line
                sWeap = sWeap & "C"
            Case "FF00"     'Nuke/Gravity
                sWeap = sWeap & "N"
            Case "0", "C0C0C0"
                Exit For
        End Select
        
    Next rEach
    WeaponList = sWeap
End Function

Public Function CriticalLevel() As Boolean
    
    Dim rows$
    
    rows = AnalyzeRow(14) & AnalyzeRow(15)
    If InStr(rows, "1") Then CriticalLevel = True Else CriticalLevel = False

End Function

Public Function GetQuakable() As Integer
    Dim sPlayers$, sNum$, fMatrix$
    
    Dim getEachPlayer%, eachPlayer%, lastItem%, parse%
    Dim spotSpace%, item%, fMaxDiff%, fPosition%
    
        sPlayers = NumbersPlaying
        For getEachPlayer = 1 To Len(sPlayers)
            
            sNum = Mid(sPlayers, getEachPlayer, 1)
            eachPlayer = GetPositionOfNumber(Val(sNum))
            
            If IsPlaying(eachPlayer) Then
                fMatrix = OpponentFieldMatrix(eachPlayer)
                lastItem = 0
                parse = 0
                
                Do: DoEvents
                        spotSpace = InStr(parse + 1, fMatrix, Chr(32))
                        If spotSpace = 0 Then Exit Do
                        item = Val(Mid(fMatrix, parse + 1, spotSpace - parse - 1))
                        If lastItem = 0 Then lastItem = item
                        If Abs(lastItem - item) >= fMaxDiff Then
                            fMaxDiff = Abs(lastItem - item)
                            fPosition = eachPlayer
                        End If
                        lastItem = item
                        parse = spotSpace
                Loop Until spotSpace = Len(fMatrix) + 1
            
            End If
        Next getEachPlayer
        If fMaxDiff >= 4 Then GetQuakable = GetNumberofPosition(fPosition)
End Function

Public Sub UseWeapon()

    Dim wList$
    Dim rHisHeight%, rMyA%

    wList = WeaponList
    If CriticalLevel = True And Left(wList, 1) <> vbNullString Then
        'If we have dont have an S or an N then...
        If InStr(wList, "S") = 0 And InStr(wList, "N") = 0 Then
            
            Do: DoEvents
                Select Case Left(WeaponList, 1)
                    Case "C"
                        PressKey Asc(myNumber)
                    Case Else
                        If LowestPlayer <> 0 Then
                            PressKey Asc(LowestPlayer)
                        Else
                            PressKey Asc("D")
                        End If
               End Select
               
               Timeout 0.001
               
            Loop Until WeaponList = ""
        
        Else
        
            Do: DoEvents
                
                Select Case Left(WeaponList, 1)
                    Case "S"
                        If LowestPlayer <> 0 Then
                            PressKey Asc(LowestPlayer)
                        Else
                            PressKey Asc("D")
                        End If
                        Exit Sub
                    Case "N"
                        PressKey Asc(myNumber)
                        Exit Sub
                    Case Else
                        If LowestPlayer <> 0 Then
                            PressKey Asc(LowestPlayer)
                        Else
                            PressKey Asc("D")
                        End If
               End Select
                Timeout 0.001
            Loop Until WeaponList = ""
        
        End If
    End If
    
    'If only 1 other player is playing and my weapon list
    'contains a sufficient # of A's to destroy him, then
    'use them
    
    If Len(NumbersPlaying) = 1 And NumbersPlaying <> 0 Then
       
       If InStr(WeaponList, "A") Then
         
         rHisHeight = MatrixItemHighest(OpponentFieldMatrix(GetPositionOfNumber(NumbersPlaying)))
         rMyA = CountChr(WeaponList, "A")

         If (20 - rHisHeight) <= rMyA Or Len(wList) = 18 Then

              Do: DoEvents
                  Select Case Left(WeaponList, 1)
                      Case "C", "N", "S"
                          PressKey Asc(myNumber)
                      Case Else
                          PressKey Asc(NumbersPlaying)
                  End Select
                  Timeout 0.05
                  
              Loop Until WeaponList = ""
              Exit Sub
          End If
        End If
    End If

    'How to utilize all of the specials
    
    Dim sLetter$
    Dim PlayerPos%, PlayerPosO%
    Dim PlayerQuake%, moveRight%
    
    sLetter = Left(wList, 1)
    Select Case sLetter
        Case "B"
            If sLetter = "B" And Mid(wList, 2, 1) = "O" Then
                PlayerPos = AnalyzeWRTBQR
                PlayerPosO = AnalyzeWRTO
                If PlayerPosO = PlayerPos And PlayerPos <> 0 Then
                    PressKey Asc("D")
                    Exit Sub
                End If
            End If
            PlayerPos = AnalyzeWRTBQR
            If PlayerPos <> 0 Then PressKey Asc(PlayerPos): Timeout 0.001
        
        Case "R"
            PlayerPos = AnalyzeWRTBQR
            If PlayerPos <> 0 Then
                Do: DoEvents
                    PressKey Asc(PlayerPos)
                    Timeout 0.001
                Loop Until Left(WeaponList, 1) <> "R"
            End If
        
        Case "Q"
            PlayerPos = AnalyzeWRTBQR
            PlayerQuake = GetQuakable
            If PlayerPos <> 0 And PlayerQuake = 0 Then
                PressKey Asc(PlayerPos)
                Exit Sub
            End If
            If PlayerPos = 0 And PlayerQuake <> 0 Then
                PressKey Asc(PlayerQuake)
                Exit Sub
            End If
            If PlayerPos = 0 And PlayerQuake = 0 Then
                Exit Sub
            End If
            If PlayerPos <> 0 And PlayerQuake <> 0 Then
                Randomize Timer
                If Int(Rnd * 2) + 1 = 1 Then
                    PressKey Asc(PlayerPos)
                Else
                    PressKey Asc(PlayerQuake)
                End If
            End If
        Case "C"
            PlayerPos = AnalyzeWRTC
            Select Case PlayerPos
                Case 7
                    PressKey Asc(myNumber)
                Case 0
                    If Replace(wList, "C", vbNullString) <> vbNullString Then
                        PressKey Asc("D")
                    End If
                    Exit Sub
                Case Else
                    PressKey Asc(PlayerPos)
            End Select
        Case "A"
            PlayerPos = AnalyzeWRTA
            If PlayerPos <> 0 Then
                Do: DoEvents
                    PressKey Asc(PlayerPos)
                    If Left(WeaponList, 1) <> "A" Then Exit Sub Else Timeout 0.01
                Loop
            End If
        
        Case "O"
            PlayerPos = AnalyzeWRTO
            If PlayerPos <> 0 Then
                PressKey Asc(PlayerPos)
                Exit Sub
            End If
            If Left(WeaponList, 1) = "O" And InStr(Mid(WeaponList, 2), "O") Then
                PressKey Asc("D")
                Exit Sub
            End If
            If Len(WeaponList) >= 10 And Left(WeaponList, 1) = "O" Then
                PressKey Asc("D")
                Exit Sub
            End If
            If Len(NumbersPlaying) = 1 And Replace(WeaponList, "O", vbNullString) <> vbNullString Then
                PressKey Asc("D")
                Exit Sub
            End If

        Case "N", "G"
            If CriticalLevel = True Then PressKey Asc(myNumber)

        Case "S"
            PlayerPos = LowestPlayer(True)
            If PlayerPos <> 0 Then
                Do While CriticalLevel = False
                    PressKey VK_DOWN
                    Timeout 0.001
                    For moveRight = 1 To 3
                        PressKey VK_RIGHT
                        Timeout 0.001
                    Next moveRight
                    PressKey VK_SPACE
                    Timeout 1.1
                Loop
                PressKey Asc(PlayerPos)
                Timeout 3
            End If
     End Select
End Sub

Public Function SpecialsInRange() As String
    
    Dim playerMatrix$, sWeap$, sColor$
    Dim cLow%, toLeft%, upDown%, x%, y%
    
    playerMatrix = MatrixFieldSkilled
    cLow = MatrixItemLowest(playerMatrix)
    For toLeft = 1 To 12
        For upDown = 1 To 18 - cLow
            x = 6 + 16 * toLeft
            y = 198 + 16 * upDown
            sColor = GetColor(x, y)
            Select Case sColor
                Case "808000"   'Add Line
                    sWeap = sWeap & "A"
                Case "4080FF"   'Clear Special Blocks
                    sWeap = sWeap & "B"
                Case "80"       'Quake
                    sWeap = sWeap & "Q"
                Case "800080"   'Random Clear Blocks
                    sWeap = sWeap & "R"
                Case "FFFF00"   'Switch
                    sWeap = sWeap & "S"
                Case "FF"       'Block Bomb
                    sWeap = sWeap & "O"
                Case "FFFF"     'Clear Line
                    sWeap = sWeap & "C"
                Case "FF00"     'Nuke/Gravity
                    sWeap = sWeap & "N"
            End Select
        Next upDown
    Next toLeft
    SpecialsInRange = sWeap
End Function
