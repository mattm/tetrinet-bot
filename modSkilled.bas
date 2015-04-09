Attribute VB_Name = "modSkilled"
Option Explicit



Public Function MatrixItemLowest(Matrix As String) As Long
    Dim Search&, spotSpace&, parse&, item&, Lowest&
    parse = 0
    Lowest = 23
    Do: DoEvents
            spotSpace = InStr(parse + 1, Matrix, Chr(32))
            If spotSpace = 0 Then Exit Do
            item = Val(Mid(Matrix, parse + 1, spotSpace - parse - 1))
            If item < Lowest Then Lowest = item
            parse = spotSpace
    Loop Until spotSpace = Len(Matrix) + 1
    MatrixItemLowest = Lowest
End Function

Public Function MatrixItemHighest(Matrix As String) As Long
    Dim Search&, spotSpace&, parse&, item&, Highest&
    parse = 0
    Highest = -1
    Do: DoEvents
            spotSpace = InStr(parse + 1, Matrix, Chr(32))
            If spotSpace = 0 Then Exit Do
            item = Val(Mid(Matrix, parse + 1, spotSpace - parse - 1))
            If item > Highest Then Highest = item
            parse = spotSpace
    Loop Until spotSpace = Len(Matrix) + 1
    MatrixItemHighest = Highest
End Function

Public Function MatrixSubtract(MatrixOrig$, offSet&)
    Dim Search&, spotSpace&, parse&, item&, NewString$
    parse = 0
    Do: DoEvents
            spotSpace = InStr(parse + 1, MatrixOrig, Chr(32))
            If spotSpace = 0 Then Exit Do
            item = Val(Mid(MatrixOrig, parse + 1, spotSpace - parse - 1))
            NewString = NewString & item - offSet & Chr(32)
            parse = spotSpace
    Loop Until spotSpace = Len(MatrixOrig) + 1
    MatrixSubtract = Chr(32) & NewString
End Function

Public Function MatrixFieldSkilled() As String
    Dim toLeft%, upDown%, xCoor%, yCoor%, columnHeight%
    Dim heights$
    
    For toLeft = 1 To 12
        For upDown = 1 To 22
            xCoor = 6 + 16 * toLeft
            yCoor = 198 + 16 * upDown
            If GetColor(xCoor, yCoor) <> "FFFFFF" Then
                columnHeight = 19 - upDown
                heights = heights & columnHeight & " "
                Exit For
            End If
        Next upDown
    Next toLeft
    MatrixFieldSkilled = heights
End Function

Public Function MatrixBlockSkilled() As String
    
    Dim toLeft%, upDown%, xCoor%, yCoor%
    Dim hLatest%, hGreatest%
    
    Dim rStart%, spot%
    Dim mHeights$, mChr$, mFinal$
    
    For toLeft = 0 To 3
        
        hLatest = 0
        
        For upDown = 0 To 3
            xCoor = 86 + 16 * toLeft
            yCoor = 150 + 16 * upDown
            If GetColor(xCoor, yCoor) <> "FFFFFF" Then
                hLatest = upDown + 1
                If hLatest > hGreatest Then hGreatest = hLatest
            End If
         Next upDown
    
        If hLatest <> 0 Then mHeights = mHeights & hLatest & Chr(32)
    
    Next toLeft
    
    
    If mHeights$ = vbNullString Then
        MatrixBlockSkilled = vbNullString
    Else
    
        rStart = 1
        Do: DoEvents
            spot = InStr(rStart, mHeights, " ")
            mChr = Mid(mHeights, rStart, spot - rStart)
            mFinal$ = mFinal$ & hGreatest - Val(mChr$) & Chr(32)
            rStart = spot + 1
        Loop Until rStart > Len(mHeights$)
                
        MatrixBlockSkilled = Chr(32) & mFinal$
    End If
End Function

Public Function AnalyzeRow(Row As Integer) As String

    Dim toLeft%, xCoor%, yCoor%
    Dim HorizonBottom$
    
    For toLeft = 1 To 12
            xCoor = 6 + 16 * toLeft
            yCoor = 198 + 16 * (19 - Row)
            If GetColor(xCoor, yCoor) <> "FFFFFF" Then HorizonBottom = HorizonBottom & "1" Else HorizonBottom = HorizonBottom & "0"
    Next toLeft
    
    AnalyzeRow = HorizonBottom

End Function

Public Function DropSkilled() As Boolean


Dim fLow%, fHigh%, dLevel%, Position%, rColumns%, offSet%, rotate%, rotations%
Dim mSkilled$, mField$, Block$, HorizonBottom$, rCheck$, inRange$


mSkilled = MatrixFieldSkilled
fLow = MatrixItemLowest(mSkilled)
fHigh = MatrixItemHighest(mSkilled)
HorizonBottom = AnalyzeRow(fLow)

Do: DoEvents
    Block = MatrixBlockSkilled
    If Block = "" Then PressKey VK_DOWN: Timeout 0.001
Loop Until Block <> ""

DropSkilled = False

    If frmMain.chkOthers.Value = 0 Then
        Select Case fHigh
            Case Is < 8: dLevel = 6
            Case Is < 11: dLevel = 4
            Case Is < 13: dLevel = 2
            Case Else: dLevel = 0
        End Select
    Else
        Select Case fHigh
            Case Is < 8: dLevel = 3
            Case Is < 11: dLevel = 2
            Case Is < 13: dLevel = 1
            Case Else: dLevel = 0
        End Select
        inRange = SpecialsInRange
        If InStr(inRange, "N") Or InStr(inRange, "S") Or InStr(inRange, "O") Then dLevel = 0
    End If

Select Case MatrixBlock
    Case "1111" & vbCrLf, "1" & vbCrLf & "1" & vbCrLf & "1" & vbCrLf & "1" & vbCrLf: rotations = 2
    Case "011" & vbCrLf & "110", "10" & vbCrLf & "11" & vbCrLf & "01", "01" & vbCrLf & "11" & vbCrLf & "10", "110" & vbCrLf & "011": rotations = 2
    Case "11" & vbCrLf & "11": rotations = 1
    Case Else: rotations = 4
End Select


For offSet = 0 To dLevel
    mField = MatrixSubtract(mSkilled, fLow + offSet)
    
    For rotate = 1 To rotations
        UpdateSeconds
        If InStr(mField, Block) Then
            Position = InStr(mField, Block)
            rColumns = CountChr(Left(mField, Position), Chr(32))
                rCheck = Mid(HorizonBottom, rColumns, Len(Replace(Block, " ", vbNullString)))

                If InStr(rCheck$, "0") Then
                    Exit For
                Else

                     DropBlockSkilled mField, Block
                     
                     Timeout 0.001
                     DropSkilled = True
                     Exit Function
                End If
        End If
        PressKey VK_UP
        Timeout 0.001
        Block = MatrixBlockSkilled
    Next rotate
Next offSet
PressKey VK_DOWN

End Function

Public Sub DropBlockSkilled(mField As String, mBlock As String)
    
    Dim rSpot%, SpacesOver%, setLeft%, setRight%
    Dim editedString$
    
    rSpot = InStr(mField, mBlock)
    editedString$ = Left(mField, rSpot)
    SpacesOver = CountChr(editedString, Chr(32)) - 1
    
    CheckSticks mBlock
    
    For setLeft = 1 To 8
        PressKey VK_LEFT
        Timeout 0.0001
    Next setLeft
    For setRight = 1 To SpacesOver
        PressKey VK_RIGHT
        Timeout 0.0001
    Next setRight
   
    PressKey VK_SPACE

End Sub

