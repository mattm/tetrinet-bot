Attribute VB_Name = "modNormal"
Option Explicit

Public Function MatrixBlock() As String
 
Dim getColumn%, rowCount%, blockHeight%
Dim eachRow%

Dim Block$, thisRow$, newRow$, bColumn$
Dim newMatrix$, lastColumn$

    For getColumn = 1 To 4
        bColumn = BlockColumn(getColumn)
        If InStr(bColumn, "1") Then Block = Block & bColumn & vbCrLf
    Next getColumn

    rowCount = MatrixRowCount(Block)
    If rowCount = 0 Then MatrixBlock = vbNullString: Exit Function

    'Trim the right columns of zeros of the block
    
    Do: DoEvents
        newMatrix = vbNullString
        blockHeight = InStr(Block, vbCrLf) - 1
        lastColumn = MatrixColumn(Block, blockHeight)
    
        If InStr(lastColumn, "1") = 0 Then
            For eachRow = 1 To rowCount
                thisRow = MatrixRow(Block, eachRow)
                newRow = Left(thisRow, Len(thisRow) - 1)
                newMatrix = newMatrix & newRow & vbCrLf
            Next eachRow
        Else
            Exit Do
        End If
        Block = newMatrix
    Loop
    
    MatrixBlock = Block

End Function

Public Function BlockColumn(column As Integer) As String
    Dim mRow%, xPos%, yPos%
    Dim blockColor$
    For mRow = 1 To 4
        xPos = 6 + 16 * (column + 4)
        yPos = 134 + 16 * mRow
        blockColor = GetColor(xPos, yPos)
        If blockColor = "FFFFFF" Then
            BlockColumn = BlockColumn & "0"
        Else
            BlockColumn = BlockColumn & "1"
        End If
    Next mRow
End Function

Public Function MatrixField() As String

    Dim toLeft%, upDown%, xPos%, yPos%
    
    Dim blockColor$, rBin$

    For toLeft = 1 To 12
        For upDown = 1 To 18
            xPos = 6 + 16 * toLeft
            yPos = 198 + 16 * upDown
            blockColor$ = GetColor(xPos, yPos)
            Select Case blockColor$
                Case "FFFFFF": rBin$ = rBin$ & "0"
                Case Else: rBin$ = rBin$ & "1"
            End Select
        Next upDown
    rBin$ = rBin$ & vbCrLf
    Next toLeft
    MatrixField = rBin$

End Function

Public Function MatrixRowCount(matrixSource As String) As Integer
    
    MatrixRowCount = CountChr(matrixSource, Chr(13))

End Function

Public Function MatrixRow(matrixSource As String, rowNum As Integer) As String
    
    Dim posParse%, rowTimes%, spot%
    
    posParse = 1
    rowTimes = 0
    Do: DoEvents
        
        spot = InStr(posParse, matrixSource, vbCrLf)
        If spot = 0 Then Exit Do
        rowTimes = rowTimes + 1
        If rowTimes = rowNum Then
            MatrixRow = Mid(matrixSource, posParse, spot - posParse)
        Else
            posParse = spot + 2
        End If
    
    Loop

End Function

Public Function MatrixColumn(matrixSource As String, columnNum As Integer) As String

    Dim parse%
    Dim matrixNew$
    
    For parse = 1 To MatrixRowCount(matrixSource)
        matrixNew = matrixNew & Mid(MatrixRow(matrixSource, parse), columnNum, 1) & vbCrLf
    Next parse
    
    MatrixColumn = matrixNew

End Function

Public Function FieldColumnHeight(columnNum As Integer) As Integer
    
    Dim upDown%, xPos%, yPos%
    Dim blockColor$
    
    For upDown = 1 To 18
        xPos = 6 + 16 * columnNum
        yPos = 198 + 16 * upDown
        blockColor$ = GetColor(xPos, yPos)
        If blockColor <> "FFFFFF" Then FieldColumnHeight = 18 - upDown + 1: Exit Function
    Next upDown
    
End Function

Public Sub DropBlockNormal()
    Dim tAnalyze$, tLeft$, tPiece$, modifiedMatrix$
    Dim spot1%, spot2%, rotatePiece%, oLeft%, oRight%
    
    tAnalyze = frmNormDisplay.lstPositions.List(0)
    spot1 = InStr(tAnalyze, "L:")
    spot2 = InStr(spot1 + 1, tAnalyze, "P:")
    tLeft = Mid(tAnalyze, spot1 + 3, 2)
    tPiece = Mid(tAnalyze, spot2 + 3)
    CheckSticks tPiece
    
    For rotatePiece = 1 To 4
        PressKey VK_UP
        Timeout 0.01
        modifiedMatrix = Replace(MatrixBlock, "1", "P")
        modifiedMatrix = Replace(modifiedMatrix, "0", "H")
        If modifiedMatrix = tPiece Then
            For oLeft = 1 To 6
                PressKey VK_LEFT
                Timeout 0.0001
            Next oLeft
            For oRight = 1 To tLeft - 1
                PressKey VK_RIGHT
                Timeout 0.0001
            Next oRight
            PressKey VK_SPACE
            Exit For
        End If
    Next rotatePiece
End Sub


Public Function AnalyzeField(matrixSource As String, Block As String, Position As Integer) As String
    
    Dim blockLength%, blockHeight%, analyzeRows%
    Dim bHoles%, range2%, bTotalHoles%, lastP%, bTop%, bBottom%
    Dim p2%, p1%, p0%, a0%, a1%, a2%
    Dim min%, max%, analyzeSection%, analyzeCheck%, bLinesCleared%
    
    Dim mRow$, thisCol$, beforeCol$, afterCol$, allNew$, tEntry$

    Dim createsHole As Boolean, clearsLines As Boolean, isProblem As Boolean

    blockLength = MatrixRowCount(Block)
    blockHeight = InStr(Block, Chr(13)) - 1
    For analyzeRows = Position To (Position + blockLength - 1)
        mRow = MatrixRow(matrixSource, analyzeRows)
        If InStr(mRow, "P0") Then
            bHoles = bHoles + 1
            range2 = InStr(mRow, "01")
            If range2 = 0 Then range2 = Len(mRow)
            bTotalHoles = bTotalHoles + (range2 - InStr(mRow, "P0"))
        End If
        lastP = Len(mRow) - InStr(mRow, "P") + 1
        If lastP > bTop Then bTop = lastP
    Next analyzeRows
    bBottom = bTop - blockHeight + 1
    
    If bTop >= 19 Then
        AnalyzeField = "mxError: reached max analyzation level"
        Exit Function
    End If
    
    'p = previous, a = after
    p2 = FieldColumnHeight(Position - 2)
    p1 = FieldColumnHeight(Position - 1)
    p0 = 19 - InStr(MatrixRow(matrixSource, Position), "P")
    a0 = 19 - InStr(MatrixRow(matrixSource, Position + blockLength - 1), "P")
    a1 = FieldColumnHeight(Position + blockLength)
    a2 = FieldColumnHeight(Position + blockLength + 1)
             
    createsHole = False
    If (p2 - p1) >= 3 And (p0 - p1) >= 3 Then
        createsHole = True
    End If

    If (a2 - a1) >= 3 And (a0 - a1) >= 3 Then
        createsHole = True
    End If
            
    clearsLines = False
    min = 22: max = 0
    For analyzeSection = (Val(bTop) - blockHeight + 1) To bTop
        thisCol = MatrixColumn(matrixSource, 18 - analyzeSection + 1)
        If InStr(thisCol, "0") = 0 Then
            If analyzeSection < min Then min = analyzeSection
            If analyzeSection > max Then max = analyzeSection
            clearsLines = True
        End If
    Next analyzeSection
    bLinesCleared = max - min + 1

    If bLinesCleared < 0 Then bLinesCleared = 0
    
    isProblem = False
    If clearsLines = True Then
        If bLinesCleared <= 2 Then '(test it with =1)
            For analyzeCheck = Position To (Position + blockLength - 1)
                beforeCol = MatrixRow(matrixSource, analyzeCheck)
                afterCol = "0" & Left(beforeCol, 18 - max) & Mid(beforeCol, 18 - min + 2)
                allNew = allNew & afterCol & vbCrLf
                If InStr(afterCol, "P0") Then isProblem = True
            Next analyzeCheck
        End If
    End If

    'Format everything all nice
            
    Dim tLeft, tHoles, tTotalHoles, tTop, tBottom, tLinesCleared, tCreatesHole
            
    tLeft = Format(Position, "00")
    tHoles = Format(bHoles, "00")
    tTotalHoles = Format(bTotalHoles, "00")
    tTop = Format(bTop, "00")
    tBottom = Format(bTop - InStr(Block, vbCrLf) - 1, "00")
    tEntry$ = "mHole: " & isProblem & " LC: " & 4 - bLinesCleared & _
            " H: " & tHoles & _
            " PRB: " & createsHole & _
            " T: " & tTop & _
            " L: " & tLeft & _
            " P: " & Block
    
    AnalyzeField = tEntry$
End Function

Public Function FallBlock(matrixSource As String, Block As String, Position As Integer) As String
    
    Dim blockLength%
    Dim matrixPrior$, matrixSection$
    Dim matrixFallen$, matrixAfter$, matrixComplete$
    
    blockLength = MatrixRowCount(Block)
    matrixPrior = MatrixPart(matrixSource, 1, Position - 1)
    matrixSection = MatrixPart(matrixSource, Position, Position + blockLength - 1)
    matrixFallen = CompleteFall(matrixSection, Block)
    matrixAfter = MatrixPart(matrixSource, Position + blockLength, 12)
    matrixComplete = matrixPrior & matrixFallen & matrixAfter
    
    FallBlock = matrixComplete
End Function

Public Function CompleteFall(Field As String, mBlock As String) As String

    Dim blockLength%, placePiece%, fallPiece%
    Dim lastP%, makeFit%
    Dim fieldRow$, pieceRow$, mNew$
    Dim mFalling$, mRow$, nRow$, sFinal$
    
    Dim stopFalling As Boolean
    
    'MsgBox Field$
    'Clipboard.SetText Field$
    'MsgBox "Go"
    'Clipboard.SetText mBlock
    'MsgBox "Go2"
    blockLength = MatrixRowCount(mBlock)
    mBlock = Replace(mBlock, "1", "P")
    mBlock = Replace(mBlock, "0", "H")
    Field = Replace(Field, "1", "F")
    Field = Replace(Field$, "0", "H")
    
    For placePiece = 1 To blockLength
        fieldRow = MatrixRow(Field, placePiece)
        pieceRow = MatrixRow(mBlock, placePiece)
        mNew = mNew & pieceRow & fieldRow & vbCrLf
    Next placePiece
    
    If InStr(mNew, "PF") = 0 Then
        
        Do: DoEvents
            mFalling = vbNullString
            
            For fallPiece = 1 To blockLength
                mRow = MatrixRow(mNew, fallPiece)
                lastP = LastOccurance(mRow, "P")
                If Mid(mRow, lastP + 2, 1) = vbNullString Then stopFalling = True
                nRow = "H" & Left(mRow, lastP) & Mid(mRow, lastP + 2)
                mFalling = mFalling & nRow & vbCrLf
            Next fallPiece
            
            mNew = mFalling
        Loop Until InStr(mNew, "PF") Or stopFalling = True
    
    End If

    For makeFit = 1 To MatrixRowCount(mNew$)
        sFinal$ = sFinal$ & Right(MatrixRow(mNew$, makeFit), 18) & vbCrLf
    Next makeFit
    
    sFinal$ = Replace(sFinal$, "H", "0")
    sFinal$ = Replace(sFinal$, "F", "1")
    CompleteFall = sFinal$
End Function

Public Function MatrixPart(matrixSource As String, Optional StartCol = 1, Optional EndCol = 12) As String
    
    Dim getRow%
    Dim mPart$
        
    For getRow = StartCol To EndCol
        mPart = mPart & MatrixRow(matrixSource, getRow) & vbCrLf
    Next getRow
    
    MatrixPart = mPart
End Function

Public Function LastOccurance(Source As String, Character As String) As Integer
    
    Dim searchString%
    Dim thisChr$
    
    For searchString = 1 To Len(Source$)
        thisChr = Mid(Source, searchString, 1)
        If thisChr = Character Then LastOccurance = searchString
    Next searchString
    
End Function

Public Function RotateBlock(Matrix As String, xRotations As Integer) As String

    Dim matrixTest$, bRow$, nRow$
    Dim bLength%, parseString%, parseAgain%

    matrixTest = Matrix
    Select Case xRotations
        Case 1
            RotateBlock = matrixTest
        Case 2
            bLength = InStr(matrixTest, Chr(13)) - 1
            For parseString = bLength To 1 Step -1
                For parseAgain = 1 To MatrixRowCount(matrixTest)
                    bRow = MatrixRow(matrixTest, parseAgain)
                    nRow = nRow & Mid(bRow, parseString, 1)
                Next parseAgain
                nRow = nRow & vbCrLf
            Next parseString
            RotateBlock = nRow
        Case 3
            For parseString = MatrixRowCount(matrixTest) To 1 Step -1
                bRow = MatrixRow(matrixTest, parseString)
                For parseAgain = Len(bRow) To 1 Step -1
                    nRow = nRow & Mid(bRow, parseAgain, 1)
                Next parseAgain
                nRow = nRow & vbCrLf
            Next parseString
            RotateBlock = nRow
        Case 4
            bLength = InStr(matrixTest, Chr(13)) - 1
            For parseAgain = 1 To bLength
                For parseString = MatrixRowCount(matrixTest) To 1 Step -1
                    bRow = MatrixRow(matrixTest, parseString)
                    nRow = nRow & Mid(bRow, parseAgain, 1)
                Next parseString
                nRow = nRow & vbCrLf
            Next parseAgain
            RotateBlock = nRow
    End Select
    
End Function

Public Sub DropNormal()
Dim rBlock%, blockLength%, iPiece%
Dim initial_piece$, Block$, matrixFallen$, tAnalyze$, aInfo$


    If InStr(AnalyzeRow(18), "1") Then
        UpdateStatus vbBlack, "Reached max analyzation level " & Time & vbCrLf
        PressKey VK_SPACE
        PressKey VK_SPACE
        frmMain.chkRun.Value = 0
        Exit Sub
    End If

    With frmNormDisplay
        .lstPositions.Clear
        .lstBlocks.Clear
        
        initial_piece = MatrixBlock
        If initial_piece = vbNullString Then Exit Sub
        
        For rBlock = 1 To 4
            Block = RotateBlock(initial_piece, rBlock)
            If InList(.lstBlocks, Block) = False Then
                .lstBlocks.AddItem Block
                
                blockLength = MatrixRowCount(Block)
                
                For iPiece = 1 To 13 - blockLength
                    matrixFallen = FallBlock(MatrixField, Block, iPiece)
                    tAnalyze = AnalyzeField(matrixFallen, Block, iPiece)
                    If InList(.lstPositions, tAnalyze) = False Then
                        .lstPositions.AddItem tAnalyze
                        .Caption = "Combination list -- Possible combinations [" & .lstPositions.ListCount & "]"
                        If .lstPositions.ListCount / 5 = Int(.lstPositions.ListCount / 5) Then
                            aInfo$ = "Position: " & vbTab & Val(Mid(.lstPositions.List(0), InStr(.lstPositions.List(0), "L:") + 3, 2)) & vbCrLf & _
                                    "Top: " & vbTab & vbTab & Val(Mid(tAnalyze, InStr(tAnalyze, "T:") + 3, 2)) & vbCrLf & _
                                    "Num Holes: " & Val(Mid(tAnalyze, InStr(tAnalyze, "H:") + 3, 2)) & vbCrLf
                            ShowNormDisplay Block, aInfo
                        End If
                    End If
                Next iPiece
            End If
        
        Next rBlock
        
        DropBlockNormal
    End With

End Sub
