VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tetris AI by Matt Mazur"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   Icon            =   "frmTAI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   45
      Picture         =   "frmTAI.frx":08CA
      ScaleHeight     =   375
      ScaleWidth      =   1830
      TabIndex        =   31
      Top             =   5760
      Width           =   1830
   End
   Begin VB.PictureBox picBPMBack 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   90
      ScaleHeight     =   375
      ScaleWidth      =   3705
      TabIndex        =   28
      Top             =   6165
      Width           =   3700
      Begin VB.PictureBox picBPM 
         BackColor       =   &H00D57D5A&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   3375
         TabIndex        =   29
         Top             =   0
         Width           =   3375
         Begin VB.Label lblDispBPM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 BPM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   240
            Left            =   2655
            TabIndex        =   30
            Top             =   90
            Width           =   675
         End
      End
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   135
      Picture         =   "frmTAI.frx":2CFC
      ScaleHeight     =   285
      ScaleWidth      =   1875
      TabIndex        =   13
      Top             =   2700
      Width           =   1875
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1905
      Left            =   45
      TabIndex        =   14
      Top             =   2745
      Width           =   3750
      Begin VB.Label lblSticks 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 (0%)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2600
         TabIndex        =   26
         Top             =   990
         Width           =   555
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sticks:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   405
         TabIndex        =   25
         Top             =   990
         Width           =   585
      End
      Begin VB.Label lblAPM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2600
         TabIndex        =   24
         Top             =   765
         Width           =   105
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lines APM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   405
         TabIndex        =   23
         Top             =   765
         Width           =   960
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seconds Playing:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   405
         TabIndex        =   22
         Top             =   1215
         Width           =   1575
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Blocks Dropped:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   405
         TabIndex        =   21
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label lblSecondsPlaying 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2600
         TabIndex        =   20
         Top             =   1215
         Width           =   105
      End
      Begin VB.Label lblBlocksDropped 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2600
         TabIndex        =   19
         Top             =   1440
         Width           =   105
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instantaneous BPM:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   405
         TabIndex        =   18
         Top             =   540
         Width           =   1770
      End
      Begin VB.Label lbliBPM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2600
         TabIndex        =   17
         Top             =   540
         Width           =   105
      End
      Begin VB.Label lblBPM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2600
         TabIndex        =   16
         Top             =   315
         Width           =   105
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Average BPM:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   405
         TabIndex        =   15
         Top             =   315
         Width           =   1305
      End
   End
   Begin VB.CheckBox chkOthers 
      BackColor       =   &H00808080&
      Caption         =   " With Opponents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   1710
      TabIndex        =   8
      Top             =   5310
      Value           =   1  'Checked
      Width           =   2130
   End
   Begin RichTextLib.RichTextBox rtbStatus 
      Height          =   960
      Left            =   90
      TabIndex        =   6
      Top             =   6615
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   1693
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      TextRTF         =   $"frmTAI.frx":4926
   End
   Begin VB.Timer tmrMonitor 
      Interval        =   1
      Left            =   135
      Top             =   7785
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BiG TeStEr"
      Height          =   465
      Left            =   2520
      TabIndex        =   5
      Top             =   7740
      Width           =   1230
   End
   Begin VB.CheckBox chkRun 
      BackColor       =   &H00808080&
      Caption         =   "RUN PAWN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   1710
      TabIndex        =   3
      Top             =   4905
      Width           =   2085
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   90
      TabIndex        =   0
      Top             =   945
      Width           =   3750
      Begin VB.PictureBox picStandard 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   720
         Picture         =   "frmTAI.frx":49A8
         ScaleHeight     =   345
         ScaleWidth      =   2910
         TabIndex        =   12
         Top             =   900
         Width           =   2910
      End
      Begin VB.PictureBox picSkilled 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   720
         Picture         =   "frmTAI.frx":7E62
         ScaleHeight     =   330
         ScaleWidth      =   2055
         TabIndex        =   11
         Top             =   450
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   135
         Picture         =   "frmTAI.frx":A20C
         ScaleHeight     =   285
         ScaleWidth      =   1815
         TabIndex        =   10
         Top             =   45
         Width           =   1815
      End
      Begin VB.OptionButton optnStandard 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   450
         TabIndex        =   2
         Top             =   945
         Width           =   240
      End
      Begin VB.OptionButton optnSkilled 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   450
         TabIndex        =   1
         Top             =   495
         Value           =   -1  'True
         Width           =   240
      End
   End
   Begin VB.Timer tmrRun 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   585
      Top             =   7785
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H8000000E&
      Height          =   2580
      Left            =   0
      ScaleHeight     =   2520
      ScaleWidth      =   3855
      TabIndex        =   9
      Top             =   0
      Width           =   3915
      Begin VB.Image Image2 
         Height          =   780
         Left            =   495
         Picture         =   "frmTAI.frx":BD52
         Top             =   45
         Width           =   2835
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Height          =   915
      Left            =   0
      TabIndex        =   27
      Top             =   4770
      Width           =   3885
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Height          =   2220
      Left            =   0
      TabIndex        =   7
      Top             =   2565
      Width           =   3885
   End
   Begin VB.Label lblStart 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   1170
      TabIndex        =   4
      Top             =   7875
      Width           =   90
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -------------------------------------------------------------------------
' Name:     PAWN, TetriNET Bot
' Author:   Matt Mazur
' Website:  http://www.mattmazur.com
' -------------------------------------------------------------------------

Private Sub chkRun_Click()
   
    If chkRun.Value = 1 Then
        If GameOn = False Then
            MsgBox "Tetrinet game is not initiated.", vbCritical, "Error"
            chkRun.Value = 0
            Exit Sub
        Else
            'Initiation procedures
            SetForm2
            Timeout 0.5
            lblBlocksDropped.Caption = 0
            lblStart.Caption = GetTickCount
            lblSticks.Caption = "0 (0%)"
            SendMessageByString FindWindowEx(tForm2, 0, "TSRichEdit", vbNullString), WM_SETTEXT, 0, vbNullString
            tmrRun.Enabled = True
        End If
    Else
        tmrRun.Enabled = False
        UpdateStatus &HC95331, "PAWN shut off " & Time & vbCrLf
    End If
    
End Sub

Private Sub Command1_Click()
    SetForm2
    Timeout 0.1
    Dim x$
DropNormal

    MsgBox tButton
End Sub

Private Sub Form_Load()
Dim doNoDelay%
    If tForm1 = 0 Then
        dNoDelay% = MsgBox("Load TetriNET without delay?", vbYesNoCancel + vbQuestion, "Load TetriNET")
        If dNoDelay = vbYes Then
            Shell "C:\Program Files\Microsoft Visual Studio\VB98\TAI Alpha\zeroTetrinet Plus.exe"
        ElseIf dNoDelay = vbNo Then
            Shell "C:\Program Files\Microsoft Visual Studio\VB98\TAI Alpha\Tetrinet - Morpher - 1sec Delay.exe"
        End If
        
        UpdateStatus &H808080, "Loaded Tetrinet v1.13 ", True
        UpdateStatus vbBlack, Time & vbCrLf, False
    
    End If
    
    SetForm2
    Me.Top = 200
    Me.Left = Screen.Width - Me.Width
    
    UpdateStatus &HD57D5A, "Loaded PAWN ", True
    UpdateStatus vbBlack, Time & vbCrLf, False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub optnSkilled_Click()
    frmNormDisplay.Visible = False
End Sub

Private Sub optnStandard_Click()

    frmNormDisplay.Show
    frmNormDisplay.Top = Me.Top + Me.Height
    frmNormDisplay.Left = Screen.Width - frmNormDisplay.Width
End Sub

Private Sub picSkilled_Click()
optnSkilled.Value = 1
optnStandard.Value = 0
End Sub

Private Sub picStandard_Click()
optnSkilled.Value = 0
optnStandard.Value = 1
End Sub

Private Sub slideBPM_Click()
lblDispBPM.Caption = slideBPM.Value
End Sub

Private Sub tmrMonitor_Timer()
    Static gOn As Boolean, notFirstTime As Boolean
    'Dim gOn As Boolean, gOnNow As Boolean
    If notFirstTime = False Then
        notFirstTime = True
    Else
        gOnNow = GameOn
        If GameOn <> gOn Then
            If gOnNow = False And gOn = True Then
                UpdateStatus vbRed, "Game Ended ", True
                UpdateStatus vbBlack, Time & vbCrLf, False
            ElseIf gOnNow = True And gOn = False Then
                UpdateStatus vbRed, "Game Started ", True
                UpdateStatus vbBlack, Time & vbCrLf, False
            End If
            gOn = GameOn
        End If
    End If
    
End Sub



Private Sub tmrRun_Timer()
    Dim skilledSuccess As Boolean
    Static prevBPM&
    Dim nowBPM&, diffBPM&, iBPM%
    Dim timeDiff, avgBPM%
    
    'Check for any errors
        If GetForegroundWindow <> tForm2 Then
            UpdateStatus vbBlack, "Lost focus on TForm2 " & Time & vbCrLf
            chkRun.Value = 0
            Exit Sub
        End If
        If GameOn = False Then
           chkRun.Value = False
           Exit Sub
        End If
    
    'Drop the piece
        If optnSkilled.Value = True Then
            FixIfDied
            skilledSuccess = DropSkilled
            If skilledSuccess = False Then Exit Sub
        ElseIf optnStandard.Value = True Then
            DropNormal
        End If
        If chkOthers.Value = 1 Then UseWeapon
    
    'Instantaneous BPM
        If prevBPM <> 0 Then
            nowBPM = GetTickCount
            diffBPM = (nowBPM - prevBPM)
            iBPM = (60 / diffBPM) * 1000
            lbliBPM.Caption = iBPM
        End If
        prevBPM = GetTickCount
    
    'Average BPM
        lblBlocksDropped.Caption = Val(lblBlocksDropped.Caption) + 1
        timeDiff = (GetTickCount - lblStart.Caption) / 1000
        lblSecondsPlaying.Caption = timeDiff & "s"
        If timeDiff <> 0 Then avgBPM = (lblBlocksDropped.Caption / timeDiff) * 60
        lblBPM.Caption = avgBPM
        picBPM.Width = Int(3700 * (avgBPM / 250))
        lblDispBPM.Caption = avgBPM & " BPM"
        lblDispBPM.Left = picBPM.Width - lblDispBPM.Width - 25
    'APM
        lblAPM.Caption = APM
End Sub


