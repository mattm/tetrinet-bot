VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmNormDisplay 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Standard Play Choice Method List"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   1395
      TabIndex        =   20
      Top             =   3735
      Width           =   1185
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   90
      Picture         =   "frmNormDisplay.frx":0000
      ScaleHeight     =   525
      ScaleWidth      =   4215
      TabIndex        =   19
      Top             =   0
      Width           =   4215
   End
   Begin RichTextLib.RichTextBox rtfInfo 
      Height          =   1365
      Left            =   1845
      TabIndex        =   18
      Top             =   540
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   2408
      _Version        =   393217
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmNormDisplay.frx":73A6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   0
      Left            =   135
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   17
      Top             =   675
      Width           =   400
   End
   Begin VB.PictureBox picBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   1
      Left            =   510
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   16
      Top             =   675
      Width           =   400
   End
   Begin VB.PictureBox picBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   2
      Left            =   885
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   15
      Top             =   675
      Width           =   400
   End
   Begin VB.PictureBox picBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   3
      Left            =   1260
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   14
      Top             =   675
      Width           =   400
   End
   Begin VB.PictureBox picBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   4
      Left            =   135
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   13
      Top             =   1050
      Width           =   400
   End
   Begin VB.PictureBox picBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   5
      Left            =   510
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   12
      Top             =   1050
      Width           =   400
   End
   Begin VB.PictureBox picBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   6
      Left            =   885
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   11
      Top             =   1050
      Width           =   400
   End
   Begin VB.PictureBox picBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   7
      Left            =   1260
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   10
      Top             =   1050
      Width           =   400
   End
   Begin VB.PictureBox picBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   8
      Left            =   135
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   9
      Top             =   1425
      Width           =   400
   End
   Begin VB.PictureBox picBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   9
      Left            =   510
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   8
      Top             =   1425
      Width           =   400
   End
   Begin VB.PictureBox picBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   10
      Left            =   885
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   7
      Top             =   1425
      Width           =   400
   End
   Begin VB.PictureBox picBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   11
      Left            =   1260
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   6
      Top             =   1425
      Width           =   400
   End
   Begin VB.PictureBox picBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   12
      Left            =   135
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   5
      Top             =   1800
      Width           =   400
   End
   Begin VB.PictureBox picBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   13
      Left            =   510
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   1800
      Width           =   400
   End
   Begin VB.PictureBox picBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   14
      Left            =   885
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   3
      Top             =   1800
      Width           =   400
   End
   Begin VB.PictureBox picBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   15
      Left            =   1260
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   2
      Top             =   1800
      Width           =   400
   End
   Begin VB.ListBox lstBlocks 
      Height          =   840
      Left            =   3510
      TabIndex        =   1
      Top             =   3600
      Width           =   1230
   End
   Begin VB.ListBox lstPositions 
      BackColor       =   &H00F7B371&
      Height          =   840
      Left            =   45
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   2655
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   0
      Top             =   540
      Width           =   1815
   End
End
Attribute VB_Name = "frmNormDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
