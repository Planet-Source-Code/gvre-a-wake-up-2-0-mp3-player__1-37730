VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Wake Up 2.0"
   ClientHeight    =   2775
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4800
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1915.354
   ScaleMode       =   0  'User
   ScaleWidth      =   4507.448
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   120
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3360
      TabIndex        =   0
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ver 2.0 Final"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wake Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Lblemail 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "email: gvre@freemail.gr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   3885
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(c) 2001 by Gvre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   450
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    frmAbout.Picture = Form1.Picture
End Sub
