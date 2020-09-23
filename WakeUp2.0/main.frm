VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wake Up 2.0 Final"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerHide 
      Interval        =   125
      Left            =   2640
      Top             =   6360
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6720
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   800
      ImageHeight     =   600
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":CCEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1BCBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2AF3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3C56E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":45708
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrameLanguage 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Ãëþóóá / Language"
      ForeColor       =   &H8000000A&
      Height          =   735
      Left            =   3480
      TabIndex        =   30
      Top             =   6360
      Width           =   2655
      Begin VB.OptionButton OptLang 
         BackColor       =   &H80000001&
         Caption         =   "ÁããëéêÜ / English"
         ForeColor       =   &H8000000A&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   32
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton OptLang 
         BackColor       =   &H80000001&
         Caption         =   "ÅëëçíéêÜ / Greek"
         ForeColor       =   &H8000000A&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   2415
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3960
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":51C73
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":51DCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5221F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":52671
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtwakeMin 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton CmdStartWake 
      Appearance      =   0  'Flat
      Caption         =   "Åíåñãïðïßçóç Áöýðíéóçò"
      Height          =   485
      Left            =   2790
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1515
   End
   Begin VB.Timer TimerTimes 
      Interval        =   10
      Left            =   5400
      Top             =   1080
   End
   Begin VB.Timer TimerCountSongs 
      Interval        =   100
      Left            =   7200
      Top             =   1080
   End
   Begin VB.Timer Timertag 
      Enabled         =   0   'False
      Interval        =   9000
      Left            =   2880
      Top             =   3240
   End
   Begin VB.OptionButton OptionRandom 
      BackColor       =   &H80000001&
      Caption         =   "Ôõ÷áßá óåéñÜ"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   6840
      Width           =   2175
   End
   Begin VB.OptionButton OptionAll 
      BackColor       =   &H80000001&
      Caption         =   "ÅðáíÜëçøç üëùí"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   6600
      Width           =   2175
   End
   Begin VB.OptionButton OptionCur 
      BackColor       =   &H80000001&
      Caption         =   "ÅðáíÜëçøç ôñÝ÷ïíôïò"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton CmdDown 
      Height          =   375
      Left            =   7560
      Picture         =   "main.frx":52AC3
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6120
      Width           =   375
   End
   Begin VB.CommandButton CmdUp 
      Height          =   375
      Left            =   7560
      Picture         =   "main.frx":52F05
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5640
      Width           =   375
   End
   Begin VB.Timer TimerWake 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   600
   End
   Begin VB.TextBox TxtWakeHour 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   1080
   End
   Begin VB.CommandButton CmdPause 
      Caption         =   "Pause"
      Height          =   255
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "Stop"
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton CmdPlay 
      Caption         =   "Play"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   615
   End
   Begin VB.CheckBox ChkMute 
      Caption         =   "Mute"
      Height          =   255
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   615
   End
   Begin MSComctlLib.Slider SliderVolume 
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   5640
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Min             =   -2000
      Max             =   0
      TickStyle       =   3
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Arrows          =   65536
      Orientation     =   8323073
   End
   Begin VB.Timer TimerFlat 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3360
      Top             =   5160
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   953
      ButtonWidth     =   3281
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Äéá÷åßñéóç Ôñáãïõäéþí"
            ImageIndex      =   3
            Style           =   5
            Object.Width           =   1e-4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Add"
                  Text            =   "ÐñïóèÞêç Ôñáãïõäéïý"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Remove"
                  Text            =   "ÄéáãñáöÞ ÅðéëåãìÝíùí"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "clear"
                  Text            =   "ÄéáãñáöÞ ¼ëùí"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Äéá÷åßñçóç Ëßóôáò"
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "ÁðïèÞêåõóç Ëßóôáò"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "ÁíÜêôçóç Ëßóôáò"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Key             =   "About"
            ImageIndex      =   2
            Object.Width           =   1e-4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "About"
                  Text            =   "About"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "¸îïäïò"
            ImageIndex      =   1
         EndProperty
      EndProperty
      OLEDropMode     =   1
      Begin VB.Timer TimerScrollCaption 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   480
         Top             =   -120
      End
   End
   Begin MSComctlLib.ListView listviewOnAir 
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   8454143
      BackColor       =   7375868
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Song Title"
         Object.Width           =   13671
      EndProperty
   End
   Begin VB.Label LblHearingChoices 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ÅðéëïãÝò Áêñüáóçò"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   6120
      Width           =   2195
   End
   Begin VB.Label LblLang 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ãëþóóá  /  Language"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   3480
      TabIndex        =   33
      Top             =   6120
      Width           =   2685
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6645
      Picture         =   "main.frx":53347
      Top             =   5595
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label LblUpDown 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   345
      Left            =   1920
      TabIndex        =   29
      Top             =   1080
      Width           =   150
   End
   Begin VB.Label LblStartWake 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   2760
      TabIndex        =   28
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "×ñüíïò Áêñüáóçò"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   5640
      TabIndex        =   27
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ÄéÜñêåéá Ôñáãïõäéïý"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   4560
      TabIndex        =   26
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label LblElapsed 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   5640
      TabIndex        =   25
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label LblDur 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   4560
      TabIndex        =   24
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label LabelNumSon 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Áñéèìüò Ôñáãïõäéþí"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   6720
      TabIndex        =   23
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label LblNumSongs 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   6720
      TabIndex        =   22
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "¿ñá Áöýðíéóçò"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   1440
      TabIndex        =   16
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ÔñÝ÷ïõóá ¿ñá"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label LblTime 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label LabelCurSong 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ÔñÝ÷ïí Ôñáãïýäé"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   7815
   End
   Begin VB.Label LblCurSong 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   7815
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6720
      Picture         =   "main.frx":53789
      Top             =   5520
      Width           =   480
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   255
      Left            =   7560
      TabIndex        =   5
      Top             =   6720
      Visible         =   0   'False
      Width           =   495
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim song As String
Dim pause As Boolean
Dim cap$, i%, spaces%
Private OldX As Integer
Private OldY As Integer
Private DragMode As Boolean
Dim MoveMe As Boolean
Dim iImageNum As Integer

Private Sub ChkMute_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ChkMute.BackColor = &HF2C58C
End Sub

Private Sub CmdPause_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CmdPause.BackColor = &HF2C58C
End Sub

Private Sub CmdPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CmdPlay.BackColor = &HF2C58C
End Sub

Private Sub CmdStartWake_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CmdStartWake.BackColor = &HF2C58C
End Sub

Private Sub CmdStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CmdStop.BackColor = &HF2C58C
End Sub

Private Sub Form_DblClick()
    If iImageNum = 6 Then
        iImageNum = 1
    Else
        iImageNum = iImageNum + 1
    End If
    
    Me.Picture = ImageList2.ListImages(iImageNum).Picture
    Form2.Picture = ImageList2.ListImages(iImageNum).Picture
    Form3.Picture = ImageList2.ListImages(iImageNum).Picture
    frmAbout.Picture = ImageList2.ListImages(iImageNum).Picture
    FrmListPreview.Picture = ImageList2.ListImages(iImageNum).Picture
    FrmSaveList.Picture = ImageList2.ListImages(iImageNum).Picture
    findcolor (iImageNum)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        shAddForm
    End If
    
    MoveMe = True
    
    OldX = X
    OldY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MoveMe = True Then
        Me.Left = Me.Left + (X - OldX)
        Me.Top = Me.Top + (Y - OldY)
    End If
    clearBackcolor
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Me.Left = Me.Left + (X - OldX)
    'Me.Top = Me.Top + (Y - OldY)
    
    MoveMe = False
End Sub

Private Sub ChkMute_Click()
    If MediaPlayer1.Mute = True Then
        MediaPlayer1.Mute = False
        Image2.Visible = False
    Else
        MediaPlayer1.Mute = True
        Image2.Visible = True
    End If
End Sub

Private Sub CmdDown_Click()
Dim itmx As ListItem
If listviewOnAir.ListItems.count > 0 Then
If listviewOnAir.SelectedItem.Index < listviewOnAir.ListItems.count Then
    If listviewOnAir.Tag = "" Then
        listviewOnAir.Tag = listviewOnAir.SelectedItem.Index
    End If

If listviewOnAir.SelectedItem.Index = listviewOnAir.ListItems.count Then
    Set listviewOnAir.SelectedItem = listviewOnAir.ListItems(listviewOnAir.ListItems.count)
    Set listviewOnAir.DropHighlight = listviewOnAir.SelectedItem
Else
    Set itmx = listviewOnAir.ListItems.Add(listviewOnAir.SelectedItem.Index + 2, , listviewOnAir.SelectedItem.Text)
        itmx.SubItems(1) = listviewOnAir.SelectedItem.SubItems(1)
        itmx.SubItems(2) = listviewOnAir.SelectedItem.SubItems(2)
        listviewOnAir.ListItems.Remove (listviewOnAir.SelectedItem.Index)
       
        If listviewOnAir.SelectedItem.Index = Int(listviewOnAir.Tag) Then
            listviewOnAir.Tag = Int(listviewOnAir.Tag) + 1
        ElseIf listviewOnAir.SelectedItem.Index = listviewOnAir.Tag - 1 Then
            listviewOnAir.Tag = listviewOnAir.Tag - 1
        End If
       
    Set listviewOnAir.SelectedItem = listviewOnAir.ListItems(listviewOnAir.SelectedItem.Index + 1)
    Set listviewOnAir.DropHighlight = listviewOnAir.SelectedItem
Timertag.Enabled = True
End If
End If
End If
End Sub

Private Sub CmdPause_Click()
    On Error Resume Next
    If MediaPlayer1.CurrentPosition > 0 Then
        If pause = False Then
            MediaPlayer1.pause
            pause = True
        Else
            MediaPlayer1.Play
            pause = False
        End If
    End If
End Sub

Private Sub Cmdplay_Click()
    If MediaPlayer1.CurrentPosition > 0 Then
        FlatScrollBar1.Enabled = True
        MediaPlayer1.Play
        LblCurSong.Caption = listviewOnAir.SelectedItem.ListSubItems(2)
        TimerScrollCaption.Enabled = True
    End If
    If listviewOnAir.ListItems.count > 0 Then
        If listviewOnAir.SelectedItem.Selected = True Then
            LblCurSong.Caption = listviewOnAir.SelectedItem.ListSubItems(2)
            FlatScrollBar1.Enabled = True
            MediaPlayer1.FileName = listviewOnAir.SelectedItem.ListSubItems(1) & _
                    listviewOnAir.SelectedItem.ListSubItems(2) & ".mp3"
            TimerFlat.Enabled = True
            listviewOnAir.Tag = listviewOnAir.SelectedItem.Index
            TimerScrollCaption.Enabled = True
        End If
    End If
End Sub

Private Sub CmdStartWake_Click()
    If Len(Trim$(TxtWakeHour.Text)) And Len(Trim$(txtwakeMin.Text)) And listviewOnAir.ListItems.count > 0 Then
        TimerWake.Enabled = True
        If gLang = "Greek" Then
            LblStartWake.Caption = "Åíåñãü"
        Else
            LblStartWake.Caption = "Active"
        End If
    End If
End Sub

Private Sub CmdStop_Click()
    MediaPlayer1.Stop
    MediaPlayer1.CurrentPosition = 0
    LblCurSong.Caption = ""
End Sub

Private Sub CmdUp_Click()
Dim itmx As ListItem
If listviewOnAir.ListItems.count > 0 Then
        If listviewOnAir.Tag = "" Then
            listviewOnAir.Tag = listviewOnAir.SelectedItem.Index
        End If
        If listviewOnAir.SelectedItem.Index = CInt(listviewOnAir.Tag) And listviewOnAir.SelectedItem.Index > 1 Then
            listviewOnAir.Tag = CInt(listviewOnAir.Tag - 1)
        End If
        If listviewOnAir.SelectedItem.Index = 1 Then
            Set listviewOnAir.DropHighlight = listviewOnAir.SelectedItem
            
        Else
        If listviewOnAir.SelectedItem.Index = listviewOnAir.ListItems.count Then
            Set itmx = listviewOnAir.ListItems.Add(listviewOnAir.SelectedItem.Index - 1, , listviewOnAir.SelectedItem.Text)
                itmx.SubItems(1) = listviewOnAir.SelectedItem.SubItems(1)
                itmx.SubItems(2) = listviewOnAir.SelectedItem.SubItems(2)
                listviewOnAir.ListItems.Remove (listviewOnAir.SelectedItem.Index)
            Set listviewOnAir.SelectedItem = listviewOnAir.ListItems(listviewOnAir.SelectedItem.Index - 1)
            Set listviewOnAir.DropHighlight = listviewOnAir.SelectedItem
        Else
            Set itmx = listviewOnAir.ListItems.Add(listviewOnAir.SelectedItem.Index - 1, , listviewOnAir.SelectedItem.Text)
                itmx.SubItems(1) = listviewOnAir.SelectedItem.SubItems(1)
                itmx.SubItems(2) = listviewOnAir.SelectedItem.SubItems(2)
                listviewOnAir.ListItems.Remove (listviewOnAir.SelectedItem.Index)
                If listviewOnAir.SelectedItem.Index = listviewOnAir.Tag Then
                    listviewOnAir.Tag = listviewOnAir.Tag + 1
                End If
            Set listviewOnAir.SelectedItem = listviewOnAir.ListItems(listviewOnAir.SelectedItem.Index - 2)
            Set listviewOnAir.DropHighlight = listviewOnAir.SelectedItem
        End If
    Timertag.Enabled = True
        End If
End If
End Sub

Private Sub FlatScrollBar1_Change()
    If FlatScrollBar1.Value = MediaPlayer1.Duration Or FlatScrollBar1.Value < 0 Then
        Exit Sub
    End If
End Sub

Private Sub FlatScrollBar1_Scroll()
        MediaPlayer1.CurrentPosition = FlatScrollBar1.Value
End Sub

Private Sub Form_Load()
    
    gLang = GetSetting(App.Title, "Configure", "Language")
    If gLang = "Greek" Then
        OptLang(0).Value = True
    Else
        OptLang(1).Value = True
    End If
    'Dim SInfo As SYSTEM_INFO
    'GetSystemInfo SInfo
    'If str$(SInfo.dwProcessorType) = 0 Then
        cap$ = App.Title '& " on AMD Processor!!!"
    'Else
        'cap$ = App.Title & " on INTEL Processor..."
    'End If
    MediaPlayer1.Volume = SliderVolume.Value
    OptionAll.Value = True
    i% = 1
    spaces% = 0
    Me.Caption = cap$
    skin
    
End Sub

Private Sub Image1_dblClick()
    Dim sSave As String, Ret As Long
    Dim result As Long
    On Error GoTo ErrorHandler
    
    sSave = Space$(255)
    'Get the system directory
    Ret = GetSystemDirectory(sSave, 255)
    'Remove all unnecessary chr$(0)'s
    sSave = Left$(sSave, Ret)
    result = Shell(sSave & "\Sndvol32.exe", vbNormalFocus)
    Exit Sub
ErrorHandler:
    sSave = Mid$(sSave, 1, 10)
    result = Shell(sSave & "\Sndvol32.exe", vbNormalFocus)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
       shAddForm
    End If
End Sub



Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
       shAddForm
    End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clearBackcolor
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
       shAddForm
    End If
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clearBackcolor
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
       shAddForm
    End If
End Sub

Private Sub LabelCurSong_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
       shAddForm
    End If
End Sub

Private Sub LabelNumSon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
       shAddForm
    End If
End Sub

Private Sub LblCurSong_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
       shAddForm
    End If
End Sub

Private Sub LblDur_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
       shAddForm
    End If
End Sub

Private Sub LblDur_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clearBackcolor
End Sub

Private Sub LblElapsed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
       shAddForm
    End If
End Sub



Private Sub LblLang_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Form1.Height = 6090
     MoveMe = True
    
    OldX = X
    OldY = Y
End Sub

Private Sub LblLang_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MoveMe = True Then
        Me.Left = Me.Left + (X - OldX)
        Me.Top = Me.Top + (Y - OldY)
    End If
    clearBackcolor
End Sub

Private Sub LblLang_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveMe = False
    Me.Height = 7695
End Sub

Private Sub LblNumSongs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
       shAddForm
    End If
End Sub

Private Sub LblStartWake_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
       shAddForm
    End If
End Sub

Private Sub LblStartWake_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clearBackcolor
End Sub

Private Sub LblTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
       shAddForm
    End If
End Sub

Private Sub listviewOnAir_Click()
   Set listviewOnAir.DropHighlight = listviewOnAir.SelectedItem
    Timertag.Enabled = True
    
End Sub

Private Sub listviewOnAir_DblClick()
    Dim song As String
    If listviewOnAir.ListItems.count > 0 Then
        song = listviewOnAir.SelectedItem.ListSubItems(1) & _
                listviewOnAir.SelectedItem.ListSubItems(2) & ".mp3"
        LblCurSong.Caption = listviewOnAir.SelectedItem.ListSubItems(2)
        FlatScrollBar1.Enabled = True
        MediaPlayer1.FileName = song
        listviewOnAir.Tag = Int(listviewOnAir.SelectedItem.Index)

        TimerFlat.Enabled = True
        TimerScrollCaption.Enabled = True
    End If
End Sub

Private Sub listviewOnAir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clearBackcolor
End Sub

Private Sub listviewOnAir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cou%
    If Button = vbRightButton Then
        If listviewOnAir.ListItems.count > 0 Then
            
            If listviewOnAir.SelectedItem.ListSubItems(2) <> LblCurSong.Caption Then
                listviewOnAir.ListItems.Remove (listviewOnAir.SelectedItem.Index)
                
                For cou% = 1 To listviewOnAir.ListItems.count
                    If listviewOnAir.ListItems(cou%).ListSubItems(2) = LblCurSong.Caption Then
                        listviewOnAir.Tag = cou%
                        Exit For
                    End If
                Next cou%
                Timertag.Enabled = True
                
            Else
                If gLang = "Greek" Then
                    MsgBox "Äåí ìðïñåßò íá äéáãñÜøåéò ôï ôñáãïýäé ðïõ ðáßæåé.", vbInformation
                Else
                    MsgBox "You can 't remove the current song.", vbInformation
                End If
            End If
        End If
    End If
End Sub

Private Sub listviewOnAir_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call loadsongs(Form2)
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal result As Long)
    Dim myvalue As Integer
    If OptionRandom.Value = True Then
        Randomize Timer
        myvalue = Int((listviewOnAir.ListItems.count * Rnd))
        If myvalue = 0 Then
            myvalue = 1
        End If
        If myvalue = listviewOnAir.SelectedItem.Index Then
            myvalue = Int((listviewOnAir.ListItems.count * Rnd))
            If myvalue = 0 Then myvalue = 1
        End If
        listviewOnAir.Tag = myvalue
        listviewOnAir.SelectedItem = listviewOnAir.ListItems(myvalue)
        song = listviewOnAir.SelectedItem.ListSubItems(1) & _
                listviewOnAir.SelectedItem.ListSubItems(2) & ".mp3"
            Call listviewDeselect(listviewOnAir)
            Set listviewOnAir.DropHighlight = listviewOnAir.ListItems(CInt(listviewOnAir.Tag))
            MediaPlayer1.FileName = song
            LblCurSong.Caption = listviewOnAir.SelectedItem.ListSubItems(2)
        Exit Sub
    End If
    
    If listviewOnAir.SelectedItem.Index = listviewOnAir.ListItems.count Then
        If OptionAll.Value = True Then
            listviewOnAir.Tag = 1
            
            listviewOnAir.SelectedItem = listviewOnAir.ListItems(1)
             song = listviewOnAir.SelectedItem.ListSubItems(1) & _
                listviewOnAir.SelectedItem.ListSubItems(2) & ".mp3"
            Call listviewDeselect(listviewOnAir)
            Set listviewOnAir.DropHighlight = listviewOnAir.ListItems(CInt(listviewOnAir.Tag))
            MediaPlayer1.FileName = song
            LblCurSong.Caption = listviewOnAir.SelectedItem.ListSubItems(2)
            Exit Sub
        Else
            Exit Sub
        End If
    End If
    
    If listviewOnAir.SelectedItem.Index < listviewOnAir.ListItems.count Then
        If OptionCur.Value = False Then
            listviewOnAir.Tag = listviewOnAir.SelectedItem.Index
            listviewOnAir.Tag = listviewOnAir.Tag + 1
            listviewOnAir.SelectedItem = listviewOnAir.ListItems(CInt(listviewOnAir.Tag))
        Else
            listviewOnAir.SelectedItem = listviewOnAir.ListItems(CInt(listviewOnAir.Tag))
        End If
    End If
    song = listviewOnAir.SelectedItem.ListSubItems(1) & _
                listviewOnAir.SelectedItem.ListSubItems(2) & ".mp3"
    Call listviewDeselect(listviewOnAir)
    Set listviewOnAir.DropHighlight = listviewOnAir.ListItems(CInt(listviewOnAir.Tag))
    MediaPlayer1.FileName = song
    LblCurSong.Caption = listviewOnAir.SelectedItem.ListSubItems(2)
    
End Sub

Private Sub Option1_Click(Index As Integer)

End Sub

Private Sub OptLang_Click(Index As Integer)
    Select Case Index
        Case 0
            Call Greek
        Case 1
            Call English
    End Select
End Sub

Private Sub SliderVolume_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clearBackcolor
End Sub

Private Sub SliderVolume_Scroll()
    MediaPlayer1.Volume = SliderVolume.Value
End Sub

Private Sub Timer1_Timer()
    LblTime.Caption = Format$(Time, "hh:mm")
End Sub

Private Sub TimerHide_Timer()
    Dim fre
    If GetAsyncKeyState(VK_F11) Then fre = ShowWindow(Me.hwnd, SW_HIDE)
    If GetAsyncKeyState(VK_F12) Then fre = ShowWindow(Me.hwnd, SW_NORMAL)

End Sub

Private Sub TimerScrollCaption_Timer()
    Dim str$
    Dim strlen%
    If MediaPlayer1.CurrentPosition > 0 Then
        str$ = LblCurSong.Caption
        strlen% = Len(LblCurSong.Caption)
        If spaces% > 0 Then
            TimerScrollCaption.Interval = 100
            Me.Caption = String(spaces%, " ") & str$
            spaces% = spaces% - 1
        Else
            TimerScrollCaption.Interval = 200
            Me.Caption = Mid$(str$, i%, strlen%)
            If i% <= strlen% Then
                i% = i% + 1
            Else
                Me.Caption = cap$
                Sleep (600)
                spaces% = 0
                i% = 1
            End If
        End If
    Else
        Me.Caption = cap$
    End If
End Sub

Private Sub TimerCountSongs_Timer()
    LblNumSongs.Caption = listviewOnAir.ListItems.count
End Sub

Private Sub TimerFlat_Timer()
    FlatScrollBar1.Max = MediaPlayer1.Duration
    If MediaPlayer1.CurrentPosition >= 0 Then
        FlatScrollBar1.Value = MediaPlayer1.CurrentPosition
    Else
        FlatScrollBar1.Value = 0
    End If
End Sub

Private Sub Timertag_Timer()
    If listviewOnAir.ListItems.count > 0 Then
        Call listviewDeselect(listviewOnAir)
        If Len(listviewOnAir.Tag) Then
            Set listviewOnAir.SelectedItem = listviewOnAir.ListItems(CInt(listviewOnAir.Tag))
            Set listviewOnAir.DropHighlight = listviewOnAir.SelectedItem
        End If
        Timertag.Enabled = False
    End If
End Sub

Private Sub TimerTimes_Timer()
Dim min%, sec%, dur%, rmin%, rsec%
Dim sMin As String, sSec As String, sRmin As String, sRsec As String
If MediaPlayer1.Duration > 0 Then
    dur% = MediaPlayer1.CurrentPosition
    min% = dur% \ 60
    sec% = dur% - (min% * 60)
    If min% < 10 Then sMin = "0" & CStr(min%) Else sMin = CStr(min%)
    If sec% < 10 Then sSec = "0" & CStr(sec%) Else sSec = CStr(sec%)
    LblElapsed.Caption = sMin & ":" & sSec
    
    rmin% = (MediaPlayer1.Duration \ 60)
    rsec% = MediaPlayer1.Duration - (rmin% * 60)
    If rmin% < 10 Then sRmin = "0" & CStr(rmin%) Else sRmin = CStr(rmin%)
    If rsec% < 10 Then sRsec = "0" & CStr(rsec%) Else sRsec = CStr(rsec%)
    LblDur.Caption = sRmin & ":" & sRsec
    
End If
End Sub

Private Sub TimerWake_Timer()
    If Trim$(TxtWakeHour.Text) & ":" & Trim$(txtwakeMin.Text) = LblTime.Caption Then
        If listviewOnAir.ListItems.count > 0 Then
            MediaPlayer1.FileName = listviewOnAir.SelectedItem.ListSubItems(1) & _
                listviewOnAir.SelectedItem.ListSubItems(2) & ".mp3"
            LblCurSong.Caption = listviewOnAir.SelectedItem.ListSubItems(2)
            TimerWake.Enabled = False
            TimerFlat.Enabled = True
            FlatScrollBar1.Enabled = True
            TimerScrollCaption.Enabled = True
            listviewOnAir.DropHighlight = listviewOnAir.SelectedItem
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim ans%, st%
    Select Case Button
        Case "About"
            frmAbout.Show
        Case "¸îïäïò", "Exit"
            If MediaPlayer1.CurrentPosition > 0 Then
                    If gLang = "Greek" Then
                        MsgBox "Äåí åßíáé åöéêôÞ ç Ýîïäïò êáôÜ ôçí äéÜñêåéá áêñüáóçò ôñáãïõäéïý.", vbInformation
                    Else
                        MsgBox "Program cannot close while song plays."
                    End If
                    Exit Sub
            End If
            If gLang = "Greek" Then
                ans% = MsgBox("Åßóáé óßãïõñïò ãéá ôçí Ýîïäï áðü ôï ðñüãñáììá;", vbOKCancel + vbDefaultButton2)
            Else
                ans% = MsgBox("Are you really want to exit?", vbOKCancel + vbDefaultButton2)
            End If
                If ans% = vbOK Then
                    For st% = 0 To 100
                        Me.Move (Me.Left - (st% * 300))
                    Next st%
                    End
                Else
                    Exit Sub
                End If
    End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
   Dim ans%, cou%
   Select Case ButtonMenu
        Case "ÐñïóèÞêç Ôñáãïõäéïý", "Add Songs"
            Form2.Show
        Case "ÄéáãñáöÞ ÅðéëåãìÝíùí", "Remove Selected"
            If listviewOnAir.ListItems.count > 0 Then
                        If listviewOnAir.SelectedItem.ListSubItems(2) <> LblCurSong.Caption Then
                            For cou% = listviewOnAir.ListItems.count To 1 Step -1
                                If listviewOnAir.ListItems(cou%).Selected = True Then
                                    If listviewOnAir.ListItems(cou%).ListSubItems(2) <> LblCurSong.Caption Then
                                        listviewOnAir.ListItems.Remove (cou%)
                                    End If
                                End If
                            Next cou%
                            For cou% = 1 To listviewOnAir.ListItems.count
                                If listviewOnAir.ListItems(cou%).ListSubItems(2) = LblCurSong.Caption Then
                                    listviewOnAir.Tag = cou%
                                    Exit For
                                End If
                            Next cou%
                        
                        Else
                            If gLang = "Greek" Then
                                MsgBox "Äåí ìðïñåßò íá äéáãñÜøåéò ôï ôñáãïýäé ðïõ ðáßæåé.", vbInformation
                            Else
                                MsgBox "You can 't remove the current song.", vbInformation
                            End If
                        End If
            Else
                If gLang = "Greek" Then
                    MsgBox "Äåí õðÜñ÷ïõí åðéëåãìÝíá ôñáãïýäéá.", vbInformation
                Else
                    MsgBox "There are no songs selected.", vbInformation
                End If
            End If
        Case "ÄéáãñáöÞ ¼ëùí", "Remove All"
            If listviewOnAir.ListItems.count > 0 Then
                If MediaPlayer1.CurrentPosition > 0 Then
                    If gLang = "Greek" Then
                        MsgBox "ÄÝí ìðïñåßò íá äéáãñÜøåéò üëá ôá ôñáãïýäéá êáôÜ ôçí äéÜñêåéá ôçò áêñüáóçò." _
                            , vbInformation
                    Else
                        MsgBox "You can 't remove all songs while one plays.", vbInformation
                    End If
                    Exit Sub
                Else
                If gLang = "Greek" Then
                    ans% = MsgBox("Åßóáé óßãïõñïò ãéá ôçí äéáãñáöÞ üëçò ôçò ëßóôáò;", vbOKCancel + vbDefaultButton2)
                Else
                    ans% = MsgBox("Are you really want to remove the list?", vbOKCancel + vbDefaultButton2)
                End If
                If ans% = vbOK Then
                    listviewOnAir.ListItems.Clear
                    Timertag.Enabled = False
                Else
                    Exit Sub
                End If
                End If
            Else
                If gLang = "Greek" Then
                    MsgBox "ÄÝí õðÜñ÷ïõí ôñáãïýäéá ãéá äéáãñáöÞ.", vbInformation
                Else
                    MsgBox "There are no songs to remove.", vbInformation
                End If
            End If
        Case "About"
            frmAbout.Show
            
        Case "ÁðïèÞêåõóç Ëßóôáò", "Save List               "
            If listviewOnAir.ListItems.count > 0 Then
                FrmSaveList.Show
            Else
                If gLang = "Greek" Then
                    MsgBox "Äåí õðÜñ÷ïõí ôñáãïýäéá ðñïò áðïèÞêåõóç.", vbInformation
                Else
                    MsgBox "There are no songs to save.", vbInformation
                End If
            End If
        Case "ÁíÜêôçóç Ëßóôáò", "Load List               "
            Form3.Show
    End Select
    Exit Sub
    
End Sub



Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clearBackcolor
End Sub

Private Sub TxtWakehour_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
            If Len(TxtWakeHour.Text) = 1 Then
                txtwakeMin.SetFocus
                Exit Sub
            ElseIf Len(TxtWakeHour.Text) > 1 Then
                KeyAscii = 0
            End If
        Case 8
            If Len(TxtWakeHour.Text) Then
                Exit Sub
            End If
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtWakeHour_LostFocus()
If Len(Trim$(TxtWakeHour.Text)) Then
    If IsNumeric(TxtWakeHour.Text) Then
        If CInt(TxtWakeHour.Text) > 23 Then
            If gLang = "Greek" Then
                MsgBox "Ðáñáêáëþ äéïñèþóôå ôçí þñá.", vbInformation
            Else
                MsgBox "Please correct the hour.", vbInformation
            End If
            'KeyAscii = 0
            TxtWakeHour.SetFocus
            Exit Sub
        End If
     Else
        TxtWakeHour.Text = ""
     End If
End If
End Sub

Private Sub TxtWakeHour_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If gLang = "English" Then
            MsgBox "Operation not permited..."
        Else
            MsgBox "Ìç åðéôñåðüìåíç êßíçóç..."
        End If
    End If
End Sub

Private Sub txtwakeMin_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
        Case 48 To 57
            If Len(txtwakeMin.Text) = 1 Then
                Exit Sub
            ElseIf Len(txtwakeMin.Text) > 1 Then
                KeyAscii = 0
            End If
        Case 8
            If Len(txtwakeMin.Text) Then
                Exit Sub
            End If
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtwakeMin_LostFocus()
If Len(txtwakeMin.Text) Then
    If IsNumeric(txtwakeMin.Text) Then
        If CInt(txtwakeMin.Text) > 59 Then
            If gLang = "Greek" Then
                MsgBox "Ðáñáêáëþ äéïñèþóôå ôá ëåðôÜ.", vbInformation
            Else
                MsgBox "Please correct the minutes.", vbInformation
            End If
            'KeyAscii = 0
            txtwakeMin.SetFocus
            Exit Sub
        End If
    Else
        txtwakeMin.Text = ""
    End If
End If
End Sub
Private Sub clearBackcolor()
    CmdPlay.BackColor = &H8000000F
    CmdPause.BackColor = &H8000000F
    CmdStop.BackColor = &H8000000F
    ChkMute.BackColor = &H8000000F
    CmdStartWake.BackColor = &H8000000F
End Sub

Private Sub txtwakeMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If gLang = "English" Then
            MsgBox "Operation not permited..."
        Else
            MsgBox "Ìç åðéôñåðüìåíç êßíçóç..."
        End If
    End If
End Sub

Private Sub txtwakeMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clearBackcolor
End Sub
