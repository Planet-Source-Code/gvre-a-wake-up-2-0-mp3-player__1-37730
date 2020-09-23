VERSION 5.00
Begin VB.Form FrmListPreview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ðñïåðéóêüðçóç Ëßóôáò"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   Icon            =   "FrmListPreview.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdClose 
      Cancel          =   -1  'True
      Caption         =   "Êëåßóéìï Öüñìáò"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6600
      Width           =   2175
   End
   Begin VB.ListBox LstboxListPreview 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "FrmListPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdClose_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    FrmListPreview.Picture = Form1.Picture
End Sub
