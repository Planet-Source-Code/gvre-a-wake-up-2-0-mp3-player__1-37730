VERSION 5.00
Begin VB.Form frmMessagebox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wake Up 2.0"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   Icon            =   "frmMessagebox.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChoice 
      Caption         =   "No To All"
      Height          =   375
      Index           =   3
      Left            =   3720
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Yes To All"
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "No"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Yes"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lbl1 
      Caption         =   "ÔÏ ÔÑÁÃÏÕÄÉ ÕÐÁÑ×ÅÉ ÓÔÇÍ ËÉÓÔÁ.ÍÁ ÐÑÏÓÔÅÈÅÉ ÎÁÍÁ?"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmMessagebox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChoice_Click(Index As Integer)
    Select Case Index
        Case 0
            sMsgAns = "yes"
            Unload Me
        Case 1
            sMsgAns = "no"
            Unload Me
        Case 2
            sMsgAns = "yes2all"
            Unload Me
        Case 3
            sMsgAns = "no2all"
            Unload Me
        Case Else
End Sub

Private Sub Form_Load()

End Sub
