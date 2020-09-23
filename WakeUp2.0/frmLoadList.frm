VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Öüñôùìá Ëßóôáò"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   Icon            =   "frmLoadList.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdClose 
      Cancel          =   -1  'True
      Caption         =   "Êëåßóéìï Öüñìáò"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton CmdLoadList 
      Caption         =   "Öüñôùìá Ëßóôáò"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   2520
      Pattern         =   "*.wul"
      TabIndex        =   2
      Top             =   0
      Width           =   6855
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub CmdLoadList_Click()
    Call loadlist
    SaveSetting App.Title, "Configure", "SaveListPath", Me.Dir1.path
End Sub

Private Sub Dir1_Change()
    File1 = Dir1
End Sub

Private Sub Drive1_Change()
    Dim res%
    On Error GoTo errhandler
    Dir1 = Drive1
    Exit Sub
errhandler:
     Select Case Err
        Case 68
            res% = MsgBox("Ç óõóêåõÞ äåí åßíáé äéáèÝóéìç.", vbAbortRetryIgnore)
            Select Case res%
                Case vbAbort
                    Exit Sub
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
        End Select
End Sub

Private Sub File1_dblClick()
    Call loadlist
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ln%, songs$

     Select Case Button
        Case vbRightButton 'Right mouse button
            If Len(Form3.File1.FileName) Then
                ln% = Len(Form3.File1.FileName)
                FrmListPreview.Show
                Open Form3.Dir1.path & "\" & Mid$(Form3.File1.FileName, 1, (ln% - 4)) & ".son" For Input As #2
                    Do Until EOF(2)
                        Input #2, songs$
                        FrmListPreview.LstboxListPreview.AddItem (songs$)
                    Loop
                Close 2
            End If
    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Dir1.path = GetSetting(App.Title, "Configure", "SaveListPath")
   Form3.Picture = Form1.Picture
End Sub
