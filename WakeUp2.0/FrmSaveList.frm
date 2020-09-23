VERSION 5.00
Begin VB.Form FrmSaveList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÁðïèÞêåõóç Ëßóôáò"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   Icon            =   "FrmSaveList.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   3720
      Pattern         =   "*.wul"
      TabIndex        =   6
      Top             =   840
      Width           =   4575
   End
   Begin VB.CommandButton CmdCloseForm 
      Cancel          =   -1  'True
      Caption         =   "Êëåßóéìï Öüñìáò"
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton cmdSaveList 
      Caption         =   "ÁðïèÞêåõóç Ëßóôáò"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtFilename 
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3495
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label LblSaveList 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "¼íïìá Áñ÷åßïõ Ðñïò ÁðïèÞêåõóç"
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
      Left            =   3720
      TabIndex        =   3
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "FrmSaveList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCloseForm_Click()
    Unload Me
End Sub

Private Sub cmdSaveList_Click()
Dim name$, cou%, path$, songs$
            If Len(Trim$(txtFilename.Text)) Then
                name$ = Trim$(txtFilename.Text)
                
                If Right$(Dir1.path, 1) <> "\" Then
                    Dir1.path = Dir1.path & "\"
                End If
                
                
                Open Dir1.path & name$ & ".wul" For Output As #1
                    Open Dir1.path & name$ & ".son" For Output As #2
                    For cou% = 1 To Form1.listviewOnAir.ListItems.count
                        path$ = Form1.listviewOnAir.ListItems(cou%).SubItems(1)
                        songs$ = Form1.listviewOnAir.ListItems(cou%).SubItems(2)
                        Write #1, path$ 'An valw print kai to tragoudi periexei koma (,) tote xwrizei ton titlo
                        Write #2, songs$
                    Next cou%
                    Close 2
                Close 1
                SaveSetting App.Title, "Configure", "SaveListPath", Me.Dir1.path
             End If
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

Private Sub Form_Load()
Dim iAns As Integer

On Error GoTo ErrorHandler
    Dir1.path = GetSetting(App.Title, "Configure", "SaveListPath")
    FrmSaveList.Picture = Form1.Picture
    Exit Sub
ErrorHandler:
   Select Case Err
        Case 68
            iAns = MsgBox("Ç óõóêåõÞ äåí åßíáé äéáèÝóéìç.", vbAbortRetryIgnore)
            Select Case iAns
                Case vbAbort
                    Exit Sub
                Case vbRetry
                    Resume
                Case vbIgnore
                    Resume Next
            End Select
        Case 76
            DeleteSetting App.Title, "Configure", "SaveListPath"
        Case Else
        
        End Select
End Sub
