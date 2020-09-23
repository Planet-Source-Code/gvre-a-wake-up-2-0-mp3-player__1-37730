VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "ÅéóáãùãÞ Ôñáãïõäéþí"
   ClientHeight    =   3615
   ClientLeft      =   1065
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "FrmAddSongs.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3615
   ScaleWidth      =   9480
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   4440
      Top             =   3000
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Êëåßóéìï Öüñìáò"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "ÐñïóèÞêç Ôñáãïõäéþí"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   3120
      MultiSelect     =   2  'Extended
      OLEDragMode     =   1  'Automatic
      Pattern         =   "*.mp3"
      TabIndex        =   2
      Top             =   0
      Width           =   6495
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAdd_Click()
    Call loadsongs(Form2)
End Sub

Private Sub CmdClose_Click()
    SaveSetting App.Title, "Configure", "Path", Me.Dir1.path
    Unload Me
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
    Dim itm As ListItem
    Dim count%, ans%, metr%, i%
    If Form1.listviewOnAir.ListItems.count > 0 Then
        If File1.ListCount > 0 Then
            For metr% = 0 To File1.ListCount - 1
                If File1.Selected(metr%) = True Then
                    File1.Tag = metr%
                    Exit For
                End If
            Next metr%
        End If
        For count% = 1 To Form1.listviewOnAir.ListItems.count
                If Trim$(Mid$(File1.List(File1.Tag), 1, Len(File1.List(File1.Tag)) - 4)) = Trim$(Form1.listviewOnAir.ListItems(count%).SubItems(2)) Then
                    ans% = MsgBox("Ôï ôñáãïýäé '" & Mid$(File1.List(File1.ListIndex), 1, Len(File1.FileName) - 4) & "' õðÜñ÷åé óôçí ëßóôá.Íá ðñïóôåèåß îáíÜ;", vbOKCancel)
                    If ans% = vbOK Then
                    Set itm = Form1.listviewOnAir.ListItems.Add
                        itm.SubItems(1) = File1.path & "\"
                        itm.SubItems(2) = Mid$(File1.List(File1.ListIndex), 1, Len(File1.FileName) - 4)
                    Else
                        Exit Sub
                    End If
                Else
                    If count% = Form1.listviewOnAir.ListItems.count Then
                    Set itm = Form1.listviewOnAir.ListItems.Add
                        itm.SubItems(1) = File1.path & "\"
                        itm.SubItems(2) = Mid$(File1.List(File1.ListIndex), 1, Len(File1.FileName) - 4)
                        Exit For
                    End If
                End If
        Next count%
    Else
    Set itm = Form1.listviewOnAir.ListItems.Add
        itm.SubItems(1) = File1.path & "\"
        itm.SubItems(2) = Mid$(File1.List(File1.ListIndex), 1, Len(File1.FileName) - 4)
    End If
End Sub

Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = 65 Then
        File1_KeyPress (Asc("+"))
    End If
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)
Dim iCount As Integer
    Select Case KeyAscii
        Case Asc("!"), Asc("+")
            If File1.ListCount > 0 Then
                For iCount = 0 To File1.ListCount - 1
                    File1.Selected(iCount) = True
                Next iCount
            End If
        Case 13
            Call loadsongs(Form2)
        Case Else
        
    End Select
End Sub

Private Sub File1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case vbRightButton
        File1_KeyPress (Asc("+"))
    Case Else
    
    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
        Me.Dir1.path = GetSetting(App.Title, "Configure", "Path")
    Form2.Picture = Form1.Picture
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 0 Then 'Normal
        If Me.Width > 9720 Then
            File1.Width = Me.Width - 3225
        Else
            Me.Width = 9720
        End If
    ElseIf Me.WindowState = 2 Then 'Maximized
        Me.WindowState = 0
        Me.Height = 9510
        Me.Width = 12120
        Me.File1.Width = Form2.Width - 3135
    ElseIf Me.WindowState = 1 Then 'Minimized
        Exit Sub
    End If
    If Me.Height < 3990 Then
        Me.Height = 3990
        Exit Sub
    End If
    If Me.Height > 9510 Then
        Me.Height = 9510
        Exit Sub
    End If
    If Me.Width > 12120 Then
        Me.Width = 12120
    End If
    
   cmdClose.Top = Form2.Height - (2 * cmdClose.Height) - 100
   CmdAdd.Top = Form2.Height - (2 * CmdAdd.Height) - 100
   File1.Height = Me.Height - 975
   Dir1.Height = Me.Height - 1425
   
End Sub

Private Sub Timer1_Timer()
     If Me.Height < 3990 Then
        Me.Height = 3990
        Exit Sub
    End If
    If Me.Width > 9720 Then
        File1.Width = Me.Width - 3225
    End If
   cmdClose.Top = Form2.Height - (2 * cmdClose.Height) - 100
   CmdAdd.Top = Form2.Height - (2 * CmdAdd.Height) - 100
   File1.Height = Me.Height - 975
   Dir1.Height = Me.Height - 1425
End Sub
