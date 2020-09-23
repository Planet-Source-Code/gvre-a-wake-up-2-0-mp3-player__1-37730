Attribute VB_Name = "ModuleSystem"
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long 'Pairnei to system directory

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const SW_HIDE = 0
Public Const SW_NORMAL = 1

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Sub skin()
    Randomize Timer
    iImageNum = Round(6 * Rnd, 0)
        If iImageNum = 0 Then
            iImageNum = 6
        End If
    Form1.Picture = Form1.ImageList2.ListImages(iImageNum).Picture
    Form2.Picture = Form1.ImageList2.ListImages(iImageNum).Picture
    Form3.Picture = Form1.ImageList2.ListImages(iImageNum).Picture
    frmAbout.Picture = Form1.ImageList2.ListImages(iImageNum).Picture
    FrmListPreview.Picture = Form1.ImageList2.ListImages(iImageNum).Picture
    FrmSaveList.Picture = Form1.ImageList2.ListImages(iImageNum).Picture
    
    findcolor (iImageNum)
End Sub

Sub findcolor(iNum As Integer)
    Select Case iNum
        Case 1, 2, 6
            Form1.listviewOnAir.BackColor = &HFAA36D
            Form1.listviewOnAir.ForeColor = &H80FFFF
        Case 3
            Form1.listviewOnAir.BackColor = &HA8BEC4
            Form1.listviewOnAir.ForeColor = &HFF0000
        Case 4
            Form1.listviewOnAir.BackColor = &H708BFC
            Form1.listviewOnAir.ForeColor = &HFF0000
        Case 5
            Form1.listviewOnAir.BackColor = &H6FB6FD
            Form1.listviewOnAir.ForeColor = &HFF0000
        Case Else
    End Select
End Sub
