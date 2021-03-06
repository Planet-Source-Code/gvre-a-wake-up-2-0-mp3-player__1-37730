Attribute VB_Name = "ModuleLanguage"
Option Explicit
Global gLang As String


Public Sub English()
    gLang = "English"
    SaveSetting App.Title, "Configure", "Language", gLang
    With Form1
        .LblHearingChoices.Caption = "Hearing Choices"
        .CmdStartWake.Caption = "Rousement Activation"
        .OptionCur.Caption = "Repeat Current"
        .OptionAll.Caption = "Repeat All"
        .OptionRandom.Caption = "Random Play"
        .LabelCurSong.Caption = "Current Song"
        .Label1.Caption = "Current Time"
        .Label2.Caption = "Rousement Time"
        .Label3.Caption = "Song Duration"
        .Label4.Caption = "Hearing Time"
        .LabelNumSon.Caption = "Number Of Songs"
    End With
    
    
    With Form1.Toolbar1.Buttons(1)
        .Caption = "      Songs Manipulate     "
        .ButtonMenus(1).Text = "Add Songs"
        .ButtonMenus(2).Text = "Remove Selected"
        .ButtonMenus(3).Text = "Remove All"
    End With
    
    With Form1.Toolbar1.Buttons(2)
        .Caption = "    List Manipulate    "
        .ButtonMenus(1).Text = "Save List               "
        .ButtonMenus(2).Text = "Load List               "
    End With
    
    
    
    With Form2
        .Caption = "Song Addition"
        .CmdAdd.Caption = "Add Selected"
        .CmdClose.Caption = "Close Form"
    End With
    
    With Form3
        .Caption = "Load List"
        .CmdLoadList.Caption = "Load List"
        .CmdClose.Caption = "Close Form"
    End With
    
    With FrmListPreview
        .Caption = "List Preview"
        .CmdClose.Caption = "Close Form"
    End With
    
    With FrmSaveList
        .Caption = "Save List"
        .cmdSaveList.Caption = "Save List"
        .LblSaveList.Caption = "File Name To Save"
        .CmdCloseForm.Caption = "Close Form"
    End With
    
    Form1.Toolbar1.Buttons(4).Caption = "Exit"
End Sub

Public Sub Greek()
    gLang = "Greek"
    SaveSetting App.Title, "Configure", "Language", gLang
    With Form1
        .LblHearingChoices.Caption = "�������� ��������"
        .CmdStartWake.Caption = "������������ ���������"
        .OptionCur.Caption = "��������� ���������"
        .OptionAll.Caption = "��������� ����"
        .OptionRandom.Caption = "������ �����"
        .LabelCurSong.Caption = "������ ��������"
        .Label1.Caption = "�������� ���"
        .Label2.Caption = "��� ���������"
        .Label3.Caption = "�������� ����������"
        .Label4.Caption = "������ ��������"
        .LabelNumSon.Caption = "������� ����������"
    End With
    
    
    With Form1.Toolbar1.Buttons(1)
        .Caption = "���������� ����������"
        .ButtonMenus(1).Text = "�������� ����������"
        .ButtonMenus(2).Text = "�������� �����������"
        .ButtonMenus(3).Text = "�������� ����"
    End With
    
    With Form1.Toolbar1.Buttons(2)
        .Caption = "���������� ������"
        .ButtonMenus(1).Text = "���������� ������"
        .ButtonMenus(2).Text = "�������� ������"
    End With
    
    
    
    With Form2
        .Caption = "�������� ����������"
        .CmdAdd.Caption = "�������� ����������"
        .CmdClose.Caption = "�������� ������"
    End With
    
    With Form3
        .Caption = "������� ������"
        .CmdLoadList.Caption = "������� ������"
        .CmdClose.Caption = "�������� ������"
    End With
    
    With FrmListPreview
        .Caption = "������������� ������"
        .CmdClose.Caption = "�������� ������"
    End With
    
    With FrmSaveList
        .Caption = "���������� ������"
        .cmdSaveList.Caption = "���������� ������"
        .LblSaveList.Caption = "����� ������� ���� ����������"
        .CmdCloseForm.Caption = "�������� ������"
    End With
    
    Form1.Toolbar1.Buttons(4).Caption = "������"
End Sub
