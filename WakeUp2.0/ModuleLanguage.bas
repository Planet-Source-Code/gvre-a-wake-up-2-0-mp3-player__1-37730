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
        .LblHearingChoices.Caption = "ÅðéëïãÝò Áêñüáóçò"
        .CmdStartWake.Caption = "Åíåñãïðïßçóç Áöýðíéóçò"
        .OptionCur.Caption = "ÅðáíÜëçøç ÔñÝ÷ïíôïò"
        .OptionAll.Caption = "ÅðáíÜëçøç ¼ëùí"
        .OptionRandom.Caption = "Ôõ÷áßá ÓåéñÜ"
        .LabelCurSong.Caption = "ÔñÝ÷ïí Ôñáãïýäé"
        .Label1.Caption = "ÔñÝ÷ïõóá ¿ñá"
        .Label2.Caption = "¿ñá Áöýðíéóçò"
        .Label3.Caption = "ÄéÜñêåéá Ôñáãïõäéïý"
        .Label4.Caption = "×ñüíïò Áêñüáóçò"
        .LabelNumSon.Caption = "Áñéèìüò Ôñáãïõäéþí"
    End With
    
    
    With Form1.Toolbar1.Buttons(1)
        .Caption = "Äéá÷åßñçóç Ôñáãïõäéþí"
        .ButtonMenus(1).Text = "ÐñïóèÞêç Ôñáãïõäéïý"
        .ButtonMenus(2).Text = "ÄéáãñáöÞ ÅðéëåãìÝíùí"
        .ButtonMenus(3).Text = "ÄéáãñáöÞ ¼ëùí"
    End With
    
    With Form1.Toolbar1.Buttons(2)
        .Caption = "Äéá÷åßñçóç Ëßóôáò"
        .ButtonMenus(1).Text = "ÁðïèÞêåõóç Ëßóôáò"
        .ButtonMenus(2).Text = "ÁíÜêôçóç Ëßóôáò"
    End With
    
    
    
    With Form2
        .Caption = "ÅéóáãùãÞ Ôñáãïõäéþí"
        .CmdAdd.Caption = "ÐñïóèÞêç Ôñáãïõäéþí"
        .CmdClose.Caption = "Êëåßóéìï Öüñìáò"
    End With
    
    With Form3
        .Caption = "Öüñôùìá Ëßóôáò"
        .CmdLoadList.Caption = "Öüñôùìá Ëßóôáò"
        .CmdClose.Caption = "Êëåßóéìï Öüñìáò"
    End With
    
    With FrmListPreview
        .Caption = "Ðñïåðéóêüðçóç Ëßóôáò"
        .CmdClose.Caption = "Êëåßóéìï Öüñìáò"
    End With
    
    With FrmSaveList
        .Caption = "ÁðïèÞêåõóç Ëßóôáò"
        .cmdSaveList.Caption = "ÁðïèÞêåõóç Ëßóôáò"
        .LblSaveList.Caption = "¼íïìá Áñ÷åßïõ Ðñïò ÁðïèÞêåõóç"
        .CmdCloseForm.Caption = "Êëåßóéìï Öüñìáò"
    End With
    
    Form1.Toolbar1.Buttons(4).Caption = "¸îïäïò"
End Sub
