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
        .LblHearingChoices.Caption = "Επιλογές Ακρόασης"
        .CmdStartWake.Caption = "Ενεργοποίηση Αφύπνισης"
        .OptionCur.Caption = "Επανάληψη Τρέχοντος"
        .OptionAll.Caption = "Επανάληψη Όλων"
        .OptionRandom.Caption = "Τυχαία Σειρά"
        .LabelCurSong.Caption = "Τρέχον Τραγούδι"
        .Label1.Caption = "Τρέχουσα Ώρα"
        .Label2.Caption = "Ώρα Αφύπνισης"
        .Label3.Caption = "Διάρκεια Τραγουδιού"
        .Label4.Caption = "Χρόνος Ακρόασης"
        .LabelNumSon.Caption = "Αριθμός Τραγουδιών"
    End With
    
    
    With Form1.Toolbar1.Buttons(1)
        .Caption = "Διαχείρηση Τραγουδιών"
        .ButtonMenus(1).Text = "Προσθήκη Τραγουδιού"
        .ButtonMenus(2).Text = "Διαγραφή Επιλεγμένων"
        .ButtonMenus(3).Text = "Διαγραφή Όλων"
    End With
    
    With Form1.Toolbar1.Buttons(2)
        .Caption = "Διαχείρηση Λίστας"
        .ButtonMenus(1).Text = "Αποθήκευση Λίστας"
        .ButtonMenus(2).Text = "Ανάκτηση Λίστας"
    End With
    
    
    
    With Form2
        .Caption = "Εισαγωγή Τραγουδιών"
        .CmdAdd.Caption = "Προσθήκη Τραγουδιών"
        .CmdClose.Caption = "Κλείσιμο Φόρμας"
    End With
    
    With Form3
        .Caption = "Φόρτωμα Λίστας"
        .CmdLoadList.Caption = "Φόρτωμα Λίστας"
        .CmdClose.Caption = "Κλείσιμο Φόρμας"
    End With
    
    With FrmListPreview
        .Caption = "Προεπισκόπηση Λίστας"
        .CmdClose.Caption = "Κλείσιμο Φόρμας"
    End With
    
    With FrmSaveList
        .Caption = "Αποθήκευση Λίστας"
        .cmdSaveList.Caption = "Αποθήκευση Λίστας"
        .LblSaveList.Caption = "Όνομα Αρχείου Προς Αποθήκευση"
        .CmdCloseForm.Caption = "Κλείσιμο Φόρμας"
    End With
    
    Form1.Toolbar1.Buttons(4).Caption = "Έξοδος"
End Sub
