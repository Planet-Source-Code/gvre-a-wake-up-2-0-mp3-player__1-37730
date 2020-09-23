Attribute VB_Name = "ModuleListview"
Sub listviewDeselect(name As ListView)
Dim i%
For i% = 1 To name.ListItems.count
    If name.ListItems(i%).Selected = True Then
       name.ListItems(i%).Selected = False
    End If
Next i%
End Sub

Sub loadlist()
Dim ln%, songs$, path$
Dim li As ListItem
On Error GoTo notEx
ln% = Len(Form3.File1.FileName)

 
    If Right$(Form3.Dir1.path, 1) <> "\" Then
        Form3.Dir1.path = Form3.Dir1.path & "\"
    End If
                

Open Form3.Dir1.path & Form3.File1.FileName For Input As #1
Open Form3.Dir1.path & Mid$(Form3.File1.FileName, 1, (ln% - 4)) & ".son" For Input As #2
Do Until EOF(1)
    Set li = Form1.listviewOnAir.ListItems.Add
    Input #1, path$
    Input #2, songs$
    li.SubItems(1) = path$
    li.SubItems(2) = songs$
Loop
Close 2
Close 1
Exit Sub
notEx:
    If gLang = "Greek" Then
        MsgBox "ÄÝí âñÝèçêå ôï áðáéôïýìåíï áñ÷åßï ãéá ôï öüñôùìá ôçò ëßóôáò."
    Else
        MsgBox "One needful file didn 't found."
    End If
End Sub

Sub loadsongs(source_form As Form)
Dim counter%, count%, ans%
Dim itmx As ListItem
    If Form2.File1.ListCount > 0 Then
        For counter% = 0 To Form2.File1.ListCount - 1
            If source_form.File1.Selected(counter%) = True Then
                If Form1.listviewOnAir.ListItems.count > 0 Then
                    For count% = 1 To Form1.listviewOnAir.ListItems.count
                        If Trim$(Mid$(source_form.File1.List(counter%), 1, Len(source_form.File1.List(counter%)) - 4)) = Trim$(Form1.listviewOnAir.ListItems(count%).SubItems(2)) Then
                            If gLang = "Greek" Then
                                ans% = MsgBox("Ôï ôñáãïýäé ' " & Mid$(Form2.File1.List(counter%), 1, Len(Form2.File1.FileName) - 4) & " ' õðÜñ÷åé óôçí ëßóôá.Íá ðñïóôåèåß îáíÜ;", vbYesNo + vbDefaultButton2)
                            Else
                                ans% = MsgBox("The Song ' " & Mid$(Form2.File1.List(counter%), 1, Len(Form2.File1.FileName) - 4) & " ' exists in the list.Add it again?", vbYesNo + vbDefaultButton2)
                            End If
                            If ans% = vbOK Then
                                Set itmx = Form1.listviewOnAir.ListItems.Add(, , "")
                                    itmx.SubItems(1) = source_form.File1.path & "\"
                                    itmx.SubItems(2) = Mid$(source_form.File1.List(counter%), 1, Len(source_form.File1.List(counter%)) - 4)
                            Else
                                Exit For
                            End If
                        Else
                            If count% = Form1.listviewOnAir.ListItems.count Then
                                Set itmx = Form1.listviewOnAir.ListItems.Add(, , "")
                                    itmx.SubItems(1) = source_form.File1.path & "\"
                                    itmx.SubItems(2) = Mid$(source_form.File1.List(counter%), 1, Len(source_form.File1.List(counter%)) - 4)
                            End If
                        End If
                    Next count%
                Else
                    Set itmx = Form1.listviewOnAir.ListItems.Add(, , "")
                        itmx.SubItems(1) = source_form.File1.path & "\"
                        itmx.SubItems(2) = Mid$(source_form.File1.List(counter%), 1, Len(source_form.File1.List(counter%)) - 4)
                End If
            End If
        Next counter%
    End If
End Sub

Public Sub shAddForm()
    Form2.Show , Form1
End Sub
