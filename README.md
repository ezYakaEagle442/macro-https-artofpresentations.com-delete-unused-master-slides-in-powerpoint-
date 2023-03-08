# Delete unused master slides in powerpoint

Credits: [https://artofpresentations.com/delete-unused-master-slides-in-powerpoint/](https://artofpresentations.com/delete-unused-master-slides-in-powerpoint/)

In the opened presentation, click on the “View” tab in the menu button. Then click on the “Macros” option.

You can alternatively press the “Alt+F8” keys on your keyboard. This will open a “Macro” dialog box.

In the “Macro” dialog box, type in a title inside the “Macro name” box.

Then click on the “Create” button to create a new macro. This will open a “Microsoft Visual Basic for Applications” window.

Step-3: Paste the Macro code

```sh
Sub SlideMasterCleanup()
Dim i As Integer
Dim j As Integer
Dim oPres As Presentation
Set oPres = ActivePresentation
On Error Resume Next
With oPres
    For i = 1 To .Designs.Count
        For j = .Designs(i).SlideMaster.CustomLayouts.Count To 1 Step -1
            .Designs(i).SlideMaster.CustomLayouts(j).Delete
        Next
    Next i
End With
End Sub
```


Now all you have to do is copy the macro code mentioned below and paste it into the description box.

Then click on the “Run sub” icon which looks like the play button in the toolbar above the description box. Alternatively, you can press the “F5” key on your keyboard.

This will delete all the unused master slides.
