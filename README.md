# VBA Control Array

Add the power of control arrays to VBA and simplify event handling.  

Here is a list of the controls that can be added to the control array collection.
```
Checkbox
Combobox
CommandButton
Frame
Image
Label
ListBox
MultiPage
OptionButton
ScrollBar
SpinButton
TabStrip
TextBox
ToggleButton
```
### Prerequisites

Any VBA product that uses Microsoft Forms 2.0 library by default.  Microsoft Access will not work with this code.
Tested on Excel and Word 2010-2016.

### Quick Start
Download one of the blank templates from the quickstart folder, there is an [Excel](https://github.com/bohicajr/vba-controlarray/blob/master/quickstart/VBA-ControlArray_Blank.xlsm) and [Word](https://github.com/bohicajr/vba-controlarray/blob/master/quickstart/VBA-ControlArray_Blank.docm) version available.

### Manual Install

Download all the .cls files from github and import the files into your VBA project.
**Copy and paste will not work!**

Here is a list of the class files you must import into the VBA IDE.

```
CACheckbox.cls
CACombobox.cls
CACommandButton.cls
CAFrame.cls
CAImage.cls
CALabel.cls
CAListBox.cls
CAMultiPage.cls
CAOptionButton.cls
CAScrollBar.cls
CASpinButton.cls
CATabStrip.cls
CATextBox.cls
CAToggleButton.cls
ControlArray.cls
IControl.cls
```

### Basic Example
Add a user form to your VBA project and add as many command buttons as you want, then add the following code behind the form.

```VBA
Option Explicit

Private WithEvents Buttons as ControlArray

Private Sub UserForm_Initialize()
    
    Set Buttons = ControlArray.Create(Me.Controls)

End Sub

Private Sub Buttons_onClick(ctl As IControl)
    
    MsgBox "You clicked the button " & ctl.Name
    
End Sub
```

note that after you add the line that has WithEvents, that in the editor you now can choose the Buttons object, and select from it's many events to handle.  Every event also sends back a reference to the object that raised it!

### Known Issues

- AfterUpdate, BeforeUpdate, Enter and Exit events are not available. this is a limitation of MSForms and VBA.  If you want to use those events you'll have to work with them in the traditional manor.

- Frames, MultiPage and TabStrip have several events that are slightly unique in its parameters, you have to listen for their distinctive events.
