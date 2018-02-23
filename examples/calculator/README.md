# VBA Control Array Calculator Example

Demonstrate how the control array can drastically cut down lines of code by consolidating common events.  Simply download the Excel file and run the form frmCalculator.

### How it works

When the form opens up we initialize two control array objects, NumericButtons and OperatorButtons.
The buttons have a tag property set to either "num" or "op" depending on if they are used to enter a number value or used as a calculator operation.

```VBA
Private WithEvents NumericButtons As ControlArray
Private WithEvents OperatorButtons As ControlArray
```
```VBA
Private Sub UserForm_Initialize()
    
    Set NumericButtons = ControlArray.Create(Me.Controls, ctCommandButton, "num")
    Set OperatorButtons = ControlArray.Create(Me.Controls, ctCommandButton, "op")
    
End Sub
```

The next part is to handle the events when the user clicks a numeric button. Notice how a reference to the button is passed into the event handling procedure, we will use this to evaluate which button was pushed and how it should be concatenated to the display value.

```VBA
Private Sub NumericButtons_onClick(ctl As IControl)

    Dim btn As MSForms.CommandButton
    Set btn = ctl.Object

    If lblDisplay.Caption = "0" Then
        lblDisplay.Caption = btn.Caption
    Else
    
        If mblnResetDisplay = True Then
            lblDisplay.Caption = btn.Caption
            mblnResetDisplay = False
        ElseIf IsNumeric(lblDisplay.Caption & btn.Caption) Then
            lblDisplay.Caption = lblDisplay.Caption & btn.Caption
        End If
            
    End If
    
End Sub
```

The last part is to handle the events when the user clicks an operator key, again note how the button reference is passed in.

```VBA
Private Sub OperatorButtons_onClick(ctl As IControl)

    Dim btn As MSForms.CommandButton
    Set btn = ctl.Object
    
    Select Case btn.Caption
        
        Case "C"
            
            mdblValueOne = 0
            mdblValueTwo = 0
            lblDisplay.Caption = "0"
            mMode = None
            
        Case "/"
            
            mdblValueOne = lblDisplay.Caption
            lblDisplay.Caption = 0
            mMode = Divide
        
        Case "*"
        
            mdblValueOne = lblDisplay.Caption
            lblDisplay.Caption = 0
            mMode = Multiply
        
        Case "+"
            
            mdblValueOne = lblDisplay.Caption
            lblDisplay.Caption = 0
            mMode = Add
        
        Case "-"
            
            mdblValueOne = lblDisplay.Caption
            lblDisplay.Caption = 0
            mMode = Subtract
        
        Case "="
            
            mdblValueTwo = lblDisplay.Caption
            mblnResetDisplay = True

            Select Case mMode
            
                Case Add
                    
                    lblDisplay.Caption = mdblValueOne + mdblValueTwo
                    
                Case Subtract
                    
                    lblDisplay.Caption = mdblValueOne - mdblValueTwo
                    
                Case Multiply
                    
                    lblDisplay.Caption = mdblValueOne * mdblValueTwo
                    
                Case Divide
                    
                    If mdblValueTwo = 0 Then
                        lblDisplay.Caption = "Div by 0 error"
                        mblnResetDisplay = True
                    Else
                        lblDisplay.Caption = mdblValueOne / mdblValueTwo
                    End If
                
            End Select
            
            'reset calculator
            mdblValueOne = 0
            mdblValueTwo = 0
            mMode = None
        
    End Select
End Sub
```

### Bonus Sample Code Included!

Included in the sample calculator is button highlighting made easy. something we take for granted in other languages is really a pain in VBA.  Not anymore!

This code sets the button that is being moused over to a shade of light red
```VBA
Private Sub NumericButtons_onMouseMove(ctl As IControl, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ctl.Object.BackColor = &H8080FF
    mblnIsHighlighted = True
    
End Sub
```
This code sets the button that is being moused over to a shade of light blue
```VBA
Private Sub OperatorButtons_onMouseMove(ctl As IControl, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ctl.Object.BackColor = &H8000000D
    mblnIsHighlighted = True
    
End Sub
```

This code sets all the buttons to its original backcolor, but only if the highlight flag is TRUE.
```VBA
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    'resets the button backcolors, if one is highlighted.
    
    Dim ctl As IControl
    
    If mblnIsHighlighted = True Then
    
        For Each ctl In NumericButtons
            ctl.Object.BackColor = &H8000000F
        Next
        
        For Each ctl In OperatorButtons
            ctl.Object.BackColor = &H8000000F
        Next
        mblnIsHighlighted = False
        
    End If
    
End Sub
```
