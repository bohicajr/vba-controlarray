VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalculator 
   Caption         =   "Calculator Sample"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4020
   OleObjectBlob   =   "frmCalculator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' VBA-ControlArray Calculator Sample
' (c) J Martinez - https://github.com/bohicajr/vba-controlarray
'
'
''----------------------------------------------------------------------------''
Option Compare Binary
Option Explicit

Private WithEvents NumericButtons As ControlArray
Attribute NumericButtons.VB_VarHelpID = -1
Private WithEvents OperatorButtons As ControlArray
Attribute OperatorButtons.VB_VarHelpID = -1

Private Enum enMode
    Add
    Subtract
    Divide
    Multiply
    None
End Enum

Private mdblValueOne As Double
Private mdblValueTwo As Double
Private mMode As enMode
Private mblnResetDisplay As Boolean
Private mblnIsHighlighted As Boolean

Private Sub UserForm_Initialize()
    
    'load the two control arrays with command buttons based on there tag property
    'NumericButtons will add the number and decimal buttons
    'OperatorButtons will add the remaining buttons used to operate the calculator
    Set NumericButtons = ControlArray.Create(Me.Controls, ctCommandButton, "num")
    Set OperatorButtons = ControlArray.Create(Me.Controls, ctCommandButton, "op")
    
End Sub

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

Private Sub NumericButtons_onMouseMove(ctl As IControl, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    'highlight the button that the cursor is over
    ctl.Object.BackColor = &H8080FF
    mblnIsHighlighted = True
    
End Sub

Private Sub OperatorButtons_onMouseMove(ctl As IControl, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    'highlight the button that the cursor is over
    ctl.Object.BackColor = &H8000000D
    mblnIsHighlighted = True
    
End Sub

Private Sub NumericButtons_onClick(ctl As IControl)

    'This step is not required but I like to work with strongly typed variables
    'you could use ctl.Object.Caption instead of btn.Caption
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

Private Sub OperatorButtons_onClick(ctl As IControl)

    'This step is not required but I like to work with strongly typed variables
    'you could use ctl.Object.Caption instead of btn.Caption
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

'fix VBA_IDE enum bug that doesn't correct case when typing
#If False Then
    #Const Add = 0
    #Const Subtract = 0
    #Const Divide = 0
    #Const Multiply = 0
    #Const None = 0
#End If


