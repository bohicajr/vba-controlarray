VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CATextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Binary
Option Explicit

Implements IControl

Private WithEvents MSFormsCtl As MSForms.TextBox
Attribute MSFormsCtl.VB_VarHelpID = -1

Private Const CONTROL_TYPE As Long = eControlTypes.ctTextBox

Private Type udtMembers
    Name As String
    ControlObject As MSForms.Control
    Parent As ControlArray
End Type

Private This As udtMembers

Public Function Create(ByRef ctl As MSForms.Control, ByRef Parent As ControlArray) As IControl

    With New CATextBox

        This.Name = ctl.Name
        Set This.ControlObject = ctl
        Set This.Parent = Parent
        Call .setMembers(This)
    
        Set Create = .Self
        
    End With
    
End Function

Friend Function Self() As IControl
    Set Self = Me
End Function

Friend Sub setMembers(ByRef Data As udtMembers)

    This = Data
    
    Set MSFormsCtl = Data.ControlObject
    
End Sub

Public Property Get ControlType() As eControlTypes
    ControlType = CONTROL_TYPE
End Property

Public Property Get Name() As String
    Name = This.Name
End Property

Public Property Get Object() As MSForms.Control
    Set Object = MSFormsCtl
End Property

Private Property Get iControl_ControlType() As eControlTypes
    iControl_ControlType = Me.ControlType
End Property

Private Property Get iControl_Name() As String
    iControl_Name = Me.Name
End Property

Private Property Get iControl_Object() As MSForms.Control
    Set iControl_Object = Me.Object
End Property

Private Sub MSFormsCtl_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Call This.Parent.RaiseBeforeDragOver(Me, Cancel, Data, X, Y, DragState, Effect, Shift)
End Sub

Private Sub MSFormsCtl_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Call This.Parent.RaiseBeforeDropOrPaste(Me, Cancel, Action, Data, X, Y, Effect, Shift)
End Sub

Private Sub MSFormsCtl_Change()
    Call This.Parent.RaiseChange(Me)
End Sub

Private Sub MSFormsCtl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call This.Parent.RaiseDblClick(Me, Cancel)
End Sub

Private Sub MSFormsCtl_DropButtonClick()
    Call This.Parent.RaiseDropButtonClick(Me)
End Sub

Private Sub MSFormsCtl_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
    Call This.Parent.RaiseError(Me, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub

Private Sub MSFormsCtl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call This.Parent.RaiseKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub MSFormsCtl_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call This.Parent.RaiseKeyPress(Me, KeyAscii)
End Sub

Private Sub MSFormsCtl_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call This.Parent.RaiseKeyUp(Me, KeyCode, Shift)
End Sub

Private Sub MSFormsCtl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call This.Parent.RaiseMouseDown(Me, Button, Shift, X, Y)
End Sub

Private Sub MSFormsCtl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call This.Parent.RaiseMouseMove(Me, Button, Shift, X, Y)
End Sub

Private Sub MSFormsCtl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call This.Parent.RaiseMouseUp(Me, Button, Shift, X, Y)
End Sub
