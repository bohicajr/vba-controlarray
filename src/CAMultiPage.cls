VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAMultiPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Binary
Option Explicit

Implements IControl

Private WithEvents MSFormsCtl As MSForms.MultiPage
Attribute MSFormsCtl.VB_VarHelpID = -1

Private Const CONTROL_TYPE As Long = eControlTypes.ctMultiPage

Private Type udtMembers
    Name As String
    ControlObject As MSForms.Control
    Parent As ControlArray
End Type

Private This As udtMembers

Public Function Create(ByRef ctl As MSForms.Control, ByRef Parent As ControlArray) As IControl

    With New CAMultiPage

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

Private Sub MSFormsCtl_AddControl(ByVal Index As Long, ByVal Control As MSForms.Control)
    Call This.Parent.RaiseAddControlMultiPage(Me, Index, Control)
End Sub

Private Sub MSFormsCtl_BeforeDragOver(ByVal Index As Long, ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal State As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Call This.Parent.RaiseBeforeDragOverMultiPage(Me, Index, Cancel, Control, Data, X, Y, State, Effect, Shift)
End Sub

Private Sub MSFormsCtl_BeforeDropOrPaste(ByVal Index As Long, ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Call This.Parent.RaiseBeforeDropOrPasteMultiPage(Me, Index, Cancel, Control, Action, Data, X, Y, Effect, Shift)
End Sub

Private Sub MSFormsCtl_Change()
    Call This.Parent.RaiseChange(Me)
End Sub

Private Sub MSFormsCtl_Click(ByVal Index As Long)
    Call This.Parent.RaiseClickMultiPage(Me, Index)
End Sub

Private Sub MSFormsCtl_DblClick(ByVal Index As Long, ByVal Cancel As MSForms.ReturnBoolean)
    Call This.Parent.RaiseDblClickMultiPage(Me, Index, Cancel)
End Sub

Private Sub MSFormsCtl_Error(ByVal Index As Long, ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
    Call This.Parent.RaiseErrorMultiPage(Me, Index, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay)
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

Private Sub MSFormsCtl_Layout(ByVal Index As Long)
    Call This.Parent.RaiseLayoutMultiPage(Me, Index)
End Sub

Private Sub MSFormsCtl_MouseDown(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call This.Parent.RaiseMouseDownMultiPage(Me, Index, Button, Shift, X, Y)
End Sub

Private Sub MSFormsCtl_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call This.Parent.RaiseMouseMoveMultiPage(Me, Index, Button, Shift, X, Y)
End Sub

Private Sub MSFormsCtl_MouseUp(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call This.Parent.RaiseMouseUpMultiPage(Me, Index, Button, Shift, X, Y)
End Sub

Private Sub MSFormsCtl_RemoveControl(ByVal Index As Long, ByVal Control As MSForms.Control)
    Call This.Parent.RaiseRemoveControlMultiPage(Me, Index, Control)
End Sub

Private Sub MSFormsCtl_Scroll(ByVal Index As Long, ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
    Call This.Parent.RaiseScrollMultiPage(Me, Index, ActionX, ActionY, RequestDx, RequestDy, ActualDx, ActualDy)
End Sub

Private Sub MSFormsCtl_Zoom(ByVal Index As Long, Percent As Integer)
    Call This.Parent.RaiseZoomMultiPage(Me, Index, Percent)
End Sub
