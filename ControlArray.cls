VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''
' VBA-ControlArray v1.0.0
' (c) J Martinez - https://github.com/bohicajr/vba-controlarray
'
'
''----------------------------------------------------------------------------''
Option Compare Binary
Option Explicit

Public Enum eControlTypes
    ctAll
    ctCheckbox
    ctComboBox
    ctCommandButton
    ctFrame
    ctImage
    ctLabel
    ctListBox
    ctMultiPage
    ctOptionButton
    ctScrollBar
    ctSpinButton
    ctTabStrip
    ctTextBox
    ctToggleButton
End Enum

Private Type udtMembers
    Col As Collection
End Type

Private This As udtMembers

'Common Events
Public Event onAddControl(ByRef ctl As IControl, ByVal Control As MSForms.Control)
Public Event onBeforeDragOver(ByRef ctl As IControl, ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
Public Event onBeforeDropOrPaste(ByRef ctl As IControl, ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
Public Event onChange(ByRef ctl As IControl)
Public Event onClick(ByRef ctl As IControl)
Public Event onDblClick(ByRef ctl As IControl, Cancel As MSForms.ReturnBoolean)
Public Event onDropButtonClick(ByRef ctl As IControl)
Public Event onError(ByRef ctl As IControl, ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
Public Event onKeyDown(ByRef ctl As IControl, ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Public Event onKeyPress(ByRef ctl As IControl, ByVal KeyAscii As MSForms.ReturnInteger)
Public Event onKeyUp(ByRef ctl As IControl, ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Public Event onLayout(ByRef ctl As IControl)
Public Event onMouseDown(ByRef ctl As IControl, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event onMouseMove(ByRef ctl As IControl, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event onMouseUp(ByRef ctl As IControl, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event onRemoveControl(ByRef ctl As IControl, ByVal Control As MSForms.Control)
Public Event onScroll(ByRef ctl As IControl)
Public Event onScrollFrame(ByRef ctl As IControl, ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
Public Event onSpinUp(ByRef ctl As IControl)
Public Event onSpinDown(ByRef ctl As IControl)
Public Event onZoom(ByRef ctl As IControl, Percent As Integer)

'Frame Control events with special parameters
Public Event onBeforeDragOverFrame(ByRef ctl As IControl, ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal State As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
Public Event onBeforeDropOrPasteFrame(ByRef ctl As IControl, ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

'TabStrip Control events with special parameters
Public Event onBeforeDragOverTabStrip(ByRef ctl As IControl, ByVal Index As Long, ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
Public Event onBeforeDropOrPasteTabStrip(ByRef ctl As IControl, ByVal Index As Long, ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
Public Event onClickTabStrip(ByRef ctl As IControl, ByVal Index As Long)
Public Event onDblClickTabStrip(ByRef ctl As IControl, ByVal Index As Long, ByVal Cancel As MSForms.ReturnBoolean)
Public Event onMouseDownTabStrip(ByRef ctl As IControl, ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event onMouseMoveTabStrip(ByRef ctl As IControl, ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event onMouseUpTabStrip(ByRef ctl As IControl, ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

'MultiPage Control events with special parameters
Public Event onAddControlMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal Control As MSForms.Control)
Public Event onBeforeDragOverMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal State As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
Public Event onBeforeDropOrPasteMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
Public Event onClickMultiPage(ByRef ctl As IControl, ByVal Index As Long)
Public Event onDblClickMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal Cancel As MSForms.ReturnBoolean)
Public Event onErrorMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
Public Event onLayoutMultiPage(ByRef ctl As IControl, ByVal Index As Long)
Public Event onMouseDownMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event onMouseMoveMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event onMouseUpMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event onRemoveControlMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal Control As MSForms.Control)
Public Event onScrollMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
Public Event onZoomMultiPage(ByRef ctl As IControl, ByVal Index As Long, Percent As Integer)

Public Function Create(Optional ByRef ctls As MSForms.Controls, Optional ByVal eType As eControlTypes = ctAll, Optional ByVal strTagFilter As String) As ControlArray

    Dim ctl As MSForms.Control

    With New ControlArray
        
        Set This.Col = New Collection
    
        If ctls Is Nothing = False Then

            For Each ctl In ctls

                If Len(strTagFilter) = 0 Or StrComp(ctl.Tag & vbNullString, strTagFilter, vbTextCompare) = 0 Then
                    
                    If TypeOf ctl Is MSForms.OptionButton And (eType = ctAll Or eType = ctOptionButton) Then
                        
                        Call This.Col.Add(CAOptionButton.Create(ctl, .Self), ctl.Name)
                    
                    ElseIf TypeOf ctl Is MSForms.ToggleButton And (eType = ctAll Or eType = ctToggleButton) Then
                        
                        Call This.Col.Add(CAToggleButton.Create(ctl, .Self), ctl.Name)
                        
                    ElseIf TypeOf ctl Is MSForms.CheckBox And (eType = ctAll Or eType = ctCheckbox) Then
                        
                        Call This.Col.Add(CACheckBox.Create(ctl, .Self), ctl.Name)
                    
                    ElseIf TypeOf ctl Is MSForms.ComboBox And (eType = ctAll Or eType = ctComboBox) Then
                        
                        Call This.Col.Add(CAComboBox.Create(ctl, .Self), ctl.Name)
                                                
                    ElseIf TypeOf ctl Is MSForms.CommandButton And (eType = ctAll Or eType = ctCommandButton) Then
                        
                        Call This.Col.Add(CACommandButton.Create(ctl, .Self), ctl.Name)
                                                        
                    ElseIf TypeOf ctl Is MSForms.Frame And (eType = ctAll Or eType = ctFrame) Then
                        
                        Call This.Col.Add(CAFrame.Create(ctl, .Self), ctl.Name)
                                                                            
                    ElseIf TypeOf ctl Is MSForms.Image And (eType = ctAll Or eType = ctImage) Then
                        
                        Call This.Col.Add(CAImage.Create(ctl, .Self), ctl.Name)
                            
                    ElseIf TypeOf ctl Is MSForms.Label And (eType = ctAll Or eType = ctLabel) Then
                        
                        Call This.Col.Add(CALabel.Create(ctl, .Self), ctl.Name)
                                                
                    ElseIf TypeOf ctl Is MSForms.MultiPage And (eType = ctAll Or eType = ctMultiPage) Then
                        
                        Call This.Col.Add(CAMultiPage.Create(ctl, .Self), ctl.Name)
                                                                        
                    ElseIf TypeOf ctl Is MSForms.ScrollBar And (eType = ctAll Or eType = ctScrollBar) Then
                        
                        Call This.Col.Add(CAScrollBar.Create(ctl, .Self), ctl.Name)
                                                                        
                    ElseIf TypeOf ctl Is MSForms.SpinButton And (eType = ctAll Or eType = ctSpinButton) Then
                        
                        Call This.Col.Add(CASpinButton.Create(ctl, .Self), ctl.Name)
                                                                        
                    ElseIf TypeOf ctl Is MSForms.TabStrip And (eType = ctAll Or eType = ctTabStrip) Then
                        
                        Call This.Col.Add(CATabStrip.Create(ctl, .Self), ctl.Name)
                        
                    ElseIf TypeOf ctl Is MSForms.TextBox And (eType = ctAll Or eType = ctTextBox) Then
                        
                        Call This.Col.Add(CATextBox.Create(ctl, .Self), ctl.Name)
                            
                    Else
                        'unsupported control for control array,ignore
                    End If
                
                End If
                
            Next
        
        End If
        
        Call .setMembers(This)
        
        Set Create = .Self
        
    End With

End Function

Friend Function Self() As ControlArray
    Set Self = Me
End Function

Friend Sub setMembers(ByRef Data As udtMembers)
    This = Data
End Sub

Public Sub Add(ByRef ctl As IControl)
    Call This.Col.Add(ctl, ctl.Name)
End Sub

Public Sub Remove(ByVal Index As Variant)
    Call This.Col.Remove(Index)
End Sub

Public Function Count() As Long
    Count = This.Col.Count
End Function

Public Sub Clear()
    Set This.Col = Nothing
    Set This.Col = New Collection
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = This.Col.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As IControl
Attribute Item.VB_UserMemId = 0
    Set Item = This.Col.Item(Index)
End Function

Public Property Get Exists(ByVal Index As Variant) As Boolean
    
    On Error Resume Next

    Call Me.Item(Index)

    If Err.Number = 0 Then
        Exists = True
    Else
        Exists = False
    End If

End Property

Friend Sub RaiseAddControl(ByRef ctl As IControl, ByVal Control As MSForms.Control)
    RaiseEvent onAddControl(ctl, Control)
End Sub

Friend Sub RaiseBeforeDragOver(ByRef ctl As IControl, ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    RaiseEvent onBeforeDragOver(ctl, Cancel, Data, X, Y, DragState, Effect, Shift)
End Sub

Friend Sub RaiseBeforeDropOrPaste(ByRef ctl As IControl, ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    RaiseEvent onBeforeDropOrPaste(ctl, Cancel, Action, Data, X, Y, Effect, Shift)
End Sub

Friend Sub RaiseChange(ByRef ctl As IControl)
    RaiseEvent onChange(ctl)
End Sub

Friend Sub RaiseClick(ByRef ctl As IControl)
    RaiseEvent onClick(ctl)
End Sub

Friend Sub RaiseDblClick(ByRef ctl As IControl, Cancel As MSForms.ReturnBoolean)
    RaiseEvent onDblClick(ctl, Cancel)
End Sub

Friend Sub RaiseDropButtonClick(ByRef ctl As IControl)
    RaiseEvent onDropButtonClick(ctl)
End Sub

Friend Sub RaiseError(ByRef ctl As IControl, ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
    RaiseEvent onError(ctl, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub

Friend Sub RaiseKeyDown(ByRef ctl As IControl, ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    RaiseEvent onKeyDown(ctl, KeyCode, Shift)
End Sub

Friend Sub RaiseKeyPress(ByRef ctl As IControl, ByVal KeyAscii As MSForms.ReturnInteger)
    RaiseEvent onKeyPress(ctl, KeyAscii)
End Sub

Friend Sub RaiseKeyUp(ByRef ctl As IControl, ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    RaiseEvent onKeyUp(ctl, KeyCode, Shift)
End Sub

Friend Sub RaiseLayout(ByRef ctl As IControl)
    RaiseEvent onLayout(ctl)
End Sub

Friend Sub RaiseMouseDown(ByRef ctl As IControl, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RaiseEvent onMouseDown(ctl, Button, Shift, X, Y)
End Sub

Friend Sub RaiseMouseMove(ByRef ctl As IControl, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RaiseEvent onMouseMove(ctl, Button, Shift, X, Y)
End Sub

Friend Sub RaiseMouseUp(ByRef ctl As IControl, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RaiseEvent onMouseUp(ctl, Button, Shift, X, Y)
End Sub

Friend Sub RaiseRemoveControl(ByRef ctl As IControl, ByVal Control As MSForms.Control)
    RaiseEvent onRemoveControl(ctl, Control)
End Sub

Friend Sub RaiseScroll(ByRef ctl As IControl)
    RaiseEvent onScroll(ctl)
End Sub

Friend Sub RaiseSpinUp(ByRef ctl As IControl)
    RaiseEvent onSpinUp(ctl)
End Sub

Friend Sub RaiseSpinDown(ByRef ctl As IControl)
    RaiseEvent onSpinDown(ctl)
End Sub

Friend Sub RaiseZoom(ByRef ctl As IControl, Percent As Integer)
    RaiseEvent onZoom(ctl, Percent)
End Sub

'Frame Control special cases
Friend Sub RaiseBeforeDragOverFrame(ByRef ctl As IControl, ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal State As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    RaiseEvent onBeforeDragOverFrame(ctl, Cancel, Control, Data, X, Y, State, Effect, Shift)
End Sub

Friend Sub RaiseBeforeDropOrPasteFrame(ByRef ctl As IControl, ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    RaiseEvent onBeforeDropOrPasteFrame(ctl, Cancel, Control, Action, Data, X, Y, Effect, Shift)
End Sub

Friend Sub RaiseScrollFrame(ByRef ctl As IControl, ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
    RaiseEvent onScrollFrame(ctl, ActionX, ActionY, RequestDx, RequestDy, ActualDx, ActualDy)
End Sub

'TabStrip Control special cases
Friend Sub RaiseBeforeDragOverTabStrip(ByRef ctl As IControl, ByVal Index As Long, ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    RaiseEvent onBeforeDragOverTabStrip(ctl, Index, Cancel, Data, X, Y, DragState, Effect, Shift)
End Sub

Friend Sub RaiseBeforeDropOrPasteTabStrip(ByRef ctl As IControl, ByVal Index As Long, ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    RaiseEvent onBeforeDropOrPasteTabStrip(ctl, Index, Cancel, Action, Data, X, Y, Effect, Shift)
End Sub

Friend Sub RaiseClickTabStrip(ByRef ctl As IControl, ByVal Index As Long)
    RaiseEvent onClickTabStrip(ctl, Index)
End Sub

Friend Sub RaiseDblClickTabStrip(ByRef ctl As IControl, ByVal Index As Long, ByVal Cancel As MSForms.ReturnBoolean)
    RaiseEvent onDblClickTabStrip(ctl, Index, Cancel)
End Sub

Friend Sub RaiseMouseDownTabStrip(ByRef ctl As IControl, ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RaiseEvent onMouseDownTabStrip(ctl, Index, Button, Shift, X, Y)
End Sub

Friend Sub RaiseMouseMoveTabStrip(ByRef ctl As IControl, ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RaiseEvent onMouseMoveTabStrip(ctl, Index, Button, Shift, X, Y)
End Sub

Friend Sub RaiseMouseUpTabStrip(ByRef ctl As IControl, ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RaiseEvent onMouseUpTabStrip(ctl, Index, Button, Shift, X, Y)
End Sub

'MutliPage Control special cases
Friend Sub RaiseAddControlMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal Control As MSForms.Control)
    RaiseEvent onAddControlMultiPage(ctl, Index, Control)
End Sub

Friend Sub RaiseBeforeDragOverMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal State As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    RaiseEvent onBeforeDragOverMultiPage(ctl, Index, Cancel, Control, Data, X, Y, State, Effect, Shift)
End Sub

Friend Sub RaiseBeforeDropOrPasteMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    RaiseEvent onBeforeDropOrPasteMultiPage(ctl, Index, Cancel, Control, Action, Data, X, Y, Effect, Shift)
End Sub

Friend Sub RaiseClickMultiPage(ByRef ctl As IControl, ByVal Index As Long)
    RaiseEvent onClickMultiPage(ctl, Index)
End Sub

Friend Sub RaiseDblClickMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal Cancel As MSForms.ReturnBoolean)
    RaiseEvent onDblClickMultiPage(ctl, Index, Cancel)
End Sub

Friend Sub RaiseErrorMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
    RaiseEvent onErrorMultiPage(ctl, Index, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub

Friend Sub RaiseLayoutMultiPage(ByRef ctl As IControl, ByVal Index As Long)
    RaiseEvent onLayoutMultiPage(ctl, Index)
End Sub

Friend Sub RaiseMouseDownMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RaiseEvent onMouseDownMultiPage(ctl, Index, Button, Shift, X, Y)
End Sub

Friend Sub RaiseMouseMoveMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RaiseEvent onMouseMoveMultiPage(ctl, Index, Button, Shift, X, Y)
End Sub

Friend Sub RaiseMouseUpMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RaiseEvent onMouseUpMultiPage(ctl, Index, Button, Shift, X, Y)
End Sub

Friend Sub RaiseRemoveControlMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal Control As MSForms.Control)
    RaiseEvent onRemoveControlMultiPage(ctl, Index, Control)
End Sub

Friend Sub RaiseScrollMultiPage(ByRef ctl As IControl, ByVal Index As Long, ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
    RaiseEvent onScrollMultiPage(ctl, Index, ActionX, ActionY, RequestDx, RequestDy, ActualDx, ActualDy)
End Sub

Friend Sub RaiseZoomMultiPage(ByRef ctl As IControl, ByVal Index As Long, Percent As Integer)
    RaiseEvent onZoomMultiPage(ctl, Index, Percent)
End Sub

'Fix for VBIDE bug that does not correct case on enumerations
'values are irrelavent
#If False Then
    #Const ctAll = 0
    #Const ctCheckbox = 0
    #Const ctComboBox = 0
    #Const ctCommandButton = 0
    #Const ctFrame = 0
    #Const ctImage = 0
    #Const ctLabel = 0
    #Const ctListBox = 0
    #Const ctMultiPage = 0
    #Const ctOptionButton = 0
    #Const ctScrollBar = 0
    #Const ctSpinButton = 0
    #Const ctTabStrip = 0
    #Const ctTextBox = 0
    #Const ctToggleButton = 0
#End If

